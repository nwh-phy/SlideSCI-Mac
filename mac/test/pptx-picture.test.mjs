import test from 'node:test';
import assert from 'node:assert/strict';
import path from 'node:path';
import JSZip from 'jszip';
import sharp from 'sharp';
import PptxGenJS from 'pptxgenjs';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { parsePresentation, extractImages, imageFragment, replacePicture } from '../lib/pptx.mjs';

const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const TAGS = 'application/vnd.openxmlformats-officedocument.presentationml.tags+xml';
const serializer = new XMLSerializer();

const xml = text => new DOMParser({ onError(level) { if (level !== 'warning') throw new Error('invalid fixture XML'); } })
  .parseFromString(text, 'application/xml');
function children(node, name) {
  const result = [];
  for (let item = node?.firstChild; item; item = item.nextSibling) if (item.nodeType === 1 && (!name || item.localName === name)) result.push(item);
  return result;
}
const child = (node, name) => children(node, name)[0] || null;
function descendants(node, name) {
  const result = [];
  for (const item of children(node)) {
    if (item.localName === name) result.push(item);
    result.push(...descendants(item, name));
  }
  return result;
}
const first = (node, name) => descendants(node, name)[0] || null;
const attr = (node, key, fallback = '') => node?.hasAttribute(key) ? node.getAttribute(key) : fallback;
function element(doc, namespace, name, attributes = {}) {
  const node = doc.createElementNS(namespace, name);
  for (const [key, value] of Object.entries(attributes)) node.setAttribute(key, String(value));
  return node;
}
const picture = (doc, id) => descendants(doc, 'pic').find(node => attr(first(node, 'cNvPr'), 'id') === String(id));
const shapeById = (model, id) => model.slides[0].shapes
  .flatMap(function flatten(shape) { return [shape, ...(shape.children || []).flatMap(flatten)]; })
  .find(shape => shape.id === String(id));
const dataUri = bytes => 'image/png;base64,' + bytes.toString('base64');
const resolveTarget = target => target.startsWith('/') ? target.slice(1) : path.posix.normalize(path.posix.join('ppt/slides', target));

async function twoTone(width, height, left, right) {
  const half = Math.floor(width / 2);
  const overlay = await sharp({ create: { width: width - half, height, channels: 4, background: right } }).png().toBuffer();
  return sharp({ create: { width, height, channels: 4, background: left } })
    .composite([{ input: overlay, left: half, top: 0 }]).png().toBuffer();
}

function addTagReference(doc, pic, relationshipId) {
  const container = element(doc, P, 'p:custDataLst');
  const tags = element(doc, P, 'p:tags');
  tags.setAttributeNS(R, 'r:id', relationshipId);
  container.appendChild(tags);
  child(child(pic, 'nvPicPr'), 'nvPr').appendChild(container);
}
function tagPart(values) {
  const doc = xml(`<p:tagLst xmlns:p="${P}"/>`);
  for (const [name, value] of Object.entries(values)) {
    const tag = element(doc, P, 'p:tag', { name, val: value });
    doc.documentElement.appendChild(tag);
  }
  return serializer.serializeToString(doc);
}
function addRelationship(doc, id, type, target) {
  const relationship = element(doc, REL, 'Relationship', { Id: id, Type: type, Target: target });
  doc.documentElement.appendChild(relationship);
  return relationship;
}
function addOverride(doc, part, contentType) {
  doc.documentElement.appendChild(element(doc, CT, 'Override', { PartName: part, ContentType: contentType }));
}
function wrap(doc, tree, members, { id, name, cx, cy }) {
  const group = element(doc, P, 'p:grpSp');
  const nonVisual = element(doc, P, 'p:nvGrpSpPr');
  nonVisual.appendChild(element(doc, P, 'p:cNvPr', { id, name }));
  nonVisual.appendChild(element(doc, P, 'p:cNvGrpSpPr'));
  nonVisual.appendChild(element(doc, P, 'p:nvPr'));
  group.appendChild(nonVisual);
  const properties = element(doc, P, 'p:grpSpPr');
  const transform = element(doc, A, 'a:xfrm');
  transform.appendChild(element(doc, A, 'a:off', { x: 0, y: 0 }));
  transform.appendChild(element(doc, A, 'a:ext', { cx, cy }));
  transform.appendChild(element(doc, A, 'a:chOff', { x: 0, y: 0 }));
  transform.appendChild(element(doc, A, 'a:chExt', { cx, cy }));
  properties.appendChild(transform);
  group.appendChild(properties);
  tree.insertBefore(group, members[0]);
  for (const member of members) { tree.removeChild(member); group.appendChild(member); }
  return group;
}
function addCrop(doc, pic, crop) {
  const fill = child(pic, 'blipFill');
  const source = element(doc, A, 'a:srcRect', {
    l: Math.round(crop.left * 100000), t: Math.round(crop.top * 100000),
    r: Math.round(crop.right * 100000), b: Math.round(crop.bottom * 100000) });
  fill.insertBefore(source, child(fill, 'stretch'));
}
function relationship(doc, id) {
  return children(doc.documentElement, 'Relationship').find(node => attr(node, 'Id') === id);
}
function slideSize(doc) {
  const size = first(doc, 'sldSz');
  return { cx: attr(size, 'cx'), cy: attr(size, 'cy') };
}
async function zipDocument(zip, name) { return xml(await zip.file(name).async('string')); }

// A real deck whose target picture is nested in a group next to a shape and another picture
// that references the very same media part. The target carries crop, rotation, alt text and tags.
async function groupedDeck() {
  const sharedBytes = await twoTone(4, 2, '#e11d48ff', '#1d4ed8ff');
  const replacementBytes = await sharp({ create: { width: 3, height: 1, channels: 4, background: '#16a34aff' } }).png().toBuffer();
  const deck = new PptxGenJS();
  deck.layout = 'LAYOUT_WIDE';
  const slide = deck.addSlide();
  slide.addImage({ data: dataUri(sharedBytes), x: 1, y: 1, w: 2, h: 1 });
  slide.addShape(deck.ShapeType.rect, { x: 7, y: 1, w: 1, h: 1, fill: { color: '00FF00' }, line: { color: '00FF00' } });
  slide.addImage({ data: dataUri(sharedBytes), x: 4, y: 1, w: 2, h: 1 });
  const zip = await JSZip.loadAsync(Buffer.from(await deck.write({ outputType: 'base64' }), 'base64'));
  const slidePath = 'ppt/slides/slide1.xml', relsPath = 'ppt/slides/_rels/slide1.xml.rels';
  const doc = await zipDocument(zip, slidePath), rels = await zipDocument(zip, relsPath);
  const tree = first(doc, 'spTree');
  const target = picture(doc, 2), shared = picture(doc, 4);
  const rectangle = descendants(doc, 'sp').find(node => attr(first(node, 'cNvPr'), 'id') === '3');
  const targetCrop = { left: .1, top: .05, right: .2, bottom: .15 };
  const sharedCrop = { left: .02, top: .01, right: .03, bottom: .04 };
  const targetCaption = first(target, 'cNvPr');
  targetCaption.setAttribute('descr', '目标图片 α');
  targetCaption.setAttribute('title', 'Target title');
  child(child(target, 'spPr'), 'xfrm').setAttribute('rot', '900000');
  addCrop(doc, target, targetCrop);
  addTagReference(doc, target, 'rIdSlideSCITagsTarget');
  const sharedCaption = first(shared, 'cNvPr');
  sharedCaption.setAttribute('descr', '共享源图');
  sharedCaption.setAttribute('title', 'Shared source');
  addCrop(doc, shared, sharedCrop);
  addTagReference(doc, shared, 'rIdSlideSCITagsShared');
  addRelationship(rels, 'rIdSlideSCITagsTarget', R + '/tags', '../tags/tag1.xml');
  addRelationship(rels, 'rIdSlideSCITagsShared', R + '/tags', '../tags/tag2.xml');
  const targetRid = child(child(target, 'blipFill'), 'blip').getAttributeNS(R, 'embed');
  const sharedBlip = child(child(shared, 'blipFill'), 'blip');
  const sharedRid = sharedBlip.getAttributeNS(R, 'embed');
  const targetPath = resolveTarget(attr(relationship(rels, targetRid), 'Target'));
  const sharedRelation = relationship(rels, sharedRid);
  const sharedPath = resolveTarget(attr(sharedRelation, 'Target'));
  sharedBlip.setAttributeNS(R, 'r:embed', targetRid);
  if (sharedPath !== targetPath) {
    zip.remove(sharedPath);
    sharedRelation.parentNode.removeChild(sharedRelation);
  }
  const types = await zipDocument(zip, '[Content_Types].xml');
  addOverride(types, '/ppt/tags/tag1.xml', TAGS);
  addOverride(types, '/ppt/tags/tag2.xml', TAGS);
  const size = slideSize(await zipDocument(zip, 'ppt/presentation.xml'));
  wrap(doc, tree, [target, rectangle, shared], { id: '101', name: 'Grouped figures', ...size });
  zip.file(slidePath, serializer.serializeToString(doc));
  zip.file(relsPath, serializer.serializeToString(rels));
  zip.file('[Content_Types].xml', serializer.serializeToString(types));
  zip.file('ppt/tags/tag1.xml', tagPart({ SLIDESCI_IMAGE_UUID: 'TARGET-UUID', SLIDESCI_FOV_VISIBLE: '1' }));
  zip.file('ppt/tags/tag2.xml', tagPart({ SLIDESCI_IMAGE_UUID: 'SHARED-UUID' }));
  return { pptx: await zip.generateAsync({ type: 'base64', compression: 'DEFLATE' }),
    targetId: '2', rectangleId: '3', sharedId: '4', groupId: '101',
    sharedBytes, replacementBytes, sharedPath: targetPath,
    targetCrop, sharedCrop };
}

// A real deck with one free-standing picture and one picture nested in a group. Both are cropped
// and rotated, so original:true has to recover the source geometry rather than the slide geometry.
async function fragmentDeck() {
  const plainBytes = await twoTone(4, 2, '#f97316ff', '#0ea5e9ff');
  const overlay = await sharp({ create: { width: 4, height: 4, channels: 4, background: '#eab308ff' } }).png().toBuffer();
  const groupedBytes = await sharp({ create: { width: 8, height: 4, channels: 4, background: '#22c55eff' } })
    .composite([{ input: overlay, left: 0, top: 0 }]).png().toBuffer();
  const deck = new PptxGenJS();
  deck.layout = 'LAYOUT_WIDE';
  const slide = deck.addSlide();
  slide.addImage({ data: dataUri(plainBytes), x: 1, y: 1, w: 2, h: 1 });
  slide.addShape(deck.ShapeType.rect, { x: 8, y: 1, w: 1, h: .5, fill: { color: '000000' } });
  slide.addImage({ data: dataUri(groupedBytes), x: 4, y: 3, w: 4, h: 2 });
  const zip = await JSZip.loadAsync(Buffer.from(await deck.write({ outputType: 'base64' }), 'base64'));
  const slidePath = 'ppt/slides/slide1.xml', relsPath = 'ppt/slides/_rels/slide1.xml.rels';
  const doc = await zipDocument(zip, slidePath), rels = await zipDocument(zip, relsPath);
  const tree = first(doc, 'spTree');
  const plain = picture(doc, 2), grouped = picture(doc, 4);
  const rectangle = descendants(doc, 'sp').find(node => attr(first(node, 'cNvPr'), 'id') === '3');
  const plainCaption = first(plain, 'cNvPr');
  plainCaption.setAttribute('descr', 'plain-alt');
  child(child(plain, 'spPr'), 'xfrm').setAttribute('rot', '1800000');
  addCrop(doc, plain, { left: .25, top: 0, right: 0, bottom: 0 });
  const groupedCaption = first(grouped, 'cNvPr');
  groupedCaption.setAttribute('descr', 'grouped-alt');
  child(child(grouped, 'spPr'), 'xfrm').setAttribute('rot', '-2700000');
  addCrop(doc, grouped, { left: 0, top: .1, right: 0, bottom: .2 });
  addTagReference(doc, grouped, 'rIdSlideSCITagsGrouped');
  addRelationship(rels, 'rIdSlideSCITagsGrouped', R + '/tags', '../tags/tag1.xml');
  const types = await zipDocument(zip, '[Content_Types].xml');
  addOverride(types, '/ppt/tags/tag1.xml', TAGS);
  const size = slideSize(await zipDocument(zip, 'ppt/presentation.xml'));
  wrap(doc, tree, [grouped, rectangle], { id: '101', name: 'Measurements', ...size });
  zip.file(slidePath, serializer.serializeToString(doc));
  zip.file(relsPath, serializer.serializeToString(rels));
  zip.file('[Content_Types].xml', serializer.serializeToString(types));
  zip.file('ppt/tags/tag1.xml', tagPart({ SLIDESCI_FOV_VISIBLE: '1', SLIDESCI_IMAGE_UUID: 'GROUPED-UUID' }));
  return { pptx: await zip.generateAsync({ type: 'base64', compression: 'DEFLATE' }),
    plainId: '2', groupedId: '4', groupId: '101', plainBytes, groupedBytes };
}

function pixel(data, info, x, y) { const offset = (y * info.width + x) * 4; return [...data.subarray(offset, offset + 4)]; }

test('replacePicture rewrites a nested picture while the group hierarchy, siblings and appearance stay intact', async () => {
  const fixture = await groupedDeck();
  const parsed = await parsePresentation(fixture.pptx);
  const originalTarget = shapeById(parsed, fixture.targetId), originalShared = shapeById(parsed, fixture.sharedId);
  assert.equal(originalTarget.image.path, originalShared.image.path, 'fixture pictures must share one media part');
  const result = await replacePicture(fixture.pptx, { slideIndex: 0, id: fixture.targetId, rootId: fixture.groupId,
    data: fixture.replacementBytes.toString('base64'), mime: 'image/png' });
  assert.deepEqual(result.shapeIds, [fixture.groupId]);
  assert.equal(result.name, 'Grouped figures');

  const replaced = await parsePresentation(result.pptx);
  assert.equal(replaced.slides[0].shapes.length, 1);
  const group = replaced.slides[0].shapes[0];
  assert.equal(group.type, 'group');
  assert.equal(group.id, fixture.groupId);
  assert.equal(group.name, 'Grouped figures');
  assert.deepEqual(group.children.map(shape => shape.id), [fixture.targetId, fixture.rectangleId, fixture.sharedId]);
  const target = group.children[0], rectangle = group.children[1], shared = group.children[2];

  assert.equal(target.type, 'picture');
  assert.equal(target.parentId, fixture.groupId);
  assert.deepEqual({ left: target.left, top: target.top, width: target.width, height: target.height, rotation: target.rotation },
    { left: 72, top: 72, width: 144, height: 72, rotation: 15 });
  assert.deepEqual(target.crop, fixture.targetCrop);
  assert.equal(target.alternativeText, '目标图片 α');
  assert.equal(target.title, 'Target title');
  assert.deepEqual(target.tags, { SLIDESCI_IMAGE_UUID: 'TARGET-UUID', SLIDESCI_FOV_VISIBLE: '1' });

  assert.equal(rectangle.type, 'shape');
  assert.deepEqual({ left: rectangle.left, top: rectangle.top, width: rectangle.width, height: rectangle.height },
    { left: 504, top: 72, width: 72, height: 72 });
  assert.equal(rectangle.parentId, fixture.groupId);

  assert.equal(shared.type, 'picture');
  assert.equal(shared.parentId, fixture.groupId);
  assert.deepEqual({ left: shared.left, top: shared.top, width: shared.width, height: shared.height },
    { left: 288, top: 72, width: 144, height: 72 });
  assert.deepEqual(shared.crop, fixture.sharedCrop);
  assert.equal(shared.alternativeText, '共享源图');
  assert.deepEqual(shared.tags, { SLIDESCI_IMAGE_UUID: 'SHARED-UUID' });

  assert.notEqual(target.image.path, shared.image.path);
  assert.equal(shared.image.path, originalShared.image.path);
  assert.notEqual(target.image.digest, shared.image.digest);
  assert.notEqual(target.image.digest, originalTarget.image.digest);
  assert.match(target.image.path, /^ppt\/media\/slidesci-[0-9a-f]{64}\.png$/);

  const targetImages = await extractImages(result.pptx, { slideIndex: 0, ids: [fixture.targetId] });
  assert.equal(targetImages[0].mime, 'image/png');
  assert.equal(targetImages[0].width, 3);
  assert.equal(targetImages[0].height, 1);
  assert.deepEqual(Buffer.from(targetImages[0].data, 'base64'), fixture.replacementBytes);
  const sharedImages = await extractImages(result.pptx, { slideIndex: 0, ids: [fixture.sharedId] });
  assert.deepEqual(Buffer.from(sharedImages[0].data, 'base64'), fixture.sharedBytes);
  const sharedPixels = await sharp(fixture.sharedBytes).ensureAlpha().raw().toBuffer({ resolveWithObject: true });
  assert.deepEqual(pixel(sharedPixels.data, sharedPixels.info, 0, 0), [225, 29, 72, 255]);
  assert.deepEqual(pixel(sharedPixels.data, sharedPixels.info, sharedPixels.info.width - 1, 0), [29, 78, 216, 255]);
});

test('replacePicture never overwrites media that another picture still references', async () => {
  const fixture = await groupedDeck();
  const result = await replacePicture(fixture.pptx, { slideIndex: 0, id: fixture.targetId, rootId: fixture.groupId,
    data: fixture.replacementBytes.toString('base64') });
  const images = await extractImages(result.pptx, { slideIndex: 0, ids: [fixture.sharedId, fixture.targetId] });
  const shared = images.find(item => item.id === fixture.sharedId), target = images.find(item => item.id === fixture.targetId);
  assert.deepEqual(Buffer.from(shared.data, 'base64'), fixture.sharedBytes);
  assert.notDeepEqual(Buffer.from(target.data, 'base64'), fixture.sharedBytes);
  const zip = await JSZip.loadAsync(result.pptx, { base64: true });
  assert.deepEqual(await zip.file(fixture.sharedPath).async('nodebuffer'), fixture.sharedBytes);
  assert.deepEqual(Buffer.from(target.data, 'base64'), fixture.replacementBytes);
});

test('replacePicture honors preserveCrop:false for the selected picture only', async () => {
  const fixture = await groupedDeck();
  const result = await replacePicture(fixture.pptx, { slideIndex: 0, id: fixture.targetId, rootId: fixture.groupId,
    data: fixture.replacementBytes.toString('base64'), preserveCrop: false });
  const parsed = await parsePresentation(result.pptx);
  const target = shapeById(parsed, fixture.targetId), shared = shapeById(parsed, fixture.sharedId);
  assert.deepEqual(target.crop, { left: 0, top: 0, right: 0, bottom: 0 });
  assert.equal(target.rotation, 15);
  assert.equal(target.alternativeText, '目标图片 α');
  assert.deepEqual(target.tags, { SLIDESCI_IMAGE_UUID: 'TARGET-UUID', SLIDESCI_FOV_VISIBLE: '1' });
  assert.deepEqual(shared.crop, fixture.sharedCrop);
  const zip = await JSZip.loadAsync(result.pptx, { base64: true });
  const doc = await zipDocument(zip, 'ppt/slides/slide1.xml');
  const fill = child(picture(doc, fixture.targetId), 'blipFill');
  assert.equal(child(fill, 'srcRect'), null);
  assert.ok(child(fill, 'stretch'));
});

test('replacePicture keeps the SVG original and embeds a real PNG compatibility preview', async () => {
  const fixture = await groupedDeck();
  const svgBytes = Buffer.from('<svg xmlns="http://www.w3.org/2000/svg" width="6" height="3" viewBox="0 0 6 3"><rect width="6" height="3" fill="#0044ff"/><rect width="3" height="3" fill="#ffcc00"/></svg>');
  const result = await replacePicture(fixture.pptx, { slideIndex: 0, id: fixture.targetId, rootId: fixture.groupId,
    data: svgBytes.toString('base64'), mime: 'image/svg+xml' });
  const parsed = await parsePresentation(result.pptx);
  const target = shapeById(parsed, fixture.targetId);
  assert.equal(target.image.mime, 'image/svg+xml');
  assert.deepEqual(target.crop, fixture.targetCrop);
  const images = await extractImages(result.pptx, { slideIndex: 0, ids: [fixture.targetId] });
  assert.equal(images[0].mime, 'image/svg+xml');
  assert.equal(images[0].width, 6);
  assert.equal(images[0].height, 3);
  assert.deepEqual(Buffer.from(images[0].data, 'base64'), svgBytes);
  assert.deepEqual(Buffer.from((await extractImages(result.pptx, { slideIndex: 0, ids: [fixture.sharedId] }))[0].data, 'base64'), fixture.sharedBytes);

  const zip = await JSZip.loadAsync(result.pptx, { base64: true });
  const doc = await zipDocument(zip, 'ppt/slides/slide1.xml');
  const fill = child(picture(doc, fixture.targetId), 'blipFill');
  const blip = child(fill, 'blip');
  const svgBlip = first(child(blip, 'extLst'), 'svgBlip');
  assert.ok(svgBlip);
  const primaryRid = blip.getAttributeNS(R, 'embed'), svgRid = svgBlip.getAttributeNS(R, 'embed');
  assert.ok(primaryRid && svgRid);
  assert.notEqual(primaryRid, svgRid);
  assert.equal(blip.hasAttributeNS(R, 'link'), false);
  const rels = await zipDocument(zip, 'ppt/slides/_rels/slide1.xml.rels');
  const primaryPath = resolveTarget(attr(relationship(rels, primaryRid), 'Target'));
  const svgPath = resolveTarget(attr(relationship(rels, svgRid), 'Target'));
  assert.match(primaryPath, /\.png$/);
  assert.match(svgPath, /\.svg$/);
  assert.notEqual(primaryPath, svgPath);
  const preview = await zip.file(primaryPath).async('nodebuffer');
  const svgPart = await zip.file(svgPath).async('nodebuffer');
  assert.equal(preview.subarray(0, 8).toString('hex'), '89504e470d0a1a0a');
  assert.deepEqual(svgPart, svgBytes);
  const metadata = await sharp(preview).metadata();
  assert.equal(metadata.width / metadata.height, 2);
  const pixels = await sharp(preview).ensureAlpha().raw().toBuffer({ resolveWithObject: true });
  assert.deepEqual(pixel(pixels.data, pixels.info, Math.floor(pixels.info.width * .25), Math.floor(pixels.info.height / 2)), [255, 204, 0, 255]);
  assert.deepEqual(pixel(pixels.data, pixels.info, Math.floor(pixels.info.width * .75), Math.floor(pixels.info.height / 2)), [0, 68, 255, 255]);
  const types = await zipDocument(zip, '[Content_Types].xml');
  const defaults = new Map(children(types.documentElement, 'Default').map(node => [attr(node, 'Extension'), attr(node, 'ContentType')]));
  assert.equal(defaults.get('svg'), 'image/svg+xml');
  assert.equal(defaults.get('png'), 'image/png');
});

test('imageFragment original:true isolates uncropped, unrotated sources at intrinsic page bounds', async () => {
  const fixture = await fragmentDeck();
  const grouped = await imageFragment(fixture.pptx, { slideIndex: 0, id: fixture.groupedId });
  assert.deepEqual(grouped.bounds, { left: 0, top: 0, width: 6, height: 3 });
  const groupedModel = await parsePresentation(grouped.pptx);
  assert.equal(groupedModel.width, 6);
  assert.equal(groupedModel.height, 3);
  assert.equal(groupedModel.slides[0].shapes.length, 1);
  const groupedShape = groupedModel.slides[0].shapes[0];
  assert.equal(groupedShape.type, 'picture');
  assert.equal(groupedShape.id, fixture.groupedId);
  assert.equal(groupedShape.parentId, undefined);
  assert.deepEqual({ left: groupedShape.left, top: groupedShape.top, width: groupedShape.width, height: groupedShape.height, rotation: groupedShape.rotation },
    { left: 0, top: 0, width: 6, height: 3, rotation: 0 });
  assert.deepEqual(groupedShape.crop, { left: 0, top: 0, right: 0, bottom: 0 });
  assert.equal(groupedShape.alternativeText, 'grouped-alt');
  assert.deepEqual(groupedShape.tags, { SLIDESCI_FOV_VISIBLE: '1', SLIDESCI_IMAGE_UUID: 'GROUPED-UUID' });
  assert.deepEqual(Buffer.from((await extractImages(grouped.pptx, { slideIndex: 0, ids: [fixture.groupedId] }))[0].data, 'base64'), fixture.groupedBytes);

  const plain = await imageFragment(fixture.pptx, { slideIndex: 0, id: fixture.plainId });
  assert.deepEqual(plain.bounds, { left: 0, top: 0, width: 3, height: 1.5 });
  const plainModel = await parsePresentation(plain.pptx);
  assert.equal(plainModel.width, 3);
  assert.equal(plainModel.height, 1.5);
  assert.equal(plainModel.slides[0].shapes.length, 1);
  const plainShape = plainModel.slides[0].shapes[0];
  assert.equal(plainShape.type, 'picture');
  assert.equal(plainShape.parentId, undefined);
  assert.deepEqual({ left: plainShape.left, top: plainShape.top, width: plainShape.width, height: plainShape.height, rotation: plainShape.rotation },
    { left: 0, top: 0, width: 3, height: 1.5, rotation: 0 });
  assert.deepEqual(plainShape.crop, { left: 0, top: 0, right: 0, bottom: 0 });
  assert.equal(plainShape.alternativeText, 'plain-alt');
  assert.deepEqual(Buffer.from((await extractImages(plain.pptx, { slideIndex: 0, ids: [fixture.plainId] }))[0].data, 'base64'), fixture.plainBytes);
});
