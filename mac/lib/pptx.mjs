import { createHash } from 'node:crypto';
import path from 'node:path';
import JSZip from 'jszip';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { readImageDimensions } from './images.mjs';
import { prunePresentation } from './pptx-prune.mjs';
import { applyTextFormatting } from './pptx-format.mjs';

const EMU = 12700;
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const serializer = new XMLSerializer();
const assetRevisions = new WeakMap();
const assetIndexes = new WeakMap();
const packageCaches = new WeakMap();

function cachesFor(zip) {
  let cache = packageCaches.get(zip);
  if (!cache) {
    cache = { relationships: new Map(), themes: new Map(), tags: new Map() };
    packageCaches.set(zip, cache);
  }
  return cache;
}
function invalidatePackage(zip) { packageCaches.delete(zip); }
function indexFor(zip) {
  let index = assetIndexes.get(zip);
  if (!index) { index = { bytes: null, ranges: null }; assetIndexes.set(zip, index); }
  return index;
}

function compressedAssetRanges(bytes) {
  const result = new Map();
  let end = bytes.length - 22;
  const lower = Math.max(0, bytes.length - 65557);
  while (end >= lower && bytes.readUInt32LE(end) !== 0x06054b50) end--;
  if (end < lower) return result;
  const count = bytes.readUInt16LE(end + 10);
  let offset = bytes.readUInt32LE(end + 16);
  for (let index = 0; index < count && offset + 46 <= bytes.length; index++) {
    if (bytes.readUInt32LE(offset) !== 0x02014b50) break;
    const length = bytes.readUInt32LE(offset + 20);
    const nameLength = bytes.readUInt16LE(offset + 28);
    const extraLength = bytes.readUInt16LE(offset + 30);
    const commentLength = bytes.readUInt16LE(offset + 32);
    const local = bytes.readUInt32LE(offset + 42);
    const name = bytes.subarray(offset + 46, offset + 46 + nameLength).toString('utf8');
    if (name.startsWith('ppt/media/') && local + 30 <= bytes.length && bytes.readUInt32LE(local) === 0x04034b50) {
      const start = local + 30 + bytes.readUInt16LE(local + 26) + bytes.readUInt16LE(local + 28);
      if (start + length <= bytes.length) result.set(name, { start, length });
    }
    offset += 46 + nameLength + extraLength + commentLength;
  }
  return result;
}
async function assetRevision(zip, name) {
  let cache = assetRevisions.get(zip);
  if (!cache) { cache = new Map(); assetRevisions.set(zip, cache); }
  if (cache.has(name)) return cache.get(name);
  const value = (async () => {
    const index = indexFor(zip);
    if (index.bytes) {
      index.ranges ??= compressedAssetRanges(index.bytes);
      const range = index.ranges.get(name);
      if (range) return createHash('sha256').update(index.bytes.subarray(range.start, range.start + range.length)).digest('hex');
    }
    const file = zip.file(name);
    if (!file) return null;
    return createHash('sha256').update(await file.async('nodebuffer')).digest('hex');
  })();
  cache.set(name, value);
  return value;
}
const mimeTypes = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif',
  tif: 'image/tiff', tiff: 'image/tiff', bmp: 'image/bmp', svg: 'image/svg+xml', emf: 'image/emf',
  wmf: 'image/wmf', mp4: 'video/mp4', mov: 'video/quicktime', avi: 'video/x-msvideo' };

function xml(text) {
  if (/<!DOCTYPE|<!ENTITY/i.test(text)) throw new Error('演示文稿 XML 包含不允许的外部实体。');
  const doc = new DOMParser({ onError(level) {
    if (level !== 'warning') throw new Error('演示文稿 XML 格式无效。');
  } }).parseFromString(text, 'application/xml');
  if (!doc.documentElement || doc.getElementsByTagName('parsererror').length) throw new Error('演示文稿 XML 格式无效。');
  return doc;
}
function children(node, name) {
  const result = [];
  for (let child = node?.firstChild; child; child = child.nextSibling) {
    if (child.nodeType === 1 && (!name || child.localName === name)) result.push(child);
  }
  return result;
}
function child(node, name) { return children(node, name)[0] || null; }
function descendants(node, name) {
  const result = [];
  for (const item of children(node)) {
    if (item.localName === name) result.push(item);
    result.push(...descendants(item, name));
  }
  return result;
}
function first(node, name) { return descendants(node, name)[0] || null; }
function attr(node, key, fallback = '') { return node?.hasAttribute(key) ? node.getAttribute(key) : fallback; }
function number(node, key, fallback = 0) { const value = Number(attr(node, key, fallback)); return Number.isFinite(value) ? value : fallback; }
function bool(node, key, fallback = false) { const value = attr(node, key, ''); return value === '' ? fallback : value === '1' || value === 'true'; }
function relationPath(owner, target) {
  const result = target.startsWith('/') ? target.slice(1) : path.posix.normalize(path.posix.join(path.posix.dirname(owner), target));
  if (result.startsWith('../') || result.includes('\\')) throw new Error('演示文稿资源路径无效。');
  return result;
}
function relsPath(owner) { return path.posix.join(path.posix.dirname(owner), '_rels', path.posix.basename(owner) + '.rels'); }
async function document(zip, name) {
  const file = zip.file(name);
  return file ? xml(await file.async('string')) : null;
}
async function relationships(zip, owner) {
  const cache = cachesFor(zip).relationships;
  let pending = cache.get(owner);
  if (!pending) {
    pending = (async () => {
      const doc = await document(zip, relsPath(owner));
      const result = new Map();
      for (const node of children(doc?.documentElement, 'Relationship')) {
        result.set(attr(node, 'Id'), { type: attr(node, 'Type'), external: attr(node, 'TargetMode') === 'External',
          target: attr(node, 'Target'), path: attr(node, 'TargetMode') === 'External' ? null : relationPath(owner, attr(node, 'Target')) });
      }
      return result;
    })();
    cache.set(owner, pending);
  }
  return pending;
}
const emptyTheme = () => ({ colors: {}, fonts: {} });
function themeFor(zip, slideRels) {
  const layout = [...slideRels.values()].find(item => item.type.endsWith('/slideLayout'));
  if (!layout?.path) return Promise.resolve(emptyTheme());
  const cache = cachesFor(zip).themes;
  let pending = cache.get(layout.path);
  if (!pending) { pending = resolveTheme(zip, layout.path); cache.set(layout.path, pending); }
  return pending;
}
async function resolveTheme(zip, layoutPath) {
  const layoutRels = await relationships(zip, layoutPath);
  const master = [...layoutRels.values()].find(item => item.type.endsWith('/slideMaster'));
  if (!master?.path) return emptyTheme();
  const masterRels = await relationships(zip, master.path);
  const theme = [...masterRels.values()].find(item => item.type.endsWith('/theme'));
  if (!theme?.path) return emptyTheme();
  const doc = await document(zip, theme.path);
  const colors = {};
  for (const item of children(first(doc, 'clrScheme'))) {
    const color = children(item)[0];
    colors[item.localName] = '#' + (attr(color, 'val') || attr(color, 'lastClr') || '000000');
  }
  const fonts = {};
  for (const kind of ['majorFont', 'minorFont']) {
    const container = first(doc, kind);
    fonts[kind] = { latin: attr(child(container, 'latin'), 'typeface', 'Arial'),
      eastAsian: attr(child(container, 'ea'), 'typeface'), complexScript: attr(child(container, 'cs'), 'typeface') };
  }
  return { colors, fonts };
}
function color(node, theme, fallback = '#000000') {
  if (!node) return fallback;
  const own = ['srgbClr', 'sysClr', 'schemeClr', 'prstClr', 'scrgbClr'].includes(node.localName) ? node
    : children(node).find(item => ['srgbClr', 'sysClr', 'schemeClr', 'prstClr', 'scrgbClr'].includes(item.localName));
  if (!own) return fallback;
  let value;
  if (own.localName === 'schemeClr') value = theme.colors[attr(own, 'val')] || fallback;
  else if (own.localName === 'sysClr') value = '#' + attr(own, 'lastClr', '000000');
  else if (own.localName === 'srgbClr') value = '#' + attr(own, 'val', '000000');
  else if (own.localName === 'scrgbClr') {
    value = '#' + ['r', 'g', 'b'].map(key => Math.round(Math.min(1, Math.max(0, number(own, key) / 100000)) * 255).toString(16).padStart(2, '0')).join('');
  } else value = attr(own, 'val', fallback);
  if (!/^#[0-9a-f]{6}$/i.test(value)) return value;
  const rgb = [1, 3, 5].map(offset => parseInt(value.slice(offset, offset + 2), 16));
  for (const modifier of children(own)) {
    const factor = number(modifier, 'val') / 100000;
    if (modifier.localName === 'tint') rgb.forEach((v, i) => { rgb[i] = v + (255 - v) * factor; });
    if (modifier.localName === 'shade' || modifier.localName === 'lumMod') rgb.forEach((v, i) => { rgb[i] = v * factor; });
    if (modifier.localName === 'lumOff') rgb.forEach((v, i) => { rgb[i] = v + 255 * factor; });
  }
  return '#' + rgb.map(value => Math.round(Math.min(255, Math.max(0, value))).toString(16).padStart(2, '0')).join('').toUpperCase();
}
function transparency(node) {
  const alpha = first(node, 'alpha');
  return alpha ? 1 - number(alpha, 'val', 100000) / 100000 : 0;
}
function fill(node, theme) {
  if (child(node, 'noFill')) return { type: 'none', transparency: 1 };
  const solid = child(node, 'solidFill');
  if (solid) return { type: 'solid', color: color(solid, theme), transparency: transparency(solid) };
  const gradient = child(node, 'gradFill');
  if (gradient) return { type: 'gradient', stops: descendants(gradient, 'gs').map(stop => ({
    position: number(stop, 'pos') / 100000, color: color(stop, theme), transparency: transparency(stop) })),
    angle: number(child(gradient, 'lin'), 'ang') / 60000, raw: serializer.serializeToString(gradient) };
  const picture = child(node, 'blipFill');
  if (picture) return { type: 'picture', raw: serializer.serializeToString(picture) };
  return { type: 'none', transparency: 1 };
}
function line(node, theme) {
  if (!node) return { visible: false, width: 0 };
  return { visible: !child(node, 'noFill'), width: number(node, 'w', EMU) / EMU,
    color: color(child(node, 'solidFill'), theme), transparency: transparency(child(node, 'solidFill')),
    dashStyle: attr(child(node, 'prstDash'), 'val', 'solid'),
    beginArrow: attr(child(node, 'headEnd'), 'type', 'none'), endArrow: attr(child(node, 'tailEnd'), 'type', 'none'),
    raw: serializer.serializeToString(node) };
}
function font(node, theme, inherited = {}) {
  const result = { ...inherited };
  if (!node) return result;
  const latin = child(node, 'latin'), ea = child(node, 'ea'), cs = child(node, 'cs');
  const resolve = name => name === '+mn-lt' ? theme.fonts.minorFont?.latin || 'Arial'
    : name === '+mj-lt' ? theme.fonts.majorFont?.latin || 'Arial' : name;
  if (latin) result.name = result.nameAscii = resolve(attr(latin, 'typeface'));
  if (ea) result.nameFarEast = attr(ea, 'typeface');
  if (cs) result.nameComplexScript = attr(cs, 'typeface');
  if (node.hasAttribute('sz')) result.size = number(node, 'sz') / 100;
  if (node.hasAttribute('b')) result.bold = bool(node, 'b');
  if (node.hasAttribute('i')) result.italic = bool(node, 'i');
  if (node.hasAttribute('u')) result.underline = attr(node, 'u') !== 'none';
  if (node.hasAttribute('strike')) result.strikethrough = attr(node, 'strike') !== 'noStrike';
  if (node.hasAttribute('baseline')) {
    result.baselineOffset = number(node, 'baseline') / 100000;
    result.superscript = result.baselineOffset > 0; result.subscript = result.baselineOffset < 0;
  }
  const solid = child(node, 'solidFill');
  if (solid) { result.color = color(solid, theme); result.transparency = transparency(solid); }
  const highlight = child(node, 'highlight');
  if (highlight) result.highlight = color(highlight, theme);
  const effect = child(node, 'effectLst');
  if (effect) {
    const glow = child(effect, 'glow');
    if (glow) result.glow = { color: color(glow, theme), radius: number(glow, 'rad') / EMU, transparency: transparency(glow) };
    result.shadow = !!(child(effect, 'outerShdw') || child(effect, 'innerShdw'));
    const reflection = child(effect, 'reflection');
    if (reflection) result.reflection = { type: 'custom', raw: serializer.serializeToString(reflection) };
  }
  if (child(node, 'ln')) result.outline = line(child(node, 'ln'), theme);
  const raw = node.cloneNode(true);
  for (const item of descendants(raw).filter(item => ['schemeClr', 'sysClr', 'scrgbClr'].includes(item.localName))) {
    const resolved = color(item, theme);
    if (!/^#[0-9a-f]{6}$/i.test(resolved)) continue;
    const replacement = raw.ownerDocument.createElementNS(A, 'a:srgbClr');
    replacement.setAttribute('val', resolved.slice(1));
    for (const modifier of children(item).filter(item => item.localName.startsWith('alpha'))) replacement.appendChild(modifier.cloneNode(true));
    item.parentNode.replaceChild(replacement, item);
  }
  result.raw = serializer.serializeToString(raw);
  return result;
}
function parseText(node, theme) {
  const body = child(node, 'txBody');
  if (!body) return {};
  const bodyPr = child(body, 'bodyPr');
  const baseFont = { name: theme.fonts.minorFont?.latin || 'Arial', size: 18, color: '#000000',
    bold: false, italic: false, underline: false, strikethrough: false };
  const paragraphs = children(body, 'p').map(paragraph => {
    const pr = child(paragraph, 'pPr');
    const defaults = font(child(pr, 'defRPr'), theme, baseFont);
    const runs = children(paragraph).flatMap(run => {
      if (run.localName === 'r' || run.localName === 'fld') {
        const properties = child(run, 'rPr');
        const hyperlink = child(properties, 'hlinkClick');
        return [{ text: child(run, 't')?.textContent || '', font: font(properties, theme, defaults),
          ...(hyperlink ? { hyperlinkId: hyperlink.getAttributeNS(R, 'id') } : {}) }];
      }
      if (run.localName === 'br') return [{ text: '\n', font: defaults }];
      if (run.localName === 'm') return [{ text: descendants(run, 't').map(item => item.textContent).join(''),
        font: defaults, math: serializer.serializeToString(run) }];
      return [];
    });
    const bulletChar = child(pr, 'buChar'), bulletNumber = child(pr, 'buAutoNum');
    return { runs, level: number(pr, 'lvl'), alignment: ({ l: 'left', ctr: 'center', r: 'right', just: 'justify' })[attr(pr, 'algn')] || attr(pr, 'algn', 'left'),
      marginLeft: number(pr, 'marL') / EMU, indent: number(pr, 'indent') / EMU,
      spaceBefore: number(child(child(pr, 'spcBef'), 'spcPts'), 'val') / 100,
      spaceAfter: number(child(child(pr, 'spcAft'), 'spcPts'), 'val') / 100,
      bullet: bulletChar ? { type: 'character', character: attr(bulletChar, 'char') }
        : bulletNumber ? { type: 'numbered', style: attr(bulletNumber, 'type'), startAt: number(bulletNumber, 'startAt', 1) } : null };
  });
  const allRuns = paragraphs.flatMap((paragraph, index) => index ? [{ text: '\n', font: baseFont }, ...paragraph.runs] : paragraph.runs);
  return { text: allRuns.map(run => run.text).join(''), runs: allRuns, paragraphs,
    font: allRuns.find(run => run.text.length)?.font || font(first(body, 'defRPr'), theme, baseFont),
    textFrame: { wordWrap: attr(bodyPr, 'wrap', 'square') !== 'none',
      autoSize: child(bodyPr, 'spAutoFit') ? 'shapeToFitText' : child(bodyPr, 'normAutofit') ? 'textToFitShape' : 'none',
      marginLeft: number(bodyPr, 'lIns', 91440) / EMU, marginRight: number(bodyPr, 'rIns', 91440) / EMU,
      marginTop: number(bodyPr, 'tIns', 45720) / EMU, marginBottom: number(bodyPr, 'bIns', 45720) / EMU,
      verticalAnchor: ({ t: 'top', ctr: 'middle', b: 'bottom' })[attr(bodyPr, 'anchor', 't')] || 'top' } };
}
async function tagsFor(zip, node, rels) {
  const result = {};
  const cache = cachesFor(zip).tags;
  for (const reference of descendants(node, 'tags')) {
    const target = rels.get(reference.getAttributeNS(R, 'id'));
    if (!target?.path) continue;
    let pending = cache.get(target.path);
    if (!pending) { pending = document(zip, target.path); cache.set(target.path, pending); }
    const doc = await pending;
    for (const item of descendants(doc, 'tag')) result[attr(item, 'name')] = attr(item, 'val');
  }
  return result;
}
function transformFor(node) {
  const properties = child(node, 'spPr') || child(node, 'grpSpPr');
  const transform = child(properties, 'xfrm') || child(node, 'xfrm');
  const offset = child(transform, 'off'), extent = child(transform, 'ext');
  return { left: number(offset, 'x') / EMU, top: number(offset, 'y') / EMU,
    width: number(extent, 'cx') / EMU, height: number(extent, 'cy') / EMU,
    rotation: number(transform, 'rot') / 60000, flipH: bool(transform, 'flipH'), flipV: bool(transform, 'flipV'),
    childLeft: number(child(transform, 'chOff'), 'x') / EMU, childTop: number(child(transform, 'chOff'), 'y') / EMU,
    childWidth: number(child(transform, 'chExt'), 'cx') / EMU, childHeight: number(child(transform, 'chExt'), 'cy') / EMU };
}
async function parseShape(zip, node, rels, theme, slideIndex, shapePath, zOrder, parent = null) {
  node = resolvedShape(node);
  const nv = children(node).find(item => item.localName.startsWith('nv'));
  const idNode = first(nv, 'cNvPr');
  const id = String(attr(idNode, 'id'));
  const properties = child(node, 'spPr') || child(node, 'grpSpPr');
  const geometry = transformFor(node);
  const preset = attr(child(properties, 'prstGeom'), 'prst', 'rectangle');
  const media = first(nv, 'videoFile') || first(nv, 'audioFile') || first(nv, 'media');
  let type = node.localName === 'grpSp' ? 'group' : node.localName === 'pic' ? 'picture'
    : first(node, 'tbl') ? 'table' : media ? 'media' : child(node, 'txBody') ? 'text' : 'shape';
  if (media) type = 'media';
  const shapeFill = fill(properties, theme);
  const shapeLine = line(child(properties, 'ln'), theme);
  const effects = child(properties, 'effectLst');
  const item = { id, name: attr(idNode, 'name', 'Shape ' + id), type, ...geometry,
    alternativeText: attr(idNode, 'descr'), title: attr(idNode, 'title'), tags: await tagsFor(zip, nv, rels),
    zOrder, shapeType: preset, isPlaceholder: !!first(nv, 'ph'),
    ref: { slideIndex, shapePath }, fill: shapeFill, line: shapeLine,
    style: { fill: shapeFill, line: shapeLine,
      effects: effects ? serializer.serializeToString(effects) : null,
      raw: { properties: properties ? serializer.serializeToString(properties) : null,
        reference: child(node, 'style') ? serializer.serializeToString(child(node, 'style')) : null } },
    ...parseText(node, theme) };
  if (parent) item.parentId = parent.id;
  const blipFill = child(node, 'blipFill');
  if (blipFill) {
    const svg = first(blipFill, 'svgBlip'), blip = first(blipFill, 'blip');
    const embeddedId = svg?.getAttributeNS(R, 'embed') || blip?.getAttributeNS(R, 'embed');
    const linkedId = blip?.getAttributeNS(R, 'link');
    const relation = rels.get(embeddedId || linkedId);
    if (relation) item.image = { path: relation.path, mime: mimeTypes[path.posix.extname(relation.target).slice(1).toLowerCase()] || 'application/octet-stream',
      external: relation.external, source: relation.external ? relation.target : undefined,
      digest: relation.path ? await assetRevision(zip, relation.path) : undefined };
    const source = child(blipFill, 'srcRect');
    item.crop = { left: number(source, 'l') / 100000, top: number(source, 't') / 100000,
      right: number(source, 'r') / 100000, bottom: number(source, 'b') / 100000 };
  }
  if (media) {
    const relation = rels.get(media.getAttributeNS(R, 'link') || media.getAttributeNS(R, 'embed'));
    if (relation) item.media = { path: relation.path, source: relation.external ? relation.target : undefined,
      mime: mimeTypes[path.posix.extname(relation.target).slice(1).toLowerCase()] || 'application/octet-stream' };
  }
  if (type === 'group') {
    const nodes = shapeNodes(node);
    item.children = await Promise.all(nodes.map((childNode, index) => parseShape(zip, childNode, rels, theme,
      slideIndex, [...shapePath, index + 1], index, item)));
  }
  return item;
}
function resolvedShape(node) {
  if (node.localName !== 'AlternateContent') return node;
  const branch = child(node, 'Choice') || child(node, 'Fallback');
  return children(branch).find(item => ['sp', 'pic', 'graphicFrame', 'cxnSp', 'grpSp'].includes(item.localName));
}
function shapeNodes(parent) {
  return children(parent).filter(item => ['sp', 'pic', 'graphicFrame', 'cxnSp', 'grpSp'].includes(resolvedShape(item)?.localName));
}
function flatten(items) { return items.flatMap(item => [item, ...flatten(item.children || [])]); }
function bounds(items) {
  if (!items.length) return { left: 0, top: 0, width: 0, height: 0 };
  const left = Math.min(...items.map(item => item.left)), top = Math.min(...items.map(item => item.top));
  return { left, top, width: Math.max(...items.map(item => item.left + item.width)) - left,
    height: Math.max(...items.map(item => item.top + item.height)) - top };
}
async function openZip(zip) {
  const presentation = await document(zip, 'ppt/presentation.xml');
  if (!presentation) throw new Error('文件中没有 PowerPoint 演示文稿。');
  const rels = await relationships(zip, 'ppt/presentation.xml');
  const ids = children(first(presentation, 'sldIdLst'), 'sldId');
  const slides = ids.map((item, index) => ({ id: attr(item, 'id'), index, relation: item.getAttributeNS(R, 'id'),
    path: rels.get(item.getAttributeNS(R, 'id'))?.path, node: item }));
  if (slides.some(item => !item.path)) throw new Error('演示文稿幻灯片关系不完整。');
  const size = first(presentation, 'sldSz');
  return { zip, presentation, slides, width: number(size, 'cx') / EMU, height: number(size, 'cy') / EMU };
}
async function open(pptx) {
  if (typeof pptx !== 'string' || !pptx.length) throw new Error('演示文稿数据为空。');
  const bytes = Buffer.from(pptx, 'base64');
  const zip = await JSZip.loadAsync(bytes);
  assetIndexes.set(zip, { bytes, ranges: null });
  return openZip(zip);
}
async function refreshPackage(pkg) {
  const refreshed = await openZip(pkg.zip);
  pkg.presentation = refreshed.presentation;
  pkg.slides = refreshed.slides;
  pkg.width = refreshed.width;
  pkg.height = refreshed.height;
  return pkg;
}
async function parsePackage(pkg, options = {}) {
  const requested = options.indices ? new Set(options.indices) : null;
  const slides = [];
  for (const slide of pkg.slides) {
    if (requested && !requested.has(slide.index)) continue;
    const doc = await document(pkg.zip, slide.path);
    const tree = first(doc, 'spTree'), rels = await relationships(pkg.zip, slide.path);
    const theme = await themeFor(pkg.zip, rels);
    const nodes = shapeNodes(tree);
    slides.push({ id: slide.id, index: slide.index, name: attr(first(doc, 'cSld'), 'name'),
      shapes: await Promise.all(nodes.map((node, index) => parseShape(pkg.zip, node, rels, theme, slide.index, [index + 1], index))) });
  }
  return { width: pkg.width, height: pkg.height, slides };
}
export async function parsePresentation(pptx, options = {}) {
  return parsePackage(await open(pptx), options);
}
async function selectPackage(pkg, indices) {
  const wanted = [...new Set(indices)];
  if (!wanted.length || wanted.some(index => !Number.isInteger(index) || index < 0 || index >= pkg.slides.length)) throw new Error('导出页码超出演示文稿范围。');
  const list = first(pkg.presentation, 'sldIdLst');
  for (const node of children(list)) list.removeChild(node);
  for (const index of wanted) list.appendChild(pkg.slides[index].node);
  pkg.zip.file('ppt/presentation.xml', serializer.serializeToString(pkg.presentation));
  await prunePresentation(pkg.zip, { indices: wanted });
  invalidatePackage(pkg.zip);
  return pkg.zip;
}
export async function selectSlides(pptx, indices) {
  const pkg = await open(pptx);
  return (await selectPackage(pkg, indices)).generateAsync({ type: 'base64', compression: 'DEFLATE' });
}
export async function extractImages(pptx, { slideIndex = 0, ids, original = true } = {}) {
  const pkg = await open(pptx);
  const model = await parsePackage(pkg, { indices: [slideIndex] });
  const wanted = ids ? new Set(ids.map(String)) : null;
  const output = [];
  for (const shape of flatten(model.slides[0]?.shapes || [])) {
    if (wanted && !wanted.has(shape.id)) continue;
    const asset = shape.image || shape.media;
    if (!asset?.path || asset.external) continue;
    const file = pkg.zip.file(asset.path);
    if (!file) throw new Error('找不到内嵌资源：' + shape.name);
    const bytes = await file.async('nodebuffer');
    let metadata = {};
    if (shape.image) metadata = await readImageDimensions(bytes, asset.mime);
    output.push({ id: shape.id, data: bytes.toString('base64'), mime: asset.mime,
      width: metadata.width, height: metadata.height, source: asset.source,
      name: shape.name, alternativeText: shape.alternativeText, tags: shape.tags, crop: shape.crop,
      original, media: shape.type === 'media' });
  }
  return output;
}
async function extractPackage(pptx, { slideIndex = 0, ids } = {}, refresh = false) {
  if (!ids?.length) throw new Error('请先选择要保存的对象。');
  const pkg = await open(pptx);
  const wantedList = ids.map(String);
  const wanted = new Set(wantedList);
  await selectPackage(pkg, [slideIndex]);
  await refreshPackage(pkg);
  const model = await parsePackage(pkg);
  const available = new Set(flatten(model.slides[0].shapes).map(item => item.id));
  if ([...wanted].some(id => !available.has(id))) throw new Error('选择的对象已改变，请重新选择。');
  const slidePath = pkg.slides[0].path;
  const doc = await document(pkg.zip, slidePath);
  const tree = first(doc, 'spTree');
  function prune(parent) {
    for (const node of shapeNodes(parent)) {
      const actual = resolvedShape(node);
      const nv = children(actual).find(item => item.localName.startsWith('nv'));
      const id = attr(first(nv, 'cNvPr'), 'id');
      if (wanted.has(id)) continue;
      if (actual.localName === 'grpSp') {
        prune(actual);
        if (shapeNodes(actual).length) continue;
      }
      parent.removeChild(node);
    }
  }
  prune(tree);
  pkg.zip.file(slidePath, serializer.serializeToString(doc));
  await prunePresentation(pkg.zip, { indices: [0] });
  invalidatePackage(pkg.zip);
  if (refresh) await refreshPackage(pkg);
  const chosen = flatten(model.slides[0].shapes).filter(item => wanted.has(item.id));
  return { pkg, model, wantedList, chosen, bounds: bounds(chosen), name: chosen.map(item => item.name).join(', ') };
}
export async function extractFragment(pptx, options = {}) {
  const fragment = await extractPackage(pptx, options);
  return { pptx: await fragment.pkg.zip.generateAsync({ type: 'base64', compression: 'DEFLATE' }),
    shapeIds: fragment.wantedList, bounds: fragment.bounds, name: fragment.name };
}
export async function imageFragment(pptx, { slideIndex = 0, id, original = true } = {}) {
  const fragment = await extractPackage(pptx, { slideIndex, ids: [String(id)] }, original);
  if (!original) return { pptx: await fragment.pkg.zip.generateAsync({ type: 'base64', compression: 'DEFLATE' }),
    shapeIds: fragment.wantedList, bounds: fragment.bounds, name: fragment.name };
  const { pkg, model } = fragment;
  const shape = flatten(model.slides[0].shapes).find(item => item.id === String(id));
  if (!shape?.image?.path || shape.image.external) throw new Error('请选择包含内嵌原图的图片对象。');
  const bytes = await pkg.zip.file(shape.image.path).async('nodebuffer');
  const image = await readImageDimensions(bytes, shape.image.mime);
  const width = image.width * .75, height = image.height * .75;
  const doc = await document(pkg.zip, pkg.slides[0].path);
  const tree = first(doc, 'spTree');
  let picture;
  function findPicture(parent) {
    for (const node of shapeNodes(parent)) {
      const actual = resolvedShape(node);
      const nv = children(actual).find(item => item.localName.startsWith('nv'));
      if (attr(first(nv, 'cNvPr'), 'id') === String(id)) picture = actual;
      if (actual.localName === 'grpSp') findPicture(actual);
    }
  }
  findPicture(tree);
  if (!picture || picture.localName !== 'pic') throw new Error('无法提取原生图片节点。');
  const isolated = picture.cloneNode(true);
  for (const node of shapeNodes(tree)) tree.removeChild(node);
  tree.appendChild(isolated);
  const properties = child(isolated, 'spPr');
  let transform = child(properties, 'xfrm');
  if (!transform) { transform = doc.createElementNS(A, 'a:xfrm'); properties.insertBefore(transform, properties.firstChild); }
  for (const key of ['rot', 'flipH', 'flipV']) transform.removeAttribute(key);
  for (const [name, values] of [['off', { x: 0, y: 0 }], ['ext', { cx: Math.round(width * EMU), cy: Math.round(height * EMU) }]]) {
    let node = child(transform, name);
    if (!node) { node = doc.createElementNS(A, 'a:' + name); transform.appendChild(node); }
    for (const [key, value] of Object.entries(values)) node.setAttribute(key, String(value));
  }
  for (const name of ['ln', 'effectLst', 'effectDag']) {
    const node = child(properties, name);
    if (node) properties.removeChild(node);
  }
  const blipFill = child(isolated, 'blipFill');
  for (const name of ['srcRect', 'tile']) {
    const node = child(blipFill, name);
    if (node) blipFill.removeChild(node);
  }
  if (!child(blipFill, 'stretch')) {
    const stretch = doc.createElementNS(A, 'a:stretch');
    stretch.appendChild(doc.createElementNS(A, 'a:fillRect'));
    blipFill.appendChild(stretch);
  }
  const size = first(pkg.presentation, 'sldSz');
  size.setAttribute('cx', String(Math.round(width * EMU)));
  size.setAttribute('cy', String(Math.round(height * EMU)));
  pkg.zip.file(pkg.slides[0].path, serializer.serializeToString(doc));
  pkg.zip.file('ppt/presentation.xml', serializer.serializeToString(pkg.presentation));
  await prunePresentation(pkg.zip, { indices: [0] });
  invalidatePackage(pkg.zip);
  return { pptx: await pkg.zip.generateAsync({ type: 'base64', compression: 'DEFLATE' }),
    shapeIds: [String(id)], bounds: { left: 0, top: 0, width, height }, image };
}
export async function replacePicture(pptx, { slideIndex = 0, id, rootId = id, rootName, data, mime = 'image/png', preserveCrop = true } = {}) {
  const fragment = await extractPackage(pptx, { slideIndex, ids: [String(rootId)] });
  const { pkg } = fragment;
  const doc = await document(pkg.zip, pkg.slides[0].path);
  const picture = descendants(doc, 'pic').find(node => attr(first(child(node, 'nvPicPr'), 'cNvPr'), 'id') === String(id));
  if (rootName) for (const node of descendants(doc, 'cNvPr')) {
    if (attr(node, 'id') === String(rootId)) node.setAttribute('name', String(rootName));
  }
  if (!picture) throw new Error('替换目标不是原生图片。');
  if (typeof data !== 'string' || !data) throw new Error('缺少替换图片的数据。');
  const bytes = Buffer.from(data, 'base64');
  await readImageDimensions(bytes, mime);
  const extension = Object.entries(mimeTypes).find(([, value]) => value === mime)?.[0]
    || ({ 'image/x-emf': 'emf', 'image/x-wmf': 'wmf', 'image/webp': 'webp' })[mime];
  if (!extension) throw new Error('不支持的替换图片类型：' + mime);
  const relPath = relsPath(pkg.slides[0].path);
  const rels = await document(pkg.zip, relPath);
  const types = await document(pkg.zip, '[Content_Types].xml');
  function addImage(buffer, suffix, contentType) {
    const digest = createHash('sha256').update(buffer).digest('hex');
    const filename = 'slidesci-' + digest + '.' + suffix;
    const rid = 'rIdSlideSCI' + digest;
    pkg.zip.file('ppt/media/' + filename, buffer);
    if (!children(rels.documentElement).some(node => attr(node, 'Id') === rid)) {
      const relationship = rels.createElementNS('http://schemas.openxmlformats.org/package/2006/relationships', 'Relationship');
      relationship.setAttribute('Id', rid);
      relationship.setAttribute('Type', R + '/image');
      relationship.setAttribute('Target', '../media/' + filename);
      rels.documentElement.appendChild(relationship);
    }
    if (!children(types.documentElement).some(node => attr(node, 'Extension') === suffix)) {
      const type = types.createElementNS('http://schemas.openxmlformats.org/package/2006/content-types', 'Default');
      type.setAttribute('Extension', suffix); type.setAttribute('ContentType', contentType);
      types.documentElement.appendChild(type);
    }
    return rid;
  }
  const fill = child(picture, 'blipFill');
  const blip = child(fill, 'blip');
  const oldExtension = child(blip, 'extLst');
  if (oldExtension) blip.removeChild(oldExtension);
  blip.removeAttributeNS(R, 'link');
  let primary = addImage(bytes, extension, mime);
  if (mime === 'image/svg+xml') {
    const { default: sharp } = await import('sharp');
    primary = addImage(await sharp(bytes, { density: 144 }).png().toBuffer(), 'png', 'image/png');
    const list = doc.createElementNS(A, 'a:extLst'), ext = doc.createElementNS(A, 'a:ext');
    ext.setAttribute('uri', '{96DAC541-7B7A-43D3-8B79-37D633B846F1}');
    const svg = doc.createElementNS('http://schemas.microsoft.com/office/drawing/2016/SVG/main', 'asvg:svgBlip');
    svg.setAttributeNS(R, 'r:embed', addImage(bytes, extension, mime));
    ext.appendChild(svg); list.appendChild(ext); blip.appendChild(list);
  }
  blip.setAttributeNS(R, 'r:embed', primary);
  const obsoleteFill = child(child(picture, 'spPr'), 'blipFill');
  if (obsoleteFill) obsoleteFill.parentNode.removeChild(obsoleteFill);
  if (!preserveCrop) { const crop = child(fill, 'srcRect'); if (crop) fill.removeChild(crop); }
  pkg.zip.file(pkg.slides[0].path, serializer.serializeToString(doc));
  pkg.zip.file(relPath, serializer.serializeToString(rels));
  pkg.zip.file('[Content_Types].xml', serializer.serializeToString(types));
  await prunePresentation(pkg.zip, { indices: [0] });
  invalidatePackage(pkg.zip);
  return { pptx: await pkg.zip.generateAsync({ type: 'base64', compression: 'DEFLATE' }),
    shapeIds: fragment.wantedList, bounds: fragment.bounds, name: fragment.name };
}
export async function formatText(pptx, { slideIndex = 0, rootId, rootName, updates } = {}) {
  if (!Array.isArray(updates) || !updates.length) throw new Error('缺少文字格式更新。');
  const fragment = await extractPackage(pptx, { slideIndex, ids: [String(rootId)] });
  const { pkg } = fragment;
  const doc = await document(pkg.zip, pkg.slides[0].path);
  for (const update of updates) applyTextFormatting(doc, update.id, update.properties);
  if (rootName) {
    const nv = descendants(doc, 'cNvPr').find(node => attr(node, 'id') === String(rootId));
    if (!nv) throw new Error('文字对象所属组合已改变。');
    nv.setAttribute('name', String(rootName));
  }
  pkg.zip.file(pkg.slides[0].path, serializer.serializeToString(doc));
  await prunePresentation(pkg.zip, { indices: [0] });
  invalidatePackage(pkg.zip);
  return { pptx: await pkg.zip.generateAsync({ type: 'base64', compression: 'DEFLATE' }),
    shapeIds: fragment.wantedList, bounds: fragment.bounds, name: fragment.name };
}

export const pptx = { parsePresentation, selectSlides, extractImages, extractFragment, imageFragment, replacePicture, formatText };
