import test from 'node:test';
import assert from 'node:assert/strict';
import path from 'node:path';
import JSZip from 'jszip';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { prunePresentation } from '../lib/pptx-prune.mjs';

const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG_RELATIONSHIPS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CONTENT_TYPES = 'http://schemas.openxmlformats.org/package/2006/content-types';
const VT = 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes';
const serializer = new XMLSerializer();

const T = {
  officeDocument: R + '/officeDocument',
  coreProperties: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
  extendedProperties: R + '/extended-properties',
  thumbnail: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail',
  slide: R + '/slide',
  slideLayout: R + '/slideLayout',
  slideMaster: R + '/slideMaster',
  theme: R + '/theme',
  notesSlide: R + '/notesSlide',
  image: R + '/image',
  chart: R + '/chart',
  media: R + '/media',
  package: R + '/package',
  oleObject: R + '/oleObject',
  tags: R + '/tags',
  hyperlink: R + '/hyperlink',
  presProps: R + '/presProps',
  viewProps: R + '/viewProps',
  tableStyles: R + '/tableStyles',
  custom: 'http://example.com/relationships/custom'
};

const png = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9Zl1sAAAAASUVORK5CYII=', 'base64');

function escape(value) {
  return String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
function parseXml(text) {
  const doc = new DOMParser({ onError(level) {
    if (level !== 'warning') throw new Error('invalid fixture XML');
  } }).parseFromString(text, 'application/xml');
  assert.ok(doc.documentElement);
  return doc;
}
function children(node, name) {
  const result = [];
  for (let item = node?.firstChild; item; item = item.nextSibling) {
    if (item.nodeType === 1 && (!name || item.localName === name)) result.push(item);
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
function attr(node, name, fallback = '') { return node?.hasAttribute(name) ? node.getAttribute(name) : fallback; }

function relationships(entries) {
  return `<Relationships xmlns="${PKG_RELATIONSHIPS}">${entries.map(([id, type, target, external]) => {
    const mode = external ? ' TargetMode="External"' : '';
    return `<Relationship Id="${id}" Type="${type}" Target="${escape(target)}"${mode}/>`;
  }).join('')}</Relationships>`;
}

function contentTypes() {
  const overrides = [
    ['/ppt/presentation.xml', 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml'],
    ['/ppt/slides/slide1.xml', 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'],
    ['/ppt/slides/slide2.xml', 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'],
    ['/ppt/slides/slide3.xml', 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'],
    ['/ppt/slideLayouts/slideLayout1.xml', 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml'],
    ['/ppt/slideMasters/slideMaster1.xml', 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml'],
    ['/ppt/theme/theme1.xml', 'application/vnd.openxmlformats-officedocument.theme+xml'],
    ['/ppt/notesSlides/notesSlide1.xml', 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml'],
    ['/ppt/notesSlides/notesSlide2.xml', 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml'],
    ['/ppt/notesSlides/notesSlide3.xml', 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml'],
    ['/ppt/charts/chart1.xml', 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'],
    ['/ppt/charts/chart3.xml', 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'],
    ['/ppt/charts/chart-orphan.xml', 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'],
    ['/ppt/tags/tag1.xml', 'application/vnd.openxmlformats-officedocument.presentationml.tags+xml'],
    ['/ppt/tags/tags-orphan.xml', 'application/vnd.openxmlformats-officedocument.presentationml.tags+xml'],
    ['/ppt/presProps.xml', 'application/vnd.openxmlformats-officedocument.presentationml.presProps+xml'],
    ['/ppt/viewProps.xml', 'application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml'],
    ['/ppt/tableStyles.xml', 'application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml'],
    ['/ppt/custom/thing.xml', 'application/xml'],
    ['/ppt/embeddings/ole1.bin', 'application/vnd.openxmlformats-officedocument.oleObject'],
    ['/ppt/embeddings/ole-orphan.bin', 'application/vnd.openxmlformats-officedocument.oleObject'],
    ['/ppt/embeddings/book1.xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'],
    ['/docProps/core.xml', 'application/vnd.openxmlformats-package.core-properties+xml'],
    ['/docProps/app.xml', 'application/vnd.openxmlformats-officedocument.extended-properties+xml']
  ];
  return `<Types xmlns="${CONTENT_TYPES}">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Default Extension="png" ContentType="image/png"/>
    <Default Extension="svg" ContentType="image/svg+xml"/>
    <Default Extension="jpeg" ContentType="image/jpeg"/>
    <Default Extension="mp4" ContentType="video/mp4"/>
    ${overrides.map(([part, type]) => `<Override PartName="${part}" ContentType="${type}"/>`).join('')}
  </Types>`;
}

const GROUP_PROPS = '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';
const ROOT_SHAPE = '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>';
function shapeTree(content = '') { return `<p:spTree>${ROOT_SHAPE}${GROUP_PROPS}${content}</p:spTree>`; }

function textShape(id, name, text, properties = '') {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${escape(name)}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US">${properties}</a:rPr>` +
    `<a:t>${escape(text)}</a:t></a:r></a:p></p:txBody></p:sp>`;
}
function picture(id, name, relationshipId, extra = '') {
  const blip = extra ? `<a:blip r:embed="${relationshipId}">${extra}</a:blip>` : `<a:blip r:embed="${relationshipId}"/>`;
  return `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${escape(name)}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
    `<p:blipFill>${blip}<a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr/></p:pic>`;
}
function chartShape(id, name, relationshipId) {
  return `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="${escape(name)}"/>` +
    '<p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>' +
    '<p:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></p:xfrm>' +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
    `<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="${relationshipId}"/>` +
    '</a:graphicData></a:graphic></p:graphicFrame>';
}
const SVG_BLIP = '<a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">' +
  `<asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rId6"/>` +
  '</a:ext></a:extLst>';

function slideXml(name, tree) {
  return `<p:sld xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}"><p:cSld name="${escape(name)}">${tree}</p:cSld></p:sld>`;
}
function notesSlideXml() {
  return `<p:notes xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}"><p:cSld>${shapeTree()}</p:cSld></p:notes>`;
}
function slideLayoutXml() {
  return `<p:sldLayout xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}" type="title" preserve="1">` +
    `<p:cSld name="Title">${shapeTree()}</p:cSld></p:sldLayout>`;
}
function slideMasterXml() {
  return `<p:sldMaster xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}">` +
    `<p:cSld name="Master">${shapeTree()}</p:cSld>` +
    '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>';
}
const themeXml = `<a:theme xmlns:a="${A}" name="Office Theme"><a:themeElements>` +
  '<a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>' +
  '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1></a:clrScheme>' +
  '<a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont>' +
  '<a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme>' +
  '<a:fmtScheme name="Office"><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme>' +
  '</a:themeElements></a:theme>';
const chartXml = `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="${R}">` +
  '<c:chart><c:plotArea/></c:chart><c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData>' +
  '</c:chartSpace>';
const plainChartXml = '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
  '<c:chart><c:plotArea/></c:chart></c:chartSpace>';

function presentationXml() {
  return `<p:presentation xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}">` +
    '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId4"/></p:sldMasterIdLst>' +
    '<p:sldIdLst><p:sldId id="256" r:id="rId1"/><p:sldId id="257" r:id="rId2"/><p:sldId id="258" r:id="rId3"/></p:sldIdLst>' +
    '<p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/>' +
    '<p:custShowLst>' +
    '<p:custShow name="Mixed" id="0"><p:sldLst><p:sld r:id="rId3"/><p:sld r:id="rId2"/></p:sldLst></p:custShow>' +
    '<p:custShow name="Gone" id="1"><p:sldLst><p:sld r:id="rId2"/></p:sldLst></p:custShow>' +
    '</p:custShowLst></p:presentation>';
}

function appXml() {
  return `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="${VT}">` +
    '<Application>Microsoft Office PowerPoint</Application><Slides>3</Slides><Notes>3</Notes>' +
    '<HeadingPairs><vt:vector size="4" baseType="variant">' +
    '<vt:variant><vt:lpstr>Theme</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant>' +
    '<vt:variant><vt:lpstr>Slide Titles</vt:lpstr></vt:variant><vt:variant><vt:i4>3</vt:i4></vt:variant>' +
    '</vt:vector></HeadingPairs>' +
    '<TitlesOfParts><vt:vector size="4" baseType="lpstr">' +
    '<vt:lpstr>Office Theme</vt:lpstr><vt:lpstr>Alpha</vt:lpstr><vt:lpstr>Beta</vt:lpstr><vt:lpstr>Gamma</vt:lpstr>' +
    '</vt:vector></TitlesOfParts><Company>SlideSCI</Company></Properties>';
}

function deckFiles() {
  const rootRels = relationships([
    ['rId1', T.officeDocument, 'ppt/presentation.xml'],
    ['rId2', T.coreProperties, 'docProps/core.xml'],
    ['rId3', T.extendedProperties, 'docProps/app.xml'],
    ['rId4', T.thumbnail, 'docProps/thumbnail.jpeg']
  ]);
  const presentationRels = relationships([
    ['rId1', T.slide, 'slides/slide1.xml'],
    ['rId2', T.slide, 'slides/slide2.xml'],
    ['rId3', T.slide, 'slides/slide3.xml'],
    ['rId4', T.slideMaster, 'slideMasters/slideMaster1.xml'],
    ['rId5', T.presProps, 'presProps.xml'],
    ['rId6', T.viewProps, 'viewProps.xml'],
    ['rId7', T.tableStyles, 'tableStyles.xml'],
    ['rId8', T.custom, 'custom/thing.xml']
  ]);
  const slide1Rels = relationships([
    ['rId1', T.slideLayout, '../slideLayouts/slideLayout1.xml'],
    ['rId2', T.notesSlide, '../notesSlides/notesSlide1.xml'],
    ['rId3', T.image, '../media/shared.png'],
    ['rId4', T.image, '../media/one.png'],
    ['rId5', T.image, '../media/svg-fallback.png'],
    ['rId6', T.image, '../media/vector.svg'],
    ['rId7', T.chart, '../charts/chart1.xml'],
    ['rId8', T.oleObject, '../embeddings/ole1.bin'],
    ['rId9', T.tags, '../tags/tag1.xml'],
    ['rId10', T.hyperlink, 'https://example.com/keep', true],
    ['rId11', T.hyperlink, 'slide2.xml'],
    ['rId12', T.hyperlink, 'slide2.xml'],
    ['rId13', T.image, '../media/orphan.png'],
    ['rId14', T.chart, '../charts/chart-orphan.xml'],
    ['rId15', T.oleObject, '../embeddings/ole-orphan.bin'],
    ['rId16', T.tags, '../tags/tags-orphan.xml'],
    ['rId17', T.media, '../media/clip.mp4'],
    ['rId18', T.media, '../media/unused.mp4']
  ]);
  const slide2Rels = relationships([
    ['rId1', T.slideLayout, '../slideLayouts/slideLayout1.xml'],
    ['rId2', T.notesSlide, '../notesSlides/notesSlide2.xml'],
    ['rId3', T.image, '../media/beta.png'],
    ['rId4', T.hyperlink, 'slide1.xml']
  ]);
  const slide3Rels = relationships([
    ['rId1', T.slideLayout, '../slideLayouts/slideLayout1.xml'],
    ['rId2', T.notesSlide, '../notesSlides/notesSlide3.xml'],
    ['rId3', T.image, '../media/shared.png'],
    ['rId4', T.chart, '../charts/chart3.xml'],
    ['rId5', T.hyperlink, 'slide1.xml'],
    ['rId6', T.hyperlink, 'https://example.com/gamma', true]
  ]);
  const layoutRels = relationships([['rId1', T.slideMaster, '../slideMasters/slideMaster1.xml']]);
  const masterRels = relationships([
    ['rId1', T.slideLayout, '../slideLayouts/slideLayout1.xml'],
    ['rId2', T.theme, '../theme/theme1.xml']
  ]);
  const chart1Rels = relationships([['rId1', T.package, '../embeddings/book1.xlsx']]);

  return {
    '_rels/.rels': rootRels,
    '[Content_Types].xml': contentTypes(),
    'ppt/presentation.xml': presentationXml(),
    'ppt/_rels/presentation.xml.rels': presentationRels,
    'ppt/slides/slide1.xml': slideXml('Alpha', shapeTree(
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Alpha text"/><p:cNvSpPr/><p:nvPr>' +
      '<p:custDataLst><p:tags r:id="rId9"/></p:custDataLst><p:videoFile r:link="rId17"/></p:nvPr></p:nvSpPr><p:spPr/>' +
      '<p:txBody><a:bodyPr/><a:lstStyle/>' +
      '<a:p><a:r><a:rPr lang="en-US"><a:hlinkClick r:id="rId10"/><a:hlinkMouseOver r:id="rId10"/></a:rPr><a:t>Alpha keep</a:t></a:r></a:p>' +
      '<a:p><a:r><a:rPr lang="en-US"><a:hlinkClick r:id="rId11"/><a:hlinkMouseOver r:id="rId12"/></a:rPr><a:t>Beta jump</a:t></a:r></a:p>' +
      '</p:txBody></p:sp>' +
      picture(10, 'Shared picture', 'rId3') +
      picture(11, 'Unique picture', 'rId4') +
      picture(12, 'Vector picture', 'rId5', SVG_BLIP) +
      chartShape(20, 'Chart 1', 'rId7') +
      '<p:oleObj r:id="rId8" imgW="100" imgH="100" progId="Package"><p:embed/></p:oleObj>'
    )),
    'ppt/slides/_rels/slide1.xml.rels': slide1Rels,
    'ppt/slides/slide2.xml': slideXml('Beta', shapeTree(
      textShape(2, 'Beta text', 'Beta body', '<a:hlinkClick r:id="rId4"/>') +
      picture(10, 'Beta picture', 'rId3')
    )),
    'ppt/slides/_rels/slide2.xml.rels': slide2Rels,
    'ppt/slides/slide3.xml': slideXml('Gamma', shapeTree(
      textShape(2, 'Gamma text', 'Gamma back', '<a:hlinkClick r:id="rId5"/>') +
      textShape(3, 'Gamma external', 'Gamma web', '<a:hlinkClick r:id="rId6"/>') +
      picture(10, 'Shared picture', 'rId3') +
      chartShape(20, 'Chart 3', 'rId4')
    )),
    'ppt/slides/_rels/slide3.xml.rels': slide3Rels,
    'ppt/slideLayouts/slideLayout1.xml': slideLayoutXml(),
    'ppt/slideLayouts/_rels/slideLayout1.xml.rels': layoutRels,
    'ppt/slideMasters/slideMaster1.xml': slideMasterXml(),
    'ppt/slideMasters/_rels/slideMaster1.xml.rels': masterRels,
    'ppt/theme/theme1.xml': themeXml,
    'ppt/notesSlides/notesSlide1.xml': notesSlideXml(),
    'ppt/notesSlides/notesSlide2.xml': notesSlideXml(),
    'ppt/notesSlides/notesSlide3.xml': notesSlideXml(),
    'ppt/charts/chart1.xml': chartXml,
    'ppt/charts/_rels/chart1.xml.rels': chart1Rels,
    'ppt/charts/chart3.xml': plainChartXml,
    'ppt/charts/chart-orphan.xml': plainChartXml,
    'ppt/tags/tag1.xml': `<p:tagLst xmlns:p="${P}"><p:tag name="stage" val="final"/></p:tagLst>`,
    'ppt/tags/tags-orphan.xml': `<p:tagLst xmlns:p="${P}"><p:tag name="removed" val="yes"/></p:tagLst>`,
    'ppt/presProps.xml': `<p:presentationPr xmlns:p="${P}"/>`,
    'ppt/viewProps.xml': `<p:viewPr xmlns:p="${P}"/>`,
    'ppt/tableStyles.xml': `<a:tblStyleLst xmlns:a="${A}" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>`,
    'ppt/custom/thing.xml': '<custom xmlns="http://example.com/custom"><value>kept</value></custom>',
    'ppt/embeddings/ole1.bin': Buffer.from('ole object'),
    'ppt/embeddings/ole-orphan.bin': Buffer.from('orphan ole object'),
    'ppt/embeddings/book1.xlsx': Buffer.from('embedded workbook'),
    'ppt/media/shared.png': png,
    'ppt/media/one.png': png,
    'ppt/media/beta.png': png,
    'ppt/media/orphan.png': png,
    'ppt/media/svg-fallback.png': png,
    'ppt/media/vector.svg': '<svg xmlns="http://www.w3.org/2000/svg" width="1" height="1"/>',
    'ppt/media/clip.mp4': Buffer.from('video clip'),
    'ppt/media/unused.mp4': Buffer.from('unused video'),
    'docProps/core.xml': '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
      'xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Prune test</dc:title></cp:coreProperties>',
    'docProps/app.xml': appXml(),
    'docProps/thumbnail.jpeg': Buffer.from([0xff, 0xd8, 0xff, 0xd9])
  };
}

async function makeDeck() {
  const zip = new JSZip();
  for (const [name, data] of Object.entries(deckFiles())) zip.file(name, data);
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

function fileNames(zip) {
  return new Set(Object.keys(zip.files).filter(name => !zip.files[name].dir));
}
async function documentOf(zip, name) {
  const file = zip.file(name);
  assert.ok(file, 'missing part ' + name);
  return parseXml(await file.async('string'));
}

async function restrictSlides(zip, ids) {
  const doc = await documentOf(zip, 'ppt/presentation.xml');
  const list = descendants(doc.documentElement, 'sldIdLst')[0];
  const byId = new Map(children(list, 'sldId').map(node => [node.getAttributeNS(R, 'id'), node]));
  for (const node of children(list, 'sldId')) list.removeChild(node);
  for (const id of ids) {
    assert.ok(byId.has(id), 'missing selected slide relationship ' + id);
    list.appendChild(byId.get(id));
  }
  zip.file('ppt/presentation.xml', serializer.serializeToString(doc));
}

async function relationshipEntries(zip, name) {
  const doc = await documentOf(zip, name);
  return children(doc.documentElement, 'Relationship').map(node => ({
    id: attr(node, 'Id'), type: attr(node, 'Type'), target: attr(node, 'Target'),
    external: attr(node, 'TargetMode') === 'External'
  }));
}
function ownerForRels(name) {
  if (name === '_rels/.rels') return '';
  const marker = name.lastIndexOf('/_rels/');
  const directory = name.slice(0, marker);
  const base = name.slice(marker + 7).replace(/\.rels$/, '');
  return directory ? directory + '/' + base : base;
}
function resolvePart(owner, target) {
  const part = target.split('#')[0];
  const directory = path.posix.dirname(owner || '');
  return path.posix.normalize(path.posix.join(directory === '.' ? '' : directory, part.replace(/^\/+/, '')));
}

async function assertPackageIntegrity(zip) {
  const names = fileNames(zip);
  for (const name of names) {
    if (!name.endsWith('.rels')) continue;
    const entries = await relationshipEntries(zip, name);
    assert.equal(new Set(entries.map(entry => entry.id)).size, entries.length, 'duplicate relationship id in ' + name);
    const owner = ownerForRels(name);
    for (const entry of entries) {
      if (entry.external) continue;
      const target = resolvePart(owner, entry.target);
      assert.ok(names.has(target), `${name} points at missing part ${target}`);
    }
  }
  for (const name of names) {
    if (name === '[Content_Types].xml' || name.endsWith('.rels') || !/\.(xml|vml)$/i.test(name)) continue;
    const doc = await documentOf(zip, name);
    const ownerRels = names.has(name.replace(/([^/]+)$/, '_rels/$1.rels'))
      ? name.replace(/([^/]+)$/, '_rels/$1.rels') : null;
    const ids = new Set();
    if (ownerRels) {
      for (const entry of await relationshipEntries(zip, ownerRels)) ids.add(entry.id);
    }
    for (const node of Array.from(doc.getElementsByTagName('*'))) {
      const attributes = node.attributes;
      for (let index = 0; index < attributes.length; index++) {
        const attribute = attributes.item(index);
        if (attribute.namespaceURI !== R && attribute.localName !== 'relid' && !/^r:/.test(attribute.name || '')) continue;
        assert.ok(ids.has(attribute.value), `${name} references missing relationship ${attribute.value}`);
      }
    }
  }
  const types = await documentOf(zip, '[Content_Types].xml');
  for (const override of children(types.documentElement, 'Override')) {
    const part = attr(override, 'PartName').replace(/^\//, '');
    assert.ok(names.has(part), 'content type override for missing part ' + part);
  }
}

test('prunes excluded slides while keeping reachable rich content and metadata', async () => {
  const zip = await JSZip.loadAsync(await makeDeck());
  await restrictSlides(zip, ['rId3', 'rId1']);
  await prunePresentation(zip, { indices: [2, 0] });

  const names = fileNames(zip);
  assert.ok(names.has('ppt/slides/slide1.xml'));
  assert.ok(names.has('ppt/slides/slide3.xml'));
  assert.ok(!names.has('ppt/slides/slide2.xml'));
  assert.ok(!names.has('ppt/slides/_rels/slide2.xml.rels'));
  assert.ok(names.has('ppt/notesSlides/notesSlide1.xml'));
  assert.ok(names.has('ppt/notesSlides/notesSlide3.xml'));
  assert.ok(!names.has('ppt/notesSlides/notesSlide2.xml'));
  assert.ok(names.has('ppt/media/shared.png'));
  assert.ok(names.has('ppt/media/one.png'));
  assert.ok(names.has('ppt/media/svg-fallback.png'));
  assert.ok(names.has('ppt/media/vector.svg'));
  assert.ok(names.has('ppt/media/clip.mp4'));
  assert.ok(!names.has('ppt/media/unused.mp4'));
  assert.ok(!names.has('ppt/media/beta.png'));
  assert.ok(!names.has('ppt/media/orphan.png'));
  assert.ok(names.has('ppt/charts/chart1.xml'));
  assert.ok(names.has('ppt/charts/chart3.xml'));
  assert.ok(!names.has('ppt/charts/chart-orphan.xml'));
  assert.ok(names.has('ppt/embeddings/book1.xlsx'));
  assert.ok(names.has('ppt/embeddings/ole1.bin'));
  assert.ok(!names.has('ppt/embeddings/ole-orphan.bin'));
  assert.ok(names.has('ppt/tags/tag1.xml'));
  assert.ok(!names.has('ppt/tags/tags-orphan.xml'));
  assert.ok(names.has('ppt/slideLayouts/slideLayout1.xml'));
  assert.ok(names.has('ppt/slideMasters/slideMaster1.xml'));
  assert.ok(names.has('ppt/theme/theme1.xml'));
  assert.ok(names.has('ppt/presProps.xml'));
  assert.ok(names.has('ppt/viewProps.xml'));
  assert.ok(names.has('ppt/tableStyles.xml'));
  assert.ok(names.has('ppt/custom/thing.xml'));
  assert.ok(names.has('docProps/app.xml'));
  assert.ok(names.has('docProps/core.xml'));
  assert.ok(names.has('docProps/thumbnail.jpeg'));

  const presentation = await documentOf(zip, 'ppt/presentation.xml');
  const slideIds = children(descendants(presentation, 'sldIdLst')[0], 'sldId')
    .map(node => node.getAttributeNS(R, 'id'));
  assert.deepEqual(slideIds, ['rId3', 'rId1']);
  const customShows = descendants(presentation, 'custShow');
  assert.equal(customShows.length, 1);
  assert.equal(attr(customShows[0], 'name'), 'Mixed');
  const customSlides = children(child(customShows[0], 'sldLst'), 'sld')
    .map(node => node.getAttributeNS(R, 'id'));
  assert.deepEqual(customSlides, ['rId3']);
  assert.ok(!serializer.serializeToString(presentation).includes('rId2'));

  const presentationRels = await relationshipEntries(zip, 'ppt/_rels/presentation.xml.rels');
  assert.ok(!presentationRels.some(entry => entry.id === 'rId2'));
  assert.deepEqual(presentationRels
    .filter(entry => entry.type === T.slide)
    .map(entry => entry.target)
    .sort(), ['slides/slide1.xml', 'slides/slide3.xml']);
  for (const id of ['rId4', 'rId5', 'rId6', 'rId7', 'rId8']) {
    assert.ok(presentationRels.some(entry => entry.id === id), 'preserved presentation relationship ' + id);
  }

  const slide1 = await documentOf(zip, 'ppt/slides/slide1.xml');
  const slide1Text = serializer.serializeToString(slide1);
  assert.match(slide1Text, /rId10/);
  assert.doesNotMatch(slide1Text, /rId11|rId12/);
  const links = Array.from(slide1.getElementsByTagName('*'))
    .filter(node => node.localName === 'hlinkClick' || node.localName === 'hlinkMouseOver');
  assert.equal(links.length, 2);
  for (const link of links) assert.equal(link.getAttributeNS(R, 'id'), 'rId10');

  const slide1Rels = await relationshipEntries(zip, 'ppt/slides/_rels/slide1.xml.rels');
  const slide1RelIds = new Set(slide1Rels.map(entry => entry.id));
  for (const id of ['rId1', 'rId2', 'rId3', 'rId4', 'rId5', 'rId6', 'rId7', 'rId8', 'rId9', 'rId10', 'rId17']) {
    assert.ok(slide1RelIds.has(id), 'preserved slide relationship ' + id);
  }
  for (const id of ['rId11', 'rId12', 'rId13', 'rId14', 'rId15', 'rId16', 'rId18']) {
    assert.ok(!slide1RelIds.has(id), 'removed slide relationship ' + id);
  }
  assert.deepEqual(slide1Rels.find(entry => entry.id === 'rId10'), {
    id: 'rId10', type: T.hyperlink, target: 'https://example.com/keep', external: true
  });

  const slide3 = await documentOf(zip, 'ppt/slides/slide3.xml');
  assert.match(serializer.serializeToString(slide3), /rId5/);
  const slide3Rels = await relationshipEntries(zip, 'ppt/slides/_rels/slide3.xml.rels');
  assert.deepEqual(slide3Rels.find(entry => entry.id === 'rId5'),
    { id: 'rId5', type: T.hyperlink, target: 'slide1.xml', external: false });
  assert.deepEqual(slide3Rels.find(entry => entry.id === 'rId6'),
    { id: 'rId6', type: T.hyperlink, target: 'https://example.com/gamma', external: true });

  const app = await documentOf(zip, 'docProps/app.xml');
  assert.equal(descendants(app, 'Slides')[0].textContent.trim(), '2');
  assert.equal(descendants(app, 'Notes')[0].textContent.trim(), '3');
  const titleVector = children(descendants(app, 'TitlesOfParts')[0], 'vector')[0];
  assert.deepEqual(children(titleVector, 'lpstr').map(node => node.textContent),
    ['Office Theme', 'Gamma', 'Alpha']);
  assert.ok(!serializer.serializeToString(app).includes('Beta'));
  const variants = children(children(descendants(app, 'HeadingPairs')[0], 'vector')[0], 'variant');
  const headingIndex = variants.findIndex(node => children(node, 'lpstr')[0]?.textContent === 'Slide Titles');
  assert.ok(headingIndex >= 0);
  assert.equal(children(variants[headingIndex + 1], 'i4')[0].textContent, '2');

  const types = await documentOf(zip, '[Content_Types].xml');
  const overrides = new Set(children(types.documentElement, 'Override').map(node => attr(node, 'PartName')));
  assert.ok(overrides.has('/ppt/slides/slide1.xml'));
  assert.ok(overrides.has('/ppt/slides/slide3.xml'));
  assert.ok(!overrides.has('/ppt/slides/slide2.xml'));
  assert.ok(!overrides.has('/ppt/notesSlides/notesSlide2.xml'));
  assert.ok(!overrides.has('/ppt/charts/chart-orphan.xml'));
  assert.ok(!overrides.has('/ppt/tags/tags-orphan.xml'));
  assert.ok(overrides.has('/ppt/charts/chart1.xml'));
  const defaults = new Set(children(types.documentElement, 'Default').map(node => attr(node, 'Extension')));
  assert.ok(defaults.has('rels'));
  assert.ok(defaults.has('png'));
  assert.ok(defaults.has('svg'));

  await assertPackageIntegrity(zip);
});

test('rejects selections that do not match the presentation slide list', async () => {
  const zip = await JSZip.loadAsync(await makeDeck());
  await assert.rejects(prunePresentation(zip, { indices: [0] }), /导出页码与演示文稿幻灯片数量不一致/);
  await assert.rejects(prunePresentation(zip, { indices: [0, 0, 1] }), /导出页码不能重复/);
});
