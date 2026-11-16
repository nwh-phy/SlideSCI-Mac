import path from 'node:path';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';

const RELATIONSHIPS_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const serializer = new XMLSerializer();

// These relationships point to explicitly referenced resources. If the owning
// part no longer contains a relationship attribute for them they are residue
// from removed shapes and can be dropped. Implicit relationships (slideLayout,
// notesSlide, slideMaster, theme, document properties, ...) and unrecognized
// relationship types are intentionally kept.
const EXPLICIT_RELATIONSHIP_TYPES = new Set([
  'image', 'chart', 'chartUserShapes', 'media', 'video', 'audio', 'oleObject',
  'package', 'tags', 'hyperlink', 'diagramData', 'diagramLayout',
  'diagramQuickStyle', 'diagramColors'
]);

function parseXml(text) {
  if (/<!DOCTYPE|<!ENTITY/i.test(text)) throw new Error('演示文稿 XML 包含不允许的外部实体。');
  const doc = new DOMParser({ onError(level) {
    if (level !== 'warning') throw new Error('演示文稿 XML 格式无效。');
  } }).parseFromString(text, 'application/xml');
  if (!doc.documentElement || doc.getElementsByTagName('parsererror').length) throw new Error('演示文稿 XML 格式无效。');
  return doc;
}

function children(node, name) {
  const result = [];
  for (let item = node?.firstChild; item; item = item.nextSibling) {
    if (item.nodeType === 1 && (!name || item.localName === name)) result.push(item);
  }
  return result;
}
function descendants(node, name) {
  const result = [];
  for (const item of children(node)) {
    if (item.localName === name) result.push(item);
    result.push(...descendants(item, name));
  }
  return result;
}
function allElements(doc) { return Array.from(doc.getElementsByTagName('*')); }
function attr(node, key, fallback = '') { return node?.hasAttribute(key) ? node.getAttribute(key) : fallback; }
function textOf(node) { return node?.textContent || ''; }
function setText(node, value) {
  while (node.firstChild) node.removeChild(node.firstChild);
  node.appendChild(node.ownerDocument.createTextNode(String(value)));
}
function relationshipId(node) {
  return node?.getAttributeNS(RELATIONSHIPS_NS, 'id') || attr(node, 'r:id');
}
function relationshipTypeName(type) {
  return String(type || '').split('/').pop().replace(/[?#].*$/, '');
}

function relsPath(owner) {
  if (!owner) return '_rels/.rels';
  const directory = path.posix.dirname(owner);
  return path.posix.join(directory === '.' ? '' : directory, '_rels', path.posix.basename(owner) + '.rels');
}

function joinTarget(owner, target) {
  if (target.startsWith('/')) return path.posix.normalize(target.slice(1));
  const directory = path.posix.dirname(owner || '');
  return path.posix.normalize(path.posix.join(directory === '.' ? '' : directory, target));
}

function resolveTarget(owner, target, fileNames) {
  if (typeof target !== 'string') return null;
  const part = target.split('#')[0];
  if (!part) return null;
  const candidates = [];
  try { candidates.push(joinTarget(owner, part)); } catch { return null; }
  if (/%[0-9a-f]{2}/i.test(part)) {
    try {
      const decoded = decodeURIComponent(part);
      if (decoded !== part) candidates.push(joinTarget(owner, decoded));
    } catch { /* keep the literal target */ }
  }
  for (const candidate of candidates) if (fileNames.has(candidate)) return candidate;
  return candidates[0];
}

function isRelationshipAttribute(attribute) {
  if (!attribute) return false;
  return attribute.namespaceURI === RELATIONSHIPS_NS
    || attribute.localName === 'relid'
    || /^r:/.test(attribute.name || '');
}

function relationshipReferences(doc) {
  const ids = new Set();
  for (const node of allElements(doc)) {
    const attributes = node.attributes;
    for (let index = 0; index < attributes.length; index++) {
      const attribute = attributes.item(index);
      if (isRelationshipAttribute(attribute)) ids.add(attribute.value);
    }
  }
  return ids;
}

// Removes links (and any other direct relationship attributes) that point at
// relationships which are no longer part of the package. hlinkClick and
// hlinkMouseOver are removed as a whole so a dropped jump target cannot leave
// an empty hyperlink element behind; slide references inside custom shows are
// removed in the same pass.
function cleanRelationshipReferences(doc, removedIds) {
  if (!doc || !removedIds?.size) return false;
  let changed = false;
  for (const node of allElements(doc)) {
    if (!node.parentNode) continue;
    const matches = [];
    const attributes = node.attributes;
    for (let index = 0; index < attributes.length; index++) {
      const attribute = attributes.item(index);
      if (isRelationshipAttribute(attribute) && removedIds.has(attribute.value)) matches.push(attribute);
    }
    if (!matches.length) continue;
    if (node.localName === 'hlinkClick' || node.localName === 'hlinkMouseOver' || node.localName === 'sld') {
      node.parentNode.removeChild(node);
      changed = true;
      continue;
    }
    for (const attribute of matches) {
      try { node.removeAttributeNode(attribute); } catch { node.removeAttribute(attribute.name); }
      changed = true;
    }
  }
  return changed;
}

function pruneEmptyCustomShows(doc) {
  let changed = false;
  for (const list of allElements(doc).filter(node => node.localName === 'custShowLst')) {
    for (const show of children(list, 'custShow')) {
      if (children(children(show, 'sldLst')[0], 'sld').length) continue;
      list.removeChild(show);
      changed = true;
    }
    if (!children(list, 'custShow').length && list.parentNode) {
      list.parentNode.removeChild(list);
      changed = true;
    }
  }
  return changed;
}

async function readRelationships(zip, owner) {
  const name = relsPath(owner);
  const file = zip.file(name);
  if (!file) return { path: name, exists: false, doc: null, rels: [], dirty: false };
  const doc = parseXml(await file.async('string'));
  const rels = children(doc.documentElement, 'Relationship').map(node => ({
    id: attr(node, 'Id'),
    type: attr(node, 'Type'),
    target: attr(node, 'Target'),
    external: attr(node, 'TargetMode').toLowerCase() === 'external',
    node
  }));
  return { path: name, exists: true, doc, rels, dirty: false };
}

async function loadOwnerXml(zip, owner, fileNames) {
  if (!owner || !fileNames.has(owner) || !/\.(xml|vml)$/i.test(owner)) return null;
  const file = zip.file(owner);
  if (!file) return null;
  try {
    return { doc: parseXml(await file.async('string')), dirty: false };
  } catch {
    // An owner we cannot inspect is treated conservatively: its relationships
    // remain intact and are still traversed.
    return null;
  }
}

function normalizePartName(value) {
  let part = String(value || '');
  if (part.startsWith('/')) part = part.slice(1);
  try { part = decodeURIComponent(part); } catch { /* keep the literal name */ }
  return path.posix.normalize(part);
}

async function pruneContentTypes(zip, reachable) {
  const file = zip.file('[Content_Types].xml');
  if (!file) return;
  const doc = parseXml(await file.async('string'));
  const reachableLower = new Set([...reachable].map(name => name.toLowerCase()));
  let changed = false;
  for (const override of children(doc.documentElement, 'Override')) {
    const part = normalizePartName(attr(override, 'PartName'));
    if (part && (reachable.has(part) || reachableLower.has(part.toLowerCase()))) continue;
    override.parentNode.removeChild(override);
    changed = true;
  }
  if (changed) zip.file('[Content_Types].xml', serializer.serializeToString(doc));
}

async function updateAppProperties(zip, indices, reachable) {
  if (!reachable.has('docProps/app.xml')) return;
  const file = zip.file('docProps/app.xml');
  if (!file) return;
  const doc = parseXml(await file.async('string'));
  const root = doc.documentElement;
  const slidesNode = descendants(root, 'Slides')[0] || null;
  const originalSlides = Number(textOf(slidesNode).trim()) || 0;
  if (slidesNode) setText(slidesNode, indices.length);

  const headingVector = descendants(descendants(root, 'HeadingPairs')[0] || null, 'vector')[0] || null;
  const titleVector = descendants(descendants(root, 'TitlesOfParts')[0] || null, 'vector')[0] || null;
  const titleEntries = titleVector ? children(titleVector, 'lpstr') : [];
  let offset = 0;
  let count = 0;
  let found = false;
  let headingCountNode = null;
  if (headingVector) {
    const variants = children(headingVector, 'variant');
    let running = 0;
    for (let index = 0; index + 1 < variants.length; index += 2) {
      const label = textOf(descendants(variants[index], 'lpstr')[0]).trim();
      const countNode = descendants(variants[index + 1], 'i4')[0] || descendants(variants[index + 1], 'int')[0];
      const amount = Math.max(0, Number(textOf(countNode).trim()) || 0);
      if (label === 'Slide Titles') {
        offset = running;
        count = amount;
        found = true;
        headingCountNode = countNode;
        break;
      }
      running += amount;
    }
  }

  // Some producers omit HeadingPairs. Only assume the vector starts with
  // slide titles when its size makes that unambiguous.
  let canMap = found;
  if (!found && titleEntries.length > 0 && titleEntries.length === originalSlides) {
    offset = 0;
    count = originalSlides;
    canMap = true;
  }
  if (headingCountNode) setText(headingCountNode, indices.length);
  const write = () => zip.file('docProps/app.xml', serializer.serializeToString(doc));
  if (!titleVector || !canMap || !titleEntries.length) { write(); return; }

  const available = Math.max(0, titleEntries.length - offset);
  if (count > available) count = available;
  if (offset > titleEntries.length) { write(); return; }
  const template = titleEntries[0];
  const titles = indices.map(index => {
    const source = index >= 0 && index < count ? titleEntries[offset + index] : null;
    return source ? textOf(source) : 'Slide ' + (index + 1);
  });
  const replacement = titles.map(title => {
    const node = template.cloneNode(false);
    while (node.firstChild) node.removeChild(node.firstChild);
    node.appendChild(doc.createTextNode(title));
    return node;
  });
  const keptBefore = titleEntries.slice(0, offset);
  const keptAfter = titleEntries.slice(offset + count);
  for (const node of titleEntries) if (node.parentNode === titleVector) titleVector.removeChild(node);
  for (const node of [...keptBefore, ...replacement, ...keptAfter]) titleVector.appendChild(node);
  titleVector.setAttribute('size', String(keptBefore.length + replacement.length + keptAfter.length));
  write();
}

export async function prunePresentation(zip, { indices } = {}) {
  if (!zip || typeof zip.file !== 'function' || typeof zip.remove !== 'function') {
    throw new TypeError('缺少可裁剪的演示文稿 ZIP。');
  }
  if (!Array.isArray(indices) || !indices.length) throw new Error('导出页码不能为空。');
  if (indices.some(index => !Number.isInteger(index) || index < 0)) throw new Error('导出页码超出演示文稿范围。');
  if (new Set(indices).size !== indices.length) throw new Error('导出页码不能重复。');

  const fileNames = new Set(Object.keys(zip.files).filter(name => !zip.files[name].dir));
  if (!fileNames.has('ppt/presentation.xml')) throw new Error('文件中没有 PowerPoint 演示文稿。');
  if (!fileNames.has('_rels/.rels')) throw new Error('演示文稿关系数据不完整。');

  const presentationFile = zip.file('ppt/presentation.xml');
  const presentationDoc = parseXml(await presentationFile.async('string'));
  const sldIdLst = descendants(presentationDoc.documentElement, 'sldIdLst')[0] || null;
  const retainedRelIds = new Set(children(sldIdLst, 'sldId').map(relationshipId).filter(Boolean));
  if (retainedRelIds.size !== indices.length) throw new Error('导出页码与演示文稿幻灯片数量不一致。');

  const presentationRelationships = await readRelationships(zip, 'ppt/presentation.xml');
  const presentationRelById = new Map(presentationRelationships.rels.map(rel => [rel.id, rel]));
  const retainedSlidePaths = new Set();
  for (const id of retainedRelIds) {
    const relationship = presentationRelById.get(id);
    if (!relationship || relationship.external) throw new Error('演示文稿幻灯片关系不完整。');
    const target = resolveTarget('ppt/presentation.xml', relationship.target, fileNames);
    if (!target || !fileNames.has(target)) throw new Error('演示文稿幻灯片关系不完整。');
    retainedSlidePaths.add(target);
  }

  const slidePaths = new Set([...fileNames].filter(name => /^ppt\/slides\/[^/]+\.xml$/i.test(name)));
  for (const relationship of presentationRelationships.rels) {
    if (relationshipTypeName(relationship.type) !== 'slide') continue;
    const target = resolveTarget('ppt/presentation.xml', relationship.target, fileNames);
    if (target) slidePaths.add(target);
  }
  const excludedSlidePaths = new Set([...slidePaths].filter(name => !retainedSlidePaths.has(name)));

  const removedByOwner = new Map();
  const removedPresentationIds = new Set();
  for (const relationship of presentationRelationships.rels) {
    if (relationshipTypeName(relationship.type) !== 'slide' || retainedRelIds.has(relationship.id)) continue;
    relationship.node.parentNode.removeChild(relationship.node);
    removedPresentationIds.add(relationship.id);
    presentationRelationships.dirty = true;
  }
  presentationRelationships.rels = presentationRelationships.rels
    .filter(relationship => !removedPresentationIds.has(relationship.id));
  removedByOwner.set('ppt/presentation.xml', removedPresentationIds);

  const presentationOwner = { doc: presentationDoc, dirty: false };
  if (cleanRelationshipReferences(presentationDoc, removedPresentationIds)) presentationOwner.dirty = true;
  if (pruneEmptyCustomShows(presentationDoc)) presentationOwner.dirty = true;

  const relsCache = new Map([['ppt/presentation.xml', presentationRelationships]]);
  const ownerCache = new Map([['ppt/presentation.xml', presentationOwner]]);

  const reachable = new Set(['[Content_Types].xml', '_rels/.rels']);
  const visitedOwners = new Set();
  const queue = [''];
  while (queue.length) {
    const owner = queue.shift();
    if (visitedOwners.has(owner)) continue;
    visitedOwners.add(owner);
    if (owner) reachable.add(owner);

    const relsData = relsCache.get(owner) || await readRelationships(zip, owner);
    let ownerData = ownerCache.get(owner);
    if (!ownerData) {
      ownerData = await loadOwnerXml(zip, owner, fileNames);
      if (ownerData) ownerCache.set(owner, ownerData);
    }
    let removed = removedByOwner.get(owner);
    if (!removed) { removed = new Set(); removedByOwner.set(owner, removed); }

    if (ownerData && cleanRelationshipReferences(ownerData.doc, removed)) ownerData.dirty = true;
    const references = ownerData ? relationshipReferences(ownerData.doc) : null;

    for (const relationship of relsData.rels) {
      if (removed.has(relationship.id)) continue;
      const type = relationshipTypeName(relationship.type);
      if (EXPLICIT_RELATIONSHIP_TYPES.has(type) && references && !references.has(relationship.id)) {
        relationship.node.parentNode.removeChild(relationship.node);
        removed.add(relationship.id);
        relsData.dirty = true;
        continue;
      }
      if (relationship.external) continue;
      const target = resolveTarget(owner, relationship.target, fileNames);
      if (!target || !fileNames.has(target) || excludedSlidePaths.has(target)) {
        relationship.node.parentNode.removeChild(relationship.node);
        removed.add(relationship.id);
        relsData.dirty = true;
        continue;
      }
      if (!reachable.has(target)) {
        reachable.add(target);
        queue.push(target);
      }
    }

    if (ownerData) {
      if (cleanRelationshipReferences(ownerData.doc, removed)) ownerData.dirty = true;
      if (ownerData.dirty) zip.file(owner, serializer.serializeToString(ownerData.doc));
    }
    if (relsData.dirty) zip.file(relsData.path, serializer.serializeToString(relsData.doc));
  }

  const reachableRelationships = new Set();
  for (const owner of visitedOwners) {
    const name = relsPath(owner);
    if (fileNames.has(name)) reachableRelationships.add(name);
  }
  for (const name of fileNames) {
    if (!reachable.has(name) && !reachableRelationships.has(name)) zip.remove(name);
  }

  await pruneContentTypes(zip, reachable);
  await updateAppProperties(zip, indices, reachable);
  return zip;
}
