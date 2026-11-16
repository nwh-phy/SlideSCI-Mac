import { FOV_PREFIX, SCALE_PREFIX, UUID_TAG, RECORDS_TAG, ZOOM_TAG, SCALE_OWNER, uuid, units, unitOptions, positive, tag, parse, content, flatten, selected, pictures, isImage, fovOf, visibleFov, validFov, defaultScale, scaleOf, scaleLayout, writeRecord, localBounds, markerBounds, relativeRegion, geometry, near, guideSegments, lineBounds, element, control, imageFor, ensureUUID, findUUID } from './images-science-core.js';
import { imageIdentity, serializeTag } from './images-science-core.js';

const roles = (context, key, role) => context.shapes.filter(s => { const r = parse(tag(s, ZOOM_TAG)); return r?.key === key && r.role === role; });
const recordsOf = source => parse(tag(source, RECORDS_TAG), []);
const sourceSignature = source => JSON.stringify({ width: source.width, height: source.height, rotation: source.rotation || 0, crop: content(source).crop, fov: fovOf(source), color: tag(source, 'SLIDESCI_PSEUDO_COLOR'), image: imageIdentity(content(source).image) });
const cropSignature = (source, region) => JSON.stringify({ region, aspect: source.width / source.height, crop: content(source).crop, fov: fovOf(source), color: tag(source, 'SLIDESCI_PSEUDO_COLOR'), image: imageIdentity(content(source).image) });
const zoomDefaults = { name: '放大图', region: { x: .25, y: .25, width: .25, height: .25, rotation: 0 }, color: '#ff0000', lineWidth: 1.5, dash: 'solid', useColor: true, guides: true, extent: 'across', placement: 'right', scaleLength: null, scaleShowText: true };
const defaultLine = options => ({ color: options.color, width: +options.lineWidth, dashStyle: options.dash, visible: true });
const valuesOf = controls => Object.fromEntries(Object.entries(controls).map(([key, c]) => [key, c.type === 'checkbox' ? c.checked : c.type === 'number' ? +c.value : c.value]));
function setControls(controls, values) { for (const [key, c] of Object.entries(controls)) { const value = values[key]; if (c.type === 'checkbox') c.checked = !!value; else c.value = value ?? ''; } }
function errorStatus(status, error) { status.textContent = error.message; status.setAttribute('role', 'alert'); }
const mutations = new WeakMap();
async function scientificMutation(ctx, run) {
  mutations.set(ctx.host, (mutations.get(ctx.host) || 0) + 1);
  try { return await run(); } finally { mutations.set(ctx.host, mutations.get(ctx.host) - 1); }
}
export const applyScaleBar = (ctx, ...args) => scientificMutation(ctx, async () => (await applyScaleBarInternal(ctx, ...args)).result);
export const applyZoomEntries = (ctx, ...args) => scientificMutation(ctx, () => applyZoomEntriesInternal(ctx, ...args));

const snapshotKeys = ['left', 'top', 'width', 'height', 'rotation', 'name', 'alternativeText'];
function shapeSnapshot(shape) {
  const snapshot = { type: shape.type, tags: { ...(shape.tags || {}) } };
  for (const key of snapshotKeys) snapshot[key] = shape[key] ?? (key === 'rotation' ? 0 : '');
  for (const key of ['fill', 'line', 'crop']) if (shape[key] !== undefined) snapshot[key] = shape[key];
  return snapshot;
}
function snapshotSlide(context) {
  const shapes = new Map();
  const walk = (items, parent) => { for (const shape of items) { const snapshot = shapeSnapshot(shape); snapshot.parent = parent; shapes.set(shape.id, snapshot); walk(shape.children || [], shape.id); } };
  walk(context.shapes || [], null);
  return shapes;
}
const identical = (left, right) => JSON.stringify(left) === JSON.stringify(right);
function restoreProperties(snapshot, shape) {
  const properties = {};
  for (const key of snapshotKeys) if (!identical(shape[key] ?? (key === 'rotation' ? 0 : ''), snapshot[key])) properties[key] = snapshot[key];
  for (const key of ['fill', 'line', 'crop']) if (snapshot[key] !== undefined && !identical(shape[key], snapshot[key])) properties[key] = snapshot[key];
  const tags = { ...(shape.tags || {}) }, names = new Set(Object.keys(snapshot.tags).map(name => name.toUpperCase()));
  for (const key of Object.keys(tags)) if (!names.has(key.toUpperCase())) tags[key] = '';
  for (const [key, value] of Object.entries(snapshot.tags)) {
    for (const existing of Object.keys(tags)) if (existing !== key && existing.toUpperCase() === key.toUpperCase()) delete tags[existing];
    tags[key] = value;
  }
  if (!identical(tags, shape.tags || {})) properties.tags = tags;
  return Object.keys(properties).length ? properties : null;
}
// Roll back a failed entry batch: remove objects created by the attempt, restore surviving originals and report direct replacements.
async function rollbackSlide(ctx, before) {
  const context = await ctx.host.context({ scope: 'slide' }), current = new Map();
  const walk = (items, parent) => { for (const shape of items) { current.set(shape.id, { shape, parent }); walk(shape.children || [], shape.id); } };
  walk(context.shapes || []);
  const replacement = new Map();
  for (const [id, snapshot] of before) {
    if (current.has(id)) continue;
    const key = tag({ tags: snapshot.tags }, UUID_TAG), owner = tag({ tags: snapshot.tags }, SCALE_OWNER), part = tag({ tags: snapshot.tags }, 'SLIDESCI_SCALE_PART');
    let match = [...current.values()].find(({ shape }) => shape.id !== id && (key && tag(shape, UUID_TAG) === key && shape.type === snapshot.type
      || owner && part && tag(shape, SCALE_OWNER) === owner && tag(shape, 'SLIDESCI_SCALE_PART') === part));
    if (!match && snapshot.parent) {
      const parent = replacement.get(snapshot.parent) ?? snapshot.parent;
      match = [...current.values()].find(({ shape, parent: ownerId }) => shape.id !== id && ownerId === parent && shape.type === snapshot.type);
    }
    if (match) replacement.set(id, match.shape.id);
  }
  const missing = [...before.keys()].filter(id => !current.has(id) && !replacement.has(id)), commands = [];
  const retained = new Set(replacement.values()), insideRetained = id => { let parent = current.get(id)?.parent; while (parent) { if (retained.has(parent)) return true; parent = current.get(parent)?.parent; } return false; };
  // A host may rebuild a group as separate children; keep objects carrying its identity, remove unrelated leftovers.
  const missingKeys = new Set(missing.map(id => tag({ tags: before.get(id).tags }, UUID_TAG) || tag({ tags: before.get(id).tags }, SCALE_OWNER)).filter(Boolean));
  const created = [...current.values()].filter(({ shape }) => !before.has(shape.id) && !retained.has(shape.id) && !insideRetained(shape.id)
    && !(missing.length && (missingKeys.has(tag(shape, UUID_TAG)) || missingKeys.has(tag(shape, SCALE_OWNER))))).map(({ shape }) => shape.id);
  if (created.length) commands.push({ op: 'delete', ids: created });
  for (const [id, snapshot] of before) {
    const resolved = current.has(id) ? id : replacement.get(id), entry = resolved && current.get(resolved);
    if (!entry) continue;
    const properties = restoreProperties(snapshot, entry.shape);
    if (properties) commands.push({ op: 'update', id: resolved, properties });
  }
  if (commands.length) await ctx.host.apply(commands);
  return { missing };
}

async function applyScaleBarInternal(ctx, original, fov, settings) {
  validFov(fov);
  const s = { ...defaultScale, ...settings }, layout = scaleLayout(original, visibleFov(original, fov), s);
  let key = await ensureUUID(ctx, original), context = await ctx.host.context({ scope: 'slide' }), source = findUUID(context, key);
  if (!source) throw new Error('图片已不在当前幻灯片中。');
  const savedLinks = tag(source, RECORDS_TAG), originalGeometry = geometry(source), preservedTags = { ...source.tags, SLIDESCI_SCALE_CALIBRATION: serializeTag(visibleFov(original, fov)) };
  if (tag(source, 'SLIDESCI_SCALE_GROUP') === '1') {
    const childId = content(source).id;
    context = await ctx.host.apply([{ op: 'ungroup', id: source.id }]);
    source = flatten(context.shapes).find(x => x.id === childId) || context.shapes.find(x => tag(x, UUID_TAG) === key && isImage(x));
    if (!source) throw new Error('取消比例尺组合后未找到原图。');
    context = await ctx.host.apply([{ op: 'update', id: source.id, properties: { ...originalGeometry, tags: { ...preservedTags, [UUID_TAG]: key, ...(savedLinks ? { [RECORDS_TAG]: savedLinks } : {}) } } }]);
    source = context.shapes.find(x => x.id === source.id);
  }
  const oldAnnotations = context.shapes.filter(x => tag(x, SCALE_OWNER) === key).map(x => x.id);
  const alternativeText = writeRecord(writeRecord(source.alternativeText, FOV_PREFIX, fov), SCALE_PREFIX, s);
  const commands = [...(oldAnnotations.length ? [{ op: 'delete', ids: oldAnnotations }] : []), { op: 'update', id: source.id, properties: { alternativeText, tags: { [UUID_TAG]: key, SLIDESCI_SCALE_CALIBRATION: preservedTags.SLIDESCI_SCALE_CALIBRATION, SLIDESCI_SCALE_CONTENT: s.groupWithImage && layout.bar ? '1' : '', SLIDESCI_SCALE_GROUP: '' }, ...(s.groupWithImage && layout.bar ? { rotation: 0 } : {}) } }];
  const annotationKey = uuid(), rotate = s.groupWithImage ? 0 : originalGeometry.rotation;
  if (layout.bar) commands.push({ op: 'insertShape', shapeType: 'rectangle', ...localBounds(source, layout.bar, rotate), name: '比例尺', fill: { type: 'solid', color: s.barColor }, line: { visible: false }, tags: { [SCALE_OWNER]: key, SLIDESCI_SCALE_PART: 'bar', SLIDESCI_CREATION: annotationKey } });
  if (layout.text) commands.push({ op: 'insertText', ...localBounds(source, layout.text, rotate), name: '比例尺文字', text: layout.label, font: { name: s.fontName, size: +s.fontSize, color: s.fontColor, bold: s.bold }, paragraphs: [{ runs: [{ text: layout.label, font: { name: s.fontName, size: +s.fontSize, color: s.fontColor, bold: s.bold } }], alignment: s.corner.endsWith('Right') ? 'right' : 'left' }], textFrame: { marginLeft: 0, marginRight: 0, marginTop: 0, marginBottom: 0, wordWrap: false, autoSize: 'none' }, fill: { type: 'none' }, line: { visible: false }, tags: { [SCALE_OWNER]: key, SLIDESCI_SCALE_PART: 'text', SLIDESCI_CREATION: annotationKey } });
  context = await ctx.host.apply(commands);
  if (s.groupWithImage && layout.bar) {
    const ids = [source.id, ...context.shapes.filter(x => tag(x, 'SLIDESCI_CREATION') === annotationKey).map(x => x.id)];
    context = await ctx.host.apply([{ op: 'group', ids, name: original.name, tags: { ...preservedTags, [UUID_TAG]: key, SLIDESCI_SCALE_GROUP: '1', ...(savedLinks ? { [RECORDS_TAG]: savedLinks } : {}) } }]);
    const group = findUUID(context, key);
    if (!group || group.type !== 'group') throw new Error('比例尺组合未返回有效的组对象。');
    context = await ctx.host.apply([{ op: 'update', id: group.id, properties: { rotation: originalGeometry.rotation, alternativeText } }]);
  }
  const result = findUUID(context, key);
  // Select-only applies may answer with a geometry-only context, so keep the rich one for the caller.
  if (result) await ctx.host.apply([{ op: 'select', ids: [result.id] }]);
  await ctx.services.request('images/settings', { values: { scale: s } });
  return { result, context };
}
async function synchronizeUngroupedScale(ctx, context) {
  const commands = [];
  for (const source of context.shapes.filter(isImage)) {
    const settings = scaleOf(source), key = tag(source, UUID_TAG);
    if (!settings || !key) continue;
    if (tag(source, 'SLIDESCI_SCALE_GROUP') === '1') {
      if (tag(source, 'SLIDESCI_SCALE_CALIBRATION') !== serializeTag(visibleFov(source))) {
        context = (await applyScaleBarInternal(ctx, source, fovOf(source), settings)).context;
      }
      continue;
    }
    if (settings.groupWithImage) continue;
    const annotations = context.shapes.filter(s => tag(s, SCALE_OWNER) === key);
    if (!annotations.length) continue;
    const layout = scaleLayout(source, visibleFov(source), settings);
    for (const annotation of annotations) {
      const part = tag(annotation, 'SLIDESCI_SCALE_PART'), bounds = layout[part];
      if (!bounds) { commands.push({ op: 'delete', ids: [annotation.id] }); continue; }
      const expected = localBounds(source, bounds);
      if (!near(expected, annotation)) commands.push({ op: 'update', id: annotation.id, properties: expected });
    }
  }
  const live = new Set(context.shapes.filter(isImage).map(s => tag(s, UUID_TAG)));
  const orphans = context.shapes.filter(s => tag(s, SCALE_OWNER) && !live.has(tag(s, SCALE_OWNER))).map(s => s.id);
  if (orphans.length) commands.push({ op: 'delete', ids: orphans });
  return commands.length ? ctx.host.apply(commands) : context;
}
function zoomPlacement(source, region, placement, occupied) {
  const gap = .5 * 72 / 2.54, rw = region.width * source.width, rh = region.height * source.height;
  const horizontal = placement === 'left' || placement === 'right', scale = horizontal ? source.height / rh : source.width / rw;
  const width = rw * scale, height = rh * scale;
  const bounds = { left: placement === 'left' ? source.left - gap - width : placement === 'right' ? source.left + source.width + gap : source.left, top: placement === 'top' ? source.top - gap - height : placement === 'bottom' ? source.top + source.height + gap : source.top, width, height, rotation: source.rotation || 0 };
  for (let i = 0; i <= occupied.length; i++) {
    const collision = occupied.find(b => bounds.left < b.left + b.width && bounds.left + bounds.width > b.left && bounds.top < b.top + b.height && bounds.top + bounds.height > b.top);
    if (!collision) break;
    if (placement === 'left') bounds.left = collision.left - gap - width;
    if (placement === 'right') bounds.left = collision.left + collision.width + gap;
    if (placement === 'top') bounds.top = collision.top - gap - height;
    if (placement === 'bottom') bounds.top = collision.top + collision.height + gap;
  }
  return bounds;
}
function cropFov(source, region) {
  const fov = visibleFov(source); if (!fov) return null;
  const angle = (region.rotation || 0) * Math.PI / 180, c = Math.cos(angle), s = Math.sin(angle);
  const calibrated = (x, y) => (Math.abs(x) < 1e-10 || fov.width > 0) && (Math.abs(y) < 1e-10 || fov.height > 0);
  const width = calibrated(c, s) ? region.width * source.width * Math.hypot(fov.width / source.width * c, fov.height / source.height * s) : 0;
  const height = calibrated(s, c) ? region.height * source.height * Math.hypot(fov.width / source.width * s, fov.height / source.height * c) : 0;
  return { width, height, unit: 'μm', source: '放大区域可见 FOV' };
}
function fitZoomScale(source, region, options, previous) {
  const scale = previous || scaleOf(source); if (!scale) return null;
  const fov = cropFov(source, region), field = scale.vertical ? fov?.height : fov?.width;
  const lengthKey = scale.vertical ? 'heightMicrometers' : 'widthMicrometers';
  const requested = options.scaleLength == null || options.scaleLength === '' ? +scale[lengthKey] : +options.scaleLength;
  if (requested < 0 || !Number.isFinite(requested)) throw new Error('放大图比例尺长度必须大于或等于 0；0 表示隐藏。');
  if (requested && !positive(field)) throw new Error('放大图缺少所需方向的 FOV 标定。');
  return { ...scale, [lengthKey]: requested ? Math.min(requested, field * .9) : 0, showText: options.scaleShowText ?? scale.showText };
}
function validateRegion(region) { if (!region || ![region.x, region.y, region.width, region.height, region.rotation || 0].every(Number.isFinite) || region.width <= 0 || region.height <= 0) throw new Error('请为每个放大图框选有效区域。'); }
function roleTags(key, role, sourceUUID) { return { [ZOOM_TAG]: serializeTag({ key, role, sourceUUID }) }; }
async function createOrUpdateZoom(ctx, context, source, entry, occupied = [], { force = false } = {}) {
  const sourceUUID = tag(source, UUID_TAG), key = entry.key || uuid(), options = { ...zoomDefaults, ...entry, key }, region = options.region;
  validateRegion(region);
  if (!/^#[a-f\d]{6}$/i.test(options.color) || !positive(options.lineWidth) || +options.lineWidth > 1000 || !['left', 'right', 'top', 'bottom'].includes(options.placement)) throw new Error('放大图样式参数无效。');
  let marker = roles(context, key, 'marker')[0], zoom = roles(context, key, 'zoom')[0], guides = roles(context, key, 'guide');
  const markerGeometry = markerBounds(source, region), commands = [];
  const markerProperties = { ...markerGeometry, fill: { type: 'none' }, line: defaultLine(options), name: `${options.name} 区域`, tags: roleTags(key, 'marker', sourceUUID) };
  if (marker) commands.push({ op: 'update', id: marker.id, properties: markerProperties });
  else if (options.manualId && context.shapes.some(s => s.id === options.manualId)) commands.push({ op: 'update', id: options.manualId, properties: markerProperties });
  else commands.push({ op: 'insertShape', shapeType: 'rectangle', ...markerProperties });
  const needsPixels = force || !zoom || entry.cropSignature !== cropSignature(source, region);
  let rendered;
  if (needsPixels) rendered = await ctx.services.request('images/crop', { ...await imageFor(ctx, source), region, displayAspect: source.width / source.height });
  const previousScale = zoom ? scaleOf(zoom) : null;
  let bounds = zoom ? geometry(zoom) : zoomPlacement(source, region, options.placement, occupied);
  if (zoom && entry.placement !== entry.initialPlacement && entry.initialPlacement) bounds = zoomPlacement(source, region, options.placement, occupied);
  else if (zoom && needsPixels) {
    if (options.placement === 'left' || options.placement === 'right') bounds.width = bounds.height * (rendered.width / rendered.height);
    else bounds.height = bounds.width * (rendered.height / rendered.width);
  }
  const fov = cropFov(source, region), baseAlt = zoom ? content(zoom).alternativeText : content(source).alternativeText;
  const scale = fitZoomScale(source, region, options, previousScale);
  if (scale && fov) scaleLayout(bounds, { width: fov.width, height: fov.height }, scale);
  let alternativeText = writeRecord(writeRecord(baseAlt, SCALE_PREFIX, null), FOV_PREFIX, fov);
  const zoomUUID = zoom ? tag(zoom, UUID_TAG) || uuid() : uuid(), line = options.useColor ? defaultLine(options) : zoom?.line || { visible: false };
  if (zoom && needsPixels) {
    commands.push({ op: 'replaceImage', id: content(zoom).id, data: rendered.data, mime: rendered.mime, preserveCrop: false });
    commands.push({ op: 'update', id: zoom.id, properties: { ...bounds, name: options.name, alternativeText, line, tags: { ...roleTags(key, 'zoom', sourceUUID), [UUID_TAG]: zoomUUID, SLIDESCI_FOV_VISIBLE: '1' } } });
    if (content(zoom).id !== zoom.id) commands.push({ op: 'update', id: content(zoom).id, properties: { alternativeText, tags: { SLIDESCI_FOV_VISIBLE: '1' } } });
  } else if (zoom) commands.push({ op: 'update', id: zoom.id, properties: { ...bounds, name: options.name, line, tags: roleTags(key, 'zoom', sourceUUID) } });
  else commands.push({ op: 'insertImage', ...rendered, ...bounds, name: options.name, alternativeText, line, tags: { ...roleTags(key, 'zoom', sourceUUID), [UUID_TAG]: zoomUUID, SLIDESCI_FOV_VISIBLE: '1' } });
  context = await ctx.host.apply(commands);
  marker = roles(context, key, 'marker')[0]; zoom = roles(context, key, 'zoom')[0];
  if (!marker || !zoom) throw new Error('未找到新建放大图或区域矩形。');
  if (scale && fov && (needsPixels || JSON.stringify(scale) !== JSON.stringify(previousScale))) {
    zoom = await applyScaleBar(ctx, zoom, fov, scale);
    // Scale grouping retains scientific link tags on the outer object.
    context = await ctx.host.apply([{ op: 'update', id: zoom.id, properties: { tags: roleTags(key, 'zoom', sourceUUID) } }]);
  }
  zoom = roles(context, key, 'zoom')[0] || findUUID(context, zoomUUID);
  const guideCommands = [], segments = options.guides ? guideSegments(source, marker, zoom, options.extent) : [];
  guides = roles(context, key, 'guide');
  for (let i = 0; i < Math.max(guides.length, segments.length); i++) {
    if (i >= segments.length) guideCommands.push({ op: 'delete', ids: [guides[i].id] });
    else {
      const properties = { ...lineBounds(segments[i]), line: { ...defaultLine(options), visible: !!segments[i] }, name: `${options.name} 辅助线 ${i + 1}`, tags: { ...roleTags(key, 'guide', sourceUUID), SLIDESCI_ZOOM_GUIDE_INDEX: String(i) } };
      guideCommands.push(guides[i] ? { op: 'update', id: guides[i].id, properties } : { op: 'insertShape', shapeType: 'line', ...properties });
    }
  }
  if (guideCommands.length) context = await ctx.host.apply(guideCommands);
  // Every branch above ended in host.apply, which already returns the updated slide context.
  source = findUUID(context, sourceUUID);
  marker = roles(context, key, 'marker')[0];
  zoom = roles(context, key, 'zoom')[0];
  if (!source || !marker || !zoom) throw new Error('未找到新建放大图或区域矩形。');
  return { context, record: { ...options, manualId: undefined, initialPlacement: options.placement, sourceState: geometry(source), markerState: geometry(marker), zoomState: geometry(zoom), sourceSignature: sourceSignature(source), cropSignature: cropSignature(source, region) } };
}
async function applyZoomEntriesInternal(ctx, sources, entries) {
  const added = entries.filter(e => !e.key), errors = [];
  for (let index = 0; index < sources.length; index++) {
    const key = await ensureUUID(ctx, sources[index]);
    let context = await ctx.host.context({ scope: 'slide' }), source = findUUID(context, key);
    if (!source) { errors.push(`${sources[index].name}：图片已不在当前幻灯片中。`); continue; }
    const before = snapshotSlide(context);
    const old = recordsOf(source), target = index === 0 ? entries : [...old, ...added.map(e => ({ ...e, key: null, manualId: null }))];
    const removed = old.filter(e => !target.some(t => t.key === e.key)), finished = [], occupied = target.filter(e => e.key).map(e => roles(context, e.key, 'zoom')[0]).filter(Boolean).map(geometry);
    try {
      for (const entry of target) {
        const result = await createOrUpdateZoom(ctx, context, source, entry, occupied);
        context = result.context; finished.push(result.record); occupied.push(result.record.zoomState);
        source = findUUID(context, key);
        if (!source) throw new Error('更新放大区域后未找到原图。');
      }
      const removedKeys = new Set(removed.map(e => e.key)), removedOwners = new Set();
      for (const entry of removed) { const zoom = roles(context, entry.key, 'zoom')[0]; if (zoom) removedOwners.add(tag(zoom, UUID_TAG)); }
      const obsolete = context.shapes.filter(shape => {
        const link = parse(tag(shape, ZOOM_TAG));
        return !!(link?.key && removedKeys.has(link.key) || tag(shape, SCALE_OWNER) && removedOwners.has(tag(shape, SCALE_OWNER)));
      });
      const commands = [{ op: 'update', id: source.id, properties: { tags: { [RECORDS_TAG]: serializeTag(finished) } } }];
      if (obsolete.length) commands.push({ op: 'delete', ids: obsolete.map(shape => shape.id) });
      await ctx.host.apply(commands);
    } catch (error) {
      try {
        const { missing } = await rollbackSlide(ctx, before);
        errors.push(`${source.name}：${error.message}${missing.length ? '；回滚时宿主已重建部分对象，请核对幻灯片。' : ''}`);
      } catch (rollbackError) { errors.push(`${source.name}：${error.message}；回滚失败：${rollbackError.message}`); }
    }
  }
  if (errors.length) throw new Error(errors.join('\n'));
  await ctx.services.request('images/settings', { values: { zoom: entries.at(-1) ? { ...entries.at(-1), key: undefined, sourceState: undefined, markerState: undefined, zoomState: undefined } : zoomDefaults } });
}

const reconcilers = new WeakMap();
export function startScientificReconciliation(ctx, onError = error => ctx.services.reportError?.(error)) {
  const existing = reconcilers.get(ctx.host); if (existing) { existing.users++; return () => existing.release(); }
  let busy = false, pending = false, disposed = false, timer;
  const state = { users: 1, release() { if (--state.users > 0) return; disposed = true; clearTimeout(timer); unsubscribe?.(); reconcilers.delete(ctx.host); } };
  const run = async () => {
    if (busy) { pending = true; return; }
    if (disposed) return;
    if (mutations.get(ctx.host)) { timer = setTimeout(run, 80); return; }
    busy = true;
    try {
      let context = await ctx.host.context({ scope: 'slide' });
      context = await synchronizeUngroupedScale(ctx, context);
      const sources = context.shapes.filter(s => isImage(s) && tag(s, RECORDS_TAG));
      for (const initial of sources) {
        const sourceUUID = tag(initial, UUID_TAG); let source = findUUID(context, sourceUUID);
        const records = recordsOf(source), next = [], commands = []; let changed = false;
        if (context.shapes.filter(s => tag(s, UUID_TAG) === sourceUUID).length > 1) continue;
        for (const record of records) {
          const markers = roles(context, record.key, 'marker'), zooms = roles(context, record.key, 'zoom');
          // Same-slide copied links are ambiguous: preserve them without connecting unrelated shapes.
          if (markers.length > 1 || zooms.length > 1) { next.push(record); continue; }
          const marker = markers[0], zoom = zooms[0], guides = roles(context, record.key, 'guide');
          if (!zoom) { const ids = [marker, ...guides].filter(Boolean).map(s => s.id); if (ids.length) commands.push({ op: 'delete', ids }); changed = true; continue; }
          if (!marker) { commands.push({ op: 'update', id: zoom.id, properties: { tags: { [ZOOM_TAG]: '' } } }); if (guides.length) commands.push({ op: 'delete', ids: guides.map(s => s.id) }); changed = true; continue; }
          const sourceMoved = !near(record.sourceState, geometry(source)), markerMoved = !near(record.markerState, geometry(marker)), zoomMoved = !near(record.zoomState, geometry(zoom));
          const signatureChanged = sourceSignature(source) !== record.sourceSignature;
          const expectedGuides = record.guides ? guideSegments(source, marker, zoom, record.extent) : [];
          const guidesChanged = guides.length !== expectedGuides.length || guides.some((guide, index) => !near(geometry(guide), { ...geometry(guide), ...lineBounds(expectedGuides[index]) }));
          if (!sourceMoved && !markerMoved && !zoomMoved && !signatureChanged && !guidesChanged) { next.push(record); continue; }
          const options = { ...record, region: markerMoved ? relativeRegion(source, marker) : record.region };
          if (marker.line) { options.color = marker.line.color || options.color; options.lineWidth = marker.line.width || options.lineWidth; options.dash = marker.line.dashStyle || options.dash; }
          const result = await createOrUpdateZoom(ctx, context, source, options);
          context = result.context; next.push(result.record); changed = true; source = findUUID(context, sourceUUID);
        }
        if (changed) {
          commands.push({ op: 'update', id: source.id, properties: { tags: { [RECORDS_TAG]: serializeTag(next) } } });
          context = await ctx.host.apply(commands);
        }
      }
      // Removing the source detaches retained zoom pictures and removes its region/guide annotations.
      const live = new Set(context.shapes.filter(isImage).map(s => tag(s, UUID_TAG))), orphanCommands = [];
      for (const shape of context.shapes) { const link = parse(tag(shape, ZOOM_TAG)); if (!link || live.has(link.sourceUUID)) continue; orphanCommands.push(link.role === 'zoom' ? { op: 'update', id: shape.id, properties: { tags: { [ZOOM_TAG]: '' } } } : { op: 'delete', ids: [shape.id] }); }
      if (orphanCommands.length) await ctx.host.apply(orphanCommands);
    } catch (error) { onError(error); }
    finally { busy = false; if (pending && !disposed) { pending = false; timer = setTimeout(run, 80); } }
  };
  const schedule = () => { if (busy) pending = true; else { clearTimeout(timer); timer = setTimeout(run, 80); } };
  const unsubscribe = ctx.host.onChanged(schedule);
  reconcilers.set(ctx.host, state); schedule();
  return () => state.release();
}

async function mountScale(container, ctx) {
  const toolbar = element('div'), load = element('button', '读取所选图片'), apply = element('button', '应用到所选图片'), status = element('p', '', { role: 'status' }), fields = element('fieldset'), preview = element('canvas', null, { width: 600, height: 340 });
  fields.append(element('legend', 'FOV 与可编辑比例尺')); toolbar.append(load, apply); container.append(toolbar, fields, preview, status);
  const fovControls = { width: control(fields, '原图 FOV 宽度（0 为未标定）', 'number', 0), height: control(fields, '原图 FOV 高度（0 为未标定）', 'number', 0), unit: control(fields, 'FOV 单位', 'select', 'μm', unitOptions) };
  const controls = {
    vertical: control(fields, '纵向比例尺', 'checkbox', false), widthMicrometers: control(fields, '横向长度（μm，0 隐藏）', 'number', 50), heightMicrometers: control(fields, '纵向长度（μm，0 隐藏）', 'number', 50), lengthUnit: control(fields, '显示单位', 'select', 'μm', unitOptions), thicknessPoints: control(fields, '厚度（pt）', 'number', 3), showText: control(fields, '显示文字', 'checkbox', true), barColor: control(fields, '比例尺颜色', 'color', '#ffffff'), fontColor: control(fields, '文字颜色', 'color', '#ffffff'), fontName: control(fields, '字体', 'text', 'Arial'), fontSize: control(fields, '字号（pt）', 'number', 14), bold: control(fields, '加粗', 'checkbox', true), corner: control(fields, '位置', 'select', 'BottomRight', [{ value: 'TopLeft', label: '左上' }, { value: 'BottomLeft', label: '左下' }, { value: 'TopRight', label: '右上' }, { value: 'BottomRight', label: '右下' }]), groupWithImage: control(fields, '与图片组合', 'checkbox', true)
  };
  let shapes = [], image = null, disposed = false;
  function draw() {
    if (!shapes[0]) return;
    const g = preview.getContext('2d'), shape = shapes[0], fov = { ...valuesOf(fovControls), source: '手动设置' }, s = valuesOf(controls);
    g.clearRect(0, 0, preview.width, preview.height);
    const scale = Math.min((preview.width - 32) / shape.width, (preview.height - 32) / shape.height), x = (preview.width - shape.width * scale) / 2, y = (preview.height - shape.height * scale) / 2;
    g.save(); g.translate(x, y); g.scale(scale, scale);
    if (image) g.drawImage(image, 0, 0, shape.width, shape.height); else { g.fillStyle = '#30343a'; g.fillRect(0, 0, shape.width, shape.height); }
    try {
      validFov(fov); const visible = visibleFov(shape, fov), layout = scaleLayout(shape, visible, s);
      if (layout.bar) { g.fillStyle = s.barColor; g.fillRect(layout.bar.left, layout.bar.top, layout.bar.width, layout.bar.height); }
      if (layout.text) { g.fillStyle = s.fontColor; g.font = `${s.bold ? 'bold ' : ''}${s.fontSize}px "${s.fontName}"`; g.textBaseline = 'top'; g.fillText(layout.label, layout.text.left, layout.text.top); }
      status.textContent = `可见 FOV：${visible.width.toFixed(2)} × ${visible.height.toFixed(2)} μm。${fovOf(shape)?.source || ''}`; apply.disabled = false;
    } catch (error) { errorStatus(status, error); apply.disabled = true; }
    g.restore();
  }
  fields.addEventListener('input', draw);
  load.onclick = async () => { try { shapes = pictures(await ctx.host.context()); const fov = fovOf(shapes[0]) || { width: 0, height: 0, unit: 'μm' }; const saved = await ctx.services.request('images/settings', {}); setControls(fovControls, fov); setControls(controls, scaleOf(shapes[0]) || saved.scale || defaultScale); const rendered = await ctx.services.request('images/crop', { ...await imageFor(ctx, shapes[0]), region: { x: 0, y: 0, width: 1, height: 1, rotation: 0 }, maximum: 800 }); image = new Image(); image.onload = () => { if (!disposed) draw(); }; image.src = `data:${rendered.mime};base64,${rendered.data}`; draw(); } catch (error) { errorStatus(status, error); } };
  apply.onclick = async () => { apply.disabled = true; try { if (!shapes.length) throw new Error('请先读取所选图片。'); const fov = validFov({ ...valuesOf(fovControls), source: '手动设置' }), s = valuesOf(controls); for (const shape of shapes) scaleLayout(shape, visibleFov(shape, fov), s); for (const shape of shapes) await applyScaleBar(ctx, shape, fov, s); status.textContent = `已更新 ${shapes.length} 张图片的 FOV 与比例尺。`; } catch (error) { errorStatus(status, error); } finally { apply.disabled = false; } };
  return () => { disposed = true; };
}

async function mountZoom(container, ctx) {
  const toolbar = element('div'), load = element('button', '读取所选图片/矩形'), add = element('button', '新增区域'), remove = element('button', '删除当前区域'), fit = element('button', '适合窗口'), apply = element('button', '应用全部区域'), status = element('p', '滚轮缩放；空格/中键拖动平移；拖动绘制区域、内部移动、角点缩放；Delete 删除。', { role: 'status' });
  toolbar.append(load, add, remove, fit, apply); const list = control(toolbar, '当前区域', 'select', '', []);
  const canvas = element('canvas', null, { width: 760, height: 460, tabIndex: 0 }); canvas.setAttribute('aria-label', '多个放大区域编辑画布'); canvas.setAttribute('role', 'application');
  const fields = element('fieldset'); fields.append(element('legend', '每个区域独立设置')); container.append(toolbar, canvas, fields, status);
  const layoutPreview = element('canvas', null, { width: 760, height: 360 });
  layoutPreview.setAttribute('aria-label', '全部放大图、方向、辅助线和比例尺布局预览');
  container.insertBefore(layoutPreview, fields);
  const controls = {
    name: control(fields, '名称', 'text', '放大图 1'), color: control(fields, '矩形与辅助线颜色', 'color', zoomDefaults.color), lineWidth: control(fields, '轮廓宽度（pt）', 'number', 1.5), dash: control(fields, '线型', 'select', 'solid', [{ value: 'solid', label: '实线' }, { value: 'dash', label: '虚线' }, { value: 'dot', label: '点线' }, { value: 'dashDot', label: '点划线' }, { value: 'dashDotDot', label: '双点划线' }]), useColor: control(fields, '放大图使用矩形颜色', 'checkbox', true), guides: control(fields, '添加辅助线', 'checkbox', true), extent: control(fields, '辅助线范围', 'select', 'across', [{ value: 'across', label: '连接原图与放大图' }, { value: 'inside', label: '仅原图内部' }]), placement: control(fields, '放大图方向', 'select', 'right', [{ value: 'right', label: '右' }, { value: 'left', label: '左' }, { value: 'top', label: '上' }, { value: 'bottom', label: '下' }]), scaleLength: control(fields, '比例尺长度（μm，0 隐藏，留空继承）', 'number', ''), scaleShowText: control(fields, '比例尺显示文字', 'checkbox', true), rotation: control(fields, '区域旋转（度）', 'number', 0), x: control(fields, '区域 X（原图比例）', 'number', .25), y: control(fields, '区域 Y（原图比例）', 'number', .25), width: control(fields, '区域宽度（原图比例）', 'number', .25), height: control(fields, '区域高度（原图比例）', 'number', .25)
  };
  for (const key of ['x', 'y', 'width', 'height']) controls[key].step = '.001'; controls.rotation.step = '1'; controls.lineWidth.min = '.01'; controls.lineWidth.step = '.1'; controls.scaleLength.min = '0';
  let sources = [], entries = [], active = -1, image = null, previewKey = '', zoom = 1, pan = { x: 0, y: 0 }, drag = null, space = false, disposed = false, imageGeneration = 0;
  const g = canvas.getContext('2d');
  function view() { if (!sources[0]) return { scale: 1, x: 0, y: 0 }; const scale = Math.min((canvas.width - 50) / sources[0].width, (canvas.height - 50) / sources[0].height) * zoom; return { scale, x: (canvas.width - sources[0].width * scale) / 2 + pan.x, y: (canvas.height - sources[0].height * scale) / 2 + pan.y }; }
  function point(event) { const b = canvas.getBoundingClientRect(), v = view(); return { x: ((event.clientX - b.left) * canvas.width / b.width - v.x) / v.scale / (sources[0]?.width || 1), y: ((event.clientY - b.top) * canvas.height / b.height - v.y) / v.scale / (sources[0]?.height || 1) }; }
  function updateList() { list.replaceChildren(...entries.map((e, i) => element('option', e.name, { value: String(i) }))); list.value = String(active); }
  function selectEntry(index) { active = index; const entry = entries[active]; if (entry) setControls(controls, { ...entry, ...entry.region, scaleLength: entry.scaleLength ?? '' }); updateList(); draw(); }
  function saveFields() { const entry = entries[active]; if (!entry) return; const values = valuesOf(controls); entry.region = { x: values.x, y: values.y, width: values.width, height: values.height, rotation: values.rotation }; for (const key of Object.keys(zoomDefaults).filter(k => k !== 'region')) if (key in values) entry[key] = values[key]; entry.scaleLength = controls.scaleLength.value === '' ? null : +controls.scaleLength.value; updateList(); draw(); }
  function drawView() {
    g.clearRect(0, 0, canvas.width, canvas.height); g.fillStyle = '#1e2228'; g.fillRect(0, 0, canvas.width, canvas.height);
    if (!sources[0]) return;
    const v = view(), source = sources[0]; g.save(); g.translate(v.x, v.y); g.scale(v.scale, v.scale);
    if (image) g.drawImage(image, 0, 0, source.width, source.height);
    entries.forEach((entry, index) => {
      const r = entry.region, w = r.width * source.width, h = r.height * source.height, cx = (r.x + r.width / 2) * source.width, cy = (r.y + r.height / 2) * source.height;
      g.save(); g.translate(cx, cy); g.rotate((r.rotation || 0) * Math.PI / 180); g.strokeStyle = entry.color; g.lineWidth = (index === active ? 2.5 : 1.5) / v.scale; g.setLineDash(entry.dash === 'solid' ? [] : entry.dash === 'dot' ? [2 / v.scale, 3 / v.scale] : [7 / v.scale, 4 / v.scale]); g.strokeRect(-w / 2, -h / 2, w, h); g.setLineDash([]);
      g.fillStyle = entry.color; g.font = `${12 / v.scale}px Arial`; g.fillText(entry.name, -w / 2, -h / 2 - 5 / v.scale);
      if (index === active) for (const [x, y] of [[-w / 2, -h / 2], [w / 2, -h / 2], [w / 2, h / 2], [-w / 2, h / 2]]) g.fillRect(x - 4 / v.scale, y - 4 / v.scale, 8 / v.scale, 8 / v.scale);
      g.restore();
    }); g.restore();
  }
  // Pan and zoom never change the region layout, so they repaint only the editing canvas.
  function draw() { drawView(); drawLayout(); }
  function drawLayout() {
    const paint = layoutPreview.getContext('2d'); paint.clearRect(0, 0, layoutPreview.width, layoutPreview.height);
    if (!sources[0]) return;
    const source = { ...sources[0], left: 0, top: 0, rotation: 0 }, occupied = [], layouts = [];
    for (const entry of entries) {
      if (!(entry.region.width > 0 && entry.region.height > 0)) continue;
      const old = entry.zoomState;
      const bounds = old && entry.placement === entry.initialPlacement ? { ...old, left: old.left - sources[0].left, top: old.top - sources[0].top, rotation: (old.rotation || 0) - (sources[0].rotation || 0) } : zoomPlacement(source, entry.region, entry.placement, occupied);
      occupied.push(bounds); layouts.push({ entry, bounds });
    }
    const minX = Math.min(0, ...occupied.map(b => b.left)), minY = Math.min(0, ...occupied.map(b => b.top)), maxX = Math.max(source.width, ...occupied.map(b => b.left + b.width)), maxY = Math.max(source.height, ...occupied.map(b => b.top + b.height));
    const scale = Math.min((layoutPreview.width - 32) / (maxX - minX), (layoutPreview.height - 40) / (maxY - minY));
    paint.save(); paint.translate(16 - minX * scale, 24 - minY * scale); paint.scale(scale, scale);
    if (image) paint.drawImage(image, 0, 0, source.width, source.height);
    const paintBar = (bounds, fov, settings) => {
      if (!settings || !fov) return;
      const layout = scaleLayout(bounds, fov, settings);
      if (layout.bar) { paint.fillStyle = settings.barColor; paint.fillRect(bounds.left + layout.bar.left, bounds.top + layout.bar.top, layout.bar.width, layout.bar.height); }
      if (layout.text) { paint.fillStyle = settings.fontColor; paint.font = `${settings.bold ? 'bold ' : ''}${settings.fontSize}px "${settings.fontName}"`; paint.textBaseline = 'top'; paint.fillText(layout.label, bounds.left + layout.text.left, bounds.top + layout.text.top); }
    };
    let invalid = null;
    try { paintBar(source, visibleFov(source), scaleOf(source)); } catch (error) { invalid = error; }
    for (const { entry, bounds } of layouts) {
      const r = entry.region, centerX = (r.x + r.width / 2) * source.width, centerY = (r.y + r.height / 2) * source.height;
      paint.save(); paint.beginPath(); paint.rect(bounds.left, bounds.top, bounds.width, bounds.height); paint.clip();
      paint.translate(bounds.left + bounds.width / 2, bounds.top + bounds.height / 2); paint.scale(bounds.width / (r.width * source.width), bounds.height / (r.height * source.height)); paint.rotate(-(r.rotation || 0) * Math.PI / 180); paint.translate(-centerX, -centerY);
      if (image) paint.drawImage(image, 0, 0, source.width, source.height); paint.restore();
      paint.strokeStyle = entry.color; paint.lineWidth = +entry.lineWidth; paint.setLineDash(entry.dash === 'solid' ? [] : entry.dash === 'dot' ? [1, 2] : [5, 3]);
      const marker = markerBounds(source, r); paint.save(); paint.translate(marker.left + marker.width / 2, marker.top + marker.height / 2); paint.rotate((marker.rotation || 0) * Math.PI / 180); paint.strokeRect(-marker.width / 2, -marker.height / 2, marker.width, marker.height); paint.restore();
      if (entry.useColor) paint.strokeRect(bounds.left, bounds.top, bounds.width, bounds.height);
      if (entry.guides) for (const segment of guideSegments(source, marker, bounds, entry.extent)) { if (!segment) continue; paint.beginPath(); paint.moveTo(segment[0].x, segment[0].y); paint.lineTo(segment[1].x, segment[1].y); paint.stroke(); }
      paint.setLineDash([]); paint.fillStyle = entry.color; paint.font = `${12 / scale}px Arial`; paint.textBaseline = 'bottom'; paint.fillText(`${entry.name}  ${(bounds.width / (r.width * source.width)).toFixed(2)}×`, bounds.left, bounds.top - 3 / scale);
      try { const settings = fitZoomScale(source, r, entry, null); paintBar(bounds, cropFov(source, r), settings); } catch (error) { invalid = error; }
    }
    paint.restore(); apply.disabled = !!invalid;
    if (invalid) errorStatus(status, invalid);
  }
  function addEntry(region = null, extra = {}) { const saved = entries[active] || zoomDefaults; entries.push({ ...saved, ...extra, key: null, name: `放大图 ${Math.max(0, ...entries.map(e => +(e.name.match(/\d+$/)?.[0] || 0))) + 1}`, region: region || { ...zoomDefaults.region }, sourceState: undefined, markerState: undefined, zoomState: undefined, cropSignature: undefined, initialPlacement: undefined }); selectEntry(entries.length - 1); }
  function deleteEntry() { if (active < 0) return; entries.splice(active, 1); selectEntry(Math.min(active, entries.length - 1)); }
  list.onchange = () => selectEntry(+list.value); fields.addEventListener('input', saveFields); add.onclick = () => { if (sources.length) addEntry(); }; remove.onclick = deleteEntry; fit.onclick = () => { zoom = 1; pan = { x: 0, y: 0 }; drawView(); };
  canvas.addEventListener('wheel', event => { if (!sources.length) return; event.preventDefault(); const p = point(event), before = view(), next = Math.max(.2, Math.min(20, zoom * Math.exp(-event.deltaY * .0015))); zoom = next; const after = view(); pan.x += before.x + p.x * sources[0].width * before.scale - after.x - p.x * sources[0].width * after.scale; pan.y += before.y + p.y * sources[0].height * before.scale - after.y - p.y * sources[0].height * after.scale; drawView(); }, { passive: false });
  canvas.addEventListener('keydown', event => { if (event.code === 'Space') { event.preventDefault(); space = true; } if (event.key === 'Delete' || event.key === 'Backspace') { event.preventDefault(); deleteEntry(); } }); canvas.addEventListener('keyup', event => { if (event.code === 'Space') space = false; }); canvas.addEventListener('blur', () => { space = false; });
  const localPoint = (p, r) => { const w = sources[0].width, h = sources[0].height, cx = (r.x + r.width / 2) * w, cy = (r.y + r.height / 2) * h, a = -(r.rotation || 0) * Math.PI / 180, x = p.x * w - cx, y = p.y * h - cy; return { x: x * Math.cos(a) - y * Math.sin(a), y: x * Math.sin(a) + y * Math.cos(a) }; };
  canvas.addEventListener('pointerdown', event => {
    if (!sources.length) return; canvas.focus(); canvas.setPointerCapture(event.pointerId); const p = point(event), v = view();
    if (space || event.button === 1) { drag = { mode: 'pan', clientX: event.clientX, clientY: event.clientY, pan: { ...pan } }; return; }
    const order = [active, ...entries.map((_, i) => i).reverse().filter(i => i !== active)].filter(i => i >= 0);
    for (const i of order) {
      const r = entries[i].region, q = localPoint(p, r), w = r.width * sources[0].width, h = r.height * sources[0].height, tolerance = 8 / v.scale;
      const handles = [[-1, -1], [1, -1], [1, 1], [-1, 1]], corner = handles.findIndex(([x, y]) => Math.abs(q.x - x * w / 2) < tolerance && Math.abs(q.y - y * h / 2) < tolerance);
      if (corner >= 0 || Math.abs(q.x) <= w / 2 && Math.abs(q.y) <= h / 2) { selectEntry(i); drag = { mode: corner >= 0 ? 'resize' : 'move', start: p, initial: { ...r }, corner: handles[corner], entry: i }; return; }
    }
    addEntry({ x: p.x, y: p.y, width: .001, height: .001, rotation: 0 }); drag = { mode: 'draw', start: p, entry: active };
  });
  canvas.addEventListener('pointermove', event => {
    if (!drag) return;
    if (drag.mode === 'pan') { const b = canvas.getBoundingClientRect(); pan.x = drag.pan.x + (event.clientX - drag.clientX) * canvas.width / b.width; pan.y = drag.pan.y + (event.clientY - drag.clientY) * canvas.height / b.height; drawView(); return; }
    const p = point(event), entry = entries[drag.entry];
    if (drag.mode === 'draw') entry.region = { x: Math.min(drag.start.x, p.x), y: Math.min(drag.start.y, p.y), width: Math.max(.001, Math.abs(p.x - drag.start.x)), height: Math.max(.001, Math.abs(p.y - drag.start.y)), rotation: 0 };
    if (drag.mode === 'move') entry.region = { ...drag.initial, x: drag.initial.x + p.x - drag.start.x, y: drag.initial.y + p.y - drag.start.y };
    if (drag.mode === 'resize') {
      const r = drag.initial, a = (r.rotation || 0) * Math.PI / 180, c = Math.cos(a), s = Math.sin(a), sw = sources[0].width, sh = sources[0].height;
      const opposite = { x: (r.x + r.width / 2) * sw - drag.corner[0] * r.width * sw / 2 * c + drag.corner[1] * r.height * sh / 2 * s, y: (r.y + r.height / 2) * sh - drag.corner[0] * r.width * sw / 2 * s - drag.corner[1] * r.height * sh / 2 * c };
      const dx = p.x * sw - opposite.x, dy = p.y * sh - opposite.y, localX = dx * c + dy * s, localY = -dx * s + dy * c;
      const width = Math.max(sw * .001, Math.abs(localX)), height = Math.max(sh * .001, Math.abs(localY)), cx = opposite.x + (localX * c - localY * s) / 2, cy = opposite.y + (localX * s + localY * c) / 2;
      entry.region = { x: (cx - width / 2) / sw, y: (cy - height / 2) / sh, width: width / sw, height: height / sh, rotation: r.rotation || 0 };
    }
    setControls(controls, { ...entry, ...entry.region, scaleLength: entry.scaleLength ?? '' }); draw();
  });
  const endDrag = event => { if (canvas.hasPointerCapture(event.pointerId)) canvas.releasePointerCapture(event.pointerId); drag = null; }; canvas.addEventListener('pointerup', endDrag); canvas.addEventListener('pointercancel', endDrag);
  async function reload({ refreshPreview = true } = {}) {
    load.disabled = true;
    try {
      const context = await ctx.host.context(), items = selected(context), rectangle = items.filter(s => !isImage(s)); sources = items.filter(isImage);
      if (!sources.length || rectangle.length > 1 || rectangle.length && sources.length !== 1 || rectangle.some(s => s.type !== 'shape' || s.shapeType && s.shapeType !== 'rectangle')) throw new Error('请选择一张或多张图片，或一张图片和一个普通矩形。');
      const saved = await ctx.services.request('images/settings', {}); entries = structuredClone(recordsOf(sources[0]));
      for (const entry of entries) { const marker = roles(context, entry.key, 'marker')[0]; if (marker) entry.region = relativeRegion(sources[0], marker); }
      active = entries.length ? 0 : -1;
      if (rectangle[0]) { const existing = entries.findIndex(e => roles(context, e.key, 'marker').some(s => s.id === rectangle[0].id)); if (existing >= 0) active = existing; else { addEntry(relativeRegion(sources[0], rectangle[0]), { ...(saved.zoom || {}), manualId: rectangle[0].id }); } }
      if (!entries.length) setControls(controls, { ...zoomDefaults, ...(saved.zoom || {}), ...zoomDefaults.region });
      selectEntry(active); zoom = 1; pan = { x: 0, y: 0 };
      // Zoom regions never rewrite the source pixels, so a reload after applying keeps the decoded preview on screen.
      const key = sources[0].id;
      if (refreshPreview || key !== previewKey) {
        const ownGeneration = ++imageGeneration, rendered = await ctx.services.request('images/crop', { ...await imageFor(ctx, sources[0]), maximum: 1600, region: { x: 0, y: 0, width: 1, height: 1, rotation: 0 }, displayAspect: sources[0].width / sources[0].height });
        const nextImage = new Image(); nextImage.onload = () => { if (!disposed && imageGeneration === ownGeneration) { image = nextImage; previewKey = key; draw(); } }; nextImage.src = `data:${rendered.mime};base64,${rendered.data}`;
      }
      status.textContent = `已读取 ${sources.length} 张图片、${entries.length} 个区域；新增区域会同步到其余所选图片。`;
    } catch (error) { errorStatus(status, error); } finally { load.disabled = false; }
  }
  load.onclick = () => reload();
  apply.onclick = async () => { apply.disabled = true; try { if (!sources.length) throw new Error('请先读取所选图片。'); saveFields(); entries.forEach(e => validateRegion(e.region)); await applyZoomEntries(ctx, sources, entries); status.textContent = '全部放大区域已更新；幻灯片中的移动、缩放、旋转和删除将自动同步。'; await reload({ refreshPreview: false }); } catch (error) { errorStatus(status, error); } finally { apply.disabled = false; } };
  return () => { disposed = true; imageGeneration++; };
}

async function setImageFov(ctx, shape, fov) {
  const scale = scaleOf(shape);
  if (scale) return applyScaleBar(ctx, shape, fov, scale);
  const commands = [{ op: 'update', id: content(shape).id, properties: { alternativeText: writeRecord(content(shape).alternativeText, FOV_PREFIX, fov) } }];
  if (shape.id !== content(shape).id) commands.push({ op: 'update', id: shape.id, properties: { alternativeText: writeRecord(shape.alternativeText, FOV_PREFIX, fov) } });
  return ctx.host.apply(commands);
}
export const actions = [
  { id: 'science.setFov', label: '设置原图 FOV', section: '科学图片', fields: [{ name: 'width', label: '原图物理宽度（0 为未标定）', type: 'number', default: 0, min: 0 }, { name: 'height', label: '原图物理高度（0 为未标定）', type: 'number', default: 0, min: 0 }, { name: 'unit', label: '单位', type: 'select', default: 'μm', options: unitOptions }], run: (ctx, values) => scientificMutation(ctx, async () => {
    const shapes = pictures(await ctx.host.context()), fov = validFov({ width: +values.width, height: +values.height, unit: values.unit, source: '手动设置' });
    for (const shape of shapes) { const scale = scaleOf(shape); if (scale) scaleLayout(shape, visibleFov(shape, fov), scale); }
    for (const shape of shapes) await setImageFov(ctx, shape, fov);
    return { count: shapes.length };
  }) },
  { id: 'science.readFov', label: '重新读取 TIFF 标定', section: '科学图片', run: ctx => scientificMutation(ctx, async () => {
    const shapes = pictures(await ctx.host.context()), calibrated = [];
    for (const shape of shapes) {
      const sourceTag = tag(shape, 'SLIDESCI_SOURCE_PATH'), originalPath = (content(shape).alternativeText || '').split(/[\r\n]+/).filter(l => l.startsWith('原图路径：')).at(-1)?.slice('原图路径：'.length) || parse(sourceTag, sourceTag);
      // Read the recorded file directly; only fall back to native bytes when the object has no source path.
      let fov;
      if (originalPath) fov = await ctx.services.request('images/fov', { path: originalPath });
      else {
        const image = await imageFor(ctx, shape);
        fov = await ctx.services.request('images/fov', image.source ? { path: image.source } : image);
      }
      if (!fov) throw new Error(`${shape.name} 没有可读取的 TIFF 物理标定。`);
      const scale = scaleOf(shape); if (scale) scaleLayout(shape, visibleFov(shape, fov), scale);
      calibrated.push({ shape, fov });
    }
    for (const { shape, fov } of calibrated) await setImageFov(ctx, shape, fov);
    return { count: shapes.length };
  }) }
];
export const panels = [{ id: 'science.scale', label: 'FOV 与比例尺预览', section: '科学图片', mount: mountScale }, { id: 'science.zoom', label: '多个放大区域', section: '科学图片', mount: mountZoom }];
// Mount this once for the taskpane lifetime, including while scientific panels are closed.
export const initialize = ctx => startScientificReconciliation(ctx);
