export const FOV_PREFIX = 'FOV（SlideSCI）：';
export const SCALE_PREFIX = '比例尺（SlideSCI）：';
export const UUID_TAG = 'SLIDESCI_IMAGE_UUID';
export const RECORDS_TAG = 'SLIDESCI_ZOOM_RECORDS';
export const ZOOM_TAG = 'SLIDESCI_ZOOM_RECORD';
export const SCALE_OWNER = 'SLIDESCI_SCALE_OWNER';
export const uuid = () => crypto.randomUUID().toUpperCase();
export const units = { nm: .001, 'μm': 1, um: 1, mm: 1000, cm: 10000, m: 1e6, in: 25400 };
export const unitOptions = Object.keys(units).filter(x => x !== 'um').map(value => ({ value, label: value }));
export const positive = x => Number.isFinite(+x) && +x > 0;
export function tag(shape, name) { const key = Object.keys(shape?.tags || {}).find(k => k.toUpperCase() === name.toUpperCase()); return key ? shape.tags[key] : undefined; }
export const serializeTag = value => 'HEX:' + Array.from(new TextEncoder().encode(JSON.stringify(value)), byte => byte.toString(16).padStart(2, '0')).join('').toUpperCase();
export function parse(value, fallback = null) {
  if (!value) return fallback;
  try {
    if (value.startsWith('HEX:')) { const hex = value.slice(4); if (hex.length % 2 || /[^0-9A-F]/i.test(hex)) return fallback; value = new TextDecoder().decode(Uint8Array.from(hex.match(/../g) || [], pair => parseInt(pair, 16))); }
    return JSON.parse(value);
  } catch { return fallback; }
}
export function content(shape) { return shape?.children?.find(s => tag(s, 'SLIDESCI_SCALE_CONTENT') === '1') || shape; }
export const isImage = s => s?.type === 'picture' || tag(s, 'SLIDESCI_SCALE_GROUP') === '1' || s?.fill?.type === 'picture';
export const flatten = shapes => shapes.flatMap(s => [s, ...flatten(s.children || [])]);
export function selected(context) { const all = flatten(context.shapes); return context.selectionIds.map(id => all.find(s => s.id === id)).filter(Boolean); }
export function pictures(context, min = 1, max = Infinity) { const items = selected(context); if (items.length < min || items.length > max || items.some(s => !isImage(s))) throw new Error(max === Infinity ? '请选择一张或多张图片或图片比例尺组合。' : `请选择 ${min}–${max} 张图片或图片比例尺组合。`); return items; }
export function readRecord(shape, prefix) {
  const line = (content(shape)?.alternativeText || shape?.alternativeText || '').split(/[\r\n]+/).filter(x => x.startsWith(prefix)).at(-1);
  const record = line ? parse(line.slice(prefix.length)) : null;
  if (!record) return null;
  // Windows documents use PascalCase JSON property names.
  return Object.fromEntries(Object.entries(record).map(([key, value]) => [key[0].toLowerCase() + key.slice(1), value]));
}
export function writeRecord(text, prefix, value) { const lines = (text || '').split(/\r\n|[\r\n]/).filter(x => !x.startsWith(prefix)); while (lines.at(-1) === '') lines.pop(); if (value) lines.push(prefix + JSON.stringify(value)); return lines.join('\n'); }
export function fovOf(shape) { return readRecord(shape, FOV_PREFIX); }
// Windows native intersections store an already-visible FOV and a source calibration stamp.
// Mac derived rasters store full-raster FOV, including older Mac pictures with only FOV_VISIBLE.
export function visibleFov(shape, fov = fovOf(shape)) {
  if (!fov || !units[fov.unit]) return null;
  const image = content(shape);
  if (tag(image, 'SLIDESCI_FOV_VISIBLE') === '1' && tag(image, 'SLIDESCI_FOV_SOURCE')) {
    return { width: +fov.width * units[fov.unit], height: +fov.height * units[fov.unit] };
  }
  const crop = image?.crop || {};
  return { width: +fov.width * units[fov.unit] * (1 - (crop.left || 0) - (crop.right || 0)), height: +fov.height * units[fov.unit] * (1 - (crop.top || 0) - (crop.bottom || 0)) };
}
export function validFov(fov) { if (!fov || !units[fov.unit] || ![+fov.width, +fov.height].every(x => Number.isFinite(x) && x >= 0) || !(+fov.width > 0 || +fov.height > 0)) throw new Error('请填写至少一个大于 0 的 FOV 尺寸，并选择物理长度单位。'); return fov; }
export const defaultScale = { vertical: false, widthMicrometers: 50, heightMicrometers: 50, lengthUnit: 'μm', thicknessPoints: 3, showText: true, barColor: '#ffffff', fontColor: '#ffffff', fontSize: 14, fontName: 'Arial', bold: true, corner: 'BottomRight', groupWithImage: true };
const argbHex = value => '#' + (Number(value) >>> 0).toString(16).padStart(8, '0').slice(-6);
export function scaleOf(shape) {
  const s = readRecord(shape, SCALE_PREFIX) || readRecord(shape, '标尺（SlideSCI）：');
  if (!s) return null;
  return { ...defaultScale, ...s, corner: typeof s.corner === 'number' ? ['TopLeft', 'BottomLeft', 'TopRight', 'BottomRight'][s.corner] : s.corner || defaultScale.corner, barColor: s.barColor || (s.barColorArgb == null ? defaultScale.barColor : argbHex(s.barColorArgb)), fontColor: s.fontColor || (s.fontColorArgb == null ? defaultScale.fontColor : argbHex(s.fontColorArgb)) };
}
let measureDocument, measureContext;
export function textSize(text, font) {
  if (measureDocument !== document) {
    measureContext = document.createElement('canvas').getContext('2d');
    measureDocument = document;
  }
  const ctx = measureContext;
  ctx.font = `${font.bold ? 'bold ' : ''}${font.size * 96 / 72}px "${font.name.replaceAll('"', '')}"`;
  const m = ctx.measureText(text);
  return { width: (m.width + 2) * 72 / 96, height: Math.max(font.size * 1.18, ((m.actualBoundingBoxAscent || 0) + (m.actualBoundingBoxDescent || 0)) * 72 / 96) };
}
// Six significant digits keep small physical lengths (4 nm, 10 nm) from collapsing to zero while
// trimming binary floating-point noise such as 0.30000000000000004.
export function formatAmount(value) {
  const amount = +value;
  if (!Number.isFinite(amount)) return '';
  if (amount === 0) return '0';
  const magnitude = Math.floor(Math.log10(Math.abs(amount)));
  const decimals = Math.min(20, Math.max(0, 5 - magnitude));
  const fixed = amount.toFixed(decimals);
  const trimmed = fixed.includes('.') ? fixed.replace(/0+$/, '').replace(/\.$/, '') : fixed;
  return +trimmed === 0 ? (+amount.toPrecision(6)).toString() : trimmed;
}
export function scaleLayout(image, fov, settings, measured) {
  const s = { ...defaultScale, ...settings }, amount = +(s.vertical ? s.heightMicrometers : s.widthMicrometers);
  if (![+s.widthMicrometers, +s.heightMicrometers].every(x => Number.isFinite(x) && x >= 0) || !positive(s.thicknessPoints) || !positive(s.fontSize) || +s.fontSize > 400 || !units[s.lengthUnit] || !['TopLeft', 'BottomLeft', 'TopRight', 'BottomRight'].includes(s.corner)) throw new Error('比例尺参数无效。');
  const label = `${formatAmount(amount / units[s.lengthUnit])}${s.lengthUnit}`;
  if (!amount) return { label, bar: null, text: null };
  const field = s.vertical ? fov?.height : fov?.width;
  if (!positive(field)) throw new Error(s.vertical ? '纵向比例尺需要填写 FOV 高度。' : '横向比例尺需要填写 FOV 宽度。');
  const length = amount * (s.vertical ? image.height : image.width) / field;
  const w = s.vertical ? +s.thicknessPoints : length, h = s.vertical ? length : +s.thicknessPoints;
  const t = measured || textSize(label, { name: s.fontName, size: +s.fontSize, bold: s.bold });
  const margin = Math.min(8, Math.min(image.width, image.height) * .04), gap = 3;
  const blockW = s.vertical ? w + (s.showText ? gap + t.width : 0) : Math.max(w, s.showText ? t.width : 0);
  const blockH = s.vertical ? Math.max(h, s.showText ? t.height : 0) : h + (s.showText ? gap + t.height : 0);
  if (blockW + margin * 2 > image.width || blockH + margin * 2 > image.height) throw new Error('比例尺或文字超出图片范围，请减小长度、厚度或字号。');
  const right = s.corner.endsWith('Right'), bottom = s.corner.startsWith('Bottom');
  const x = right ? image.width - margin - w : margin, y = bottom ? image.height - margin - h : margin;
  return { label, bar: { left: x, top: y, width: w, height: h }, text: s.showText ? { left: s.vertical ? right ? x - gap - t.width : x + w + gap : right ? x + w - t.width : x, top: s.vertical ? bottom ? y + h - t.height : y : bottom ? y - gap - t.height : y + h + gap, width: t.width, height: t.height } : null };
}
export function rotate(point, center, degrees) { const a = degrees * Math.PI / 180, c = Math.cos(a), s = Math.sin(a), x = point.x - center.x, y = point.y - center.y; return { x: center.x + x * c - y * s, y: center.y + x * s + y * c }; }
export function localBounds(source, bounds, rotation = source.rotation || 0) {
  const p = rotate({ x: source.left + bounds.left + bounds.width / 2, y: source.top + bounds.top + bounds.height / 2 }, { x: source.left + source.width / 2, y: source.top + source.height / 2 }, rotation);
  return { left: p.x - bounds.width / 2, top: p.y - bounds.height / 2, width: bounds.width, height: bounds.height, rotation };
}
export function markerBounds(source, region) { return { ...localBounds(source, { left: region.x * source.width, top: region.y * source.height, width: region.width * source.width, height: region.height * source.height }), rotation: (source.rotation || 0) + (region.rotation || 0) }; }
export function relativeRegion(source, marker) {
  const p = rotate({ x: marker.left + marker.width / 2, y: marker.top + marker.height / 2 }, { x: source.left + source.width / 2, y: source.top + source.height / 2 }, -(source.rotation || 0));
  return { x: (p.x - source.left - marker.width / 2) / source.width, y: (p.y - source.top - marker.height / 2) / source.height, width: marker.width / source.width, height: marker.height / source.height, rotation: (marker.rotation || 0) - (source.rotation || 0) };
}
export function corners(shape) { return [{ x: shape.left, y: shape.top }, { x: shape.left + shape.width, y: shape.top }, { x: shape.left + shape.width, y: shape.top + shape.height }, { x: shape.left, y: shape.top + shape.height }].map(p => rotate(p, { x: shape.left + shape.width / 2, y: shape.top + shape.height / 2 }, shape.rotation || 0)); }
export function clipLine(start, end, source) {
  const center = { x: source.left + source.width / 2, y: source.top + source.height / 2 }, a = rotate(start, center, -(source.rotation || 0)), b = rotate(end, center, -(source.rotation || 0));
  const dx = b.x - a.x, dy = b.y - a.y; let first = 0, last = 1;
  for (const [p, q] of [[-dx, a.x - source.left], [dx, source.left + source.width - a.x], [-dy, a.y - source.top], [dy, source.top + source.height - a.y]]) {
    if (Math.abs(p) < 1e-8) { if (q < 0) return null; continue; }
    const r = q / p;
    if (p < 0) first = Math.max(first, r); else last = Math.min(last, r);
    if (first > last) return null;
  }
  return [rotate({ x: a.x + dx * first, y: a.y + dy * first }, center, source.rotation || 0), rotate({ x: a.x + dx * last, y: a.y + dy * last }, center, source.rotation || 0)];
}
export function guideSegments(source, marker, zoom, extent = 'across') {
  const a = corners(marker), b = corners(zoom), dx = (zoom.left + zoom.width / 2 - source.left - source.width / 2) / source.width, dy = (zoom.top + zoom.height / 2 - source.top - source.height / 2) / source.height;
  const pairs = Math.abs(dx) >= Math.abs(dy) ? dx >= 0 ? [[a[1], b[0]], [a[2], b[3]]] : [[a[0], b[1]], [a[3], b[2]]] : dy >= 0 ? [[a[3], b[0]], [a[2], b[1]]] : [[a[0], b[3]], [a[1], b[2]]];
  return extent === 'inside' ? pairs.map(p => clipLine(...p, source)) : pairs;
}
export function lineBounds(segment) {
  if (!segment) return { left: 0, top: 0, width: 0, height: 0, line: { visible: false } };
  const [a, b] = segment, length = Math.hypot(b.x - a.x, b.y - a.y);
  return { left: (a.x + b.x - length) / 2, top: (a.y + b.y) / 2, width: length, height: 0, rotation: Math.atan2(b.y - a.y, b.x - a.x) * 180 / Math.PI };
}
export function geometry(shape) { return { left: shape.left, top: shape.top, width: shape.width, height: shape.height, rotation: shape.rotation || 0 }; }
export function near(a, b, epsilon = .02) {
  return !!a && !!b && Object.keys(a).every(key => {
    if (typeof a[key] !== 'number') return a[key] === b[key];
    const delta = key === 'rotation' ? ((a[key] - (b[key] || 0) + 540) % 360 + 360) % 360 - 180 : a[key] - (b[key] || 0);
    return Math.abs(delta) <= epsilon;
  });
}
export function imageIdentity(image) {
  if (!image) return null;
  let digest = image.digest || image.revision;
  if (!digest && image.data) { let hash = 2166136261; for (let i = 0; i < image.data.length; i++) hash = Math.imul(hash ^ image.data.charCodeAt(i), 16777619); digest = `${image.data.length}:${hash >>> 0}`; }
  return { mime: image.mime, width: image.width, height: image.height, digest };
}
export const element = (name, text, props = {}) => { const node = document.createElement(name); if (text != null) node.textContent = text; Object.assign(node, props); return node; };
export function control(parent, label, type, value, options) { const wrap = element('label', label), node = element(type === 'select' ? 'select' : 'input', null, type === 'select' ? {} : { type }); if (options) for (const item of options) node.append(element('option', item.label ?? item, { value: item.value ?? item })); if (type === 'checkbox') node.checked = !!value; else node.value = value ?? ''; wrap.append(node); parent.append(wrap); return node; }
export async function imageFor(ctx, shape) {
  const result = await ctx.host.images([content(shape).id], { original: true });
  const image = result[0];
  return prepareImage(ctx, shape, image);
}
export async function imagesFor(ctx, shapes) {
  if (!shapes.length) return [];
  const ids = [...new Set(shapes.map(shape => String(content(shape).id)))];
  const originals = await ctx.host.images(ids, { original: true });
  const byId = new Map(originals.map(image => [String(image.id), image]));
  const images = [];
  for (const shape of shapes) images.push(await prepareImage(ctx, shape, byId.get(String(content(shape).id))));
  return images;
}
async function prepareImage(ctx, shape, image) {
  if (!image?.data && !image?.path) throw new Error('无法读取原始图片字节。');
  const crop = content(shape).crop || image.crop;
  if (/image\/(?:x-)?(?:emf|wmf)|image\/(?:x-ms-)?bmp/i.test(image.mime || '') || /\.(?:emf|wmf|bmp)$/i.test(image.source || '')) {
    if (typeof ctx.host.rasterizeImage !== 'function') throw new Error('当前宿主缺少原始 BMP/EMF/WMF 图片的原生栅格导出接口。');
    const raster = await ctx.host.rasterizeImage(content(shape).id, { original: true, rotation: 0, maximum: 4096 });
    return { ...image, ...raster, crop };
  }
  if (/tiff/i.test(image.mime || '') || /^(?:SUkq|SUsr|TU0A)/.test(image.data || '')) {
    const normalized = await ctx.services.request('images/normalize', image);
    return { ...image, ...normalized, crop };
  }
  return { ...image, crop };
}
export function findUUID(context, key) { return context.shapes.find(s => tag(s, UUID_TAG) === key) || flatten(context.shapes).find(s => tag(s, UUID_TAG) === key); }
export async function ensureUUID(ctx, shape) { const key = tag(shape, UUID_TAG) || uuid(); if (!tag(shape, UUID_TAG)) await ctx.host.apply([{ op: 'update', id: shape.id, properties: { tags: { [UUID_TAG]: key } } }], { readback: false }); return key; }
