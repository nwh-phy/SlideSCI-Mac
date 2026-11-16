import {
  TRUE, FALSE, items, optional, readTags, writeTags, shapeIndex,
  findNative, findNativeEntry, readCrop, readFont, setCrop, properties, zOrder,
  shapeCodes, sameText,
} from './wps-host.js';
import { alignment } from './features/editing-geometry.js';

// WPS 演示的 JS API 是与 PowerPoint COM 同构的同步接口。Every document
// mutation and every SaveCopyAs snapshot stays on one promise chain so native
// collections cannot be mutated while a temporary presentation is open.
const PP_SELECTION_TEXT = 3;
const PP_SELECTION_SHAPES = 2;
const PP_SELECTION_SLIDES = 1;
const SAVE_AS_OPENXML_PRESENTATION = 24;
const FIXED_FORMAT_PDF = 2;
const BLANK_LAYOUT = 12;
const TEXT_ORIENTATION_HORIZONTAL = 1;
const MAX_RASTER_DPI = 2400;
const POLL_INTERVAL = 700;
// 位置/名称/文字的变化用轻量指纹每轮检测；填充、轮廓、字体细节等富格式
// 只按这个轮次间隔做一次完整扫描，避免每 0.7 秒重读整页富格式。
const DEEP_POLL_INTERVAL = 4;

const MIME_EXTENSIONS = {
  'image/png': 'png', 'image/jpeg': 'jpg', 'image/gif': 'gif', 'image/tiff': 'tif',
  'image/bmp': 'bmp', 'image/svg+xml': 'svg', 'image/x-emf': 'emf', 'image/emf': 'emf',
  'image/x-wmf': 'wmf', 'image/wmf': 'wmf', 'application/pdf': 'pdf',
  'application/vnd.openxmlformats-officedocument.presentationml.presentation': 'pptx',
  'video/mp4': 'mp4', 'video/quicktime': 'mov', 'video/webm': 'webm', 'video/x-msvideo': 'avi',
};
const IMAGE_FORMATS = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', bmp: 'image/bmp' };
const EXPORT_FILTERS = { png: 'PNG', jpg: 'JPG', jpeg: 'JPG', bmp: 'BMP' };
// 只改位置、尺寸与旋转的更新不需要 PPTX 快照，也不影响文字、图片与标记。
const GEOMETRY_KEYS = new Set(['left', 'top', 'width', 'height', 'rotation']);
function geometryProperties(value) {
  const keys = Object.keys(value || {});
  return keys.length > 0 && keys.every(key => GEOMETRY_KEYS.has(key));
}
function nativeShapeKind(type) {
  const value = Number(type);
  if (value === 6) return 'group';
  if (value === 11 || value === 13) return 'picture';
  if (value === 16) return 'media';
  if (value === 19) return 'table';
  if (value === 14 || value === 17) return 'text';
  return 'shape';
}

function boolean(value) { return value === TRUE || value === 1 || value === true; }
function round(value) { return Math.round(Number(value || 0) * 100) / 100; }
function clamp(value, minimum, maximum) { return Math.min(maximum, Math.max(minimum, value)); }
function canonical(value) { return String(value).replace(/\r\n|[\r\n]/g, '\r'); }
function clone(value) { return value === undefined ? undefined : structuredClone(value); }
function extensionFor(mime, fallback = 'bin') { return MIME_EXTENSIONS[String(mime || '').toLowerCase()] || fallback; }
function safeStem(value, fallback = 'SlideSCI') {
  const stem = String(value || fallback).trim().replace(/[\x00-\x1f/\\:*?"<>|]/gu, '_')
    .replace(/\.(pptx|ppt|pdf|png|jpe?g|bmp|svg)$/iu, '');
  return stem && stem !== '.' && stem !== '..' ? stem.slice(0, 180) : fallback;
}
function shapeCode(name) {
  const direct = shapeCodes[name];
  if (direct !== undefined) return direct;
  const normalized = Object.fromEntries(Object.entries(shapeCodes).map(([key, value]) => [key.toLowerCase(), value]));
  return normalized[String(name).toLowerCase()];
}
function isPicture(shape) { return Number(shape.Type) === 13 || Number(shape.Type) === 11; }
function isMedia(shape) { return Number(shape.Type) === 16; }
function isGroup(shape) { return Number(shape.Type) === 6; }
function isVideoMime(mime) { return /^(video|audio)\//i.test(String(mime || '')); }
function asBase64Bytes(data) {
  const binary = atob(String(data));
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index++) bytes[index] = binary.charCodeAt(index);
  return bytes;
}
async function digest(data) {
  const value = String(data || '');
  if (!value) return undefined;
  try {
    if (globalThis.crypto?.subtle) {
      const bytes = asBase64Bytes(value);
      const hash = new Uint8Array(await crypto.subtle.digest('SHA-256', bytes));
      return 'sha256:' + Array.from(hash, byte => byte.toString(16).padStart(2, '0')).join('');
    }
  } catch { /* The deterministic fallback below still yields a stable digest. */ }
  let hash = 2166136261;
  for (let index = 0; index < value.length; index++) hash = Math.imul(hash ^ value.charCodeAt(index), 16777619);
  return 'fnv1a:' + (hash >>> 0).toString(16) + ':' + value.length;
}
function intrinsicSize(data) {
  if (!data) return undefined;
  try {
    const bytes = asBase64Bytes(data);
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    if (bytes.length >= 26 && bytes[0] === 0x42 && bytes[1] === 0x4d) {
      return { width: Math.abs(view.getInt32(18, true)), height: Math.abs(view.getInt32(22, true)) };
    }
    if (bytes.length >= 40 && view.getUint32(0, true) === 1 && view.getUint32(40, true) === 0x464d4520) {
      return { width: Math.abs(view.getInt32(24, true) - view.getInt32(32, true)) / 100 * 72 / 25.4,
        height: Math.abs(view.getInt32(28, true) - view.getInt32(36, true)) / 100 * 72 / 25.4 };
    }
    if (bytes.length >= 22 && view.getUint32(0, true) === 0x9ac6cdd7) {
      return { width: Math.abs(view.getInt16(6, true) - view.getInt16(10, true)) / 20,
        height: Math.abs(view.getInt16(8, true) - view.getInt16(12, true)) / 20 };
    }
    const png = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];
    if (bytes.length >= 24 && png.every((value, index) => bytes[index] === value)) {
      return { width: view.getUint32(16), height: view.getUint32(20) };
    }
    if (bytes.length >= 10 && bytes[0] === 0x47 && bytes[1] === 0x49 && bytes[2] === 0x46) {
      return { width: view.getUint16(6, true), height: view.getUint16(8, true) };
    }
    if (bytes.length > 4 && bytes[0] === 0xff && bytes[1] === 0xd8) {
      let offset = 2;
      while (offset + 8 < bytes.length) {
        if (bytes[offset] !== 0xff) { offset++; continue; }
        const marker = bytes[offset + 1]; offset += 2;
        if (marker === 0xd8 || marker === 0xd9) continue;
        const length = view.getUint16(offset, false);
        if ([0xc0, 0xc1, 0xc2, 0xc3, 0xc5, 0xc6, 0xc7, 0xc9, 0xca, 0xcb, 0xcd, 0xce, 0xcf].includes(marker)) {
          return { width: view.getUint16(offset + 5, false), height: view.getUint16(offset + 3, false) };
        }
        if (length < 2) break;
        offset += length;
      }
    }
    const text = bytes.length < 200000 ? new TextDecoder().decode(bytes) : '';
    const svg = /<svg\b([^>]*)>/i.exec(text);
    if (svg) {
      const attributes = svg[1];
      const width = /\bwidth\s*=\s*"([\d.]+)(?:pt|px)?"/i.exec(attributes);
      const height = /\bheight\s*=\s*"([\d.]+)(?:pt|px)?"/i.exec(attributes);
      if (width && height) return { width: Number(width[1]), height: Number(height[1]) };
      const box = /\bviewBox\s*=\s*"[\d.\-+]+\s+[\d.\-+]+\s+([\d.]+)\s+([\d.]+)"/i.exec(attributes);
      if (box) return { width: Number(box[1]), height: Number(box[2]) };
    }
  } catch { /* Unsupported container; callers can still rasterize through WPS. */ }
  return undefined;
}
async function decodedSize(data, mime) {
  if (typeof createImageBitmap !== 'function') return undefined;
  try {
    const bitmap = await createImageBitmap(new Blob([asBase64Bytes(data)], { type: mime || 'application/octet-stream' }));
    const size = { width: bitmap.width, height: bitmap.height };
    bitmap.close?.();
    return size;
  } catch { return undefined; }
}
function parsedMap(shapes) {
  const byId = new Map(), byPath = new Map();
  for (const shape of shapes || []) {
    byId.set(String(shape.id), shape);
    if (shape.ref?.shapePath) byPath.set(shape.ref.shapePath.join('.'), shape);
  }
  return { byId, byPath };
}
function mergeFont(base, override) {
  if (!base) return override;
  if (!override) return base;
  const result = { ...base };
  for (const [key, value] of Object.entries(override)) {
    if (value === undefined) continue;
    if (value && typeof value === 'object' && !Array.isArray(value) && result[key] && typeof result[key] === 'object' && !Array.isArray(result[key])) {
      result[key] = mergeFont(result[key], value);
    } else result[key] = value;
  }
  return result;
}
function mergeRuns(nativeRuns, parsedRuns) {
  if (!nativeRuns?.length) return parsedRuns;
  if (!parsedRuns?.length || parsedRuns.length !== nativeRuns.length) return nativeRuns;
  return nativeRuns.map((run, index) => ({ ...parsedRuns[index], ...run, font: mergeFont(parsedRuns[index]?.font, run.font) }));
}
function mergeParagraphs(nativeParagraphs, parsedParagraphs) {
  if (!nativeParagraphs?.length) return parsedParagraphs;
  return nativeParagraphs.map((paragraph, index) => {
    const parsed = parsedParagraphs?.[index] || {};
    return { ...parsed, ...paragraph, runs: mergeRuns(paragraph.runs, parsed.runs),
      bullet: paragraph.bullet !== undefined ? paragraph.bullet : parsed.bullet };
  });
}
function mergeParsedShape(native, parsed) {
  const item = { ...native };
  if (!parsed) {
    item.style = item.style || { fill: item.fill, line: item.line, font: item.font, textFrame: item.textFrame };
    return item;
  }
  const type = parsed.type && native.type === 'shape' ? parsed.type : native.type;
  item.type = type;
  item.shapeType = native.shapeType ?? parsed.shapeType;
  if (native.parentId) {
    for (const key of ['left', 'top', 'width', 'height', 'rotation']) if (Number.isFinite(parsed[key])) item[key] = parsed[key];
  }
  if (parsed.fill?.type) {
    if (parsed.fill.type === 'gradient') item.fill = { ...(item.fill || {}), ...parsed.fill };
    else if (!item.fill || item.fill.type === 'none' || item.fill.type === undefined) item.fill = parsed.fill;
    else item.fill = { ...parsed.fill, ...item.fill };
  }
  if (parsed.line) item.line = { ...parsed.line, ...item.line };
  if (type === 'picture') item.crop = parsed.crop ?? item.crop;
  if (parsed.image) item.image = { ...(item.image || {}), ...parsed.image };
  if (parsed.media) item.media = { ...(item.media || {}), ...parsed.media };
  item.textFrame = { ...parsed.textFrame, ...item.textFrame };
  item.paragraphs = mergeParagraphs(item.paragraphs, parsed.paragraphs);
  item.runs = mergeRuns(item.runs, parsed.runs);
  item.font = mergeFont(parsed.font, item.font);
  if (item.text === undefined && parsed.text !== undefined) item.text = parsed.text;
  if (!item.alternativeText && parsed.alternativeText) item.alternativeText = parsed.alternativeText;
  item.tags = { ...parsed.tags, ...item.tags };
  if (parsed.title) item.title = parsed.title;
  if (item.isPlaceholder === undefined) item.isPlaceholder = parsed.isPlaceholder;
  item.style = { fill: item.fill, line: item.line, font: item.font, textFrame: item.textFrame, shapeType: item.shapeType };
  if (type === 'group') item.children = mergeShapes(item.children || [], parsed.children || []);
  return item;
}
function mergeShapes(nativeShapes, parsedShapes) {
  const { byId, byPath } = parsedMap(parsedShapes);
  return nativeShapes.map(shape => mergeParsedShape(shape, byId.get(String(shape.id)) || byPath.get(shape.ref.shapePath.join('.'))));
}
function flattenShapes(shapes) { return shapes.flatMap(shape => [shape, ...flattenShapes(shape.children || [])]); }
function boundsOf(shapes) {
  if (!shapes.length) return { left: 0, top: 0, width: 0, height: 0 };
  const left = Math.min(...shapes.map(shape => shape.left)), top = Math.min(...shapes.map(shape => shape.top));
  return { left, top, width: Math.max(...shapes.map(shape => shape.left + shape.width)) - left,
    height: Math.max(...shapes.map(shape => shape.top + shape.height)) - top };
}

export function connectWps(services = {}) {
  if (!window.Application) throw new Error('WPS 演示 JS 接口不可用。请从 WPS 的 SlideSCI 选项卡打开。');
  const app = window.Application;
  let queue = Promise.resolve();
  let mutationDepth = 0;
  let suppressDepth = 0;
  let watching = false;
  let pollTimer;
  let dispatchTimer;
  let lastCheap, lastDeep, deepPolls = 0;
  let pictureProbeAt = 0, pictureProbeScope = '', pictureRevisions = [];
  const listeners = new Set();
  const selectionMemory = new Map();
  const eventHandlers = [];
  const temporaryAssets = new Set();

  function serial(operation) {
    const result = queue.then(operation, operation);
    queue = result.catch(() => {});
    return result;
  }
  function muted() { return mutationDepth > 0 || suppressDepth > 0; }
  async function suppressed(operation) {
    suppressDepth++;
    try { return await operation(); }
    finally { suppressDepth--; }
  }
  async function tracked(operation) {
    return serial(async () => {
      try { return await operation(); }
      finally {
        try { await release([...temporaryAssets]); } catch { /* exported or inserted media is already embedded */ }
        temporaryAssets.clear();
      }
    });
  }
  function request(namespace, payload) {
    if (typeof services?.request !== 'function') throw new Error('WPS 宿主缺少本地服务连接，请重新打开 SlideSCI。');
    return services.request(namespace, payload);
  }
  function nativeExecute(action, payload = {}) { return request('native/execute', { host: 'wps', action, ...payload }); }
  function release(paths) {
    const unique = [...new Set((paths || []).filter(Boolean))];
    if (!unique.length) return Promise.resolve();
    return nativeExecute('releaseFiles', { paths: unique });
  }
  async function withTemporary(extension, operation) {
    const reserved = await nativeExecute('temporaryPath', { extension });
    if (!reserved?.path) throw new Error('无法创建 WPS 临时文件。');
    let result, failure;
    try { result = await operation(reserved.path); }
    catch (error) { failure = error; }
    try { await release([reserved.path]); }
    catch (error) { failure = failure ? new AggregateError([failure, error], '临时文件清理失败。') : error; }
    if (failure) throw failure;
    return result;
  }
  async function materialize(value) {
    if (value.path) return { path: value.path };
    if (!value.data) throw new Error('缺少图片数据或本地文件路径。');
    const file = await request('files/materialize', { data: value.data, mime: value.mime || 'image/png', name: value.name, host: 'wps' });
    if (!file?.path) throw new Error('无法将图片写入 WPS 临时目录。');
    temporaryAssets.add(file.path);
    return file;
  }
  async function readFile(path) {
    const file = await nativeExecute('readFile', { path });
    if (!file?.data) throw new Error('WPS 导出的文件为空。');
    return file;
  }
  function activeData() {
    const presentation = app.ActivePresentation;
    if (!presentation) throw new Error('请先打开演示文稿。');
    const window = app.ActiveWindow;
    if (!window) throw new Error('请先打开演示文稿窗口。');
    let slide = optional(() => window.View.Slide);
    if (!slide) slide = optional(() => window.Selection.SlideRange.Item(1));
    if (!slide && Number(presentation.Slides.Count) > 0) slide = presentation.Slides.Item(1);
    if (!slide) throw new Error('当前演示文稿没有幻灯片。');
    return { presentation, window, slide, slideIndex: Number(slide.SlideIndex) - 1 };
  }
  function presentationKey(presentation) { return String(optional(() => presentation.Name) || 'WPS'); }
  function slideAt(presentation, index) {
    const value = Number(index);
    if (!Number.isInteger(value) || value < 0 || value >= Number(presentation.Slides.Count)) throw new Error(`幻灯片索引无效：${index}`);
    return presentation.Slides.Item(value + 1);
  }
  function selectionState(window) {
    const selection = window.Selection;
    const type = Number(selection.Type);
    if (type === PP_SELECTION_SHAPES) {
      let range = optional(() => boolean(optional(() => selection.HasChildShapeRange)) ? selection.ChildShapeRange : selection.ShapeRange);
      if (!range) range = optional(() => selection.ShapeRange);
      if (!range) return { type, ids: [], shapes: [] };
      const shapes = items(range);
      return { type, ids: shapes.map(shape => String(shape.Id)), shapes };
    }
    if (type === PP_SELECTION_TEXT) {
      const shape = optional(() => selection.ShapeRange.Item(1));
      const range = optional(() => selection.TextRange);
      if (!shape || !range) return { type, ids: [], shapes: [], textSelection: null };
      const start = Math.max(0, Number(range.Start) - 1), length = Math.max(0, Number(range.Length));
      const second = optional(() => shape.TextFrame2.TextRange.Characters(start + 1, length).Font);
      return { type, ids: [String(shape.Id)], shapes: [shape],
        textSelection: { shapeId: String(shape.Id), start, length, text: String(range.Text),
          font: readFont(range.Font, second) } };
    }
    if (type === PP_SELECTION_SLIDES) {
      const range = optional(() => selection.SlideRange);
      return { type, ids: [], shapes: [], slideIndices: range ? items(range).map(slide => Number(slide.SlideIndex) - 1) : [] };
    }
    return { type, ids: [], shapes: [] };
  }
  function orderedSelection(key, ids) {
    const current = new Set(ids);
    const previous = selectionMemory.get(key) || [];
    const order = previous.filter(id => current.has(id));
    for (const id of ids) if (!order.includes(id)) order.push(id);
    if (selectionMemory.size > 24 && !selectionMemory.has(key)) selectionMemory.delete(selectionMemory.keys().next().value);
    selectionMemory.set(key, order);
    return order;
  }
  // 快速路径只读当前原生选区：文字选区取所属形状且不读字体，
  // 形状选区保留 ChildShapeRange（组内选中）语义。childRange 为 undefined
  // 表示当前 WPS 构建没有实现 HasChildShapeRange，需要形状树兜底判断。
  function liveSelection(window) {
    const selection = window.Selection;
    const type = Number(selection.Type);
    if (type === PP_SELECTION_TEXT) {
      const childRange = optional(() => boolean(selection.HasChildShapeRange));
      const shape = optional(() => selection.ShapeRange.Item(1));
      return { type, childRange, ids: shape ? [String(shape.Id)] : [], shapes: shape ? [shape] : [] };
    }
    if (type === PP_SELECTION_SHAPES) {
      const childRange = optional(() => boolean(selection.HasChildShapeRange));
      let range = optional(() => childRange === true ? selection.ChildShapeRange : selection.ShapeRange);
      if (!range) range = optional(() => selection.ShapeRange);
      const shapes = range ? items(range) : [];
      return { type, childRange, ids: shapes.map(shape => String(shape.Id)), shapes };
    }
    if (type === PP_SELECTION_SLIDES) {
      const range = optional(() => selection.SlideRange);
      return { type, childRange: false, ids: [], shapes: [], slideIndices: range ? items(range).map(slide => Number(slide.SlideIndex) - 1) : [] };
    }
    return { type, childRange: false, ids: [], shapes: [] };
  }
  // 组内选中判定：先信任 HasChildShapeRange，再试 ParentGroup，最后用形状树兜底
  // （兜底时把树一并返回，避免同一次调用再遍历一次）。WPS 原生子形状坐标与富模型
  // （mergeParsedShape 用解析子坐标系覆盖）不一致，组内选中不能拿原生坐标做页面数学。
  function selectionGroupState(slide, selection) {
    if (selection.childRange !== undefined) return { group: selection.childRange };
    if ((selection.shapes || []).some(shape => Boolean(optional(() => shape.ParentGroup)))) return { group: true };
    const wanted = new Set((selection.ids || []).map(String));
    if (!wanted.size) return { group: false };
    const nodes = geometryTree(slide);
    return { group: nodes.some(node => wanted.has(node.id) && node.ancestors.length > 0), nodes };
  }
  function selectionOrder(active, selection) {
    return orderedSelection(presentationKey(active.presentation) + '\u0000' + active.slideIndex, selection.ids || []);
  }
  function presentationTitle(presentation) { return String(optional(() => presentation.Name) || 'WPS 演示'); }
  function pageSize(presentation) {
    const width = Number(optional(() => presentation.PageSetup.SlideWidth)) || 0;
    const height = Number(optional(() => presentation.PageSetup.SlideHeight)) || 0;
    if (!(width > 0 && height > 0)) throw new Error('无法读取 WPS 幻灯片尺寸。');
    return { width, height };
  }
  // 只读 Id/Type 建出形状树（供几何读取定位父子关系），不触碰富格式。
  function geometryTree(slide) {
    const nodes = [];
    const visit = (collection, path, parentId, ancestors) => {
      items(collection).forEach((native, offset) => {
        const nativeType = Number(native.Type);
        const node = { native, nativeType, id: String(native.Id), path: [...path, offset + 1], parentId, ancestors };
        nodes.push(node);
        if (nativeType === 6) visit(native.GroupItems, node.path, node.id, [...ancestors, node.id]);
      });
    };
    visit(slide.Shapes, [], undefined, []);
    return nodes;
  }
  function geometryShape(node, slideIndex) {
    const native = node.native;
    const shape = { id: node.id, name: String(native.Name), type: nativeShapeKind(node.nativeType), nativeType: node.nativeType,
      left: Number(native.Left), top: Number(native.Top), width: Number(native.Width), height: Number(native.Height),
      rotation: Number(native.Rotation), ref: { slideIndex, shapePath: node.path } };
    if (node.parentId !== undefined) shape.parentId = node.parentId;
    return shape;
  }
  // 选中对象的几何树：包含组内选中对象所在的祖先组与其子树，保持
  // flatten/selected/bounds/alignment 现有的数学语义。ids 为空集合时只看选择，
  // 传 null 时返回整页几何。组内子形状返回的是原生坐标（可能位于组坐标系），
  // 只用于结构、选择与基线，不能当页面坐标做数学：选择范围遇到组内选中会回落到富上下文。
  function geometryShapes(nodes, slideIndex, ids) {
    const wanted = ids === null ? undefined : new Set((ids || []).map(String));
    if (wanted && !wanted.size) return [];
    const included = new Set();
    if (wanted) {
      for (const node of nodes) {
        if (!wanted.has(node.id)) continue;
        included.add(node.id);
        for (const ancestor of node.ancestors) included.add(ancestor);
      }
      for (const node of nodes) if (node.ancestors.some(id => included.has(id))) included.add(node.id);
    }
    const byParent = new Map();
    for (const node of nodes) {
      if (wanted && !included.has(node.id)) continue;
      if (!byParent.has(node.parentId)) byParent.set(node.parentId, []);
      byParent.get(node.parentId).push(node);
    }
    const build = parentId => (byParent.get(parentId) || []).map(node => {
      const shape = geometryShape(node, slideIndex);
      const children = build(node.id);
      if (children.length) shape.children = children;
      return shape;
    });
    return build(undefined);
  }
  // 轻量几何上下文：页面尺寸、选择顺序与几何树，不导出、不解析、不读字体。
  // 组内选中（HasChildShapeRange）子形状的原生坐标取自组坐标系，与 mergeParsedShape
  // 覆盖后的富模型（解析子坐标系）不一致，选择范围因此回落到富上下文，绝不把
  // 组坐标当作页面坐标参与对齐/排版数学。整页几何仅用于结构、选择与基线签名。
  function geometryContext(options = {}, extraIds = []) {
    const active = activeData();
    const selection = liveSelection(active.window);
    const selectionIds = selectionOrder(active, selection);
    const { width, height } = pageSize(active.presentation);
    const scope = options.scope === 'slide' ? 'slide' : 'selection';
    let nodes;
    if (scope === 'selection') {
      const state = selectionGroupState(active.slide, selection);
      if (state.group) return undefined;
      nodes = state.nodes;
    }
    const ids = scope === 'slide' ? null : [...new Set([...selectionIds, ...extraIds.map(String)])];
    return { presentationId: presentationKey(active.presentation), title: presentationTitle(active.presentation),
      slideId: String(optional(() => active.slide.SlideID) || active.slideIndex + 1), slideIndex: active.slideIndex,
      width, height, selectionIds, geometry: true,
      shapes: ids === null || ids.length ? geometryShapes(nodes || geometryTree(active.slide), active.slideIndex, ids) : [],
      selectedSlideIndices: selection.type === PP_SELECTION_SLIDES ? (selection.slideIndices || []) : [] };
  }
  // 组内选中的对齐沿用原有富坐标路径：以解析子坐标系坐标计算、原生写回，
  // 与 mergeParsedShape 的选择完全一致；只导出一份快照，不再二次构建上下文。
  async function alignRich(active, order, edge) {
    const snapshot = await buildContext({ scope: 'selection' });
    const byId = new Map(flattenShapes(snapshot.shapes).map(shape => [String(shape.id), shape]));
    const shapes = order.map(id => byId.get(id)).filter(Boolean);
    if (!shapes.length) throw new Error('所选对象已改变，请重新选择。');
    const commands = alignment(snapshot, shapes, edge);
    mutationDepth++;
    let changed = false;
    try {
      for (const command of commands) {
        await applyUpdate(active.slide, command, new Map());
        changed = true;
      }
    } finally {
      mutationDepth--;
      if (changed) { await rebaselineFingerprint(); schedule('apply'); }
    }
    return { title: presentationTitle(active.presentation), slideIndex: active.slideIndex,
      width: snapshot.width, height: snapshot.height, selectionIds: order, aligned: commands.length };
  }
  async function snapshotPptx(presentation) {
    return suppressed(() => withTemporary('pptx', async path => {
      presentation.SaveCopyAs(path, SAVE_AS_OPENXML_PRESENTATION, FALSE);
      return (await readFile(path)).data;
    }));
  }
  async function parsedPresentation(pptx, indices) {
    return request('pptx/parse', { pptx, ...(indices ? { indices } : {}) });
  }
  async function imagesInternal(ids, slideIndex, pptx, { original = true } = {}) {
    const wanted = ids.map(String);
    if (!wanted.length) return [];
    const result = await request('pptx/images', { pptx, slideIndex, ids: wanted, original });
    const list = Array.isArray(result) ? result : Array.isArray(result?.images) ? result.images : [];
    const byId = new Map(list.map(image => [String(image.id), image]));
    const output = [];
    for (const id of wanted) {
      const image = byId.get(id);
      if (!image) throw new Error(`无法读取对象 ${id} 的原始图片字节。请确认该对象是内嵌图片而非链接图片。`);
      let width = Number(image.width), height = Number(image.height);
      if (!(width > 0 && height > 0)) { const size = intrinsicSize(image.data); if (size) { width = size.width; height = size.height; } }
      if (!(width > 0 && height > 0) && /^image\//i.test(image.mime || '')) { const size = await decodedSize(image.data, image.mime); if (size) { width = size.width; height = size.height; } }
      if (!(width > 0 && height > 0) && (image.media || isVideoMime(image.mime))) {
        const file = await materialize({ data: image.data, mime: image.mime || 'application/octet-stream', name: 'media' });
        try {
          const info = await nativeExecute('mediaInfo', { path: file.path });
          width = Number(info?.width ?? info?.videoWidth); height = Number(info?.height ?? info?.videoHeight);
        } finally { temporaryAssets.add(file.path); }
      }
      output.push({ ...image, id: String(image.id), width, height, digest: await digest(image.data) });
    }
    return output;
  }
  async function attachImages(context, pptx) {
    const groups = new Map();
    if (context.slides) {
      for (const slide of context.slides) for (const shape of flattenShapes(slide.shapes)) {
        if (!['picture', 'media'].includes(shape.type)) continue;
        if (!groups.has(slide.index)) groups.set(slide.index, []);
        groups.get(slide.index).push(shape.id);
      }
    } else {
      groups.set(context.slideIndex, flattenShapes(context.shapes).filter(shape => ['picture', 'media'].includes(shape.type)).map(shape => shape.id));
    }
    const images = new Map();
    for (const [slideIndex, ids] of groups) {
      if (!ids.length) continue;
      for (const image of await imagesInternal(ids, slideIndex, pptx, { original: true })) images.set(image.id, image);
    }
    const apply = shapes => shapes.map(shape => {
      const image = images.get(String(shape.id));
      if (image) shape.image = { ...(shape.image || {}), data: image.data, mime: image.mime, source: image.source, width: image.width, height: image.height, digest: image.digest, crop: image.crop };
      if (shape.children) shape.children = apply(shape.children);
      return shape;
    });
    if (context.slides) context.slides.forEach(slide => { slide.shapes = apply(slide.shapes); });
    else context.shapes = apply(context.shapes);
    return context;
  }
  async function buildContext(options = {}) {
    const scope = options.scope || 'selection';
    const active = activeData();
    const selection = selectionState(active.window);
    const key = presentationKey(active.presentation) + '\u0000' + active.slideIndex;
    const selectionIds = orderedSelection(key, selection.ids);
    const snapshot = await snapshotPptx(active.presentation);
    const parsed = await parsedPresentation(snapshot, scope === 'presentation' ? undefined : [active.slideIndex]);
    const parsedSlides = new Map((parsed.slides || []).map(slide => [Number(slide.index), slide]));
    const width = Number(optional(() => active.presentation.PageSetup.SlideWidth)) || Number(parsed.width);
    const height = Number(optional(() => active.presentation.PageSetup.SlideHeight)) || Number(parsed.height);
    if (!(width > 0 && height > 0)) throw new Error('无法读取 WPS 幻灯片尺寸。');
    const mergeSlide = slide => {
      const index = Number(slide.SlideIndex) - 1;
      const model = parsedSlides.get(index);
      const nativeShapes = shapeIndex(slide, index);
      return { id: String(optional(() => slide.SlideID) || index + 1), index,
        name: String(optional(() => slide.Name) || `幻灯片${index + 1}`),
        shapes: mergeShapes(nativeShapes, model?.shapes || []) };
    };
    const result = {
      presentationId: presentationKey(active.presentation),
      title: String(optional(() => active.presentation.Name) || 'WPS 演示'),
      slideId: String(optional(() => active.slide.SlideID) || active.slideIndex + 1),
      slideIndex: active.slideIndex, width, height, selectionIds,
      shapes: mergeSlide(active.slide).shapes,
      selectedSlideIndices: selection.type === PP_SELECTION_SLIDES ? (selection.slideIndices || []) : [],
    };
    if (selection.textSelection) result.textSelection = selection.textSelection;
    if (scope === 'presentation') {
      result.slides = [];
      for (const slide of items(active.presentation.Slides)) result.slides.push(mergeSlide(slide));
    }
    if (options.images === true) await attachImages(result, snapshot);
    return result;
  }
  async function fragmentFromSnapshot(ids, slideIndex, pptx) {
    const wanted = ids.map(String);
    if (!wanted.length) throw new Error('请先选择需要保存的对象。');
    const fragment = await request('pptx/fragment', { pptx, slideIndex, ids: wanted });
    if (!fragment?.pptx) throw new Error('未能生成可编辑的 PPTX 片段。');
    return { ...fragment, shapeIds: (fragment.shapeIds || wanted).map(String), bounds: fragment.bounds };
  }
  async function fragmentInternal(ids) {
    const active = activeData();
    const pptx = await snapshotPptx(active.presentation);
    return fragmentFromSnapshot(ids, active.slideIndex, pptx);
  }
  async function mediaInfo(path) {
    const info = await nativeExecute('mediaInfo', { path });
    const width = Number(info?.width ?? info?.videoWidth), height = Number(info?.height ?? info?.videoHeight);
    if (!(width > 0 && height > 0)) throw new Error('无法读取视频画面的原始尺寸。');
    return { width, height, duration: Number(info?.duration ?? info?.length) || undefined };
  }
  function pictureAsset(value) { return materialize({ data: value.data, mime: value.mime || 'image/png', name: 'picture-fill' }).then(file => file.path); }
  async function applyStyle(shape, style) {
    if (!style) return;
    if (style.fragment?.pptx) {
      const file = await materialize({ data: style.fragment.pptx, mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', name: 'style.pptx' });
      const geometry = { left: shape.Left, top: shape.Top, width: shape.Width, height: shape.Height, rotation: shape.Rotation, locked: shape.LockAspectRatio };
      const metadata = { name: shape.Name, alternativeText: shape.AlternativeText, tags: readTags(shape) };
      const originalText = { text: String(optional(() => shape.TextFrame.TextRange.Text) || '') };
      let source;
      try {
        source = app.Presentations.Open(file.path, TRUE, FALSE, FALSE);
        if (!source) throw new Error('无法打开格式片段。');
        const sourceShape = source.Slides.Item(1).Shapes.Item(1);
        sourceShape.PickUp();
        shape.Apply();
        shape.LockAspectRatio = FALSE;
        shape.Left = geometry.left; shape.Top = geometry.top;
        shape.Width = geometry.width; shape.Height = geometry.height; shape.Rotation = geometry.rotation;
        shape.LockAspectRatio = geometry.locked === TRUE ? TRUE : FALSE;
        shape.Name = metadata.name; shape.AlternativeText = metadata.alternativeText;
        writeTags(shape, metadata.tags);
        if (optional(() => shape.HasTextFrame) && !sameText(shape.TextFrame.TextRange.Text, originalText.text)) {
          shape.TextFrame.TextRange.Text = canonical(originalText.text);
        }
      } finally {
        if (source) { try { source.Saved = TRUE; source.Close(); } catch { try { source.Close(); } catch { /* the style operation's error is reported first */ } } }
      }
      return;
    }
    const value = {};
    for (const key of ['fill', 'line', 'font', 'textFrame']) if (style[key] !== undefined) value[key] = style[key];
    if (Object.keys(value).length) await properties(shape, value, pictureAsset);
  }
  async function applyProperties(shape, value) {
    const input = { ...(value || {}) };
    const style = input.style;
    delete input.style;
    if (input.crop !== undefined && !isPicture(shape)) throw new Error('裁剪只能应用于图片对象。');
    if (style) { await applyStyle(shape, style); delete input.style; }
    if (input.tags !== undefined) input.tags = { ...readTags(shape), ...input.tags };
    if (!Object.keys(input).length) return;
    await properties(shape, input, pictureAsset);
  }
  function shapeRange(slide, entries) {
    if (!entries.length) throw new Error('请先选择对象。');
    if (entries.some(entry => !entry)) throw new Error('所选对象已改变，请重新选择。');
    const parent = entries[0].parent;
    if (entries.some(entry => entry.parent !== parent)) throw new Error('所选对象不在同一个组合层级，无法一起处理。');
    const collection = parent ? parent.GroupItems : slide.Shapes;
    const indices = entries.map(entry => {
      const list = items(collection);
      const index = list.findIndex(shape => String(shape.Id) === String(entry.shape.Id));
      if (index < 0) throw new Error('所选对象已改变，请重新选择。');
      return index + 1;
    });
    return collection.Range(indices);
  }
  function resolveEntry(slide, id, remap) { return findNativeEntry(slide, remap.get(String(id)) || id); }
  function resolveShape(slide, id, remap) {
    const entry = resolveEntry(slide, id, remap);
    if (!entry) throw new Error(`WPS 当前页面中找不到对象 ${id}，请重新选择。`);
    return entry.shape;
  }
  async function addPicture(slide, value, path) {
    const args = [path, FALSE, TRUE, Number(value.left || 0), Number(value.top || 0)];
    if (value.width !== undefined) args.push(Number(value.width));
    if (value.height !== undefined) args.push(Number(value.height));
    return slide.Shapes.AddPicture(...args);
  }
  async function insertImage(slide, command) {
    const mime = command.mime || 'image/png';
    const file = command.data ? await materialize({ data: command.data, mime, name: command.name || 'image' }) : (command.path ? { path: command.path } : undefined);
    if (!file?.path) throw new Error('插入图片需要 data 或 path。');
    let shape;
    if (isVideoMime(mime) || command.op === 'insertMedia') {
      let width = Number(command.width), height = Number(command.height);
      if (!(width > 0) || !(height > 0)) {
        const info = await mediaInfo(file.path);
        if (!(width > 0) && !(height > 0)) {
          const scale = Math.min(1, 0.9 * (Number(optional(() => slide.Parent.PageSetup.SlideWidth)) || 960) / (info.width * 0.75),
            0.9 * (Number(optional(() => slide.Parent.PageSetup.SlideHeight)) || 540) / (info.height * 0.75));
          width = info.width * 0.75 * scale; height = info.height * 0.75 * scale;
        } else if (!(width > 0)) width = height * info.width / info.height;
        else if (!(height > 0)) height = width * info.height / info.width;
      }
      const left = Number(command.left ?? 0), top = Number(command.top ?? 0);
      shape = slide.Shapes.AddMediaObject2(file.path, FALSE, TRUE, left, top, width, height);
    } else {
      shape = await addPicture(slide, command, file.path);
    }
    const metadata = {};
    for (const key of ['rotation', 'name', 'alternativeText', 'tags', 'fill', 'line', 'crop', 'textFrame', 'font', 'text', 'runs', 'paragraphs']) {
      if (command[key] !== undefined) metadata[key] = command[key];
    }
    if (metadata.name === undefined && command.name) metadata.name = command.name;
    if (metadata.tags !== undefined) metadata.tags = { ...metadata.tags };
    if (Object.keys(metadata).length) await applyProperties(shape, metadata);
    return shape;
  }
  async function insertText(slide, command) {
    const shape = slide.Shapes.AddTextbox(TEXT_ORIENTATION_HORIZONTAL,
      Number(command.left || 0), Number(command.top || 0), Number(command.width || 1), Number(command.height || 1));
    const metadata = {};
    for (const key of ['rotation', 'name', 'alternativeText', 'tags', 'fill', 'line', 'textFrame', 'font', 'text', 'runs', 'paragraphs']) {
      if (command[key] !== undefined) metadata[key] = command[key];
    }
    if (Object.keys(metadata).length) await applyProperties(shape, metadata);
    return shape;
  }
  async function insertShape(slide, command) {
    const type = String(command.shapeType || 'rectangle');
    const left = Number(command.left || 0), top = Number(command.top || 0);
    const width = Number(command.width || 0), height = Number(command.height || 0);
    let shape;
    if (type === 'line' || type === 'straightConnector' || type === 'connector') shape = slide.Shapes.AddLine(left, top, left + width, top + height);
    else {
      const code = shapeCode(type) ?? (/^\d+$/.test(type) ? Number(type) : undefined);
      if (code === undefined) throw new Error(`WPS 不支持形状类型：${type}`);
      shape = slide.Shapes.AddShape(code, left, top, width, height);
    }
    const metadata = {};
    for (const key of ['rotation', 'name', 'alternativeText', 'tags', 'fill', 'line', 'font', 'text', 'runs', 'paragraphs', 'textFrame']) {
      if (command[key] !== undefined) metadata[key] = command[key];
    }
    if (Object.keys(metadata).length) await applyProperties(shape, metadata);
    return shape;
  }
  async function replaceObject(slide, command, remap, operation = 'replacePicture') {
    const entry = resolveEntry(slide, command.id, remap);
    if (!entry) throw new Error(`WPS 当前页面中找不到对象 ${command.id}，请重新选择。`);
    const targetId = String(entry.shape.Id);
    if (operation === 'replacePicture' && !isPicture(entry.shape)) throw new Error('替换图片只适用于图片对象。');
    let root = entry;
    while (root.parent) root = findNativeEntry(slide, root.parent.Id);
    const rootShape = root.shape;
    function capture(shape) {
      return { id: String(shape.Id), name: String(shape.Name), alternativeText: String(shape.AlternativeText || ''),
        tags: readTags(shape), children: isGroup(shape) ? items(shape.GroupItems).map(capture) : [] };
    }
    const previous = capture(rootShape);
    const geometry = { left: Number(rootShape.Left), top: Number(rootShape.Top), width: Number(rootShape.Width),
      height: Number(rootShape.Height), rotation: Number(rootShape.Rotation), z: Number(rootShape.ZOrderPosition) - 1 };
    const presentation = activeData().presentation;
    const original = await withTemporary('pptx', async filename => {
      presentation.SaveCopyAs(filename, SAVE_AS_OPENXML_PRESENTATION);
      return readFile(filename);
    });
    const fragment = await request('pptx/' + operation, { ...command, pptx: original.data,
      slideIndex: Number(slide.SlideIndex) - 1, id: String(entry.shape.Id), rootId: previous.id,
      updates: operation === 'formatText' ? [{ id: String(entry.shape.Id), properties: command.properties }] : undefined });
    const file = await materialize({ data: fragment.pptx,
      mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', name: 'replacement.pptx' });
    const clipboard = await nativeExecute('snapshotClipboard');
    let source, replacement, committed = false;
    const pairs = [];
    try {
      source = await openHidden(file.path);
      source.Slides.Item(1).Shapes.Range().Copy();
      const pasted = slide.Shapes.Paste();
      if (!pasted || Number(pasted.Count) !== 1) throw new Error('替换后未能取得完整的原生对象。');
      replacement = pasted.Item(1);
      function pair(before, after) {
        const children = isGroup(after) ? items(after.GroupItems) : [];
        if (children.length !== before.children.length) throw new Error('替换后的组合成员发生变化，已保留原对象。');
        pairs.push([before, after]);
        before.children.forEach((child, index) => pair(child, children[index]));
      }
      pair(previous, replacement);
      const locked = replacement.LockAspectRatio;
      replacement.LockAspectRatio = FALSE;
      replacement.Left = geometry.left; replacement.Top = geometry.top;
      replacement.Width = geometry.width; replacement.Height = geometry.height;
      replacement.Rotation = geometry.rotation; replacement.LockAspectRatio = locked;
      for (const [before, after] of pairs) {
        after.Name = before.name; after.AlternativeText = before.alternativeText; writeTags(after, before.tags);
      }
      zOrder(replacement, 'index', geometry.z);
      rootShape.Delete(); committed = true;
      for (const [before, after] of pairs) remap.set(before.id, String(after.Id));
      remap.set(String(command.id), remap.get(targetId));
      return replacement;
    } finally {
      try { if (!committed && replacement) replacement.Delete(); }
      finally {
        try { closeHidden(source); }
        finally { await nativeExecute('restoreClipboard', { items: clipboard.items }); }
      }
    }
  }
  async function applyUpdate(slide, command, remap) {
    const richKeys = new Set(['font', 'fontScope', 'runs', 'paragraphs', 'textRange']);
    const ordinary = {}, rich = {};
    for (const [key, value] of Object.entries(command.properties || {})) {
      (richKeys.has(key) ? rich : ordinary)[key] = value;
    }
    if (rich.paragraphs) ordinary.paragraphs = rich.paragraphs.map(({ runs, text, font, ...layout }) => layout);
    if (rich.font?.emboss !== undefined && !rich.font.raw) {
      ordinary.font = { emboss: rich.font.emboss };
      rich.font = { ...rich.font }; delete rich.font.emboss;
    }
    if (rich.font?.reflection && !rich.font.raw && !rich.font.reflection.raw) {
      ordinary.font = { ...ordinary.font, reflection: rich.font.reflection };
      rich.font = { ...rich.font }; delete rich.font.reflection;
    }
    if (rich.textRange?.font?.emboss !== undefined && !rich.textRange.font.raw) {
      ordinary.textRange = { ...rich.textRange, font: { emboss: rich.textRange.font.emboss } };
      delete ordinary.textRange.text;
      rich.textRange = { ...rich.textRange, font: { ...rich.textRange.font } };
      delete rich.textRange.font.emboss;
    }
    if (Object.keys(ordinary).length) await applyProperties(resolveShape(slide, command.id, remap), ordinary);
    if (Object.keys(rich).some(key => key !== 'fontScope')) {
      await replaceObject(slide, { ...command, properties: rich }, remap, 'formatText');
    }
  }
  async function duplicateShape(slide, command, remap) {
    const source = resolveShape(slide, command.id, remap);
    const entry = resolveEntry(slide, command.id, remap);
    const before = new Set(items(slide.Shapes).map(shape => String(shape.Id)));
    let result;
    if (typeof source.Duplicate === 'function') result = source.Duplicate();
    else {
      if (entry.parent) throw new Error('WPS 无法复制组合内单个对象，请先取消组合。');
      const clipboard = await nativeExecute('snapshotClipboard');
      try { source.Copy(); result = slide.Shapes.Paste(); }
      finally { try { await nativeExecute('restoreClipboard', { items: clipboard.items }); } catch { /* clipboard restoration must not mask the copy result */ } }
    }
    const ranges = result?.Count !== undefined ? items(result) : [result];
    let copy = ranges[0];
    if (!copy) {
      const added = items(slide.Shapes).filter(shape => !before.has(String(shape.Id)));
      if (!added.length) throw new Error('复制对象后未找到新对象。');
      copy = added.at(-1);
    }
    if (command.properties) await applyProperties(copy, command.properties);
    if (command.name !== undefined) copy.Name = String(command.name);
    if (command.tags !== undefined) writeTags(copy, { ...readTags(copy), ...command.tags });
    return copy;
  }
  async function openHidden(path) {
    const presentation = app.Presentations.Open(path, TRUE, FALSE, FALSE);
    if (!presentation) throw new Error('WPS 无法打开临时演示文稿。');
    return presentation;
  }
  function closeHidden(presentation) {
    if (!presentation) return;
    try { presentation.Saved = TRUE; } catch { /* read-only presentations may reject Saved */ }
    try { presentation.Close(); } catch (error) { throw new Error(`无法关闭临时演示文稿：${error.message}`); }
  }
  function exportPdf(presentation, path) {
    if (typeof presentation.ExportAsFixedFormat === 'function') presentation.ExportAsFixedFormat(path, FIXED_FORMAT_PDF);
    else presentation.SaveCopyAs(path, 32);
  }
  async function withRenderedPdf(pptxPath, operation) {
    return suppressed(() => withTemporary('pdf', async pdfPath => {
      let source, result, failure;
      try {
        source = await openHidden(pptxPath);
        exportPdf(source, pdfPath);
        result = await operation(pdfPath);
      } catch (error) { failure = error; }
      try { closeHidden(source); }
      catch (error) { failure = failure ? new AggregateError([failure, error], '临时演示文稿未能关闭。') : error; }
      if (failure) throw failure;
      return result;
    }));
  }
  async function renderSelectedPdf(indices, presentation, operation) {
    const snapshot = await snapshotPptx(presentation);
    const selected = indices.length === Number(presentation.Slides.Count)
      ? snapshot : (await request('pptx/selectSlides', { pptx: snapshot, indices })).pptx;
    const file = await materialize({ data: selected, mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', name: 'export.pptx' });
    return withRenderedPdf(file.path, operation);
  }
  async function exportSlideRaster(slide, index, format, dpi, slideWidth, slideHeight, name, presentation) {
    const width = Math.max(1, Math.round(slideWidth * dpi / 72));
    const height = Math.max(1, Math.round(slideHeight * dpi / 72));
    let failure;
    const extension = format === 'jpeg' ? 'jpg' : format;
    try {
      return await withTemporary(extension, async path => {
        slide.Export(path, EXPORT_FILTERS[format], width, height);
        const file = await readFile(path);
        return { name: `${name}.${extension}`, mime: IMAGE_FORMATS[format], data: file.data, width, height };
      });
    } catch (error) { failure = error; }
    try {
      return await renderSelectedPdf([index], presentation, async pdfPath => {
        const bounds = { left: 0, top: 0, width: slideWidth, height: slideHeight };
        const pages = await nativeExecute('rasterPDF', { path: pdfPath, indices: [0], dpi: Math.min(dpi, MAX_RASTER_DPI), format: format === 'jpeg' ? 'jpg' : format, bounds });
        const page = pages?.[0];
        if (!page?.data) throw new Error('WPS 未能渲染幻灯片。');
        return { name: `${name}.${format === 'jpeg' ? 'jpg' : format}`, mime: page.mime || IMAGE_FORMATS[format], data: page.data, width: page.width, height: page.height };
      });
    } catch (fallbackError) { throw new AggregateError([failure, fallbackError], `WPS 导出幻灯片 ${index + 1} 失败。`); }
  }
  async function deliver(files, folder) {
    if (!folder) return files;
    const output = [];
    for (const file of files) {
      const saved = await request('files/write', { folder, name: file.name, data: file.data, mime: file.mime });
      const path = typeof saved === 'string' ? saved : saved?.path;
      if (!path) throw new Error(`无法写入导出文件：${file.name}`);
      output.push({ ...file, path });
    }
    return output;
  }
  async function finishExport(files, options) {
    const delivered = await deliver(files, options.folder);
    if (options.copyToClipboard && delivered.length) {
      if (typeof services?.writeClipboard !== 'function') throw new Error('WPS 宿主缺少剪贴板写入服务。');
      const first = delivered[0];
      if (!/^(image|application\/pdf)/i.test(first.mime || '')) throw new Error('该导出格式不能直接写入剪贴板。');
      await services.writeClipboard({ data: first.data, mime: first.mime });
    }
    return delivered;
  }
  async function exportSelection(ids, format, dpi, options) {
    const fragment = await fragmentInternal(ids);
    const stem = safeStem(options.filename, 'selection');
    if (format === 'pptx') return finishExport([{ name: `${stem}.pptx`, mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', data: fragment.pptx }], options);
    const file = await materialize({ data: fragment.pptx, mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', name: 'selection.pptx' });
    if (format === 'pdf') {
      return withRenderedPdf(file.path, async pdfPath => {
        let data;
        if (fragment.bounds) {
          const cropped = await nativeExecute('slicePDF', { path: pdfPath, indices: [0], bounds: fragment.bounds });
          data = cropped.data;
        } else data = (await readFile(pdfPath)).data;
        return finishExport([{ name: `${stem}.pdf`, mime: 'application/pdf', data }], options);
      });
    }
    if (!IMAGE_FORMATS[format]) throw new Error(`WPS 不支持导出格式：${format}`);
    return withRenderedPdf(file.path, async pdfPath => {
      const bounds = fragment.bounds || { left: 0, top: 0, width: 0, height: 0 };
      const pages = await nativeExecute('rasterPDF', { path: pdfPath, indices: [0], dpi: Math.min(dpi, MAX_RASTER_DPI), format: format === 'jpeg' ? 'jpg' : format, bounds });
      const page = pages?.[0];
      if (!page?.data) throw new Error('WPS 未能渲染选中对象。');
      const mime = page.mime || IMAGE_FORMATS[format];
      return finishExport([{ name: `${stem}.${format === 'jpeg' ? 'jpg' : format}`, mime, data: page.data, width: page.width, height: page.height }], options);
    });
  }
  async function exportSlides(options) {
    const active = activeData();
    const format = String(options.format || 'png').toLowerCase();
    const dpi = Number(options.dpi ?? 300);
    if (!Number.isFinite(dpi) || dpi < 1 || dpi > MAX_RASTER_DPI) throw new Error('DPI 必须在 1 到 2400 之间。');
    const count = Number(active.presentation.Slides.Count);
    const indices = options.indices?.length ? [...new Set(options.indices.map(Number))] : Array.from({ length: count }, (_, index) => index);
    if (!indices.length || indices.some(index => !Number.isInteger(index) || index < 0 || index >= count)) throw new Error('导出页码超出演示文稿范围。');
    const stem = safeStem(options.filename, 'slides');
    if (format === 'pdf') {
      if (options.separatePdf === true) {
        // 一次导出整份演示文稿 PDF，再交给原生侧一次性按页切分：
        // 不做逐页 PPTX 快照，也不逐页导出，最后只交付一次。
        return suppressed(() => withTemporary('pdf', async path => {
          exportPdf(active.presentation, path);
          const pages = await nativeExecute('slicePDF', { path, indices, separate: true });
          if (!Array.isArray(pages) || pages.length !== indices.length || pages.some(page => !page?.data)) {
            throw new Error('WPS 未能按页拆分导出的 PDF。');
          }
          const files = pages.map((page, offset) => ({ name: `${stem}_页面${indices[offset] + 1}.pdf`,
            mime: page.mime || 'application/pdf', data: page.data }));
          return finishExport(files, options);
        }));
      }
      if (indices.length === count) {
        const data = await suppressed(() => withTemporary('pdf', async path => { exportPdf(active.presentation, path); return (await readFile(path)).data; }));
        return finishExport([{ name: `${stem}.pdf`, mime: 'application/pdf', data }], options);
      }
      const snapshot = await snapshotPptx(active.presentation);
      const selected = (await request('pptx/selectSlides', { pptx: snapshot, indices })).pptx;
      const file = await materialize({ data: selected, mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', name: 'export.pptx' });
      return withRenderedPdf(file.path, async path => finishExport([{ name: `${stem}.pdf`, mime: 'application/pdf', data: (await readFile(path)).data }], options));
    }
    if (format === 'pptx') {
      let data;
      if (indices.length === count) data = await snapshotPptx(active.presentation);
      else data = (await request('pptx/selectSlides', { pptx: await snapshotPptx(active.presentation), indices })).pptx;
      return finishExport([{ name: `${stem}.pptx`, mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', data }], options);
    }
    if (!IMAGE_FORMATS[format]) throw new Error(`WPS 不支持导出格式：${format}`);
    const slideWidth = Number(optional(() => active.presentation.PageSetup.SlideWidth)) || 0;
    const slideHeight = Number(optional(() => active.presentation.PageSetup.SlideHeight)) || 0;
    if (!(slideWidth > 0 && slideHeight > 0)) throw new Error('无法读取幻灯片尺寸。');
    const files = [];
    for (const index of indices) {
      const name = indices.length === 1 ? stem : `${stem}-${index + 1}`;
      files.push(await exportSlideRaster(slideAt(active.presentation, index), index, format, dpi, slideWidth, slideHeight, name, active.presentation));
    }
    return finishExport(files, options);
  }
  async function exportOriginalImages(images, options) {
    const prefix = options.filename ? safeStem(options.filename, 'image') : undefined;
    const used = new Set();
    const files = images.map((image, index) => {
      const extension = extensionFor(image.mime, 'bin');
      let base = prefix ? `${prefix}${images.length > 1 ? '-' + (index + 1) : ''}` : safeStem(image.name || `image-${index + 1}`, `image-${index + 1}`);
      base = base.replace(new RegExp(`\\.${extension.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}$`, 'i'), '');
      let name = `${base}.${extension}`, suffix = 2;
      while (used.has(name.toLocaleLowerCase())) name = `${base} (${suffix++}).${extension}`;
      used.add(name.toLocaleLowerCase());
      return { name, mime: image.mime || 'application/octet-stream', data: image.data,
        width: image.width, height: image.height, source: image.source, digest: image.digest };
    });
    return finishExport(files, options);
  }
  async function exportInternal(options = {}) {
    const kind = options.kind || 'slides';
    if (kind === 'original-images') {
      const active = activeData();
      const ids = options.ids?.length ? options.ids.map(String) : selectionOrder(active, liveSelection(active.window));
      if (!ids.length) throw new Error('请先选择图片。');
      const pptx = await snapshotPptx(active.presentation);
      return exportOriginalImages(await imagesInternal(ids, active.slideIndex, pptx, { original: true }), options);
    }
    if (kind === 'selection') {
      const active = activeData();
      const ids = options.ids?.length ? options.ids.map(String) : liveSelection(active.window).ids;
      if (!ids.length) throw new Error('请先选择对象。');
      const format = String(options.format || 'png').toLowerCase();
      const dpi = Number(options.dpi ?? 300);
      if (!Number.isFinite(dpi) || dpi < 1 || dpi > MAX_RASTER_DPI) throw new Error('DPI 必须在 1 到 2400 之间。');
      return exportSelection(ids, format, dpi, options);
    }
    if (kind !== 'slides') throw new Error(`未知的导出类型：${kind}`);
    return exportSlides(options);
  }
  // 轻量指纹：Id、类型、几何、层级、名称与文字，每轮轮询都读；
  // 富格式（填充、轮廓、裁剪、标记、逐段字体）交给 fingerprint 深度扫描。
  function cheapShapeSignature(native) {
    const type = Number(native.Type);
    const text = boolean(optional(() => native.HasTextFrame)) ? String(optional(() => native.TextFrame.TextRange.Text) || '') : '';
    const signature = [String(native.Id), type, round(native.Left), round(native.Top), round(native.Width), round(native.Height),
      round(native.Rotation), String(native.Name), Number(native.ZOrderPosition) - 1, text];
    if (type === 6) signature.push(items(native.GroupItems).map(child => cheapShapeSignature(child)));
    return signature;
  }
  function cheapFingerprint(active) {
    const selection = liveSelection(active.window);
    return JSON.stringify([presentationKey(active.presentation), active.slideIndex, Number(active.presentation.Slides.Count),
      selection.type, selection.ids.join(','), items(active.slide.Shapes).map(native => cheapShapeSignature(native))]);
  }
  async function fingerprint(active) {
    const selection = liveSelection(active.window);
    const shapes = shapeIndex(active.slide, active.slideIndex);
    const shapeSignature = shape => [String(shape.id), round(shape.left), round(shape.top), round(shape.width), round(shape.height),
      round(shape.rotation), shape.name, shape.alternativeText, shape.zOrder, shape.text, JSON.stringify(shape.tags),
      JSON.stringify(shape.crop), JSON.stringify(shape.font), JSON.stringify(shape.fill), JSON.stringify(shape.line),
      shape.children ? shape.children.map(shapeSignature) : []];
    const scope = presentationKey(active.presentation) + ':' + active.slideIndex;
    const linked = flattenShapes(shapes).some(shape => Object.keys(shape.tags || {}).some(key => key.toUpperCase() === 'SLIDESCI_ZOOM_RECORDS'));
    if (scope !== pictureProbeScope || linked && Date.now() - pictureProbeAt >= 2500) {
      pictureProbeScope = scope; pictureProbeAt = Date.now(); pictureRevisions = [];
      if (linked) {
        const parsed = await parsedPresentation(await snapshotPptx(active.presentation), [active.slideIndex]);
        pictureRevisions = flattenShapes(parsed.slides[0]?.shapes || []).filter(shape => shape.image).map(shape => [shape.id, shape.image.digest]);
      }
    }
    return JSON.stringify([presentationKey(active.presentation), active.slideIndex, Number(active.presentation.Slides.Count),
      selection.ids.join(','), selection.type, shapes.map(shapeSignature), pictureRevisions]);
  }
  function dispatch(reason) {
    const event = { host: 'wps', reason };
    for (const listener of [...listeners]) {
      try {
        const result = listener(event);
        if (result?.catch) result.catch(error => console.error('SlideSCI WPS 变更回调失败：' + error.message));
      } catch (error) { console.error('SlideSCI WPS 变更回调失败：' + error.message); }
    }
  }
  function schedule(reason) {
    if (!listeners.size || muted()) return;
    clearTimeout(dispatchTimer);
    dispatchTimer = setTimeout(() => {
      dispatchTimer = undefined;
      if (!listeners.size || muted()) return;
      dispatch(reason);
    }, 70);
  }
  function registerEvents() {
    const api = optional(() => app.ApiEvent);
    if (!api || typeof api.AddApiEventListener !== 'function') return;
    for (const name of ['WindowSelectionChange', 'SlideSelectionChanged', 'PresentationNewSlide', 'PresentationOpen', 'PresentationClose', 'PresentationSave', 'WindowActivate']) {
      const handler = () => schedule(name);
      try { api.AddApiEventListener(name, handler); eventHandlers.push([name, handler]); } catch { /* polling below remains the complete fallback */ }
    }
  }
  function unregisterEvents() {
    const api = optional(() => app.ApiEvent);
    for (const [name] of eventHandlers) { try { api?.RemoveApiEventListener?.(name); } catch { /* the web page is closing */ } }
    eventHandlers.length = 0;
  }
  function invalidateFingerprint() { lastCheap = undefined; lastDeep = undefined; deepPolls = 0; }
  async function rebaselineFingerprint() {
    try {
      const active = activeData();
      lastCheap = cheapFingerprint(active);
      lastDeep = await fingerprint(active);
      deepPolls = 0;
    } catch { invalidateFingerprint(); }
  }
  function poll() {
    if (!watching) return;
    pollTimer = setTimeout(async () => {
      if (!watching) return;
      try {
        if (!muted()) {
          const active = activeData();
          const cheap = cheapFingerprint(active);
          if (lastCheap === undefined) { lastCheap = cheap; lastDeep = await fingerprint(active); deepPolls = 0; }
          else if (cheap !== lastCheap) {
            lastCheap = cheap; deepPolls = 0;
            const deep = await fingerprint(active);
            if (deep !== lastDeep) { lastDeep = deep; schedule('poll'); }
          } else if (++deepPolls >= DEEP_POLL_INTERVAL) {
            deepPolls = 0;
            const deep = await fingerprint(active);
            if (deep !== lastDeep) { lastDeep = deep; schedule('poll'); }
          }
        }
      } catch { /* no active presentation or a transient WPS view change */ }
      if (watching) poll();
    }, POLL_INTERVAL);
  }
  function startWatching() {
    if (watching) return;
    watching = true;
    invalidateFingerprint();
    registerEvents();
    poll();
  }
  function stopWatching() {
    watching = false;
    clearTimeout(pollTimer); clearTimeout(dispatchTimer);
    pollTimer = dispatchTimer = undefined;
    unregisterEvents();
    invalidateFingerprint();
  }
  async function finishMutation(touched) {
    mutationDepth--;
    if (!touched) return;
    await rebaselineFingerprint();
    schedule('apply');
  }

  const host = {
    name: 'WPS 演示',
    kind: 'wps',

    async context(options = {}) {
      // 纯排版只读几何：不导出、不解析、不读字体。
      const geometry = options.geometry === true && (options.scope === undefined || options.scope === 'selection' || options.scope === 'slide');
      return tracked(() => suppressed(() => {
        if (!geometry) return buildContext(options);
        const snapshot = geometryContext(options);
        // 组内选中时几何上下文不可用（子形状为组坐标），回落到富上下文。
        return snapshot === undefined ? buildContext(options) : snapshot;
      }));
    },

    // 任务窗格与导出规划用的轻量摘要：不导出、不解析、不读形状富格式。
    // slides:false 跳过逐页 ID 枚举（UI 轮询只取标题/页码/选择数时用），此时
    // 仍返回 slides: [] 以保持返回结构一致。
    async summary(options = {}) {
      return tracked(() => {
        const active = activeData();
        const selection = liveSelection(active.window);
        const { width, height } = pageSize(active.presentation);
        let slides = [];
        if (options.slides !== false) {
          const collection = active.presentation.Slides;
          const count = Number(collection.Count);
          slides = Array.from({ length: count }, (_, index) => {
            const slide = collection.Item(index + 1);
            return { id: String(optional(() => slide.SlideID) || index + 1), index };
          });
        }
        return { presentationId: presentationKey(active.presentation), title: presentationTitle(active.presentation),
          slideId: String(optional(() => active.slide.SlideID) || active.slideIndex + 1), slideIndex: active.slideIndex,
          width, height, slideCount: Number(optional(() => active.presentation.Slides.Count) || 0),
          selectionIds: selectionOrder(active, selection), slides,
          selectedSlideIndices: selection.type === PP_SELECTION_SLIDES ? (selection.slideIndices || []) : [] };
      });
    },

    // 快速对齐：直接读取当前选区几何并写回 Left/Top，六种对齐共用
    // editing-geometry 的对齐数学。不导出、不解析、不重算富指纹。
    async align(edge) {
      return tracked(async () => {
        const active = activeData();
        const selection = liveSelection(active.window);
        const order = selectionOrder(active, selection);
        if (!order.length) throw new Error('请先选择至少 1 个形状。');
        // 组内选中的子形状原生坐标属于组坐标系，不能当页面坐标做数学。
        if (selectionGroupState(active.slide, selection).group) return alignRich(active, order, edge);
        // liveSelection 已经返回原生形状对象，按点选顺序重排即可，无需再全页查找。
        const natives = new Map(selection.shapes.map(shape => [String(shape.Id), shape]));
        const shapes = order.map(id => {
          const native = natives.get(id);
          if (!native) throw new Error(`WPS 当前页面中找不到对象 ${id}，请重新选择。`);
          return { id, name: String(optional(() => native.Name) || id),
            left: Number(native.Left), top: Number(native.Top),
            width: Number(native.Width), height: Number(native.Height) };
        });
        const { width, height } = pageSize(active.presentation);
        // 多选以点选顺序的基准形状为准，单选以幻灯片页面为基准。
        const commands = alignment({ width, height, selectionIds: order, shapes }, shapes, edge);
        mutationDepth++;
        let changed = false;
        try {
          for (const command of commands) {
            const shape = natives.get(String(command.id));
            if (!shape) throw new Error('所选对象已改变，请重新选择。');
            await properties(shape, command.properties);
            changed = true;
          }
        } finally {
          mutationDepth--;
          if (changed) { invalidateFingerprint(); schedule('apply'); }
        }
        return { title: presentationTitle(active.presentation), slideIndex: active.slideIndex, width, height,
          selectionIds: order, aligned: commands.length };
      });
    },

    async apply(commands, options = {}) {
      if (!Array.isArray(commands) || !commands.length) return this.context({ scope: 'slide' });
      return tracked(async () => {
        mutationDepth++;
        let touched = false;
        // 命中条件时才返回轻量几何模型，其余调用保持原有富上下文契约。
        let lightweight = false;
        let rich = false, geometryUpdate = false;
        const movedIds = new Set();
        try {
          const active = activeData();
          const presentation = active.presentation;
          const remap = new Map();
          for (const command of commands) {
            if (!command || typeof command.op !== 'string') throw new Error('无效的原生命令。');
            if (command.op === 'update' && geometryProperties(command.properties)) { geometryUpdate = true; movedIds.add(String(command.id)); }
            else if (command.op !== 'select') rich = true;
            // 轻量模型只描述当前页，跨页命令一律回落到完整上下文。
            if (command.slideIndex !== undefined) rich = true;
            const slide = command.slideIndex === undefined ? active.slide : slideAt(presentation, command.slideIndex);
            switch (command.op) {
              case 'update': {
                await applyUpdate(slide, command, remap);
                touched = true;
                break;
              }
              case 'insertText': {
                await insertText(slide, command);
                touched = true;
                break;
              }
              case 'insertImage':
              case 'insertMedia': {
                await insertImage(slide, command);
                touched = true;
                break;
              }
              case 'insertShape': {
                await insertShape(slide, command);
                touched = true;
                break;
              }
              case 'delete': {
                for (const id of command.ids || []) {
                  const entry = resolveEntry(slide, id, remap);
                  if (!entry) throw new Error(`WPS 当前页面中找不到对象 ${id}，请重新选择。`);
                  entry.shape.Delete();
                }
                touched = true;
                break;
              }
              case 'duplicate': {
                await duplicateShape(slide, command, remap);
                touched = true;
                break;
              }
              case 'group': {
                const entries = (command.ids || []).map(id => resolveEntry(slide, id, remap));
                if (entries.length < 2) throw new Error('编组至少需要两个对象。');
                let group = shapeRange(slide, entries).Group();
                if (group?.Count !== undefined) group = items(group)[0];
                if (!group) throw new Error('WPS 编组后未返回组合对象。');
                if (command.name !== undefined) group.Name = String(command.name);
                if (command.tags !== undefined) writeTags(group, { ...readTags(group), ...command.tags });
                touched = true;
                break;
              }
              case 'ungroup': {
                const shape = resolveShape(slide, command.id, remap);
                shape.Ungroup();
                touched = true;
                break;
              }
              case 'select': {
                const entries = (command.ids || []).map(id => resolveEntry(slide, id, remap));
                if (entries.some(entry => !entry)) throw new Error('所选对象已改变，请重新选择。');
                const targetIndex = Number(optional(() => slide.SlideIndex)) - 1;
                if (targetIndex !== active.slideIndex) {
                  try { active.window.View.GotoSlide(targetIndex + 1); } catch { slide.Select(); }
                }
                shapeRange(slide, entries).Select();
                selectionMemory.set(presentationKey(presentation) + '\u0000' + targetIndex, entries.map(entry => String(entry.shape.Id)));
                break;
              }
              case 'zOrder': {
                const shape = resolveShape(slide, command.id, remap);
                zOrder(shape, command.position, command.index);
                touched = true;
                break;
              }
              case 'replaceImage': {
                await replaceObject(slide, command, remap);
                touched = true;
                break;
              }
              case 'insertSlide': {
                const index = command.index === undefined ? Number(presentation.Slides.Count) : Number(command.index);
                if (!Number.isInteger(index) || index < 0 || index > Number(presentation.Slides.Count)) throw new Error('新增幻灯片索引越界。');
                if (command.layout && command.layout !== 'blank') throw new Error('WPS 只支持空白版式新增幻灯片。');
                let added;
                if (typeof presentation.Slides.Add === 'function') added = presentation.Slides.Add(index + 1, BLANK_LAYOUT);
                else {
                  const layout = optional(() => presentation.Slides.Item(1).CustomLayout);
                  if (!layout) throw new Error('WPS 无法取得幻灯片版式。');
                  added = presentation.Slides.AddSlide(index + 1, layout);
                }
                if (added?.Count !== undefined) added = items(added)[0];
                if (!added) throw new Error('WPS 未能新增幻灯片。');
                if (command.name !== undefined) added.Name = String(command.name);
                touched = true;
                break;
              }
              case 'deleteSlides': {
                const indices = [...new Set(command.indices || [])].map(Number);
                if (!indices.length || indices.some(value => !Number.isInteger(value) || value < 0 || value >= Number(presentation.Slides.Count))) throw new Error('删除页面索引无效。');
                if (Number(presentation.Slides.Count) - indices.length < 1) throw new Error('必须保留至少一页。');
                for (const index of indices.sort((a, b) => b - a)) presentation.Slides.Item(index + 1).Delete();
                touched = true;
                break;
              }
              case 'moveSlide': {
                const from = Number(command.index), to = Number(command.toIndex);
                if (!Number.isInteger(from) || !Number.isInteger(to) || from < 0 || to < 0 || from >= Number(presentation.Slides.Count) || to >= Number(presentation.Slides.Count)) throw new Error('移动页面索引越界。');
                if (from !== to) presentation.Slides.Item(from + 1).MoveTo(to + 1);
                touched = true;
                break;
              }
              default: throw new Error(`未知的 WPS 原生编辑命令：${command.op}`);
            }
            if (['insertSlide', 'deleteSlides', 'moveSlide'].includes(command.op)) {
              const current = activeData();
              active.slide = current.slide;
              active.slideIndex = current.slideIndex;
            }
          }
          if (options.readback === false) {
            lightweight = true;
            return { applied: commands.length };
          }
          lightweight = !rich && options.geometry !== false && (options.geometry === true || !geometryUpdate);
          if (lightweight) {
            const model = geometryContext({ scope: 'selection' }, [...movedIds]);
            if (model) return model;
            // 组内选中无法用原生几何建模：按富契约返回，并重新校正指纹。
            lightweight = false;
          }
          const context = await buildContext({ scope: 'slide' });
          return context;
        } finally {
          try { await release([...temporaryAssets]); } catch { /* inserted media is embedded; release is cleanup only */ }
          temporaryAssets.clear();
          if (lightweight) {
            mutationDepth--;
            if (touched) { invalidateFingerprint(); schedule('apply'); }
          } else await finishMutation(touched);
        }
      });
    },

    async images(ids, options = {}) {
      if (!ids?.length) return [];
      return tracked(() => suppressed(async () => {
        const active = activeData();
        const pptx = await snapshotPptx(active.presentation);
        return imagesInternal(ids.map(String), active.slideIndex, pptx, { original: options.original !== false });
      }));
    },

    async fragment(ids) {
      return tracked(() => suppressed(() => fragmentInternal(ids || [])));
    },

    async insertFragment(fragment, options = {}) {
      if (!fragment?.pptx) throw new Error('片段缺少可编辑的 PPTX 数据。');
      return tracked(async () => {
        const active = activeData();
        const destinationSlide = active.slide;
        const destinationWidth = Number(optional(() => active.presentation.PageSetup.SlideWidth)) || 0;
        const destinationHeight = Number(optional(() => active.presentation.PageSetup.SlideHeight)) || 0;
        const file = await materialize({ data: fragment.pptx, mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', name: 'fragment.pptx' });
        let clipboard, source, context, changed = false, failure, insertedIds = [], committed = false;
        mutationDepth++;
        try {
          clipboard = await nativeExecute('snapshotClipboard');
          source = await openHidden(file.path);
          const sourceSlide = source.Slides.Item(1);
          const sourceShapes = sourceSlide.Shapes;
          if (!Number(sourceShapes.Count)) throw new Error('片段中没有可插入的对象。');
          const metadata = items(sourceShapes).map(shape => ({ name: String(shape.Name), alternativeText: String(shape.AlternativeText || ''), tags: readTags(shape) }));
          sourceShapes.Range().Copy();
          const pasted = destinationSlide.Shapes.Paste();
          if (!pasted || !Number(pasted.Count)) throw new Error('WPS 无法将片段粘贴到当前幻灯片。');
          changed = true;
          const pastedShapes = items(pasted);
          insertedIds = pastedShapes.map(shape => String(shape.Id));
          const sourceBounds = fragment.bounds || boundsOf(pastedShapes.map(shape => ({ left: Number(shape.Left), top: Number(shape.Top), width: Number(shape.Width), height: Number(shape.Height) })));
          let scaleX = 1, scaleY = 1;
          if (options.width !== undefined && sourceBounds.width > 0) scaleX = Number(options.width) / sourceBounds.width;
          if (options.height !== undefined && sourceBounds.height > 0) scaleY = Number(options.height) / sourceBounds.height;
          if (options.width !== undefined && options.height === undefined) scaleY = scaleX;
          if (options.height !== undefined && options.width === undefined) scaleX = scaleY;
          const targetWidth = sourceBounds.width * scaleX, targetHeight = sourceBounds.height * scaleY;
          const left = options.left !== undefined ? Number(options.left) : options.center && destinationWidth
            ? (destinationWidth - targetWidth) / 2 : sourceBounds.left || 0;
          const top = options.top !== undefined ? Number(options.top) : options.center && destinationHeight
            ? (destinationHeight - targetHeight) / 2 : sourceBounds.top || 0;
          pastedShapes.forEach((shape, index) => {
            const original = { left: Number(shape.Left), top: Number(shape.Top), width: Number(shape.Width), height: Number(shape.Height) };
            const locked = shape.LockAspectRatio;
            shape.LockAspectRatio = FALSE;
            shape.Width = original.width * scaleX;
            shape.Height = original.height * scaleY;
            shape.Left = left + (original.left - sourceBounds.left) * scaleX;
            shape.Top = top + (original.top - sourceBounds.top) * scaleY;
            shape.LockAspectRatio = locked === TRUE ? TRUE : FALSE;
            const originalMetadata = metadata[index];
            if (originalMetadata) {
              shape.Name = originalMetadata.name;
              shape.AlternativeText = originalMetadata.alternativeText;
              writeTags(shape, originalMetadata.tags);
            }
          });
        } catch (error) { failure = error; }
        try { closeHidden(source); } catch (error) { failure = failure ? new AggregateError([failure, error], '片段已插入，但临时演示文稿未能关闭。') : error; }
        try { if (clipboard) await nativeExecute('restoreClipboard', { items: clipboard.items }); }
        catch (error) { if (!failure) failure = error; }
        try { await release([file.path]); } catch (error) { if (!failure) failure = error; }
        temporaryAssets.delete(file.path);
        if (!failure) {
          try { context = await buildContext({ scope: 'slide' }); committed = true; }
          catch (error) { failure = error; }
        }
        if (failure && !committed && insertedIds.length) {
          try {
            for (const id of insertedIds) { const entry = findNativeEntry(destinationSlide, id); if (entry) entry.shape.Delete(); }
          } catch (error) { failure = new AggregateError([failure, error], '插入片段失败且无法撤销新增对象。'); }
        }
        await finishMutation(changed && committed);
        if (failure) throw failure;
        return context;
      });
    },

    async rasterizeImage(id, options = {}) {
      return tracked(() => suppressed(async () => {
        const active = activeData();
        const maximum = Number(options.maximum ?? 4096);
        if (!(maximum > 0)) throw new Error('栅格最大尺寸必须大于 0。');
        const ids = [String(id)];
        if (options.original === false) {
          const fragment = await fragmentInternal(ids);
          const file = await materialize({ data: fragment.pptx, mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', name: 'raster.pptx' });
          return withRenderedPdf(file.path, async pdfPath => {
            const bounds = fragment.bounds || { left: 0, top: 0, width: 0, height: 0 };
            const dpi = clamp(Math.floor(maximum / Math.max(bounds.width, bounds.height) * 72), 1, MAX_RASTER_DPI);
            const pages = await nativeExecute('rasterPDF', { path: pdfPath, indices: [0], dpi, format: 'png', bounds });
            const page = pages?.[0];
            if (!page?.data) throw new Error('WPS 未能渲染该对象。');
            return { data: page.data, mime: page.mime || 'image/png', width: page.width, height: page.height, digest: await digest(page.data) };
          });
        }
        const pptx = await snapshotPptx(active.presentation);
        const isolated = await request('pptx/imageFragment', { pptx, slideIndex: active.slideIndex, id: String(id), original: true });
        if (!isolated?.pptx) throw new Error('未能生成原始图片的隔离演示文稿。');
        if (!isolated.bounds?.width || !isolated.bounds?.height) throw new Error('原始图片的边界无效。');
        const file = await materialize({ data: isolated.pptx, mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', name: 'image-fragment.pptx' });
        return withRenderedPdf(file.path, async pdfPath => {
          const dpi = clamp(Math.floor(maximum / Math.max(isolated.bounds.width, isolated.bounds.height) * 72), 1, MAX_RASTER_DPI);
          const pages = await nativeExecute('rasterPDF', { path: pdfPath, indices: [0], dpi, format: 'png', bounds: isolated.bounds });
          const page = pages?.[0];
          if (!page?.data) throw new Error('WPS 未能栅格化原始图片。');
          const result = { data: page.data, mime: page.mime || 'image/png', width: page.width, height: page.height,
            digest: await digest(page.data) };
          if (isolated.source !== undefined) result.source = isolated.source;
          if (isolated.mime !== undefined) result.mimeOriginal = isolated.mime;
          return result;
        });
      }));
    },

    async export(options = {}) {
      return tracked(() => suppressed(() => exportInternal(options)));
    },

    onChanged(callback) {
      if (typeof callback !== 'function') throw new Error('变更回调必须是函数。');
      listeners.add(callback);
      startWatching();
      return () => {
        listeners.delete(callback);
        if (!listeners.size) stopWatching();
      };
    },

    dispose() {
      listeners.clear();
      stopWatching();
    },

    // 兼容现有基础任务窗格；完整功能走 context/apply。
    async selection(transform) {
      return tracked(async () => {
        const active = activeData();
        const selection = selectionState(active.window);
        if (selection.type !== PP_SELECTION_SHAPES) { if (transform) transform([]); return []; }
        const shapes = selection.shapes.map(shape => ({ id: String(shape.Id), name: String(shape.Name),
          left: Number(shape.Left), top: Number(shape.Top), width: Number(shape.Width), height: Number(shape.Height) }));
        if (!transform) return shapes;
        const patches = transform(shapes);
        selection.shapes.forEach((shape, index) => {
          const patch = patches[index] || {};
          if ('left' in patch) shape.Left = Number(patch.left);
          if ('top' in patch) shape.Top = Number(patch.top);
        });
        return shapes.map((shape, index) => ({ ...shape, ...patches[index] }));
      });
    },
  };
  return host;
}
