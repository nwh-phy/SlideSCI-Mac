import { UUID_TAG, FOV_PREFIX, SCALE_PREFIX, uuid, tag, parse, serializeTag, pictures, selected, content, flatten, imagesFor, visibleFov, scaleOf, writeRecord, element, control, findUUID } from './images-science-core.js';

export const colorPresets = [{ value: '#00ff00', label: '绿' }, { value: '#ff0000', label: '红' }, { value: '#0000ff', label: '蓝' }, { value: '#ffffff', label: '灰' }, { value: '#00ffff', label: '青' }, { value: '#ff00ff', label: '品红' }, { value: '#ffff00', label: '黄' }];
const videoPattern = /\.(mp4|m4v|mov|webm|avi|wmv|asf|mpg|mpeg|mpe)$/i;
const imagePattern = /\.(png|jpe?g|bmp|gif|tiff?|svg|emf|wmf)$/i;
const videoTypes = { mp4: 'video/mp4', m4v: 'video/mp4', mov: 'video/quicktime', webm: 'video/webm', avi: 'video/x-msvideo', wmv: 'video/x-ms-wmv', asf: 'video/x-ms-asf', mpg: 'video/mpeg', mpeg: 'video/mpeg', mpe: 'video/mpeg' };
const videoMime = (name, mime) => mime?.startsWith('video/') ? mime : videoTypes[name.split('.').at(-1).toLowerCase()] || 'video/mp4';
// Native media inspection returns the intrinsic picture size, never the poster-frame of the inserted player.
async function videoSize(ctx, path) {
  const info = await ctx.services.request('native/execute', { host: ctx.host.kind, action: 'mediaInfo', path });
  const width = Number(info?.width) || Number(info?.videoWidth), height = Number(info?.height) || Number(info?.videoHeight);
  if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0) throw new Error('无法读取视频画面的原始尺寸。');
  return { width, height };
}

// Pure decode/render work may overlap, native host mutations never do: bounded, order-preserving.
const RENDER_LIMIT = 2;
async function concurrentMap(items, limit, run) {
  const results = new Array(items.length), workers = Math.min(limit, items.length);
  let next = 0;
  const worker = async () => { while (next < items.length) { const index = next++; try { results[index] = { value: await run(items[index], index) }; } catch (error) { results[index] = { error }; } } };
  await Promise.all(Array.from({ length: workers }, () => worker()));
  return results;
}
// Surface the first failure in input order, like a sequential loop would.
function settle(results) {
  for (const result of results) if (result.error) throw result.error;
  return results.map(result => result.value);
}

const sourcePath = shape => {
  const lines = (content(shape).alternativeText || shape.alternativeText || '').split(/[\r\n]+/);
  const record = lines.filter(x => x.startsWith('原图路径：') || x.startsWith('视频路径：')).at(-1);
  return record ? record.slice(record.indexOf('：') + 1) : parse(tag(shape, 'SLIDESCI_SOURCE_PATH'), tag(shape, 'SLIDESCI_SOURCE_PATH'));
};
const tiffPattern = /\.tiff?$/i;
function availableName(name, used) { let value = name, i = 2; const dot = name.lastIndexOf('.'), stem = dot > 0 ? name.slice(0, dot) : name, ext = dot > 0 ? name.slice(dot) : ''; while (used.has(value.toLocaleLowerCase())) value = `${stem} (${i++})${ext}`; used.add(value.toLocaleLowerCase()); return value; }
export async function insertSourceFiles(ctx, files, { position, shapeIds = [] } = {}) {
  if (!files?.length) return { count: 0 };
  const updating = Array.isArray(shapeIds) && shapeIds.length > 0;
  let context = await ctx.host.context({ scope: 'slide' });
  const used = new Set(context.shapes.map(s => s.name.toLocaleLowerCase())), errors = [], inserted = [], directUpdates = [];
  let replaced = false;
  async function flushDirectUpdates() {
    if (!directUpdates.length) return;
    const pending = directUpdates.splice(0, directUpdates.length);
    try {
      await ctx.host.apply(pending.map(item => item.command), { readback: false });
      for (const { existing, metadata } of pending) {
        existing.name = metadata.name;
        existing.alternativeText = metadata.alternativeText;
        existing.tags = { ...(existing.tags || {}), ...metadata.tags };
        inserted.push(existing.id);
      }
    } catch (error) {
      for (const item of pending) errors.push(`${item.name}：${error.message}`);
    }
  }
  const media = files.map(file => {
    const source = file.path, name = file.name || source?.split('/').at(-1) || '图片.png';
    const tiff = tiffPattern.test(name) || /tiff/i.test(file.mime || '');
    return { name, tiff, video: videoPattern.test(name) || !!file.mime?.startsWith('video/'), image: imagePattern.test(name) || !!file.mime?.startsWith('image/') };
  });
  const needsNormalize = index => {
    const kind = media[index];
    return kind.image && !kind.video && (!updating || kind.tiff);
  };
  // Decode the next small batch while this one is inserted, so a multi-file paste does not wait decode→insert→decode.
  const load = start => concurrentMap(files.slice(start, start + RENDER_LIMIT), RENDER_LIMIT, (file, offset) => {
    return needsNormalize(start + offset) ? ctx.services.request('images/normalize', file) : null;
  });
  const batches = [];
  let started = 0, taken = 0, current = null;
  const ensure = count => { while (started < count) { batches[started] = load(started * RENDER_LIMIT); started++; } };
  ensure(1);
  for (let index = 0; index < files.length; index++) {
    const file = files[index], { name, video, image, tiff } = media[index], source = file.path;
    try {
      if (!video && !image) throw new Error('不支持的图片或视频格式。');
      if (!source && video) throw new Error('插入视频需要从原生文件选择器取得文件路径。');
      const existing = updating ? context.shapes.find(s => String(s.id) === String(shapeIds[index])) : null;
      // Updating pasted shapes must never fall back to inserting duplicate copies if a shape was
      // already replaced, deleted, or belongs to a different slide.
      if (updating && !existing) continue;
      const target = existing ? content(existing) : null;
      const group = Math.floor(index / RENDER_LIMIT);
      if (group >= taken) { taken = group + 1; current = batches[group]; batches[group] = null; ensure(group + 2); }
      const normalized = (await current)[index - group * RENDER_LIMIT];
      let prepared = null;
      if (needsNormalize(index)) {
        if (normalized.error) throw normalized.error;
        prepared = normalized.value;
      } else if (video && !existing) {
        prepared = { ...await videoSize(ctx, source), mime: videoMime(name, file.mime) };
      }
      let alternativeText = existing?.alternativeText || '';
      if (source) alternativeText += `${alternativeText ? '\n' : ''}${video ? '视频路径：' : '原图路径：'}${source}`;
      if (prepared?.fov) alternativeText = writeRecord(alternativeText, FOV_PREFIX, prepared.fov);
      const metadata = { name: existing?.name || availableName(name, used), alternativeText, tags: { [UUID_TAG]: tag(existing, UUID_TAG) || uuid(), SLIDESCI_SOURCE_PATH: source ? serializeTag(source) : '', SLIDESCI_SOURCE_NAME: serializeTag(name) } };
      if (existing) {
        if (!video && tiff && prepared) {
          await flushDirectUpdates();
          context = await ctx.host.apply([{ op: 'update', id: existing.id, properties: metadata }]);
          const vectorSource = /image\/(?:x-)?(?:emf|wmf)|image\/(?:x-ms-)?bmp/i.test(prepared.mime || '');
          context = await ctx.host.apply([{ op: 'replaceImage', id: target.id, data: prepared.data, mime: prepared.mime, source, preserveCrop: true, ...(vectorSource && source ? { path: source } : {}) }]);
          replaced = true;
          const created = context.shapes.find(s => tag(s, UUID_TAG) === metadata.tags[UUID_TAG]);
          if (created) inserted.push(created.id);
        } else {
          // Non-TIFF images and videos pasted by the host already carry their embedded media;
          // batch metadata-only updates so pasting many figures takes one lightweight host call.
          directUpdates.push({ existing, metadata, name, command: { op: 'update', id: existing.id, properties: metadata } });
        }
      } else {
        await flushDirectUpdates();
        const scale = Math.min(1, context.width * .9 / (prepared.width * .75), context.height * .9 / (prepared.height * .75));
        const width = prepared.width * .75 * scale, height = prepared.height * .75 * scale;
        const left = position ? Math.max(0, Math.min(context.width - width, position.x - width / 2)) : (context.width - width) / 2;
        const top = position ? Math.max(0, Math.min(context.height - height, position.y - height / 2)) : (context.height - height) / 2;
        const vectorSource = /image\/(?:x-)?(?:emf|wmf)|image\/(?:x-ms-)?bmp/i.test(prepared.mime || '');
        const imageArguments = vectorSource && source ? { path: source } : { data: prepared.data };
        const command = { op: video ? 'insertMedia' : 'insertImage', ...(video ? { path: source } : imageArguments), mime: prepared.mime, source, left, top, width, height, ...metadata };
        try { context = await ctx.host.apply([command]); }
        catch (error) { await removeInserted(ctx, [metadata.tags[UUID_TAG]], error); throw error; }
        const created = context.shapes.find(s => tag(s, UUID_TAG) === metadata.tags[UUID_TAG]);
        if (created) inserted.push(created.id);
      }
    } catch (error) { errors.push(`${name}：${error.message}`); }
  }
  await flushDirectUpdates();
  if (inserted.length && (!updating || replaced)) await ctx.host.apply([{ op: 'select', ids: inserted }]);
  if (errors.length) throw new Error(`已插入或记录 ${inserted.length} 个文件；以下文件失败：\n${errors.join('\n')}`);
  return { count: inserted.length };
}
// Native paste/drop interception calls this with file capabilities and optional existing shape IDs.
export const handleNativePaste = (ctx, event) => insertSourceFiles(ctx, event.files, event);
const rawPseudoColor = shape => tag(shape, 'SLIDESCI_PSEUDO_COLOR') ?? tag(content(shape), 'SLIDESCI_PSEUDO_COLOR');
const normalizePseudo = value => value?.startsWith('#') ? value : value ? '#' + (Number(value) >>> 0).toString(16).padStart(8, '0').slice(-6) : '#ffffff';
const pseudoColorOf = shape => normalizePseudo(rawPseudoColor(shape));
const peakOf = hex => Math.max(...[1, 3, 5].map(index => parseInt(hex.slice(index, index + 2), 16)));
// A retained pseudo-color copy is a plain picture: zoom links, scale ownership and the old UUID must not follow it.
function copyTags(source, outer) {
  const tags = {};
  for (const shape of [outer, source]) for (const [key, value] of Object.entries(shape?.tags || {})) {
    const name = key.toUpperCase();
    if (name === UUID_TAG || name.startsWith('SLIDESCI_ZOOM_') || name.startsWith('SLIDESCI_SCALE_')) continue;
    tags[key] = value;
  }
  return tags;
}
async function removeInserted(ctx, keys, error) {
  if (!keys.length) return;
  try {
    const context = await ctx.host.context({ scope: 'slide' });
    const ids = flatten(context.shapes).filter(shape => keys.includes(tag(shape, UUID_TAG))).map(shape => shape.id);
    if (ids.length) await ctx.host.apply([{ op: 'delete', ids }]);
  } catch (cleanup) { error.message += `；清理新插入的伪彩图失败：${cleanup.message}`; }
}
export async function applyPseudoColors(ctx, shapes, colors, keepOriginal = false) {
  // One host round trip for every original; render every channel before editing so a codec or I/O failure never leaves a half-applied selection.
  const originals = await imagesFor(ctx, shapes);
  const rendered = settle(await concurrentMap(shapes, RENDER_LIMIT, async (shape, index) => {
    const color = colors[index], original = originals[index];
    return { color, source: original.source, image: await ctx.services.request('images/color', { ...original, crop: null, color, sourcePeak: peakOf(pseudoColorOf(shape)) }) };
  }));
  // replaceImage may return a new picture id; a persistent UUID lets us find it and prevents group content lookup from breaking.
  const keys = keepOriginal ? [] : shapes.map(shape => tag(shape, UUID_TAG) || tag(content(shape), UUID_TAG));
  const seeds = [];
  if (!keepOriginal) keys.forEach((key, index) => { if (!key) { keys[index] = uuid(); seeds.push({ op: 'update', id: shapes[index].id, properties: { tags: { [UUID_TAG]: keys[index] } } }); } });
  if (seeds.length) await ctx.host.apply(seeds, { readback: false });
  const offset = keepOriginal ? Math.max(...shapes.map(shape => shape.left + shape.width)) - Math.min(...shapes.map(shape => shape.left)) + 18 : 0;
  const created = [], commands = [];
  for (let index = 0; index < shapes.length; index++) {
    const shape = shapes[index], source = content(shape), { color, image, source: imageSource } = rendered[index];
    if (keepOriginal) {
      const key = uuid(), tags = copyTags(source, shape);
      const rotation = (shape.rotation || 0) + (source === shape ? 0 : source.rotation || 0);
      tags[UUID_TAG] = key; tags.SLIDESCI_PSEUDO_COLOR = color;
      const alternativeText = writeRecord(writeRecord(source.alternativeText, SCALE_PREFIX, null), '标尺（SlideSCI）：', null);
      commands.push({ op: 'insertImage', data: image.data, mime: image.mime, source: imageSource, left: shape.left + offset, top: shape.top, width: source.width, height: source.height, name: `${source.name}_伪彩`, alternativeText, tags, rotation, crop: source.crop, line: source.line });
      created.push({ key, source, rotation });
    } else commands.push({ op: 'replaceImage', id: source.id, data: image.data, mime: image.mime, source: imageSource, preserveCrop: true });
  }
  let context;
  try { context = await ctx.host.apply(commands); }
  catch (error) { await removeInserted(ctx, created.map(item => item.key), error); throw error; }
  const updates = [];
  try {
    for (const item of created) {
      const shape = findUUID(context, item.key);
      if (!shape) throw new Error('插入伪彩图后未找到新对象。');
      const properties = {};
      properties.rotation = item.rotation;
      if (item.source.crop) properties.crop = item.source.crop;
      if (item.source.line) properties.line = item.source.line;
      if (Object.keys(properties).length) updates.push({ op: 'update', id: shape.id, properties });
    }
    if (!keepOriginal) for (let index = 0; index < shapes.length; index++) {
      const shape = shapes[index], target = content(shape) === shape ? findUUID(context, keys[index]) : shape;
      if (!target) throw new Error('替换伪彩后未找到原图对象。');
      updates.push({ op: 'update', id: target.id, properties: { tags: { SLIDESCI_PSEUDO_COLOR: colors[index] } } });
    }
    if (updates.length) await ctx.host.apply(updates, { readback: false });
  } catch (error) { await removeInserted(ctx, created.map(item => item.key), error); throw error; }
  await ctx.services.request('images/settings', { values: { pseudoColor: colors.at(-1) } });
  const ids = (keepOriginal ? created.map(item => findUUID(context, item.key)?.id) : shapes.map((shape, index) => content(shape) === shape ? findUUID(context, keys[index])?.id ?? shape.id : shape.id)).filter(Boolean);
  if (ids.length) await ctx.host.apply([{ op: 'select', ids }]);
  return { count: shapes.length };
}
async function mergeChannels(ctx) {
  const context = await ctx.host.context(), shapes = pictures(context, 2, 7), first = shapes[0], aspect = content(first).width / content(first).height;
  if (shapes.some(s => Math.abs((content(s).width / content(s).height) / aspect - 1) > .005)) throw new Error('各通道可见画面的宽高比必须一致；合并按画面坐标叠加，不自动配准。');
  if (shapes.some(s => Math.abs(((s.rotation - first.rotation + 540) % 360) - 180) > .01)) throw new Error('各通道图片的旋转角度必须一致。');
  let calibration = null;
  for (const shape of shapes) {
    const f = visibleFov(shape); if (!f) continue;
    if (calibration && ['width', 'height'].some(k => calibration[k] > 0 && f[k] > 0 && Math.abs(f[k] / calibration[k] - 1) > .005)) throw new Error('各通道的可见 FOV 标定不一致，请先统一标定和裁剪范围。');
    calibration = { width: calibration?.width || f.width, height: calibration?.height || f.height, unit: 'μm', source: '由合并通道的可见画面标定' };
  }
  const images = await imagesFor(ctx, shapes), rendered = await ctx.services.request('images/merge', { images });
  let alternativeText = '合并通道（RGB 相加，超过 255 截断；不自动配准）：\n' + shapes.map(s => s.name).join('\n');
  if (calibration) alternativeText = writeRecord(alternativeText, FOV_PREFIX, calibration);
  // The merged raster is already cropped to each channel's visible area, so this FOV calibrates the
  // whole new picture; a later native crop must still reduce it.
  const id = uuid(), result = await ctx.host.apply([{ op: 'insertImage', ...rendered, left: Math.max(...shapes.map(s => s.left + s.width)) + 18, top: first.top, width: content(first).width, height: content(first).height, rotation: first.rotation || 0, name: `合并通道_${shapes.length}通道`, alternativeText, tags: { [UUID_TAG]: id } }]);
  const merged = result.shapes.find(s => tag(s, UUID_TAG) === id), scale = shapes.map(scaleOf).find(Boolean);
  if (merged && scale && calibration) { const { applyScaleBar } = await import('./science.js'); await applyScaleBar(ctx, merged, calibration, scale); }
  else if (merged) await ctx.host.apply([{ op: 'select', ids: [merged.id] }]);
  return { count: shapes.length };
}
async function mountPseudo(container, ctx) {
  const status = element('p', '选择图片后读取预览。', { role: 'status' }), list = element('div'), toolbar = element('div');
  const reload = element('button', '读取所选图片'), applyAll = element('button', '当前颜色应用到全部'), apply = element('button', '应用伪彩');
  const keep = control(toolbar, '保留原图', 'checkbox', false); toolbar.append(reload, applyAll, apply); container.append(toolbar, list, status);
  let shapes = [], colors = [], images = [], rows = [], active = 0, generation = 0, disposed = false;
  // Preview renders are CPU work in the local service; queue them instead of firing one request per selected image.
  let rendering = 0;
  const waiting = [];
  const enter = () => rendering < RENDER_LIMIT ? (rendering++, Promise.resolve()) : new Promise(resolve => waiting.push(() => { rendering++; resolve(); }));
  const leave = () => { const next = waiting.shift(); if (next) next(); else rendering--; };
  async function preview(index, row, ownGeneration) {
    const color = colors[index], shape = shapes[index], image = images[index], token = ++row.token;
    if (!image) return;
    await enter();
    try {
      if (disposed || ownGeneration !== generation || token !== row.token) return;
      const rendered = await ctx.services.request('images/color', { ...image, maximum: 640, color, sourcePeak: peakOf(pseudoColorOf(shape)) });
      if (!disposed && ownGeneration === generation && token === row.token) { row.rendered = color; row.image.src = `data:${rendered.mime};base64,${rendered.data}`; }
    } finally { leave(); }
  }
  function render() {
    list.replaceChildren(); rows = []; const ownGeneration = ++generation;
    shapes.forEach((shape, index) => {
      const row = element('fieldset'), legend = element('legend', shape.name), image = element('img', null, { alt: `${shape.name} 伪彩预览`, width: 240 });
      const presets = control(row, '预设', 'select', colors[index], colorPresets), custom = control(row, '自定义颜色', 'color', colors[index]);
      row.prepend(legend); row.append(image); list.append(row);
      const state = rows[index] = { image, presets, custom, token: 0, rendered: null };
      const changed = async color => { active = index; colors[index] = color; custom.value = color; presets.value = color; try { await preview(index, state, ownGeneration); status.textContent = ''; } catch (error) { status.textContent = error.message; } };
      presets.addEventListener('change', () => changed(presets.value)); custom.addEventListener('input', () => changed(custom.value));
      preview(index, state, ownGeneration).catch(error => { status.textContent = error.message; });
    });
  }
  reload.onclick = async () => { try { shapes = pictures(await ctx.host.context()); images = await imagesFor(ctx, shapes); const saved = await ctx.services.request('images/settings', {}); colors = shapes.map((s, i) => saved.pseudoColor || (rawPseudoColor(s) ? normalizePseudo(rawPseudoColor(s)) : colorPresets[i % 7].value)); render(); status.textContent = `已读取 ${shapes.length} 张图片。`; } catch (error) { status.textContent = error.message; } };
  // Rows already showing the active color keep their rendered preview instead of re-running the render service.
  applyAll.onclick = () => {
    if (!colors.length) return;
    const color = colors[active];
    colors = colors.map(() => color);
    rows.forEach((row, index) => {
      row.custom.value = color; row.presets.value = color;
      if (row.rendered !== color) preview(index, row, generation).catch(error => { status.textContent = error.message; });
    });
  };
  apply.onclick = async () => { apply.disabled = true; try { if (!shapes.length) throw new Error('请先读取所选图片。'); await applyPseudoColors(ctx, shapes, colors, keep.checked); await reload.onclick(); status.textContent = '伪彩已应用。'; } catch (error) { status.textContent = error.message; } finally { apply.disabled = false; } };
  return () => { disposed = true; generation++; };
}
export const actions = [
  { id: 'images.insert', label: '插入图片并记录来源', section: '图片', run: async ctx => insertSourceFiles(ctx, await ctx.services.pickFiles({ kind: 'images', multiple: true })) },
  { id: 'images.insertVideo', label: '插入视频并记录来源', section: '图片', run: async ctx => insertSourceFiles(ctx, await ctx.services.pickFiles({ kind: 'video', multiple: true })) },
  { id: 'images.pasteFiles', label: '粘贴图片或视频文件', section: '图片', run: async ctx => { const files = await ctx.services.readClipboardFiles(); if (!files?.length) throw new Error('剪贴板中没有图片或视频文件。'); return insertSourceFiles(ctx, files); } },
  { id: 'images.reveal', label: '打开原图或视频所在位置', section: '图片', run: async ctx => { const items = selected(await ctx.host.context()); if (items.length !== 1) throw new Error('请选择一个图片或视频对象。'); const path = sourcePath(items[0]); if (!path?.startsWith('/') || /[\0\r\n]/.test(path)) throw new Error('该对象没有有效的原始文件路径。'); await ctx.services.reveal(path); } },
  { id: 'images.merge', label: '合并通道', section: '图片', description: '2–7 个通道进行 RGB 饱和相加；保留原图，不自动配准。', run: mergeChannels }
];
export const panels = [{ id: 'images.pseudoColor', label: '伪彩预览与设置', section: '图片', mount: mountPseudo }];
