import { element, button, input, select, statusNode, descriptorData, filesOf } from './library-ai-ui.js';

const formats = ['png', 'jpg', 'bmp', 'pdf'];
const formatOptions = formats.map(value => ({ value, label: value.toUpperCase() }));
const rangeOptions = [{ value: 'current', label: '当前页' }, { value: 'selected', label: '选中的页面' }, { value: 'all', label: '全部页面' }, { value: 'custom', label: '自定义页码' }];
// Mac questions go to the Mac fork; the Windows author keeps attribution and the feature guide.
const links = {
  feedback: 'https://github.com/nwh-phy/SlideSCI-Mac/issues',
  github: 'https://github.com/nwh-phy/SlideSCI-Mac',
  install: 'https://github.com/nwh-phy/SlideSCI-Mac/blob/main/mac/DISTRIBUTION.md',
  help: 'https://www.yuque.com/achuan-2/blog/etzcergpmb4rr2sk/',
  developer: 'https://www.github.com/achuan-2',
  sponsor: 'https://fastly.jsdelivr.net/gh/Achuan-2/PicBed/assets/20241128221208-2024-11-28.png'
};
const linkLabels = { feedback: 'Mac 版问题反馈', github: 'Mac 版源代码', install: 'Mac 版安装与常见问题', help: '功能使用文档（原版）', developer: '原版作者主页', sponsor: '赞助原版作者' };
const maintainer = 'Mac 版：nwh-phy（问题请在 Mac 版 GitHub 反馈）；原版作者：Achuan-2';
function safeName(value, fallback = 'SlideSCI') {
  const result = String(value || fallback).trim().replace(/[\x00-\x1f/\\:*?"<>|]/gu, '_').replace(/\.(pptx|ppt|pdf|png|jpg|bmp|svg)$/iu, '');
  if (!result || result === '.' || result === '..' || result.length > 180) throw new Error('请输入有效的文件名（最多180字）。');
  return result;
}
export function pageIndices(range, context, specification = '') {
  const count = context.slides?.length || 0;
  if (!count) throw new Error('请先打开至少包含一页的演示文稿。');
  if (range === 'current') return [context.slideIndex];
  if (range === 'all') return context.slides.map(slide => slide.index);
  if (range === 'selected') { if (!context.selectedSlideIndices?.length) throw new Error('请先在页面缩略图中选择需要导出的页面。'); return [...new Set(context.selectedSlideIndices)].sort((a, b) => a - b); }
  if (range !== 'custom') throw new Error('无效的导出范围。');
  const result = new Set();
  for (const token of String(specification).split(/[,，;；\s]+/u).filter(Boolean)) {
    const match = /^(\d+)(?:[-–](\d+))?$/u.exec(token);
    if (!match) throw new Error('页码格式应为1,3-5，页码从1开始。');
    const first = Number(match[1]); const last = Number(match[2] || match[1]);
    if (!Number.isInteger(first) || first < 1 || last < first || last > count) throw new Error(`页码必须在1到${count}之间，范围起点不能大于终点。`);
    for (let page = first; page <= last; page++) result.add(page - 1);
  }
  if (!result.size) throw new Error('请输入要导出的页码。');
  return [...result].sort((a, b) => a - b);
}
async function defaults(ctx) {
  return { range: 'current', format: await ctx.settings.get('export.format', 'png'), dpi: await ctx.settings.get('export.dpi', 300), separatePdf: await ctx.settings.get('export.separatePdf', false), openFolder: await ctx.settings.get('export.openFolder', true) };
}
async function exportSlides(ctx, values) {
  const options = { ...await defaults(ctx), ...values };
  if (!formats.includes(options.format)) throw new Error('请选择有效的导出格式。');
  const dpi = Number(options.dpi);
  if (!Number.isFinite(dpi) || dpi < 36 || dpi > 2400) throw new Error('DPI必须介于36与2400之间。');
  // Page planning only needs the deck outline, not the exported package.
  const context = await ctx.host.summary();
  const indices = pageIndices(options.range, context, options.pages);
  const filename = safeName(options.filename || context.title);
  const folder = await ctx.services.chooseFolder(); if (!folder) return;
  // One host request per export: the host splits per-page PDFs itself
  // (separatePdf is only honoured for PDF and ignored by other formats).
  const results = filesOf(await ctx.host.export({ kind: 'slides', indices, format: options.format, separatePdf: options.separatePdf, dpi, folder, filename }));
  if (!results.length) throw new Error('宿主没有返回任何导出文件。');
  for (const key of ['format', 'dpi', 'separatePdf', 'openFolder']) await ctx.settings.set(`export.${key}`, options[key]);
  if (options.openFolder) await ctx.services.reveal(folder);
  ctx.notify(`已导出${indices.length}页，共${results.length}个文件。`); return { files: results, folder };
}
function selectedPictures(context) {
  const selected = new Set(context.selectionIds); const found = [];
  function visit(shapes, parentSelected = false) {
    for (const shape of shapes) {
      const chosen = parentSelected || selected.has(shape.id);
      if (chosen && shape.type === 'picture') found.push(shape);
      if (shape.children) visit(shape.children, chosen);
    }
  }
  visit(context.shapes); return found;
}
async function originalImages(ctx, copy) {
  const context = await ctx.host.context({ scope: 'selection', images: false });
  const pictures = selectedPictures(context);
  if (!pictures.length) throw new Error('请先选择图片或包含图片的组合。');
  if (copy && pictures.length !== 1) throw new Error('复制原图时请只选择一张图片。');
  const images = await ctx.host.images(pictures.map(shape => shape.id), { original: true });
  if (images.length !== pictures.length) throw new Error('未能读取全部选中图片的原始资源。');
  if (copy) {
    const image = await descriptorData(images[0], ctx.services);
    await ctx.services.writeClipboard({ data: image.data, mime: image.mime }); ctx.notify('已将未裁剪的原始图片复制到剪贴板。'); return image;
  }
  const folder = await ctx.services.chooseFolder(); if (!folder) return;
  const files = [];
  const mimeExtension = { 'image/png': 'png', 'image/jpeg': 'jpg', 'image/gif': 'gif', 'image/tiff': 'tif', 'image/bmp': 'bmp', 'image/svg+xml': 'svg', 'image/x-emf': 'emf', 'image/x-wmf': 'wmf', 'image/webp': 'webp' };
  const used = new Set();
  for (let index = 0; index < images.length; index++) {
    const image = await descriptorData(images[index], ctx.services);
    const original = String(image.source || '').split(/[\\/]/u).pop()?.split('?')[0];
    let name = original && /\.[a-z0-9]{2,6}$/iu.test(original) ? original.replace(/[\x00-\x1f/\\:*?"<>|]/gu, '_') : `${safeName(pictures[index].name, `原图${index + 1}`)}.${mimeExtension[image.mime] || 'bin'}`;
    if (used.has(name.toLocaleLowerCase())) name = `${index + 1}_${name}`;
    used.add(name.toLocaleLowerCase());
    files.push(await ctx.services.request('files/write', { folder, name, data: image.data, mime: image.mime }));
  }
  if (await ctx.settings.get('export.openFolder', true)) await ctx.services.reveal(folder);
  ctx.notify(`已导出${files.length}张原始图片，保留原始文件格式与数据。`); return { files, folder };
}
export const actions = [
  { id: 'export.slides', label: '导出页面', section: '导出', fields: [
    { name: 'range', label: '导出范围', type: 'select', default: 'current', options: rangeOptions },
    { name: 'pages', label: '自定义页码（如1,3-5）', type: 'text' },
    { name: 'format', label: '格式', type: 'select', default: 'png', options: formatOptions },
    { name: 'dpi', label: '分辨率DPI', type: 'number', min: 36, max: 2400, default: 300 },
    { name: 'separatePdf', label: '每页单独导出PDF', type: 'checkbox', default: false },
    { name: 'openFolder', label: '导出后打开文件夹', type: 'checkbox', default: true },
    { name: 'filename', label: '文件名前缀（留空使用文稿名称）', type: 'text' }
  ], run: exportSlides },
  { id: 'export.originalImages', label: '导出原始图片', section: '导出', run: ctx => originalImages(ctx, false) },
  { id: 'export.copyOriginal', label: '复制原始图片', section: '导出', run: ctx => originalImages(ctx, true) },
  { id: 'export.copySelection', label: '复制选中对象为图片', section: '导出', run: async ctx => {
    const { selectionIds } = await ctx.host.summary({ slides: false }); if (!selectionIds.length) throw new Error('请先选择对象。');
    const result = await ctx.host.export({ kind: 'selection', ids: selectionIds, format: 'png', dpi: 300 });
    const file = await descriptorData(filesOf(result)[0], ctx.services); await ctx.services.writeClipboard({ data: file.data, mime: file.mime || 'image/png' }); ctx.notify('已复制选中对象的PNG图片。'); return result;
  } },
  ...Object.entries(links).map(([name, url]) => ({ id: `export.${name}`, label: linkLabels[name], section: '设置与帮助', run: ctx => ctx.services.openExternal(url) })),
  { id: 'export.version', label: '当前版本', section: '设置与帮助', run: async ctx => { const info = await ctx.services.request('library/info', {}); ctx.notify(`SlideSCI Mac ${info.version} · Windows原版 ${info.upstreamVersion}`); return info; } },
  { id: 'export.about', label: '关于', section: '设置与帮助', run: ctx => { ctx.notify(`${maintainer}。Mac 版问题反馈：${links.feedback}`); return { maintainer, feedback: links.feedback, upstream: links.developer }; } }
];
export const panels = [{ id: 'export.options', label: '导出与设置', section: '导出', mount: async (container, ctx) => {
  const lifetime = new AbortController(); const { signal } = lifetime; const options = await defaults(ctx);
  const root = element('section', undefined, { className: 'export-settings' }); root.append(element('h3', '导出页面与默认设置'));
  const range = select('导出范围', rangeOptions, options.range); const pages = input('自定义页码（1,3-5）');
  const format = select('格式', formatOptions, options.format); const dpi = input('DPI', 'number', String(options.dpi), { min: 36, max: 2400, step: 1 });
  const filename = input('文件名前缀（留空使用文稿名称）');
  const separate = input('每页导出为单独PDF', 'checkbox'); separate.control.checked = options.separatePdf;
  const open = input('导出后打开文件夹', 'checkbox'); open.control.checked = options.openFolder;
  const status = statusNode();
  const update = () => { pages.wrapper.hidden = range.control.value !== 'custom'; separate.wrapper.hidden = format.control.value !== 'pdf'; dpi.wrapper.hidden = format.control.value === 'pdf'; };
  range.control.addEventListener('change', update, { signal }); format.control.addEventListener('change', update, { signal }); update();
  root.append(range.wrapper, pages.wrapper, format.wrapper, dpi.wrapper, separate.wrapper, filename.wrapper, open.wrapper,
    button('选择文件夹并导出', async () => { const result = await exportSlides(ctx, { range: range.control.value, pages: pages.control.value, format: format.control.value, dpi: Number(dpi.control.value), filename: filename.control.value, separatePdf: separate.control.checked, openFolder: open.control.checked }); if (result) status.textContent = `已导出至${result.folder}`; }, ctx, signal),
    button('保存默认设置', async () => { const value = Number(dpi.control.value); if (!Number.isFinite(value) || value < 36 || value > 2400) throw new Error('DPI必须介于36与2400之间。'); for (const [key, setting] of Object.entries({ format: format.control.value, dpi: value, separatePdf: separate.control.checked, openFolder: open.control.checked })) await ctx.settings.set(`export.${key}`, setting); status.textContent = '默认设置已保存。'; }, ctx, signal), status
  );
  const about = await ctx.services.request('library/info', {}).catch(() => null);
  if (about?.version) root.append(element('p', `SlideSCI Mac ${about.version} · Windows原版 ${about.upstreamVersion}`));
  root.append(element('p', maintainer));
  for (const [name, url] of Object.entries(links)) root.append(button(linkLabels[name], () => ctx.services.openExternal(url), ctx, signal));
  container.append(root); return () => { lifetime.abort(); root.remove(); };
} }, { id: 'export.preferences', label: '偏好设置与帮助', section: '设置与帮助', mount: async (container, ctx) => {
  const lifetime = new AbortController(); const { signal } = lifetime;
  const root = element('section', undefined, { className: 'export-settings' });
  root.append(element('h3', '后台监听与运行方式'));
  const tracking = input('自动追踪从 Finder 复制粘贴/拖入的图片（剪贴板监听）', 'checkbox');
  tracking.control.checked = Boolean(await ctx.settings.get('background.autoPasteTracking', false));
  const hint = element('p', '默认关闭：避免后台监听干扰系统粘贴，也不产生后台轮询。需要记录原图路径或导入 16 位 TIFF 时，直接使用「图片 → 插入图片并记录来源 / 粘贴图片文件」即可。');
  const status = statusNode();
  tracking.control.addEventListener('change', async () => {
    const enabled = tracking.control.checked;
    try {
      await ctx.settings.set('background.autoPasteTracking', enabled);
      globalThis.dispatchEvent?.(new CustomEvent('slidesci:tracking-toggle', { detail: { enabled } }));
      status.textContent = enabled ? '已开启自动粘贴监听。' : '已关闭自动粘贴监听。';
    } catch (error) {
      tracking.control.checked = !enabled;
      status.textContent = '无法保存设置：' + error.message;
    }
  }, { signal });
  root.append(tracking.wrapper, hint, status);
  const about = await ctx.services.request('library/info', {}).catch(() => null);
  if (about?.version) root.append(element('p', `SlideSCI Mac ${about.version} · Windows原版 ${about.upstreamVersion}`));
  root.append(element('p', maintainer));
  for (const [name, url] of Object.entries(links)) root.append(button(linkLabels[name], () => ctx.services.openExternal(url), ctx, signal));
  container.append(root);
  return () => { lifetime.abort(); root.remove(); };
} }];
