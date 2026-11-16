import { element, button, input, select, statusNode, descriptorData, filesOf } from './library-ai-ui.js';

async function selectionFragment(ctx) {
  const { selectionIds } = await ctx.host.summary({ slides: false });
  if (!selectionIds.length) throw new Error('请先选择需要保存的形状或组合。');
  const fragment = await ctx.host.fragment(selectionIds);
  if (!fragment?.pptx) throw new Error('宿主未返回可编辑的PPTX素材。');
  if (!fragment.preview) {
    const result = await ctx.host.export({ kind: 'selection', ids: selectionIds, format: 'png', dpi: 150 });
    fragment.preview = (await descriptorData(filesOf(result)[0], ctx.services)).data;
  }
  return fragment;
}
export const actions = [
  { id: 'library.save', label: '将选中形状保存至素材库', section: '素材库', fields: [{ name: 'name', label: '素材名称', type: 'text', required: true, default: '自定义素材' }], run: async (ctx, values) => {
    const result = await ctx.services.request('library/save', { name: values.name, folder: 'root', fragment: await selectionFragment(ctx) });
    ctx.notify(`已保存「${result.asset.name}」，包含原生可编辑对象。`); return result;
  } },
  { id: 'library.export', label: '导出完整素材库ZIP', section: '素材库', run: async ctx => {
    const archive = await ctx.services.request('library/export', {}); await ctx.services.download(archive); ctx.notify(`已导出${archive.count}个素材。`); return archive;
  } },
  { id: 'library.import', label: '导入素材库ZIP', section: '素材库', fields: [{ name: 'conflict', label: '重名素材', type: 'select', default: 'skip', options: [{ value: 'skip', label: '跳过' }, { value: 'overwrite', label: '覆盖' }, { value: 'rename', label: '保留两份' }] }], run: async (ctx, values) => {
    const files = await ctx.services.pickFiles({ kind: 'all', multiple: false }); if (!files?.length) return;
    const result = await ctx.services.request('library/import', { path: files[0].path, data: files[0].data, conflict: values.conflict });
    ctx.notify(`已导入${result.imported}个素材，跳过${result.skipped}个重名素材。`); return result;
  } }
];
export const panels = [{ id: 'library.browser', label: '素材库', section: '素材库', mount: async (container, ctx) => {
  const lifetime = new AbortController(); const { signal } = lifetime;
  let current = 'root'; let selected = null; let listing; let requestVersion = 0;
  let view = await ctx.settings.get('library.view', 'cards');
  const root = element('section', undefined, { className: 'library-browser' });
  const heading = element('h3', '可编辑素材库'); const status = statusNode();
  const toolbar = element('div', undefined, { className: 'feature-toolbar' });
  const breadcrumbs = element('nav', undefined, { 'aria-label': '素材库路径' });
  const search = input('搜索全部素材', 'search', '', { placeholder: '素材名称或文件夹名称' });
  const names = input('素材 / 新文件夹名称', 'text', '', { placeholder: '输入名称' });
  const results = element('div', undefined, { className: 'library-results', role: 'list' });
  const details = element('section', undefined, { 'aria-label': '素材详细信息' });
  const destination = select('移动 / 保存到', [{ value: 'root', label: '素材库' }], current);
  const destructive = input('我确认删除所选素材或文件夹及其内容', 'checkbox');
  const controls = element('div', undefined, { className: 'feature-toolbar' });
  const request = (action, payload = {}) => ctx.services.request(`library/${action}`, payload);
  const announce = message => { status.textContent = message; ctx.notify(message); };
  const requireSelection = () => { if (!selected) throw new Error('请先选择素材或文件夹。'); return selected; };
  async function insert(item) {
    if (item.kind === 'folder') { current = item.id; search.control.value = ''; selected = null; await load(); return; }
    const { fragment } = await request('get', { id: item.id });
    let options = { center: true };
    if (fragment.bounds) {
      const { width, height } = await ctx.host.summary({ slides: false });
      options = { left: (width - fragment.bounds.width) / 2, top: (height - fragment.bounds.height) / 2 };
    }
    await ctx.host.insertFragment(fragment, options); await ctx.refresh(); announce(`已插入可编辑素材「${item.name}」。`);
  }
  function preview(item) {
    selected = item; details.replaceChildren();
    for (const node of results.children) node.setAttribute('aria-selected', String(node.dataset.id === item.id));
    details.append(element('h4', item.name));
    if (item.kind === 'asset') {
      const image = element('img', undefined, { src: `data:image/png;base64,${item.preview}`, alt: `${item.name}预览`, className: 'library-preview' });
      details.append(image, element('p', item.path || '素材库根目录'), button('插入当前幻灯片', () => insert(item), ctx, signal));
    } else details.append(element('p', '文件夹'), button('进入文件夹', () => insert(item), ctx, signal));
    names.control.value = item.name; destructive.control.checked = false;
  }
  async function load() {
    const version = ++requestVersion;
    const data = await request('list', { folder: current, query: search.control.value });
    if (version !== requestVersion || signal.aborted) return;
    listing = data; results.replaceChildren(); breadcrumbs.replaceChildren(); details.replaceChildren(); selected = null;
    const chain = []; let cursor = data.folder;
    while (cursor) { chain.unshift(cursor); cursor = data.folders.find(item => item.id === cursor.parent); }
    for (const item of chain) breadcrumbs.append(button(item.name, async () => { current = item.id; search.control.value = ''; await load(); }, ctx, signal));
    const chosenDestination = destination.control.value;
    destination.control.replaceChildren(...data.folders.map(item => element('option', item.path || '素材库', { value: item.id })));
    destination.control.value = data.folders.some(item => item.id === chosenDestination) ? chosenDestination : current;
    results.className = `library-results library-${view}`;
    for (const item of data.items) {
      const card = element('article', undefined, { className: 'library-card', role: 'listitem', tabIndex: 0, draggable: true, 'aria-label': item.kind === 'folder' ? `文件夹 ${item.name}` : item.name, 'aria-selected': 'false' });
      card.dataset.id = item.id;
      if (item.kind === 'asset') card.append(element('img', undefined, { src: `data:image/png;base64,${item.preview}`, alt: '', loading: 'lazy', className: 'library-thumbnail' }));
      card.append(element('strong', item.kind === 'folder' ? `文件夹：${item.name}` : item.name));
      if (item.path && search.control.value) card.append(element('small', item.path));
      card.addEventListener('click', () => preview(item), { signal });
      card.addEventListener('dblclick', () => insert(item).catch(error => ctx.notify(error.message, 'error')), { signal });
      card.addEventListener('keydown', event => { if (event.key === 'Enter') { event.preventDefault(); insert(item).catch(error => ctx.notify(error.message, 'error')); } else if (event.key === ' ') { event.preventDefault(); preview(item); } }, { signal });
      card.addEventListener('dragstart', event => event.dataTransfer.setData('application/x-slidesci-asset', item.id), { signal });
      if (item.kind === 'folder') {
        card.addEventListener('dragover', event => { if (event.dataTransfer.types.includes('application/x-slidesci-asset')) event.preventDefault(); }, { signal });
        card.addEventListener('drop', async event => { event.preventDefault(); const id = event.dataTransfer.getData('application/x-slidesci-asset'); if (!id) return; try { await request('move', { id, folder: item.id }); await load(); } catch (error) { ctx.notify(error.message, 'error'); } }, { signal });
      }
      results.append(card);
    }
    status.textContent = `${data.items.length}项${search.control.value ? '搜索结果' : ''}；双击或按Enter插入 / 打开，空格键选择。`;
  }
  toolbar.append(
    button('返回上级', async () => { current = listing.folder.parent || 'root'; search.control.value = ''; await load(); }, ctx, signal),
    button('刷新', load, ctx, signal),
    button('卡片 / 列表', async () => { view = view === 'cards' ? 'list' : 'cards'; await ctx.settings.set('library.view', view); results.className = `library-results library-${view}`; }, ctx, signal),
    button('打开素材库文件夹', () => request('reveal'), ctx, signal)
  );
  controls.append(
    button('保存选中形状', async () => { const result = await request('save', { name: names.control.value, folder: destination.control.value, fragment: await selectionFragment(ctx) }); current = result.asset.folder; await load(); announce(`已保存「${result.asset.name}」。`); }, ctx, signal),
    button('新建文件夹', async () => { await request('mkdir', { name: names.control.value, folder: current }); await load(); }, ctx, signal),
    button('重命名', async () => { await request('rename', { id: requireSelection().id, name: names.control.value }); await load(); }, ctx, signal),
    button('移动', async () => { await request('move', { id: requireSelection().id, folder: destination.control.value }); await load(); }, ctx, signal),
    button('用选中形状更新素材', async () => { const item = requireSelection(); if (item.kind !== 'asset') throw new Error('请选择需要更新的素材。'); await request('save', { id: item.id, name: item.name, fragment: await selectionFragment(ctx) }); await load(); announce(`已更新「${item.name}」。`); }, ctx, signal),
    button('显示素材源文件', async () => { const item = requireSelection(); if (item.kind !== 'asset') throw new Error('请选择素材。'); await request('reveal', { id: item.id }); }, ctx, signal),
    button('删除', async () => { const item = requireSelection(); if (!destructive.control.checked) throw new Error('请先勾选删除确认。'); await request('delete', { id: item.id }); destructive.control.checked = false; await load(); }, ctx, signal)
  );
  const conflict = select('ZIP导入重名处理', [{ value: 'skip', label: '跳过' }, { value: 'overwrite', label: '覆盖' }, { value: 'rename', label: '保留两份' }], 'skip');
  const archives = element('div', undefined, { className: 'feature-toolbar' });
  archives.append(conflict.wrapper,
    button('导入ZIP到当前文件夹', async () => { const files = await ctx.services.pickFiles({ kind: 'all', multiple: false }); if (!files?.length) return; const result = await request('import', { path: files[0].path, data: files[0].data, folder: current, conflict: conflict.control.value }); await load(); announce(`已导入${result.imported}个素材，跳过${result.skipped}个。`); }, ctx, signal),
    button('导出完整素材库ZIP', async () => { const archive = await request('export'); await ctx.services.download(archive); announce(`已导出${archive.count}个素材。`); }, ctx, signal)
  );
  let timer;
  search.control.addEventListener('input', () => { clearTimeout(timer); timer = setTimeout(() => load().catch(error => ctx.notify(error.message, 'error')), 180); }, { signal });
  root.append(heading, toolbar, breadcrumbs, search.wrapper, status, results, details, names.wrapper, destination.wrapper, controls, destructive.wrapper, archives); container.append(root);
  await load();
  return () => { lifetime.abort(); clearTimeout(timer); root.remove(); };
} }];
