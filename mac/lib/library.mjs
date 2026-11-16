import { mkdir, readFile, writeFile, rename, rm } from 'node:fs/promises';
import path from 'node:path';
import { randomUUID } from 'node:crypto';
import JSZip from 'jszip';

const locks = new Map();
const MAX_BYTES = 512 * 1024 * 1024;
const rootId = 'root';
function text(value, label) {
  const result = String(value ?? '').trim().normalize('NFC');
  if (!result || result.length > 160 || /[\x00-\x1f/\\]/u.test(result) || result === '.' || result === '..') throw new Error(`${label}不能为空，且不能包含路径分隔符或控制字符（最多160字）。`);
  return result;
}
function id(value) {
  if (typeof value !== 'string' || !/^(root|[a-f0-9-]{36})$/u.test(value)) throw new Error('无效的素材或文件夹标识。');
  return value;
}
function base64(value, label) {
  if (typeof value !== 'string' || !value.length || value.length > MAX_BYTES * 1.4 || !/^[A-Za-z0-9+/\r\n]*={0,2}$/u.test(value)) throw new Error(`${label}不是有效的Base64文件。`);
  const bytes = Buffer.from(value, 'base64');
  if (!bytes.length || bytes.length > MAX_BYTES) throw new Error(`${label}大小超过512MB限制。`);
  return bytes;
}
async function validatePptx(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  if (!zip.file('[Content_Types].xml') || !zip.file('ppt/presentation.xml') || !Object.keys(zip.files).some(name => /^ppt\/slides\/slide\d+\.xml$/u.test(name))) throw new Error('素材必须包含有效、可编辑的PowerPoint演示文稿。');
}
async function atomic(file, contents) {
  const temp = `${file}.${randomUUID()}.tmp`;
  try { await writeFile(temp, contents, { mode: 0o600 }); await rename(temp, file); }
  finally { await rm(temp, { force: true }); }
}
async function store(env, callback) {
  const root = path.join(env.storageDir, 'library');
  const previous = locks.get(root) || Promise.resolve();
  const operation = previous.catch(() => {}).then(async () => {
    await mkdir(path.join(root, 'assets'), { recursive: true, mode: 0o700 });
    let db;
    try { db = JSON.parse(await readFile(path.join(root, 'index.json'), 'utf8')); }
    catch (error) { if (error.code !== 'ENOENT') throw error; db = { version: 1, folders: [{ id: rootId, parent: null, name: '素材库' }], assets: [] }; }
    if (db.version !== 1 || !Array.isArray(db.folders) || !Array.isArray(db.assets)) throw new Error('素材库索引格式无效，请先恢复备份。');
    const save = () => atomic(path.join(root, 'index.json'), JSON.stringify(db));
    return callback({ root, db, save });
  });
  locks.set(root, operation);
  try { return await operation; } finally { if (locks.get(root) === operation) locks.delete(root); }
}
function folder(db, value = rootId) {
  const found = db.folders.find(item => item.id === id(value));
  if (!found) throw new Error('文件夹已不存在，请刷新素材库。');
  return found;
}
function target(db, value) {
  id(value);
  const found = db.assets.find(item => item.id === value) || db.folders.find(item => item.id === value);
  if (!found || value === rootId) throw new Error('素材不存在或操作指向了素材库根目录。');
  return found;
}
function unique(db, parent, name, except, kind = 'asset') {
  const items = kind === 'folder' ? db.folders.filter(item => item.parent === parent) : db.assets.filter(item => item.folder === parent);
  if (items.some(item => item.id !== except && item.name.toLocaleLowerCase() === name.toLocaleLowerCase())) throw new Error(`当前文件夹已经存在「${name}」。`);
}
function folderPaths(db) {
  const byId = new Map(db.folders.map(item => [item.id, item]));
  const paths = new Map([[rootId, '']]);
  return value => {
    if (paths.has(value)) return paths.get(value);
    const chain = []; const visited = new Set();
    let cursor = value, base = '';
    while (cursor && cursor !== rootId) {
      if (paths.has(cursor)) { base = paths.get(cursor); break; }
      if (visited.has(cursor)) throw new Error('素材库文件夹存在循环引用。');
      visited.add(cursor);
      const key = id(cursor);
      const item = byId.get(key);
      if (!item) throw new Error('文件夹已不存在，请刷新素材库。');
      chain.push({ id: key, name: item.name });
      cursor = item.parent;
    }
    let prefix = base;
    for (const item of chain.reverse()) {
      prefix = prefix ? `${prefix}/${item.name}` : item.name;
      paths.set(item.id, prefix);
    }
    return prefix;
  };
}
function metadata(fragment) {
  const result = {};
  if (Array.isArray(fragment.shapeIds)) result.shapeIds = fragment.shapeIds.map(String);
  if (fragment.bounds && ['left', 'top', 'width', 'height'].every(key => Number.isFinite(fragment.bounds[key]))) result.bounds = fragment.bounds;
  if (fragment.name) result.name = String(fragment.name);
  return result;
}
async function mapLimit(items, limit, callback) {
  const results = new Array(items.length);
  let index = 0;
  const worker = async () => {
    while (index < items.length) { const current = index++; results[current] = await callback(items[current], current); }
  };
  await Promise.all(Array.from({ length: Math.min(limit, items.length) }, worker));
  return results;
}
async function readFragmentFiles(root, asset) {
  const directory = path.join(root, 'assets', id(asset.id));
  const [meta, pptx, preview] = await Promise.all([
    readFile(path.join(directory, 'fragment.json'), 'utf8'),
    readFile(path.join(directory, 'fragment.pptx')),
    readFile(path.join(directory, 'preview.png'))
  ]);
  return { meta: JSON.parse(meta), pptx, preview };
}
async function readFragment(root, asset) {
  const files = await readFragmentFiles(root, asset);
  return { ...files.meta, pptx: files.pptx.toString('base64'), preview: files.preview.toString('base64') };
}
async function putFragment(root, asset, fragment) {
  const pptx = Buffer.isBuffer(fragment?.pptx) ? fragment.pptx : base64(fragment?.pptx, '可编辑素材');
  const preview = Buffer.isBuffer(fragment?.preview) ? fragment.preview : base64(fragment?.preview, 'PNG预览');
  if (!pptx.length || pptx.length > MAX_BYTES) throw new Error('可编辑素材大小超过512MB限制。');
  if (!preview.subarray(0, 8).equals(Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]))) throw new Error('素材预览必须为PNG图片。');
  await validatePptx(pptx);
  // A generation directory makes updates atomic, including their preview and metadata.
  const generation = randomUUID();
  const directory = path.join(root, 'assets', generation);
  await mkdir(directory, { mode: 0o700 });
  try {
    await writeFile(path.join(directory, 'fragment.pptx'), pptx, { mode: 0o600 });
    await writeFile(path.join(directory, 'preview.png'), preview, { mode: 0o600 });
    await writeFile(path.join(directory, 'fragment.json'), JSON.stringify(metadata(fragment)), { mode: 0o600 });
  } catch (error) { await rm(directory, { recursive: true, force: true }); throw error; }
  return { ...asset, id: generation, updatedAt: new Date().toISOString() };
}
async function listing({ root, db }, input) {
  const current = folder(db, input.folder || rootId);
  const query = String(input.query || '').trim().toLocaleLowerCase();
  const pathOf = folderPaths(db);
  const assets = query ? db.assets.filter(asset => asset.name.toLocaleLowerCase().includes(query) || pathOf(asset.folder).toLocaleLowerCase().includes(query)) : db.assets.filter(asset => asset.folder === current.id);
  const entries = await mapLimit(assets, 16, async asset => ({ ...asset, kind: 'asset', path: pathOf(asset.folder), preview: (await readFile(path.join(root, 'assets', id(asset.id), 'preview.png'))).toString('base64') }));
  const folders = query ? [] : db.folders.filter(item => item.parent === current.id).map(item => ({ ...item, kind: 'folder' }));
  return { folder: current, folders: db.folders.map(item => ({ ...item, path: pathOf(item.id) })), items: [...folders, ...entries].sort((a, b) => (a.kind === b.kind ? a.name.localeCompare(b.name, 'zh') : a.kind === 'folder' ? -1 : 1)) };
}
function archivePath(value) {
  if (typeof value !== 'string' || value.startsWith('/') || value.includes('\\') || /[\x00-\x1f:]/u.test(value)) throw new Error('ZIP包含不安全的文件路径。');
  const parts = value.split('/').filter(Boolean);
  if (parts.some(part => part === '.' || part === '..')) throw new Error('ZIP包含越界文件路径。');
  return parts;
}
export const libraryHandlers = {
  info: async () => {
    const pkg = JSON.parse(await readFile(new URL('../package.json', import.meta.url), 'utf8'));
    return { name: 'SlideSCI Mac', version: pkg.version, upstreamVersion: '2.2.4.0', maintainer: 'nwh-phy', feedback: 'https://github.com/nwh-phy/SlideSCI-Mac/issues', developer: 'Achuan-2' };
  },
  list: (input, env) => store(env, state => listing(state, input)),
  get: (input, env) => store(env, async ({ root, db }) => {
    const asset = db.assets.find(item => item.id === id(input.id));
    if (!asset) throw new Error('素材不存在。');
    return { asset, fragment: await readFragment(root, asset) };
  }),
  save: (input, env) => store(env, async ({ root, db, save }) => {
    const parent = folder(db, input.folder || rootId).id; const name = text(input.name, '素材名称');
    const existing = input.id ? db.assets.find(item => item.id === id(input.id)) : null;
    if (input.id && !existing) throw new Error('待更新素材不存在。');
    const actualParent = existing ? existing.folder : parent;
    unique(db, actualParent, name, existing?.id);
    const asset = await putFragment(root, { folder: actualParent, name, createdAt: existing?.createdAt || new Date().toISOString() }, input.fragment);
    if (existing) db.assets.splice(db.assets.indexOf(existing), 1, asset); else db.assets.push(asset);
    await save();
    if (existing) await rm(path.join(root, 'assets', existing.id), { recursive: true, force: true });
    return { asset };
  }),
  mkdir: (input, env) => store(env, async ({ db, save }) => {
    const parent = folder(db, input.folder || rootId).id; const name = text(input.name, '文件夹名称'); unique(db, parent, name, null, 'folder');
    const item = { id: randomUUID(), parent, name }; db.folders.push(item); await save(); return item;
  }),
  rename: (input, env) => store(env, async ({ db, save }) => {
    const item = target(db, input.id); const name = text(input.name, '名称'); unique(db, item.folder || item.parent, name, item.id, 'folder' in item ? 'asset' : 'folder');
    item.name = name; await save(); return item;
  }),
  move: (input, env) => store(env, async ({ db, save }) => {
    const item = target(db, input.id); const destination = folder(db, input.folder).id;
    if ('parent' in item) {
      let cursor = destination;
      while (cursor) { if (cursor === item.id) throw new Error('不能将文件夹移入其自身或子文件夹。'); cursor = folder(db, cursor).parent; }
    }
    unique(db, destination, item.name, item.id, 'folder' in item ? 'asset' : 'folder');
    if ('folder' in item) item.folder = destination; else item.parent = destination;
    await save(); return item;
  }),
  delete: (input, env) => store(env, async ({ root, db, save }) => {
    const item = target(db, input.id); const removed = new Set([item.id]);
    const childrenByParent = new Map();
    for (const child of db.folders) {
      const list = childrenByParent.get(child.parent);
      if (list) list.push(child.id); else childrenByParent.set(child.parent, [child.id]);
    }
    const queue = [item.id];
    while (queue.length) {
      for (const id of childrenByParent.get(queue.pop()) || []) {
        if (removed.has(id)) continue;
        removed.add(id); queue.push(id);
      }
    }
    const assets = db.assets.filter(asset => removed.has(asset.id) || removed.has(asset.folder));
    db.assets = db.assets.filter(asset => !assets.includes(asset)); db.folders = db.folders.filter(child => !removed.has(child.id)); await save();
    await Promise.all(assets.map(asset => rm(path.join(root, 'assets', asset.id), { recursive: true, force: true })));
    return { deleted: assets.length, folders: removed.size - ('folder' in item ? 1 : 0) };
  }),
  reveal: (input, env) => store(env, async ({ root, db }) => {
    const asset = input.id ? db.assets.find(item => item.id === id(input.id)) : null;
    if (input.id && !asset) throw new Error('素材不存在。');
    const location = asset ? path.join(root, 'assets', asset.id, 'fragment.pptx') : root;
    await env.files.reveal(location); return { path: location };
  }),
  export: (input, env) => store(env, async ({ root, db }) => {
    const zip = new JSZip(); const manifest = { version: 1, assets: {} };
    const pathOf = folderPaths(db);
    for (const item of db.folders) if (item.id !== rootId) zip.folder(pathOf(item.id));
    const fragments = await mapLimit(db.assets, 8, asset => readFragmentFiles(root, asset));
    db.assets.forEach((asset, index) => {
      const prefix = [pathOf(asset.folder), asset.name].filter(Boolean).join('/');
      const fragment = fragments[index];
      zip.file(`${prefix}.pptx`, fragment.pptx);
      zip.file(`${prefix}.png`, fragment.preview);
      manifest.assets[`${prefix}.pptx`] = metadata(fragment.meta);
    });
    zip.file('.slidesci.json', JSON.stringify(manifest));
    return { name: 'SlideSCI_ShapeLibrary.zip', mime: 'application/zip', data: await zip.generateAsync({ type: 'base64', compression: 'DEFLATE' }), count: db.assets.length };
  }),
  import: (input, env) => store(env, async ({ root, db, save }) => {
    if (!['skip', 'overwrite', 'rename'].includes(input.conflict || 'skip')) throw new Error('请选择重名素材的处理方式。');
    const source = input.path ? (await env.files.read(input.path)).data : input.data;
    const zip = await JSZip.loadAsync(base64(source, '素材库ZIP'));
    const entries = Object.values(zip.files);
    if (entries.length > 20000) throw new Error('ZIP文件数超过20000。');
    let total = 0;
    for (const entry of entries) {
      archivePath(entry.unsafeOriginalName || entry.name);
      total += Number(entry._data?.uncompressedSize || 0);
      if (total > MAX_BYTES) throw new Error('ZIP解压后超过512MB。');
    }
    let manifest = {};
    if (zip.file('.slidesci.json')) manifest = JSON.parse(await zip.file('.slidesci.json').async('string'));
    const destination = folder(db, input.folder || rootId).id;
    const ensureFolder = parts => {
      let parent = destination;
      for (const part of parts) {
        const name = text(part, '文件夹名称');
        let found = db.folders.find(item => item.parent === parent && item.name.toLocaleLowerCase() === name.toLocaleLowerCase());
        if (!found) { unique(db, parent, name, null, 'folder'); found = { id: randomUUID(), parent, name }; db.folders.push(found); }
        parent = found.id;
      }
      return parent;
    };
    const staged = []; const oldAssets = []; let skipped = 0;
    try {
      for (const entry of entries.filter(item => item.dir)) ensureFolder(archivePath(entry.name));
      for (const entry of entries.filter(item => !item.dir && /\.pptx$/iu.test(item.name))) {
        const parts = archivePath(entry.name); let name = text(parts.pop().replace(/\.pptx$/iu, ''), '素材名称'); const parent = ensureFolder(parts);
        const previous = db.assets.find(item => item.folder === parent && item.name.toLocaleLowerCase() === name.toLocaleLowerCase());
        if (previous && (input.conflict || 'skip') === 'skip') { skipped++; continue; }
        if (previous && input.conflict === 'rename') { const base = name; for (let n = 2; db.assets.some(item => item.folder === parent && item.name.toLocaleLowerCase() === name.toLocaleLowerCase()); n++) name = `${base} (${n})`; }
        else unique(db, parent, name, previous?.id);
        const png = zip.file(entry.name.replace(/\.pptx$/iu, '.png'));
        if (!png) throw new Error(`素材「${entry.name}」缺少同名PNG预览。`);
        const fragment = { ...manifest.assets?.[entry.name], pptx: await entry.async('nodebuffer'), preview: await png.async('nodebuffer') };
        const asset = await putFragment(root, { name, folder: parent, createdAt: new Date().toISOString() }, fragment); staged.push(asset);
        if (previous && input.conflict === 'overwrite') { db.assets.splice(db.assets.indexOf(previous), 1); oldAssets.push(previous); }
        db.assets.push(asset);
      }
      if (!staged.length && !skipped && !entries.some(item => item.dir)) throw new Error('ZIP中没有素材（需要配对的.pptx和.png文件）。');
      await save();
    } catch (error) { await Promise.all(staged.map(asset => rm(path.join(root, 'assets', asset.id), { recursive: true, force: true }))); throw error; }
    await Promise.all(oldAssets.map(asset => rm(path.join(root, 'assets', asset.id), { recursive: true, force: true })));
    return { imported: staged.length, skipped };
  })
};
