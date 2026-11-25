import os from 'node:os';
import path from 'node:path';
import { promises as fs } from 'node:fs';
import { randomUUID } from 'node:crypto';
import { fileURLToPath } from 'node:url';

const mimeByExtension = {
  '.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.gif': 'image/gif',
  '.tif': 'image/tiff', '.tiff': 'image/tiff', '.bmp': 'image/bmp', '.svg': 'image/svg+xml',
  '.webp': 'image/webp', '.mp4': 'video/mp4', '.mov': 'video/quicktime', '.m4v': 'video/mp4',
  '.avi': 'video/x-msvideo', '.mp3': 'audio/mpeg', '.wav': 'audio/wav', '.pdf': 'application/pdf',
  '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  '.zip': 'application/zip', '.json': 'application/json'
};
const extensionByMime = Object.fromEntries(Object.entries(mimeByExtension).map(([extension, mime]) => [mime, extension]));

export function safeFilename(value, fallback = 'SlideSCI') {
  const name = String(value || fallback).replace(/[\\/\u0000-\u001f\u007f:]/g, '_').trim();
  if (!name || name === '.' || name === '..') return fallback;
  return name;
}
function contained(file, directory) {
  const relative = path.relative(directory, file);
  return relative === '' || (!relative.startsWith('..' + path.sep) && relative !== '..' && !path.isAbsolute(relative));
}
function filePath(value) {
  if (typeof value !== 'string' || !value.length || value.includes('\0')) throw new Error('文件路径无效。');
  const target = value.startsWith('file:') ? fileURLToPath(value) : value;
  if (!path.isAbsolute(target)) throw new Error('请选择本机上的原始文件。');
  return target;
}
export async function createFileServices(env) {
  const storageDir = path.join(os.homedir(), 'Library', 'Application Support', 'SlideSCI Mac');
  // The PID lets a later service tell an abandoned work folder from one still in use.
  const workDir = path.join(os.tmpdir(), `slidesci-mac-${process.pid}-${randomUUID()}`);
  const hostDirs = {
    powerpoint: path.join(os.homedir(), 'Library/Containers/com.microsoft.Powerpoint/Data/tmp/SlideSCI-Mac'),
    wps: path.join(os.homedir(), 'Library/Containers/com.kingsoft.wpsoffice.mac/Data/tmp/SlideSCI-Mac')
  };
  await fs.mkdir(storageDir, { recursive: true, mode: 0o700 });
  await fs.mkdir(workDir, { recursive: true, mode: 0o700 });
  const roots = [await fs.realpath(storageDir), await fs.realpath(workDir)];
  const grantedFiles = new Set(), grantedFolders = new Set();
  // The home folder may sit behind a symlink; host caches are matched in both spellings without
  // resolving anything inside another app's container.
  const home = os.homedir(), realHome = await fs.realpath(home).catch(() => home);
  async function hostDirectory(host) {
    if (!Object.hasOwn(hostDirs, host)) return workDir;
    await fs.mkdir(hostDirs[host], { recursive: true, mode: 0o700 });
    const resolved = await fs.realpath(hostDirs[host]);
    if (!roots.includes(resolved)) roots.push(resolved);
    return resolved;
  }
  async function allowFile(value) {
    const resolved = await fs.realpath(filePath(typeof value === 'string' ? value : value.path));
    if (!(await fs.stat(resolved)).isFile()) throw new Error('选择的项目不是文件。');
    grantedFiles.add(resolved);
    return { path: resolved, name: path.basename(resolved), mime: mimeByExtension[path.extname(resolved).toLowerCase()] || 'application/octet-stream' };
  }
  // A host cache is registered when a file inside it is first requested, never at startup: touching
  // another app's container raises macOS's "access data from other apps" prompt, which must not
  // appear when the login item starts. Only the container that holds the file is touched; a denial
  // is reported by the host operation that needs the cache.
  async function hostRoot(target) {
    for (const host of Object.keys(hostDirs)) {
      const real = path.join(realHome, path.relative(home, hostDirs[host]));
      if (!contained(target, hostDirs[host]) && !contained(target, real)) continue;
      try { await hostDirectory(host); }
      catch { /* the caller then reports the file as not granted */ }
    }
  }
  async function requireFile(value) {
    const spelled = path.resolve(filePath(value));
    const resolved = await fs.realpath(spelled);
    const allowed = () => grantedFiles.has(resolved) || roots.some(root => contained(resolved, root))
      || [...grantedFolders].some(root => contained(resolved, root));
    if (allowed()) return resolved;
    await hostRoot(spelled);
    if (resolved !== spelled) await hostRoot(resolved);
    if (allowed()) return resolved;
    throw new Error('请先通过文件选择器选择该文件。');
  }
  async function requireFolder(value) {
    const resolved = await fs.realpath(filePath(value));
    if (!(await fs.stat(resolved)).isDirectory()) throw new Error('目标不是文件夹。');
    if (roots.some(root => contained(resolved, root)) || [...grantedFolders].some(root => contained(resolved, root))) return resolved;
    throw new Error('请先选择输出文件夹。');
  }
  async function chooseFiles(options = {}) {
    const selected = await env.native.pickFiles(options);
    const entries = Array.isArray(selected) ? selected : selected?.files || [];
    return Promise.all(entries.map(allowFile));
  }
  async function chooseFolder(options = {}) {
    const selected = await env.native.chooseFolder(options);
    const value = typeof selected === 'string' ? selected : selected?.path;
    if (!value) return null;
    const resolved = await fs.realpath(filePath(value));
    if (!(await fs.stat(resolved)).isDirectory()) throw new Error('请选择文件夹。');
    grantedFolders.add(resolved);
    return resolved;
  }
  async function read(value) {
    const resolved = await requireFile(typeof value === 'string' ? value : value.path);
    const data = await fs.readFile(resolved);
    return { path: resolved, name: path.basename(resolved), data: data.toString('base64'),
      mime: mimeByExtension[path.extname(resolved).toLowerCase()] || 'application/octet-stream' };
  }
  async function materialize({ data, mime, name, host } = {}) {
    if (typeof data !== 'string') throw new Error('缺少文件数据。');
    if (host === 'wps' || host === 'powerpoint') {
      const result = await env.native.materialize({ data, mime, name, host });
      grantedFiles.add(await fs.realpath(result.path));
      return { ...result, name: name || result.name };
    }
    const root = await hostDirectory(host);
    const extension = extensionByMime[mime] || path.extname(safeFilename(name || '')) || '.bin';
    const target = path.join(root, randomUUID() + extension);
    await fs.writeFile(target, Buffer.from(data, 'base64'), { flag: 'wx', mode: 0o600 });
    grantedFiles.add(target);
    return { path: target, name: name || path.basename(target), mime: mime || mimeByExtension[extension] || 'application/octet-stream' };
  }
  async function write({ folder, name, data, mime } = {}) {
    const directory = await requireFolder(folder);
    const base = safeFilename(name || 'SlideSCI' + (extensionByMime[mime] || '.bin'));
    const extension = path.extname(base), stem = path.basename(base, extension);
    const bytes = Buffer.isBuffer(data) ? data : Buffer.from(data || '', 'base64');
    for (let index = 1; index <= 10000; index++) {
      const target = path.join(directory, index === 1 ? base : stem + '-' + index + extension);
      try {
        await fs.writeFile(target, bytes, { flag: 'wx', mode: 0o600 });
        grantedFiles.add(target);
        return { path: target, name: path.basename(target), mime: mime || mimeByExtension[extension.toLowerCase()] || 'application/octet-stream' };
      } catch (error) {
        if (error.code !== 'EEXIST') throw error;
      }
    }
    throw new Error('同名输出文件过多，请选择新的文件名。');
  }
  async function readClipboardFiles() {
    const selected = await env.native.readClipboardFiles({});
    const files = Array.isArray(selected) ? selected : selected?.files || [];
    return Promise.all(files.map(allowFile));
  }
  async function reveal(value) {
    const resolved = filePath(typeof value === 'string' ? value : value.path);
    await fs.stat(resolved).catch(() => { throw new Error('原始文件已不存在或已移动：' + resolved); });
    return env.native.reveal({ path: resolved });
  }
  const files = { read, materialize, write, chooseFiles, chooseFolder, reveal,
    pickFiles: chooseFiles, readClipboardFiles, writeClipboard: input => env.native.writeClipboard(input),
    grant: async values => Promise.all(values.map(allowFile)), requireFile, requireFolder, hostDirectory,
    clipboard: { readFiles: readClipboardFiles, write: input => env.native.writeClipboard(input) } };
  return { storageDir, workDir, files, hostDirs };
}
