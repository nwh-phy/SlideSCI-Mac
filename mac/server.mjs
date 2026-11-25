import http from 'node:http';
import https from 'node:https';
import os from 'node:os';
import path from 'node:path';
import { lookup } from 'node:dns/promises';
import { lstat, mkdir, readFile, readdir, realpath, rm, stat, writeFile } from 'node:fs/promises';
import { copyFileSync, fstatSync, ftruncateSync, statSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { parseArgs } from 'node:util';
import { randomBytes, randomUUID, timingSafeEqual } from 'node:crypto';
import { EventEmitter } from 'node:events';
import { contentHandlers } from './lib/content.mjs';
import { imageHandlers } from './lib/images.mjs';
import { libraryHandlers } from './lib/library.mjs';
import { nativeHandlers } from './lib/native.mjs';
import { pptx } from './lib/pptx.mjs';
import { createFileServices } from './lib/storage.mjs';
import { acceptanceHandlers } from './lib/acceptance.mjs';

const { values } = parseArgs({ options: { cert: { type: 'string' }, key: { type: 'string' } } });
if (Boolean(values.cert) !== Boolean(values.key)) throw new Error('--cert 和 --key 必须一起提供。');
const root = fileURLToPath(new URL('.', import.meta.url));
const version = JSON.parse(await readFile(new URL('./package.json', import.meta.url), 'utf8')).version;
const token = randomBytes(32).toString('hex');
// Identifies this server process, so a pane can tell a restarted service from a dropped connection.
const instance = randomUUID();
// The add-in manifests fix 3000/3443; the overrides exist only for isolated integration tests.
const httpPort = Number(process.env.SLIDESCI_HTTP_PORT) || 3000;
const httpsPort = Number(process.env.SLIDESCI_HTTPS_PORT) || 3443;
const origins = new Set([`http://127.0.0.1:${httpPort}`, `http://localhost:${httpPort}`, `https://localhost:${httpsPort}`, `https://127.0.0.1:${httpsPort}`, `https://[::1]:${httpsPort}`]);
const events = new EventEmitter();
events.setMaxListeners(100);
const env = { events, pptx, version };
env.native = Object.fromEntries(Object.entries(nativeHandlers).map(([name, fn]) => [name, input => fn(input || {}, env)]));
Object.assign(env, await createFileServices(env));
// Stay completely passive at startup so a login LaunchAgent never triggers macOS app-data or
// automation prompts until the user invokes an advanced native action.
const accessWarnings = [];
// service.log is what users attach to bug reports. Each event is one timestamped line, the home
// folder is written as ~ so a shared log does not carry the account name, and an error that keeps
// repeating is counted rather than written again for ten minutes, so a loop cannot push the useful
// lines out of the 5 MB log.
const home = os.homedir();
const REPEAT_WINDOW = 10 * 60 * 1000;
let lastLogged = null;
function flushRepeats() {
  if (lastLogged?.repeats) console.log(`${new Date().toISOString()} （上一条又出现 ${lastLogged.repeats} 次）`);
  if (lastLogged) lastLogged.repeats = 0;
}
function log(message, { error = false } = {}) {
  const text = String(message).replaceAll(home, '~').replace(/\s*[\r\n]+\s*/g, ' ⏎ ').slice(0, 4000), now = Date.now();
  if (lastLogged?.text === text && now - lastLogged.at < REPEAT_WINDOW) { lastLogged.repeats++; return; }
  flushRepeats();
  lastLogged = { text, at: now, repeats: 0 };
  (error ? console.error : console.log)(`${new Date(now).toISOString()} ${text}`);
}
// The pane lists the last 30 service errors; each one is also written to service.log, in the
// fuller form given as `logged` when there is one.
const diagnostics = [];
function diagnose(message, logged = message) {
  diagnostics.push({ at: new Date().toISOString(), message: String(message || '服务错误').slice(0, 4000) });
  if (diagnostics.length > 30) diagnostics.shift();
  log(logged || '服务错误', { error: true });
}
// A long-lived local service must log, not exit, when one request leaves a promise unhandled.
process.on('unhandledRejection', error => diagnose(error?.message || error, 'SlideSCI 未处理的异步错误：' + (error?.stack || error)));
// launchd opens the log once at login in append mode and the service may run for weeks, so the
// service rotates its own log: copy it aside, then truncate in place (appends continue at the end).
const LOG_LIMIT = 5 * 1024 * 1024;
function rotateLog(file = process.env.SLIDESCI_LOG_FILE, fd = 1, limit = LOG_LIMIT) {
  if (!file) return false;
  try {
    const output = fstatSync(fd), current = statSync(file);
    if (!output.isFile() || output.ino !== current.ino || output.dev !== current.dev || output.size <= limit) return false;
    copyFileSync(file, file + '.1');
    ftruncateSync(fd, 0);
    return true;
  } catch { return false; }
}
rotateLog();
setInterval(rotateLog, 3600 * 1000).unref();
const clients = new Set();
const watchLeases = new Map();
const LEASE_GRACE = 60000;
const dropTimers = new Map();
let watchState = null;
let watchQueue = Promise.resolve();
function watchOperation(operation) {
  const result = watchQueue.then(operation, operation);
  watchQueue = result.catch(() => {});
  return result;
}
const clientId = value => typeof value === 'string' && /^[A-Za-z0-9_-]{8,64}$/.test(value) ? value : null;
const connected = client => [...clients].some(item => item.client === client);
const leasesOf = client => [...watchLeases].filter(([, lease]) => lease.client === client).map(([id]) => id);
async function releaseWatcherIfUnused() {
  if (!watchLeases.size && watchState) {
    const current = watchState;
    watchState = null;
    await nativeHandlers.watchStop({ id: current.id }, env);
  }
}
// Leases follow the pane's event stream: a pane that disappears without pagehide (quit, crash,
// sleep) loses its lease after a grace period, so the clipboard watcher never outlives every pane.
function scheduleLeaseDrop(client) {
  if (!client || dropTimers.has(client) || connected(client)) return;
  dropTimers.set(client, setTimeout(() => {
    dropTimers.delete(client);
    if (connected(client)) return;
    watchOperation(async () => {
      for (const id of leasesOf(client)) watchLeases.delete(id);
      await releaseWatcherIfUnused();
    }).catch(error => diagnose('释放媒体监听失败：' + error.message));
  }, LEASE_GRACE));
}
async function startWatch(input) {
  return watchOperation(async () => {
    const client = clientId(input.client);
    if (!watchState) watchState = await nativeHandlers.watchStart({ interceptTIFF: input.interceptTIFF === true }, env);
    // One lease per pane: a pane that re-acquires after reconnecting keeps its existing lease.
    const existing = client ? leasesOf(client)[0] : null;
    const id = existing || randomUUID();
    if (!existing) watchLeases.set(id, { host: input.host, client, created: Date.now() });
    scheduleLeaseDrop(client);
    return { ...watchState, id, instance, shared: true };
  });
}
async function stopWatch(input) {
  return watchOperation(async () => {
    if (input.id) watchLeases.delete(input.id);
    await releaseWatcherIfUnused();
    return { stopped: true };
  });
}
const nativePublic = Object.fromEntries(['execute', 'pickFiles', 'chooseFolder', 'reveal', 'fonts',
  'readClipboardFiles', 'writeClipboard', 'normalizeTIFF'].map(name => [name, nativeHandlers[name]]));
nativePublic.watchStart = startWatch;
nativePublic.watchStop = stopWatch;
const fileHandlers = {
  read: input => env.files.read(input),
  materialize: input => env.files.materialize(input),
  write: input => env.files.write(input),
  pickFiles: input => env.files.chooseFiles(input),
  chooseFolder: input => env.files.chooseFolder(input),
  reveal: input => env.files.reveal(input),
  readClipboardFiles: () => env.files.readClipboardFiles(),
  writeClipboard: input => env.files.writeClipboard(input),
  download: async input => {
    const folder = input.folder || await env.files.chooseFolder({ title: '保存 ' + (input.name || 'SlideSCI 文件') });
    if (!folder) return { cancelled: true };
    if (typeof input.data !== 'string') throw new Error('缺少要保存的文件数据。');
    return env.files.write({ ...input, folder });
  },
  openExternal: async input => {
    const url = new URL(input.url);
    if (!['http:', 'https:', 'mailto:'].includes(url.protocol)) throw new Error('不允许打开此链接类型。');
    const { execFile } = await import('node:child_process');
    await new Promise((resolve, reject) => execFile('/usr/bin/open', [url.href], error => error ? reject(error) : resolve()));
    return { opened: true };
  }
};
const handlers = {
  content: contentHandlers, images: imageHandlers, library: libraryHandlers,
  native: nativePublic, files: fileHandlers,
  pptx: {
    parse: input => pptx.parsePresentation(input.pptx, input),
    images: input => pptx.extractImages(input.pptx, input),
    fragment: input => pptx.extractFragment(input.pptx, input),
    selectSlides: async input => ({ pptx: await pptx.selectSlides(input.pptx, input.indices) }),
    imageFragment: input => pptx.imageFragment(input.pptx, input),
    formatText: input => pptx.formatText(input.pptx, input),
    replacePicture: async input => pptx.replacePicture(input.pptx, {
      ...input, data: input.data || (await env.files.read(input.path)).data
    })
  },
  client: {
    error: input => {
      const message = String(input.message || '客户端错误');
      // Which pane and action failed, plus the first frames of its stack, without the payload data.
      const context = input.context && typeof input.context === 'object' ? input.context : {};
      const where = [input.host, input.version, context.operation || context.action || context.panel]
        .filter(value => typeof value === 'string' && value).map(value => value.slice(0, 80)).join(' ');
      const frames = typeof input.stack === 'string' ? input.stack.split('\n').map(line => line.trim()).filter(line => line.startsWith('at ')).slice(0, 3) : [];
      diagnose(message, `界面错误${where ? `（${where}）` : ''}：${message.slice(0, 1000)}${frames.length ? ' | ' + frames.join(' | ') : ''}`);
      return { recorded: true };
    },
    diagnostics: () => ({ entries: diagnostics })
  }
};
if (process.env.SLIDESCI_ACCEPTANCE === '1') handlers.acceptance = acceptanceHandlers();
function sameToken(value) {
  if (typeof value !== 'string') return false;
  const bytes = Buffer.from(value);
  const expected = Buffer.from(token);
  return bytes.length === expected.length && timingSafeEqual(bytes, expected);
}
function json(response, status, data) {
  response.writeHead(status, { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store',
    'X-Content-Type-Options': 'nosniff', 'Referrer-Policy': 'no-referrer' });
  response.end(JSON.stringify(data));
}
async function body(request) {
  const chunks = [];
  let length = 0;
  for await (const chunk of request) {
    length += chunk.length;
    if (length > 256 * 1024 * 1024) throw new Error('请求数据超过256MB，请分批处理图片或页面。');
    chunks.push(chunk);
  }
  const result = JSON.parse(Buffer.concat(chunks).toString('utf8') || '{}');
  if (!result || typeof result !== 'object' || Array.isArray(result)) throw new Error('请求参数必须是对象。');
  return result;
}
function hostKind(value) {
  return /powerpoint/i.test(String(value)) ? 'powerpoint' : /wps|kingsoft/i.test(String(value)) ? 'wps' : null;
}
async function broadcast(name, input) {
  let data = input || {};
  if (data.files?.length) {
    try { data = { ...data, files: await env.files.grant(data.files) }; }
    catch (error) { broadcast('nativeError', { message: error.message }); return; }
  }
  const eventName = data.type === 'clipboardImage' ? 'clipboardImage' : name;
  const frame = 'event: ' + eventName + '\ndata: ' + JSON.stringify(data) + '\n\n';
  for (const client of clients) {
    if (client.host && hostKind(data.host) && client.host !== hostKind(data.host)) continue;
    client.response.write(frame);
  }
}
for (const name of ['changed', 'mediaSource', 'clipboardImage', 'nativeError']) events.on(name, input => {
  broadcast(name, input).catch(error => diagnose(error.message, '推送事件失败：' + error.message));
});
const mimes = { '.html': 'text/html; charset=utf-8', '.css': 'text/css; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8', '.mjs': 'text/javascript; charset=utf-8', '.xml': 'application/xml; charset=utf-8',
  '.png': 'image/png', '.svg': 'image/svg+xml', '.jpg': 'image/jpeg', '.woff2': 'font/woff2' };
async function serve(request, response) {
  try {
    const hostname = new URL('http://' + request.headers.host).hostname;
    if (!['localhost', '127.0.0.1', '[::1]'].includes(hostname) || (request.headers.origin && !origins.has(request.headers.origin))) {
      json(response, 403, { error: '仅允许本机 SlideSCI 界面访问。' }); return;
    }
    const url = new URL(request.url, 'http://localhost');
    if (url.pathname === '/api/session' && request.method === 'GET') {
      json(response, 200, { token, version, instance, warnings: accessWarnings, acceptance: process.env.SLIDESCI_ACCEPTANCE === '1' }); return;
    }
    if (url.pathname === '/api/events' && request.method === 'GET') {
      if (!sameToken(url.searchParams.get('token'))) { json(response, 403, { error: '会话授权已失效，请重新加载插件。' }); return; }
      response.writeHead(200, { 'Content-Type': 'text/event-stream; charset=utf-8', 'Cache-Control': 'no-store',
        Connection: 'keep-alive', 'X-Content-Type-Options': 'nosniff', 'Referrer-Policy': 'no-referrer' });
      response.write(': connected\n\n');
      const client = { response, host: url.searchParams.get('host'), client: clientId(url.searchParams.get('client')) };
      clients.add(client);
      if (client.client) { clearTimeout(dropTimers.get(client.client)); dropTimers.delete(client.client); }
      // Tells the pane which server it reached and which of its leases survived, so it can re-acquire after a restart.
      response.write('event: hello\ndata: ' + JSON.stringify({ instance, leases: client.client ? leasesOf(client.client) : [] }) + '\n\n');
      const heartbeat = setInterval(() => response.write(': heartbeat\n\n'), 20000);
      request.on('close', () => { clearInterval(heartbeat); clients.delete(client); scheduleLeaseDrop(client.client); });
      return;
    }
    if (url.pathname.startsWith('/api/')) {
      if (request.method !== 'POST') { json(response, 405, { error: '此接口需要POST。' }); return; }
      if (!sameToken(request.headers['x-slidesci-token'])) { json(response, 403, { error: '会话授权已失效，请重新加载插件。' }); return; }
      const parts = url.pathname.slice(5).split('/');
      if (parts.length !== 2 || !Object.hasOwn(handlers, parts[0]) || !Object.hasOwn(handlers[parts[0]], parts[1])) {
        json(response, 404, { error: '未知的 SlideSCI 操作。' }); return;
      }
      const result = await handlers[parts[0]][parts[1]](await body(request), env);
      json(response, 200, result === undefined ? {} : result); return;
    }
    if (!['GET', 'HEAD'].includes(request.method)) { response.writeHead(405, { Allow: 'GET, HEAD' }); response.end(); return; }
    let resource = decodeURIComponent(url.pathname);
    if (resource === '/') resource = '/taskpane.html';
    if (resource === '/wps/') resource = '/wps/index.html';
    if (!['/taskpane.html', '/taskpane.css'].includes(resource) && !/^\/(src|assets|wps)\//.test(resource)) {
      response.writeHead(404); response.end(); return;
    }
    const filename = await realpath(path.join(root, resource.slice(1))).catch(() => null);
    if (!filename || !filename.startsWith(root) || !(await stat(filename)).isFile() || !mimes[path.extname(filename)]) {
      response.writeHead(404); response.end(); return;
    }
    response.writeHead(200, { 'Content-Type': mimes[path.extname(filename)], 'Cache-Control': 'no-store',
      'X-Content-Type-Options': 'nosniff', 'Referrer-Policy': 'no-referrer' });
    response.end(request.method === 'HEAD' ? undefined : await readFile(filename));
  } catch (error) {
    // The path only: the event stream carries its session token in the query string.
    log(`${request.method} ${String(request.url || '').split('?')[0].slice(0, 200)} 失败：${error.message || error}`, { error: true });
    if (!response.headersSent) json(response, 400, { error: error.message || 'SlideSCI 操作失败。' });
    else response.end();
  }
}
const servers = [];
function listen(server, port, host, label) {
  return new Promise((resolve, reject) => {
    server.once('error', reject);
    server.listen(port, host, () => {
      server.off('error', reject);
      server.on('error', error => diagnose(error.message, '服务连接错误：' + error.message));
      servers.push(server);
      log('SlideSCI ' + label);
      resolve();
    });
  });
}
try {
  await listen(http.createServer(serve), httpPort, '127.0.0.1', `http://127.0.0.1:${httpPort}`);
  if (values.cert) {
    const tls = { cert: await readFile(values.cert), key: await readFile(values.key) };
    // PowerPoint resolves localhost to ::1 or 127.0.0.1; serve every loopback address it may pick.
    const addresses = [...new Set((await lookup('localhost', { all: true })).map(item => item.address))]
      .filter(address => ['127.0.0.1', '::1'].includes(address));
    let bound = 0;
    for (const address of addresses) {
      try { await listen(https.createServer(tls, serve), httpsPort, address, `https://localhost:${httpsPort} (${address})`); bound++; }
      catch (error) { if (error.code !== 'EADDRNOTAVAIL') throw error; }
    }
    if (!bound) throw new Error(`无法在本机回环地址监听 ${httpsPort} 端口。`);
  } else log('PowerPoint: 请提供已在 macOS 信任的本机证书 --cert <pem> --key <pem>。');
} catch (error) {
  log(error.code === 'EADDRINUSE' ? `SlideSCI 无法启动：端口 ${error.port} 已被其他程序占用（可能是另一个 SlideSCI 服务）。` : `SlideSCI 无法启动：${error.message}`, { error: true });
  await rm(env.workDir, { recursive: true, force: true }).catch(() => {});
  // Under the login LaunchAgent a non-zero exit is restarted every 10 s; a taken port or a bad
  // certificate does not heal by retrying, so exit cleanly and let Start.command report it.
  process.exit(process.env.SLIDESCI_SUPERVISOR === 'launchd' ? 0 : 1);
}
// The versions a bug report needs, once per start.
const macos = (await readFile('/System/Library/CoreServices/SystemVersion.plist', 'utf8').catch(() => ''))
  .match(/<key>ProductVersion<\/key>\s*<string>([^<]+)<\/string>/)?.[1] || `Darwin ${os.release()}`;
log(`SlideSCI ${version} 已启动（进程 ${process.pid}，Node ${process.version}，macOS ${macos}，${process.env.SLIDESCI_SUPERVISOR === 'launchd' ? '登录自启' : '手动启动'}）`);

// Only sweep this process's own tmpdir workspace at startup; host container caches are never
// touched here (native.mjs sweeps a host cache when the service first writes into it).
const CACHE_AGE = 24 * 3600 * 1000;
async function sweepWorkDirs() {
  const cutoff = Date.now() - CACHE_AGE;
  const remove = async file => { try { if ((await lstat(file)).mtimeMs < cutoff) await rm(file, { recursive: true, force: true }); } catch { /* best effort */ } };
  const alive = pid => { try { process.kill(pid, 0); return true; } catch (error) { return error.code === 'EPERM'; } };
  try {
    for (const name of await readdir(os.tmpdir())) {
      const file = path.join(os.tmpdir(), name), owner = /^slidesci-mac-(\d+)-[0-9a-f-]{36}$/i.exec(name);
      if (file === env.workDir) continue;
      if (owner && !alive(Number(owner[1]))) await rm(file, { recursive: true, force: true }).catch(() => {});
      else if (/^slidesci-mac-[0-9a-f-]{36}$/i.test(name)) await remove(file);
    }
  } catch { /* best effort */ }
}
sweepWorkDirs().catch(error => diagnose('清理临时文件失败：' + error.message));

let stopping = false;
for (const signal of ['SIGINT', 'SIGTERM']) process.on(signal, async () => {
  if (stopping) return;
  stopping = true;
  // Never wait forever on a keep-alive connection or a stuck helper.
  setTimeout(() => process.exit(0), 3000).unref();
  for (const client of clients) client.response.end();
  for (const timer of dropTimers.values()) clearTimeout(timer);
  try { await nativeHandlers.watchStop({}, env); } catch { /* the watcher exits with its parent */ }
  await Promise.all(servers.map(server => new Promise(resolve => { server.close(() => resolve()); server.closeAllConnections?.(); })));
  await rm(env.workDir, { recursive: true, force: true }).catch(() => {});
  log(`SlideSCI 已停止（${signal}）`);
  flushRepeats();
  process.exit(0);
});
