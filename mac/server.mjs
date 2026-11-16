import http from 'node:http';
import https from 'node:https';
import path from 'node:path';
import { readFile, realpath, stat } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { parseArgs } from 'node:util';
import { randomBytes, randomUUID, timingSafeEqual } from 'node:crypto';
import { EventEmitter } from 'node:events';
import { contentHandlers } from './lib/content.mjs';
import { imageHandlers } from './lib/images.mjs';
import { libraryHandlers } from './lib/library.mjs';
import { aiHandlers } from './lib/ai.mjs';
import { nativeHandlers } from './lib/native.mjs';
import { pptx } from './lib/pptx.mjs';
import { createFileServices } from './lib/storage.mjs';
import { acceptanceHandlers } from './lib/acceptance.mjs';

const { values } = parseArgs({ options: { cert: { type: 'string' }, key: { type: 'string' } } });
if (Boolean(values.cert) !== Boolean(values.key)) throw new Error('--cert 和 --key 必须一起提供。');
const root = fileURLToPath(new URL('.', import.meta.url));
const version = JSON.parse(await readFile(new URL('./package.json', import.meta.url), 'utf8')).version;
const token = randomBytes(32).toString('hex');
const origins = new Set(['http://127.0.0.1:3000', 'http://localhost:3000', 'https://localhost:3443', 'https://127.0.0.1:3443']);
const events = new EventEmitter();
events.setMaxListeners(100);
const env = { events, pptx, version };
env.native = Object.fromEntries(Object.entries(nativeHandlers).map(([name, fn]) => [name, input => fn(input || {}, env)]));
Object.assign(env, await createFileServices(env));
const diagnostics = [];
const clients = new Set();
const watchLeases = new Map();
let watchState = null;
let watchQueue = Promise.resolve();
function watchOperation(operation) {
  const result = watchQueue.then(operation, operation);
  watchQueue = result.catch(() => {});
  return result;
}
async function startWatch(input) {
  return watchOperation(async () => {
    if (!watchState) watchState = await nativeHandlers.watchStart({ interceptTIFF: input.interceptTIFF === true }, env);
    const id = randomUUID();
    watchLeases.set(id, { host: input.host, created: Date.now() });
    return { ...watchState, id, shared: true };
  });
}
async function stopWatch(input) {
  return watchOperation(async () => {
    if (input.id) watchLeases.delete(input.id);
    if (!watchLeases.size && watchState) {
      await nativeHandlers.watchStop({ id: watchState.id }, env);
      watchState = null;
    }
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
  grant: input => env.files.grant((input.paths || []).map(value => typeof value === 'string' ? value : value.path)),
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
  content: contentHandlers, images: imageHandlers, library: libraryHandlers, ai: aiHandlers,
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
      diagnostics.push({ at: new Date().toISOString(), message: String(input.message || '客户端错误').slice(0, 4000) });
      if (diagnostics.length > 30) diagnostics.shift();
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
  broadcast(name, input).catch(error => { diagnostics.push({ at: new Date().toISOString(), message: error.message }); });
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
      json(response, 200, { token, version, acceptance: process.env.SLIDESCI_ACCEPTANCE === '1' }); return;
    }
    if (url.pathname === '/api/events' && request.method === 'GET') {
      if (!sameToken(url.searchParams.get('token'))) { json(response, 403, { error: '会话授权已失效，请重新加载插件。' }); return; }
      response.writeHead(200, { 'Content-Type': 'text/event-stream; charset=utf-8', 'Cache-Control': 'no-store',
        Connection: 'keep-alive', 'X-Content-Type-Options': 'nosniff', 'Referrer-Policy': 'no-referrer' });
      response.write(': connected\n\n');
      const client = { response, host: url.searchParams.get('host') };
      clients.add(client);
      const heartbeat = setInterval(() => response.write(': heartbeat\n\n'), 20000);
      request.on('close', () => { clearInterval(heartbeat); clients.delete(client); });
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
    if (!response.headersSent) json(response, 400, { error: error.message || 'SlideSCI 操作失败。' });
    else response.end();
  }
}
const servers = [];
function listen(server, port, host) {
  server.on('error', error => { console.error(error.message); process.exitCode = 1; });
  server.listen(port, host, () => console.log('SlideSCI ' + (port === 3000 ? 'http://127.0.0.1:3000' : 'https://localhost:3443')));
  servers.push(server);
}
listen(http.createServer(serve), 3000, '127.0.0.1');
if (values.cert) {
  listen(https.createServer({ cert: await readFile(values.cert), key: await readFile(values.key) }, serve), 3443, 'localhost');
} else console.log('PowerPoint: 请提供已在 macOS 信任的本机证书 --cert <pem> --key <pem>。');
for (const signal of ['SIGINT', 'SIGTERM']) process.on(signal, async () => {
  for (const client of clients) client.response.end();
  await nativeHandlers.watchStop({}, env);
  for (const server of servers) server.close();
});
