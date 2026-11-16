const SETTINGS_KEY = 'slidesci:settings:v1';
const REQUEST_TIMEOUT = 300000;
const SECRET_KEY = /(?:secret|token|password|passwd|credential|authorization|bearer|api[-_. ]?key)/i;

export class ApiError extends Error {
  constructor(message, { status = 0, code = '', detail = '' } = {}) {
    super(message);
    this.name = 'ApiError';
    this.status = status;
    this.code = code;
    this.detail = detail;
    this.missingRoute = status === 404;
  }
}

function clone(value) {
  if (value === undefined) return value;
  return typeof structuredClone === 'function' ? structuredClone(value) : JSON.parse(JSON.stringify(value));
}

function abortError() {
  return new DOMException('操作已取消。', 'AbortError');
}

function isAbort(error, signal) {
  return signal?.aborted || error?.name === 'AbortError';
}

function secretFree(key, value) {
  if (SECRET_KEY.test(String(key))) return false;
  if (value && typeof value === 'object') {
    for (const [name, item] of Object.entries(value)) {
      if (SECRET_KEY.test(name) || (typeof item === 'string' && /^sk-[A-Za-z0-9_-]{16,}/.test(item))) return false;
    }
  }
  return !(typeof value === 'string' && /^sk-[A-Za-z0-9_-]{16,}/.test(value));
}

class Settings {
  read() {
    try {
      const raw = globalThis.localStorage?.getItem(SETTINGS_KEY);
      if (!raw) return {};
      const parsed = JSON.parse(raw);
      return parsed && typeof parsed === 'object' && !Array.isArray(parsed) ? parsed : {};
    } catch (error) {
      console.warn('SlideSCI 读取本地设置失败：' + error.message);
      return {};
    }
  }

  write(value) {
    const storage = globalThis.localStorage;
    if (!storage) throw new Error('无法访问本地设置存储，请检查演示软件的存储权限。');
    try {
      storage.setItem(SETTINGS_KEY, JSON.stringify(value));
    } catch {
      throw new Error('无法保存本地设置，请检查演示软件的存储权限。');
    }
  }

  get(key, fallback = undefined) {
    const store = this.read();
    return Object.prototype.hasOwnProperty.call(store, key) ? clone(store[key]) : fallback;
  }

  set(key, value) {
    if (typeof key !== 'string' || !key.trim()) throw new Error('设置名称无效。');
    if (!secretFree(key, value)) throw new Error('设置只保存非敏感配置，请勿在本地设置中保存密钥或凭据。');
    const store = this.read();
    store[key] = clone(value);
    this.write(store);
    return value;
  }

  remove(key) {
    const store = this.read();
    delete store[key];
    this.write(store);
  }

  all() {
    return clone(this.read());
  }

  clear() {
    const storage = globalThis.localStorage;
    if (!storage) throw new Error('无法访问本地设置存储，请检查演示软件的存储权限。');
    try {
      storage.removeItem(SETTINGS_KEY);
    } catch {
      throw new Error('无法清除本地设置，请检查演示软件的存储权限。');
    }
  }
}

class EventStream {
  constructor(services) {
    this.services = services;
    this.listeners = new Map();
    this.registered = new Set();
    this.source = null;
    this.connected = false;
    this.closed = false;
    this.retry = null;
    this.attempt = 0;
    this.started = false;
  }

  subscribe(name, callback) {
    if (typeof name !== 'string' || !name || typeof callback !== 'function') {
      throw new Error('事件订阅参数无效。');
    }
    const handlers = this.listeners.get(name) || new Set();
    const handler = payload => {
      try {
        callback(payload, name);
      } catch (error) {
        console.error('SlideSCI 事件处理失败：' + error.message);
      }
    };
    handlers.add(handler);
    this.listeners.set(name, handlers);
    this.ensure();
    if (this.source && !this.registered.has(name)) {
      this.registered.add(name);
      this.source.addEventListener(name, event => this.dispatchNamed(event));
    }
    return () => {
      handlers.delete(handler);
      if (!handlers.size) this.listeners.delete(name);
    };
  }

  ensure() {
    if (this.started) return;
    this.started = true;
    this.connect();
  }

  async connect() {
    if (this.closed) return;
    clearTimeout(this.retry);
    let session;
    try {
      session = await this.services.session();
    } catch (error) {
      this.schedule(error.name === 'AbortError' ? 5000 : 3000);
      return;
    }
    if (this.closed) return;
    let source;
    try {
      const url = this.services.path('/api/events?token=' + encodeURIComponent(session.token)
        + '&host=' + encodeURIComponent(this.services.host));
      source = new EventSource(url);
    } catch (error) {
      this.schedule(5000);
      return;
    }
    this.source = source;
    this.registered.clear();
    source.onopen = () => {
      this.connected = true;
      this.attempt = 0;
      this.services.emitDiagnostic('events-open');
    };
    source.onerror = () => {
      if (this.closed || this.source !== source) return;
      this.connected = false;
      source.close();
      this.source = null;
      this.schedule();
    };
    for (const name of this.listeners.keys()) {
      this.registered.add(name);
      source.addEventListener(name, event => this.dispatchNamed(event));
    }
    source.onmessage = event => this.dispatchGeneric(event);
  }

  schedule(delay) {
    if (this.closed) return;
    clearTimeout(this.retry);
    const wait = delay ?? Math.min(15000, 800 * 2 ** Math.min(this.attempt++, 4));
    this.retry = setTimeout(() => this.connect(), wait);
  }

  payload(event) {
    if (!event?.data) return null;
    try {
      return JSON.parse(event.data);
    } catch {
      return { message: String(event.data) };
    }
  }

  dispatchNamed(event) {
    const payload = this.payload(event);
    if (!payload) return;
    for (const handler of this.listeners.get(event.type) || []) handler(payload);
  }

  dispatchGeneric(event) {
    const payload = this.payload(event);
    if (!payload) return;
    const name = payload.type || payload.event;
    if (name && this.listeners.has(name)) {
      for (const handler of this.listeners.get(name)) handler(payload);
    }
  }

  close() {
    this.closed = true;
    clearTimeout(this.retry);
    this.source?.close();
    this.source = null;
    this.connected = false;
    this.listeners.clear();
  }
}

export class Services {
  constructor({ host = 'powerpoint', baseUrl = globalThis.location?.origin || 'http://127.0.0.1:3000' } = {}) {
    if (!['powerpoint', 'wps'].includes(host)) throw new Error('未知的演示软件类型。');
    this.host = host;
    this.baseUrl = String(baseUrl).replace(/\/+$/, '');
    this.version = null;
    this.token = null;
    this.sessionPromise = null;
    this.settings = new Settings();
    this.events = new EventStream(this);
    this.diagnostics = [];
  }

  path(relative) {
    return this.baseUrl + relative;
  }

  emitDiagnostic(name) {
    this.diagnostics.push({ name, at: Date.now() });
    if (this.diagnostics.length > 50) this.diagnostics.shift();
  }

  async ready() {
    const session = await this.session();
    return session;
  }

  async session(force = false) {
    if (!force && this.sessionPromise) return this.sessionPromise;
    this.sessionPromise = (async () => {
      let response;
      const controller = new AbortController();
      const timer = setTimeout(() => controller.abort(), 15000);
      try {
        response = await fetch(this.path('/api/session'), { headers: { Accept: 'application/json' }, cache: 'no-store', signal: controller.signal });
      } catch (error) {
        this.sessionPromise = null;
        throw new ApiError(controller.signal.aborted
          ? '连接 SlideSCI 本地服务超时，请重启服务后重试。'
          : '无法连接 SlideSCI 本地服务，请确认服务已启动后重试。', { detail: error.message });
      } finally {
        clearTimeout(timer);
      }
      if (!response.ok) {
        this.sessionPromise = null;
        throw new ApiError(response.status === 404
          ? '本地服务版本不匹配，请重启 SlideSCI 服务后重试。'
          : '无法建立与本地服务的安全会话，请重启 SlideSCI 服务后重试。', { status: response.status });
      }
      let body;
      try {
        body = await response.json();
      } catch {
        this.sessionPromise = null;
        throw new ApiError('本地服务返回了无法解析的会话数据，请重启服务后重试。', { status: response.status });
      }
      if (!body || typeof body.token !== 'string' || !body.token) {
        this.sessionPromise = null;
        throw new ApiError('本地服务没有返回有效会话，请重启 SlideSCI 服务后重试。');
      }
      this.version = body.version ?? null;
      this.token = body.token;
      return body;
    })();
    return this.sessionPromise;
  }

  async request(name, payload = {}, { signal, keepalive = false, timeout = REQUEST_TIMEOUT } = {}) {
    const parts = String(name).split('/');
    if (parts.length !== 2 || !parts[0] || !parts[1]) throw new ApiError('服务请求名称无效。');
    const [namespace, action] = parts;
    let refreshed = false;
    for (;;) {
      const session = await this.session();
      try {
        return await this.post(namespace, action, payload, session.token, { signal, keepalive, timeout });
      } catch (error) {
        if (error instanceof ApiError && (error.status === 401 || error.status === 403) && !refreshed) {
          refreshed = true;
          await this.session(true);
          continue;
        }
        throw error;
      }
    }
  }

  async post(namespace, action, payload, token, { signal, keepalive, timeout }) {
    if (signal?.aborted) throw abortError();
    const controller = new AbortController();
    const timer = timeout ? setTimeout(() => controller.abort(new DOMException('本地服务响应超时。', 'TimeoutError')), timeout) : null;
    const stop = () => controller.abort();
    signal?.addEventListener('abort', stop, { once: true });
    let response;
    try {
      response = await fetch(this.path('/api/' + encodeURIComponent(namespace) + '/' + encodeURIComponent(action)), {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', Accept: 'application/json', 'X-SlideSCI-Token': token },
        body: JSON.stringify(payload ?? {}),
        signal: controller.signal,
        keepalive,
        cache: 'no-store'
      });
    } catch (error) {
      if (isAbort(error, signal)) throw abortError();
      if (controller.signal.aborted) throw new ApiError('本地服务处理超时，请稍后重试。');
      throw new ApiError('无法连接 SlideSCI 本地服务，请确认服务正在运行后重试。', { detail: error.message });
    } finally {
      clearTimeout(timer);
      signal?.removeEventListener('abort', stop);
    }
    const text = await response.text();
    let body = null;
    if (text) {
      try {
        body = JSON.parse(text);
      } catch {
        if (response.ok) throw new ApiError('本地服务返回了无法解析的数据，请重启服务后重试。', { status: response.status });
      }
    }
    if (!response.ok) {
      const message = typeof body?.error === 'string' ? body.error
        : typeof body?.message === 'string' ? body.message
          : response.status === 401 || response.status === 403 ? '本地服务会话无效，请重启服务后重试。'
            : response.status === 404 ? '本地服务缺少所需功能，请更新或重启 SlideSCI 服务。'
              : response.status === 413 ? '文件过大，无法处理。'
                : '本地服务处理失败，请稍后重试。';
      if (response.status === 401 || response.status === 403) this.sessionPromise = null;
      throw new ApiError(message, { status: response.status, code: body?.code ?? '', detail: body?.detail ?? '' });
    }
    return body;
  }

  subscribe(name, callback) {
    return this.events.subscribe(name, callback);
  }

  closeEvents() {
    this.events.close();
  }

  async pickFiles({ kind = 'all', multiple = true } = {}) {
    const files = await this.request('files/pickFiles', { kind, multiple, host: this.host });
    return Array.isArray(files) ? files : files?.files || [];
  }

  async chooseFolder() {
    const result = await this.request('files/chooseFolder', { host: this.host });
    return typeof result === 'string' ? result : result?.path ?? null;
  }

  async download(file) {
    if (!file || typeof file.data !== 'string' || !file.data) throw new ApiError('没有可下载的文件数据。');
    return this.request('files/download', {
      name: file.name || 'SlideSCI',
      mime: file.mime || 'application/octet-stream',
      data: file.data,
      host: this.host
    });
  }

  async reveal(path) {
    if (typeof path !== 'string' || !path.startsWith('/')) throw new ApiError('没有可显示的本机文件位置。');
    return this.request('files/reveal', { path });
  }

  async readClipboardFiles() {
    const files = await this.request('files/readClipboardFiles', { host: this.host });
    return Array.isArray(files) ? files : files?.files || [];
  }

  async writeClipboard({ data, mime } = {}) {
    if (typeof data !== 'string' || !data || typeof mime !== 'string' || !mime) throw new ApiError('剪贴板内容无效。');
    return this.request('files/writeClipboard', { data, mime });
  }

  async grantFiles(paths, { source = 'clipboard' } = {}) {
    const values = (Array.isArray(paths) ? paths : [paths]).filter(path => typeof path === 'string' && path.startsWith('/'));
    if (!values.length) return [];
    return this.request('files/grant', { paths: values, values, source, host: this.host });
  }

  async fonts() {
    const result = await this.request('native/fonts', {});
    return Array.isArray(result) ? result : result?.fonts || [];
  }

  async openExternal(url) {
    let parsed;
    try {
      parsed = new URL(String(url));
    } catch {
      throw new ApiError('外部链接格式无效。');
    }
    if (!['http:', 'https:'].includes(parsed.protocol)) throw new ApiError('只能打开 http 或 https 链接。');
    const result = await this.request('files/openExternal', { url: parsed.href });
    return result?.opened ? parsed.href : result;
  }

  async startWatcher({ interceptTIFF = false } = {}) {
    const result = await this.request('native/watchStart', { host: this.host, interceptTIFF });
    return result?.id ? result : null;
  }

  async stopWatcher(id) {
    if (!id) return;
    try {
      if (this.token) await this.post('native', 'watchStop', { id }, this.token, { keepalive: true, timeout: 5000 });
      else await this.request('native/watchStop', { id }, { keepalive: true, timeout: 5000 });
    } catch (error) {
      if (error instanceof ApiError && (error.status === 401 || error.status === 403)) {
        try {
          await this.request('native/watchStop', { id }, { keepalive: true, timeout: 5000 });
          return;
        } catch (retryError) {
          this.reportError(retryError, { operation: 'watchStop' });
          return;
        }
      }
      if (!(error instanceof ApiError && error.missingRoute)) this.reportError(error, { operation: 'watchStop' });
    }
  }

  async reportError(error, context = {}) {
    const redact = value => String(value)
      .replace(/sk-[A-Za-z0-9_-]{16,}/g, '[密钥已隐藏]')
      .replace(/Bearer\s+[A-Za-z0-9._-]+/gi, 'Bearer [已隐藏]');
    const payload = {
      message: redact(error?.message ? String(error.message) : String(error)).slice(0, 4000),
      name: error?.name || 'Error',
      stack: typeof error?.stack === 'string' ? redact(error.stack).slice(0, 8000) : '',
      context: clone(context),
      host: this.host,
      version: this.version,
      at: new Date().toISOString()
    };
    console.error('SlideSCI 操作失败：' + payload.message, context);
    try {
      await this.request('client/error', payload, { timeout: 5000 });
    } catch {
      this.emitDiagnostic('report-error-failed');
    }
    return payload;
  }
}

export function createServices(options) {
  return new Services(options);
}

export default createServices;
