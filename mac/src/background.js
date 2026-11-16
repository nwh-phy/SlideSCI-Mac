import { handleNativePaste } from './features/images.js';
import { initialize as initializeScienceFeature } from './features/science.js';

const runtimes = new Map();
const EVENT_NAMES = ['mediaSource', 'clipboardImage', 'changed', 'nativeError'];
const HOST_IDS = {
  powerpoint: ['com.microsoft.Powerpoint'],
  wps: ['com.kingsoft.wpsoffice.mac', 'com.kingsoft.wpsoffice']
};
const PENDING_LIFETIME = 90000;
const FAST_POLL_WINDOW = 6000;
const FAST_POLL_INTERVAL = 400;

function messageOf(error) {
  return error?.message ? String(error.message) : String(error);
}

function matchesHost(payload, host) {
  const value = payload?.host;
  if (!value) return false;
  const normalized = String(value).toLowerCase();
  return normalized === host
    || (host === 'powerpoint' && normalized === 'presentation')
    || (HOST_IDS[host] || []).some(id => id.toLowerCase() === normalized);
}

function mimeOf(name = '') {
  const extension = String(name).toLowerCase().match(/\.([a-z0-9]+)$/)?.[1];
  return {
    png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif',
    tif: 'image/tiff', tiff: 'image/tiff', bmp: 'image/bmp', svg: 'image/svg+xml',
    webp: 'image/webp', emf: 'image/x-emf', wmf: 'image/x-wmf',
    mp4: 'video/mp4', mov: 'video/quicktime', m4v: 'video/mp4', avi: 'video/x-msvideo'
  }[extension] || 'application/octet-stream';
}

function validPosition(position) {
  if (!position || !Number.isFinite(position.x) || !Number.isFinite(position.y)) return undefined;
  return { x: position.x, y: position.y };
}

class BackgroundTracking {
  constructor({ host, services, hostAdapter = null, role = 'foreground', interceptTIFF = false, science, onStatus, onEvent, onHostChanged }) {
    if (!services) throw new Error('后台跟踪需要本地服务。');
    this.host = host;
    this.services = services;
    this.role = role;
    this.interceptTIFF = Boolean(interceptTIFF);
    this.shouldInitializeScience = science ?? (role === 'wps-hidden');
    this.onStatus = typeof onStatus === 'function' ? onStatus : null;
    this.onEvent = typeof onEvent === 'function' ? onEvent : null;
    this.onHostChanged = typeof onHostChanged === 'function' ? onHostChanged : null;
    this.injectedAdapter = hostAdapter;
    this.adapterPromise = hostAdapter ? Promise.resolve(hostAdapter) : null;
    this.leaseId = null;
    this.stopped = false;
    this.queue = Promise.resolve();
    this.unsubscribers = [];
    this.controller = new AbortController();
    this.pending = [];
    this.knownIds = null;
    this.handledIds = new Set();
    this.pendingPoll = null;
    this.pendingDeadline = 0;
    this.scienceCleanup = null;
    this.scienceFailed = false;
    this.scienceRetryTimer = null;
    this.scienceAttempts = 0;
    this.lastScienceAttempt = 0;
    this.lastFailure = null;
    this.hostUnsubscribe = null;
    this.started = false;
  }

  status(message, kind = 'info') {
    try {
      this.onStatus?.({ message, kind, host: this.host, role: this.role });
    } catch (error) {
      console.warn('SlideSCI 后台状态回调失败：' + messageOf(error));
    }
  }

  notify(message, kind = 'info') {
    this.status(message, kind);
  }

  report(error, operation, context = {}) {
    const message = messageOf(error);
    const now = Date.now();
    if (this.lastFailure?.message === message && now - this.lastFailure.at < 30000) return;
    this.lastFailure = { message, at: now };
    this.services.reportError(error, { operation, host: this.host, role: this.role, ...context })
      .catch(reportError => console.warn('SlideSCI 错误上报失败：' + messageOf(reportError)));
  }

  async start() {
    if (this.started) return this;
    this.started = true;
    this.subscribeEvents();
    globalThis.addEventListener?.('pagehide', this.pagehide, { once: true });
    try {
      const watcher = await this.services.startWatcher({ interceptTIFF: this.interceptTIFF });
      this.leaseId = watcher?.id || null;
      if (!this.leaseId) throw new Error('原生媒体监听没有返回有效租约。');
    } catch (error) {
      this.stopped = true;
      this.unsubscribeEvents();
      this.report(error, 'background-start');
      throw error;
    }
    if (this.shouldInitializeScience) await this.startScience();
    try {
      const adapter = await this.hostAdapter();
      const context = await adapter.context({ scope: 'slide', geometry: true });
      this.knownIds = new Set((context.shapes || []).map(shape => String(shape.id)));
      if (typeof adapter.onChanged === 'function') {
        this.hostUnsubscribe = adapter.onChanged(() => {
          if (!this.stopped && this.pending.length) {
            this.checkPending().catch(error => this.report(error, 'pending-host-change'));
          }
        });
      }
    } catch (error) {
      this.report(error, 'shape-baseline');
    }
    return this;
  }

  pagehide = () => {
    this.stop({ keepalive: true }).catch(error => console.warn('SlideSCI 后台跟踪停止失败：' + messageOf(error)));
  };

  subscribeEvents() {
    for (const name of EVENT_NAMES) {
      this.unsubscribers.push(this.services.subscribe(name, payload => this.receive(name, payload)));
    }
  }

  unsubscribeEvents() {
    for (const unsubscribe of this.unsubscribers) {
      try {
        unsubscribe();
      } catch {
        // Unsubscription is idempotent.
      }
    }
    this.unsubscribers = [];
  }

  receive(name, payload) {
    if (this.stopped) return;
    if (name === 'nativeError') {
      const error = new Error(payload?.message ? String(payload.message) : '原生媒体监听失败。');
      this.status(error.message, 'error');
      this.report(error, 'native-watcher');
      return;
    }
    if (name === 'changed') {
      if (!matchesHost(payload, this.host)) return;
      this.checkPending().catch(error => this.report(error, 'pending-check'));
      if (this.shouldInitializeScience && !this.scienceCleanup && !this.scienceRetryTimer
        && Date.now() - this.lastScienceAttempt > 15000) this.retryScience(true);
      try {
        this.onHostChanged?.(payload);
      } catch (error) {
        this.report(error, 'host-changed');
      }
      return;
    }
    this.considerMedia(name, payload);
  }

  considerMedia(name, payload) {
    if (!payload || !matchesHost(payload, this.host)) return;
    const clipboardImage = name === 'clipboardImage' || payload.type === 'clipboardImage';
    const files = Array.isArray(payload.files)
      ? payload.files.filter(file => typeof file?.path === 'string' && file.path.startsWith('/'))
        .map(file => ({ ...file, mime: file.mime || mimeOf(file.name || file.path) }))
      : [];
    if (!files.length && !(clipboardImage && typeof payload.data === 'string' && payload.data)) return;
    if (clipboardImage && this.pending.some(item => item.intercepted && item.files.length && Date.now() - item.at < 5000)) return;
    const signature = clipboardImage
      ? `${this.host}|image|${payload.mime || ''}|${payload.data.length}`
      : `${this.host}|files|${files.map(file => file.path).join('|')}`;
    const now = Date.now();
    this.pending = this.pending.filter(item => now - item.at < PENDING_LIFETIME && item.signature !== signature);
    const candidate = {
      signature,
      at: now,
      baseline: this.knownIds ? new Set(this.knownIds) : null,
      baselineReady: this.captureBaseline(),
      files,
      clipboard: clipboardImage ? { data: payload.data, mime: payload.mime || 'image/tiff' } : null,
      intercepted: payload.intercepted === true,
      reason: payload.reason || '',
      position: validPosition(payload.position)
    };
    this.pending.push(candidate);
    if (this.pending.length > 4) this.pending.shift();
    const paths = files.map(file => file.path);
    if (paths.length) {
      this.services.grantFiles(paths, { source: clipboardImage ? 'clipboardImage' : candidate.reason || 'clipboard' })
        .catch(error => {
          if (!error?.missingRoute) this.report(error, 'grant-files');
        });
    }
    this.beginPendingObservation();
  }

  captureBaseline() {
    return this.hostAdapter()
      .then(adapter => adapter.context({ scope: 'slide', geometry: true }))
      .then(context => ({
        ids: new Set((context.shapes || []).map(shape => String(shape.id))),
        scope: `${context.presentationId || ''}|${context.slideId || ''}`
      }))
      .catch(error => {
        this.report(error, 'paste-baseline');
        return null;
      });
  }

  beginPendingObservation() {
    this.pendingDeadline = Math.max(this.pendingDeadline, Date.now() + FAST_POLL_WINDOW);
    if (this.pendingPoll) return;
    const tick = async () => {
      if (this.stopped || !this.pending.length) {
        this.pendingPoll = null;
        return;
      }
      try {
        await this.checkPending();
      } catch (error) {
        this.report(error, 'pending-poll');
      }
      if (!this.stopped && this.pending.length) {
        this.pendingPoll = setTimeout(tick, Date.now() < this.pendingDeadline ? FAST_POLL_INTERVAL : 3000);
      } else {
        this.pendingPoll = null;
      }
    };
    this.pendingPoll = setTimeout(tick, FAST_POLL_INTERVAL);
  }

  async checkPending() {
    if (this.stopped || !this.pending.length) return;
    const now = Date.now();
    this.pending = this.pending.filter(item => now - item.at < PENDING_LIFETIME);
    if (!this.pending.length) return;
    let context;
    try {
      const adapter = await this.hostAdapter();
      context = await adapter.context({ scope: 'slide', images: false });
    } catch (error) {
      this.report(error, 'pending-context');
      return;
    }
    if (this.stopped) return;
    const scope = `${context.presentationId || ''}|${context.slideId || ''}`;
    const existing = new Set((context.shapes || []).map(shape => String(shape.id)));
    const selected = new Set((context.selectionIds || []).map(String));
    const handled = this.handledIds || (this.handledIds = new Set());
    const handledKey = id => `${scope}|${id}`;
    const candidates = this.pending.slice().sort((a, b) => b.at - a.at);
    for (const item of candidates) {
      if (!item.baselineReady) continue;
      const snapshot = await item.baselineReady;
      item.baselineReady = null;
      if (!item.intercepted) item.baseline = snapshot && snapshot.scope === scope ? snapshot.ids : null;
    }
    for (const item of candidates) {
      const baseline = item.baseline || new Set();
      const fresh = (context.shapes || []).filter(shape => {
        const id = String(shape.id);
        if (baseline.has(id) || handled.has(handledKey(id))) return false;
        if (shape.type !== 'picture' && shape.type !== 'media') return false;
        if (hasSourceRecord(shape)) return false;
        if (selected.has(id) || item.intercepted) return true;
        return !item.baseline && now - item.at < 3000;
      });
      if (!fresh.length) continue;
      this.pending = this.pending.filter(entry => entry !== item);
      for (const shape of fresh) handled.add(handledKey(String(shape.id)));
      this.queue = this.queue.then(() => this.applySource(item, fresh)).catch(error => {
        this.status('更新粘贴图片的来源失败：' + messageOf(error), 'error');
        this.report(error, 'native-paste');
      });
    }
    this.knownIds = existing;
    if (handled.size > 300) this.handledIds = new Set([...handled].slice(-150));
  }

  async applySource(candidate, fresh) {
    if (this.stopped) return;
    const adapter = await this.hostAdapter();
    if (this.stopped) return;
    const files = candidate.files.length
      ? candidate.files
      : [await this.materializeClipboard(candidate)];
    const limit = Math.min(files.length, fresh.length);
    if (!limit) return;
    const shapeIds = fresh.slice(0, limit).map(shape => shape.id);
    const ctx = {
      host: adapter,
      services: this.services,
      settings: this.services.settings,
      signal: this.controller.signal,
      notify: (message, kind = 'info') => this.notify(message, kind),
      refresh: () => adapter.summary()
    };
    await handleNativePaste(ctx, { files: files.slice(0, limit), position: candidate.position, shapeIds });
    this.onEvent?.({ ...candidate, shapeIds });
    this.status(`已记录 ${limit} 个粘贴图片的来源。`, 'success');
  }

  async materializeClipboard(candidate) {
    if (typeof candidate.clipboard?.data !== 'string') throw new Error('剪贴板图片数据不可用。');
    return {
      name: '剪贴板图片.tiff',
      mime: candidate.clipboard.mime || 'image/tiff',
      data: candidate.clipboard.data
    };
  }

  hostAdapter() {
    if (this.injectedAdapter) return Promise.resolve(this.injectedAdapter);
    if (!this.adapterPromise) {
      this.adapterPromise = (async () => {
        let adapter;
        if (this.host === 'wps') {
          const module = await import('./wps.js');
          adapter = module.connectWps(this.services);
        } else {
          const module = await import('./powerpoint.js');
          adapter = await module.connectPowerPoint(this.services);
        }
        adapter = await adapter;
        for (const name of ['context', 'apply', 'images', 'onChanged']) {
          if (typeof adapter?.[name] !== 'function') throw new Error('演示软件连接不完整，无法跟踪粘贴图片。请更新 SlideSCI 后重试。');
        }
        return adapter;
      })().catch(error => {
        this.adapterPromise = null;
        throw error;
      });
    }
    return this.adapterPromise;
  }

  async startScience() {
    if (this.stopped || this.scienceCleanup) return;
    clearTimeout(this.scienceRetryTimer);
    this.scienceRetryTimer = null;
    this.lastScienceAttempt = Date.now();
    try {
      const adapter = await this.hostAdapter();
      const ctx = {
        host: adapter,
        services: this.services,
        settings: this.services.settings,
        signal: this.controller.signal,
        notify: (message, kind = 'info') => this.notify(message, kind),
        refresh: async () => { await adapter.context({ scope: 'selection' }); }
      };
      const cleanup = await initializeScienceFeature(ctx);
      if (this.stopped) {
        if (typeof cleanup === 'function') cleanup();
        return;
      }
      this.scienceCleanup = typeof cleanup === 'function' ? cleanup : null;
      this.scienceFailed = false;
      this.scienceAttempts = 0;
    } catch (error) {
      this.scienceFailed = true;
      this.status('科学图片自动同步初始化失败：' + messageOf(error), 'error');
      this.report(error, 'science-initialize');
      this.retryScience();
    }
  }

  retryScience(immediate = false) {
    if (this.stopped || this.scienceCleanup || this.scienceRetryTimer) return;
    const attempt = this.scienceAttempts++;
    const delay = immediate ? 0 : attempt < 8 ? Math.min(15000, 1500 * 2 ** attempt) : 60000;
    this.scienceRetryTimer = setTimeout(() => {
      this.scienceRetryTimer = null;
      this.startScience().catch(error => console.warn('SlideSCI 科学图片同步重试失败：' + messageOf(error)));
    }, delay);
  }

  async stop({ keepalive = false } = {}) {
    if (this.stopped) return;
    this.stopped = true;
    this.started = false;
    clearTimeout(this.pendingPoll);
    this.pendingPoll = null;
    clearTimeout(this.scienceRetryTimer);
    this.scienceRetryTimer = null;
    this.controller.abort();
    this.unsubscribeEvents();
    try {
      this.hostUnsubscribe?.();
    } catch (error) {
      console.warn('SlideSCI 宿主变更监听清理失败：' + messageOf(error));
    }
    this.hostUnsubscribe = null;
    try {
      this.scienceCleanup?.();
    } catch (error) {
      console.warn('SlideSCI 科学图片同步清理失败：' + messageOf(error));
    }
    this.scienceCleanup = null;
    const leaseId = this.leaseId;
    this.leaseId = null;
    if (leaseId) {
      if (keepalive) this.services.stopWatcher(leaseId).catch(() => {});
      else await this.services.stopWatcher(leaseId);
    }
    runtimes.delete(this.host);
  }
}

function hasSourceRecord(shape) {
  const tags = shape.tags || {};
  if (Object.keys(tags).some(name => /SLIDESCI_SOURCE_PATH|SLIDESCI_IMAGE_UUID/i.test(name))) return true;
  return /^(?:原图路径|视频路径)：/m.test(String(shape.alternativeText || ''));
}

export async function startBackgroundTracking(options) {
  const { host, services } = options || {};
  if (!['powerpoint', 'wps'].includes(host)) throw new Error('后台跟踪需要指定演示软件类型。');
  if (!services) throw new Error('后台跟踪需要本地服务。');
  const existing = runtimes.get(host);
  if (existing && !existing.stopped) return existing;
  const runtime = new BackgroundTracking({ ...options, host, services });
  runtimes.set(host, runtime);
  try {
    await runtime.start();
  } catch (error) {
    runtimes.delete(host);
    throw error;
  }
  return runtime;
}

export const startup = startBackgroundTracking;

export function backgroundStatus(host) {
  const runtime = runtimes.get(host);
  return runtime && !runtime.stopped
    ? { host, role: runtime.role, watcher: Boolean(runtime.leaseId), pending: runtime.pending.length, science: Boolean(runtime.scienceCleanup), scienceFailed: Boolean(runtime.scienceFailed) }
    : { host, role: null, watcher: false, pending: 0, science: false, scienceFailed: false };
}

export async function stopBackgroundTracking(host) {
  const runtime = runtimes.get(host);
  if (runtime) await runtime.stop();
}
