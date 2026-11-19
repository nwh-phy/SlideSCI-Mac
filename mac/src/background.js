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
    this.leasePromise = null;
    this.stopped = false;
    this.queue = Promise.resolve();
    this.unsubscribers = [];
    this.controller = new AbortController();
    this.pending = [];
    this.knownByScope = new Map();
    this.handledIds = new Set();
    this.checkPromise = null;
    this.checkAgain = false;
    this.applyingDepth = 0;
    this.syncingKnown = false;
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
    // A missing service is not fatal: the event stream keeps retrying, and its hello
    // re-acquires the watcher lease once the service is (re)started.
    try {
      await this.acquireLease();
    } catch (error) {
      this.report(error, 'background-start');
      this.status('自动粘贴监听暂未启动：' + messageOf(error) + '。本地服务恢复后会自动重试。', 'error');
    }
    if (this.shouldInitializeScience) await this.startScience();
    try {
      const adapter = await this.hostAdapter();
      const context = await adapter.context({ scope: 'slide', geometry: true });
      this.recordKnown(context);
      if (typeof adapter.onChanged === 'function') {
        this.hostUnsubscribe = adapter.onChanged(change => {
          if (this.stopped || this.applyingDepth > 0) return;
          if (this.pending.length) {
            this.checkPending().catch(error => this.report(error, 'pending-host-change'));
          } else if (change?.kind !== 'selection') {
            this.refreshKnown().catch(error => this.report(error, 'shape-baseline'));
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
    this.unsubscribers.push(this.services.subscribe('hello', payload => this.hello(payload)));
  }

  acquireLease() {
    if (this.leasePromise) return this.leasePromise;
    this.leasePromise = this.services.startWatcher({ interceptTIFF: this.interceptTIFF })
      .then(watcher => {
        if (!watcher?.id) throw new Error('原生媒体监听没有返回有效租约。');
        if (this.stopped) {
          this.services.stopWatcher(watcher.id).catch(() => {});
          return null;
        }
        this.leaseId = watcher.id;
        return watcher;
      })
      .finally(() => { this.leasePromise = null; });
    return this.leasePromise;
  }

  // Every (re)connected event stream reports the leases the service still holds for this pane.
  // A restarted service or an expired grace period leaves ours out, so take a new one.
  hello(payload) {
    if (this.stopped) return;
    const leases = Array.isArray(payload?.leases) ? payload.leases : [];
    if (this.leaseId && leases.includes(this.leaseId)) return;
    const recovering = !this.leaseId;
    this.acquireLease()
      .then(watcher => { if (watcher && recovering) this.status('自动粘贴监听已恢复。', 'success'); })
      .catch(error => this.report(error, 'background-reacquire'));
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

  scopeOf(context) {
    return `${context?.presentationId || ''}|${context?.slideId || ''}`;
  }

  recordKnown(context) {
    if (!context) return new Set();
    const scope = this.scopeOf(context);
    const ids = new Set((context.shapes || []).map(shape => String(shape.id)));
    this.knownByScope.delete(scope);
    this.knownByScope.set(scope, ids);
    if (this.knownByScope.size > 32) this.knownByScope.delete(this.knownByScope.keys().next().value);
    return ids;
  }

  async refreshKnown() {
    if (this.stopped || this.syncingKnown || this.applyingDepth > 0) return;
    this.syncingKnown = true;
    try {
      const adapter = await this.hostAdapter();
      const context = await adapter.context({ scope: 'slide', geometry: true });
      if (!this.stopped && this.applyingDepth === 0) this.recordKnown(context);
    } finally {
      this.syncingKnown = false;
    }
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
    const reason = payload.reason || '';
    const channel = reason === 'drag' || reason === 'drop' ? 'drag' : 'clipboard';
    const now = Date.now();
    // Each pasteboard holds one current value; a newer clipboard or drag capture supersedes the old one.
    this.pending = this.pending.filter(item => now - item.at < PENDING_LIFETIME && item.signature !== signature && item.channel !== channel);
    const baselines = new Map([...this.knownByScope].map(([scope, ids]) => [scope, new Set(ids)]));
    const candidate = {
      signature,
      channel,
      at: now,
      baselines,
      baselineReady: this.captureBaseline(),
      files,
      clipboard: clipboardImage ? { data: payload.data, mime: payload.mime || 'image/tiff' } : null,
      intercepted: payload.intercepted === true,
      reason,
      position: validPosition(payload.position)
    };
    this.pending.push(candidate);
    if (this.pending.length > 4) this.pending.shift();
    // The service grants watcher-reported files before broadcasting them.
    this.beginPendingObservation();
  }

  captureBaseline() {
    return this.hostAdapter()
      .then(adapter => adapter.context({ scope: 'slide', geometry: true }))
      .then(context => ({
        ids: new Set((context.shapes || []).map(shape => String(shape.id))),
        scope: this.scopeOf(context)
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

  checkPending() {
    if (this.stopped || !this.pending.length) return Promise.resolve();
    if (this.applyingDepth > 0) {
      this.checkAgain = true;
      return this.queue;
    }
    if (this.checkPromise) {
      this.checkAgain = true;
      return this.checkPromise;
    }
    this.checkPromise = (async () => {
      try {
        do {
          this.checkAgain = false;
          if (this.stopped || this.applyingDepth > 0 || !this.pending.length) break;
          await this.runPendingCheck();
        } while (!this.stopped && this.applyingDepth === 0 && this.checkAgain && this.pending.length);
      } finally {
        this.checkPromise = null;
      }
    })();
    return this.checkPromise;
  }

  async runPendingCheck() {
    const now = Date.now();
    this.pending = this.pending.filter(item => now - item.at < PENDING_LIFETIME);
    if (!this.pending.length) return;
    let adapter, light;
    try {
      adapter = await this.hostAdapter();
      // Poll with lightweight geometry first; only export full slide metadata when a new picture/media appears.
      light = await adapter.context({ scope: 'slide', geometry: true });
    } catch (error) {
      this.report(error, 'pending-context');
      return;
    }
    if (this.stopped || this.applyingDepth > 0) return;
    const scope = this.scopeOf(light);
    const existing = new Set((light.shapes || []).map(shape => String(shape.id)));
    const handled = this.handledIds || (this.handledIds = new Set());
    const handledKey = id => `${scope}|${id}`;
    const candidates = this.pending.slice().sort((a, b) => b.at - a.at);
    for (const item of candidates) {
      if (!item.baselineReady) continue;
      const snapshot = await item.baselineReady;
      item.baselineReady = null;
      if (!snapshot) continue;
      // On drop, the picture may already have landed; keep any pre-drop baseline recorded for this slide.
      if (!(item.reason === 'drop' && item.baselines.has(snapshot.scope))) {
        item.baselines.set(snapshot.scope, snapshot.ids);
      }
      if (!this.knownByScope.has(snapshot.scope)) this.knownByScope.set(snapshot.scope, new Set(snapshot.ids));
    }
    let rich = null;
    let scheduled = false;
    let deferred = false;
    for (const item of candidates) {
      let baseline = item.baselines.get(scope);
      if (!baseline) {
        if (this.knownByScope.has(scope)) {
          baseline = new Set(this.knownByScope.get(scope));
        } else {
          // Never treat pre-existing pictures on a newly visited slide as freshly pasted.
          baseline = new Set(existing);
        }
        item.baselines.set(scope, baseline);
      }
      const potential = (light.shapes || []).filter(shape => {
        const id = String(shape.id);
        if (baseline.has(id) || handled.has(handledKey(id))) return false;
        return shape.type === 'picture' || shape.type === 'media';
      });
      if (!potential.length) continue;
      // Multi-file pastes (e.g., 8 images from Finder) can land across several hundred ms in
      // PowerPoint; wait briefly until all expected shapes appear or the count settles.
      const expected = item.files.length || 1;
      if (expected > 1 && potential.length < expected) {
        const prev = item.partial;
        if (!prev || prev.scope !== scope || prev.count !== potential.length) {
          item.partial = { scope, count: potential.length, at: now };
          deferred = true;
          this.beginPendingObservation();
          break;
        }
        if (now - prev.at < 800) {
          deferred = true;
          this.beginPendingObservation();
          break;
        }
      }
      item.partial = null;
      if (!rich) {
        try {
          rich = await adapter.context({ scope: 'slide', images: false });
        } catch (error) {
          this.report(error, 'pending-context');
          return;
        }
        if (this.stopped || this.applyingDepth > 0) return;
        if (this.scopeOf(rich) !== scope) return;
      }
      const potentialIds = new Set(potential.map(shape => String(shape.id)));
      const fresh = [];
      for (const shape of rich.shapes || []) {
        const id = String(shape.id);
        if (!potentialIds.has(id)) continue;
        if (hasSourceRecord(shape)) {
          baseline.add(id);
          handled.add(handledKey(id));
          continue;
        }
        fresh.push(shape);
      }
      if (!fresh.length) continue;
      this.pending = this.pending.filter(entry => entry !== item);
      for (const shape of fresh) {
        const id = String(shape.id);
        handled.add(handledKey(id));
        baseline.add(id);
      }
      this.applyingDepth++;
      scheduled = true;
      this.queue = this.queue.then(() => this.applySource(item, fresh)).catch(error => {
        this.status('更新粘贴图片的来源失败：' + messageOf(error), 'error');
        this.report(error, 'native-paste');
      }).finally(() => {
        this.applyingDepth = Math.max(0, this.applyingDepth - 1);
        if (!this.stopped && this.checkAgain && this.pending.length) {
          this.checkPending().catch(error => this.report(error, 'pending-after-apply'));
        }
      });
      break;
    }
    if (!scheduled && !deferred) this.recordKnown(light);
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
    const result = await handleNativePaste(ctx, { files: files.slice(0, limit), position: candidate.position, shapeIds });
    try {
      const after = await adapter.context({ scope: 'slide', geometry: true });
      const scope = this.scopeOf(after);
      const ids = this.recordKnown(after);
      for (const item of this.pending) {
        if (item.baselines?.has(scope)) {
          const set = item.baselines.get(scope);
          for (const id of ids) set.add(id);
        }
      }
    } catch { /* best effort baseline refresh after apply */ }
    if (result?.count) {
      this.onEvent?.({ ...candidate, shapeIds });
      this.status(`已记录 ${result.count} 个粘贴图片的来源。`, 'success');
    }
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
