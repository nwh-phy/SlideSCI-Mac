import test from 'node:test';
import assert from 'node:assert/strict';
import { startBackgroundTracking, stopBackgroundTracking } from '../src/background.js';
import { insertSourceFiles } from '../src/features/images.js';
import { shouldSuppressWatchEvent } from '../lib/native.mjs';

function fakeServices({ normalize } = {}) {
  const listeners = new Map();
  return {
    normalizeCalls: 0,
    settings: { get: (_, fallback) => fallback, set() {} },
    subscribe(name, callback) {
      if (!listeners.has(name)) listeners.set(name, new Set());
      listeners.get(name).add(callback);
      return () => listeners.get(name)?.delete(callback);
    },
    emit(name, payload) {
      for (const callback of listeners.get(name) || []) callback(payload);
    },
    async startWatcher() { return { id: 'lease-1' }; },
    async stopWatcher() {},
    async reportError() {},
    async request(route, input) {
      if (route === 'images/normalize') {
        this.normalizeCalls++;
        if (normalize) return normalize(input);
        return { data: 'AAAA', mime: 'image/png', width: 100, height: 80, fov: null, highBitDepth: false };
      }
      throw new Error('Unexpected route: ' + route);
    }
  };
}

const delay = ms => new Promise(resolve => setTimeout(resolve, ms));

test('concurrent pending checks never duplicate pasted TIFF images when replaceImage swaps shape IDs', async () => {
  const services = fakeServices({
    normalize: () => ({ data: 'AAAA', mime: 'image/png', width: 100, height: 80, fov: null, highBitDepth: true })
  });
  let nextId = 10;
  let slide = {
    presentationId: 'deck-1',
    slideId: 'slide-1',
    width: 960,
    height: 540,
    selectionIds: [],
    shapes: []
  };
  const ops = [];
  const changeListeners = new Set();
  const adapter = {
    kind: 'powerpoint',
    async context() {
      await delay(30);
      return structuredClone(slide);
    },
    async summary() { return structuredClone(slide); },
    async apply(commands) {
      for (const command of commands) {
        ops.push(command.op);
        if (command.op === 'update') {
          const shape = slide.shapes.find(item => item.id === command.id);
          if (shape) {
            Object.assign(shape, command.properties, { tags: { ...(shape.tags || {}), ...(command.properties.tags || {}) } });
          }
        } else if (command.op === 'replaceImage') {
          const index = slide.shapes.findIndex(item => item.id === command.id);
          assert.ok(index >= 0, 'replaceImage target must exist');
          const old = slide.shapes[index];
          const replaced = { ...old, id: 'replaced-' + ++nextId };
          slide.shapes.splice(index, 1, replaced);
          slide.selectionIds = [replaced.id];
        } else if (command.op === 'insertImage') {
          slide.shapes.push({ id: 'inserted-' + ++nextId, type: 'picture', name: command.name, tags: command.tags || {} });
        } else if (command.op === 'select') {
          slide.selectionIds = [...command.ids];
        }
      }
      return structuredClone(slide);
    },
    onChanged(cb) { changeListeners.add(cb); return () => changeListeners.delete(cb); }
  };

  const runtime = await startBackgroundTracking({ host: 'powerpoint', services, hostAdapter: adapter });
  try {
    // 1. User copies a TIFF file in Finder; baseline captures empty slide-1.
    services.emit('mediaSource', {
      host: 'com.microsoft.Powerpoint',
      reason: 'clipboard',
      files: [{ path: '/tmp/sample.tif', name: 'sample.tif' }]
    });
    await delay(50);

    // 2. User pastes the TIFF into slide-1; host + mouseUp + onChanged fire concurrently.
    slide.shapes.push({ id: 'pasted-1', type: 'picture', name: 'Picture 1', tags: {}, alternativeText: '' });
    slide.selectionIds = ['pasted-1'];
    services.emit('changed', { host: 'powerpoint' });
    services.emit('changed', { host: 'powerpoint' });
    for (const cb of changeListeners) cb({ kind: 'content' });
    await Promise.all([runtime.checkPending(), runtime.checkPending(), runtime.checkPending()]);
    await runtime.queue;
    await delay(80);

    assert.equal(slide.shapes.length, 1, 'slide must still contain exactly one picture');
    assert.equal(ops.filter(op => op === 'insertImage').length, 0, 'must never fall back to insertImage duplication');
    assert.equal(ops.filter(op => op === 'replaceImage').length, 1, 'TIFF replacement runs exactly once');
  } finally {
    await stopBackgroundTracking('powerpoint');
  }
});

test('switching to the next slide never treats pre-existing selected pictures as freshly pasted', async () => {
  const services = fakeServices();
  const slides = {
    'slide-1': {
      presentationId: 'deck-1',
      slideId: 'slide-1',
      width: 960,
      height: 540,
      selectionIds: ['s1-a'],
      shapes: [{ id: 's1-a', type: 'picture', name: 'A', tags: {}, alternativeText: '' }]
    },
    'slide-2': {
      presentationId: 'deck-1',
      slideId: 'slide-2',
      width: 960,
      height: 540,
      selectionIds: ['s2-a', 's2-b', 's2-c'],
      shapes: [
        { id: 's2-a', type: 'picture', name: 'Picture 1', tags: {}, alternativeText: '' },
        { id: 's2-b', type: 'picture', name: 'Picture 2', tags: {}, alternativeText: '' },
        { id: 's2-c', type: 'picture', name: 'Picture 3', tags: {}, alternativeText: '' }
      ]
    }
  };
  let current = 'slide-1';
  let richReads = 0;
  const applied = [];
  const adapter = {
    kind: 'powerpoint',
    async context(options = {}) {
      if (!options.geometry) richReads++;
      return structuredClone(slides[current]);
    },
    async summary() { return structuredClone(slides[current]); },
    async apply(commands) {
      applied.push(...commands);
      for (const command of commands) {
        if (command.op === 'update') {
          const shape = slides[current].shapes.find(item => item.id === command.id);
          if (shape) Object.assign(shape, command.properties, { tags: { ...(shape.tags || {}), ...(command.properties.tags || {}) } });
        }
      }
      return structuredClone(slides[current]);
    },
    onChanged() { return () => {}; }
  };

  const runtime = await startBackgroundTracking({ host: 'powerpoint', services, hostAdapter: adapter });
  try {
    // Pending clipboard capture from slide 1.
    services.emit('mediaSource', {
      host: 'com.microsoft.Powerpoint',
      reason: 'clipboard',
      files: [{ path: '/tmp/fig1.png', name: 'fig1.png' }]
    });
    await delay(30);

    // User switches to slide 2 where 3 existing unrecorded figures are selected for arranging.
    current = 'slide-2';
    await runtime.checkPending();
    await runtime.queue;

    assert.equal(applied.length, 0, 'pre-existing figures on slide 2 must not be modified or replaced');
    assert.equal(richReads, 0, 'no expensive full-slide export should run when nothing new was pasted');

    // Now user actually pastes a new picture onto slide 2.
    slides['slide-2'].shapes.push({ id: 's2-new', type: 'picture', name: 'Picture 4', tags: {}, alternativeText: '' });
    slides['slide-2'].selectionIds = ['s2-new'];
    await runtime.checkPending();
    await runtime.queue;

    assert.equal(applied.filter(c => c.op === 'insertImage').length, 0);
    assert.equal(applied.filter(c => c.op === 'replaceImage').length, 0, 'plain PNG paste updates metadata without replaceImage');
    assert.equal(applied.filter(c => c.op === 'update' && c.id === 's2-new').length, 1);
    assert.equal(services.normalizeCalls, 0, 'plain PNG metadata update skips unnecessary image decode');
  } finally {
    await stopBackgroundTracking('powerpoint');
  }
});

test('insertSourceFiles in update mode skips missing shapeIds instead of inserting duplicate copies', async () => {
  const services = fakeServices();
  const applied = [];
  const ctx = {
    host: {
      kind: 'powerpoint',
      async context() {
        return { width: 960, height: 540, shapes: [{ id: 'other', type: 'picture', name: 'Other', tags: {}, alternativeText: '' }] };
      },
      async apply(commands) { applied.push(...commands); return { width: 960, height: 540, shapes: [] }; }
    },
    services
  };
  const result = await insertSourceFiles(ctx, [{ path: '/tmp/a.png', name: 'a.png', mime: 'image/png' }], { shapeIds: ['already-replaced'] });
  assert.deepEqual(result, { count: 0 });
  assert.equal(applied.length, 0);
  assert.equal(services.normalizeCalls, 0);
});

test('shouldSuppressWatchEvent blocks delayed internal clipboard restore echoes and allows later user copies', () => {
  // Active internal clipboard operation suppresses hostChanged and pasteboard events.
  assert.equal(shouldSuppressWatchEvent({ type: 'mediaSource', reason: 'clipboard', changeCount: 50 }, { depth: 1, ignoredCount: 40 }), true);
  assert.equal(shouldSuppressWatchEvent({ type: 'hostChanged' }, { depth: 1, ignoredCount: 40 }), true);
  // Errors always surface.
  assert.equal(shouldSuppressWatchEvent({ type: 'nativeError', message: 'boom' }, { depth: 1, ignoredCount: 40 }), false);
  // Delayed 200ms watcher poll seeing the changeCount written by restoreClipboard is suppressed even after depth returns to 0.
  assert.equal(shouldSuppressWatchEvent({ type: 'mediaSource', reason: 'clipboard', changeCount: 42, pasteboard: 'Apple CFPasteboard general' }, { depth: 0, ignoredCount: 42 }), true);
  assert.equal(shouldSuppressWatchEvent({ type: 'clipboardImage', reason: 'clipboard', changeCount: 41, pasteboard: 'Apple CFPasteboard general' }, { depth: 0, ignoredCount: 42 }), true);
  // A subsequent real user Cmd+C increments changeCount beyond ignoredCount and passes through.
  assert.equal(shouldSuppressWatchEvent({ type: 'mediaSource', reason: 'clipboard', changeCount: 43, pasteboard: 'Apple CFPasteboard general' }, { depth: 0, ignoredCount: 42 }), false);
  // Drag pasteboard has its own counter and is not blocked by general pasteboard's ignoredCount.
  assert.equal(shouldSuppressWatchEvent({ type: 'mediaSource', reason: 'drag', changeCount: 5, pasteboard: 'Apple CFPasteboard drag' }, { depth: 0, ignoredCount: 42 }), false);
});

test('pasting 8 images from Finder on slide 1 then 8 more on slide 2 waits for full landing and batches metadata in one call per slide', async () => {
  const services = fakeServices();
  const slides = {
    'slide-1': { presentationId: 'deck-1', slideId: 'slide-1', width: 960, height: 540, selectionIds: [], shapes: [] },
    'slide-2': { presentationId: 'deck-1', slideId: 'slide-2', width: 960, height: 540, selectionIds: [], shapes: [] }
  };
  let current = 'slide-1';
  const applyBatches = [];
  const adapter = {
    kind: 'powerpoint',
    async context() { return structuredClone(slides[current]); },
    async summary() { return structuredClone(slides[current]); },
    async apply(commands, options) {
      applyBatches.push({ slide: current, commands: structuredClone(commands), options });
      for (const command of commands) {
        if (command.op === 'update') {
          const shape = slides[current].shapes.find(item => item.id === command.id);
          if (shape) Object.assign(shape, command.properties, { tags: { ...(shape.tags || {}), ...(command.properties.tags || {}) } });
        }
      }
      return structuredClone(slides[current]);
    },
    onChanged() { return () => {}; }
  };

  const runtime = await startBackgroundTracking({ host: 'powerpoint', services, hostAdapter: adapter });
  try {
    const batch1 = Array.from({ length: 8 }, (_, i) => ({ path: `/tmp/s1-${i + 1}.png`, name: `s1-${i + 1}.png` }));
    services.emit('mediaSource', { host: 'com.microsoft.Powerpoint', reason: 'clipboard', files: batch1 });
    await delay(20);

    // PowerPoint lands 3 of the 8 pictures first, then the remaining 5.
    for (let i = 1; i <= 3; i++) slides['slide-1'].shapes.push({ id: `p1-${i}`, type: 'picture', name: `Picture ${i}`, tags: {}, alternativeText: '' });
    await runtime.checkPending();
    assert.equal(applyBatches.length, 0, 'waits while multi-file paste is still landing');

    for (let i = 4; i <= 8; i++) slides['slide-1'].shapes.push({ id: `p1-${i}`, type: 'picture', name: `Picture ${i}`, tags: {}, alternativeText: '' });
    await runtime.checkPending();
    await runtime.queue;

    assert.equal(slides['slide-1'].shapes.length, 8);
    assert.equal(applyBatches.length, 1, 'all 8 PNG metadata updates execute in one lightweight apply batch');
    assert.equal(applyBatches[0].commands.length, 8);
    assert.equal(applyBatches[0].options?.readback, false);
    assert.ok(applyBatches[0].commands.every(c => c.op === 'update'));

    // Switch to slide 2 and paste 8 different images from Finder.
    current = 'slide-2';
    const batch2 = Array.from({ length: 8 }, (_, i) => ({ path: `/tmp/s2-${i + 1}.png`, name: `s2-${i + 1}.png` }));
    services.emit('mediaSource', { host: 'com.microsoft.Powerpoint', reason: 'clipboard', files: batch2 });
    await delay(20);
    for (let i = 1; i <= 8; i++) slides['slide-2'].shapes.push({ id: `p2-${i}`, type: 'picture', name: `Picture ${i}`, tags: {}, alternativeText: '' });
    await Promise.all([runtime.checkPending(), runtime.checkPending()]);
    await runtime.queue;

    assert.equal(slides['slide-2'].shapes.length, 8);
    assert.equal(applyBatches.length, 2, 'slide 2 batch also runs once with zero duplication');
    assert.equal(applyBatches[1].slide, 'slide-2');
    assert.equal(applyBatches[1].commands.length, 8);
    assert.equal(services.normalizeCalls, 0);
  } finally {
    await stopBackgroundTracking('powerpoint');
  }
});
