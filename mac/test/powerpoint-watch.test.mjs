import test from 'node:test';
import assert from 'node:assert/strict';
import { connectPowerPoint } from '../src/powerpoint.js';

const loadable = value => Object.assign(value, { load() { return this; } });
function collection(items) {
  return loadable({ items, getCount() { return { value: items.length }; }, getItemAt(index) { return items[index]; },
    getItem(id) { return items.find(item => String(item.id) === String(id)); } });
}
function shapeTags() {
  const tags = collection([]);
  tags.add = (key, value) => { tags.delete(key); tags.items.push({ key, value }); };
  tags.delete = key => { const index = tags.items.findIndex(item => item.key === key); if (index >= 0) tags.items.splice(index, 1); };
  return tags;
}

// A text shape exposes textFrame.textRange.text; the watch must read that, not a TextRange.textFrame.
test('powerpoint change watch detects text edits on slides with text shapes and tolerates % in file names', { timeout: 15000 }, async t => {
  const originals = { Office: globalThis.Office, PowerPoint: globalThis.PowerPoint };
  t.after(() => { for (const [name, value] of Object.entries(originals)) { if (value === undefined) delete globalThis[name]; else globalThis[name] = value; } });
  const range = loadable({ text: 'before' });
  const shapes = [loadable({ id: 't1', name: 'title', left: 10, top: 10, width: 200, height: 40, rotation: 0, type: 'TextBox', zOrderPosition: 1,
    textFrame: { textRange: range } })];
  const slides = collection([loadable({ id: 'slide1', index: 0, shapes: collection(shapes) })]);
  let selection = [];
  const presentation = loadable({ id: 'deck', slides, pageSetup: loadable({ slideWidth: 720, slideHeight: 540 }),
    getSelectedSlides() { return collection(slides.items); }, getSelectedShapes() { return collection(selection); } });
  globalThis.Office = { HostType: { PowerPoint: 'PowerPoint' }, onReady: async () => ({ host: 'PowerPoint' }),
    context: { document: { url: '/Users/someone/增长 20%.pptx' }, requirements: { isSetSupported: () => true } } };
  globalThis.PowerPoint = { run: async operation => operation({ presentation, sync: async () => {} }) };
  const reported = [];
  const services = {
    async request(route, input) {
      if (route === 'native/execute' && input.action === 'documentInfo') return { name: '增长 20%.pptx' };
      throw new Error(`unexpected ${route}`);
    },
    async reportError(error) { reported.push(error); }
  };
  const host = await connectPowerPoint(services);
  t.after(() => host.dispose());
  assert.equal((await host.summary({ slides: false })).title, '增长 20%', 'a % in the file name is not treated as URI encoding');
  const changes = [];
  host.onChanged(change => { changes.push(change?.kind); });
  await new Promise(resolve => setTimeout(resolve, 1700));
  assert.deepEqual(reported, [], 'the first fingerprint succeeds on a slide with a text shape');
  range.text = 'after';
  await new Promise(resolve => setTimeout(resolve, 1700));
  assert.deepEqual(changes, ['content'], 'editing the text is reported as a content change');
  selection = [shapes[0]];
  await new Promise(resolve => setTimeout(resolve, 1700));
  assert.deepEqual(changes, ['content', 'selection'], 'a click that only changes the selection says so');
  assert.deepEqual(reported, []);
});

// Connect and background watch must stay pure Office.js: zero AppleScript and zero background PPTX exports.
test('powerpoint connect and watch stay pure Office.js without AppleScript or background PPTX exports', { timeout: 20000 }, async t => {
  const originals = { Office: globalThis.Office, PowerPoint: globalThis.PowerPoint };
  t.after(() => { for (const [name, value] of Object.entries(originals)) { if (value === undefined) delete globalThis[name]; else globalThis[name] = value; } });
  let exports = 0;
  const shapes = [loadable({ id: 'p1', name: 'picture', left: 10, top: 10, width: 200, height: 40, rotation: 0, type: 'Image', zOrderPosition: 1 })];
  const slides = collection([loadable({ id: 'slide1', index: 0, shapes: collection(shapes), exportAsBase64() {} })]);
  slides.exportAsBase64Presentation = () => { exports++; return { value: 'UEsDBA==' }; };
  const presentation = loadable({ id: 'deck', slides, pageSetup: loadable({ slideWidth: 720, slideHeight: 540 }),
    getSelectedSlides() { return collection(slides.items); }, getSelectedShapes() { return collection([]); } });
  globalThis.Office = { HostType: { PowerPoint: 'PowerPoint' }, onReady: async () => ({ host: 'PowerPoint' }),
    context: { document: { url: '/Users/someone/deck.pptx' }, requirements: { isSetSupported: () => true } } };
  globalThis.PowerPoint = { run: async operation => operation({ presentation, sync: async () => {} }) };
  const requests = [], reported = [];
  const services = {
    async request(route, input) { requests.push([route, input]); throw new Error(`unexpected ${route}`); },
    async reportError(error) { reported.push(error); }
  };
  const host = await connectPowerPoint(services);
  t.after(() => host.dispose());
  host.onChanged(() => {});
  await new Promise(resolve => setTimeout(resolve, 1500 * 5 + 300));
  assert.deepEqual(reported, []);
  assert.deepEqual(requests, [], 'connect and polling never call native AppleScript or backend routes');
  assert.equal(exports, 0, 'polling never runs background PPTX slide exports');
});

test('core figure labels, titles and selectAllText run in pure Office.js without backend requests', async t => {
  const originals = { Office: globalThis.Office, PowerPoint: globalThis.PowerPoint };
  t.after(() => { for (const [name, value] of Object.entries(originals)) { if (value === undefined) delete globalThis[name]; else globalThis[name] = value; } });
  const { actions } = await import('../src/features/editing.js');
  let nextId = 10;
  const makeShape = ({ id, name, type, left, top, width, height, text = '' }) => {
    const font = { name: 'Arial', size: 12, bold: false };
    const paragraphFormat = { horizontalAlignment: 'Left' };
    const textRange = loadable({ text, font, paragraphFormat });
    const textFrame = loadable({ textRange, wordWrap: false, autoSizeSetting: 'AutoSizeNone', hasText: Boolean(text) });
    return loadable({ id, name, type, left, top, width, height, rotation: 0, zOrderPosition: 1, textFrame,
      tags: shapeTags(), delete() { const i = shapes.findIndex(s => s.id === id); if (i >= 0) shapes.splice(i, 1); } });
  };
  const shapes = [
    makeShape({ id: 'img-1', name: 'Fig1', type: 'Image', left: 20, top: 30, width: 120, height: 90 }),
    makeShape({ id: 'img-2', name: 'Fig2', type: 'Image', left: 160, top: 30, width: 120, height: 90 })
  ];
  let selectedIds = ['img-1', 'img-2'];
  const shapeCollection = Object.assign(collection(shapes), {
    addTextBox(text, options) {
      const box = makeShape({ id: `tb-${++nextId}`, name: 'TextBox', type: 'TextBox',
        left: options.left, top: options.top, width: options.width, height: options.height, text });
      shapes.push(box);
      return box;
    }
  });
  const slide = loadable({ id: 'slide1', index: 0, shapes: shapeCollection, setSelectedShapes(ids) { selectedIds = [...ids]; } });
  const slides = collection([slide]);
  const presentation = loadable({ id: 'deck', slides, pageSetup: loadable({ slideWidth: 720, slideHeight: 540 }),
    getSelectedSlides() { return collection(slides.items); },
    getSelectedShapes() { return collection(selectedIds.map(id => shapes.find(s => s.id === id)).filter(Boolean)); } });
  globalThis.Office = { HostType: { PowerPoint: 'PowerPoint' }, onReady: async () => ({ host: 'PowerPoint' }),
    context: { document: { url: '/Users/someone/figures.pptx' }, requirements: { isSetSupported: () => true } } };
  globalThis.PowerPoint = { run: async operation => operation({ presentation, sync: async () => {} }) };

  let backendCalls = 0;
  const services = { async request() { backendCalls++; throw new Error('backend should not be called'); }, async reportError() {} };
  const host = await connectPowerPoint(services);
  t.after(() => host.dispose());
  const settingsStore = new Map();
  const ctx = { host, services, settings: { get: (k, f) => settingsStore.get(k) ?? f, set: (k, v) => settingsStore.set(k, v) } };

  const addLabels = actions.find(a => a.id === 'editing.addLabels');
  const updateLabels = actions.find(a => a.id === 'editing.updateLabels');
  const addTitlesBelow = actions.find(a => a.id === 'editing.addTitlesBelow');
  const selectAllText = actions.find(a => a.id === 'editing.selectAllText');

  await addLabels.run(ctx, { template: 'A', startIndex: 1, fontName: 'Arial', fontSize: 14, offsetX: -10, offsetY: -5, bold: true });
  const labels = shapes.filter(s => s.type === 'TextBox');
  assert.equal(labels.length, 2);
  assert.deepEqual(labels.map(l => l.textFrame.textRange.text), ['A', 'B']);
  assert.equal(labels[0].textFrame.textRange.font.bold, true);
  assert.equal(labels[0].textFrame.textRange.font.size, 14);

  await updateLabels.run(ctx, { template: '(a)', startIndex: 1, fontName: 'Helvetica', fontSize: 16, bold: false });
  assert.deepEqual(labels.map(l => l.textFrame.textRange.text), ['(a)', '(b)']);
  assert.equal(labels[0].textFrame.textRange.font.name, 'Helvetica');
  assert.equal(labels[0].textFrame.textRange.font.size, 16);

  selectedIds = ['img-1'];
  await addTitlesBelow.run(ctx, { fontName: 'Arial', fontSize: 12, distance: 4, text: 'Sample A', autoGroup: false, center: true });
  const titleBox = shapes.at(-1);
  assert.equal(titleBox.textFrame.textRange.text, 'Sample A');
  assert.equal(titleBox.textFrame.textRange.paragraphFormat.horizontalAlignment, 'Center');

  await selectAllText.run(ctx, {});
  assert.equal(selectedIds.length, 3, 'all 3 text boxes are selected');
  assert.equal(backendCalls, 0, 'zero backend or AppleScript calls for core figure composition');
});

test('PowerPoint rollback and text deletion stay on the captured slide after selection changes', async t => {
  const originals = { Office: globalThis.Office, PowerPoint: globalThis.PowerPoint };
  t.after(() => { for (const [name, value] of Object.entries(originals)) { if (value === undefined) delete globalThis[name]; else globalThis[name] = value; } });
  const changes = [];
  const slides = collection(['256', '257'].map(id => loadable({ id, shapes: collection([
    loadable({ id: '4', type: 'Group', group: { shapes: collection([]), ungroup: () => changes.push(['ungroup', id]) } }),
    loadable({ id: '5', type: 'TextBox', delete: () => changes.push(['delete', id]) })
  ]) })));
  const presentation = loadable({ id: 'deck', slides, getSelectedSlides: () => collection([slides.items[1]]) });
  globalThis.Office = { HostType: { PowerPoint: 'PowerPoint' }, onReady: async () => ({ host: 'PowerPoint' }),
    context: { document: { url: '/tmp/fixture.pptx' }, requirements: { isSetSupported: () => true } } };
  globalThis.PowerPoint = { run: operation => operation({ presentation, sync: async () => {} }) };
  const host = await connectPowerPoint({ request() { assert.fail('unexpected native call'); } });
  t.after(() => host.dispose());
  const options = { target: { presentationId: 'deck', slideId: '256' }, readback: false };
  await host.apply([{ op: 'ungroup', id: '4' }], options);
  await host.apply([{ op: 'delete', ids: ['5'] }], options);
  await assert.rejects(host.apply([{ op: 'delete', ids: ['5'] }],
    { ...options, target: { ...options.target, presentationId: 'another deck' } }), /找不到操作开始时/);
  assert.deepEqual(changes, [['ungroup', '256'], ['delete', '256']]);
});

test('PowerPoint reads retry a transient GeneralException but give up on other errors', async t => {
  const original = globalThis.PowerPoint;
  t.after(() => { if (original === undefined) delete globalThis.PowerPoint; else globalThis.PowerPoint = original; });
  const { readPowerPoint } = await import('../src/powerpoint.js');
  let calls = 0;
  globalThis.PowerPoint = { run: async operation => { calls++; if (calls < 3) throw Object.assign(new Error('GeneralException'), { code: 'GeneralException' }); return operation(); } };
  assert.equal(await readPowerPoint(() => 'read'), 'read');
  assert.equal(calls, 3);
  calls = 0;
  globalThis.PowerPoint = { run: async () => { calls++; throw Object.assign(new Error('InvalidArgument'), { code: 'InvalidArgument' }); } };
  await assert.rejects(readPowerPoint(() => 'never'), /InvalidArgument/);
  assert.equal(calls, 1, 'a real error is not retried');
});

test('powerpoint selection watch uses DocumentSelectionChanged without polling and pauses polling while hidden', { timeout: 15000 }, async t => {
  const originals = { Office: globalThis.Office, PowerPoint: globalThis.PowerPoint, document: globalThis.document };
  t.after(() => { for (const [name, value] of Object.entries(originals)) { if (value === undefined) delete globalThis[name]; else globalThis[name] = value; } });

  const visibilityListeners = new Set();
  const docMock = {
    visibilityState: 'visible',
    addEventListener(name, fn) { if (name === 'visibilitychange') visibilityListeners.add(fn); },
    removeEventListener(name, fn) { if (name === 'visibilitychange') visibilityListeners.delete(fn); },
    setVisibility(state) {
      this.visibilityState = state;
      for (const fn of visibilityListeners) fn({});
    }
  };
  globalThis.document = docMock;

  const addedHandlers = [];
  const removedHandlers = [];
  const shapes = [loadable({ id: 'p1', name: 'Pic1', left: 10, top: 10, width: 100, height: 80, rotation: 0, type: 'Image', zOrderPosition: 1 })];
  const slides = collection([loadable({ id: 'slide1', index: 0, shapes: collection(shapes) })]);
  let selection = [];
  const presentation = loadable({ id: 'deck', slides, pageSetup: loadable({ slideWidth: 720, slideHeight: 540 }),
    getSelectedSlides() { return collection(slides.items); }, getSelectedShapes() { return collection(selection); } });

  globalThis.Office = {
    HostType: { PowerPoint: 'PowerPoint' },
    EventType: { DocumentSelectionChanged: 'DocumentSelectionChanged' },
    AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
    onReady: async () => ({ host: 'PowerPoint' }),
    context: {
      document: {
        url: '/Users/someone/event-watch.pptx',
        addHandlerAsync(type, handler, cb) {
          assert.equal(type, 'DocumentSelectionChanged');
          addedHandlers.push(handler);
          queueMicrotask(() => cb({ status: 'succeeded' }));
        },
        removeHandlerAsync(type, options, cb) {
          assert.equal(type, 'DocumentSelectionChanged');
          removedHandlers.push(options?.handler);
          cb?.({ status: 'succeeded' });
        }
      },
      requirements: { isSetSupported: () => true }
    }
  };

  let powerPointRuns = 0;
  globalThis.PowerPoint = { run: async operation => { powerPointRuns++; return operation({ presentation, sync: async () => {} }); } };
  const services = { async request() { throw new Error('unexpected backend request'); }, async reportError(err) { throw err; } };
  const host = await connectPowerPoint(services);
  t.after(() => host.dispose());

  const selectionChanges = [];
  host.onChanged(change => selectionChanges.push(change?.kind), { mode: 'selection' });
  await new Promise(resolve => setTimeout(resolve, 1750));
  assert.equal(addedHandlers.length, 1, 'DocumentSelectionChanged handler registered once');
  assert.equal(powerPointRuns, 0, 'selection-only watch performs zero background polling when event is supported');

  addedHandlers[0]({});
  await new Promise(resolve => setTimeout(resolve, 20));
  assert.deepEqual(selectionChanges, ['selection'], 'DocumentSelectionChanged triggers immediate selection notification');
  assert.equal(powerPointRuns, 0, 'selection event does not run full surfaceFingerprint');

  const allChanges = [];
  const unsubAll = host.onChanged(change => allChanges.push(change?.kind), { mode: 'all' });
  selection = [shapes[0]];
  addedHandlers[0]({});
  await Promise.resolve();
  assert.deepEqual(selectionChanges, ['selection', 'selection'], 'the first click with an all-mode listener reaches the pane before any fingerprint exists');
  assert.deepEqual(allChanges, ['selection']);
  docMock.setVisibility('hidden');
  await new Promise(resolve => setTimeout(resolve, 1750));
  assert.equal(powerPointRuns, 0, 'background polling pauses when document is hidden');

  const countBeforeVisible = selectionChanges.length;
  docMock.setVisibility('visible');
  await new Promise(resolve => setTimeout(resolve, 20));
  assert.ok(selectionChanges.length > countBeforeVisible, 'returning visible emits a refresh notification');
  await new Promise(resolve => setTimeout(resolve, 1750));
  assert.ok(powerPointRuns >= 1, 'polling resumes when visible with an all-mode listener');

  unsubAll();
  powerPointRuns = 0;
  await new Promise(resolve => setTimeout(resolve, 1750));
  assert.equal(powerPointRuns, 0, 'polling stops immediately once the last all-mode listener unsubscribes');

  host.dispose();
  assert.equal(removedHandlers.length, 1, 'dispose removes DocumentSelectionChanged handler');
});

test('addTitlesBelow batches text insertion and auto-grouping without per-shape round trips', async t => {
  const originals = { Office: globalThis.Office, PowerPoint: globalThis.PowerPoint };
  t.after(() => { for (const [name, value] of Object.entries(originals)) { if (value === undefined) delete globalThis[name]; else globalThis[name] = value; } });
  const { actions } = await import('../src/features/editing.js');
  let nextId = 20;
  const makeShape = ({ id, name, type, left, top, width, height, text = '', children = [] }) => {
    const font = { name: 'Arial', size: 12, bold: false };
    const paragraphFormat = { horizontalAlignment: 'Left' };
    const textRange = loadable({ text, font, paragraphFormat });
    const textFrame = loadable({ textRange, wordWrap: false, autoSizeSetting: 'AutoSizeNone', hasText: Boolean(text) });
    return loadable({ id, name, type, left, top, width, height, rotation: 0, zOrderPosition: 1, textFrame, children,
      ...(type === 'Group' ? { group: { shapes: collection(children), ungroup() {
        shapes.splice(shapes.findIndex(shape => shape.id === id), 1, ...children);
      } } } : {}),
      tags: shapeTags(), delete() { const i = shapes.findIndex(s => s.id === id); if (i >= 0) shapes.splice(i, 1); } });
  };
  const shapes = [1, 2, 3, 4].map(i => makeShape({ id: `img-${i}`, name: `Fig${i}`, type: 'Image', left: i * 100, top: 40, width: 80, height: 60 }));
  let selectedIds = shapes.map(s => s.id);
  let activeRunStats = null;
  const runLogs = [];

  const shapeCollection = Object.assign(collection(shapes), {
    getItem(id) { return shapes.find(shape => shape.id === id); },
    addTextBox(text, options) {
      if (activeRunStats) activeRunStats.addedBoxes++;
      const box = makeShape({ id: `tb-${++nextId}`, name: 'TextBox', type: 'TextBox',
        left: options.left, top: options.top, width: options.width, height: options.height, text });
      shapes.push(box);
      return box;
    }
  });
  const slide = loadable({ id: 'slide1', index: 0, shapes: shapeCollection, setSelectedShapes(ids) { selectedIds = [...ids]; } });
  const slides = collection([slide]);
  const presentation = loadable({
    id: 'deck', slides, pageSetup: loadable({ slideWidth: 720, slideHeight: 540 }),
    getSelectedSlides() { return collection(slides.items); },
    getSelectedShapes() {
      const selected = selectedIds.map(id => shapes.find(s => s.id === id)).filter(Boolean);
      return Object.assign(collection(selected), {
        group() {
          if (activeRunStats) activeRunStats.grouped++;
          const members = selectedIds.map(id => {
            const idx = shapes.findIndex(s => s.id === id);
            return idx >= 0 ? shapes.splice(idx, 1)[0] : null;
          }).filter(Boolean);
          const groupShape = makeShape({ id: `grp-${++nextId}`, name: 'Group', type: 'Group',
            left: members[0]?.left ?? 0, top: members[0]?.top ?? 0, width: 80, height: 90, children: members });
          shapes.push(groupShape);
          return groupShape;
        }
      });
    }
  });
  globalThis.Office = { HostType: { PowerPoint: 'PowerPoint' }, onReady: async () => ({ host: 'PowerPoint' }),
    context: { document: { url: '/Users/someone/titles.pptx' }, requirements: { isSetSupported: () => true } } };
  globalThis.PowerPoint = {
    run: async operation => {
      const stats = { addedBoxes: 0, grouped: 0 };
      runLogs.push(stats);
      activeRunStats = stats;
      try { return await operation({ presentation, sync: async () => {} }); }
      finally { activeRunStats = null; }
    }
  };

  let backendCalls = 0;
  const services = { async request() { backendCalls++; throw new Error('unexpected backend request'); }, async reportError() {} };
  const host = await connectPowerPoint(services);
  t.after(() => host.dispose());
  const settingsStore = new Map();
  const ctx = { host, services, settings: { get: (k, f) => settingsStore.get(k) ?? f, set: (k, v) => settingsStore.set(k, v) } };
  const addTitlesBelow = actions.find(a => a.id === 'editing.addTitlesBelow');

  await addTitlesBelow.run(ctx, { fontName: 'Arial', fontSize: 12, distance: 4, text: 'Panel', autoGroup: false, center: true });
  assert.deepEqual(runLogs.filter(r => r.addedBoxes > 0), [{ addedBoxes: 4, grouped: 0 }], 'all 4 titles insert in one PowerPoint.run batch');
  assert.equal(selectedIds.length, 4);
  assert.equal(backendCalls, 0);

  runLogs.length = 0;
  selectedIds = ['img-1', 'img-2', 'img-3', 'img-4'];
  await addTitlesBelow.run(ctx, { fontName: 'Arial', fontSize: 12, distance: 4, text: 'Grouped', autoGroup: true, center: false });
  assert.deepEqual(runLogs.filter(r => r.addedBoxes > 0), [{ addedBoxes: 4, grouped: 0 }], 'autoGroup still inserts all 4 titles in one batch');
  assert.deepEqual(runLogs.filter(r => r.grouped > 0), [{ addedBoxes: 0, grouped: 4 }], 'autoGroup batches all 4 groupings in one PowerPoint.run');
  assert.equal(shapes.filter(s => s.type === 'Group').length, 4);
  assert.equal(selectedIds.length, 4);
  const grouped = shapes.find(shape => shape.type === 'Group');
  grouped.tags.add('CHECK_GROUP', '1');
  grouped.children[1].tags.add('CHECK_CHILD', '1');
  const tree = await host.context({ scope: 'slide', geometry: true, tags: true });
  const matchedGroup = tree.shapes.find(shape => shape.id === grouped.id);
  assert.equal(matchedGroup.tags.CHECK_GROUP, '1');
  assert.equal(matchedGroup.children[1].tags.CHECK_CHILD, '1');
  await host.apply(shapes.filter(shape => shape.type === 'Group').map(shape => ({ op: 'ungroup', id: shape.id })), { readback: false });
  assert.deepEqual(shapes.filter(shape => shape.type === 'Image').map(shape => shape.id).sort(), ['img-1', 'img-2', 'img-3', 'img-4']);
  assert.equal(backendCalls, 0);
});

class MockElement {
  constructor(tag) { this.tagName = tag; this.children = []; this.events = new Map(); this.attributes = {}; this.value = ''; }
  append(...nodes) { this.children.push(...nodes); if (this.tagName === 'select' && !this.value) this.value = this.children[0]?.value || ''; }
  replaceChildren(...nodes) { this.children = nodes; }
  setAttribute(key, value) { this.attributes[key] = value; }
  addEventListener(name, listener) { this.events.set(name, listener); }
  dispatch(name) { this.events.get(name)?.({ preventDefault() {} }); }
}
const walkElements = root => [root, ...root.children.filter(c => c instanceof MockElement).flatMap(walkElements)];

test('scaleTogether and live controls avoid PPTX export for pure pictures and keep rich font scaling for text', { timeout: 10000 }, async t => {
  const originals = { Office: globalThis.Office, PowerPoint: globalThis.PowerPoint, document: globalThis.document };
  t.after(() => { for (const [name, value] of Object.entries(originals)) { if (value === undefined) delete globalThis[name]; else globalThis[name] = value; } });
  globalThis.document = { createElement: tag => new MockElement(tag) };

  const { actions, panels } = await import('../src/features/editing.js');
  const shapes = [
    loadable({ id: 'img-1', name: 'Fig1', type: 'Image', left: 20, top: 40, width: 120, height: 90, rotation: 0 }),
    loadable({ id: 'img-2', name: 'Fig2', type: 'Image', left: 160, top: 40, width: 100, height: 80, rotation: 0 }),
    loadable({ id: 'tb-1', name: 'Text1', type: 'TextBox', left: 40, top: 180, width: 100, height: 40, rotation: 0,
      fill: loadable({ type: 'NoFill' }), lineFormat: loadable({ visible: false }), tags: loadable({ items: [] }),
      textFrame: loadable({ textRange: loadable({ text: 'hi', font: loadable({ name: 'Arial', size: 12 }), paragraphFormat: { horizontalAlignment: 'Left' } }) }) })
  ];
  let selectedIds = ['img-1', 'img-2'];
  let exports = 0;
  const backendRequests = [];
  const slides = collection([loadable({ id: 'slide1', index: 0, shapes: collection(shapes) })]);
  slides.exportAsBase64Presentation = () => { exports++; return { value: 'UEsDBA==' }; };
  const presentation = loadable({ id: 'deck', slides, pageSetup: loadable({ slideWidth: 720, slideHeight: 540 }),
    getSelectedSlides() { return collection(slides.items); },
    getSelectedShapes() { return collection(selectedIds.map(id => shapes.find(s => s.id === id)).filter(Boolean)); } });
  globalThis.Office = { HostType: { PowerPoint: 'PowerPoint' }, onReady: async () => ({ host: 'PowerPoint' }),
    context: { document: { url: '/Users/someone/scale.pptx' }, requirements: { isSetSupported: () => true } } };
  globalThis.PowerPoint = { run: async operation => operation({ presentation, sync: async () => {} }) };

  const services = {
    async request(route, input) {
      backendRequests.push([route, input]);
      if (route === 'pptx/parse') {
        return { width: 720, height: 540, slides: [{ index: 0, shapes: [{ id: 'tb-1', name: 'Text1', type: 'text',
          left: 40, top: 180, width: 100, height: 40, rotation: 0, text: 'hi', font: { size: 12 }, ref: { slideIndex: 0, shapePath: [3] } }] }] };
      }
      if (route === 'native/execute' && input.action === 'textSelection') return null;
      if (route === 'native/execute' && input.action === 'apply') return { applied: input.commands.length };
      throw new Error('unexpected route ' + route);
    },
    async reportError(error) { throw error; }
  };

  const host = await connectPowerPoint(services);
  t.after(() => host.dispose());
  const settingsStore = new Map();
  const ctxBase = { host, services, settings: { get: (k, f) => settingsStore.get(k) ?? f, set: (k, v) => settingsStore.set(k, v) } };
  const scaleTogether = actions.find(a => a.id === 'editing.scaleTogether');

  await scaleTogether.run(ctxBase, { percent: 150 });
  assert.equal(shapes[0].width, 180);
  assert.equal(shapes[0].height, 135);
  assert.equal(exports, 0, 'pure-image scaleTogether performs zero PPTX exports');
  assert.equal(backendRequests.length, 0, 'pure-image scaleTogether makes zero backend calls');

  let refreshedCount = 0;
  let refreshDone = Promise.withResolvers();
  const panelCtx = { ...ctxBase, notify(msg, kind) { if (kind === 'error') throw new Error(msg); },
    refresh() { refreshedCount++; refreshDone.resolve(); } };
  const root = new MockElement('div');
  const cleanup = await panels[0].mount(root, panelCtx);
  t.after(cleanup);
  assert.equal(exports, 0, 'mounting live controls reads lightweight geometry without PPTX export');

  const scaleSlider = walkElements(root).find(node => node.type === 'range' && node.attributes['aria-label'] === '图文比例 (%)');
  scaleSlider.value = '200';
  scaleSlider.dispatch('input');
  await refreshDone.promise;
  assert.equal(shapes[0].width, 360);
  assert.equal(exports, 0, 'dragging scale slider on pure images stays zero-export');
  assert.equal(backendRequests.length, 0);

  refreshDone = Promise.withResolvers();
  const gapSlider = walkElements(root).find(node => node.type === 'range' && node.attributes['aria-label'] === '水平间距 (cm)');
  gapSlider.value = '1';
  gapSlider.dispatch('input');
  await refreshDone.promise;
  assert.equal(exports, 0, 'dragging spacing slider stays zero-export');
  assert.equal(backendRequests.length, 0);
  cleanup();

  selectedIds = ['tb-1'];
  await scaleTogether.run(ctxBase, { percent: 200 });
  assert.equal(exports, 1, 'text scaling still exports once to read full rich typography');
  assert.equal(shapes[2].textFrame.textRange.font.size, 24, 'text font size scales proportionally');
  assert.equal(shapes[2].width, 200);
  assert.equal(shapes[2].height, 80);
});
