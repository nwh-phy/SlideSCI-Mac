import test from 'node:test';
import assert from 'node:assert/strict';
import { connectPowerPoint } from '../src/powerpoint.js';

const loadable = value => Object.assign(value, { load() { return this; } });
function collection(items) {
  return loadable({ items, getCount() { return { value: items.length }; }, getItemAt(index) { return items[index]; } });
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
      tags: { add() {}, delete() {} }, delete() { const i = shapes.findIndex(s => s.id === id); if (i >= 0) shapes.splice(i, 1); } });
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
