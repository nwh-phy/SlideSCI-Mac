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

// The deep pass must not run AppleScript: it targets the active window and fails for every other deck.
test('powerpoint deep watch pass reads the slide without native text-cursor calls', { timeout: 20000 }, async t => {
  const originals = { Office: globalThis.Office, PowerPoint: globalThis.PowerPoint };
  t.after(() => { for (const [name, value] of Object.entries(originals)) { if (value === undefined) delete globalThis[name]; else globalThis[name] = value; } });
  const shapes = [loadable({ id: 'p1', name: 'picture', left: 10, top: 10, width: 200, height: 40, rotation: 0, type: 'Image', zOrderPosition: 1 })];
  const slides = collection([loadable({ id: 'slide1', index: 0, shapes: collection(shapes), exportAsBase64() {} })]);
  slides.exportAsBase64Presentation = () => ({ value: 'UEsDBA==' });
  const presentation = loadable({ id: 'deck', slides, pageSetup: loadable({ slideWidth: 720, slideHeight: 540 }),
    getSelectedSlides() { return collection(slides.items); }, getSelectedShapes() { return collection([]); } });
  globalThis.Office = { HostType: { PowerPoint: 'PowerPoint' }, onReady: async () => ({ host: 'PowerPoint' }),
    context: { document: { url: '/Users/someone/deck.pptx' }, requirements: { isSetSupported: () => true } } };
  globalThis.PowerPoint = { run: async operation => operation({ presentation, sync: async () => {} }) };
  const nativeActions = [], reported = [];
  const services = {
    async request(route, input) {
      if (route === 'native/execute') { nativeActions.push(input.action); if (input.action === 'documentInfo') return { name: 'deck.pptx' }; return null; }
      if (route === 'pptx/parse') return { width: 720, height: 540, slides: [{ index: 0, shapes: [] }] };
      throw new Error(`unexpected ${route}`);
    },
    async reportError(error) { reported.push(error); }
  };
  const host = await connectPowerPoint(services);
  t.after(() => host.dispose());
  host.onChanged(() => {});
  await new Promise(resolve => setTimeout(resolve, 1500 * 6 + 800));
  assert.deepEqual(reported, []);
  assert.deepEqual(nativeActions, ['documentInfo'], 'only the one-time document lookup reaches AppleScript');
});
