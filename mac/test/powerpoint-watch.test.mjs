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
  const presentation = loadable({ id: 'deck', slides, pageSetup: loadable({ slideWidth: 720, slideHeight: 540 }),
    getSelectedSlides() { return collection(slides.items); }, getSelectedShapes() { return collection([]); } });
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
  let changes = 0;
  host.onChanged(() => { changes++; });
  await new Promise(resolve => setTimeout(resolve, 1700));
  assert.deepEqual(reported, [], 'the first fingerprint succeeds on a slide with a text shape');
  range.text = 'after';
  await new Promise(resolve => setTimeout(resolve, 1700));
  assert.equal(changes, 1, 'editing the text is reported as a change');
  assert.deepEqual(reported, []);
});
