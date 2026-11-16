import test from 'node:test';
import assert from 'node:assert/strict';
import { connectPowerPoint } from '../src/powerpoint.js';
import { connectWps } from '../src/wps.js';
import { actions, panels } from '../src/features/editing.js';

const loadable = value => Object.assign(value, { load() { return this; } });
function collection(items) {
  return loadable({ items, get Count() { return items.length; }, Item(index) { return items[index - 1]; },
    getCount() { return { value: items.length }; },
    getItemAt(index) { return items[index]; }, getItem(id) { return items.find(item => String(item.id) === String(id)); } });
}
async function fixture(t, kind) {
  const originals = { Office: globalThis.Office, PowerPoint: globalThis.PowerPoint, window: globalThis.window };
  t.after(() => { for (const [name, value] of Object.entries(originals)) {
    if (value === undefined) delete globalThis[name]; else globalThis[name] = value;
  } });
  const shapes = [
    { id: 'a', name: 'reference', left: 80, top: 40, width: 200, height: 120, rotation: 0, type: 'GeometricShape' },
    { id: 'b', name: 'target', left: 320, top: 250, width: 100, height: 50, rotation: 0, type: 'GeometricShape' },
    { id: 'c', name: 'unselected', left: 510, top: 300, width: 40, height: 60, rotation: 0, type: 'GeometricShape' }
  ].map(loadable);
  const state = { selection: shapes.slice(0, 2), shapes };
  const services = { async request(route, input) {
    if (route === 'native/execute' && input.action === 'documentInfo') return { name: 'SlideSCI-Performance.pptx' };
    throw new Error(`Unexpected slow service call: ${route}/${input?.action || ''}`);
  }, reportError(error) { throw error; } };
  let host;
  if (kind === 'powerpoint') {
    const slides = collection([loadable({ id: 'slide1', index: 0, shapes: collection(shapes), setSelectedShapes(ids) {
      state.selection = ids.map(id => shapes.find(shape => shape.id === id));
    } })]);
    const presentation = loadable({ id: 'presentation', slides,
      pageSetup: loadable({ slideWidth: 720, slideHeight: 540 }),
      getSelectedSlides() { return collection(slides.items); }, getSelectedShapes() { return collection(state.selection); } });
    globalThis.Office = { HostType: { PowerPoint: 'PowerPoint' }, onReady: async () => ({ host: 'PowerPoint' }),
      context: { document: { url: 'file:///SlideSCI-Performance.pptx' }, requirements: { isSetSupported: () => true } } };
    globalThis.PowerPoint = { run: async operation => operation({ presentation, sync: async () => {} }) };
    host = await connectPowerPoint(services);
  } else {
    for (const shape of shapes) {
      for (const [native, key] of Object.entries({ Id: 'id', Name: 'name', Left: 'left', Top: 'top', Width: 'width', Height: 'height', Rotation: 'rotation' })) {
        Object.defineProperty(shape, native, { get() { return this[key]; }, set(value) { this[key] = value; } });
      }
      shape.Type = 1; shape.LockAspectRatio = -1;
    }
    const slide = { SlideID: 1, SlideIndex: 1, Shapes: collection(shapes) };
    slide.Shapes.Range = indices => ({ Select() { state.selection = indices.map(index => shapes[index - 1]); } });
    const selection = { get Type() { return state.selection.length ? 2 : 0; }, HasChildShapeRange: 0,
      get ShapeRange() { return collection(state.selection); } };
    globalThis.window = { Application: { ActivePresentation: { Name: 'SlideSCI-Performance.pptx', Slides: collection([slide]),
      PageSetup: { SlideWidth: 720, SlideHeight: 540 } }, ActiveWindow: { View: { Slide: slide }, Selection: selection } } };
    host = connectWps(services);
  }
  t.after(() => host.dispose());
  return { host, state, shapes };
}

for (const kind of ['powerpoint', 'wps']) {
  test(`${kind}: six alignments retain the first reference, dimensions and unselected objects`, async t => {
    const { host, shapes } = await fixture(t, kind);
    const [reference, target, unselected] = shapes;
    const untouched = { left: unselected.left, top: unselected.top };
    const expected = { left: ['left', 80], right: ['left', 180], top: ['top', 40], bottom: ['top', 110],
      horizontalCenter: ['left', 130], verticalCenter: ['top', 75] };
    const ctx = { host, settings: { get: () => ({}) } };
    for (const [edge, [axis, value]] of Object.entries(expected)) {
      target.left = 320; target.top = 250;
      await actions.find(action => action.id === `editing.align.${edge}`).run(ctx);
      assert.equal(target[axis], value, edge);
      assert.deepEqual([reference.left, reference.top, target.width, target.height], [80, 40, 100, 50]);
      assert.deepEqual({ left: unselected.left, top: unselected.top }, untouched);
    }
    const summary = await host.summary({ slides: false });
    assert.deepEqual([summary.width, summary.height, summary.slideIndex], [720, 540, 0]);
    assert.deepEqual(summary.selectionIds, ['a', 'b']);
  });
  test(`${kind}: page centering and a changed selection use fresh native geometry`, async t => {
    const { host, state, shapes } = await fixture(t, kind);
    state.selection = [shapes[1]];
    await host.align('horizontalCenter');
    assert.equal(shapes[1].left, 310);
    shapes[1].width = 180;
    await host.align('right');
    assert.equal(shapes[1].left, 540);
    state.selection = [];
    await assert.rejects(host.align('left'), /选择/);
    state.selection = [shapes[2]];
    await host.align('verticalCenter');
    assert.equal(shapes[2].top, 240);
  });
  test(`${kind}: position and size copying remain correct through geometry-only reads`, async t => {
    const { host, state, shapes } = await fixture(t, kind);
    const settings = { get: () => ({}), set() {} };
    const run = id => actions.find(action => action.id === id).run({ host, settings });
    state.selection = [shapes[0]];
    await run('editing.copy.dimensions');
    state.selection = [shapes[1]];
    await run('editing.paste.dimensions');
    assert.deepEqual([shapes[1].width, shapes[1].height, shapes[1].left, shapes[1].top], [200, 120, 320, 250]);
    assert.deepEqual([shapes[2].width, shapes[2].height], [40, 60]);
  });
  test(`${kind}: metadata-only writes can discard readback without changing other objects`, async t => {
    const { host, shapes } = await fixture(t, kind);
    const result = await host.apply([{ op: 'update', id: 'b', properties: { name: 'Renamed target' } }], { readback: false });
    assert.equal(shapes[1].name, 'Renamed target');
    assert.equal(shapes[0].name, 'reference');
    assert.deepEqual(result, { applied: 1 });
  });
  test(`${kind}: explicit selection replaces the previous alignment reference`, async t => {
    const { host, state, shapes } = await fixture(t, kind);
    state.selection = [shapes[1]];
    await host.summary({ slides: false });
    await host.apply([{ op: 'select', ids: ['a', 'b'] }], { readback: false });
    await host.align('right');
    assert.equal(shapes[0].left, 80);
    assert.equal(shapes[1].left, 180);
  });
}

class Element {
  constructor(tag) { this.tagName = tag; this.children = []; this.events = new Map(); this.attributes = {}; this.value = ''; }
  append(...nodes) { this.children.push(...nodes); if (this.tagName === 'select' && !this.value) this.value = this.children[0]?.value || ''; }
  replaceChildren(...nodes) { this.children = nodes; }
  setAttribute(key, value) { this.attributes[key] = value; }
  addEventListener(name, listener) { this.events.set(name, listener); }
  dispatch(name) { this.events.get(name)?.({ preventDefault() {} }); }
}
const descendants = element => [element, ...element.children.filter(child => child instanceof Element).flatMap(descendants)];
test('rapid slider changes finish at the latest value without replaying obsolete intermediate edits', async t => {
  const previous = globalThis.document;
  globalThis.document = { createElement: tag => new Element(tag) };
  t.after(() => { if (previous === undefined) delete globalThis.document; else globalThis.document = previous; });
  const model = { presentationId: 'p', slideId: 's', selectionIds: ['a'], shapes: [
    { id: 'a', name: 'A', type: 'text', left: 20, top: 40, width: 100, height: 50, font: { size: 10 }, text: 'A' } ] };
  const started = Promise.withResolvers(), release = Promise.withResolvers(), refreshed = Promise.withResolvers();
  let calls = 0;
  const errors = [];
  const ctx = { settings: { get: () => ({}), set() {} }, notify: message => errors.push(message),
    refresh() { refreshed.resolve(); }, host: { context: async () => structuredClone(model), onChanged: () => () => {},
      async apply(commands) { calls++; if (calls === 1) { started.resolve(); await release.promise; }
        for (const command of commands) Object.assign(model.shapes.find(shape => shape.id === command.id), command.properties);
        return structuredClone(model);
      } } };
  const root = new Element('div');
  const cleanup = await panels[0].mount(root, ctx);
  t.after(cleanup);
  const slider = descendants(root).find(node => node.type === 'range' && node.attributes['aria-label'] === '图文比例 (%)');
  slider.value = '120'; slider.dispatch('input');
  await started.promise;
  for (let value = 121; value <= 200; value++) { slider.value = String(value); slider.dispatch('input'); }
  release.resolve();
  await refreshed.promise;
  assert.deepEqual(errors, []);
  assert.equal(calls, 2);
  assert.deepEqual([model.shapes[0].width, model.shapes[0].height, model.shapes[0].font.size], [200, 100, 20]);
});
