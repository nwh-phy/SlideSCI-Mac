import test from 'node:test';
import assert from 'node:assert/strict';
import { connectWps } from '../src/wps.js';

function collection(items) {
  return { items, get Count() { return items.length; }, Item(index) { return items[index - 1]; } };
}
function shape(id, type, parent, children = []) {
  const node = { Id: id, Type: type, Name: `shape ${id}`, Left: 0, Top: 0, Width: 10, Height: 10, Rotation: 0, deleted: false };
  node.Tags = { Count: 1, Name: () => 'AUDIT_ID', Value: () => `SHAPE-${id}` };
  node.GroupItems = collection(children);
  node.Delete = () => {
    if (node.deleted) throw new Error(`shape ${id} was already deleted`);
    node.deleted = true;
    const owner = parent();
    owner.items.splice(owner.items.indexOf(node), 1);
  };
  return node;
}

// Rolling back a scale bar or zoom deletes a new group together with its new children.
test('wps delete removes a group once even when its members are listed too', async t => {
  const original = globalThis.window;
  t.after(() => { if (original === undefined) delete globalThis.window; else globalThis.window = original; });
  const slideShapes = collection([]);
  const members = collection([]);
  const bar = shape(11, 1, () => members), label = shape(12, 17, () => members);
  members.items.push(bar, label);
  const group = shape(10, 6, () => slideShapes, members.items);
  group.GroupItems = members;
  const keep = shape(20, 1, () => slideShapes);
  slideShapes.items.push(group, keep);
  const slide = { SlideID: 1, SlideIndex: 1, Shapes: slideShapes };
  globalThis.window = { Application: {
    ActivePresentation: { Name: 'delete.pptx', Slides: collection([slide]), PageSetup: { SlideWidth: 720, SlideHeight: 540 } },
    ActiveWindow: { View: { Slide: slide }, Selection: { Type: 0 } }
  } };
  const host = connectWps({ async request(route) { throw new Error(`unexpected ${route}`); }, async reportError() {} });
  t.after(() => host.dispose());
  const snapshot = await host.context({ scope: 'slide', geometry: true, tags: true });
  assert.deepEqual(snapshot.shapes[0].tags, { AUDIT_ID: 'SHAPE-10' });
  assert.deepEqual(snapshot.shapes[0].children[0].tags, { AUDIT_ID: 'SHAPE-11' });
  const result = await host.apply([{ op: 'delete', ids: ['10', '11', '12'] }], { readback: false });
  assert.deepEqual(result, { applied: 1 });
  assert.equal(group.deleted, true);
  assert.equal(bar.deleted, false, 'members go with their group instead of being deleted again');
  assert.deepEqual(slideShapes.items.map(item => item.Id), [20]);
  await assert.rejects(host.apply([{ op: 'delete', ids: ['20', '99'] }], { readback: false }), /找不到对象 99/);
  assert.equal(keep.deleted, false, 'a missing id is reported before anything is deleted');
});

test('wps discards a new textbox if the host rejects its properties', async t => {
  const original = globalThis.window;
  t.after(() => { if (original === undefined) delete globalThis.window; else globalThis.window = original; });
  const shapes = collection([]);
  shapes.items.push(shape(20, 1, () => shapes));
  shapes.AddTextbox = () => {
    const box = shape(21, 17, () => shapes);
    Object.defineProperty(box, 'Name', { set() { throw new Error('WPS rejected the title name'); } });
    shapes.items.push(box);
    return box;
  };
  const slide = { SlideID: 1, SlideIndex: 1, Shapes: shapes };
  globalThis.window = { Application: {
    ActivePresentation: { Name: 'titles.pptx', Slides: collection([slide]), PageSetup: { SlideWidth: 720, SlideHeight: 540 } },
    ActiveWindow: { View: { Slide: slide }, Selection: { Type: 0 } }
  } };
  const host = connectWps();
  t.after(() => host.dispose());
  await assert.rejects(host.apply([{ op: 'insertText', name: 'title', text: 'Sample' }], { readback: false }), /rejected the title name/);
  assert.deepEqual(shapes.items.map(item => item.Id), [20]);
});

test('a WPS delete addressed to the original slide ignores the same shape ID on a newly selected slide', async t => {
  const original = globalThis.window;
  t.after(() => { if (original === undefined) delete globalThis.window; else globalThis.window = original; });
  const slides = [41, 42].map((id, index) => {
    const shapes = collection([]);
    shapes.items.push(shape(20, 17, () => shapes));
    return { SlideID: id, SlideIndex: index + 1, Shapes: shapes };
  });
  globalThis.window = { Application: {
    ActivePresentation: { Name: 'fixture.pptx', FullName: '/tmp/fixture.pptx', Slides: collection(slides) },
    ActiveWindow: { View: { Slide: slides[1] }, Selection: { Type: 0 } }
  } };
  const host = connectWps();
  t.after(() => host.dispose());
  await host.apply([{ op: 'delete', ids: ['20'] }],
    { target: { presentationId: '/tmp/fixture.pptx', slideId: '41' }, readback: false });
  assert.equal(slides[0].Shapes.Count, 0);
  assert.equal(slides[1].Shapes.Count, 1);
});

// Labels drawn as rectangles are autoshapes (type 1); the geometry context still reports their text.
test('wps geometry context marks autoshapes that hold text when textMeta is requested', async t => {
  const original = globalThis.window;
  t.after(() => { if (original === undefined) delete globalThis.window; else globalThis.window = original; });
  const shapes = collection([]);
  const label = shape(31, 1, () => shapes), frame = shape(32, 1, () => shapes), box = shape(33, 17, () => shapes);
  label.HasTextFrame = -1; label.TextFrame = { TextRange: { Text: '(a)' } };
  frame.HasTextFrame = -1; frame.TextFrame = { TextRange: { Text: '' } };
  shapes.items.push(label, frame, box);
  const slide = { SlideID: 1, SlideIndex: 1, Shapes: shapes };
  globalThis.window = { Application: {
    ActivePresentation: { Name: 'labels.pptx', Slides: collection([slide]), PageSetup: { SlideWidth: 720, SlideHeight: 540 } },
    ActiveWindow: { View: { Slide: slide }, Selection: { Type: 0 } }
  } };
  const host = connectWps({ async request(route) { throw new Error(`unexpected ${route}`); }, async reportError() {} });
  t.after(() => host.dispose());
  const plain = await host.context({ scope: 'slide', geometry: true });
  assert.equal(plain.shapes.some(item => 'hasText' in item), false, 'text is only read when requested');
  const snapshot = await host.context({ scope: 'slide', geometry: true, textMeta: true });
  assert.deepEqual(snapshot.shapes.map(item => [item.id, item.type, item.hasText]), [['31', 'shape', true], ['32', 'shape', false], ['33', 'text', undefined]]);
});
