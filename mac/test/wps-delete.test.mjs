import test from 'node:test';
import assert from 'node:assert/strict';
import { connectWps } from '../src/wps.js';

function collection(items) {
  return { items, get Count() { return items.length; }, Item(index) { return items[index - 1]; } };
}
function shape(id, type, parent, children = []) {
  const node = { Id: id, Type: type, Name: `shape ${id}`, Left: 0, Top: 0, Width: 10, Height: 10, Rotation: 0, deleted: false };
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
  const result = await host.apply([{ op: 'delete', ids: ['10', '11', '12'] }], { readback: false });
  assert.deepEqual(result, { applied: 1 });
  assert.equal(group.deleted, true);
  assert.equal(bar.deleted, false, 'members go with their group instead of being deleted again');
  assert.deepEqual(slideShapes.items.map(item => item.Id), [20]);
  await assert.rejects(host.apply([{ op: 'delete', ids: ['20', '99'] }], { readback: false }), /找不到对象 99/);
  assert.equal(keep.deleted, false, 'a missing id is reported before anything is deleted');
});
