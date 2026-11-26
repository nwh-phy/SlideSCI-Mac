import test from 'node:test';
import assert from 'node:assert/strict';
import { actions, panels } from '../src/features/editing.js';

const clone = value => structuredClone(value);
const flatten = shapes => (shapes || []).flatMap(shape => [shape, ...flatten(shape.children)]);
const markerTags = shapes => flatten(shapes).flatMap(shape => shape.tags?.SLIDESCI_TITLE_BATCH ? [shape.tags.SLIDESCI_TITLE_BATCH] : []);
const newCtx = host => {
  const store = new Map();
  return { host, settings: { get: async (key, fallback) => store.get(key) ?? fallback, set: async (key, value) => { store.set(key, value); } } };
};

// A slide with four selected pictures, one preexisting group and one unrelated text box.
// Grouping and insertion can fail after some commands were already applied, like WPS native apply.
function titleHost({ insertAfter, groupAfter, selectFail = false, recreateIds = false, groupTagFails = false } = {}) {
  const originals = [1, 2, 3, 4].map(i => ({ id: `img-${i}`, name: `Fig${i}`, type: 'picture',
    left: i * 100, top: 40, width: 80, height: 60, rotation: 0, tags: {}, children: [] }));
  const keptPicture = { id: 'kept-pic', name: 'Kept', type: 'picture', left: 4, top: 4, width: 20, height: 20, tags: {}, children: [] };
  const keptGroup = { id: 'keep-group', name: 'Keep', type: 'group', left: 4, top: 4, width: 20, height: 20,
    tags: {}, children: [keptPicture] };
  const unrelated = { id: 'note', name: 'Note', type: 'text', left: 600, top: 400, width: 80, height: 30,
    text: 'keep me', tags: {}, children: [] };
  const state = { shapes: [...originals, keptGroup, unrelated], selection: originals.map(shape => shape.id), next: 0 };
  let insertions = 0;
  let groupings = 0;
  const all = () => flatten(state.shapes);
  const find = id => all().find(shape => String(shape.id) === String(id));
  function remove(id) {
    const removeFrom = list => {
      const index = list.findIndex(shape => String(shape.id) === String(id));
      if (index >= 0) { list.splice(index, 1); return true; }
      return list.some(shape => shape.children && removeFrom(shape.children));
    };
    removeFrom(state.shapes);
  }
  const host = {
    async context() {
      return { presentationId: 'deck', slideId: 'slide-1', slideIndex: 0, width: 720, height: 540,
        selectionIds: [...state.selection], shapes: clone(state.shapes) };
    },
    async apply(commands, options = {}) {
      for (const command of commands) {
        if (command.op === 'insertText') {
          if (insertAfter !== undefined && insertions >= insertAfter) throw new Error('模拟插入失败');
          state.shapes.push({ id: `title-${++state.next}`, name: command.name, type: 'text',
            left: command.left, top: command.top, width: command.width, height: command.height,
            text: command.text, tags: { ...(command.tags || {}) }, children: [] });
          insertions++;
        } else if (command.op === 'group') {
          if (groupAfter !== undefined && groupings >= groupAfter) throw new Error('模拟编组失败');
          const members = command.ids.map(find);
          if (members.some(member => !member)) throw new Error('编组对象不存在。');
          members.forEach(member => remove(member.id));
          const first = members[0];
          state.shapes.push({ id: `group-${++state.next}`, name: command.name, type: 'group',
            left: first.left, top: first.top, width: first.width, height: first.height,
            tags: groupTagFails ? {} : { ...(command.tags || {}) }, children: members });
          groupings++;
          if (groupTagFails) throw new Error('模拟编组标记失败');
        } else if (command.op === 'ungroup') {
          const index = state.shapes.findIndex(shape => String(shape.id) === String(command.id));
          if (index < 0) throw new Error('组合不存在。');
          const [group] = state.shapes.splice(index, 1);
          // Some hosts regenerate child IDs when ungrouping; ownership must survive that.
          if (recreateIds) {
            const [original, title] = group.children;
            original.id = title.id;
            title.id = `restored-${++state.next}`;
          }
          state.shapes.splice(index, 0, ...group.children);
        } else if (command.op === 'delete') {
          for (const id of command.ids) remove(id);
        } else if (command.op === 'select') {
          if (selectFail) throw new Error('模拟选择失败');
          state.selection = command.ids.map(String);
        } else if (command.op === 'update') {
          const shape = find(command.id);
          if (shape) Object.assign(shape, clone(command.properties));
        }
      }
      return options.readback === false ? { applied: commands.length } : {
        presentationId: 'deck', slideId: 'slide-1', slideIndex: 0, width: 720, height: 540,
        selectionIds: [...state.selection], shapes: clone(state.shapes) };
    }
  };
  return { host, state, originals, keptPicture, keptGroup, unrelated };
}

const addTitlesBelow = actions.find(action => action.id === 'editing.addTitlesBelow');
const titleSettings = { fontName: 'Arial', fontSize: 12, distance: 4, text: 'Panel', autoGroup: true, center: true };

function assertTitleBatchRolledBack({ state, originals, keptPicture, keptGroup, unrelated }) {
  assert.deepEqual(markerTags(state.shapes), [], 'no title or group from the failed run survives');
  assert.deepEqual(state.shapes.filter(shape => shape.type === 'group' && shape.tags.SLIDESCI_TITLE_BATCH), [],
    'created groups are gone');
  for (const original of originals) assert.ok(flatten(state.shapes).includes(original), `${original.id} is the same object`);
  assert.ok(state.shapes.includes(keptGroup), 'preexisting group stays top-level');
  assert.ok(keptGroup.children.includes(keptPicture), 'preexisting group keeps its child');
  assert.ok(flatten(state.shapes).includes(unrelated), 'unrelated shape is untouched');
}

test('addTitles ungroups a partially applied batch and removes only this run\'s titles', async () => {
  const fixture = titleHost({ groupAfter: 1, recreateIds: true });
  await assert.rejects(addTitlesBelow.run(newCtx(fixture.host), { ...titleSettings }), /模拟编组失败/);
  assertTitleBatchRolledBack(fixture);
  assert.equal(fixture.state.shapes.filter(shape => shape.type === 'group').length, 1, 'only the preexisting group remains');
});

test('addTitles reclaims a group whose tag write failed by its tagged title', async () => {
  const fixture = titleHost({ groupTagFails: true, recreateIds: true });
  await assert.rejects(addTitlesBelow.run(newCtx(fixture.host), { ...titleSettings }), /模拟编组标记失败/);
  assertTitleBatchRolledBack(fixture);
  assert.equal(fixture.state.shapes.filter(shape => shape.type === 'group').length, 1, 'only the preexisting group remains');
});

test('addTitles removes titles inserted before a batched insertion failed', async () => {
  const fixture = titleHost({ insertAfter: 2 });
  await assert.rejects(addTitlesBelow.run(newCtx(fixture.host), { ...titleSettings }), /模拟插入失败/);
  assertTitleBatchRolledBack(fixture);
  assert.equal(fixture.state.shapes.filter(shape => shape.type === 'text').length, 1, 'only the unrelated text remains');
});

test('addTitles treats a rejected final selection as a failure of the whole batch', async () => {
  const fixture = titleHost({ selectFail: true });
  await assert.rejects(addTitlesBelow.run(newCtx(fixture.host), { ...titleSettings, autoGroup: false }), /模拟选择失败/);
  assertTitleBatchRolledBack(fixture);
  assert.deepEqual(fixture.state.selection, fixture.originals.map(shape => shape.id));
});

class MockElement {
  constructor(tag) { this.tagName = tag; this.children = []; this.events = new Map(); this.attributes = {}; this.value = ''; }
  append(...nodes) { this.children.push(...nodes); if (this.tagName === 'select' && !this.value) this.value = this.children[0]?.value || ''; }
  replaceChildren(...nodes) { this.children = nodes; }
  setAttribute(key, value) { this.attributes[key] = value; }
  addEventListener(name, listener) { this.events.set(name, listener); }
  dispatch(name) { this.events.get(name)?.({ preventDefault() {} }); }
}
const descendants = root => [root, ...root.children.filter(child => child instanceof MockElement).flatMap(descendants)];

// Light selection-time geometry is frozen; rich reads reflect edits already applied on the host.
function liveHost({ mixed = true, richSlide = 'slide-1' } = {}) {
  const shapes = mixed
    ? [{ id: 'img-1', name: 'Fig1', type: 'picture', left: 20, top: 40, width: 120, height: 90, rotation: 0 },
      { id: 'tb-1', name: 'Text1', type: 'text', left: 40, top: 180, width: 100, height: 40, rotation: 0 }]
    : [{ id: 'img-1', name: 'Fig1', type: 'picture', left: 20, top: 40, width: 120, height: 90, rotation: 0 }];
  const baseline = { presentationId: 'deck', slideId: 'slide-1', slideIndex: 0, width: 720, height: 540,
    selectionIds: shapes.map(shape => shape.id), shapes: clone(shapes) };
  const current = clone(shapes);
  if (mixed) current[1].font = { name: 'Arial', size: 12 };
  const applies = [];
  let richReads = 0;
  const host = {
    async context(options = {}) {
      if (options.geometry) return clone(baseline);
      richReads++;
      return { ...clone(baseline), slideId: richSlide, shapes: clone(current) };
    },
    onChanged() { return () => {}; },
    async apply(commands) {
      applies.push(clone(commands));
      for (const command of commands) {
        if (command.op !== 'update') continue;
        const shape = current.find(item => String(item.id) === String(command.id));
        if (shape) Object.assign(shape, clone(command.properties));
      }
      return { applied: commands.length };
    }
  };
  return { host, baseline, current, applies, richReads: () => richReads };
}

async function mountPanel(host) {
  const errors = [];
  let refreshed = 0;
  const waiters = [];
  let failed;
  const failure = new Promise(resolve => { failed = resolve; });
  const ctx = {
    host,
    settings: { get: async () => ({}), set: async () => {} },
    notify(message, kind) { if (kind === 'error') { errors.push(message); failed(); } },
    async refresh() {
      refreshed++;
      for (let i = waiters.length - 1; i >= 0; i--) if (refreshed >= waiters[i].target) {
        waiters[i].resolve(); waiters.splice(i, 1);
      }
    }
  };
  const root = new MockElement('div');
  const cleanup = await panels[0].mount(root, ctx);
  const waitFor = target => refreshed >= target ? Promise.resolve() : new Promise(resolve => waiters.push({ target, resolve }));
  const slider = label => descendants(root).find(node => node.type === 'range' && node.attributes['aria-label'] === label);
  return { root, cleanup, errors, waitFor, slider, failure };
}

test('live zoom keeps selection-time geometry as baseline after spacing and restores it at 100%', async t => {
  const previous = globalThis.document;
  globalThis.document = { createElement: tag => new MockElement(tag) };
  t.after(() => { if (previous === undefined) delete globalThis.document; else globalThis.document = previous; });
  const { host, current, applies, richReads } = liveHost();
  const panel = await mountPanel(host);
  t.after(panel.cleanup);

  const spacing = panel.slider('水平间距 (cm)');
  spacing.value = '1'; spacing.dispatch('input');
  await panel.waitFor(1);
  assert.equal(applies.length, 1);
  assert.ok(current[1].left > 40, 'spacing moved the text away from its selection-time position');

  const zoom = panel.slider('图文比例 (%)');
  zoom.value = '200'; zoom.dispatch('input');
  await panel.waitFor(2);
  const scaled = applies[1];
  const update = id => scaled.find(command => String(command.id) === id).properties;
  assert.deepEqual([update('img-1').left, update('img-1').top, update('img-1').width, update('img-1').height],
    [20, 40, 240, 180], 'scaling uses the selection bounds, not the spaced geometry');
  assert.deepEqual([update('tb-1').left, update('tb-1').top, update('tb-1').width, update('tb-1').height],
    [60, 320, 200, 80]);
  assert.equal(update('tb-1').font.size, 24, 'rich font data is added lazily');

  zoom.value = '100'; zoom.dispatch('input');
  await panel.waitFor(3);
  const restored = applies[2];
  const reset = id => restored.find(command => String(command.id) === id).properties;
  assert.deepEqual([reset('img-1').left, reset('img-1').top, reset('img-1').width, reset('img-1').height],
    [20, 40, 120, 90], '100% restores the geometry captured when the selection was made');
  assert.deepEqual([reset('tb-1').left, reset('tb-1').top, reset('tb-1').width, reset('tb-1').height],
    [40, 180, 100, 40], 'the earlier spacing edit does not become the new baseline');
  assert.equal(reset('tb-1').font.size, 12);
  assert.equal(richReads(), 1, 'rich text is read once per captured selection');
  assert.deepEqual(panel.errors, []);
});

test('live zoom for pictures only stays light and never requests the rich context', async t => {
  const previous = globalThis.document;
  globalThis.document = { createElement: tag => new MockElement(tag) };
  t.after(() => { if (previous === undefined) delete globalThis.document; else globalThis.document = previous; });
  const { host, richReads } = liveHost({ mixed: false });
  const panel = await mountPanel(host);
  t.after(panel.cleanup);
  const zoom = panel.slider('图文比例 (%)');
  zoom.value = '200'; zoom.dispatch('input');
  await panel.waitFor(1);
  assert.equal(richReads(), 0, 'pictures never trigger a rich read');
  assert.equal(panel.errors.length, 0);
});

test('live zoom refuses a rich read from a different page instead of adopting its geometry', async t => {
  const previous = globalThis.document;
  globalThis.document = { createElement: tag => new MockElement(tag) };
  t.after(() => { if (previous === undefined) delete globalThis.document; else globalThis.document = previous; });
  const { host, applies } = liveHost({ richSlide: 'slide-2' });
  const panel = await mountPanel(host);
  t.after(panel.cleanup);
  const zoom = panel.slider('图文比例 (%)');
  zoom.value = '200'; zoom.dispatch('input');
  await panel.failure;
  assert.deepEqual(applies, [], 'the stale rich read is not applied to the baseline selection');
  assert.match(panel.errors[0], /选择或页面已变化/);
});

test('a selection change during live scale replaces the baseline before the next slider write', async t => {
  const previous = globalThis.document;
  globalThis.document = { createElement: tag => new MockElement(tag) };
  t.after(() => { if (previous === undefined) delete globalThis.document; else globalThis.document = previous; });
  const { host, baseline, applies } = liveHost({ mixed: false });
  let onChanged;
  host.onChanged = handler => { onChanged = handler; return () => {}; };
  host.summary = async () => ({ ...baseline });
  const apply = host.apply;
  host.apply = async commands => {
    const result = await apply(commands);
    if (applies.length === 1) {
      baseline.slideId = 'slide-2';
      baseline.selectionIds = ['img-2'];
      baseline.shapes[0] = { ...baseline.shapes[0], id: 'img-2', width: 80 };
      onChanged();
    }
    return result;
  };
  const panel = await mountPanel(host);
  t.after(panel.cleanup);
  const zoom = panel.slider('图文比例 (%)');
  zoom.value = '150'; zoom.dispatch('input');
  await panel.waitFor(1);
  zoom.value = '200'; zoom.dispatch('input');
  await panel.waitFor(2);
  assert.deepEqual(applies.map(([command]) => [command.id, command.properties.width]), [['img-1', 180], ['img-2', 160]]);
});
