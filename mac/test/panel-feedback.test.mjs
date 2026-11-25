import test from 'node:test';
import assert from 'node:assert/strict';
import { automationMessage } from '../lib/native.mjs';

// Just enough DOM for the science panels to build their controls.
class MockElement {
  constructor(tag) { this.tagName = tag; this.children = []; this.attributes = {}; this.value = ''; this.disabled = false; }
  append(...nodes) { this.children.push(...nodes); }
  insertBefore(node, before) { const i = this.children.indexOf(before); this.children.splice(i < 0 ? this.children.length : i, 0, node); }
  setAttribute(key, value) { this.attributes[key] = value; }
  addEventListener() {}
  getContext() { return stubContext; }
}
// A drawing context that accepts any call or property, since only the layout is checked here.
const stubContext = new Proxy(function () {}, { get: (target, key) => key === Symbol.toPrimitive ? () => 0 : stubContext, apply: () => stubContext, set: () => true });
const text = node => node.textContent ?? '';

test('scale and zoom panels show their status under the buttons, and scale apply waits for a picture', async t => {
  const previous = globalThis.document;
  globalThis.document = { createElement: tag => new MockElement(tag) };
  t.after(() => { if (previous === undefined) delete globalThis.document; else globalThis.document = previous; });
  const { panels } = await import('../src/features/science.js');
  const ctx = { host: {}, services: {}, settings: { get: (key, fallback) => fallback, set() {} } };
  for (const id of ['science.scale', 'science.zoom']) {
    const root = new MockElement('div');
    await panels.find(panel => panel.id === id).mount(root, ctx);
    const [toolbar, status] = root.children;
    assert.ok(toolbar.children.some(node => node.tagName === 'button'), `${id} starts with its buttons`);
    assert.equal(status.tagName, 'p', `${id} shows messages right under the buttons, above the canvases`);
    assert.equal(status.attributes.role ?? 'status', 'status');
  }
  const root = new MockElement('div');
  await panels.find(panel => panel.id === 'science.scale').mount(root, ctx);
  const apply = root.children[0].children.find(node => text(node) === '应用到所选图片');
  assert.equal(apply.disabled, true, 'apply is disabled until a picture has been read');
});

test('automation errors name the process macOS asks about', () => {
  const denied = 'execution error: Not authorized to send Apple events to Microsoft PowerPoint. (-1743)';
  const timedOut = 'execution error: Microsoft PowerPoint got an error: AppleEvent timed out. (-1712)';
  assert.match(automationMessage(denied, 'launchd'), /为“node”（SlideSCI 后台服务）打开 Microsoft PowerPoint/);
  assert.match(automationMessage(denied, ''), /为启动 SlideSCI 服务的终端打开 Microsoft PowerPoint/);
  assert.match(automationMessage(timedOut, 'launchd'), /“node想要控制 Microsoft PowerPoint”.*3 分钟/);
  assert.match(automationMessage(timedOut, undefined), /“终端想要控制 Microsoft PowerPoint”/);
  assert.equal(automationMessage('PowerPoint got an error: Can’t get slide 9.', 'launchd'), null, 'other errors keep their own text');
});
