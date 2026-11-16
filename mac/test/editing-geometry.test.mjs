import test from 'node:test';
import assert from 'node:assert/strict';
import {
  POINTS_PER_CM, anchors, point, place, selected, ordered, grid, alignment, spacing,
  scale, fontSubset, labelAt,
} from '../src/features/editing-geometry.js';

const shape = (id, properties = {}) => ({ id, name: id, type: 'picture', left: 0, top: 0, width: 100, height: 50, ...properties });
const gridSettings = { columns: 2, columnGap: 10, rowGap: 20, sort: 'selection', mode: 'columns', excludeText: true, widthCm: '', heightCm: '' };
const near = (actual, expected) => assert.ok(Math.abs(actual - expected) < 1e-8, `${actual} != ${expected}`);

test('each of nine anchor positions transfers between unequal rectangles without resizing', () => {
  const source = shape('source', { left: 30, top: 70 });
  const target = shape('target', { width: 35, height: 80 });
  for (const anchor of Object.keys(anchors)) {
    const origin = point(source, anchor);
    const moved = { ...target, ...place(target, anchor, origin) };
    assert.deepEqual(point(moved, anchor), origin);
    assert.equal(moved.width, target.width);
    assert.equal(moved.height, target.height);
  }
});

test('column grid uses column maxima and each row maximum after filtering', () => {
  const shapes = [shape('a', { left: 17, top: 23 }), shape('ignored', { type: 'text', width: 999 }),
    shape('b', { width: 40, height: 90 }), shape('c', { width: 150, height: 10 }), shape('d', { width: 20 })];
  const result = grid(shapes, gridSettings);
  assert.deepEqual(result.map(item => item.id), ['a', 'b', 'c', 'd']);
  assert.deepEqual(result.map(item => [item.properties.left, item.properties.top]), [[17, 23], [177, 23], [17, 133], [177, 133]]);
  assert.equal(shapes[1].left, 0);
});

test('centimeter sizing preserves ratio with one dimension and allows both dimensions', () => {
  const a = shape('a');
  const widthOnly = grid([a], { ...gridSettings, widthCm: '2.54 cm' })[0].properties;
  near(widthOnly.width, 72);
  near(widthOnly.height, 36);
  const both = grid([a], { ...gridSettings, widthCm: '2.54', heightCm: '2.54' })[0].properties;
  near(both.width, 72);
  near(both.height, 72);
  assert.throws(() => grid([a], { ...gridSettings, widthCm: '-1' }), /宽度/);
});

test('equal-height rows and waterfall each keep image proportions', () => {
  const items = [shape('a', { left: 7, top: 9, height: 100 }), shape('b', { width: 100, height: 20 }), shape('c', { width: 50, height: 25 })];
  const equal = grid(items, { ...gridSettings, mode: 'equalHeight' });
  assert.deepEqual(equal.map(item => [item.properties.width, item.properties.height]), [[100, 100], [500, 100], [200, 100]]);
  assert.equal(equal[2].properties.top, 129);
  const waterfall = grid(items, { ...gridSettings, mode: 'waterfall' });
  assert.deepEqual(waterfall.map(item => [item.properties.left, item.properties.top]), [[7, 9], [117, 9], [117, 49]]);
  assert.equal(waterfall[2].properties.height, 50);
});

test('position ordering groups overlapping rows rather than sorting tiny top offsets', () => {
  const right = shape('right', { left: 130, top: 10 });
  const left = shape('left', { left: 10, top: 12 });
  const lower = shape('lower', { left: 0, top: 100 });
  assert.deepEqual(ordered([right, lower, left], 'position').map(item => item.id), ['left', 'right', 'lower']);
  assert.deepEqual(ordered([right, lower, left], 'selection').map(item => item.id), ['right', 'lower', 'left']);
});

test('spacing keeps first-selected interior object fixed and supports overlap', () => {
  const middle = shape('middle', { left: 150, width: 70 });
  const left = shape('left', { left: 0, width: 30 });
  const right = shape('right', { left: 300, width: 40 });
  const commands = spacing([middle, left, right], 'horizontal', -1);
  const positions = new Map(commands.map(command => [command.id, command.properties.left]));
  assert.equal(positions.get('middle'), 150);
  near(positions.get('left'), 150 - 30 + POINTS_PER_CM);
  near(positions.get('right'), 220 - POINTS_PER_CM);
});

test('distribution fixes endpoints with unequal widths', () => {
  const commands = spacing([shape('a', { left: 10, width: 30 }), shape('b', { left: 100, width: 70 }), shape('c', { left: 300, width: 20 })], 'horizontal', 0, true);
  assert.deepEqual(commands, [{ op: 'update', id: 'b', properties: { left: 135 } }]);
});

test('all six alignments use slide bounds for one object and first-selected for a group', () => {
  const snapshot = { width: 720, height: 540 };
  const target = shape('target', { width: 100, height: 50 });
  const expectations = { left: { left: 0 }, right: { left: 620 }, top: { top: 0 }, bottom: { top: 490 }, horizontalCenter: { left: 310 }, verticalCenter: { top: 245 } };
  for (const [edge, expected] of Object.entries(expectations)) assert.deepEqual(alignment(snapshot, [target], edge)[0].properties, expected);
  const reference = shape('reference', { left: 80, top: 40, width: 200, height: 120 });
  assert.deepEqual(alignment(snapshot, [reference, target], 'right'), [{ op: 'update', id: 'target', properties: { left: 180 } }]);
});

test('joint scaling preserves mixed runs and hyperlinks in nested groups', () => {
  const child = shape('child', { type: 'text', text: 'abc', runs: [{ text: 'a', font: { size: 12, italic: true } }, { text: 'bc', font: { size: 20 }, hyperlink: 'https://example.com' }] });
  const group = shape('group', { type: 'group', left: 50, top: 40, children: [child] });
  const image = shape('image', { left: 200, top: 140 });
  const result = scale([group, image], 150);
  assert.deepEqual(result[0].properties, { left: 50, top: 40, width: 150, height: 75 });
  assert.deepEqual(result[1].properties.runs.map(run => run.font.size), [18, 30]);
  assert.equal(result[1].properties.runs[1].hyperlink, 'https://example.com');
  assert.equal(result[2].properties.left, 275);
  assert.equal(child.runs[0].font.size, 12);
});

test('selected text resolves native nested shape even without shape selection IDs', () => {
  const child = shape('child', { type: 'text' });
  const snapshot = { shapes: [shape('group', { children: [child] })], selectionIds: [], textSelection: { shapeId: 'child', start: 1, length: 3 } };
  assert.deepEqual(selected(snapshot), [child]);
});

test('selective font color contains outlines and highlight without changing font name or size', () => {
  const font = { name: 'Arial', size: 14, color: '#FF0000', transparency: 0.2, outline: { visible: false }, highlight: null, glow: { radius: 4 } };
  assert.deepEqual(fontSubset(font, 'color'), { color: '#FF0000', transparency: 0.2, outline: { visible: false }, highlight: null });
  assert.deepEqual(fontSubset(font, 'effect'), { glow: { radius: 4 } });
});

test('labels support all bracket forms, unlimited decimal numbers, and cyclic symbol alphabets', () => {
  assert.equal(labelAt('(A)', 27), '(A)');
  assert.equal(labelAt('a)', 28), 'b)');
  assert.equal(labelAt('1)', 121), '121)');
  assert.equal(labelAt('Ⅰ', 8), 'Ⅷ');
  assert.equal(labelAt('①)', 11), '①)');
  assert.equal(labelAt('一', 12), '二');
});
