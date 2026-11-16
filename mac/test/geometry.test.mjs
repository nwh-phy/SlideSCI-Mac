import test from 'node:test';
import assert from 'node:assert/strict';
import { align, capturePositions, pastePositions } from '../src/geometry.js';
const shapes = [{ left: 80, top: 120 }, { left: 200, top: 40 }];

test('multiple alignment keeps the reference position and leaves the other axis alone', () => {
  assert.deepEqual(align(shapes, 'left'), [{ left: 80 }, { left: 80 }]);
  assert.deepEqual(shapes, [{ left: 80, top: 120 }, { left: 200, top: 40 }]);
});
test('single shape aligns to page edge', () => {
  assert.deepEqual(align([shapes[0]], 'top'), [{ top: 0 }]);
});
test('one copied position applies to every target', () => {
  assert.deepEqual(pastePositions(shapes, [{ left: 15, top: 25 }]),
    [{ left: 15, top: 25 }, { left: 15, top: 25 }]);
});
test('multiple copied positions preserve pairing and reject mismatches before mutation', () => {
  assert.deepEqual(pastePositions(shapes, capturePositions(shapes)), shapes);
  assert.throws(() => pastePositions([shapes[0]], capturePositions(shapes)), /目标数量/);
});
test('copied coordinates survive later changes to source objects', () => {
  const source = [{ left: 10, top: 20 }];
  const copied = capturePositions(source);
  source[0].left = 90;
  assert.deepEqual(copied, [{ left: 10, top: 20 }]);
});
test('empty selection and empty clipboard fail with actionable messages', () => {
  assert.throws(() => align([], 'left'), /选中形状/);
  assert.throws(() => capturePositions([]), /选中形状/);
  assert.throws(() => pastePositions(shapes), /先复制位置/);
});
