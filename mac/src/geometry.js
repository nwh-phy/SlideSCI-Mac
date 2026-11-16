export function requireSelection(shapes) {
  if (!shapes.length) throw new Error('请先在幻灯片中选中形状。');
}

export function align(shapes, axis) {
  requireSelection(shapes);
  if (axis !== 'left' && axis !== 'top') throw new Error('不支持的对齐方向。');
  const position = shapes.length === 1 ? 0 : shapes[0][axis];
  return shapes.map(() => ({ [axis]: position }));
}

export function capturePositions(shapes) {
  requireSelection(shapes);
  return shapes.map(({ left, top }) => ({ left, top }));
}

export function pastePositions(shapes, positions) {
  requireSelection(shapes);
  if (!positions?.length) throw new Error('请先复制位置。');
  if (positions.length !== 1 && positions.length !== shapes.length) {
    throw new Error('多对象粘贴时，目标数量必须与复制数量一致。');
  }
  return shapes.map((_, index) => ({ ...positions[positions.length === 1 ? 0 : index] }));
}
