export const POINTS_PER_CM = 72 / 2.54;
export const anchors = Object.freeze({
  topLeft: [0, 0], topCenter: [0.5, 0], topRight: [1, 0],
  middleLeft: [0, 0.5], center: [0.5, 0.5], middleRight: [1, 0.5],
  bottomLeft: [0, 1], bottomCenter: [0.5, 1], bottomRight: [1, 1],
});

export function number(value, label, { min = -Infinity, max = Infinity, integer = false } = {}) {
  if (value === '' || value == null || typeof value === 'boolean') throw new Error(`请输入${label}。`);
  const result = Number(value);
  if (!Number.isFinite(result) || result < min || result > max || (integer && !Number.isInteger(result))) {
    throw new Error(`${label}必须是${integer ? '整数' : '有效数字'}，范围 ${min} 至 ${max}。`);
  }
  return result;
}

export function cm(value, label) {
  if (value == null || String(value).trim() === '') return undefined;
  return number(String(value).replace(/cm|厘米|\s/gi, ''), label, { min: Number.MIN_VALUE }) * POINTS_PER_CM;
}

export function point(shape, anchor) {
  const factors = anchors[anchor];
  if (!factors) throw new Error('请选择九宫格中的一个定位点。');
  return { x: shape.left + shape.width * factors[0], y: shape.top + shape.height * factors[1] };
}

export function place(shape, anchor, target) {
  const factors = anchors[anchor];
  if (!factors) throw new Error('定位点无效。');
  return { left: target.x - shape.width * factors[0], top: target.y - shape.height * factors[1] };
}

export function update(shape, properties) {
  return { op: 'update', id: shape.id, properties };
}

export function flatten(shapes) {
  return shapes.flatMap(shape => [shape, ...flatten(shape.children || [])]);
}

export function selected(snapshot, minimum = 1) {
  const all = new Map(flatten(snapshot.shapes).map(shape => [String(shape.id), shape]));
  const ids = snapshot.selectionIds?.length ? snapshot.selectionIds : snapshot.textSelection ? [snapshot.textSelection.shapeId] : [];
  const shapes = ids.map(id => all.get(String(id))).filter(Boolean);
  if (shapes.length < minimum) throw new Error(`请先选择至少 ${minimum} 个形状。`);
  return shapes;
}

export function bounds(shapes) {
  const left = Math.min(...shapes.map(shape => shape.left));
  const top = Math.min(...shapes.map(shape => shape.top));
  return { left, top, width: Math.max(...shapes.map(shape => shape.left + shape.width)) - left,
    height: Math.max(...shapes.map(shape => shape.top + shape.height)) - top };
}

export function ordered(shapes, order = 'selection') {
  if (order === 'selection') return [...shapes];
  if (order === 'horizontal') return [...shapes].sort((a, b) => a.left - b.left || a.top - b.top);
  if (order === 'vertical') return [...shapes].sort((a, b) => a.top - b.top || a.left - b.left);
  if (order !== 'position') throw new Error('排序方式无效。');
  const rows = [];
  for (const shape of shapes) {
    let row = rows.find(item => Math.min(item.bottom, shape.top + shape.height) - Math.max(item.top, shape.top) >= shape.height / 2);
    if (!row) { row = { top: shape.top, bottom: shape.top + shape.height, shapes: [] }; rows.push(row); }
    row.top = Math.min(row.top, shape.top);
    row.bottom = Math.max(row.bottom, shape.top + shape.height);
    row.shapes.push(shape);
  }
  return rows.sort((a, b) => a.top - b.top).flatMap(row => row.shapes.sort((a, b) => a.left - b.left));
}

export function grid(shapes, values) {
  const columns = number(values.columns, '列数', { min: 1, max: 10000, integer: true });
  const gapX = number(values.columnGap, '列间距 (pt)', { min: 0 });
  const gapY = number(values.rowGap, '行间距 (pt)', { min: 0 });
  const width = cm(values.widthCm, '宽度 (cm)');
  const height = cm(values.heightCm, '高度 (cm)');
  const items = ordered(shapes.filter(shape => !values.excludeText || !['text', 'shape', 'media'].includes(shape.type)), values.sort);
  if (!items.length) throw new Error('筛选后没有可以排列的对象。');
  for (const shape of items) {
    if (!(shape.width > 0 && shape.height > 0)) throw new Error(`“${shape.name}”的尺寸不能为零。`);
  }
  const start = items[0];
  const count = Math.min(columns, items.length);
  if (values.mode === 'waterfall') {
    const uniformWidth = width ?? start.width;
    const tops = Array(count).fill(start.top);
    return items.map(shape => {
      let column = 0;
      for (let i = 1; i < count; i++) if (tops[i] < tops[column]) column = i;
      const nextHeight = uniformWidth * shape.height / shape.width;
      const properties = { left: start.left + column * (uniformWidth + gapX), top: tops[column], width: uniformWidth, height: nextHeight };
      tops[column] += nextHeight + gapY;
      return update(shape, properties);
    });
  }
  if (!['columns', 'equalHeight'].includes(values.mode)) throw new Error('排列模式无效。');
  const sizes = items.map(shape => {
    if (width && height) return { width, height };
    if (width) return { width, height: width * shape.height / shape.width };
    if (height) return { width: height * shape.width / shape.height, height };
    if (values.mode === 'equalHeight') return { width: start.height * shape.width / shape.height, height: start.height };
    return { width: shape.width, height: shape.height };
  });
  const widths = Array(count).fill(0);
  sizes.forEach((size, i) => { widths[i % count] = Math.max(widths[i % count], size.width); });
  let left = start.left;
  let top = start.top;
  let rowHeight = 0;
  return items.map((shape, index) => {
    const column = index % count;
    if (index && !column) { top += rowHeight + gapY; left = start.left; rowHeight = 0; }
    const size = sizes[index];
    const command = update(shape, { ...size, left, top });
    rowHeight = Math.max(rowHeight, size.height);
    left += (values.mode === 'columns' ? widths[column] : size.width) + gapX;
    return command;
  });
}

export function alignment(snapshot, shapes, edge) {
  const reference = shapes.length === 1 ? { left: 0, top: 0, width: snapshot.width, height: snapshot.height } : shapes[0];
  return shapes.filter(shape => shapes.length === 1 || shape.id !== reference.id).map(shape => {
    switch (edge) {
      case 'left': return update(shape, { left: reference.left });
      case 'right': return update(shape, { left: reference.left + reference.width - shape.width });
      case 'top': return update(shape, { top: reference.top });
      case 'bottom': return update(shape, { top: reference.top + reference.height - shape.height });
      case 'horizontalCenter': return update(shape, { left: reference.left + (reference.width - shape.width) / 2 });
      case 'verticalCenter': return update(shape, { top: reference.top + (reference.height - shape.height) / 2 });
      default: throw new Error('对齐方式无效。');
    }
  });
}

export function spacing(shapes, axis, gap, distribute = false) {
  if (shapes.length < (distribute ? 3 : 2)) throw new Error(`请至少选择 ${distribute ? 3 : 2} 个形状。`);
  const position = axis === 'horizontal' ? 'left' : 'top';
  const dimension = axis === 'horizontal' ? 'width' : 'height';
  const items = [...shapes].sort((a, b) => a[position] - b[position]);
  const baseline = distribute ? 0 : items.findIndex(shape => shape.id === shapes[0].id);
  const positions = items.map(shape => shape[position]);
  const actualGap = distribute
    ? (positions.at(-1) + items.at(-1)[dimension] - positions[0] - items.reduce((sum, shape) => sum + shape[dimension], 0)) / (items.length - 1)
    : number(gap, '间距 (cm)', { min: -100, max: 100 }) * POINTS_PER_CM;
  for (let i = baseline + 1; i < items.length; i++) positions[i] = positions[i - 1] + items[i - 1][dimension] + actualGap;
  for (let i = baseline - 1; i >= 0; i--) positions[i] = positions[i + 1] - items[i][dimension] - actualGap;
  return items.flatMap((shape, i) => distribute && (i === 0 || i === items.length - 1) ? [] : [update(shape, { [position]: positions[i] })]);
}

function scaledFont(font, factor) {
  return font?.size > 0 ? { ...font, size: font.size * factor } : font;
}

export function scale(shapes, percent) {
  const factor = number(percent, '缩放比例 (%)', { min: 1, max: 10000 }) / 100;
  const origin = bounds(shapes);
  const commands = [];
  function text(shape) {
    const properties = {};
    if (shape.font?.size > 0) properties.font = scaledFont(shape.font, factor);
    if (shape.runs?.length) properties.runs = shape.runs.map(run => ({ ...run, font: scaledFont(run.font, factor) }));
    if (shape.paragraphs?.length) {
      properties.paragraphs = shape.paragraphs.map(paragraph => ({ ...paragraph,
        runs: paragraph.runs?.map(run => ({ ...run, font: scaledFont(run.font, factor) })) }));
      delete properties.runs;
    }
    return properties;
  }
  for (const shape of shapes) {
    commands.push(update(shape, { left: origin.left + (shape.left - origin.left) * factor,
      top: origin.top + (shape.top - origin.top) * factor, width: shape.width * factor, height: shape.height * factor, ...text(shape) }));
    for (const child of flatten(shape.children || [])) {
      const properties = text(child);
      if (Object.keys(properties).length) commands.push(update(child, properties));
    }
  }
  return commands;
}

export function fontSubset(font, option) {
  if (!font) throw new Error('所选对象没有可复制的文字格式。');
  const keys = {
    name: ['name', 'nameAscii', 'nameFarEast', 'nameComplexScript'], size: ['size'],
    color: ['color', 'transparency', 'outline', 'highlight'],
    effect: ['shadow', 'emboss', 'baselineOffset', 'subscript', 'superscript', 'glow', 'reflection'],
  };
  if (option === 'all') return structuredClone(font);
  if (!keys[option]) throw new Error('文字格式选项无效。');
  return Object.fromEntries([...keys[option], 'raw'].filter(key => font[key] !== undefined).map(key => [key, structuredClone(font[key])]));
}

export function labelAt(template, index) {
  const alphabets = { A: 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', a: 'abcdefghijklmnopqrstuvwxyz',
    'Ⅰ': 'ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ', '①': '①②③④⑤⑥⑦⑧⑨⑩', '一': '一二三四五六七八九十' };
  const key = template.replace(/[()]/g, '');
  if (key !== '1' && !alphabets[key]) throw new Error('标签模板无效。');
  const value = key === '1' ? String(index) : [...alphabets[key]][(index - 1) % [...alphabets[key]].length];
  return template.startsWith('(') ? `(${value})` : template.endsWith(')') ? `${value})` : value;
}
