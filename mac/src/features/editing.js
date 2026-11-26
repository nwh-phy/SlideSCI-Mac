import {
  POINTS_PER_CM, number, point, place, update, flatten, selected, bounds,
  ordered, grid, spacing, scale, fontSubset, labelAt,
} from './editing-geometry.js';

const geometryKeys = new Set(['left', 'top', 'width', 'height', 'rotation']);
const section = '排版与编辑';
const session = new WeakMap();
const selectField = (name, label, options, value) => ({ name, label, type: 'select', options: options.map(([value, label]) => ({ value, label })), default: value });
const numericField = (name, label, value, min, max, step = 1) => ({ name, label, type: 'number', default: value, min, max, step });
const checkField = (name, label, value) => ({ name, label, type: 'checkbox', default: value });
const textField = (name, label, value = '') => ({ name, label, type: 'text', default: value });
const anchorOptions = [
  ['topLeft', '左上'], ['topCenter', '上中'], ['topRight', '右上'],
  ['middleLeft', '左中'], ['center', '中心'], ['middleRight', '右中'],
  ['bottomLeft', '左下'], ['bottomCenter', '下中'], ['bottomRight', '右下'],
];
const orderOptions = [['horizontal', '水平位置顺序'], ['vertical', '垂直位置顺序'], ['selection', '点选顺序']];
const anchorField = value => selectField('anchor', '定位点', anchorOptions, value);
const fontOptions = [['all', '全部文字格式'], ['name', '仅字体'], ['color', '颜色、描边和高亮'], ['size', '仅字号'], ['effect', '阴影、浮雕、上下标、发光和倒影']];
const shapeOptions = [['all', '全部形状格式'], ['fill', '仅填充'], ['line', '仅轮廓']];
const gridFields = [
  selectField('mode', '排列模式', [['columns', '列对齐排列'], ['equalHeight', '统一高度排列'], ['waterfall', '统一宽度瀑布流']], 'columns'),
  selectField('sort', '排序', [['position', '根据位置排序'], ['selection', '根据选择顺序']], 'position'),
  numericField('columns', '列数', 3, 1, 10000), numericField('columnGap', '列间距 (pt)', 10, 0, undefined, 0.1),
  numericField('rowGap', '行间距 (pt)', 25, 0, undefined, 0.1), textField('widthCm', '统一宽度 (cm，空白保留比例)'),
  textField('heightCm', '统一高度 (cm，空白保留比例)'), checkField('excludeText', '排除文本框、形状和媒体', true),
];
const titleFields = [
  textField('fontName', '字体', '微软雅黑'), numericField('fontSize', '字号 (pt)', 14, 0.01, undefined, 0.1),
  numericField('distance', '图片与标题距离 (pt)', 0, undefined, undefined, 0.1), textField('text', '标题占位文字', '图片标题'),
  checkField('autoGroup', '自动编组图片与标题', false), checkField('center', '标题居中', true),
];
const labelFields = [
  textField('fontName', '字体', 'Arial'), numericField('fontSize', '字号 (pt)', 12, 0.01, undefined, 0.1),
  selectField('template', '标签模板', ['A', 'a', 'A)', 'a)', '(A)', '(a)', '1', '1)', 'Ⅰ', 'Ⅰ)', '①)', '①', '一)', '一'].map(value => [value, value]), 'A'),
  numericField('offsetX', 'X 偏移 (pt)', -20, undefined, undefined, 0.1), numericField('offsetY', 'Y 偏移 (pt)', -7, undefined, undefined, 0.1),
  numericField('startIndex', '起始编号', 1, 1, Number.MAX_SAFE_INTEGER), checkField('bold', '标签加粗', true), checkField('advanceIndex', '添加标签后自动更新起始编号', false),
];

function state(host) {
  if (!session.has(host)) session.set(host, {});
  return session.get(host);
}

function defaults(fields) {
  return Object.fromEntries(fields.map(field => [field.name, field.default]));
}

function makeAction(id, label, fields, execute, description) {
  const settingsKey = id === 'addTitlesAbove' || id === 'addTitlesBelow' ? 'editing.titles'
    : id === 'addLabels' || id === 'updateLabels' ? 'editing.labels' : `editing.${id}`;
  return { id: `editing.${id}`, label, section, description, fields, settingsKey,
    async run(ctx, values = {}) {
      const key = settingsKey;
      const saved = await ctx.settings.get(key, {});
      const effective = { ...defaults(fields), ...saved, ...values };
      const result = await execute(ctx, effective);
      if (fields.length) await ctx.settings.set(key, effective);
      return result;
    } };
}

async function selection(ctx, minimum = 1, geometry = false) {
  const snapshot = await ctx.host.context({ scope: 'selection', ...(geometry ? { geometry: true } : {}) });
  return { snapshot, shapes: selected(snapshot, minimum) };
}

async function apply(ctx, commands) {
  if (!commands.length) return ctx.host.context({ scope: 'selection' });
  const geometry = commands.every(command => command.slideIndex === undefined && (command.op === 'select' || command.op === 'update'
    && Object.keys(command.properties || {}).every(key => geometryKeys.has(key))));
  return ctx.host.apply(commands, geometry ? { geometry: true, readback: false } : { geometry: false });
}

function copied(value, label) {
  if (!value) throw new Error(`请先复制${label}。`);
  return value;
}

function nativeFont(shape) {
  return shape.font ?? shape.runs?.find(run => run.font)?.font ?? shape.paragraphs?.flatMap(paragraph => paragraph.runs || []).find(run => run.font)?.font;
}

function selectionFont(snapshot, shapes) {
  const selection = snapshot.textSelection;
  if (!selection) return nativeFont(shapes[0]);
  const shape = flatten(snapshot.shapes || []).find(shape => shape.id === selection.shapeId);
  let parsed = nativeFont(shape || shapes[0]);
  let offset = 0;
  for (const run of shape?.runs || []) {
    const length = String(run.text || '').length;
    if (selection.start >= offset && selection.start < offset + length) {
      parsed = { ...parsed, ...run.font }; break;
    }
    offset += length;
  }
  const native = selection.font || {};
  const result = { ...parsed, ...native };
  for (const key of ['reflection', 'glow', 'outline']) {
    if (parsed?.[key] && native[key]) result[key] = { ...parsed[key], ...native[key] };
  }
  return result;
}

function newShapes(before, after) {
  const existing = new Set(before.shapes.map(shape => String(shape.id)));
  return after.shapes.filter(shape => !existing.has(String(shape.id)));
}

async function rollback(ctx, ids, originalError) {
  if (ids.length) {
    try { await ctx.host.apply([{ op: 'delete', ids }]); }
    catch (cleanupError) { throw new AggregateError([originalError, cleanupError], `${originalError.message}；无法撤销本次新增对象：${cleanupError.message}`); }
  }
  throw originalError;
}

const definitions = [
  makeAction('grid', '图片自动排列', gridFields, async (ctx, values) => {
    const { shapes } = await selection(ctx, 1, true);
    return apply(ctx, grid(shapes, values));
  }, '三种排列模式；间距单位为 pt，统一尺寸单位为 cm。仅设置一项尺寸时保持宽高比。瀑布流按统一宽度保持各图比例。'),
  ...[
    ['left', '左对齐'], ['right', '右对齐'], ['top', '顶对齐'], ['bottom', '底对齐'],
    ['horizontalCenter', '水平居中'], ['verticalCenter', '垂直居中'],
  ].map(([edge, label]) => makeAction(`align.${edge}`, label, [], ctx => ctx.host.align(edge),
    '多选时以最先选中的对象为基准；单选时对齐幻灯片。')),
  ...['horizontal', 'vertical'].flatMap(axis => [
    makeAction(`spacing.${axis}`, axis === 'horizontal' ? '设置水平间距' : '设置垂直间距', [numericField('gapCm', '间距 (cm)', 0.5, -100, 100, 0.01)], async (ctx, values) => {
      const { shapes } = await selection(ctx, 2, true);
      return apply(ctx, spacing(shapes, axis, values.gapCm));
    }, '最先选中的形状保持原位，负间距允许对象重叠。'),
    makeAction(`distribute.${axis}`, axis === 'horizontal' ? '水平均匀分布' : '垂直均匀分布', [], async ctx => {
      const { shapes } = await selection(ctx, 3, true);
      return apply(ctx, spacing(shapes, axis, 0, true));
    }, '保留两端位置，均匀分配对象边缘之间的间距。'),
  ]),
  makeAction('copyPosition', '复制位置', [anchorField('center')], async (ctx, values) => {
    const { shapes } = await selection(ctx, 1, true);
    state(ctx.host).positions = { anchor: values.anchor, points: shapes.map(shape => point(shape, values.anchor)) };
    return { message: `已复制 ${shapes.length} 个位置。` };
  }),
  makeAction('pastePosition', '粘贴位置', [], async ctx => {
    const positions = copied(state(ctx.host).positions, '位置');
    const { shapes } = await selection(ctx, 1, true);
    return apply(ctx, shapes.slice(0, positions.points.length).map((shape, i) => update(shape, place(shape, positions.anchor, positions.points[i]))));
  }, '按选择顺序对应粘贴；超出已复制数量的目标保持原位。'),
  makeAction('swapPosition', '交换位置', [anchorField('topLeft')], async (ctx, values) => {
    const { shapes } = await selection(ctx, 2, true);
    if (shapes.length !== 2) throw new Error('交换位置时请只选择两个对象。');
    return apply(ctx, [update(shapes[0], place(shapes[0], values.anchor, point(shapes[1], values.anchor))),
      update(shapes[1], place(shapes[1], values.anchor, point(shapes[0], values.anchor)))]);
  }),
  makeAction('copyRelativePosition', '复制标注相对位置', [anchorField('topLeft'), selectField('order', '参考图与标注顺序', orderOptions, 'horizontal')], async (ctx, values) => {
    const { shapes } = await selection(ctx, 2);
    const items = ordered(shapes, values.order);
    const origin = point(items[0], values.anchor);
    const annotations = [];
    for (const shape of items.slice(1)) {
      const position = point(shape, values.anchor);
      annotations.push({ delta: { x: position.x - origin.x, y: position.y - origin.y }, shape: structuredClone(shape), fragment: await ctx.host.fragment([shape.id]) });
    }
    state(ctx.host).relative = { anchor: values.anchor, order: values.order, annotations };
    return { message: `已复制 ${annotations.length} 个标注的相对位置和可编辑对象。` };
  }, '排序后的第一个形状为参考图。缺少的目标标注会从原生可编辑副本中克隆，支持跨页。'),
  makeAction('pasteRelativePosition', '粘贴标注相对位置', [], async ctx => {
    const relative = copied(state(ctx.host).relative, '标注相对位置');
    const { snapshot, shapes } = await selection(ctx);
    const targets = ordered(shapes, relative.order);
    if (targets.length - 1 > relative.annotations.length) throw new Error('目标标注数量不能多于已复制的标注数量。');
    const origin = point(targets[0], relative.anchor);
    const annotations = targets.slice(1);
    const created = [];
    let current = snapshot;
    try {
      for (let i = annotations.length; i < relative.annotations.length; i++) {
        const source = relative.annotations[i];
        const position = place(source.shape, relative.anchor, { x: origin.x + source.delta.x, y: origin.y + source.delta.y });
        const next = await ctx.host.insertFragment(source.fragment, { ...position, width: source.shape.width, height: source.shape.height });
        const inserted = newShapes(current, next);
        created.push(...inserted.map(shape => shape.id));
        if (inserted.length !== 1) throw new Error('标注副本应包含一个顶层原生对象。');
        annotations.push(inserted[0]);
        current = next;
      }
      const commands = annotations.map((shape, i) => update(shape, place(shape, relative.anchor, { x: origin.x + relative.annotations[i].delta.x, y: origin.y + relative.annotations[i].delta.y })));
      commands.push({ op: 'select', ids: [targets[0].id, ...annotations.map(shape => shape.id)] });
      return await apply(ctx, commands);
    } catch (error) { return rollback(ctx, created, error); }
  }),
  ...[['dimensions', '宽高'], ['width', '宽度'], ['height', '高度']].flatMap(([kind, label]) => [
    makeAction(`copy.${kind}`, `复制${label}`, [], async ctx => {
      const { shapes } = await selection(ctx, 1, true);
      const clipboard = state(ctx.host);
      if (kind !== 'height') clipboard.width = shapes[0].width;
      if (kind !== 'width') clipboard.height = shapes[0].height;
      return { message: `已复制${label}。` };
    }),
    makeAction(`paste.${kind}`, `粘贴${label}`, [], async ctx => {
      const clipboard = state(ctx.host);
      const width = kind !== 'height' ? copied(clipboard.width, '宽度') : undefined;
      const height = kind !== 'width' ? copied(clipboard.height, '高度') : undefined;
      const { shapes } = await selection(ctx, 1, true);
      return apply(ctx, shapes.map(shape => {
        if (!(shape.width > 0 && shape.height > 0)) throw new Error('无法按比例调整零尺寸对象。');
        return update(shape, { width: width ?? shape.width * height / shape.height,
          height: height ?? shape.height * width / shape.width });
      }));
    }, kind === 'dimensions' ? '同时设置宽度和高度。' : '按原有宽高比缩放。'),
  ]),
  makeAction('copyCrop', '复制裁剪', [], async ctx => {
    const { shapes } = await selection(ctx);
    const shape = shapes[0];
    if (shape.type !== 'picture') throw new Error('请选择图片复制裁剪。');
    state(ctx.host).crop = { crop: structuredClone(shape.crop ?? { left: 0, top: 0, right: 0, bottom: 0 }), height: shape.height };
    return { message: '已复制裁剪和裁剪后高度。' };
  }),
  makeAction('pasteCrop', '粘贴裁剪', [], async ctx => {
    const source = copied(state(ctx.host).crop, '裁剪');
    const { shapes } = await selection(ctx);
    if (shapes.some(shape => shape.type !== 'picture')) throw new Error('粘贴裁剪时只选择图片。');
    return apply(ctx, shapes.map(shape => {
      const crop = shape.crop ?? { left: 0, top: 0, right: 0, bottom: 0 };
      const oldVisibleX = 1 - crop.left - crop.right;
      const oldVisibleY = 1 - crop.top - crop.bottom;
      const visibleX = 1 - source.crop.left - source.crop.right;
      const visibleY = 1 - source.crop.top - source.crop.bottom;
      if (Math.min(oldVisibleX, oldVisibleY, visibleX, visibleY, shape.height) <= 0) throw new Error('裁剪区域必须保留可见图像。');
      const originalRatio = shape.width / shape.height * oldVisibleY / oldVisibleX;
      return update(shape, { crop: source.crop, height: source.height, width: source.height * originalRatio * visibleX / visibleY, left: shape.left, top: shape.top });
    }));
  }),
  makeAction('copyShapeFormat', '复制形状格式', [selectField('option', '格式范围', shapeOptions, 'all')], async (ctx, values) => {
    const { shapes } = await selection(ctx);
    const source = shapes[0];
    if (values.option === 'all' && source.style === undefined) throw new Error('宿主未返回完整原生形状样式，请检查宿主样式接口。');
    const style = values.option === 'all' ? { ...structuredClone(source.style), fragment: await ctx.host.fragment([source.id]) } : undefined;
    state(ctx.host).shapeFormat = { option: values.option, style, fill: structuredClone(source.fill), line: structuredClone(source.line), font: structuredClone(nativeFont(source)) };
    return { message: '已复制形状格式。' };
  }),
  makeAction('pasteShapeFormat', '粘贴形状格式', [], async ctx => {
    const source = copied(state(ctx.host).shapeFormat, '形状格式');
    const { shapes } = await selection(ctx);
    return apply(ctx, shapes.map(shape => {
      const properties = source.option === 'all' ? { style: source.style } : { [source.option]: source[source.option] };
      if (source.font && nativeFont(shape)) {
        properties.fontScope = state(ctx.host).textFormat?.option ?? 'all';
        properties.font = fontSubset(source.font, properties.fontScope);
      }
      return update(shape, properties);
    }));
  }),
  makeAction('copyTextFormat', '复制文字格式', [selectField('option', '格式范围', fontOptions, 'all')], async (ctx, values) => {
    const { snapshot, shapes } = await selection(ctx);
    const font = fontSubset(selectionFont(snapshot, shapes), values.option);
    state(ctx.host).textFormat = { option: values.option, font };
    return { message: '已复制文字格式。' };
  }),
  makeAction('pasteTextFormat', '粘贴文字格式', [], async ctx => {
    const source = copied(state(ctx.host).textFormat, '文字格式');
    const { snapshot, shapes } = await selection(ctx);
    if (snapshot.textSelection) {
      const { shapeId, start, length } = snapshot.textSelection;
      return apply(ctx, [{ op: 'update', id: shapeId, properties: {
        fontScope: source.option, textRange: { start, length, font: source.font } } }]);
    }
    const targets = shapes.flatMap(shape => flatten([shape])).filter(shape => nativeFont(shape) || shape.text !== undefined);
    if (!targets.length) throw new Error('所选对象中没有可设置格式的文字。');
    return apply(ctx, targets.map(shape => update(shape, { font: source.font, fontScope: source.option })));
  }),
  makeAction('copyGroupFormat', '复制组格式', [], async ctx => {
    const { shapes } = await selection(ctx);
    state(ctx.host).groupFormat = await ctx.host.fragment(shapes.map(shape => shape.id));
    return { message: '已复制可编辑的组格式。' };
  }),
  makeAction('pasteGroupFormat', '粘贴组格式', [], async ctx => {
    const fragment = copied(state(ctx.host).groupFormat, '组格式');
    const { snapshot, shapes } = await selection(ctx);
    const origin = bounds(shapes);
    const textShapes = ordered(flatten(shapes).filter(shape => !shape.children?.length && shape.text), 'vertical');
    const inserted = await ctx.host.insertFragment(fragment, { left: origin.left, top: origin.top });
    const replacement = newShapes(snapshot, inserted);
    if (!replacement.length) throw new Error('组格式插入后未找到新增的原生对象。');
    try {
      const templateTexts = ordered(flatten(replacement).filter(shape => !shape.children?.length && shape.text), 'vertical');
      const split = textShapes.length < templateTexts.length && textShapes.some(shape => /[\r\n\v]/.test(shape.text));
      const pieces = textShapes.flatMap(shape => split ? shape.text.split(/\r\n|[\r\n\v]/).map(text => text.trim()).filter(Boolean) : [shape.text]);
      if (pieces.length && !templateTexts.length) throw new Error('复制的组格式不含文字位置，无法保留目标文字。');
      if (pieces.length > templateTexts.length) {
        pieces.splice(templateTexts.length - 1, pieces.length, pieces.slice(templateTexts.length - 1).join('\n'));
      }
      const commands = templateTexts.slice(0, pieces.length).map((shape, i) => update(shape, { text: pieces[i],
        ...(textShapes[i] ? { tags: textShapes[i].tags ?? {}, alternativeText: textShapes[i].alternativeText ?? '' } : {}) }));
      replacement.forEach((shape, i) => {
        const target = shapes[Math.min(i, shapes.length - 1)];
        commands.push(update(shape, { tags: target.tags ?? {}, alternativeText: target.alternativeText ?? '' }));
      });
      await apply(ctx, commands);
    } catch (error) { return rollback(ctx, replacement.map(shape => shape.id), error); }
    const orders = replacement.map((shape, i) => ({ shape, target: shapes[Math.min(i, shapes.length - 1)] }))
      .filter(item => Number.isInteger(item.target.zOrder))
      .sort((a, b) => a.target.zOrder - b.target.zOrder)
      .map(({ shape, target }) => ({ op: 'zOrder', id: shape.id, position: 'index', index: target.zOrder }));
    return apply(ctx, [{ op: 'delete', ids: shapes.map(shape => shape.id) }, ...orders,
      { op: 'select', ids: replacement.map(shape => shape.id) }]);
  }, '采用复制组的可编辑结构与格式，按位置把目标文字填入模板；当模板文本框更多时可按行分配。'),
  makeAction('scaleTogether', '图文同比缩放', [numericField('percent', '缩放比例 (%)', 100, 1, 10000)], async (ctx, values) => {
    const { shapes } = await selection(ctx);
    return apply(ctx, scale(shapes, values.percent));
  }, '以选择区域左上角为基准，同时缩放对象位置、宽高和各级组内文字字号。'),
  makeAction('selectAllText', '全选文本框', [], async ctx => {
    const snapshot = await ctx.host.context({ scope: 'slide' });
    const ids = snapshot.shapes.filter(shape => shape.type === 'text' || (shape.type === 'placeholder' && shape.text !== undefined) || (shape.isPlaceholder && shape.text !== undefined)).map(shape => shape.id);
    if (!ids.length) throw new Error('当前幻灯片没有文本框。');
    return apply(ctx, [{ op: 'select', ids }]);
  }),
];

async function addTitles(ctx, values, below) {
  const { snapshot, shapes } = await selection(ctx);
  const fontSize = number(values.fontSize, '字号', { min: 0.01 });
  const distance = number(values.distance, '标题距离');
  if (!String(values.fontName).trim()) throw new Error('请输入标题字体名称。');
  const created = [];
  const resultIds = [];
  let current = snapshot;
  try {
    for (const shape of shapes) {
      const top = below ? shape.top + shape.height + distance : shape.top - fontSize * 2 - distance;
      const next = await apply(ctx, [{ op: 'insertText', text: String(values.text), left: shape.left, top, width: shape.width, height: fontSize * 2,
        font: { name: values.fontName.trim(), nameFarEast: values.fontName.trim(), size: fontSize },
        paragraphs: [{ runs: [{ text: String(values.text) }], alignment: values.center ? 'center' : 'left' }],
        textFrame: { wordWrap: true, autoSize: 'shapeToFitText' }, name: `${shape.name} 标题` }]);
      const inserted = newShapes(current, next);
      created.push(...inserted.map(item => item.id));
      if (inserted.length !== 1) throw new Error('标题插入后未找到新增文本框。');
      current = next;
      if (values.autoGroup) {
        const grouped = await apply(ctx, [{ op: 'group', ids: [shape.id, inserted[0].id], name: `${shape.name} 与标题` }]);
        const old = new Set(current.shapes.map(item => String(item.id)));
        const group = grouped.shapes.find(item => !old.has(String(item.id)));
        if (!group) throw new Error('编组后未返回新增组。');
        resultIds.push(group.id);
        current = grouped;
      } else resultIds.push(inserted[0].id);
    }
    return apply(ctx, [{ op: 'select', ids: resultIds }]);
  } catch (error) {
    // A grouped original must remain intact if a later insertion fails.
    const live = await ctx.host.context({ scope: 'slide' });
    return rollback(ctx, created.filter(id => live.shapes.some(shape => String(shape.id) === String(id))), error);
  }
}

definitions.push(
  makeAction('addTitlesBelow', '添加下标题', titleFields, (ctx, values) => addTitles(ctx, values, true)),
  makeAction('addTitlesAbove', '添加上标题', titleFields, (ctx, values) => addTitles(ctx, values, false)),
  makeAction('addLabels', '添加图片标签', labelFields, async (ctx, values) => {
    const { snapshot, shapes } = await selection(ctx);
    const fontSize = number(values.fontSize, '标签字号', { min: 0.01 });
    const start = number(values.startIndex, '起始编号', { min: 1, max: Number.MAX_SAFE_INTEGER - shapes.length, integer: true });
    const offsetX = number(values.offsetX, 'X 偏移');
    const offsetY = number(values.offsetY, 'Y 偏移');
    if (!String(values.fontName).trim()) throw new Error('请输入标签字体名称。');
    const commands = ordered(shapes, 'position').map((shape, i) => ({ op: 'insertText', text: labelAt(values.template, start + i),
      left: shape.left + offsetX, top: shape.top + offsetY, width: Math.max(fontSize, 1), height: fontSize * 2,
      font: { name: values.fontName.trim(), nameFarEast: values.fontName.trim(), size: fontSize, bold: Boolean(values.bold) },
      paragraphs: [{ runs: [{ text: labelAt(values.template, start + i) }], alignment: 'left' }],
      textFrame: { wordWrap: false, autoSize: 'shapeToFitText' }, name: `${shape.name} 标签` }));
    const next = await apply(ctx, commands);
    const ids = newShapes(snapshot, next).map(shape => shape.id);
    await apply(ctx, [{ op: 'select', ids }]);
    // The advanced start index must reach the visible form, or the form's old value overwrites it.
    if (values.advanceIndex) { values.startIndex = start + shapes.length; return { message: `已添加 ${ids.length} 个标签，下次从 ${values.startIndex} 开始。`, formValues: { startIndex: values.startIndex } }; }
    return { message: `已添加 ${ids.length} 个标签。` };
  }),
  makeAction('updateLabels', '更新图片标签', labelFields, async (ctx, values) => {
    const { shapes } = await selection(ctx);
    const targets = ordered(shapes.filter(shape => shape.type === 'text'), 'position');
    if (!targets.length) throw new Error('请选择要更新标签的文本框。');
    const fontSize = number(values.fontSize, '标签字号', { min: 0.01 });
    const start = number(values.startIndex, '起始编号', { min: 1, max: Number.MAX_SAFE_INTEGER - targets.length, integer: true });
    if (!String(values.fontName).trim()) throw new Error('请输入标签字体名称。');
    const result = await apply(ctx, targets.map((shape, index) => update(shape, {
      text: labelAt(values.template, start + index),
      font: { name: values.fontName.trim(), nameFarEast: values.fontName.trim(), size: fontSize, bold: Boolean(values.bold) },
    })));
    if (values.advanceIndex) { values.startIndex = start + targets.length; return { message: `已更新 ${targets.length} 个标签，下次从 ${values.startIndex} 开始。`, formValues: { startIndex: values.startIndex } }; }
    return result;
  }),
);

export const actions = definitions;

function element(tag, text) {
  const node = document.createElement(tag);
  if (text !== undefined) node.textContent = text;
  return node;
}

function fieldControl(field, value) {
  let input;
  if (field.type === 'select') {
    input = element('select');
    for (const option of field.options) {
      const item = element('option', option.label);
      item.value = option.value;
      input.append(item);
    }
  } else {
    input = element('input');
    input.type = field.type;
    for (const key of ['min', 'max', 'step']) if (field[key] !== undefined) input[key] = field[key];
  }
  input.name = field.name;
  if (field.type === 'checkbox') input.checked = Boolean(value);
  else input.value = value ?? '';
  const label = element('label');
  label.append(element('span', field.label), input);
  return { label, input };
}

export const panels = [{
  id: 'editing.controls', label: '编辑设置与实时调整', section,
  async mount(container, ctx) {
    let disposed = false;
    let applying = false;
    let liveSnapshot;
    let signature = '';
    let activeOperation = Promise.resolve();
    let changePending = false;
    const pendingInputs = new Map();
    let liveScheduled = false;
    const chooser = element('select');
    chooser.setAttribute('aria-label', '编辑操作');
    for (const action of actions) {
      const option = element('option', action.label);
      option.value = action.id;
      chooser.append(option);
    }
    const form = element('form');
    const fields = element('div');
    const description = element('p');
    const run = element('button', '执行');
    run.type = 'submit';
    const save = element('button', '保存设置');
    save.type = 'button';
    form.append(fields, description, run, save);
    container.append(chooser, form);
    let controls = [];
    let renderId = 0;
    async function render() {
      const id = ++renderId;
      const action = actions.find(item => item.id === chooser.value);
      const saved = await ctx.settings.get(action.settingsKey, {});
      if (disposed || id !== renderId) return;
      controls = action.fields.map(field => ({ field, ...fieldControl(field, saved[field.name] ?? field.default) }));
      fields.replaceChildren(...controls.map(control => control.label));
      description.textContent = action.description ?? '';
      save.hidden = action.fields.length === 0;
    }
    chooser.addEventListener('change', render);
    save.addEventListener('click', async () => {
      if (!form.reportValidity()) return;
      const action = actions.find(item => item.id === chooser.value);
      const values = Object.fromEntries(controls.map(({ field, input }) => [field.name, field.type === 'checkbox' ? input.checked : input.value]));
      try {
        await ctx.settings.set(action.settingsKey, values);
        ctx.notify('设置已保存。', 'success');
      } catch (error) { ctx.notify(error.message, 'error'); }
    });
    form.addEventListener('submit', async event => {
      event.preventDefault();
      run.disabled = true;
      const action = actions.find(item => item.id === chooser.value);
      const values = Object.fromEntries(controls.map(({ field, input }) => [field.name, field.type === 'checkbox' ? input.checked : input.value]));
      try {
        const result = await action.run(ctx, values);
        ctx.notify(result?.message ?? `${action.label}完成。`, 'success');
        await ctx.refresh();
        await render();
      } catch (error) { ctx.notify(error.message, 'error'); }
      finally { run.disabled = false; }
    });
    await render();
    const live = element('fieldset');
    live.append(element('legend', '实时图文缩放与间距'));
    const status = element('p', '选择对象后调整。基准为选择时的几何与字号。');
    const reset = element('button', '重新读取选择');
    reset.type = 'button';
    const inputs = [];
    const liveFields = [numericField('scale', '图文比例 (%)', 100, 1, 500, 1),
      numericField('horizontal', '水平间距 (cm)', 0.5, -100, 100, 0.01), numericField('vertical', '垂直间距 (cm)', 0.5, -100, 100, 0.01)];
    live.append(status, reset);
    for (const field of liveFields) {
      const control = fieldControl(field, field.default);
      const slider = element('input');
      slider.type = 'range';
      slider.min = field.name === 'scale' ? '1' : '-5';
      slider.max = field.name === 'scale' ? '500' : '15';
      slider.step = field.name === 'scale' ? '1' : '0.1';
      slider.value = String(field.default);
      slider.setAttribute('aria-label', field.label);
      live.append(control.label, slider);
      inputs.push({ ...control, field, slider });
      const change = source => {
        if (source === slider) control.input.value = slider.value;
        else slider.value = control.input.value;
        pendingInputs.delete(field.name);
        pendingInputs.set(field.name, control.input.value);
        if (liveScheduled) return;
        liveScheduled = true;
        activeOperation = activeOperation.then(async () => {
          applying = true;
          try {
            while (!disposed && pendingInputs.size) {
              const [name, value] = pendingInputs.entries().next().value;
              pendingInputs.delete(name);
              if (!liveSnapshot) await capture();
              const shapes = selected(liveSnapshot, name === 'scale' ? 1 : 2);
              const commands = name === 'scale' ? scale(shapes, value) : spacing(shapes, name, value);
              await apply(ctx, commands);
              await ctx.settings.set(`editing.live.${name}`, Number(value));
            }
            if (!disposed) await ctx.refresh();
          } catch (error) { pendingInputs.clear(); ctx.notify(error.message, 'error'); }
          finally { applying = false; liveScheduled = false; }
        });
      };
      control.input.addEventListener('input', () => change(control.input));
      slider.addEventListener('input', () => change(slider));
    }
    container.append(live);
    async function capture() {
      const snapshot = await ctx.host.context({ scope: 'selection' });
      if (disposed) return;
      liveSnapshot = structuredClone(snapshot);
      signature = JSON.stringify([snapshot.presentationId, snapshot.slideId, snapshot.selectionIds]);
      const shapes = snapshot.selectionIds.length ? selected(snapshot) : [];
      status.textContent = `已选中 ${shapes.length} 个对象；最先选择的对象为间距基准。`;
      for (const item of inputs) {
        let value = 100;
        if (item.field.name !== 'scale') {
          const position = item.field.name === 'horizontal' ? 'left' : 'top';
          const dimension = position === 'left' ? 'width' : 'height';
          const sorted = [...shapes].sort((a, b) => a[position] - b[position]);
          value = sorted.length > 1 ? sorted.slice(1).reduce((sum, shape, i) => sum + shape[position] - sorted[i][position] - sorted[i][dimension], 0) / (sorted.length - 1) / POINTS_PER_CM : 0;
        }
        item.input.value = String(Math.round(value * 100) / 100);
        item.slider.value = String(value);
      }
    }
    reset.addEventListener('click', () => {
      activeOperation = activeOperation.then(capture).catch(error => ctx.notify(error.message, 'error'));
    });
    const unsubscribe = ctx.host.onChanged(() => {
      if (disposed || applying || changePending) return;
      changePending = true;
      activeOperation = activeOperation.then(async () => {
        try {
          if (disposed) return;
          const snapshot = await ctx.host.summary({ slides: false });
          const next = JSON.stringify([snapshot.presentationId, snapshot.slideId, snapshot.selectionIds]);
          if (next !== signature) await capture();
        } catch (error) { if (!disposed) ctx.notify(error.message, 'error'); }
        finally { changePending = false; }
      });
    });
    try { await capture(); } catch (error) { ctx.notify(error.message, 'error'); }
    return () => { disposed = true; renderId++; unsubscribe(); container.replaceChildren(); };
  },
}];
