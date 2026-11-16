// WPS desktop JavaScript API uses synchronous, one-based native collections.
// Reference: https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/jsapi/wpp/Shape/obj
export const TRUE = -1;
export const FALSE = 0;
const tri = value => value ? TRUE : FALSE;
export const items = collection => Array.from({ length: Number(collection.Count) }, (_, index) => collection.Item(index + 1));
export const flatten = shapes => shapes.flatMap(shape => [shape, ...flatten(shape.children || [])]);
export function optional(getter) { try { return getter(); } catch { return undefined; } }
export function finite(value, label = '数值') {
  const result = Number(value);
  if (!Number.isFinite(result)) throw new Error(`${label}必须为有限数值。`);
  return result;
}
export function rgb(value) {
  if (typeof value === 'number') return value;
  if (!/^#[\da-f]{6}$/i.test(value)) throw new Error(`无效的颜色：${value}`);
  return parseInt(value.slice(1, 3), 16) | parseInt(value.slice(3, 5), 16) << 8 | parseInt(value.slice(5, 7), 16) << 16;
}
export function color(value) {
  if (!Number.isFinite(value) || value < 0) return undefined;
  return '#' + [value & 255, value >> 8 & 255, value >> 16 & 255].map(item => item.toString(16).padStart(2, '0')).join('').toUpperCase();
}
function compact(object) { return Object.fromEntries(Object.entries(object).filter(([, value]) => value !== undefined)); }
function boolean(value) { return value === -2 || value == null ? undefined : value === TRUE || value === 1 || value === true; }
const dashCodes = { solid: 1, squareDot: 2, roundDot: 3, dash: 4, dashDot: 5, dashDotDot: 6, longDash: 7, longDashDot: 8, longDashDotDot: 9, sysDash: 10, sysDot: 11, sysDashDot: 12, sysDashDotDot: 13, dot: 3, lgDash: 7, lgDashDot: 8, lgDashDotDot: 9 };
const dashNames = Object.fromEntries(Object.entries({ solid: 1, squareDot: 2, roundDot: 3, dash: 4, dashDot: 5, dashDotDot: 6, longDash: 7, longDashDot: 8, longDashDotDot: 9, sysDash: 10, sysDot: 11, sysDashDot: 12, sysDashDotDot: 13 }).map(([name, code]) => [code, name]));
const arrowCodes = { none: 1, triangle: 2, open: 3, stealth: 4, diamond: 5, oval: 6 };
const arrowNames = Object.fromEntries(Object.entries(arrowCodes).map(([name, code]) => [code, name]));
export function readTags(native) {
  const tags = native.Tags;
  return Object.fromEntries(Array.from({ length: Number(tags.Count) }, (_, index) => [String(tags.Name(index + 1)), String(tags.Value(index + 1))]));
}
function tagName(tags, name) {
  const wanted = String(name).toUpperCase();
  for (let index = 1; index <= Number(tags.Count); index++) {
    const current = String(tags.Name(index));
    if (current === name || current.toUpperCase() === wanted) return current;
  }
  return undefined;
}
export function writeTags(native, tags) {
  const collection = native.Tags;
  for (const [key, value] of Object.entries(tags || {})) {
    const existing = tagName(collection, key);
    if (value === undefined) continue;
    if (value === null || value === '') {
      if (existing) collection.Delete(existing);
      continue;
    }
    if (existing) collection.Delete(existing);
    collection.Add(key, String(value));
  }
}
export function readLine(native) {
  if (!native) return undefined;
  return compact({ visible: boolean(optional(() => native.Visible)), color: color(optional(() => native.ForeColor.RGB)),
    width: optional(() => native.Weight), transparency: optional(() => native.Transparency),
    dashStyle: dashNames[optional(() => native.DashStyle)], style: optional(() => native.Style),
    beginArrow: arrowNames[optional(() => native.BeginArrowheadStyle)],
    endArrow: arrowNames[optional(() => native.EndArrowheadStyle)] });
}
export function readFont(native, native2) {
  if (!native && !native2) return undefined;
  const result = {};
  const properties = { name: 'Name', nameAscii: 'NameAscii', nameFarEast: 'NameFarEast', nameComplexScript: 'NameComplexScript', size: 'Size', baselineOffset: 'BaselineOffset' };
  for (const [key, property] of Object.entries(properties)) {
    const classic = optional(() => native?.[property]);
    const advanced = optional(() => native2?.[property]);
    const value = classic !== undefined && classic !== null && classic !== -2 ? classic : advanced;
    if (value !== undefined && value !== null && value !== -2 && (key !== 'size' || value > 0)) result[key] = value;
  }
  for (const [key, property] of Object.entries({ bold: 'Bold', italic: 'Italic', underline: 'Underline', shadow: 'Shadow', emboss: 'Emboss', superscript: 'Superscript', subscript: 'Subscript' })) {
    const classic = optional(() => native?.[property]);
    const advanced = optional(() => native2?.[property]);
    const value = boolean(classic !== undefined && classic !== -2 ? classic : advanced);
    if (value !== undefined) result[key] = value;
  }
  if (native2) {
    const strike = optional(() => native2.Strike);
    if (strike === 1 || strike === 2) result.strikethrough = true;
    else if (strike === 0) result.strikethrough = false;
    else {
      const doubleStrike = boolean(optional(() => native2.DoubleStrikeThrough));
      if (doubleStrike !== undefined) result.strikethrough = doubleStrike;
    }
    const allCaps = boolean(optional(() => native2.Allcaps ?? native2.Caps));
    if (allCaps !== undefined) result.allCaps = allCaps;
    const smallCaps = boolean(optional(() => native2.Smallcaps));
    if (smallCaps !== undefined) result.smallCaps = smallCaps;
    const spacing = optional(() => native2.Spacing);
    if (Number.isFinite(spacing)) result.spacing = spacing;
    const kerning = optional(() => native2.Kerning);
    if (Number.isFinite(kerning)) result.kerning = kerning;
    const underlineColor = color(optional(() => native2.UnderlineColor.RGB));
    if (underlineColor) result.underlineColor = underlineColor;
  }
  const nativeColor = color(optional(() => native?.Color.RGB));
  if (nativeColor) result.color = nativeColor;
  if (native2) {
    const fillColor = color(optional(() => native2.Fill.ForeColor.RGB));
    if (fillColor) result.color = fillColor;
    const transparency = optional(() => native2.Fill.Transparency);
    if (Number.isFinite(transparency)) result.transparency = transparency;
    const outline = readLine(optional(() => native2.Line));
    if (outline) result.outline = outline;
    const glow = optional(() => native2.Glow);
    if (glow) result.glow = compact({ color: color(optional(() => glow.Color.RGB)), radius: optional(() => glow.Radius), transparency: optional(() => glow.Transparency) });
    const reflection = optional(() => native2.Reflection.Type);
    if (reflection !== undefined) result.reflection = { type: reflection };
    const highlightType = optional(() => native2.Highlight.Type);
    const highlight = optional(() => native2.Highlight.RGB);
    if (highlightType !== undefined && Number(highlightType) !== 0 && highlight !== 0 && color(highlight)) result.highlight = color(highlight);
  }
  return result;
}
export function readFill(native) {
  const fill = native.Fill;
  if (!boolean(fill.Visible)) return { type: 'none', transparency: 1 };
  const type = ({ 1: 'solid', 2: 'pattern', 3: 'gradient', 4: 'texture', 5: 'background', 6: 'picture' })[fill.Type];
  const result = compact({ type, color: color(optional(() => fill.ForeColor.RGB)), backColor: color(optional(() => fill.BackColor.RGB)), transparency: optional(() => fill.Transparency) });
  if (type === 'gradient') {
    const stops = optional(() => fill.GradientStops);
    if (stops) result.stops = items(stops).map(stop => {
      const position = Number(stop.Position);
      return { position: position > 1 ? position / 100 : position, color: color(stop.Color.RGB), transparency: stop.Transparency };
    });
    result.angle = optional(() => fill.GradientAngle);
  }
  if (type === 'pattern') result.pattern = optional(() => fill.Pattern);
  return result;
}
function readRuns(range, range2) {
  const runs = range.Runs();
  return Array.from({ length: Number(runs.Count) }, (_, index) => {
    const run = range.Runs(index + 1, 1);
    const offset = Number(run.Start) - Number(range.Start);
    const second = range2 ? optional(() => range2.Characters(offset + 1, Number(run.Length))) : undefined;
    const hyperlink = optional(() => run.ActionSettings.Item(1).Hyperlink.Address);
    return { text: String(run.Text), font: readFont(run.Font, second?.Font), ...(hyperlink ? { hyperlink } : {}) };
  });
}
function stripTerminator(runs) {
  if (!runs.length) return runs;
  const last = runs.at(-1);
  if (last.text.endsWith('\r')) {
    last.text = last.text.slice(0, -1);
    if (!last.text) runs.pop();
  }
  return runs;
}
function readParagraph(range, range2) {
  const format = range.ParagraphFormat;
  const bullet = format.Bullet;
  const bulletType = Number(optional(() => bullet.Type));
  return compact({ runs: stripTerminator(readRuns(range, range2)), level: Math.max(0, Number(range.IndentLevel) - 1),
    alignment: ({ 1: 'left', 2: 'center', 3: 'right', 4: 'justify', 5: 'distribute' })[format.Alignment],
    spaceBefore: format.SpaceBefore, spaceAfter: format.SpaceAfter, lineSpacing: format.SpaceWithin,
    lineRuleWithin: boolean(format.LineRuleWithin), lineRuleBefore: boolean(format.LineRuleBefore), lineRuleAfter: boolean(format.LineRuleAfter),
    bullet: bulletType === 1 ? compact({ type: 'character', character: String.fromCodePoint(Number(bullet.Character) || 0x2022), color: color(optional(() => bullet.Font.Color.RGB)), font: optional(() => bullet.Font.Name), relativeSize: optional(() => bullet.RelativeSize) })
      : bulletType === 2 ? compact({ type: 'number', style: optional(() => bullet.Style), start: optional(() => bullet.StartValue) }) : null });
}
export function readText(native, rich = true) {
  if (!boolean(native.HasTextFrame)) return {};
  const frame = native.TextFrame;
  const range = frame.TextRange;
  const textFrame = compact({ wordWrap: boolean(frame.WordWrap), autoSize: ({ 0: 'none', 1: 'shapeToFitText', 2: 'textToFitShape' })[frame.AutoSize],
    marginLeft: frame.MarginLeft, marginRight: frame.MarginRight, marginTop: frame.MarginTop, marginBottom: frame.MarginBottom,
    verticalAnchor: ({ 1: 'top', 2: 'top', 3: 'middle', 4: 'bottom', 5: 'bottom' })[optional(() => native.TextFrame2.VerticalAnchor)],
    horizontalAnchor: ({ 1: 'left', 2: 'center', 3: 'right' })[optional(() => native.TextFrame2.HorizontalAnchor)] });
  const text = String(range.Text);
  // A saved PPTX snapshot supplies runs, paragraphs and native effects in bulk.
  // Keep only computed native defaults here; avoid repeating every font RPC for
  // both the run list and paragraph list of every shape during context reads.
  if (!rich) return { text, font: readFont(range.Font), textFrame };
  const second = optional(() => native.TextFrame2.TextRange);
  const paragraphs = range.Paragraphs();
  return { text, font: readFont(range.Font, second?.Font), runs: readRuns(range, second),
    paragraphs: Array.from({ length: Number(paragraphs.Count) }, (_, index) => {
      const paragraph = range.Paragraphs(index + 1, 1);
      const offset = Number(paragraph.Start) - Number(range.Start);
      return readParagraph(paragraph, second ? optional(() => second.Characters(offset + 1, Number(paragraph.Length))) : undefined);
    }), textFrame };
}
export function readCrop(native) {
  const crop = native.PictureFormat.Crop;
  const width = Number(crop.PictureWidth), height = Number(crop.PictureHeight);
  if (!(width > 0 && height > 0)) throw new Error('WPS 图片原始绘制尺寸无效。');
  const frameWidth = Number(crop.ShapeWidth), frameHeight = Number(crop.ShapeHeight);
  const x = Number(crop.PictureOffsetX), y = Number(crop.PictureOffsetY);
  return { left: (width - frameWidth - x * 2) / (width * 2), right: (width - frameWidth + x * 2) / (width * 2),
    top: (height - frameHeight - y * 2) / (height * 2), bottom: (height - frameHeight + y * 2) / (height * 2) };
}
export function readShape(native, slideIndex, shapePath, parentId) {
  const type = Number(native.Type);
  const item = { id: String(native.Id), name: String(native.Name),
    type: type === 6 ? 'group' : [11, 13].includes(type) ? 'picture' : type === 16 ? 'media' : type === 19 ? 'table' : [14, 17].includes(type) ? 'text' : 'shape',
    nativeType: type, left: Number(native.Left), top: Number(native.Top), width: Number(native.Width), height: Number(native.Height), rotation: Number(native.Rotation),
    alternativeText: String(native.AlternativeText || ''), tags: readTags(native), zOrder: Number(native.ZOrderPosition) - 1,
    ref: { slideIndex, shapePath }, isPlaceholder: type === 14, visible: boolean(optional(() => native.Visible)),
    flipH: boolean(optional(() => native.HorizontalFlip)), flipV: boolean(optional(() => native.VerticalFlip)),
    fill: optional(() => readFill(native)) ?? { type: 'none', transparency: 1 }, line: optional(() => readLine(native.Line)), ...readText(native, false),
    ...(parentId ? { parentId } : {}) };
  const shapeType = optional(() => native.AutoShapeType);
  if (shapeType !== undefined) item.shapeType = shapeType;
  if (item.type === 'picture') { const crop = optional(() => readCrop(native)); if (crop) item.crop = crop; }
  item.style = { fill: item.fill, line: item.line, font: item.font, textFrame: item.textFrame };
  if (item.type === 'group') item.children = items(native.GroupItems).map((child, index) => readShape(child, slideIndex, [...shapePath, index + 1], item.id));
  return item;
}
export function shapeIndex(slide, slideIndex) {
  const shapes = items(slide.Shapes).map((native, index) => readShape(native, slideIndex, [index + 1]));
  return shapes;
}
export function findNative(slide, id) {
  const entry = findNativeEntry(slide, id);
  if (!entry) throw new Error(`WPS 当前页面中找不到对象 ${id}，请重新选择。`);
  return entry.shape;
}
export function findNativeEntry(slide, id) {
  function find(collection, parent) {
    for (const shape of items(collection)) {
      if (String(shape.Id) === String(id)) return { shape, parent, collection };
      if (Number(shape.Type) === 6) { const match = find(shape.GroupItems, shape); if (match) return match; }
    }
    return undefined;
  }
  return find(slide.Shapes, undefined);
}
export function setLine(native, value) {
  if (value.visible !== undefined) native.Visible = tri(value.visible);
  if (value.color !== undefined) native.ForeColor.RGB = rgb(value.color);
  if (value.width !== undefined) native.Weight = finite(value.width, '轮廓宽度');
  if (value.transparency !== undefined) native.Transparency = finite(value.transparency);
  if (value.dashStyle !== undefined) {
    const code = typeof value.dashStyle === 'number' ? value.dashStyle : dashCodes[value.dashStyle];
    if (!code) throw new Error(`无效的轮廓线型：${value.dashStyle}`);
    native.DashStyle = code;
  }
  if (value.style !== undefined) native.Style = value.style;
  if (value.beginArrow !== undefined) native.BeginArrowheadStyle = arrowCodes[value.beginArrow] ?? finite(value.beginArrow);
  if (value.endArrow !== undefined) native.EndArrowheadStyle = arrowCodes[value.endArrow] ?? finite(value.endArrow);
}
export async function setFill(native, value, asset) {
  if (value.type === 'none') { native.Visible = FALSE; return; }
  native.Visible = TRUE;
  if (value.type === 'picture') { if (typeof asset !== 'function') throw new Error('图片填充需要可用的本地文件服务。'); native.UserPicture(await asset(value)); }
  else if (value.type === 'gradient') {
    const stops = value.stops;
    if (!stops?.length || stops.length < 2) throw new Error('渐变填充必须包含至少两个颜色位置。');
    native.ForeColor.RGB = rgb(stops[0].color);
    native.BackColor.RGB = rgb(stops.at(-1).color);
    native.TwoColorGradient(1, 1);
    if (value.angle !== undefined) native.GradientAngle = finite(value.angle);
    const target = native.GradientStops;
    for (let i = Number(target.Count); i > 2; i--) target.Delete(i);
    const first = target.Item(1), last = target.Item(2);
    first.Color.RGB = rgb(stops[0].color); first.Position = finite(stops[0].position); first.Transparency = stops[0].transparency ?? 0;
    last.Color.RGB = rgb(stops.at(-1).color); last.Position = finite(stops.at(-1).position); last.Transparency = stops.at(-1).transparency ?? 0;
    for (const stop of stops.slice(1, -1)) target.Insert2(rgb(stop.color), finite(stop.position), stop.transparency ?? 0);
  } else if (value.type === 'pattern') native.Patterned(finite(value.pattern));
  else native.Solid();
  if (value.color !== undefined) native.ForeColor.RGB = rgb(value.color);
  if (value.backColor !== undefined) native.BackColor.RGB = rgb(value.backColor);
  if (value.transparency !== undefined) native.Transparency = finite(value.transparency);
}
function clearHighlight(second) {
  const before = optional(() => second.Highlight.RGB);
  try {
    second.Highlight.ObjectThemeColor = 0;
    const after = optional(() => second.Highlight.RGB);
    if (after === -1 || after === undefined) return;
    if (before !== undefined && before >= 0) second.Highlight.RGB = before;
    throw new Error('WPS 演示接口无法清除文字高亮。请先在 WPS 中手动清除目标文字高亮。');
  } catch (error) {
    if (before !== undefined && before >= 0) { try { second.Highlight.RGB = before; } catch { /* restore is best effort; the original error is reported */ } }
    throw error;
  }
}
export function setFont(native, second, value) {
  const properties = { name: 'Name', nameAscii: 'NameAscii', nameFarEast: 'NameFarEast', nameComplexScript: 'NameComplexScript', size: 'Size', baselineOffset: 'BaselineOffset' };
  for (const [key, property] of Object.entries(properties)) if (value[key] !== undefined) native[property] = value[key];
  for (const [key, property] of Object.entries({ bold: 'Bold', italic: 'Italic', underline: 'Underline', shadow: 'Shadow', emboss: 'Emboss', superscript: 'Superscript', subscript: 'Subscript' })) {
    if (value[key] !== undefined) native[property] = tri(value[key]);
  }
  if (value.color !== undefined) native.Color.RGB = rgb(value.color);
  if (value.strikethrough !== undefined) {
    if (!second) throw new Error('WPS 演示缺少文字删除线接口。');
    second.Strike = value.strikethrough ? 1 : 0;
  }
  if (value.allCaps !== undefined && second) second.Allcaps = tri(value.allCaps);
  if (value.smallCaps !== undefined && second) second.Smallcaps = tri(value.smallCaps);
  if (value.spacing !== undefined && second) second.Spacing = finite(value.spacing);
  if (value.kerning !== undefined && second) second.Kerning = finite(value.kerning);
  if (value.underlineColor !== undefined && second) second.UnderlineColor.RGB = rgb(value.underlineColor);
  if (value.transparency !== undefined) {
    if (!second) throw new Error('WPS 演示缺少文字透明度接口。');
    second.Fill.Transparency = finite(value.transparency);
  }
  if (value.outline !== undefined) {
    if (!second) throw new Error('WPS 演示缺少文字描边接口。');
    setLine(second.Line, value.outline);
  }
  if (value.glow !== undefined) {
    if (!second) throw new Error('WPS 演示缺少文字发光接口。');
    if (value.glow.color !== undefined) second.Glow.Color.RGB = rgb(value.glow.color);
    if (value.glow.radius !== undefined) second.Glow.Radius = finite(value.glow.radius);
    if (value.glow.transparency !== undefined) second.Glow.Transparency = finite(value.glow.transparency);
  }
  if (value.reflection !== undefined) {
    if (!second) throw new Error('WPS 演示缺少文字倒影接口。');
    second.Reflection.Type = typeof value.reflection === 'number' ? value.reflection : value.reflection.type;
  }
  if (value.highlight !== undefined) {
    if (!second) throw new Error('WPS 演示缺少文字高亮接口。');
    if (value.highlight === null) clearHighlight(second);
    else second.Highlight.RGB = rgb(value.highlight);
  }
}
export function setCrop(native, value, geometry = {}) {
  const left = finite(value.left ?? 0), right = finite(value.right ?? 0), top = finite(value.top ?? 0), bottom = finite(value.bottom ?? 0);
  if (Math.min(left, right, top, bottom) < 0 || Math.max(left, right, top, bottom) > 1) throw new Error('裁剪比例必须在 0 到 1 之间。');
  if (1 - left - right <= 0 || 1 - top - bottom <= 0) throw new Error('裁剪必须保留可见区域。');
  const crop = native.PictureFormat.Crop;
  const width = geometry.width ?? native.Width, height = geometry.height ?? native.Height;
  const pictureWidth = width / (1 - left - right), pictureHeight = height / (1 - top - bottom);
  crop.PictureWidth = pictureWidth;
  crop.PictureHeight = pictureHeight;
  crop.ShapeWidth = width;
  crop.ShapeHeight = height;
  crop.PictureOffsetX = (right - left) * pictureWidth / 2;
  crop.PictureOffsetY = (bottom - top) * pictureHeight / 2;
}
export function setTextFrame(native, value) {
  const frame = native.TextFrame;
  for (const [key, property] of Object.entries({ marginLeft: 'MarginLeft', marginRight: 'MarginRight', marginTop: 'MarginTop', marginBottom: 'MarginBottom' })) {
    if (value[key] !== undefined) frame[property] = finite(value[key]);
  }
  if (value.wordWrap !== undefined) frame.WordWrap = tri(value.wordWrap);
  if (value.autoSize !== undefined) {
    const code = { none: 0, shapeToFitText: 1, textToFitShape: 2 }[value.autoSize];
    if (code === undefined) throw new Error('无效的文本自动尺寸选项。');
    frame.AutoSize = code;
  }
  if (value.verticalAnchor !== undefined) {
    const code = { top: 1, middle: 3, center: 3, bottom: 4, t: 1, ctr: 3, b: 4 }[value.verticalAnchor];
    if (code === undefined) throw new Error('无效的文本垂直对齐选项。');
    native.TextFrame2.VerticalAnchor = code;
  }
  if (value.horizontalAnchor !== undefined) {
    const code = { left: 1, none: 1, l: 1, center: 2, ctr: 2 }[value.horizontalAnchor];
    if (code === undefined) throw new Error('WPS 演示不支持该文本水平对齐选项。');
    native.TextFrame2.HorizontalAnchor = code;
  }
}
const numberedStyles = { arabicParenBoth: 12, arabicParenR: 2, arabicPeriod: 3, arabicPlain: 13,
  alphaLcParenBoth: 8, alphaLcParenR: 9, alphaLcPeriod: 0, alphaUcParenBoth: 10, alphaUcParenR: 11, alphaUcPeriod: 1,
  romanLcParenBoth: 4, romanLcParenR: 5, romanLcPeriod: 6, romanUcParenBoth: 14, romanUcParenR: 15, romanUcPeriod: 7,
  circleNumDbPlain: 18, circleNumWDBlackPlain: 20, circleNumWDWhitePlain: 19,
  arabicDbPeriod: 29, arabicDbPlain: 28, simplChinPeriod: 17, simplChinPlain: 16, tradChinPeriod: 22, tradChinPlain: 21 };
function setParagraph(range, value) {
  const format = range.ParagraphFormat;
  if (value.alignment !== undefined) format.Alignment = { left: 1, l: 1, center: 2, ctr: 2, right: 3, r: 3, justify: 4, just: 4, distribute: 5, dist: 5 }[value.alignment];
  if (value.level !== undefined) range.IndentLevel = Math.max(1, Math.min(5, finite(value.level) + 1));
  for (const [key, property] of Object.entries({ spaceBefore: 'SpaceBefore', spaceAfter: 'SpaceAfter', lineSpacing: 'SpaceWithin' })) if (value[key] !== undefined) format[property] = finite(value[key]);
  for (const [key, property] of Object.entries({ lineRuleWithin: 'LineRuleWithin', lineRuleBefore: 'LineRuleBefore', lineRuleAfter: 'LineRuleAfter' })) if (value[key] !== undefined) format[property] = tri(value[key]);
  if (value.bullet !== undefined) {
    const bullet = format.Bullet;
    const specification = value.bullet === true ? { type: 'character', character: '•' } : value.bullet;
    if (!specification || specification.type === 'none') { bullet.Type = 0; }
    else if (specification.type === 'character' || specification.type === 'bullet') {
      bullet.Type = 1;
      const character = String(specification.character || '•');
      const codePoint = character.codePointAt(0);
      if (!Number.isInteger(codePoint)) throw new Error('无效的项目符号字符。');
      bullet.Character = codePoint;
      if (specification.font) { bullet.UseTextFont = FALSE; bullet.Font.Name = specification.font; }
      if (specification.color) { bullet.UseTextColor = FALSE; bullet.Font.Color.RGB = rgb(specification.color); }
      if (specification.relativeSize !== undefined) bullet.RelativeSize = finite(specification.relativeSize);
    } else if (specification.type === 'number' || specification.type === 'numbered') {
      bullet.Type = 2;
      const start = specification.start ?? specification.startAt;
      if (start !== undefined) bullet.StartValue = finite(start);
      if (typeof specification.style === 'number') bullet.Style = specification.style;
      else if (specification.style) {
        if (numberedStyles[specification.style] === undefined) throw new Error(`无效的编号样式：${specification.style}`);
        bullet.Style = numberedStyles[specification.style];
      }
      if (specification.font) { bullet.UseTextFont = FALSE; bullet.Font.Name = specification.font; }
      if (specification.color) { bullet.UseTextColor = FALSE; bullet.Font.Color.RGB = rgb(specification.color); }
    } else throw new Error('无效的项目符号类型。');
  }
}
function canonical(value) { return String(value).replace(/\r\n|[\r\n]/g, '\r'); }
export function sameText(left, right) { return canonical(left).replace(/\r$/, '') === canonical(right).replace(/\r$/, ''); }
export function setText(native, properties) {
  let range = native.TextFrame.TextRange;
  let second = optional(() => native.TextFrame2.TextRange);
  let runs = properties.runs;
  if (properties.paragraphs?.some(paragraph => paragraph.runs || paragraph.text !== undefined)) runs = properties.paragraphs.flatMap((paragraph, index) => {
    const paragraphRuns = paragraph.runs?.length ? paragraph.runs : paragraph.text !== undefined ? [{ text: String(paragraph.text), font: paragraph.font }] : [];
    return index ? [{ text: '\r' }, ...paragraphRuns] : paragraphRuns;
  });
  let text = properties.text !== undefined ? String(properties.text) : runs ? runs.map(run => run.text || '').join('') : undefined;
  if (text !== undefined && !sameText(text, range.Text)) {
    range.Text = canonical(text);
    range = native.TextFrame.TextRange;
    second = optional(() => native.TextFrame2.TextRange);
  }
  if (properties.textRange?.text !== undefined) {
    const part = properties.textRange;
    const start = finite(part.start), length = finite(part.length);
    if (!Number.isInteger(start) || !Number.isInteger(length) || start < 0 || length < 0 || start + length > Number(range.Length)) throw new Error('文字选择范围已改变，请重新选择。');
    if (length) {
      range.Characters(start + 1, length).Text = String(part.text);
      range = native.TextFrame.TextRange;
      second = optional(() => native.TextFrame2.TextRange);
    }
  }
  if (properties.font) setFont(range.Font, second?.Font, properties.font);
  if (runs) {
    let offset = 0;
    for (const run of runs) {
      const length = canonical(run.text || '').length;
      if (length) {
        const target = range.Characters(offset + 1, length), target2 = second ? second.Characters(offset + 1, length) : undefined;
        if (run.font) setFont(target.Font, target2?.Font, run.font);
        if (run.hyperlink !== undefined) {
          const settings = target.ActionSettings.Item(1);
          settings.Hyperlink.Address = String(run.hyperlink);
          settings.Action = run.hyperlink ? 7 : 0;
        }
      }
      offset += length;
    }
  }
  if (properties.paragraphs) properties.paragraphs.forEach((paragraph, index) => setParagraph(range.Paragraphs(index + 1, 1), paragraph));
  if (properties.textRange) {
    const part = properties.textRange;
    const start = finite(part.start), length = finite(part.length);
    if (!Number.isInteger(start) || !Number.isInteger(length) || start < 0 || length < 0 || start + length > Number(range.Length)) throw new Error('文字选择范围已改变，请重新选择。');
    if (length) {
      if (part.font) setFont(range.Characters(start + 1, length).Font, second ? second.Characters(start + 1, length).Font : undefined, part.font);
    }
  }
}
export async function properties(native, value, asset) {
  if (value.width !== undefined || value.height !== undefined) {
    const locked = native.LockAspectRatio;
    native.LockAspectRatio = FALSE;
    try {
      if (value.width !== undefined) native.Width = finite(value.width, '宽度');
      if (value.height !== undefined) native.Height = finite(value.height, '高度');
    } finally { native.LockAspectRatio = locked === TRUE ? TRUE : FALSE; }
  }
  if (value.crop !== undefined) setCrop(native, value.crop, value);
  if (value.rotation !== undefined) native.Rotation = finite(value.rotation, '旋转角度');
  if (value.left !== undefined) native.Left = finite(value.left, '水平位置');
  if (value.top !== undefined) native.Top = finite(value.top, '垂直位置');
  if (value.name !== undefined) native.Name = String(value.name);
  if (value.alternativeText !== undefined) native.AlternativeText = String(value.alternativeText);
  if (value.tags !== undefined) writeTags(native, value.tags);
  if (value.fill !== undefined) await setFill(native.Fill, value.fill, asset);
  if (value.line !== undefined) setLine(native.Line, value.line);
  if (value.textFrame !== undefined) setTextFrame(native, value.textFrame);
  if (['text', 'runs', 'paragraphs', 'font', 'textRange'].some(key => value[key] !== undefined)) setText(native, value);
}
export function zOrder(shape, position, index) {
  const codes = { front: 0, back: 1, forward: 2, backward: 3 };
  if (position !== 'index') {
    if (codes[position] === undefined) throw new Error('无效的堆叠顺序。');
    shape.ZOrder(codes[position]);
    return;
  }
  const desired = finite(index) + 1;
  if (!Number.isInteger(desired) || desired < 1) throw new Error('堆叠索引必须是非负整数。');
  const current = Number(shape.ZOrderPosition);
  for (let remaining = Math.abs(current - desired); remaining > 0; remaining--) shape.ZOrder(current > desired ? 3 : 2);
}
export const shapeCodes = { rectangle: 1, rect: 1, roundRect: 5, roundedRectangle: 5, ellipse: 9, oval: 9, triangle: 7, rightTriangle: 8, diamond: 4, parallelogram: 2, trapezoid: 3, pentagon: 51, hexagon: 10, octagon: 6, star: 12, star5: 12, rightArrow: 33, leftArrow: 34, upArrow: 35, downArrow: 36, leftRightArrow: 37, chevron: 52, arc: 25, bracePair: 26, bracketPair: 27, cloud: 179, heart: 21, plus: 11 };
