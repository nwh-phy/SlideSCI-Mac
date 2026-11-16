import { DOMParser, XMLSerializer } from '@xmldom/xmldom';

const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const XML_NS = 'http://www.w3.org/XML/1998/namespace';
const EMU = 12700;
const serializer = new XMLSerializer();

const RPR_ORDER = ['ln', 'noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill',
  'effectLst', 'effectDag', 'highlight', 'uLnTx', 'uLn', 'uFillTx', 'uFill',
  'latin', 'ea', 'cs', 'sym', 'hlinkClick', 'hlinkMouseOver', 'rtl', 'extLst'];
const EFFECT_ORDER = ['blur', 'fillOverlay', 'glow', 'innerShdw', 'outerShdw', 'prstShdw', 'reflection', 'softEdge', 'extLst'];
const LN_ORDER = ['noFill', 'solidFill', 'gradFill', 'pattFill', 'prstDash', 'custDash',
  'round', 'bevel', 'miter', 'headEnd', 'tailEnd', 'extLst'];
const PPR_ORDER = ['lnSpc', 'spcBef', 'spcAft', 'buClrTx', 'buClr', 'buSzTx', 'buSzPct', 'buSzPts',
  'buFontTx', 'buFont', 'buNone', 'buAutoNum', 'buChar', 'buBlip', 'tabLst', 'defRPr', 'extLst'];
const FILL_NAMES = ['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill'];
const COLOR_NAMES = [...FILL_NAMES, 'highlight', 'ln'];
const EFFECT_NAMES = ['effectLst', 'effectDag'];
const NAME_NAMES = ['latin', 'ea', 'cs', 'sym'];
const LINK_NAMES = ['hlinkClick', 'hlinkMouseOver', 'hlinkHover'];
const COLOR_CHILD_NAMES = ['srgbClr', 'schemeClr', 'sysClr', 'prstClr', 'scrgbClr', 'hslClr'];
const EXTERNAL_ELEMENTS = ['blip', 'videoFile', 'audioFile', 'media', 'oleObj', 'svgBlip', 'wavAudioFile', 'chart'];
const TEXT_NODES = new Set(['r', 'fld', 'br', 'm']);
const SHAPE_NODES = new Set(['sp', 'grpSp', 'graphicFrame', 'cxnSp', 'pic']);

const FONT_SCOPES = new Set(['all', 'name', 'size', 'color', 'effect']);
const SCOPE_FIELDS = {
  all: ['name', 'nameAscii', 'nameFarEast', 'nameComplexScript', 'size', 'color', 'transparency',
    'bold', 'italic', 'underline', 'underlineStyle', 'strikethrough', 'subscript', 'superscript',
    'baselineOffset', 'highlight', 'outline', 'line', 'glow', 'shadow', 'reflection', 'emboss'],
  merge: ['name', 'nameAscii', 'nameFarEast', 'nameComplexScript', 'size', 'color', 'transparency',
    'bold', 'italic', 'underline', 'underlineStyle', 'strikethrough', 'subscript', 'superscript',
    'baselineOffset', 'highlight', 'outline', 'line', 'glow', 'shadow', 'reflection', 'emboss'],
  name: ['name', 'nameAscii', 'nameFarEast', 'nameComplexScript'],
  size: ['size'],
  color: ['color', 'transparency', 'highlight', 'outline', 'line'],
  effect: ['shadow', 'emboss', 'baselineOffset', 'subscript', 'superscript', 'glow', 'reflection'],
};

const UNDERLINE_TYPES = {
  none: 'none', single: 'sng', sng: 'sng', singleline: 'sng', 'single line': 'sng',
  double: 'dbl', dbl: 'dbl', doubleline: 'dbl', 'double line': 'dbl',
  heavy: 'heavy', heavyline: 'heavy', 'heavy line': 'heavy',
  dotted: 'dotted', dot: 'dotted', dotline: 'dotted', 'dot line': 'dotted', dottedheavy: 'dottedHeavy',
  dash: 'dash', dashline: 'dash', 'dash line': 'dash', dashheavy: 'dashHeavy',
  longdash: 'dashLong', dashlong: 'dashLong', longdashline: 'dashLong', dashlongline: 'dashLong',
  dashlongheavy: 'dashLongHeavy',
  dotdash: 'dotDash', dotdashline: 'dotDash', 'dot dash line': 'dotDash', dotdashheavy: 'dotDashHeavy',
  dotdotdash: 'dotDotDash', dotdotdashline: 'dotDotDash', 'dot dot dash line': 'dotDotDash',
  dotdotdashheavy: 'dotDotDashHeavy',
  wavy: 'wavy', wavyline: 'wavy', 'wavy line': 'wavy', wavyheavy: 'wavyHeavy',
  wavydouble: 'wavyDouble', wavydoubleline: 'wavyDouble', 'wavy double line': 'wavyDouble',
  words: 'words',
};
const UNDERLINE_NUMBERS = {
  0: 'none', 1: 'sng', 2: 'words', 3: 'dbl', 4: 'dotted', 6: 'dash', 7: 'dashLong',
  9: 'dotDash', 10: 'dotDotDash', 11: 'wavy', 17: 'wavyDouble', 20: 'dottedHeavy',
  21: 'dashHeavy', 23: 'dashLongHeavy', 24: 'dotDashHeavy', 25: 'dotDotDashHeavy', 26: 'wavyHeavy',
};
const DASH_TYPES = new Set(['solid', 'dot', 'dash', 'lgDash', 'dashDot', 'lgDashDot', 'lgDashDotDot',
  'sysDash', 'sysDot', 'sysDashDot', 'sysDashDotDot']);
const DASH_ALIASES = {
  solid: 'solid', rounddot: 'sysDot', dot: 'sysDot', squaredot: 'sysDot',
  dash: 'dash', dashdot: 'dashDot', dashdotdot: 'lgDashDotDot', longdash: 'lgDash',
  longdashdot: 'lgDashDot', longdashdotdot: 'lgDashDotDot', sysdash: 'sysDash', sysdot: 'sysDot',
  sysdashdot: 'sysDashDot', sysdashdotdot: 'sysDashDotDot',
};
const ARROW_TYPES = new Set(['none', 'triangle', 'stealth', 'diamond', 'oval', 'arrow']);

const REFLECTION_TYPE_NUMBERS = {
  none: 0, type1: 1, type2: 2, type3: 3, type4: 4, type5: 5, type6: 6, type7: 7, type8: 8, type9: 9,
  tight: 1, half: 2, full: 3, semitransparent: 4, glow: 5, blur: 6, opposite: 7, tilt: 8, distant: 9,
};

function fail(message) { throw new Error(message); }
function element(node) { return node && node.nodeType === 1; }
function children(node, name) {
  const result = [];
  for (let item = node?.firstChild; item; item = item.nextSibling) {
    if (item.nodeType === 1 && (!name || item.localName === name)) result.push(item);
  }
  return result;
}
function child(node, name) { return children(node, name)[0] || null; }
function descendants(node, name) {
  const result = [];
  for (const item of children(node)) {
    if (!name || item.localName === name) result.push(item);
    result.push(...descendants(item, name));
  }
  return result;
}
function allElements(node) { return element(node) ? [node, ...descendants(node)] : descendants(node); }
function attr(node, key, fallback = '') { return node?.hasAttribute(key) ? node.getAttribute(key) : fallback; }
function numberValue(node, key, fallback = 0) {
  const value = Number(attr(node, key, fallback));
  return Number.isFinite(value) ? value : fallback;
}
function create(doc, name) { return doc.createElementNS(A, 'a:' + name); }
function insertOrdered(parent, node, order) {
  const rank = order.indexOf(node.localName);
  if (rank < 0) { parent.appendChild(node); return node; }
  for (let item = parent.firstChild; item; item = item.nextSibling) {
    if (!element(item)) continue;
    const other = order.indexOf(item.localName);
    if (other > rank) { parent.insertBefore(node, item); return node; }
  }
  parent.appendChild(node);
  return node;
}
function ensureOrdered(doc, parent, name, order) {
  const existing = child(parent, name);
  if (existing) return existing;
  return insertOrdered(parent, create(doc, name), order);
}
function removeNamed(parent, names) {
  for (const item of children(parent)) if (names.includes(item.localName)) parent.removeChild(item);
}
function removeEffects(root, names) {
  const effects = child(root, 'effectLst');
  if (effects) removeNamed(effects, names);
}
function removeLinks(node) {
  for (const item of allElements(node)) {
    if (LINK_NAMES.includes(item.localName) && item !== node) item.parentNode.removeChild(item);
  }
}
function has(node, key) { return node?.[key] !== undefined; }
function finite(value, label) {
  const result = Number(value);
  if (!Number.isFinite(result)) fail(label + '无效。');
  return result;
}
function clamp(value, min, max) { return Math.min(max, Math.max(min, value)); }
function textContent(node) { return node?.textContent || ''; }

function setPreservedText(t, value) {
  t.textContent = value;
  if (/^\s|\s$/.test(value)) t.setAttributeNS(XML_NS, 'xml:space', 'preserve');
  else t.removeAttributeNS(XML_NS, 'space');
}

function parseColor(value) {
  if (typeof value !== 'string' || !value.trim()) fail('颜色必须使用 #RRGGBB。');
  let text = value.trim();
  if (/^#[0-9a-f]{3}$/i.test(text)) {
    text = '#' + [...text.slice(1)].map(item => item + item).join('');
  }
  if (/^#[0-9a-f]{6}$/i.test(text)) return { hex: text.slice(1).toUpperCase(), alpha: null };
  if (/^#[0-9a-f]{8}$/i.test(text)) {
    return { hex: text.slice(1, 7).toUpperCase(), alpha: parseInt(text.slice(7), 16) / 255 };
  }
  fail('颜色必须使用 #RRGGBB 或 #RRGGBBAA。');
}

function alphaElement(doc, alpha) {
  const alphaNode = create(doc, 'alpha');
  alphaNode.setAttribute('val', String(Math.round(clamp(alpha, 0, 1) * 100000)));
  return alphaNode;
}
function colorNode(doc, parsed, alpha) {
  const node = create(doc, 'srgbClr');
  node.setAttribute('val', parsed.hex);
  if (alpha !== null && alpha < 1) node.appendChild(alphaElement(doc, alpha));
  return node;
}
function existingAlpha(container) {
  const alpha = descendants(container, 'alpha')[0];
  return alpha ? numberValue(alpha, 'val', 100000) / 100000 : null;
}
function setColorChild(doc, container, color, transparency, existing = container) {
  const parsed = parseColor(color);
  let alpha = null;
  if (transparency !== undefined) alpha = 1 - finite(transparency, '透明度');
  else if (parsed.alpha !== null) alpha = parsed.alpha;
  else if (existing) alpha = existingAlpha(existing);
  removeNamed(container, COLOR_CHILD_NAMES);
  container.appendChild(colorNode(doc, parsed, alpha));
}
function setAlphaOnly(container, transparency) {
  const color = children(container).find(item => COLOR_CHILD_NAMES.includes(item.localName));
  if (!color) return false;
  removeNamed(color, ['alpha']);
  const alpha = 1 - finite(transparency, '透明度');
  if (alpha < 1) color.appendChild(alphaElement(color.ownerDocument, alpha));
  return true;
}
function setSolidFill(doc, parent, color, transparency) {
  const existing = children(parent).find(item => FILL_NAMES.includes(item.localName));
  const parsed = parseColor(color);
  let alpha = null;
  if (transparency !== undefined) alpha = 1 - finite(transparency, '透明度');
  else if (parsed.alpha !== null) alpha = parsed.alpha;
  else if (existing) alpha = existingAlpha(existing);
  removeNamed(parent, FILL_NAMES);
  const solid = create(doc, 'solidFill');
  solid.appendChild(colorNode(doc, parsed, alpha));
  insertOrdered(parent, solid, RPR_ORDER);
  return solid;
}

function stripLinks(root) {
  for (const node of allElements(root)) {
    if (node !== root && LINK_NAMES.includes(node.localName)) node.parentNode.removeChild(node);
  }
}
function rejectExternalResources(root, label) {
  for (const node of allElements(root)) {
    if (node.namespaceURI === R) fail(label + '包含外部关系资源。');
    if (EXTERNAL_ELEMENTS.includes(node.localName)) fail(label + '包含外部资源。');
    if (LINK_NAMES.includes(node.localName)) fail(label + '包含不允许复制的超链接。');
    if (!node.attributes) continue;
    for (let index = 0; index < node.attributes.length; index++) {
      const attribute = node.attributes.item(index);
      if (attribute.namespaceURI === R || attribute.prefix === 'r' || attribute.name.startsWith('r:')) {
        fail(label + '包含外部关系资源。');
      }
    }
  }
}
function parseRawElement(doc, raw, expected, label, { links = false } = {}) {
  if (raw === undefined || raw === null) return null;
  let text;
  if (typeof raw === 'string') text = raw;
  else if (element(raw)) text = serializer.serializeToString(raw);
  else fail(label + '必须是 XML 字符串。');
  if (/<!DOCTYPE|<!ENTITY/i.test(text)) fail(label + '不允许文档类型或实体声明。');
  if (text.length > 2_000_000) fail(label + '过大。');
  const parsed = new DOMParser({ onError(level) {
    if (level !== 'warning') fail(label + ' XML 格式无效。');
  } }).parseFromString(text, 'application/xml');
  if (!parsed.documentElement || parsed.getElementsByTagName('parsererror').length) fail(label + ' XML 格式无效。');
  const root = parsed.documentElement;
  if (root.namespaceURI !== A) fail(label + '必须使用 DrawingML 命名空间。');
  if (expected.length && !expected.includes(root.localName)) fail(label + '根元素必须是 a:' + expected.join('/a:') + '。');
  if (links) stripLinks(root);
  rejectExternalResources(root, label);
  return doc.importNode(root, true);
}

function backgroundRunProperties(run) {
  return child(run, 'rPr');
}
function targetLinks(properties) {
  return children(properties)
    .filter(item => item.localName === 'hlinkClick' || item.localName === 'hlinkMouseOver')
    .map(item => item.cloneNode(true));
}
function attachLinks(root, links) {
  for (const link of links) insertOrdered(root, link, RPR_ORDER);
}

function renameProperties(doc, root, name) {
  if (root.localName === name) return root;
  const replacement = create(doc, name);
  if (root.attributes) for (let index = 0; index < root.attributes.length; index++) {
    const attribute = root.attributes.item(index);
    if (attribute.namespaceURI) replacement.setAttributeNS(attribute.namespaceURI, attribute.name, attribute.value);
    else replacement.setAttribute(attribute.name, attribute.value);
  }
  while (root.firstChild) replacement.appendChild(root.firstChild);
  return replacement;
}

function clearColorProperties(root) {
  removeNamed(root, COLOR_NAMES);
}
function clearEffectProperties(root) {
  removeNamed(root, EFFECT_NAMES);
  root.removeAttribute('baseline');
}
function clearNameProperties(root) {
  removeNamed(root, NAME_NAMES);
}
function copyRawChildren(raw, root, names) {
  for (const item of children(raw)) {
    if (!names.includes(item.localName)) continue;
    insertOrdered(root, item, RPR_ORDER);
  }
}

function underlineValue(font, existing) {
  if (font.underline === false) return 'none';
  if (font.underlineStyle !== undefined) {
    if (font.underlineStyle === false || font.underlineStyle === null) return 'none';
    const key = String(font.underlineStyle).replace(/[\s_-]/g, '').toLowerCase();
    if (UNDERLINE_NUMBERS[font.underlineStyle] !== undefined) return UNDERLINE_NUMBERS[font.underlineStyle];
    if (UNDERLINE_TYPES[key]) return UNDERLINE_TYPES[key];
    fail('下划线样式无效：' + font.underlineStyle);
  }
  if (font.underline === true) {
    const current = attr(existing, 'u', 'none');
    return current && current !== 'none' ? current : 'sng';
  }
  return null;
}
function setBaseline(root, font) {
  if (font.baselineOffset !== undefined) {
    const offset = finite(font.baselineOffset, '基线偏移');
    if (Math.abs(offset) < 1e-9) root.removeAttribute('baseline');
    else root.setAttribute('baseline', String(Math.round(offset * 100000)));
    return;
  }
  if (font.superscript === true) {
    if (numberValue(root, 'baseline') <= 0) root.setAttribute('baseline', '30000');
    return;
  }
  if (font.subscript === true) {
    if (numberValue(root, 'baseline') >= 0) root.setAttribute('baseline', '-25000');
    return;
  }
  if (font.superscript === false || font.subscript === false) {
    const current = numberValue(root, 'baseline');
    if ((font.superscript === false && font.subscript === false)
      || (font.superscript === false && current > 0)
      || (font.subscript === false && current < 0)) root.removeAttribute('baseline');
  }
}
function setStrike(root, font) {
  if (font.strikethrough === undefined) return;
  const value = typeof font.strikethrough === 'string' ? font.strikethrough.toLowerCase() : font.strikethrough;
  if (value === 'double' || value === 'dbl' || value === 'dblstrike') root.setAttribute('strike', 'dblStrike');
  else if (value === 'none' || value === 'nostrike' || value === false) root.setAttribute('strike', 'noStrike');
  else if (value === 'single' || value === 'sngstrike' || value === true) root.setAttribute('strike', 'sngStrike');
  else fail('删除线样式无效：' + font.strikethrough);
}
function setBooleanAttribute(root, name, value) {
  if (value === undefined) return;
  root.setAttribute(name, value ? '1' : '0');
}

function setNames(doc, root, font) {
  if (font.name !== undefined || font.nameAscii !== undefined) {
    const value = font.nameAscii !== undefined && font.nameAscii !== '' ? font.nameAscii : font.name;
    if (value === false || value === null || value === '') removeNamed(root, ['latin']);
    else ensureOrdered(doc, root, 'latin', RPR_ORDER).setAttribute('typeface', String(value));
  }
  if (font.nameFarEast !== undefined) {
    if (font.nameFarEast === false || font.nameFarEast === null || font.nameFarEast === '') removeNamed(root, ['ea']);
    else ensureOrdered(doc, root, 'ea', RPR_ORDER).setAttribute('typeface', String(font.nameFarEast));
  }
  if (font.nameComplexScript !== undefined) {
    if (font.nameComplexScript === false || font.nameComplexScript === null || font.nameComplexScript === '') removeNamed(root, ['cs']);
    else ensureOrdered(doc, root, 'cs', RPR_ORDER).setAttribute('typeface', String(font.nameComplexScript));
  }
}

function setFillColor(doc, root, font) {
  if (has(font, 'color')) {
    if (font.color === false || font.color === null) removeNamed(root, FILL_NAMES);
    else setSolidFill(doc, root, font.color, font.transparency);
    return;
  }
  if (font.transparency !== undefined) {
    const existing = children(root).find(item => FILL_NAMES.includes(item.localName));
    if (existing) {
      const color = children(existing).find(item => COLOR_CHILD_NAMES.includes(item.localName));
      if (color) setAlphaOnly(color, font.transparency);
    }
  }
}

function applyHighlight(doc, root, highlight) {
  if (highlight === undefined) return;
  if (!highlight) { removeNamed(root, ['highlight']); return; }
  const node = ensureOrdered(doc, root, 'highlight', RPR_ORDER);
  if (typeof highlight === 'string') setColorChild(doc, node, highlight, undefined, node);
  else if (highlight && typeof highlight === 'object') {
    const color = highlight.color || highlight.value;
    if (!color) fail('高亮颜色缺失。');
    setColorChild(doc, node, color, highlight.transparency, node);
  } else fail('高亮颜色无效。');
}

function setLineColor(doc, line, color, transparency, existing) {
  const parsed = parseColor(color);
  let alpha = null;
  if (transparency !== undefined) alpha = 1 - finite(transparency, '透明度');
  else if (parsed.alpha !== null) alpha = parsed.alpha;
  else if (existing) alpha = existingAlpha(existing);
  removeNamed(line, ['noFill', 'solidFill', 'gradFill', 'pattFill']);
  const solid = create(doc, 'solidFill');
  solid.appendChild(colorNode(doc, parsed, alpha));
  insertOrdered(line, solid, LN_ORDER);
}

function applyOutline(doc, root, outline) {
  if (outline === undefined) return;
  if (outline === false || outline === null) {
    const line = ensureOrdered(doc, root, 'ln', RPR_ORDER);
    removeNamed(line, ['solidFill', 'gradFill', 'pattFill']);
    if (!child(line, 'noFill')) insertOrdered(line, create(doc, 'noFill'), LN_ORDER);
    return;
  }
  if (typeof outline !== 'object') fail('文字轮廓必须是对象。');
  let line;
  if (outline.raw !== undefined) {
    const imported = parseRawElement(doc, outline.raw, ['ln'], '文字轮廓', { links: false });
    const existing = child(root, 'ln');
    if (existing) root.replaceChild(imported, existing);
    else insertOrdered(root, imported, RPR_ORDER);
    line = imported;
  } else line = ensureOrdered(doc, root, 'ln', RPR_ORDER);
  if (outline.visible === false) {
    removeNamed(line, ['solidFill', 'gradFill', 'pattFill']);
    if (!child(line, 'noFill')) insertOrdered(line, create(doc, 'noFill'), LN_ORDER);
    return;
  }
  if (outline.color !== undefined || outline.visible === true) {
    const existing = child(line, 'solidFill');
    if (outline.color !== undefined) setLineColor(doc, line, outline.color, outline.transparency, existing);
    else if (!existing && !child(line, 'gradFill') && !child(line, 'pattFill')) {
      setLineColor(doc, line, '#000000', outline.transparency, existing);
    } else if (outline.transparency !== undefined && existing) {
      const color = children(existing).find(item => COLOR_CHILD_NAMES.includes(item.localName));
      if (color) setAlphaOnly(color, outline.transparency);
    }
  }
  if (outline.width !== undefined) {
    const width = finite(outline.width, '轮廓宽度');
    if (width < 0) fail('轮廓宽度不能为负数。');
    line.setAttribute('w', String(Math.round(width * EMU)));
  } else if (outline.visible === true && !line.hasAttribute('w')) line.setAttribute('w', String(EMU));
  if (outline.dashStyle !== undefined) {
    const key = String(outline.dashStyle).replace(/[\s_-]/g, '').toLowerCase();
    const dash = DASH_ALIASES[key] || outline.dashStyle;
    if (!DASH_TYPES.has(dash)) fail('轮廓线型无效：' + outline.dashStyle);
    const existing = child(line, 'prstDash');
    if (existing) existing.setAttribute('val', dash);
    else insertOrdered(line, create(doc, 'prstDash'), LN_ORDER).setAttribute('val', dash);
  }
  for (const [key, name] of [['beginArrow', 'headEnd'], ['endArrow', 'tailEnd']]) {
    if (outline[key] === undefined) continue;
    if (outline[key] === 'none' || outline[key] === false) {
      const existing = child(line, name);
      if (existing) line.removeChild(existing);
      continue;
    }
    if (!ARROW_TYPES.has(outline[key])) fail('箭头样式无效：' + outline[key]);
    const existing = child(line, name);
    if (existing) existing.setAttribute('type', outline[key]);
    else insertOrdered(line, create(doc, name), LN_ORDER).setAttribute('type', outline[key]);
  }
}

function ensureEffectList(doc, root) {
  const existing = child(root, 'effectLst');
  if (existing) return existing;
  if (child(root, 'effectDag')) fail('无法在已有的效果树中追加效果，请使用 font.raw 完整复制。');
  return ensureOrdered(doc, root, 'effectLst', RPR_ORDER);
}

function applyGlow(doc, root, glow) {
  if (glow === undefined) return;
  if (glow === false || glow === null) { removeEffects(root, ['glow']); return; }
  if (glow === true) glow = {};
  if (typeof glow !== 'object') fail('发光效果必须是对象。');
  const effect = ensureEffectList(doc, root);
  const node = ensureOrdered(doc, effect, 'glow', EFFECT_ORDER);
  if (glow.radius !== undefined) {
    const radius = finite(glow.radius, '发光半径');
    if (radius < 0) fail('发光半径不能为负数。');
    node.setAttribute('rad', String(Math.round(radius * EMU)));
  } else if (!node.hasAttribute('rad')) node.setAttribute('rad', String(5 * EMU));
  if (glow.color !== undefined) setColorChild(doc, node, glow.color, glow.transparency, node);
  else if (glow.transparency !== undefined) {
    if (!setAlphaOnly(node, glow.transparency) && !children(node).some(item => COLOR_CHILD_NAMES.includes(item.localName))) {
      setColorChild(doc, node, '#000000', glow.transparency, node);
    }
  } else if (!children(node).some(item => COLOR_CHILD_NAMES.includes(item.localName))) {
    setColorChild(doc, node, '#000000', undefined, node);
  }
}

function setShadowGeometry(doc, node, shadow) {
  let blur;
  if (shadow.blur !== undefined) blur = finite(shadow.blur, '阴影模糊');
  else if (shadow.size !== undefined) blur = finite(shadow.size, '阴影大小');
  if (blur !== undefined) {
    if (blur < 0) fail('阴影模糊不能为负数。');
    node.setAttribute('blurRad', String(Math.round(blur * EMU)));
  } else if (!node.hasAttribute('blurRad')) node.setAttribute('blurRad', String(Math.round(4 * EMU)));
  let distance = null, direction = null;
  if (shadow.distance !== undefined) distance = finite(shadow.distance, '阴影距离');
  if (shadow.direction !== undefined) direction = finite(shadow.direction, '阴影方向');
  if (shadow.offsetX !== undefined || shadow.offsetY !== undefined) {
    const x = finite(shadow.offsetX ?? 0, '阴影水平偏移');
    const y = finite(shadow.offsetY ?? 0, '阴影垂直偏移');
    distance = Math.hypot(x, y);
    direction = Math.atan2(y, x) * 180 / Math.PI;
  }
  if (distance !== null) {
    if (distance < 0) fail('阴影距离不能为负数。');
    node.setAttribute('dist', String(Math.round(distance * EMU)));
  } else if (!node.hasAttribute('dist')) node.setAttribute('dist', String(Math.round(3 * EMU)));
  if (direction !== null) {
    let value = Math.round(direction * 60000);
    value = ((value % 21600000) + 21600000) % 21600000;
    node.setAttribute('dir', String(value));
  } else if (!node.hasAttribute('dir')) node.setAttribute('dir', '5400000');
  if (!node.hasAttribute('algn')) node.setAttribute('algn', 'tl');
  if (!node.hasAttribute('rotWithShape')) node.setAttribute('rotWithShape', '0');
}

function applyShadow(doc, root, shadow, { removeInner = true } = {}) {
  if (shadow === undefined) return;
  if (shadow === false || shadow === null) {
    removeEffects(root, removeInner ? ['outerShdw', 'innerShdw'] : ['outerShdw']);
    return;
  }
  if (shadow === true) shadow = {};
  if (typeof shadow !== 'object') fail('阴影必须是布尔值或对象。');
  if (shadow.visible === false) {
    removeEffects(root, removeInner ? ['outerShdw', 'innerShdw'] : ['outerShdw']);
    return;
  }
  const effect = ensureEffectList(doc, root);
  const node = ensureOrdered(doc, effect, 'outerShdw', EFFECT_ORDER);
  setShadowGeometry(doc, node, shadow);
  if (shadow.color !== undefined) setColorChild(doc, node, shadow.color, shadow.transparency, node);
  else if (shadow.transparency !== undefined) {
    if (!setAlphaOnly(node, shadow.transparency) && !children(node).some(item => COLOR_CHILD_NAMES.includes(item.localName))) {
      setColorChild(doc, node, '#000000', shadow.transparency, node);
    }
  } else if (!children(node).some(item => COLOR_CHILD_NAMES.includes(item.localName))) {
    setColorChild(doc, node, '#000000', 0.6, node);
  }
}

function reflectionIndex(type) {
  if (typeof type === 'number' || /^\d+$/.test(String(type))) {
    const value = Number(type);
    if (Number.isInteger(value) && value >= 0 && value <= 9) return value;
  }
  const key = String(type).replace(/[\s_-]/g, '').toLowerCase();
  if (REFLECTION_TYPE_NUMBERS[key] !== undefined) return REFLECTION_TYPE_NUMBERS[key];
  if (key === 'custom' || key === 'mixed') return null;
  fail('倒影类型无效：' + type);
}

function applyReflection(doc, root, reflection) {
  if (reflection === undefined) return;
  if (reflection === false || reflection === null) { removeEffects(root, ['reflection']); return; }
  if (typeof reflection === 'number' || typeof reflection === 'string') reflection = { type: reflection };
  if (typeof reflection !== 'object') fail('倒影必须是对象。');
  let node = null;
  if (reflection.raw !== undefined) {
    node = parseRawElement(doc, reflection.raw, ['reflection'], '倒影', { links: false });
    const existing = child(child(root, 'effectLst'), 'reflection');
    if (existing) existing.parentNode.replaceChild(node, existing);
    else insertOrdered(ensureEffectList(doc, root), node, EFFECT_ORDER);
  }
  if (!node && reflection.type !== undefined && reflectionIndex(reflection.type) === 0) {
    removeEffects(root, ['reflection']); return;
  }
  if (!node) node = child(child(root, 'effectLst'), 'reflection');
  if (!node) fail('倒影样式需要原生 font.reflection.raw 数据。');
  if (reflection.opacity !== undefined) node.setAttribute('stA', String(Math.round(clamp(finite(reflection.opacity, '倒影透明度'), 0, 1) * 100000)));
  if (reflection.size !== undefined) node.setAttribute('endPos', String(Math.round(clamp(finite(reflection.size, '倒影大小'), 0, 1) * 100000)));
  if (reflection.distance !== undefined) node.setAttribute('dist', String(Math.round(finite(reflection.distance, '倒影距离') * EMU)));
  if (reflection.blur !== undefined) node.setAttribute('blurRad', String(Math.round(finite(reflection.blur, '倒影模糊') * EMU)));
  if (reflection.angle !== undefined) {
    let value = Math.round(finite(reflection.angle, '倒影角度') * 60000);
    value = ((value % 21600000) + 21600000) % 21600000;
    node.setAttribute('dir', String(value));
    node.setAttribute('fadeDir', String(value));
  }
}

function applyEmboss(doc, root, emboss) {
  if (emboss === undefined) return;
  fail('浮雕样式需要通过宿主应用设置，或提供原生 font.raw 数据。');
}

function applyFontFields(doc, root, font, mode, rawPresent) {
  const allowed = SCOPE_FIELDS[mode] || SCOPE_FIELDS.merge;
  const supports = key => allowed.includes(key);
  if (supports('nameAscii') || supports('nameFarEast') || supports('nameComplexScript')) {
    const names = {};
    if (supports('name') && font.name !== undefined) names.name = font.name;
    if (supports('nameAscii') && font.nameAscii !== undefined) names.nameAscii = font.nameAscii;
    if (supports('nameFarEast') && font.nameFarEast !== undefined) names.nameFarEast = font.nameFarEast;
    if (supports('nameComplexScript') && font.nameComplexScript !== undefined) names.nameComplexScript = font.nameComplexScript;
    setNames(doc, root, names);
  }
  if (supports('size') && font.size !== undefined) {
    const size = finite(font.size, '字号');
    if (size <= 0 || size > 4000) fail('字号必须在 0 到 4000 之间。');
    root.setAttribute('sz', String(Math.round(size * 100)));
  }
  if (supports('bold')) setBooleanAttribute(root, 'b', font.bold);
  if (supports('italic')) setBooleanAttribute(root, 'i', font.italic);
  if (supports('underline') || supports('underlineStyle')) {
    const value = underlineValue(font, root);
    if (value !== null) root.setAttribute('u', value);
  }
  if (supports('strikethrough')) setStrike(root, font);
  if (supports('superscript') || supports('subscript') || supports('baselineOffset')) setBaseline(root, font);
  const nativePaint = rawPresent && children(root).some(node => ['gradFill', 'blipFill', 'noFill'].includes(node.localName));
  if ((supports('color') || supports('transparency')) && !nativePaint) setFillColor(doc, root, font);
  if (supports('highlight')) applyHighlight(doc, root, font.highlight);
  if (!(rawPresent && child(root, 'ln'))) {
    if (supports('outline')) applyOutline(doc, root, font.outline);
    if (supports('line')) applyOutline(doc, root, font.line);
  }
  if (supports('glow') && !(rawPresent && child(child(root, 'effectLst'), 'glow'))) applyGlow(doc, root, font.glow);
  if (supports('shadow') && !(rawPresent && typeof font.shadow === 'boolean')) {
    applyShadow(doc, root, font.shadow, { removeInner: !rawPresent });
  }
  if (supports('reflection')) applyReflection(doc, root, font.reflection);
  if (supports('emboss') && !rawPresent) applyEmboss(doc, root, font.emboss);
}

function buildProperties(doc, existing, font, scope, name = 'rPr') {
  const mode = scope || (font.raw !== undefined ? 'all' : 'merge');
  if (!SCOPE_FIELDS[mode]) fail('文字格式范围无效：' + mode);
  const raw = font.raw !== undefined
    ? parseRawElement(doc, font.raw, ['rPr', 'defRPr', 'endParaRPr'], '字体样式', { links: true })
    : null;
  const links = existing ? targetLinks(existing) : [];
  let root;
  if (mode === 'all') root = raw || create(doc, 'rPr');
  else root = existing ? existing.cloneNode(true) : create(doc, 'rPr');
  removeLinks(root);
  if (mode === 'color') {
    clearColorProperties(root);
    if (raw) copyRawChildren(raw, root, COLOR_NAMES);
  } else if (mode === 'effect') {
    clearEffectProperties(root);
    if (raw) {
      if (raw.hasAttribute('baseline')) root.setAttribute('baseline', raw.getAttribute('baseline'));
      copyRawChildren(raw, root, EFFECT_NAMES);
    }
  } else if (mode === 'name') {
    clearNameProperties(root);
    if (raw) copyRawChildren(raw, root, NAME_NAMES);
  } else if (mode === 'size') {
    if (raw) {
      if (raw.hasAttribute('sz')) root.setAttribute('sz', raw.getAttribute('sz'));
      else root.removeAttribute('sz');
    }
  }
  applyFontFields(doc, root, font, mode, Boolean(raw));
  removeLinks(root);
  attachLinks(root, links);
  return renameProperties(doc, root, name);
}

function setTextRunProperties(run, properties) {
  const existing = child(run, 'rPr');
  if (existing) run.replaceChild(properties, existing);
  else run.insertBefore(properties, run.firstChild);
}

function paragraphText(paragraph) {
  return children(paragraph).filter(item => TEXT_NODES.has(item.localName))
    .map(item => segmentText(item)).join('');
}
function segmentText(node) {
  if (node.localName === 'br') return '\n';
  if (node.localName === 'm') return descendants(node, 't').map(textContent).join('');
  return textContent(child(node, 't'));
}
function isStyleable(node) { return node.localName === 'r' || node.localName === 'fld'; }

function buildLayout(txBody) {
  const paragraphs = children(txBody, 'p');
  const segments = [];
  let offset = 0;
  for (let index = 0; index < paragraphs.length; index++) {
    const paragraph = paragraphs[index];
    const paragraphSegments = [];
    for (const node of children(paragraph)) {
      if (!TEXT_NODES.has(node.localName)) continue;
      const text = segmentText(node);
      const segment = { type: node.localName, node, paragraph, text, start: offset, end: offset + text.length };
      paragraphSegments.push(segment);
      segments.push(segment);
      offset += text.length;
    }
    paragraph.paragraphSegments = paragraphSegments;
    if (index < paragraphs.length - 1) {
      const separator = { type: 'paragraph', node: null, paragraph, next: paragraphs[index + 1],
        text: '\n', start: offset, end: offset + 1 };
      paragraph.separator = separator;
      segments.push(separator);
      offset += 1;
    }
  }
  return { text: segments.map(item => item.text).join(''), paragraphs, segments };
}

function splitTextNode(doc, node, offset) {
  const t = child(node, 't');
  if (!t) return false;
  const text = t.textContent || '';
  if (offset <= 0 || offset >= text.length) return false;
  const clone = node.cloneNode(true);
  const cloneText = child(clone, 't');
  setPreservedText(t, text.slice(0, offset));
  setPreservedText(cloneText, text.slice(offset));
  node.parentNode.insertBefore(clone, node.nextSibling);
  return true;
}

function splitAt(doc, txBody, offset) {
  const layout = buildLayout(txBody);
  if (offset <= 0 || offset >= layout.text.length) return;
  const segment = layout.segments.find(item => item.start < offset && offset < item.end);
  if (segment && isStyleable(segment.node)) splitTextNode(doc, segment.node, offset - segment.start);
}

function paragraphTarget(txBody) {
  const paragraphs = children(txBody, 'p');
  if (paragraphs.length) return paragraphs[0];
  const paragraph = create(txBody.ownerDocument, 'p');
  txBody.appendChild(paragraph);
  return paragraph;
}
function ensureParagraphProperties(paragraph) {
  let properties = child(paragraph, 'pPr');
  if (!properties) {
    properties = create(paragraph.ownerDocument, 'pPr');
    paragraph.insertBefore(properties, paragraph.firstChild);
  }
  return properties;
}
function ensureEndMark(paragraph) {
  const doc = paragraph.ownerDocument;
  const existing = child(paragraph, 'endParaRPr');
  if (existing) return existing;
  const node = create(doc, 'endParaRPr');
  const ext = child(paragraph, 'extLst');
  if (ext) paragraph.insertBefore(node, ext);
  else paragraph.appendChild(node);
  return node;
}

function formatMath(doc, node, start, end, font, scope) {
  const runs = descendants(node).filter(item => item.localName === 'r' &&
    item.namespaceURI === 'http://schemas.openxmlformats.org/officeDocument/2006/math');
  let offset = 0;
  for (const run of runs) {
    const text = child(run, 't');
    const value = textContent(text);
    const length = value.length;
    const from = Math.max(0, start - offset), to = Math.min(length, end - offset);
    offset += length;
    if (from >= to) continue;
    let selected = run;
    if (from > 0) {
      selected = run.cloneNode(true);
      setPreservedText(text, value.slice(0, from));
      setPreservedText(child(selected, 't'), value.slice(from));
      run.parentNode.insertBefore(selected, run.nextSibling);
    }
    if (to < length) {
      const tail = selected.cloneNode(true);
      setPreservedText(child(tail, 't'), value.slice(to));
      setPreservedText(child(selected, 't'), value.slice(from, to));
      selected.parentNode.insertBefore(tail, selected.nextSibling);
    }
    const existing = children(selected).find(item => item.namespaceURI === A && item.localName === 'rPr');
    const properties = buildProperties(doc, existing, font, scope);
    if (existing) selected.replaceChild(properties, existing);
    else selected.insertBefore(properties, child(selected, 't'));
  }
}

function applyFontToTextNodes(doc, txBody, start, end, font, scope, { marks = false } = {}) {
  splitAt(doc, txBody, start);
  splitAt(doc, txBody, end);
  const layout = buildLayout(txBody);
  for (const segment of layout.segments) {
    if (segment.type === 'paragraph') {
      if (marks && segment.start >= start && segment.end <= end) {
        const endMark = ensureEndMark(segment.paragraph);
        segment.paragraph.replaceChild(buildProperties(doc, endMark, font, scope, 'endParaRPr'), endMark);
      }
      continue;
    }
    if (segment.type === 'm') {
      formatMath(doc, segment.node, Math.max(0, start - segment.start),
        Math.min(segment.text.length, end - segment.start), font, scope);
      continue;
    }
    if (!isStyleable(segment.node)) continue;
    if (segment.start < start || segment.end > end) continue;
    const runProperties = buildProperties(doc, backgroundRunProperties(segment.node), font, scope);
    setTextRunProperties(segment.node, runProperties);
  }
}

function formatWholeTextBody(doc, txBody, font, scope) {
  if (!child(txBody, 'p')) txBody.appendChild(create(doc, 'p'));
  for (const paragraph of children(txBody, 'p')) {
    for (const node of children(paragraph)) {
      if (node.localName === 'm') {
        formatMath(doc, node, 0, segmentText(node).length, font, scope);
        continue;
      }
      if (!isStyleable(node)) continue;
      const properties = buildProperties(doc, backgroundRunProperties(node), font, scope);
      setTextRunProperties(node, properties);
    }
    const paragraphProperties = ensureParagraphProperties(paragraph);
    const defaultProperties = child(paragraphProperties, 'defRPr');
    const next = buildProperties(doc, defaultProperties, font, scope, 'defRPr');
    if (defaultProperties) paragraphProperties.replaceChild(next, defaultProperties);
    else insertOrdered(paragraphProperties, next, PPR_ORDER);
    const endMark = ensureEndMark(paragraph);
    paragraph.replaceChild(buildProperties(doc, endMark, font, scope, 'endParaRPr'), endMark);
  }
}

function findRunsForProperties(properties, layout) {
  let runs;
  if (properties.paragraphs) {
    runs = [];
    for (const [index, paragraph] of properties.paragraphs.entries()) {
      if (index) runs.push({ text: '\n' });
      if (paragraph.runs) runs.push(...paragraph.runs);
      else if (paragraph.text !== undefined) runs.push({ text: String(paragraph.text), font: paragraph.font });
    }
  } else runs = properties.runs;
  if (!Array.isArray(runs)) return null;
  let offset = 0;
  const spans = [];
  for (const run of runs) {
    const text = run?.text === undefined ? '' : String(run.text);
    if (run?.font && text.length) spans.push({ start: offset, end: offset + text.length, font: run.font,
      scope: run.fontScope ?? properties.fontScope });
    offset += text.length;
  }
  if (offset !== layout.text.length) fail('文字运行与目标文本不一致，请刷新后重试。');
  return spans;
}

function applyRunFonts(doc, txBody, properties) {
  const layout = buildLayout(txBody);
  const spans = findRunsForProperties(properties, layout);
  if (!spans) return;
  for (const span of spans) applyFontToTextNodes(doc, txBody, span.start, span.end, span.font, span.scope);
}

function insertionStyle(doc, segments, index, paragraph) {
  const candidates = [
    segments[index],
    segments[index - 1],
    ...segments.slice(index + 1),
    ...segments.slice(0, index).reverse(),
  ].filter(Boolean);
  for (const segment of candidates) {
    if (!isStyleable(segment.node)) continue;
    const properties = backgroundRunProperties(segment.node);
    return properties ? properties.cloneNode(true) : null;
  }
  const defaults = child(child(paragraph, 'pPr'), 'defRPr');
  return defaults ? defaults.cloneNode(true) : null;
}

function insertTextRuns(doc, paragraph, before, text, sourceProperties) {
  const pieces = String(text).split('\n');
  let reference = before;
  const insert = node => {
    if (reference) paragraph.insertBefore(node, reference);
    else if (child(paragraph, 'endParaRPr')) paragraph.insertBefore(node, child(paragraph, 'endParaRPr'));
    else if (child(paragraph, 'extLst')) paragraph.insertBefore(node, child(paragraph, 'extLst'));
    else paragraph.appendChild(node);
  };
  for (let index = 0; index < pieces.length; index++) {
    if (index) {
      const br = create(doc, 'br');
      if (sourceProperties) br.appendChild(sourceProperties.cloneNode(true));
      insert(br);
    }
    if (!pieces[index]) continue;
    const run = create(doc, 'r');
    if (sourceProperties) run.appendChild(sourceProperties.cloneNode(true));
    const t = create(doc, 't');
    setPreservedText(t, pieces[index]);
    run.appendChild(t);
    insert(run);
  }
  if (!pieces.some(Boolean) && pieces.length === 1) return;
  // Keep the reference at the original position: all nodes above were inserted
  // before it, so the next insertion point remains valid.
}

function mergeParagraph(preceding, following) {
  const endMark = child(preceding, 'endParaRPr');
  if (endMark) preceding.removeChild(endMark);
  for (const node of children(following)) {
    if (node.localName === 'pPr' || node.localName === 'extLst') continue;
    preceding.appendChild(node);
  }
  if (preceding.parentNode) preceding.parentNode.removeChild(following);
}

function replaceTextRange(txBody, start, end, text) {
  const doc = txBody.ownerDocument;
  const layout = buildLayout(txBody);
  if (!layout.segments.length) {
    const paragraph = paragraphTarget(txBody);
    insertTextRuns(doc, paragraph, null, text, null);
    return;
  }
  const insideMath = layout.segments.find(item => item.type === 'm' && item.start < start && start < item.end)
    || layout.segments.find(item => item.type === 'm' && item.start < end && end < item.end);
  if (insideMath) fail('不能在公式内部替换文本；请选择完整公式或公式外文字。');

  splitAt(doc, txBody, start);
  splitAt(doc, txBody, end);
  const current = buildLayout(txBody);
  const startSegment = current.segments.find(item => item.start <= start && start < item.end)
    || current.segments.at(-1);
  const startIndex = Math.max(0, current.segments.indexOf(startSegment));
  let paragraph;
  let reference;
  if (start >= current.text.length) {
    paragraph = current.paragraphs.at(-1);
    reference = child(paragraph, 'endParaRPr') || child(paragraph, 'extLst');
  } else if (startSegment.type === 'paragraph') {
    paragraph = startSegment.paragraph;
    reference = child(paragraph, 'endParaRPr') || child(paragraph, 'extLst');
  } else {
    paragraph = startSegment.paragraph;
    reference = startSegment.node;
  }
  const properties = insertionStyle(doc, current.segments, startIndex, paragraph);
  const selected = current.segments.filter(item => item.start >= start && item.end <= end && item.end > item.start);

  // Insert first, then delete the original range. The inserted nodes are not part
  // of `selected`, so a reference that is about to be deleted still anchors them
  // at the correct position after the removal.
  insertTextRuns(doc, paragraph, reference, text, properties);
  const merges = [];
  for (const segment of selected) {
    if (segment.type === 'paragraph') merges.push(segment);
    else if (segment.node?.parentNode) segment.node.parentNode.removeChild(segment.node);
  }
  for (const separator of merges.reverse()) {
    const preceding = separator.paragraph;
    const following = separator.next;
    if (preceding?.parentNode && following?.parentNode && preceding.parentNode === following.parentNode) {
      mergeParagraph(preceding, following);
    }
  }
}

function shapeForId(document, shapeId) {
  const id = String(shapeId);
  const matches = allElements(document.documentElement || document).filter(node => node.localName === 'cNvPr'
    && String(attr(node, 'id')) === id);
  if (!matches.length) fail('未找到对象 ' + id + '。');
  for (const match of matches) {
    for (let parent = match.parentNode; parent; parent = parent.parentNode) {
      if (element(parent) && SHAPE_NODES.has(parent.localName)) return parent;
    }
  }
  fail('对象 ' + id + ' 不是可编辑的幻灯片形状。');
}

function textBodiesFor(shape) {
  if (shape.localName === 'grpSp' || shape.localName === 'graphicFrame') {
    return descendants(shape, 'txBody');
  }
  const body = child(shape, 'txBody');
  return body ? [body] : [];
}

function applyToTextBody(doc, txBody, properties) {
  if (properties.font !== undefined) formatWholeTextBody(doc, txBody, properties.font, properties.fontScope);
  if (properties.runs !== undefined || properties.paragraphs !== undefined) {
    applyRunFonts(doc, txBody, properties);
  }
  if (properties.textRange !== undefined) {
    const range = properties.textRange;
    if (!range || typeof range !== 'object') fail('文字选择范围无效。');
    const layout = buildLayout(txBody);
    const start = Number(range.start ?? 0);
    const length = Number(range.length ?? 0);
    if (!Number.isInteger(start) || start < 0) fail('文字选择起点无效。');
    if (!Number.isInteger(length) || length < 0) fail('文字选择长度无效。');
    if (start + length > layout.text.length) fail('文字选择范围已改变，请重新选择。');
    if (length === 0 && range.font !== undefined) fail('请选择至少一个文字字符来设置格式。');
    let targetLength = length;
    if (range.text !== undefined) {
      const replacement = String(range.text);
      replaceTextRange(txBody, start, start + length, replacement);
      targetLength = replacement.length;
    }
    if (range.font !== undefined) {
      applyFontToTextNodes(doc, txBody, start, start + targetLength, range.font,
        range.fontScope ?? properties.fontScope, { marks: true });
    }
  }
}

/**
 * Apply rich-text formatting to one shape inside a slide DrawingML document.
 *
 * Supported properties:
 * - font: normalized properties from lib/pptx.mjs font() (including raw XML)
 * - fontScope: all | name | size | color | effect
 * - runs / paragraphs: rich runs carrying text, font and original hyperlinks/math
 * - textRange: { start, length, font, text? } with JavaScript UTF-16 offsets
 *
 * The document is mutated synchronously. The return value summarizes the text that
 * remains in the first affected text body.
 */
export function applyTextFormatting(document, shapeId, properties = {}) {
  if (!document || (document.nodeType !== 9 && document.nodeType !== 1)) fail('缺少幻灯片 XML 文档。');
  if (!properties || typeof properties !== 'object') fail('缺少文字格式属性。');
  if (properties.fontScope !== undefined && !FONT_SCOPES.has(properties.fontScope)) {
    fail('文字格式范围必须是 all、name、size、color 或 effect。');
  }
  const shape = shapeForId(document, shapeId);
  const bodies = textBodiesFor(shape);
  if (!bodies.length) fail('所选对象没有可编辑的文字。');
  const doc = document.nodeType === 9 ? document : document.ownerDocument;
  const selectedOnly = properties.textRange !== undefined || properties.runs !== undefined || properties.paragraphs !== undefined;
  if (selectedOnly && bodies.length !== 1) fail('文字选择范围只能应用到单个文本框。');
  for (const body of bodies) applyToTextBody(doc, body, properties);
  const body = bodies[0];
  return { shapeId: String(shapeId), text: buildLayout(body).text };
}

export const pptxFormat = { applyTextFormatting };

// Exposed for the focused tests and for callers that need the same UTF-16 text
// layout that the formatter uses (paragraphs joined with "\n", line breaks as
// "\n", native math contributing its visible text).
export function textLayout(txBody) { return buildLayout(txBody); }
