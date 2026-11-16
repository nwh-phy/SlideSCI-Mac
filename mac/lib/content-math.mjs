import { readFileSync } from 'node:fs';
import vm from 'node:vm';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';

const XML_ESCAPES = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;', '\r': '&#13;', '\n': '&#10;', '\t': '&#9;' };
const XML_UNSAFE = /[&<>"'\r\n\t]/g;
export const xml = value => String(value ?? '').replace(XML_UNSAFE, char => XML_ESCAPES[char]);
const MATH_NS = 'http://www.w3.org/1998/Math/MathML';
let katexEngine;
let svgEngine;
let katexRenderScript;
let svgRenderScript;

export function normalizeLatex(source) {
  const text = String(source ?? '').trim();
  for (const [open, close] of [['$$', '$$'], ['\\(', '\\)'], ['\\[', '\\]'], ['$', '$']]) {
    if (text.startsWith(open) && text.endsWith(close) && text.length >= open.length + close.length) return text.slice(open.length, -close.length).trim();
  }
  return text;
}

function engine(file) {
  const context = vm.createContext({ console: { log() {}, warn() {}, error() {} } });
  new vm.Script(readFileSync(new URL(`../../SlideSCI/latex-converter/${file}`, import.meta.url), 'utf8'), { filename: file }).runInContext(context, { timeout: 15000 });
  return context;
}

export function latexToMathMlBatch(sources, display = false) {
  katexEngine ??= engine('katex/katex.min.js');
  katexRenderScript ??= new vm.Script(`(() => { const macros = Object.create(null); return sources.map((source, index) => katex.renderToString(source, {output:'mathml', displayMode:display[index], throwOnError:true, strict:'ignore', trust:false, maxExpand:1000, macros, globalGroup:true})); })()`);
  katexEngine.sources = sources.map(normalizeLatex);
  katexEngine.display = Array.isArray(display) ? display : sources.map(() => display);
  if (katexEngine.sources.some(source => !source)) throw new Error('LaTeX 公式不能为空。');
  try {
    return katexRenderScript.runInContext(katexEngine, { timeout: 5000 });
  } finally {
    delete katexEngine.sources;
    delete katexEngine.display;
  }
}

export function latexToSvg(source) {
  svgEngine ??= engine('mathjax/mathjax-svg.min.js');
  svgEngine.source = normalizeLatex(source);
  if (!svgEngine.source) throw new Error('LaTeX 公式不能为空。');
  svgEngine.display = /^\s*(\$\$|\\\[)/.test(source) || /\\\\|\\begin\s*\{(?:align|equation|gather|multline|cases|matrix)/.test(source);
  try {
    svgRenderScript ??= new vm.Script('slideSciToSvg(source, display)');
    const svg = svgRenderScript.runInContext(svgEngine, { timeout: 15000 });
    if (!/^<svg\b/.test(svg)) throw new Error('MathJax 未生成有效 SVG。');
    return svg;
  } finally {
    delete svgEngine.source;
    delete svgEngine.display;
  }
}

function parseXml(source) {
  if (/<!DOCTYPE|<!ENTITY/i.test(source)) throw new Error('不允许 XML 外部实体。');
  const errors = [];
  const document = new DOMParser({ onError: (level, message) => { if (level !== 'warning') errors.push(message); } }).parseFromString(source, 'text/xml');
  if (errors.length || !document.documentElement) throw new Error(`XML 无效：${errors.join('; ')}`);
  return document;
}

export function readSvg(source) {
  const document = parseXml(source);
  const root = document.documentElement;
  if (root.localName !== 'svg' || (root.namespaceURI && root.namespaceURI !== 'http://www.w3.org/2000/svg')) throw new Error('请输入完整的 SVG 文档。');
  if (!root.namespaceURI) root.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
  const externalStyle = value => /@import/i.test(value) || Array.from(value.matchAll(/\burl\s*\(([^)]*)\)/gi)).some(match => !match[1].trim().replace(/^(['"])(.*)\1$/, '$2').trim().startsWith('#'));
  const elements = [root, ...Array.from(root.getElementsByTagName('*'))];
  for (const node of elements) {
    if (['script', 'foreignObject', 'iframe', 'object', 'embed'].includes(node.localName)) throw new Error(`SVG 包含不允许的元素：${node.localName}`);
    for (const attribute of Array.from(node.attributes ?? [])) {
      if (/^on/i.test(attribute.name)) throw new Error('SVG 不允许事件处理程序。');
      if (/(?:^|:)href$/i.test(attribute.name) && attribute.value && !attribute.value.startsWith('#') && !/^data:image\/(?:png|jpeg|gif|webp);base64,/i.test(attribute.value)) throw new Error('SVG 外部资源必须嵌入文件。');
      if (externalStyle(attribute.value)) throw new Error('SVG 不允许外部样式或资源。');
    }
    if (node.localName === 'style' && externalStyle(node.textContent)) throw new Error('SVG 不允许外部样式或资源。');
  }
  const viewBox = (root.getAttribute('viewBox') || '').trim().split(/[\s,]+/).map(Number);
  const dimension = attribute => {
    const match = /^([\d.+-]+)(px|pt|in|cm|mm|em|ex)?$/.exec((root.getAttribute(attribute) || '').trim());
    return match ? Number(match[1]) * ({ px: 0.75, pt: 1, in: 72, cm: 72 / 2.54, mm: 72 / 25.4, em: 15, ex: 7.5 }[match[2] || 'px']) : 0;
  };
  const width = dimension('width') || (viewBox.length === 4 && viewBox[2] * 0.75) || 500;
  const height = dimension('height') || (viewBox.length === 4 && viewBox[3] * 0.75) || 200;
  if (!(width > 0 && height > 0)) throw new Error('SVG 尺寸必须为正数。');
  return { source: new XMLSerializer().serializeToString(document), width, height };
}

const children = node => Array.from(node.childNodes ?? []).filter(child => child.nodeType === 1);
const bool = value => value ? '1' : '0';
const prop = (tag, value) => `<m:${tag} m:val="${xml(value)}"/>`;
const element = (tag, value, properties = '') => `<m:${tag}>${properties}${value}</m:${tag}>`;
const arg = (tag, value) => element(tag, value || '<m:r><m:t></m:t></m:r>');

/** Convert the presentation MathML tree to editable Office Math, rejecting unknown semantics. */
export function mathMlToOmml(markup, options = {}) {
  const document = parseXml(markup);
  const root = document.getElementsByTagNameNS(MATH_NS, 'math')[0] ?? (document.documentElement.localName === 'math' ? document.documentElement : null);
  if (!root) throw new Error('KaTeX 未生成有效 MathML。');
  function convert(node, inherited = {}) {
    if (node.nodeType === 3) return node.nodeValue.trim() ? run(node.nodeValue, inherited) : '';
    if (node.nodeType !== 1) return '';
    const tag = node.localName;
    const c = children(node);
    const style = { ...inherited };
    if (node.hasAttribute('mathvariant')) style.variant = node.getAttribute('mathvariant');
    if (node.hasAttribute('mathcolor')) style.color = node.getAttribute('mathcolor');
    const visit = child => child ? convert(child, style) : '';
    const all = () => c.map(visit).join('');
    const script = (type, base, sub, sup) => element(type, arg('e', visit(base)) + (sub ? arg('sub', visit(sub)) : '') + (sup ? arg('sup', visit(sup)) : ''));
    const nary = (base, lower, upper, location = 'subSup') => {
      if (base?.localName !== 'mo' || !/^[∑∏∐∫∬∭∮∯∰⋀⋁⋂⋃⨀⨁⨂⨄⨆]$/u.test(base.textContent)) return null;
      return element('nary', '<m:naryPr>' + prop('chr', base.textContent) + prop('limLoc', location) + prop('subHide', bool(!lower)) + prop('supHide', bool(!upper)) + '</m:naryPr>' + arg('sub', visit(lower)) + arg('sup', visit(upper)) + arg('e', ''));
    };
    switch (tag) {
      case 'math': case 'mstyle': case 'semantics': return all();
      case 'mrow': {
        const first = c[0];
        const last = c.at(-1);
        if (c.length > 1 && first.localName === 'mo' && last.localName === 'mo' && first.getAttribute('fence') === 'true' && last.getAttribute('fence') === 'true') {
          return element('d', '<m:dPr>' + prop('begChr', first.textContent) + prop('endChr', last.textContent) + prop('grow', '1') + '</m:dPr>' + arg('e', c.slice(1, -1).map(visit).join('')));
        }
        return all();
      }
      case 'mpadded': {
        const zero = ['height', 'depth', 'width'].filter(attribute => /^0(?:px|em|ex|pt|%)?$/.test(node.getAttribute(attribute)));
        if (!zero.length) return all();
        return element('phant', '<m:phantPr>' + prop('show', '1') + zero.map(attribute => prop({ width: 'zeroWid', height: 'zeroAsc', depth: 'zeroDesc' }[attribute], '1')).join('') + '</m:phantPr>' + arg('e', all()));
      }
      case 'annotation': case 'annotation-xml': return '';
      case 'mi': case 'mn': case 'mo': case 'mtext': case 'ms': return run(node.textContent, { ...style, normal: tag === 'mtext' || tag === 'ms' || style.variant === 'normal' });
      case 'mspace': {
        const width = parseFloat(node.getAttribute('width')) || 0;
        return width > 0 ? run(width >= 1 ? '\u2003'.repeat(Math.min(20, Math.ceil(width))) : '\u2009', style) : '';
      }
      case 'mfrac': return element('f', '<m:fPr>' + prop('type', node.getAttribute('linethickness') === '0px' || node.getAttribute('linethickness') === '0' ? 'noBar' : node.getAttribute('bevelled') === 'true' ? 'skw' : 'bar') + '</m:fPr>' + arg('num', visit(c[0])) + arg('den', visit(c[1])));
      case 'msqrt': return element('rad', '<m:radPr>' + prop('degHide', '1') + '</m:radPr><m:deg/>' + arg('e', all()));
      case 'mroot': return element('rad', arg('deg', visit(c[1])) + arg('e', visit(c[0])));
      case 'msub': return nary(c[0], c[1]) || script('sSub', c[0], c[1]);
      case 'msup': return nary(c[0], null, c[1]) || script('sSup', c[0], null, c[1]);
      case 'msubsup': return nary(c[0], c[1], c[2]) || script('sSubSup', c[0], c[1], c[2]);
      case 'munder': case 'mover': case 'munderover': {
        const base = c[0];
        const lower = tag !== 'mover' ? c[1] : null;
        const upper = tag === 'munderover' ? c[2] : tag === 'mover' ? c[1] : null;
        const nativeNary = nary(base, lower, upper, 'undOvr');
        if (nativeNary) return nativeNary;
        const group = upper || lower;
        if (group && /^[⏟⏞⏜⏝]$/.test(group.textContent)) return element('groupChr', '<m:groupChrPr>' + prop('chr', group.textContent) + prop('pos', upper ? 'top' : 'bot') + prop('vertJc', upper ? 'bot' : 'top') + '</m:groupChrPr>' + arg('e', visit(base)));
        const accent = node.getAttribute(tag === 'munder' ? 'accentunder' : 'accent') === 'true';
        if (accent && upper) return element('acc', '<m:accPr>' + prop('chr', upper.textContent) + '</m:accPr>' + arg('e', visit(base)));
        if (accent && lower && /^[¯_‾―−]$/.test(lower.textContent)) return element('bar', '<m:barPr>' + prop('pos', 'bot') + '</m:barPr>' + arg('e', visit(base)));
        let result = visit(base);
        if (lower) result = element('limLow', arg('e', result) + arg('lim', visit(lower)));
        if (upper) result = element('limUpp', arg('e', result) + arg('lim', visit(upper)));
        return result;
      }
      case 'mtable': {
        const rows = c.filter(row => row.localName === 'mtr');
        if (rows.length !== c.length) throw new Error('原生公式不支持带标签的 MathML 表格。');
        const columns = Math.max(1, ...rows.map(row => children(row).length));
        const alignments = (node.getAttribute('columnalign') || 'center').trim().split(/\s+/);
        const matrixColumns = Array.from({ length: columns }, (_, index) => '<m:mc><m:mcPr>' + prop('count', '1') + prop('mcJc', { left: 'left', right: 'right', center: 'center' }[alignments[index] || alignments.at(-1)] || 'center') + '</m:mcPr></m:mc>').join('');
        return element('m', '<m:mPr><m:baseJc m:val="center"/><m:mcs>' + matrixColumns + '</m:mcs></m:mPr>' + rows.map(row => '<m:mr>' + Array.from({ length: columns }, (_, index) => arg('e', visit(children(row)[index]))).join('') + '</m:mr>').join(''));
      }
      case 'mtr': case 'mtd': return all();
      case 'mfenced': {
        const separators = node.getAttribute('separators') || ',';
        return element('d', '<m:dPr>' + prop('begChr', node.hasAttribute('open') ? node.getAttribute('open') : '(') + prop('sepChr', separators[0]) + prop('endChr', node.hasAttribute('close') ? node.getAttribute('close') : ')') + '</m:dPr>' + c.map(child => arg('e', visit(child))).join(''));
      }
      case 'mphantom': return element('phant', '<m:phantPr>' + prop('show', '0') + '</m:phantPr>' + arg('e', all()));
      case 'menclose': {
        const notation = (node.getAttribute('notation') || 'longdiv').split(/\s+/);
        const allowed = ['box', 'top', 'bottom', 'left', 'right', 'updiagonalstrike', 'downdiagonalstrike', 'verticalstrike', 'horizontalstrike'];
        if (notation.some(value => !allowed.includes(value))) throw new Error(`原生公式不支持 MathML menclose：${notation.join(' ')}`);
        return element('borderBox', '<m:borderBoxPr>' + ['Top', 'Bot', 'Left', 'Right'].map((side, index) => prop(`hide${side}`, bool(!notation.some(value => ['box', ['top', 'bottom', 'left', 'right'][index]].includes(value))))).join('') + prop('strikeH', bool(notation.includes('horizontalstrike'))) + prop('strikeV', bool(notation.includes('verticalstrike'))) + prop('strikeBLTR', bool(notation.includes('updiagonalstrike'))) + prop('strikeTLBR', bool(notation.includes('downdiagonalstrike'))) + '</m:borderBoxPr>' + arg('e', all()));
      }
      case 'mmultiscripts': {
        let result = visit(c[0]);
        const pre = c.findIndex(child => child.localName === 'mprescripts');
        const post = c.slice(1, pre < 0 ? c.length : pre);
        for (let index = 0; index < post.length; index += 2) result = element('sSubSup', arg('e', result) + arg('sub', visit(post[index])) + arg('sup', visit(post[index + 1])));
        if (pre >= 0) {
          const scripts = c.slice(pre + 1);
          for (let index = 0; index < scripts.length; index += 2) result = element('sPre', arg('sub', visit(scripts[index])) + arg('sup', visit(scripts[index + 1])) + arg('e', result));
        }
        return result;
      }
      case 'none': return '';
      default: throw new Error(`原生公式不支持 MathML 元素 <${tag}>。`);
    }
  }
  function run(text, style) {
    const variants = {
      normal: ['p', 'roman'], bold: ['b', 'roman'], italic: ['i', 'roman'], 'bold-italic': ['bi', 'roman'],
      'double-struck': ['p', 'double-struck'], script: ['p', 'script'], 'bold-script': ['b', 'script'], fraktur: ['p', 'fraktur'], 'bold-fraktur': ['b', 'fraktur'],
      'sans-serif': ['p', 'sans-serif'], 'bold-sans-serif': ['b', 'sans-serif'], 'sans-serif-italic': ['i', 'sans-serif'], 'sans-serif-bold-italic': ['bi', 'sans-serif'], monospace: ['p', 'monospace']
    };
    const variant = variants[style.variant];
    if (style.variant && !variant) throw new Error(`原生公式不支持字体变体 ${style.variant}。`);
    const properties = style.normal ? '<m:nor/>' : variant ? prop('scr', variant[1]) + prop('sty', variant[0]) : '';
    const colors = { black: '#000000', white: '#FFFFFF', red: '#FF0000', green: '#008000', blue: '#0000FF', yellow: '#FFFF00', cyan: '#00FFFF', aqua: '#00FFFF', magenta: '#FF00FF', fuchsia: '#FF00FF', gray: '#808080', grey: '#808080', silver: '#C0C0C0', maroon: '#800000', olive: '#808000', lime: '#00FF00', teal: '#008080', navy: '#000080', purple: '#800080', orange: '#FFA500', brown: '#A52A2A', pink: '#FFC0CB', violet: '#EE82EE' };
    let colorValue = style.color ? colors[style.color.toLowerCase()] || style.color : '';
    if (/^#[\da-f]{3}$/i.test(colorValue)) colorValue = '#' + Array.from(colorValue.slice(1), char => char + char).join('');
    if (colorValue && !/^#[\da-f]{6}$/i.test(colorValue)) throw new Error(`原生公式不支持颜色 ${style.color}。`);
    const color = colorValue ? `<a:solidFill><a:srgbClr val="${colorValue.slice(1)}"/></a:solidFill>` : '';
    const drawing = `<a:rPr${options.size ? ` sz="${Math.round(options.size * 100)}"` : ''}>${color}<a:latin typeface="Cambria Math"/><a:ea typeface="Cambria Math"/></a:rPr>`;
    return `<m:r>${properties ? `<m:rPr>${properties}</m:rPr>` : ''}${drawing}<m:t xml:space="preserve">${xml(text)}</m:t></m:r>`;
  }
  return `<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${convert(root)}</m:oMath>`;
}
