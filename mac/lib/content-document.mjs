import { Marked } from 'marked';
import hljs from 'highlight.js';
import { DOMParser } from '@xmldom/xmldom';
import { latexToMathMlBatch, mathMlToOmml, latexToSvg, normalizeLatex, readSvg } from './content-math.mjs';
import { measureNativeMath } from './content-metrics.mjs';

const namedEntities = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: '\u00a0' };
const decode = source => String(source).replace(/&(#(?:x[\da-f]+|\d+)|amp|lt|gt|quot|apos|nbsp);/gi, (match, entity) => {
  if (entity[0] === '#') {
    const value = entity[1].toLowerCase() === 'x' ? parseInt(entity.slice(2), 16) : parseInt(entity.slice(1), 10);
    return value > 0 && value <= 0x10ffff ? String.fromCodePoint(value) : match;
  }
  return namedEntities[entity.toLowerCase()];
});
const parser = new Marked({ gfm: true, breaks: false });
parser.use({ extensions: [
  {
    name: 'nativeMathBlock', level: 'block', start: source => source.search(/\$\$|\\\[/),
    tokenizer(source) {
      const match = /^(?:\$\$([\s\S]*?)\$\$|\\\[([\s\S]*?)\\\])(?:[ \t]*\n|$)?/.exec(source);
      return match ? { type: 'nativeMathBlock', raw: match[0], source: match[1] ?? match[2], display: true } : undefined;
    }
  },
  {
    name: 'nativeMath', level: 'inline', start: source => source.search(/\$|\\\(/),
    tokenizer(source) {
      const match = /^(?:\\\(((?:\\.|[^\\])*?)\\\)|\$(?!\$)((?:\\.|[^$\\\n])+?)\$(?!\$))/.exec(source);
      return match ? { type: 'nativeMath', raw: match[0], source: match[1] ?? match[2] } : undefined;
    }
  },
  {
    name: 'advancedInline', level: 'inline', start: source => source.search(/[~^+=*]|__/),
    tokenizer(source) {
      const match = /^(?:\^([^\s^](?:[^\n^]*?[^\s^])?)\^|~(?!~)([^\s~](?:[^\n~]*?[^\s~])?)~(?!~)|\+\+([^\n]+?)\+\+|==([^\n]+?)==|\*\*([\s\S]+?)\*\*|__([\s\S]+?)__)/.exec(source);
      if (!match) return undefined;
      const index = match.findIndex((value, index) => index > 0 && value !== undefined);
      return { type: 'advancedInline', raw: match[0], style: ['superscript', 'subscript', 'underline', 'highlight', 'bold', 'bold'][index - 1], tokens: this.lexer.inlineTokens(match[index]) };
    }
  }
] });

const themes = {
  dark: { background: '#1E1E1E', foreground: '#FFFFFF', keyword: '#569CD6', comment: '#57A64A', string: '#D69D85', number: '#B5CEA8', property: '#9CDCFE', selector: '#D7BA7D' },
  light: { background: '#FFFFFF', foreground: '#000000', keyword: '#0000FF', comment: '#008000', string: '#A31515', number: '#098658', property: '#0086D1', selector: '#A86300' }
};
const languageAliases = { js: 'javascript', 'c#': 'csharp', cs: 'csharp', py: 'python', m: 'matlab', htm: 'xml', html: 'xml', f90: 'fortran', f95: 'fortran', f03: 'fortran', f08: 'fortran', f77: 'fortran', for: 'fortran', f: 'fortran' };

export function highlightedRuns(source, language = 'text', theme = 'light', fontSize = 12) {
  const palette = themes[theme];
  if (!palette) throw new Error('代码主题必须为 light 或 dark。');
  const canonical = languageAliases[String(language).toLowerCase()] || String(language).toLowerCase();
  const base = { name: 'Menlo', size: fontSize, color: palette.foreground };
  if (['text', 'plaintext', 'txt', ''].includes(canonical)) return [{ text: source, font: base }];
  if (!hljs.getLanguage(canonical)) throw new Error(`不支持代码语言“${language}”。请选择已安装的语言或纯文本。`);
  const markup = hljs.highlight(source, { language: canonical, ignoreIllegals: true }).value;
  const document = new DOMParser().parseFromString(`<root>${markup}</root>`, 'text/xml');
  const runs = [];
  function visit(node, font) {
    if (node.nodeType === 3) { if (node.nodeValue) runs.push({ text: node.nodeValue, font }); return; }
    const classes = node.getAttribute?.('class') || '';
    let type;
    if (/comment|quote/.test(classes)) type = 'comment';
    else if (/string|regexp|symbol/.test(classes)) type = 'string';
    else if (/number|literal/.test(classes)) type = 'number';
    else if (/selector|title|section|tag/.test(classes)) type = 'selector';
    else if (/property|attr|variable|built_in/.test(classes)) type = 'property';
    else if (/keyword|type|name|meta/.test(classes)) type = 'keyword';
    const next = { ...font, ...(type ? { color: palette[type] } : {}) };
    for (const child of Array.from(node.childNodes || [])) visit(child, next);
  }
  visit(document.documentElement, base);
  return runs;
}

function safeLink(value) {
  const url = decode(value || '').trim();
  if (/^(?:https?:|mailto:|tel:)/i.test(url) || /^#[^\s]*$/.test(url)) return url;
  throw new Error(`链接协议不受支持：${url}`);
}

function inline(tokens, inherited = {}, state = { stack: [] }) {
  const result = [];
  const add = (text, font = inherited, extra = {}) => { if (text) result.push({ text: decode(text), font: { ...font }, ...extra }); };
  for (const token of tokens || []) {
    const font = { ...inherited, ...Object.assign({}, ...state.stack.map(entry => entry.font)) };
    switch (token.type) {
      case 'nativeMath': result.push({ text: token.raw, mathSource: token.source, font }); break;
      case 'advancedInline': result.push(...inline(token.tokens, { ...font, [token.style]: true }, state)); break;
      case 'strong': result.push(...inline(token.tokens, { ...font, bold: true }, state)); break;
      case 'em': result.push(...inline(token.tokens, { ...font, italic: true }, state)); break;
      case 'del': result.push(...inline(token.tokens, { ...font, strikethrough: true }, state)); break;
      case 'link': result.push(...inline(token.tokens, { ...font, color: '#0563C1', underline: true }, state).map(run => ({ ...run, hyperlink: safeLink(token.href) }))); break;
      case 'codespan': result.push({ text: token.text, font: { ...font, name: 'Menlo', color: '#C00000' } }); break;
      case 'br': add('\n', font); break;
      case 'escape': add(token.text, font); break;
      case 'html': {
        const tag = /^<\s*(\/?)\s*([a-z][\w-]*)(?:\s[^>]*)?\s*\/?>$/i.exec(token.raw);
        if (!tag) { add(token.raw, font); break; }
        const name = tag[2].toLowerCase();
        const styles = { strong: { bold: true }, b: { bold: true }, em: { italic: true }, i: { italic: true }, u: { underline: true }, s: { strikethrough: true }, del: { strikethrough: true }, strike: { strikethrough: true }, sup: { superscript: true }, sub: { subscript: true }, code: { name: 'Menlo', color: '#C00000' }, mark: { highlight: true }, span: {} };
        if (name === 'br') add('\n', font);
        else if (styles[name]) {
          if (tag[1]) { const index = state.stack.map(entry => entry.name).lastIndexOf(name); if (index >= 0) state.stack.splice(index, 1); }
          else state.stack.push({ name, font: styles[name] });
        } else if (!['input', 'p', 'div'].includes(name)) add(token.raw, font);
        break;
      }
      case 'image': add(token.text || token.href, font, { hyperlink: safeLink(token.href) }); break;
      default: if (token.tokens) result.push(...inline(token.tokens, font, state)); else add(token.text ?? token.raw ?? '', font);
    }
  }
  return result;
}

const fullWidthCharacter = /[\u2e80-\uffff]/u;
const narrowCharacter = /[ilI.,'!|:;]/;
const wideCharacter = /[MW@%]/;
const monospaceFont = /Menlo|Consolas|Mono/;
const linePart = /(\n)/;
function textWidth(text, size, monospace = false) {
  let width = 0;
  for (const character of text) width += size * (character === '\t' ? 2.4 : fullWidthCharacter.test(character) ? 1 : monospace ? 0.6 : narrowCharacter.test(character) ? 0.28 : wideCharacter.test(character) ? 0.85 : 0.53);
  return width;
}

export function estimateParagraph(paragraph, width, defaultSize) {
  const size = paragraph.fontSize || defaultSize;
  const usable = Math.max(size, width - (paragraph.marginLeft || 0) - (paragraph.marginRight || 0));
  let height = 0;
  let used = 0;
  let lineHeight = size * 1.28;
  const newline = () => { height += lineHeight; used = 0; lineHeight = size * 1.28; };
  for (const run of paragraph.runs) {
    const runSize = run.font?.size || size;
    if (run.mathMetrics) {
      const metrics = run.mathMetrics;
      if (used + metrics.width > usable && used > 0) newline();
      used += metrics.width;
      lineHeight = Math.max(lineHeight, (metrics.ascent + metrics.descent) * 1.2);
      continue;
    }
    const monospace = monospaceFont.test(run.font?.name || '');
    for (const part of run.text.split(linePart)) {
      if (part === '\n') { newline(); continue; }
      for (const character of part) {
        const advance = textWidth(character, runSize, monospace);
        if (used + advance > usable && used > 0) newline();
        used += advance;
        lineHeight = Math.max(lineHeight, runSize * 1.28);
      }
    }
  }
  return height + lineHeight + (paragraph.spaceBefore || 0) + (paragraph.spaceAfter ?? 6);
}

export function buildContentDocument(input) {
  const kind = input.kind || 'markdown';
  const source = String(input.source ?? '').replace(/\r\n?/g, '\n');
  if (!source.trim()) throw new Error('请输入要插入的内容。');
  const finite = (value, fallback, min = 0) => value === undefined ? fallback : Number.isFinite(Number(value)) && Number(value) > min ? Number(value) : (() => { throw new Error('内容尺寸或字号无效。'); })();
  const width = finite(input.width, 500);
  const fontSize = finite(input.fontSize, kind === 'code' ? 12 : 18);
  const left = Number.isFinite(Number(input.left)) ? Number(input.left) : 60;
  const top = Number.isFinite(Number(input.top)) ? Number(input.top) : 50;
  const fontFamily = String(input.fontFamily || 'Arial');
  const theme = input.theme || 'light';
  const shapes = [];
  const warnings = [];
  let y = top;
  const gap = 10;
  function text(paragraphs, options = {}) {
    const shape = { kind: 'text', left: left + (options.indent || 0), top: y, width: width - (options.indent || 0), fontSize, fontFamily, paragraphs, ...options };
    shape.height = paragraphs.reduce((sum, paragraph) => sum + estimateParagraph(paragraph, shape.width - (options.padding || 0) * 2, fontSize), 0) + (options.padding || 0) * 2;
    shapes.push(shape);
    y += shape.height + gap;
    return shape;
  }
  function svg(value, alt, options = {}) {
    const data = readSvg(value);
    const scaledWidth = Math.min(width, options.natural ? data.width * 2 : width);
    const shape = { kind: 'svg', left, top: y, width: scaledWidth, height: data.height * scaledWidth / data.width, source: data.source, alternativeText: alt, ...options };
    if (options.center) shape.left += (width - scaledWidth) / 2;
    shapes.push(shape);
    y += shape.height + gap;
  }
  function code(value, language = 'text', indent = 0) {
    const palette = themes[theme];
    if (!palette) throw new Error('代码主题必须为 light 或 dark。');
    const codeSize = input.codeFontSize || (kind === 'code' ? fontSize : 12);
    text([{ runs: highlightedRuns(value, language, theme, codeSize), fontSize: codeSize, spaceAfter: 0 }], { indent, fill: palette.background, line: '#C8C8C8', padding: 10, fontFamily: 'Menlo', name: `SlideSCI code (${language})`, source: value });
  }
  function process(tokens, level = 0, quote = false) {
    for (const token of tokens) {
      switch (token.type) {
        case 'space': break;
        case 'heading': text([{ runs: inline(token.tokens, { bold: true }), fontSize: fontSize * [2, 1.7, 1.4, 1.2, 1.1, 1][token.depth - 1], spaceAfter: 8 }], { indent: level * 24 }); break;
        case 'paragraph': case 'text': text([{ runs: inline(token.tokens || [{ type: 'text', text: token.text }]) }], { indent: level * 24 }); break;
        case 'nativeMathBlock': text([{ runs: [{ text: token.raw.trim(), mathSource: token.source, display: true, font: {} }], alignment: 'center', spaceAfter: 4 }], { indent: level * 24, name: 'SlideSCI editable equation' }); break;
        case 'code': if (String(token.lang).toLowerCase() === 'svg' || /^\s*<svg\b/i.test(token.text)) svg(token.text, token.text); else code(token.text, (token.lang || 'text').split(/\s/)[0], level * 24); break;
        case 'html': if (/^\s*<svg\b/i.test(token.text || token.raw)) svg(token.text || token.raw, token.text || token.raw); else text([{ runs: inline(parser.lexer(token.text || token.raw).flatMap(part => part.tokens || [{ type: 'text', text: part.text || part.raw }])) }], { indent: level * 24 }); break;
        case 'list': {
          let ordinal = Number(token.start) || 1;
          for (const item of token.items) {
            let first = true;
            for (const part of item.tokens) {
              // GFM task lists emit a standalone checkbox token; item.checked already supplies the ☑/☐ marker.
              if (part.type === 'checkbox') continue;
              if (part.type === 'list') { process([part], level + 1, quote); continue; }
              if (['text', 'paragraph'].includes(part.type)) {
                const marker = item.task ? (item.checked ? '☑' : '☐') : token.ordered ? `${ordinal}.` : '•';
                const hanging = Math.max(15, textWidth(marker, fontSize) + 5);
                const paragraph = { runs: inline(part.tokens || [{ type: 'text', text: part.text }]), level, marginLeft: Math.max(24, hanging + 5), hanging, bullet: first ? { type: item.task ? 'task' : token.ordered ? 'ordered' : 'unordered', marker, start: ordinal, checked: item.checked } : null, spaceAfter: 3 };
                text([paragraph], { indent: level * 24 });
                y -= gap - 2;
                first = false;
              } else process([part], level + 1, quote);
            }
            ordinal++;
          }
          break;
        }
        case 'blockquote': {
          const start = y;
          const startIndex = shapes.length;
          process(token.tokens, level + 1, true);
          const height = Math.max(fontSize * 1.5, y - start - gap + 10);
          for (const shape of shapes.slice(startIndex)) shape.top += 5;
          shapes.splice(startIndex, 0, { kind: 'rect', left: left + level * 24, top: start, width: width - level * 24, height, fill: '#F5F5F5', line: '#000000', name: 'SlideSCI quote', quoteShapes: shapes.slice(startIndex) });
          y = start + height + gap;
          break;
        }
        case 'table': {
          const rows = [token.header, ...token.rows].map((row, index) => row.map((cell, column) => ({ paragraphs: [{ runs: inline(cell.tokens, index === 0 ? { bold: true } : {}), alignment: token.align[column] || 'left', spaceAfter: 0 }], fill: index === 0 ? '#EDEDED' : '#FFFFFF' })));
          const tableWidth = width - level * 24;
          const rowHeights = rows.map(row => Math.max(...row.map(cell => cell.paragraphs.reduce((height, paragraph) => height + estimateParagraph(paragraph, tableWidth / row.length - 12, fontSize), 0))) + 12);
          const shape = { kind: 'table', left: left + level * 24, top: y, width: tableWidth, height: rowHeights.reduce((a, b) => a + b, 0), rows, rowHeights, fontSize, fontFamily, name: 'SlideSCI table' };
          shapes.push(shape); y += shape.height + gap; break;
        }
        case 'hr': shapes.push({ kind: 'rect', left: left + level * 24, top: y + 4, width: width - level * 24, height: 0.8, fill: '#808080', name: 'SlideSCI divider' }); y += 15; break;
        case 'def': break;
        default: throw new Error(`不支持的 Markdown 块：${token.type}`);
      }
    }
  }
  if (kind === 'markdown') process(parser.lexer(source));
  else if (kind === 'code') code(source, input.language || 'python');
  else if (kind === 'latex') text([{ runs: [{ text: source, mathSource: normalizeLatex(source), display: true, font: {} }], alignment: 'center' }], { name: 'SlideSCI editable equation' });
  else if (kind === 'latex-svg') svg(latexToSvg(source), normalizeLatex(source), { natural: true, center: true, name: 'SlideSCI LaTeX SVG' });
  else if (kind === 'svg') svg(source, source, { name: 'SlideSCI SVG' });
  else throw new Error(`不支持的内容类型：${kind}`);
  const mathRuns = [];
  const mathShapes = new Set();
  for (const shape of shapes) {
    const paragraphs = shape.paragraphs || shape.rows?.flatMap(row => row.flatMap(cell => cell.paragraphs)) || [];
    for (const paragraph of paragraphs) for (const run of paragraph.runs) if (run.mathSource !== undefined) {
      run.mathFontSize = run.font?.size || paragraph.fontSize || shape.fontSize || fontSize;
      mathRuns.push(run);
      mathShapes.add(shape);
    }
  }
  if (mathRuns.length) {
    try {
      const mathMl = latexToMathMlBatch(mathRuns.map(run => run.mathSource), mathRuns.map(run => Boolean(run.display)));
      mathMl.forEach((markup, index) => {
        try {
          const run = mathRuns[index];
          run.omml = mathMlToOmml(markup, { size: run.mathFontSize });
          run.mathMetrics = measureNativeMath(run.omml, run.mathFontSize);
        }
        catch (error) { warnings.push({ source: mathRuns[index].text, message: `行内公式转换失败，将保留公式源码和正文：${error.message}` }); }
      });
    } catch (error) {
      warnings.push({ source: mathRuns.map(run => run.text).join('\n'), message: `行内公式转换失败，将保留公式源码和正文：${error.message}` });
    }
  }
  // Reflow after fractions, roots, matrices and scripts acquire their native math height.
  if (mathRuns.length) {
    let offset = 0;
    for (const shape of shapes) {
      shape.top += offset;
      if (shape.kind === 'text' && mathShapes.has(shape)) {
        const height = shape.paragraphs.reduce((sum, paragraph) => sum + estimateParagraph(paragraph, shape.width - (shape.padding || 0) * 2, fontSize), 0) + (shape.padding || 0) * 2;
        offset += height - shape.height;
        shape.height = height;
      } else if (shape.kind === 'table' && mathShapes.has(shape)) {
        const heights = shape.rows.map(row => Math.max(...row.map(cell => cell.paragraphs.reduce((height, paragraph) => height + estimateParagraph(paragraph, shape.width / row.length - 12, fontSize), 0))) + 12);
        const height = heights.reduce((a, b) => a + b, 0);
        offset += height - shape.height;
        shape.height = height;
        shape.rowHeights = heights;
      }
    }
  }
  for (const shape of [...shapes].reverse()) if (shape.quoteShapes?.length) shape.height = Math.max(...shape.quoteShapes.map(child => child.top + child.height)) - shape.top + 5;
  if (!shapes.length) throw new Error('没有可插入的内容。');
  const minLeft = Math.min(...shapes.map(shape => shape.left));
  const minTop = Math.min(...shapes.map(shape => shape.top));
  const bounds = { left: minLeft, top: minTop, width: Math.max(...shapes.map(shape => shape.left + shape.width)) - minLeft, height: Math.max(...shapes.map(shape => shape.top + shape.height)) - minTop };
  return { shapes, warnings, bounds, source, kind, fontSize, fontFamily, slideWidth: finite(input.slideWidth, 960), slideHeight: finite(input.slideHeight, 540) };
}
