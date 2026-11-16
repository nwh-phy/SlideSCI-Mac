import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import { DOMParser } from '@xmldom/xmldom';
import { normalizeLatex, latexToMathMlBatch, mathMlToOmml, latexToSvg, readSvg } from '../lib/content-math.mjs';
import { buildContentDocument, highlightedRuns } from '../lib/content-document.mjs';
import { contentHandlers } from '../lib/content.mjs';
import { parsePresentation, extractFragment } from '../lib/pptx.mjs';

const MATH = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
const DRAWING = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const math = expression => `<math xmlns="http://www.w3.org/1998/Math/MathML">${expression}</math>`;
const paragraphs = document => document.shapes.flatMap(shape => shape.paragraphs || shape.rows?.flatMap(row => row.flatMap(cell => cell.paragraphs)) || []);
const runs = document => paragraphs(document).flatMap(paragraph => paragraph.runs);

 test('native math converts fractions, radicals, prescripts, limits and matrices without images', () => {
  const omml = mathMlToOmml(math('<mrow><mo fence="true">(</mo><mfrac><mi>a</mi><msqrt><mi>b</mi></msqrt></mfrac><mo fence="true">)</mo></mrow><munderover><mo>∑</mo><mi>i</mi><mi>n</mi></munderover><mmultiscripts><mi>x</mi><mprescripts/><mi>a</mi><mi>b</mi></mmultiscripts><mtable columnalign="left right"><mtr><mtd><mi>x</mi></mtd><mtd><mn>1</mn></mtd></mtr></mtable>'), { size: 24 });
  const document = new DOMParser().parseFromString(omml, 'text/xml');
  for (const tag of ['f', 'rad', 'd', 'nary', 'sPre', 'm']) assert.ok(document.getElementsByTagNameNS(MATH, tag).length, `Missing native ${tag}`);
  assert.equal(document.getElementsByTagNameNS(DRAWING, 'rPr')[0].getAttribute('sz'), '2400');
  assert.doesNotMatch(omml, /svg|blip|annotation/);
  assert.throws(() => mathMlToOmml(math('<maction><mi>x</mi></maction>')), /不支持 MathML 元素/);
});

test('equation macros are shared inside one conversion and reset for the next conversion', () => {
  const result = latexToMathMlBatch(['\\gdef\\slidescifoo{x}', '\\slidescifoo+1']);
  assert.match(result[1], /<mi>x<\/mi>/);
  assert.throws(() => latexToMathMlBatch(['\\slidescifoo']), /Undefined control sequence/);
  assert.equal(normalizeLatex(' \\[ a+b \\] '), 'a+b');
});

test('Markdown retains inline styles, CJK bold, links, scripts and editable inline equations', () => {
  const document = buildContentDocument({ kind: 'markdown', source: '【**加粗**】 *斜体* ~~删除~~ ++下划线++ ==高亮== H~2~O x^2^ <sub>n</sub> [站点](https://example.com) `a < b` 与 $\\frac{a}{b}$。' });
  const values = runs(document);
  for (const [text, property] of [['加粗', 'bold'], ['斜体', 'italic'], ['删除', 'strikethrough'], ['下划线', 'underline'], ['高亮', 'highlight'], ['2', 'subscript'], ['n', 'subscript']]) assert.ok(values.some(run => run.text === text && run.font[property]), `${text} lost ${property}`);
  assert.ok(values.some(run => run.text === '2' && run.font.superscript));
  assert.ok(values.some(run => run.hyperlink === 'https://example.com'));
  assert.ok(values.some(run => run.text === 'a < b' && run.font.name === 'Menlo'));
  assert.ok(values.some(run => run.omml?.includes('<m:f>')));
  assert.deepEqual(document.warnings, []);
});

test('nested ordered task lists keep numbering and hanging indents', () => {
  const document = buildContentDocument({ kind: 'markdown', source: '3. first\n   - [x] nested done\n   - [ ] nested open\n4. second\n\nAfter.' });
  const values = paragraphs(document).filter(paragraph => paragraph.bullet);
  assert.deepEqual(values.filter(paragraph => paragraph.bullet.type === 'ordered').map(paragraph => paragraph.bullet.start), [3, 4]);
  assert.deepEqual(values.filter(paragraph => paragraph.bullet.type === 'task').map(paragraph => paragraph.bullet.marker), ['☑', '☐']);
  assert.ok(values.every(paragraph => paragraph.marginLeft > 0 && paragraph.hanging > 0));
  assert.ok(values.some(paragraph => paragraph.level === 1));
});

test('code syntax highlighting preserves every character including string-contained comment markers', () => {
  const source = 'url = "https://example.com/#tag"\n# comment\nprint(42)';
  const values = highlightedRuns(source, 'py', 'dark', 16);
  assert.equal(values.map(run => run.text).join(''), source);
  assert.ok(values.some(run => run.text.includes('https://') && run.font.color === '#D69D85'));
  assert.ok(values.some(run => run.text.includes('# comment') && run.font.color === '#57A64A'));
  const document = buildContentDocument({ kind: 'code', source, language: 'python', fontSize: 16, theme: 'dark' });
  assert.ok(runs(document).every(run => run.font.size === 16));
});

test('failed native math keeps the original source and adjacent body text with an explicit warning', () => {
  const source = 'Before $\\unknownSlideSCI{x}$ after.';
  const document = buildContentDocument({ kind: 'markdown', source });
  assert.equal(runs(document).map(run => run.text).join(''), source);
  assert.ok(document.warnings.some(warning => /保留公式源码和正文/.test(warning.message)));
  assert.ok(runs(document).every(run => !run.omml));
});

test('PPTX fragments keep native tables, native math, hyperlinks and valid namespaces', async () => {
  const result = await contentHandlers.render({ kind: 'markdown', source: '# Heading\n\n[x](https://example.com) $x^2$\n\n| A | B |\n|---|---|\n| $\\sqrt{x}$ | **bold** |\n\n> quote\n\n$$\\frac{1}{2}$$', left: 90, top: 65, width: 400 });
  const zip = await JSZip.loadAsync(result.fragment.pptx, { base64: true });
  const markup = await zip.file('ppt/slides/slide1.xml').async('string');
  const document = new DOMParser().parseFromString(markup, 'text/xml');
  assert.ok(document.getElementsByTagNameNS(MATH, 'oMath').length >= 3);
  assert.ok(document.getElementsByTagNameNS(DRAWING, 'tbl').length);
  assert.ok(document.getElementsByTagNameNS('http://schemas.microsoft.com/office/drawing/2010/main', 'm').length);
  const relationships = await zip.file('ppt/slides/_rels/slide1.xml.rels').async('string');
  assert.match(relationships, /https:\/\/example.com/);
  assert.equal(result.fragment.bounds.left, 90);
  assert.equal(result.fragment.bounds.top, 65);
  assert.equal(result.editable, true);
  assert.equal(Object.entries(zip.files).filter(([path, part]) => !part.dir && /ppt\/media\//.test(path)).length, 0);
  const parsed = await parsePresentation(result.fragment.pptx);
  assert.ok(parsed.slides[0].shapes.some(shape => shape.type === 'table'));
  const equation = parsed.slides[0].shapes.find(shape => shape.name === 'SlideSCI editable equation');
  assert.ok(equation, 'Native equations must remain selectable by the host adapter');
  assert.match(equation.alternativeText, /\\frac\{1\}\{2\}/);
  const captured = await extractFragment(result.fragment.pptx, { ids: [equation.id] });
  const capturedZip = await JSZip.loadAsync(captured.pptx, { base64: true });
  const capturedDocument = new DOMParser().parseFromString(await capturedZip.file('ppt/slides/slide1.xml').async('string'), 'text/xml');
  assert.ok(capturedDocument.getElementsByTagNameNS(MATH, 'oMath').length);
  assert.equal(capturedDocument.getElementsByTagNameNS(DRAWING, 'tbl').length, 0);
});

test('MathJax SVG keeps source alt text, vector data and an embedded compatibility preview', async () => {
  const source = '\\frac{a}{b}';
  const result = await contentHandlers.render({ kind: 'latex-svg', source });
  const zip = await JSZip.loadAsync(result.fragment.pptx, { base64: true });
  const markup = await zip.file('ppt/slides/slide1.xml').async('string');
  const document = new DOMParser().parseFromString(markup, 'text/xml');
  assert.equal(document.getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'cNvPr')[1].getAttribute('descr'), source);
  assert.ok(Object.keys(zip.files).some(path => /\.svg$/.test(path)));
  assert.ok(Object.keys(zip.files).some(path => /\.png$/.test(path)));
  assert.equal(result.editable, false);
  assert.throws(() => latexToSvg('\\unknownSlideSCI{x}'));
});

test('raw SVG permits local paint references and rejects external resources', () => {
  const svg = '<svg viewBox="0 0 100 50"><defs><linearGradient id="g"><stop stop-color="red"/></linearGradient></defs><rect width="100" height="50" fill="url(\'#g\')"/></svg>';
  assert.match(readSvg(svg).source, /xmlns="http:\/\/www.w3.org\/2000\/svg"/);
  assert.equal(readSvg(svg).width, 75);
  assert.throws(() => readSvg('<svg><image href="https://example.com/a.png"/></svg>'), /外部资源/);
});

test('tall native matrices leave following paragraphs below their bounds', () => {
  const document = buildContentDocument({ kind: 'markdown', source: '$$\\begin{pmatrix}a\\\\b\\\\c\\\\d\\\\e\\end{pmatrix}$$\n\nFollowing paragraph.', fontSize: 18 });
  const [equation, following] = document.shapes.filter(shape => shape.kind === 'text');
  assert.ok(equation.height > 90);
  assert.ok(following.top >= equation.top + equation.height);
  assert.deepEqual(document.warnings, []);
});

test('inline code keeps entity source and math delimiters literal', () => {
  const document = buildContentDocument({ kind: 'markdown', source: '`&lt; $x$` and $x$' });
  const values = runs(document);
  assert.ok(values.some(run => run.text === '&lt; $x$' && run.font.name === 'Menlo' && !run.omml));
  assert.equal(values.filter(run => run.omml).length, 1);
});

test('cached layout templates preserve new content and refresh creation timestamps', async t => {
  t.mock.timers.enable({ apis: ['Date'], now: Date.UTC(2030, 0, 1) });
  const first = await contentHandlers.render({ kind: 'markdown', source: 'First render' });
  t.mock.timers.tick(10000);
  const second = await contentHandlers.render({ kind: 'markdown', source: 'Second render' });
  const firstZip = await JSZip.loadAsync(first.fragment.pptx, { base64: true });
  const secondZip = await JSZip.loadAsync(second.fragment.pptx, { base64: true });
  assert.match(await firstZip.file('docProps/core.xml').async('string'), /2030-01-01T00:00:00/);
  assert.match(await secondZip.file('docProps/core.xml').async('string'), /2030-01-01T00:00:10/);
  const markup = await secondZip.file('ppt/slides/slide1.xml').async('string');
  assert.match(markup, /Second render/);
  assert.doesNotMatch(markup, /First render/);
});
