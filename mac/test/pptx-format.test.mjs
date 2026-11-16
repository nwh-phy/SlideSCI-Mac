import test from 'node:test';
import assert from 'node:assert/strict';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { applyTextFormatting, textLayout } from '../lib/pptx-format.mjs';

const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const M = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const serialize = node => new XMLSerializer().serializeToString(node);
const nodes = (root, ns, name) => Array.from(root.getElementsByTagNameNS(ns, name));
const child = (root, ns, name) => Array.from(root.childNodes).find(node => node.namespaceURI === ns && node.localName === name);
const properties = '<a:rPr sz="2400" b="0" u="sng"><a:solidFill><a:srgbClr val="123456"/></a:solidFill><a:latin typeface="Arial"/><a:hlinkClick r:id="targetLink"/></a:rPr>';
function slide(body) {
  return new DOMParser().parseFromString(`<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="${A}" xmlns:r="${R}" xmlns:m="${M}" xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="Text"/></p:nvSpPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p>${body}</a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`, 'application/xml');
}
const run = (text, pr = properties) => `<a:r>${pr}<a:t>${text}</a:t></a:r>`;
const raw = body => `<a:rPr xmlns:a="${A}" xmlns:r="${R}" sz="1200" b="1">${body}<a:latin typeface="Courier New"/><a:hlinkClick r:id="sourceLink"/></a:rPr>`;

test('effect copy preserves exact native shadow/reflection, target typography and target links', () => {
  const doc = slide(run('target'));
  const reflection = `<a:reflection xmlns:a="${A}" blurRad="9234" stA="41000" endA="0" endPos="73000" dist="17321" dir="5400000" fadeDir="5400000" sx="100000" sy="-100000" algn="bl"/>`;
  const source = raw('<a:effectLst><a:innerShdw blurRad="2345" dist="6789" dir="123456"><a:srgbClr val="ABCDEF"/></a:innerShdw>' + reflection + '</a:effectLst>');
  applyTextFormatting(doc, '2', { fontScope: 'effect', font: { raw: source, shadow: true, reflection: { type: 1, raw: reflection } } });
  const pr = child(nodes(doc, A, 'r')[0], A, 'rPr');
  assert.equal(pr.getAttribute('sz'), '2400');
  assert.equal(pr.getAttribute('b'), '0');
  assert.equal(child(pr, A, 'latin').getAttribute('typeface'), 'Arial');
  assert.equal(nodes(pr, A, 'srgbClr')[0].getAttribute('val'), '123456');
  assert.equal(nodes(pr, A, 'innerShdw').length, 1);
  assert.equal(nodes(pr, A, 'innerShdw')[0].getAttribute('dist'), '6789');
  assert.equal(nodes(pr, A, 'outerShdw').length, 0);
  assert.equal(nodes(pr, A, 'reflection').length, 1);
  assert.equal(nodes(pr, A, 'reflection')[0].getAttribute('dist'), '17321');
  assert.equal(nodes(pr, A, 'reflection')[0].getAttribute('endPos'), '73000');
  assert.equal(nodes(pr, A, 'hlinkClick')[0].getAttributeNS(R, 'id'), 'targetLink');
  assert.ok(!serialize(doc).includes('sourceLink'));
});

test('name-only raw copy preserves size, colors, underline and target hyperlink', () => {
  const doc = slide(run('target'));
  applyTextFormatting(doc, '2', { fontScope: 'name', font: { raw: raw('<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>') } });
  const pr = child(nodes(doc, A, 'r')[0], A, 'rPr');
  assert.equal(child(pr, A, 'latin').getAttribute('typeface'), 'Courier New');
  assert.equal(pr.getAttribute('sz'), '2400');
  assert.equal(pr.getAttribute('b'), '0');
  assert.equal(pr.getAttribute('u'), 'sng');
  assert.equal(nodes(pr, A, 'srgbClr')[0].getAttribute('val'), '123456');
  assert.equal(nodes(pr, A, 'hlinkClick')[0].getAttributeNS(R, 'id'), 'targetLink');
});

test('UTF-16 range styling isolates a complete emoji without changing adjacent characters', () => {
  const doc = slide(run('A😀BC'));
  applyTextFormatting(doc, '2', { fontScope: 'color', textRange: { start: 1, length: 2, font: { color: '#FF0000' } } });
  const runs = nodes(doc, A, 'r');
  assert.deepEqual(runs.map(node => child(node, A, 't').textContent), ['A', '😀', 'BC']);
  assert.deepEqual(runs.map(node => nodes(node, A, 'srgbClr')[0].getAttribute('val')), ['123456', 'FF0000', '123456']);
  assert.ok(runs.every(node => child(node, A, 'rPr').getAttribute('sz') === '2400'));
  assert.ok(runs.every(node => nodes(node, A, 'hlinkClick')[0].getAttributeNS(R, 'id') === 'targetLink'));
});

const equation = '<a14:m><m:oMath><m:f><m:num><m:r><m:rPr><m:sty m:val="i"/></m:rPr><a:rPr sz="1200"><a:latin typeface="Cambria Math"/></a:rPr><m:t>ab</m:t></m:r></m:num><m:den><m:r><a:rPr sz="1200"/><m:t>x</m:t></m:r></m:den></m:f></m:oMath></a14:m>';
test('whole-shape size formatting scales native equation runs and keeps equation structure', () => {
  const doc = slide(run('!') + equation + run('?'));
  applyTextFormatting(doc, '2', { fontScope: 'size', font: { size: 28 } });
  assert.equal(nodes(doc, M, 'f').length, 1);
  assert.equal(nodes(doc, M, 'num')[0].textContent, 'ab');
  assert.equal(nodes(doc, M, 'den')[0].textContent, 'x');
  assert.equal(nodes(doc, M, 'sty')[0].getAttributeNS(M, 'val'), 'i');
  assert.ok(nodes(doc, M, 'r').every(node => child(node, A, 'rPr').getAttribute('sz') === '2800'));
  assert.equal(textLayout(nodes(doc, 'http://schemas.openxmlformats.org/presentationml/2006/main', 'txBody')[0]).text, '!abx?');
});

test('partial native equation range preserves the fraction and styles only the selected math character', () => {
  const doc = slide(run('!') + equation + run('?'));
  applyTextFormatting(doc, '2', { fontScope: 'color', textRange: { start: 2, length: 1, font: { color: '#FF0000' } } });
  assert.equal(nodes(doc, M, 'f').length, 1);
  assert.equal(nodes(doc, M, 'num')[0].textContent, 'ab');
  const mathRuns = nodes(doc, M, 'r');
  assert.deepEqual(mathRuns.map(node => child(node, M, 't').textContent), ['a', 'b', 'x']);
  assert.equal(nodes(mathRuns[0], A, 'solidFill').length, 0);
  assert.equal(nodes(mathRuns[1], A, 'srgbClr')[0].getAttribute('val'), 'FF0000');
  assert.equal(nodes(mathRuns[2], A, 'solidFill').length, 0);
  assert.equal(nodes(doc, M, 'den')[0].textContent, 'x');
});

test('whole-field formatting preserves its dynamic field identity and text', () => {
  const doc = slide(`<a:fld id="{12345678-1234-1234-1234-123456789ABC}" type="slidenum">${properties}<a:t>7</a:t></a:fld>`);
  applyTextFormatting(doc, '2', { fontScope: 'size', font: { size: 32 } });
  const fields = nodes(doc, A, 'fld');
  assert.equal(fields.length, 1);
  assert.equal(fields[0].getAttribute('type'), 'slidenum');
  assert.equal(fields[0].getAttribute('id'), '{12345678-1234-1234-1234-123456789ABC}');
  assert.equal(child(fields[0], A, 't').textContent, '7');
  assert.equal(child(fields[0], A, 'rPr').getAttribute('sz'), '3200');
});

test('native gradient paint is preserved when a host reports its first stop as foreground color', () => {
  const doc = slide(run('gradient'));
  const gradient = '<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:lin ang="5400000" scaled="1"/></a:gradFill>';
  applyTextFormatting(doc, '2', { fontScope: 'color', font: { raw: raw(gradient), color: '#FF0000' } });
  const pr = child(nodes(doc, A, 'r')[0], A, 'rPr');
  assert.equal(nodes(pr, A, 'gradFill').length, 1);
  assert.equal(nodes(pr, A, 'solidFill').length, 0);
  assert.deepEqual(nodes(pr, A, 'gs').map(node => [node.getAttribute('pos'), nodes(node, A, 'srgbClr')[0].getAttribute('val')]),
    [['0', 'FF0000'], ['100000', '0000FF']]);
  assert.equal(pr.getAttribute('sz'), '2400');
  assert.equal(nodes(pr, A, 'hlinkClick')[0].getAttributeNS(R, 'id'), 'targetLink');
});

test('explicit effect removal clears nested effects while retaining unrelated soft edges', () => {
  const effects = '<a:rPr><a:effectLst><a:glow rad="12000"><a:srgbClr val="FF0000"/></a:glow><a:innerShdw blurRad="1000" dist="1000" dir="0"><a:srgbClr val="000000"/></a:innerShdw><a:outerShdw blurRad="1000" dist="1000" dir="0"><a:srgbClr val="000000"/></a:outerShdw><a:reflection blurRad="0"/><a:softEdge rad="2000"/></a:effectLst></a:rPr>';
  const doc = slide(run('clear', effects));
  applyTextFormatting(doc, '2', { font: { glow: false, shadow: false, reflection: false } });
  const pr = child(nodes(doc, A, 'r')[0], A, 'rPr');
  for (const name of ['glow', 'innerShdw', 'outerShdw', 'reflection']) assert.equal(nodes(pr, A, name).length, 0);
  assert.equal(nodes(pr, A, 'softEdge')[0].getAttribute('rad'), '2000');
});
