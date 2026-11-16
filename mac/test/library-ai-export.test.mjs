import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import JSZip from 'jszip';
import { libraryHandlers } from '../lib/library.mjs';
import { aiHandlers, decodeCompletion, endpoint } from '../lib/ai.mjs';
import { validateNativeCommands } from '../lib/ai-commands.mjs';
import { pageIndices } from '../src/features/export.js';

const preview = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl6kGQAAAAASUVORK5CYII=';
async function fragment() {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>');
  zip.file('ppt/presentation.xml', '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>');
  zip.file('ppt/slides/slide1.xml', '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:spTree/></p:sld>');
  zip.file('ppt/media/resource.bin', Buffer.from([0, 255, 17, 38, 127]));
  return { pptx: await zip.generateAsync({ type: 'base64' }), preview, shapeIds: ['18', '24'], bounds: { left: 10, top: 20, width: 50, height: 60 } };
}
async function environment(t) {
  const storageDir = await mkdtemp(path.join(os.tmpdir(), 'slidesci-library-ai-test-'));
  t.after(() => rm(storageDir, { recursive: true, force: true }));
  const keys = new Map();
  return { storageDir, native: { keychainGet: async ({ account }) => keys.get(account) || null, keychainSet: async ({ account, secret }) => keys.set(account, secret), keychainDelete: async ({ account }) => keys.delete(account) }, files: {} };
}

test('library archives preserve nested folders, editable package bytes, preview, and fragment bounds', async t => {
  const env = await environment(t); const other = await environment(t);
  const parent = await libraryHandlers.mkdir({ name: '图表' }, env);
  const child = await libraryHandlers.mkdir({ name: '可编辑', folder: parent.id }, env);
  const source = await fragment();
  await libraryHandlers.save({ name: '显微图组合', folder: child.id, fragment: source }, env);
  const archive = await libraryHandlers.export({}, env);
  assert.deepEqual(await libraryHandlers.import({ data: archive.data }, other), { imported: 1, skipped: 0 });
  const listing = await libraryHandlers.list({ query: '显微图' }, other);
  assert.equal(listing.items[0].path, '图表/可编辑');
  const stored = await libraryHandlers.get({ id: listing.items[0].id }, other);
  assert.deepEqual(Buffer.from(stored.fragment.pptx, 'base64'), Buffer.from(source.pptx, 'base64'));
  assert.deepEqual(stored.fragment.bounds, source.bounds);
  assert.equal(stored.fragment.preview, preview);
});

test('library rejects traversal, cycles, duplicate races, and invalid updates without losing existing data', async t => {
  const env = await environment(t); const source = await fragment();
  const parent = await libraryHandlers.mkdir({ name: 'parent' }, env);
  const child = await libraryHandlers.mkdir({ name: 'child', folder: parent.id }, env);
  await assert.rejects(libraryHandlers.move({ id: parent.id, folder: child.id }, env), /自身|子文件夹/u);
  await assert.rejects(libraryHandlers.mkdir({ name: '../escape' }, env), /路径分隔符/u);
  const attempts = await Promise.allSettled([libraryHandlers.save({ name: 'same', fragment: source }, env), libraryHandlers.save({ name: 'same', fragment: source }, env)]);
  assert.equal(attempts.filter(item => item.status === 'fulfilled').length, 1);
  const asset = attempts.find(item => item.status === 'fulfilled').value.asset;
  await assert.rejects(libraryHandlers.save({ id: asset.id, name: asset.name, fragment: { ...source, pptx: Buffer.from('bad archive').toString('base64') } }, env));
  assert.equal((await libraryHandlers.get({ id: asset.id }, env)).fragment.pptx, source.pptx);
  const unsafe = new JSZip(); unsafe.file('../outside.pptx', Buffer.from(source.pptx, 'base64')); unsafe.file('../outside.png', Buffer.from(preview, 'base64'));
  await assert.rejects(libraryHandlers.import({ data: await unsafe.generateAsync({ type: 'base64' }) }, env));
  assert.equal((await libraryHandlers.get({ id: asset.id }, env)).fragment.pptx, source.pptx);
});

test('library imports Windows ZIP layout and applies explicit conflict policy', async t => {
  const env = await environment(t); const source = await fragment(); const zip = new JSZip();
  zip.file('Windows素材/组合.pptx', Buffer.from(source.pptx, 'base64')); zip.file('Windows素材/组合.png', Buffer.from(preview, 'base64'));
  const data = await zip.generateAsync({ type: 'base64' });
  assert.equal((await libraryHandlers.import({ data }, env)).imported, 1);
  assert.equal((await libraryHandlers.import({ data, conflict: 'skip' }, env)).skipped, 1);
  assert.equal((await libraryHandlers.import({ data, conflict: 'rename' }, env)).imported, 1);
  assert.equal((await libraryHandlers.list({ query: '组合' }, env)).items.length, 2);
});

const context = { slideIndex: 0, shapes: [{ id: 'one' }], slides: [{ index: 0, shapes: [{ id: 'one' }] }, { index: 1, shapes: [{ id: 'two' }] }] };
test('AI validates native target identities and properties across slides and sequential deletion', () => {
  const commands = [{ op: 'update', slideIndex: 1, id: 'two', properties: { font: { color: '#123456' } } }];
  assert.deepEqual(validateNativeCommands(commands, context), commands);
  assert.throws(() => validateNativeCommands([{ op: 'update', id: 'two', properties: { text: 'x' } }], context), /不存在/u);
  assert.throws(() => validateNativeCommands([{ op: 'update', id: 'one', properties: { width: -2 } }], context), /大于0/u);
  assert.throws(() => validateNativeCommands([{ op: 'delete', ids: ['one'] }, { op: 'update', id: 'one', properties: { text: 'stale' } }], context), /不存在/u);
  assert.throws(() => validateNativeCommands([{ op: 'deleteSlides', indices: [0, 1] }], context), /至少一页/u);
  assert.throws(() => validateNativeCommands([{ op: 'update', id: 'one', properties: { javascript: 'arbitrary code' } }], context), /格式/u);
  assert.doesNotThrow(() => validateNativeCommands([{ op: 'moveSlide', index: 0, toIndex: 1 }, { op: 'update', id: 'one', slideIndex: 1, properties: { text: 'moved' } }], context));
});

function sseResponse(frames, chunkSize = 3) {
  const bytes = new TextEncoder().encode(frames.join('\r\n\r\n') + '\r\n\r\n');
  return new Response(new ReadableStream({ start(controller) { for (let index = 0; index < bytes.length; index += chunkSize) controller.enqueue(bytes.slice(index, index + chunkSize)); controller.close(); } }), { headers: { 'Content-Type': 'text/event-stream' } });
}
test('AI assembles UTF8 SSE chunks and tool arguments, and exposes real content deltas', async () => {
  const frames = [
    'data: ' + JSON.stringify({ choices: [{ delta: { content: '处理中文' } }] }),
    'data: ' + JSON.stringify({ choices: [{ delta: { tool_calls: [{ index: 0, id: 'call_1', function: { name: 'edit_presentation', arguments: '{"commands":[' } }] } }] }),
    'data: ' + JSON.stringify({ choices: [{ delta: { tool_calls: [{ index: 0, function: { arguments: '{"op":"delete","ids":["one"]}]}' } }] } }] }),
    'data: ' + JSON.stringify({ choices: [{ delta: {}, finish_reason: 'tool_calls' }] }),
    'data: [DONE]'
  ];
  const job = { events: [] }; const message = await decodeCompletion(sseResponse(frames), job);
  assert.equal(message.content, '处理中文'); assert.equal(message.tool_calls[0].id, 'call_1');
  assert.deepEqual(JSON.parse(message.tool_calls[0].function.arguments), { commands: [{ op: 'delete', ids: ['one'] }] });
  assert.equal(job.events[0].text, '处理中文');
  await assert.rejects(decodeCompletion(sseResponse([frames[0]]), { events: [] }), /中断/u);
  await assert.rejects(decodeCompletion(new Response('{"error":{"message":"rate limited"}}', { status: 429 }), { events: [] }), /429.*rate limited/u);
});

test('AI profiles persist nonsecrets while keys are retained only by native keychain API', async t => {
  const env = await environment(t);
  const saved = await aiHandlers.saveProfile({ id: 'test', name: 'Test', baseUrl: 'http://127.0.0.1:12345/v1', model: 'fixture', secret: 'unit-test-secret', systemPrompt: 'test prompt' }, env);
  assert.equal(saved.activeProfileId, 'test'); assert.equal(saved.profiles.find(item => item.id === 'test').hasKey, true);
  const persisted = await readFile(path.join(env.storageDir, 'ai', 'profiles.json'), 'utf8');
  assert.equal(persisted.includes('unit-test-secret'), false);
  assert.equal(JSON.stringify(saved).includes('unit-test-secret'), false);
  assert.equal(endpoint('https://api.example.com/v1/'), 'https://api.example.com/v1/chat/completions');
  assert.throws(() => endpoint('http://untrusted.example/v1'), /HTTPS/u);
});

test('export parses page ranges, deduplicates overlaps and rejects unavailable pages', () => {
  const slides = Array.from({ length: 6 }, (_, index) => ({ index }));
  const snapshot = { slides, slideIndex: 2, selectedSlideIndices: [3, 1, 3] };
  assert.deepEqual(pageIndices('custom', snapshot, '1,3-5,4'), [0, 2, 3, 4]);
  assert.deepEqual(pageIndices('selected', snapshot), [1, 3]);
  assert.deepEqual(pageIndices('current', snapshot), [2]);
  assert.throws(() => pageIndices('custom', snapshot, '5-3'), /范围/u);
  assert.throws(() => pageIndices('custom', snapshot, '0,9'), /1到6/u);
});

test('AI streaming preserves UTF8 and split CRLF delimiters at every byte boundary', async () => {
  const frames = [
    'data: ' + JSON.stringify({ choices: [{ delta: { content: '中文第一行\n第二行' } }] }),
    'data: ' + JSON.stringify({ choices: [{ delta: {}, finish_reason: 'stop' }] }),
    'data: [DONE]'
  ];
  for (const size of [1, 2, 7, 1000]) {
    const message = await decodeCompletion(sseResponse(frames, size), { events: [] });
    assert.equal(message.content, '中文第一行\n第二行');
  }
});
