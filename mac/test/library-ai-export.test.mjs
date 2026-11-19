import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import JSZip from 'jszip';
import { libraryHandlers } from '../lib/library.mjs';
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

test('export parses page ranges, deduplicates overlaps and rejects unavailable pages', () => {
  const slides = Array.from({ length: 6 }, (_, index) => ({ index }));
  const snapshot = { slides, slideIndex: 2, selectedSlideIndices: [3, 1, 3] };
  assert.deepEqual(pageIndices('custom', snapshot, '1,3-5,4'), [0, 2, 3, 4]);
  assert.deepEqual(pageIndices('selected', snapshot), [1, 3]);
  assert.deepEqual(pageIndices('current', snapshot), [2]);
  assert.throws(() => pageIndices('custom', snapshot, '5-3'), /范围/u);
  assert.throws(() => pageIndices('custom', snapshot, '0,9'), /1到6/u);
});
