import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, readFile, writeFile, rm, readdir } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import sharp from 'sharp';
import { imageHandlers } from '../lib/images.mjs';

const denied = { files: { read: async () => { throw new Error('请先通过文件选择器选择该文件。'); } } };

test('negative crop (picture shrunk inside its frame) renders with transparent padding instead of failing', async () => {
  const png = await sharp({ create: { width: 100, height: 50, channels: 3, background: '#ff0000' } }).png().toBuffer();
  const rendered = await imageHandlers.crop({ data: png.toString('base64'), crop: { left: -0.1, top: 0, right: -0.1, bottom: 0 }, region: { x: 0, y: 0, width: 1, height: 1 } }, denied);
  const raw = await sharp(Buffer.from(rendered.data, 'base64')).ensureAlpha().raw().toBuffer({ resolveWithObject: true });
  assert.equal(raw.info.width, 120);
  assert.equal(raw.data[3], 0, 'the padded left edge is transparent');
  const center = (Math.floor(raw.info.height / 2) * raw.info.width + 60) * 4;
  assert.deepEqual([...raw.data.subarray(center, center + 4)], [255, 0, 0, 255]);
});

test('image settings survive concurrent writes and a damaged settings file', async () => {
  const storageDir = await mkdtemp(path.join(os.tmpdir(), 'slidesci-settings-'));
  try {
    await writeFile(path.join(storageDir, 'images-settings.json'), '{"scale":');
    const env = { storageDir };
    await Promise.all([
      imageHandlers.settings({ values: { pseudoColor: '#00ff00' } }, env),
      imageHandlers.settings({ values: { scale: { fontSize: 18 } } }, env),
      imageHandlers.settings({ values: { zoom: { color: '#ff0000' } } }, env)
    ]);
    const saved = JSON.parse(await readFile(path.join(storageDir, 'images-settings.json'), 'utf8'));
    assert.deepEqual(saved, { pseudoColor: '#00ff00', scale: { fontSize: 18 }, zoom: { color: '#ff0000' } });
    assert.equal((await readdir(storageDir)).filter(name => name.includes('.corrupt-')).length, 1, 'the damaged file is kept aside');
  } finally { await rm(storageDir, { recursive: true, force: true }); }
});

test('TIFF calibration reads a recorded source path without a fresh grant', async () => {
  const directory = await mkdtemp(path.join(os.tmpdir(), 'slidesci-fov-'));
  try {
    const file = path.join(directory, 'calibrated.tif');
    // Minimal little-endian TIFF directory: width 40, height 20 and an ImageJ description.
    const description = Buffer.from('ImageJ=1.54\nunit=micron\npixel_width=0.5\npixel_height=0.5\n\0');
    const tiff = Buffer.alloc(8 + 2 + 3 * 12 + 4 + description.length);
    tiff.write('II'); tiff.writeUInt16LE(42, 2); tiff.writeUInt32LE(8, 4); tiff.writeUInt16LE(3, 8);
    const entry = (index, tag, type, count, value) => { const at = 10 + index * 12; tiff.writeUInt16LE(tag, at); tiff.writeUInt16LE(type, at + 2); tiff.writeUInt32LE(count, at + 4); tiff.writeUInt32LE(value, at + 8); };
    entry(0, 256, 4, 1, 40); entry(1, 257, 4, 1, 20); entry(2, 270, 2, description.length, 8 + 2 + 3 * 12 + 4);
    description.copy(tiff, 8 + 2 + 3 * 12 + 4);
    await writeFile(file, tiff);
    const fov = await imageHandlers.fov({ path: file }, denied);
    assert.equal(fov.width, 20);
    assert.equal(fov.height, 10);
    await assert.rejects(imageHandlers.fov({ path: path.join(directory, 'moved.tif') }, denied), /原始文件已不存在或已移动/);
  } finally { await rm(directory, { recursive: true, force: true }); }
});
