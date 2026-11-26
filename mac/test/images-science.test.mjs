import test from 'node:test';
import assert from 'node:assert/strict';
import sharp from 'sharp';
import { imageHandlers, autoDisplayLevels, display16, firstTiffTags, tiffFov, readImageDimensions } from '../lib/images.mjs';
import { scaleLayout, defaultScale, localBounds, markerBounds, relativeRegion, visibleFov, FOV_PREFIX, SCALE_PREFIX, near, clipLine, lineBounds, serializeTag, parse, imagesFor, formatAmount } from '../src/features/images-science-core.js';
import { applyPseudoColors, actions as imageActions } from '../src/features/images.js';
import { applyZoomEntries, reconcileSlide } from '../src/features/science.js';

// Small, independently encoded uncompressed TIFF fixtures; no user files or Office documents.
function tiff16(samples, channels, little = true, associated = false) {
  const width = samples.length / channels, height = 1;
  const entries = [[256, 4, [width]], [257, 4, [height]], [258, 3, Array(channels).fill(16)], [259, 3, [1]], [262, 3, [channels === 1 ? 1 : 2]], [273, 4, [0]], [277, 3, [channels]], [278, 4, [height]], [279, 4, [samples.length * 2]], [284, 3, [1]], ...(channels === 4 ? [[338, 3, [associated ? 1 : 2]]] : [])];
  const directoryEnd = 8 + 2 + entries.length * 12 + 4, bitsSize = channels > 2 ? channels * 2 : 0, pixelOffset = directoryEnd + bitsSize;
  entries.find(e => e[0] === 273)[2][0] = pixelOffset;
  const bytes = Buffer.alloc(pixelOffset + samples.length * 2);
  const short = (value, offset) => little ? bytes.writeUInt16LE(value, offset) : bytes.writeUInt16BE(value, offset);
  const long = (value, offset) => little ? bytes.writeUInt32LE(value, offset) : bytes.writeUInt32BE(value, offset);
  bytes.write(little ? 'II' : 'MM'); short(42, 2); long(8, 4); short(entries.length, 8);
  entries.forEach(([tag, type, values], i) => { const offset = 10 + i * 12; short(tag, offset); short(type, offset + 2); long(values.length, offset + 4); if (type === 3 && values.length > 2) { long(directoryEnd, offset + 8); values.forEach((v, j) => short(v, directoryEnd + j * 2)); } else if (type === 3) values.forEach((v, j) => short(v, offset + 8 + j * 2)); else long(values[0], offset + 8); });
  samples.forEach((value, i) => short(value, pixelOffset + i * 2));
  return bytes;
}
const decode = result => sharp(Buffer.from(result.data, 'base64')).ensureAlpha().raw().toBuffer();
const png = async (data, width, height) => ({ data: (await sharp(Buffer.from(data), { raw: { width, height, channels: 4 } }).png().toBuffer()).toString('base64') });

for (const little of [true, false]) test(`16-bit grayscale TIFF ${little ? 'little' : 'big'} endian preserves measured contrast`, async () => {
  const result = await imageHandlers.normalize({ data: tiff16([1000, 1500, 2000], 1, little).toString('base64'), name: 'own-fixture.tif' }, {});
  assert.equal(result.highBitDepth, true); assert.equal(result.width, 3);
  assert.deepEqual([...await decode(result)], [0, 0, 0, 255, 128, 128, 128, 255, 255, 255, 255, 255]);
});
test('16-bit RGB uses a shared window rather than per-channel contrast', async () => {
  const result = await imageHandlers.normalize({ data: tiff16([1000, 2000, 3000, 2000, 3000, 4000], 3).toString('base64') }, {});
  assert.deepEqual([...await decode(result)], [0, 85, 171, 255, 85, 171, 255, 255]);
});
test('16-bit premultiplied RGBA unpremultiplies before contrast and ignores hidden colors', async () => {
  const result = await imageHandlers.normalize({ data: tiff16([1000, 2000, 3000, 32768, 60000, 60000, 60000, 0], 4, false, true).toString('base64') }, {});
  const pixels = [...await decode(result)]; assert.deepEqual(pixels.slice(0, 4), [0, 128, 255, 128]); assert.equal(pixels[7], 0);
});
test('ImageJ first-auto skips histogram spikes rather than percentile clipping', () => {
  const samples = new Uint16Array(10000); samples.fill(1000); for (let i = 0; i < 1000; i++) samples[i] = 10000 + Math.floor(i / 10) * 100;
  const [minimum, maximum] = autoDisplayLevels(samples, 1); assert.ok(minimum > 1000); assert.ok(maximum > minimum && maximum <= 19900);
  assert.deepEqual([...display16(new Uint16Array([1234, 1234]), 1)], [0, 0, 0, 255, 0, 0, 0, 255]);
});
test('OME and ImageJ physical calibration ignore Z spacing and respect unit conversion', () => {
  const base = { 256: [100], 257: [50] };
  const ome = tiffFov({ ...base, 270: '<OME><Image><Pixels SizeX="100" SizeY="50" PhysicalSizeX="250" PhysicalSizeXUnit="nm" PhysicalSizeY="0.5" PhysicalSizeYUnit="μm"/></Image></OME>' });
  assert.equal(ome.width, 25); assert.equal(ome.height, 25);
  const imageJ = tiffFov({ ...base, 270: 'ImageJ=1.54\nunit=um\nspacing=99', 282: [2], 283: [4] }); assert.equal(imageJ.width, 50); assert.equal(imageJ.height, 12.5);
  assert.equal(tiffFov({ ...base, 270: 'ImageJ=1.54\nunit=pixel\nspacing=99', 282: [2], 283: [4] }), null);
});
test('BigTIFF first directory metadata supports 64-bit offsets', () => {
  const bytes = Buffer.alloc(72); bytes.write('II'); bytes.writeUInt16LE(43, 2); bytes.writeUInt16LE(8, 4); bytes.writeBigUInt64LE(16n, 8); bytes.writeBigUInt64LE(2n, 16);
  for (let i = 0; i < 2; i++) { const offset = 24 + i * 20; bytes.writeUInt16LE(256 + i, offset); bytes.writeUInt16LE(16, offset + 2); bytes.writeBigUInt64LE(1n, offset + 4); bytes.writeBigUInt64LE(BigInt(i ? 50 : 100), offset + 12); }
  assert.deepEqual(firstTiffTags(bytes).tags, { 256: [100], 257: [50] });
});
test('channel recoloring recovers custom dark LUT intensity and preserves alpha', async () => {
  const original = await png([0, 64, 0, 128], 1, 1), result = await imageHandlers.color({ ...original, color: '#ff0000', sourcePeak: 128 }, {});
  assert.deepEqual([...await decode(result)], [128, 0, 0, 128]);
});
test('channel merge saturates RGB and uses alpha contribution on black', async () => {
  const a = await png([240, 20, 0, 255], 1, 1), b = await png([100, 100, 200, 128], 1, 1);
  const result = await imageHandlers.merge({ images: [a, b] }, {}); assert.deepEqual([...await decode(result)], [255, 70, 100, 255]);
});
test('rotated crop reads displayed axes and rejects regions with no overlap', async () => {
  const image = await png([255, 0, 0, 255, 0, 255, 0, 255, 0, 0, 255, 255, 255, 255, 255, 255], 2, 2);
  const result = await imageHandlers.crop({ ...image, region: { x: 0, y: 0, width: 1, height: 1, rotation: 90 } }, {});
  assert.deepEqual([...await decode(result)], [0, 255, 0, 255, 255, 255, 255, 255, 255, 0, 0, 255, 0, 0, 255, 255]);
  await assert.rejects(imageHandlers.crop({ ...image, region: { x: 2, y: 2, width: .5, height: .5, rotation: 0 } }, {}), /没有重叠/);
});
test('native scale geometry uses cropped physical FOV and validates full label fit', () => {
  const image = { width: 200, height: 100, alternativeText: FOV_PREFIX + JSON.stringify({ width: 100, height: 50, unit: 'μm' }), crop: { left: .25, right: .25, top: 0, bottom: 0 } };
  const fov = visibleFov(image), layout = scaleLayout(image, fov, { ...defaultScale, widthMicrometers: 10, fontSize: 10 }, { width: 24, height: 12 });
  assert.equal(fov.width, 50); assert.equal(layout.bar.width, 40); assert.ok(layout.text.left >= 0); assert.ok(layout.text.top >= 0);
  assert.throws(() => scaleLayout(image, fov, { ...defaultScale, widthMicrometers: 100 }, { width: 24, height: 12 }), /超出/);
});
test('legacy tagged derived FOV calibrates the full raster and later crop reduces it', () => {
  // Zoom, merged-channel and retained pseudo-color pictures historically carried SLIDESCI_FOV_VISIBLE.
  // Their recorded FOV is full-raster calibration, so cropping afterwards must still shrink it.
  const image = { width: 200, height: 100, alternativeText: FOV_PREFIX + JSON.stringify({ width: 100, height: 50, unit: 'nm' }),
    tags: { SLIDESCI_FOV_VISIBLE: '1' }, crop: { left: .25, top: 0, right: 0, bottom: 0 } };
  const visible = visibleFov(image);
  assert.ok(Math.abs(visible.width - .075) < 1e-12); assert.ok(Math.abs(visible.height - .05) < 1e-12);
  const layout = scaleLayout(image, visible, { ...defaultScale, widthMicrometers: .01, lengthUnit: 'nm', fontSize: 10 }, { width: 24, height: 12 });
  assert.equal(layout.label, '10nm');
  assert.ok(Math.abs(layout.bar.width - 200 * 10 / 75) < 1e-9);
  // Crop is cumulative rather than double-counted when only part of the raster is removed.
  image.crop = { left: .1, right: .15, top: .1, bottom: .4 };
  assert.ok(Math.abs(visibleFov(image).width - .075) < 1e-12);
  assert.ok(Math.abs(visibleFov(image).height - .025) < 1e-12);
});
test('Windows intersection FOV already accounts for the picture crop', () => {
  const image = { type: 'picture', width: 200, height: 100,
    alternativeText: FOV_PREFIX + JSON.stringify({ Width: 50, Height: 50, Unit: 'μm', Source: '由原图标定和实际取图范围换算' }),
    tags: { SLIDESCI_FOV_VISIBLE: '1', SLIDESCI_FOV_SOURCE: '100;50' }, crop: { left: .5, right: 0, top: 0, bottom: 0 } };
  assert.deepEqual(visibleFov(image), { width: 50, height: 50 });
});
test('scale labels retain small lengths in the chosen display unit without float noise', () => {
  const image = { width: 200, height: 100 }, fov = { width: 100, height: 50 };
  const label = (widthMicrometers, lengthUnit) => scaleLayout(image, fov, { ...defaultScale, widthMicrometers, lengthUnit }, { width: 24, height: 12 }).label;
  assert.equal(label(.004, 'μm'), '0.004μm');
  assert.equal(label(.01, 'nm'), '10nm');
  assert.equal(label(.004, 'nm'), '4nm');
  assert.equal(label(50, 'μm'), '50μm');
  assert.equal(label(.1 + .2, 'μm'), '0.3μm');
  assert.equal(formatAmount(1 / 3), '0.333333');
  assert.equal(formatAmount(.004), '0.004');
  assert.equal(formatAmount(.01), '0.01');
  assert.equal(formatAmount(4e-9), '0.000000004');
  assert.equal(formatAmount(0), '0');
});
for (const grouped of [true, false]) test(`reconciling a ${grouped ? 'grouped' : 'ungrouped'} source repairs only its saved scale label`, async t => {
  const previous = globalThis.document;
  globalThis.document = { createElement: () => ({ getContext: () => ({ measureText: text => ({ width: text.length * 6 }) }) }) };
  t.after(() => { if (previous === undefined) delete globalThis.document; else globalThis.document = previous; });

  const fov = { width: 1, height: .5, unit: 'μm' };
  const settings = { ...defaultScale, widthMicrometers: .004, groupWithImage: grouped };
  const sourceTags = { SLIDESCI_IMAGE_UUID: 'SOURCE' };
  const picture = { id: 'picture', name: 'Source', type: 'picture', left: 40, top: 50, width: 200, height: 100, rotation: 0,
    alternativeText: FOV_PREFIX + JSON.stringify(fov) + '\n' + SCALE_PREFIX + JSON.stringify(settings),
    tags: grouped ? { ...sourceTags, SLIDESCI_SCALE_CONTENT: '1' } : sourceTags };
  const bounds = scaleLayout(picture, fov, settings);
  const oldText = { ...bounds.text, left: bounds.text.left + 12, width: bounds.text.width - 12 };
  const bar = { id: 'bar', type: 'shape', ...localBounds(picture, bounds.bar, 0),
    tags: { SLIDESCI_SCALE_OWNER: 'SOURCE', SLIDESCI_SCALE_PART: 'bar' } };
  const label = { id: 'label', name: '比例尺文字', type: 'text', text: '0μm', ...localBounds(picture, oldText, 0),
    tags: { SLIDESCI_SCALE_OWNER: 'SOURCE', SLIDESCI_SCALE_PART: 'text' } };
  const unrelated = { ...label, id: 'unrelated', left: label.left - 40, tags: { SLIDESCI_SCALE_PART: 'text' } };
  const shapes = [picture, bar, label, unrelated];
  const state = { width: 960, height: 540, shapes: grouped ? [{ ...picture, id: 'group', type: 'group', children: shapes,
    tags: { ...sourceTags, SLIDESCI_SCALE_GROUP: '1', SLIDESCI_SCALE_CALIBRATION: serializeTag({ width: 1, height: .5 }) } }] : shapes };
  const updated = [];
  const ctx = { host: { kind: 'powerpoint', async context() { return state; }, async apply(commands) {
    for (const command of commands) {
      assert.equal(command.op, 'update');
      Object.assign(shapes.find(shape => shape.id === command.id), command.properties);
      updated.push(command.id);
    }
    return state;
  } } };
  await reconcileSlide(ctx);
  assert.equal(label.text, '0.004μm');
  assert.equal(label.width, bounds.text.width);
  assert.equal(unrelated.text, '0μm');
  assert.deepEqual(updated, ['label']);
});
test('scientific settings and source paths survive native tag uppercasing', () => {
  const record = { role: 'marker', placement: 'left', name: '放大图 α', source: '/Users/Researcher/Images/Cell.tif' };
  assert.deepEqual(parse(serializeTag(record).toUpperCase()), record);
  assert.equal(parse(serializeTag(record.source).toUpperCase()), record.source);
});
test('region anchors round-trip rotated native markers and normalize Office rotations', () => {
  const source = { left: 32, top: 48, width: 300, height: 180, rotation: 37 }, region = { x: .1, y: .25, width: .3, height: .4, rotation: 23 };
  const marker = markerBounds(source, region); assert.ok(near(relativeRegion(source, marker), region, 1e-9));
  assert.ok(near({ rotation: -90 }, { rotation: 270 }));
  const clipped = clipLine({ x: -10, y: 50 }, { x: 110, y: 50 }, { left: 0, top: 0, width: 100, height: 100, rotation: 0 });
  assert.ok(near(lineBounds(clipped), { left: 0, top: 50, width: 100, height: 0, rotation: 0 }));
});
test('readImageDimensions parses BMP, EMF and WMF headers and falls back to raster metadata', async () => {
  const bmp = Buffer.alloc(26); bmp.write('BM'); bmp.writeUInt32LE(40, 14); bmp.writeInt32LE(640, 18); bmp.writeInt32LE(-480, 22);
  assert.deepEqual(await readImageDimensions(bmp, 'image/bmp'), { width: 640, height: 480, mime: 'image/bmp' });
  const emf = Buffer.alloc(44); emf.writeUInt32LE(1, 0); emf.writeInt32LE(0, 8); emf.writeInt32LE(0, 12); emf.writeInt32LE(300, 16); emf.writeInt32LE(200, 20); emf.writeUInt32LE(0x464d4520, 40);
  assert.deepEqual(await readImageDimensions(emf, 'image/x-emf'), { width: 300, height: 200, mime: 'image/x-emf' });
  const wmf = Buffer.alloc(22); wmf.writeUInt32LE(0x9ac6cdd7, 0); wmf.writeInt16LE(0, 6); wmf.writeInt16LE(0, 8); wmf.writeInt16LE(100, 10); wmf.writeInt16LE(50, 12); wmf.writeUInt16LE(96, 14);
  assert.deepEqual(await readImageDimensions(wmf, 'image/wmf'), { width: 100, height: 50, mime: 'image/wmf' });
  const png = await sharp({ create: { width: 7, height: 5, channels: 4, background: '#336699ff' } }).png().toBuffer();
  assert.deepEqual(await readImageDimensions(png, 'image/png'), { width: 7, height: 5, mime: 'image/png' });
});
test('retained pseudo color copies a grouped picture without scale or zoom links', async () => {
  const pixels = await sharp({ create: { width: 2, height: 1, channels: 4, background: '#00ff00ff' } }).png().toBuffer();
  const fov = FOV_PREFIX + JSON.stringify({ width: 20, height: 10, unit: 'μm' }), scale = SCALE_PREFIX + JSON.stringify({ ...defaultScale });
  const picture = { id: 'picture', name: '原图', type: 'picture', left: 10, top: 20, width: 100, height: 80, rotation: 15,
    alternativeText: `${fov}\n${scale}`, tags: { SLIDESCI_SCALE_CONTENT: '1', SLIDESCI_FOV_VISIBLE: '1' }, crop: { left: .1, top: .2, right: .1, bottom: .1 } };
  const group = { id: 'group', name: '原图', type: 'group', left: 10, top: 20, width: 150, height: 80, tags: { SLIDESCI_SCALE_GROUP: '1', SLIDESCI_IMAGE_UUID: 'GROUP', SLIDESCI_ZOOM_RECORDS: 'HEX:' }, children: [picture] };
  let state = { shapes: [group], width: 960, height: 540 }, inserted = null;
  const applied = [];
  const ctx = {
    host: { kind: 'powerpoint', async context() { return state; }, async images() { return [{ id: 'picture', data: pixels.toString('base64'), mime: 'image/png', width: 2, height: 1, source: '/Users/Researcher/原图.png' }]; },
      async apply(commands) { applied.push(...commands); for (const command of commands) if (command.op === 'insertImage') { inserted = { id: 'created', type: 'picture', ...command, tags: command.tags }; state = { ...state, shapes: [...state.shapes, inserted] }; } return state; } },
    services: { async request(namespace) { if (namespace === 'images/color') return { data: pixels.toString('base64'), mime: 'image/png' }; if (namespace === 'images/settings') return {}; throw new Error(`unexpected ${namespace}`); } }
  };
  assert.deepEqual(await applyPseudoColors(ctx, [group], ['#ff0000'], true), { count: 1 });
  assert.equal(inserted.left, 178);
  assert.equal(inserted.top, 20); assert.equal(inserted.width, 100); assert.equal(inserted.height, 80); assert.equal(inserted.rotation, 15);
  assert.equal(inserted.name, '原图_伪彩'); assert.deepEqual(inserted.crop, picture.crop);
  assert.ok(inserted.alternativeText.startsWith(FOV_PREFIX)); assert.ok(!inserted.alternativeText.includes(SCALE_PREFIX));
  assert.equal(inserted.tags.SLIDESCI_PSEUDO_COLOR, '#ff0000'); assert.equal(inserted.tags.SLIDESCI_FOV_VISIBLE, '1');
  assert.notEqual(inserted.tags.SLIDESCI_IMAGE_UUID, 'GROUP');
  assert.ok(!Object.keys(inserted.tags).some(key => key.startsWith('SLIDESCI_ZOOM_') || key.startsWith('SLIDESCI_SCALE_')));
  assert.ok(applied.some(command => command.op === 'update' && command.properties.rotation === 15 && command.properties.crop === picture.crop));
});
test('partial zoom creation failure removes new annotations and keeps source records', async () => {
  const pngBytes = await sharp({ create: { width: 2, height: 2, channels: 4, background: '#808080ff' } }).png().toBuffer();
  const state = { width: 960, height: 540, shapes: [{ id: 'source', name: '原图', type: 'picture', left: 40, top: 50, width: 120, height: 90, rotation: 0,
    alternativeText: '', tags: { SLIDESCI_IMAGE_UUID: 'SOURCE' }, image: { digest: 'digest-1' } }] };
  const applied = [];
  const ctx = {
    host: { kind: 'powerpoint', async context() { return state; },
      async images() { return [{ data: pngBytes.toString('base64'), mime: 'image/png', width: 2, height: 2, source: null }]; },
      async apply(commands) {
        for (const command of commands) {
          applied.push(command);
          if (command.op === 'insertShape') state.shapes.push({ id: `${command.shapeType}-${state.shapes.length}`, type: 'shape', ...command });
          else if (command.op === 'insertImage') throw new Error('模拟插入失败');
          else if (command.op === 'update') { const shape = state.shapes.find(item => item.id === command.id); if (shape && command.properties.tags) shape.tags = { ...(shape.tags || {}), ...command.properties.tags }; }
          else if (command.op === 'delete') state.shapes = state.shapes.filter(item => !command.ids.includes(item.id));
        }
        return state;
      } },
    services: { async request(namespace) { if (namespace === 'images/crop') return { data: pngBytes.toString('base64'), mime: 'image/png', width: 1, height: 1 }; if (namespace === 'images/settings') return {}; throw new Error(`unexpected ${namespace}`); } }
  };
  const entry = { name: '放大图 1', region: { x: .25, y: .25, width: .25, height: .25, rotation: 0 }, color: '#ff0000', lineWidth: 1.5,
    dash: 'solid', useColor: true, guides: true, extent: 'across', placement: 'right', scaleLength: null, scaleShowText: true, key: null };
  await assert.rejects(applyZoomEntries(ctx, [state.shapes[0]], [entry]), /模拟插入失败/);
  assert.deepEqual(state.shapes.map(shape => shape.id), ['source']);
  assert.ok(applied.some(command => command.op === 'insertShape' && command.shapeType === 'rectangle'));
  assert.ok(!state.shapes.some(shape => shape.tags?.SLIDESCI_ZOOM_RECORD));
  assert.ok(!state.shapes[0].tags.SLIDESCI_ZOOM_RECORDS);
});

test('successful zoom keeps its source UUID distinct from the zoom entry key', async () => {
  const pixels = await sharp({ create: { width: 64, height: 48, channels: 4, background: '#808080ff' } }).png().toBuffer();
  const state = { width: 960, height: 540, shapes: [{ id: 'source', name: 'Source', type: 'picture', left: 40, top: 50, width: 120, height: 90, rotation: 0,
    alternativeText: FOV_PREFIX + JSON.stringify({ width: 100, height: 75, unit: 'μm' }),
    tags: { SLIDESCI_IMAGE_UUID: 'SOURCE' }, image: { digest: 'source-pixels' } }] };
  let nextId = 0;
  const ctx = {
    host: {
      async context() { return structuredClone(state); },
      async images() { return [{ data: pixels.toString('base64'), mime: 'image/png', width: 64, height: 48 }]; },
      async apply(commands) {
        for (const command of commands) {
          if (command.op === 'insertShape' || command.op === 'insertImage') state.shapes.push({ ...command, id: 'created-' + ++nextId,
            type: command.op === 'insertImage' ? 'picture' : 'shape' });
          else if (command.op === 'update') {
            const shape = state.shapes.find(shape => shape.id === command.id);
            const tags = { ...shape.tags, ...command.properties.tags };
            Object.assign(shape, command.properties, { tags });
          } else if (command.op === 'delete') state.shapes = state.shapes.filter(shape => !command.ids.includes(shape.id));
          else throw new Error('Unexpected zoom command: ' + command.op);
        }
        return structuredClone(state);
      }
    },
    services: { async request(route, input) {
      if (route === 'images/crop') return imageHandlers.crop(input, {});
      if (route === 'images/settings') return {};
      throw new Error('Unexpected zoom service: ' + route);
    } }
  };
  await applyZoomEntries(ctx, [state.shapes[0]], [{ name: 'Zoom', region: { x: .25, y: .25, width: .25, height: .25, rotation: 0 },
    placement: 'right', guides: true, color: '#ff0000', lineWidth: 1.5, dash: 'solid' }]);
  const record = parse(state.shapes[0].tags.SLIDESCI_ZOOM_RECORDS)[0];
  assert.notEqual(record.key, 'SOURCE');
  assert.equal(record.sourceState.left, 40);
  const linked = state.shapes.map(shape => parse(shape.tags?.SLIDESCI_ZOOM_RECORD)).filter(Boolean);
  assert.ok(linked.some(link => link.role === 'marker'));
  assert.ok(linked.some(link => link.role === 'zoom'));
  assert.ok(linked.every(link => link.sourceUUID === 'SOURCE' && link.key === record.key));
});

test('zoom creation records full-raster calibration and ignores legacy visible tagging', async () => {
  const pixels = await sharp({ create: { width: 64, height: 48, channels: 4, background: '#808080ff' } }).png().toBuffer();
  const crop = { left: .25, right: 0, top: 0, bottom: 0 };
  const state = { width: 960, height: 540, shapes: [{ id: 'source', name: 'Source', type: 'picture', left: 40, top: 50, width: 200, height: 100, rotation: 0, crop,
    alternativeText: FOV_PREFIX + JSON.stringify({ width: 100, height: 50, unit: 'nm' }),
    tags: { SLIDESCI_IMAGE_UUID: 'SOURCE' }, image: { digest: 'source-pixels' } }] };
  let nextId = 0;
  const ctx = {
    host: {
      async context() { return structuredClone(state); },
      async images() { return [{ data: pixels.toString('base64'), mime: 'image/png', width: 64, height: 48 }]; },
      async apply(commands) {
        for (const command of commands) {
          if (command.op === 'insertShape' || command.op === 'insertImage') state.shapes.push({ ...command, id: 'created-' + ++nextId,
            type: command.op === 'insertImage' ? 'picture' : 'shape' });
          else if (command.op === 'update') {
            const shape = state.shapes.find(shape => shape.id === command.id);
            shape.tags = { ...shape.tags, ...command.properties.tags };
            Object.assign(shape, command.properties, { tags: shape.tags });
          } else if (command.op === 'replaceImage') {
            const shape = state.shapes.find(shape => shape.id === command.id);
            assert.equal(command.preserveCrop, false);
            delete shape.crop;
          } else if (command.op === 'delete') state.shapes = state.shapes.filter(shape => !command.ids.includes(shape.id));
          else throw new Error('Unexpected zoom command: ' + command.op);
        }
        return structuredClone(state);
      }
    },
    services: { async request(route, input) {
      if (route === 'images/crop') return imageHandlers.crop(input, {});
      if (route === 'images/settings') return {};
      throw new Error('Unexpected zoom service: ' + route);
    } }
  };
  await applyZoomEntries(ctx, [state.shapes[0]], [{ name: 'Zoom', region: { x: 0, y: 0, width: .25, height: .5, rotation: 0 },
    placement: 'right', guides: false, color: '#ff0000', lineWidth: 1.5, dash: 'solid' }]);
  const zoom = state.shapes.find(shape => parse(shape.tags?.SLIDESCI_ZOOM_RECORD)?.role === 'zoom');
  const fov = parse(zoom.alternativeText.split('\n').find(line => line.startsWith(FOV_PREFIX)).slice(FOV_PREFIX.length));
  // 100 nm full width, 25 % cropped away, 25 % region width => 18.75 nm; the FOV is not reduced twice.
  assert.ok(Math.abs(fov.width - .01875) < 1e-12); assert.ok(Math.abs(fov.height - .025) < 1e-12);
  assert.equal(fov.unit, 'μm');
  assert.equal(zoom.tags.SLIDESCI_FOV_VISIBLE, undefined); assert.equal(zoom.crop, undefined);
  // A later native crop applies to the stored full-raster FOV instead of being suppressed by a tag.
  zoom.crop = { left: .2, right: 0, top: 0, bottom: 0 };
  assert.ok(Math.abs(visibleFov(zoom).width - .015) < 1e-12);

  // Old records contained the raw calibration in both signatures and did not detect an
  // already-miscalibrated zoom when nothing had moved. Synchronization migrates both together.
  const [legacy] = parse(state.shapes[0].tags.SLIDESCI_ZOOM_RECORDS);
  const rawFov = { width: 100, height: 50, unit: 'nm' };
  for (const key of ['sourceSignature', 'cropSignature']) legacy[key] = JSON.stringify({ ...JSON.parse(legacy[key]), fov: rawFov });
  state.shapes[0].tags.SLIDESCI_ZOOM_RECORDS = serializeTag([legacy]);
  zoom.alternativeText = FOV_PREFIX + JSON.stringify({ ...fov, width: .025 });
  await reconcileSlide(ctx);
  assert.ok(Math.abs(visibleFov(zoom).width - .01875) < 1e-12);
  assert.equal(zoom.crop, undefined);
});

test('merged channel pictures use full-raster calibration instead of a visible tag', async () => {
  const pixels = await sharp({ create: { width: 2, height: 1, channels: 4, background: '#808080ff' } }).png().toBuffer();
  const crop = { left: .25, right: 0, top: 0, bottom: 0 };
  const channel = (id, name) => ({ id, name, type: 'picture', left: 10, top: 20, width: 200, height: 100, rotation: 0, crop,
    alternativeText: FOV_PREFIX + JSON.stringify({ width: 100, height: 50, unit: 'μm' }), tags: {} });
  let state = { width: 960, height: 540, selectionIds: ['a', 'b'], shapes: [channel('a', 'DAPI'), channel('b', 'FITC')] };
  let inserted = null, applied = [];
  const ctx = {
    host: {
      async context() { return state; },
      async images() { return state.shapes.map(shape => ({ id: shape.id, data: pixels.toString('base64'), mime: 'image/png', width: 2, height: 1 })); },
      async apply(commands) {
        applied = applied.concat(commands);
        for (const command of commands) if (command.op === 'insertImage') {
          inserted = { id: 'merged', type: 'picture', ...command, tags: command.tags };
          state = { ...state, shapes: [...state.shapes, inserted] };
        }
        return state;
      }
    },
    services: { async request(route) {
      if (route === 'images/merge') return { data: pixels.toString('base64'), mime: 'image/png', width: 200, height: 100 };
      throw new Error('Unexpected merge service: ' + route);
    } }
  };
  await imageActions.find(action => action.id === 'images.merge').run(ctx);
  assert.equal(inserted.tags.SLIDESCI_FOV_VISIBLE, undefined);
  assert.ok(Math.abs(visibleFov(inserted).width - 75) < 1e-12);
  inserted.crop = { left: .2, right: 0, top: 0, bottom: 0 };
  assert.ok(Math.abs(visibleFov(inserted).width - 60) < 1e-12);
  assert.ok(applied.some(command => command.op === 'select' && command.ids[0] === 'merged'));
});

test('batched original images follow selection identities and keep each picture crop', async () => {
  const crop = { left: .2, right: .1, top: 0, bottom: 0 };
  const shapes = [{ id: 'group', children: [{ id: 'b', tags: { SLIDESCI_SCALE_CONTENT: '1' }, crop }] }, { id: 'a' }, { id: 'b', crop }];
  const calls = [];
  const ctx = { host: { async images(ids) {
    calls.push(ids);
    return [{ id: 'a', data: 'AAAA', mime: 'image/png' }, { id: 'b', data: 'BBBB', mime: 'image/png' }];
  } } };
  const images = await imagesFor(ctx, shapes);
  assert.deepEqual(images.map(image => image.data), ['BBBB', 'AAAA', 'BBBB']);
  assert.deepEqual(images.map(image => image.crop), [crop, undefined, crop]);
  assert.deepEqual(calls, [['b', 'a']]);
});

// Inserts one zoom for a picture on a 720 × 405 pt slide and returns the new inset's bounds.
async function placeZoom(source, placement) {
  const pixels = await sharp({ create: { width: 64, height: 48, channels: 4, background: '#808080ff' } }).png().toBuffer();
  const state = { width: 720, height: 405, shapes: [{ id: 'source', name: 'Source', type: 'picture', rotation: 0, ...source,
    alternativeText: '', tags: { SLIDESCI_IMAGE_UUID: 'SOURCE' }, image: { digest: 'source-pixels' } }] };
  let nextId = 0, inset;
  const ctx = {
    host: {
      async context() { return structuredClone(state); },
      async images() { return [{ data: pixels.toString('base64'), mime: 'image/png', width: 64, height: 48 }]; },
      async apply(commands) {
        for (const command of commands) {
          if (command.op === 'insertShape' || command.op === 'insertImage') {
            const shape = { ...command, id: 'created-' + ++nextId, type: command.op === 'insertImage' ? 'picture' : 'shape' };
            state.shapes.push(shape);
            if (command.op === 'insertImage') inset = shape;
          } else if (command.op === 'update') {
            const shape = state.shapes.find(shape => shape.id === command.id);
            Object.assign(shape, command.properties, { tags: { ...shape.tags, ...command.properties.tags } });
          } else throw new Error('Unexpected zoom command: ' + command.op);
        }
        return structuredClone(state);
      }
    },
    services: { async request(route, input) {
      if (route === 'images/crop') return imageHandlers.crop(input, {});
      if (route === 'images/settings') return {};
      throw new Error('Unexpected zoom service: ' + route);
    } }
  };
  await applyZoomEntries(ctx, [state.shapes[0]], [{ name: 'Zoom', region: { x: .25, y: .25, width: .25, height: .25, rotation: 0 },
    placement, guides: false, color: '#ff0000', lineWidth: 1.5, dash: 'solid' }]);
  return { left: inset.left, top: inset.top, width: inset.width, height: inset.height };
}

test('a new zoom inset that would leave the slide shrinks into the room on its side', async () => {
  const gap = .5 * 72 / 2.54;
  // The picture of the WPS regression deck: a full-height inset on its right would end at 888 pt.
  const right = await placeZoom({ left: 118.11, top: 55.12, width: 377.95, height: 251.97 }, 'right');
  assert.ok(Math.abs(right.left - (118.11 + 377.95 + gap)) < 1e-6, 'keeps its gap to the picture');
  assert.ok(Math.abs(right.left + right.width - 720) < 1e-6, 'ends at the slide edge');
  assert.ok(Math.abs(right.top - 55.12) < 1e-6, 'stays aligned with the picture top');
  assert.ok(Math.abs(right.width / right.height - 377.95 / 251.97) < 1e-6, 'keeps the region aspect');

  const left = await placeZoom({ left: 200, top: 100, width: 300, height: 200 }, 'left');
  assert.ok(Math.abs(left.left) < 1e-6, 'starts at the slide edge');
  assert.ok(Math.abs(left.left + left.width - (200 - gap)) < 1e-6, 'keeps its gap to the picture');
  assert.ok(Math.abs(left.width / left.height - 1.5) < 1e-6);

  // No room worth zooming into: the inset would end up smaller than the marked region, so it keeps its size.
  const crowded = await placeZoom({ left: 20, top: 50, width: 680, height: 300 }, 'right');
  assert.ok(Math.abs(crowded.height - 300) < 1e-6);
  assert.ok(Math.abs(crowded.width - 680) < 1e-6);
});
