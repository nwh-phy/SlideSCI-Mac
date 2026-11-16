import sharp from 'sharp';
import UTIF from 'utif';
import { DOMParser } from '@xmldom/xmldom';
import { mkdir, readFile, rename, rm, stat, writeFile } from 'node:fs/promises';
import { randomUUID } from 'node:crypto';
import path from 'node:path';

const units = { nm: .001, um: 1, 'μm': 1, micron: 1, microns: 1, micrometer: 1, micrometers: 1, micrometre: 1, micrometres: 1, mm: 1000, millimeter: 1000, millimeters: 1000, cm: 10000, centimeter: 10000, centimeters: 10000, m: 1e6, meter: 1e6, meters: 1e6, in: 25400, inch: 25400, inches: 25400, nanometer: .001, nanometers: .001 };
const factor = unit => units[String(unit ?? '').trim().replaceAll('µ', 'μ').toLowerCase()];
const positive = x => Number.isFinite(x) && x > 0;
const roundEven = x => { const n = Math.floor(x), d = x - n; return d === .5 ? n + (n % 2) : Math.round(x); };
export function autoDisplayLevels(pixels, channels) {
  let minimum = 65535, maximum = 0, count = 0;
  for (let i = 0; i < pixels.length; i += channels) {
    if (channels === 4 && !pixels[i + 3]) continue;
    for (let c = 0; c < Math.min(channels, 3); c++) { minimum = Math.min(minimum, pixels[i + c]); maximum = Math.max(maximum, pixels[i + c]); count++; }
  }
  if (!count) return [0, 1];
  if (maximum <= minimum) return [minimum, minimum + 1];
  const histogram = new Uint32Array(256), range = maximum - minimum;
  for (let i = 0; i < pixels.length; i += channels) {
    if (channels === 4 && !pixels[i + 3]) continue;
    for (let c = 0; c < Math.min(channels, 3); c++) histogram[Math.min(255, Math.floor((pixels[i + c] - minimum) * 256 / range))]++;
  }
  let first = -1, last = -1;
  for (let i = 0; i < 256; i++) if (histogram[i] > Math.floor(count / 5000) && histogram[i] <= Math.floor(count / 10)) { if (first < 0) first = i; last = i; }
  return first >= 0 && last > first ? [minimum + first * range / 256, minimum + last * range / 256] : [minimum, maximum];
}
export function display16(pixels, channels) {
  const [minimum, maximum] = autoDisplayLevels(pixels, channels), scale = 256 / (maximum - minimum + 1);
  // Every sample passes through the same two mappings, so tabulate them once instead of per channel per pixel.
  const levels = new Uint8Array(65536), alphas = new Uint8Array(65536);
  for (let value = 0; value < 65536; value++) {
    levels[value] = Math.max(0, Math.min(255, roundEven((value - minimum) * scale)));
    alphas[value] = Math.floor((value * 255 + 32767) / 65535);
  }
  const output = Buffer.alloc(pixels.length / channels * 4);
  for (let i = 0, j = 0; i < pixels.length; i += channels, j += 4) {
    for (let c = 0; c < 3; c++) output[j + c] = levels[pixels[i + (channels === 1 ? 0 : c)]];
    output[j + 3] = channels === 4 ? alphas[pixels[i + 3]] : 255;
  }
  return output;
}

// Read only the first classic/BigTIFF directory; offsets and metadata lengths are bounded.
export function firstTiffTags(bytes) {
  const b = Buffer.from(bytes), little = b.subarray(0, 2).toString() === 'II';
  if (!little && b.subarray(0, 2).toString() !== 'MM') throw new Error('不是 TIFF 文件。');
  const n = (offset, size) => {
    if (offset < 0 || offset + size > b.length) throw new Error('TIFF 数据偏移越界。');
    const value = size === 8 ? Number(little ? b.readBigUInt64LE(offset) : b.readBigUInt64BE(offset)) : little ? b.readUIntLE(offset, size) : b.readUIntBE(offset, size);
    if (!Number.isSafeInteger(value)) throw new Error('TIFF 偏移过大。');
    return value;
  };
  const version = n(2, 2), big = version === 43;
  if (version !== 42 && !big) throw new Error('未知 TIFF 版本。');
  if (big && (n(4, 2) !== 8 || n(6, 2) !== 0)) throw new Error('未知 BigTIFF 偏移类型。');
  const offset = n(big ? 8 : 4, big ? 8 : 4), count = n(offset, big ? 8 : 2);
  if (count > 4096) throw new Error('TIFF 标签过多。');
  const result = {}, wanted = new Set([256, 257, 258, 259, 262, 270, 277, 282, 283, 284, 296, 322, 323, 338]);
  for (let i = 0; i < count; i++) {
    const e = offset + (big ? 8 : 2) + i * (big ? 20 : 12), tag = n(e, 2), type = n(e + 2, 2), items = n(e + 4, big ? 8 : 4);
    if (!wanted.has(tag)) continue;
    const size = ({ 1: 1, 2: 1, 3: 2, 4: 4, 5: 8, 16: 8 })[type];
    if (!size || !items || items * size > 4 * 1024 * 1024) continue;
    const p = items * size <= (big ? 8 : 4) ? e + (big ? 12 : 8) : n(e + (big ? 12 : 8), big ? 8 : 4);
    if (p + items * size > b.length) throw new Error('TIFF 元数据不完整。');
    result[tag] = type === 2 ? b.subarray(p, p + items).toString('utf8').replace(/\0+$/, '') : Array.from({ length: items }, (_, j) => type === 5 ? n(p + j * 8, 4) / n(p + j * 8 + 4, 4) : n(p + j * size, size));
  }
  return { tags: result, little, big };
}
export function tiffFov(tags) {
  const width = tags[256]?.[0], height = tags[257]?.[0], description = tags[270] ?? '';
  const create = (w, h, unit, source) => positive(factor(unit)) && w >= 0 && h >= 0 && (positive(w) || positive(h)) ? { width: w, height: h, unit: factor(unit) === 1 ? 'μm' : unit, source } : null;
  if (description.trimStart().startsWith('<') && !/<!DOCTYPE|<!ENTITY/i.test(description)) {
    const xml = new DOMParser({ onError: () => {} }).parseFromString(description, 'application/xml');
    const pixels = Array.from(xml.getElementsByTagName('*')).find(n => n.localName === 'Pixels');
    if (xml.documentElement?.localName === 'OME' && pixels && +pixels.getAttribute('SizeX') === width && +pixels.getAttribute('SizeY') === height) {
      const x = +pixels.getAttribute('PhysicalSizeX'), y = +pixels.getAttribute('PhysicalSizeY');
      const result = create(width * x * factor(pixels.getAttribute('PhysicalSizeXUnit') || 'μm'), height * y * factor(pixels.getAttribute('PhysicalSizeYUnit') || 'μm'), 'μm', 'OME-TIFF 像素标定');
      if (result) return result;
    }
  }
  const lines = Object.fromEntries(description.split(/[\r\n]+/).filter(l => l.includes('=')).map(l => { const i = l.indexOf('='); return [l.slice(0, i).trim().toLowerCase(), l.slice(i + 1).trim()]; }));
  if (lines.imagej) {
    if (!factor(lines.unit)) return null;
    return create(width * (positive(+lines.pixel_width) ? +lines.pixel_width : 1 / tags[282]?.[0]), height * (positive(+lines.pixel_height) ? +lines.pixel_height : 1 / tags[283]?.[0]), lines.unit, 'ImageJ TIFF 标定');
  }
  const unit = tags[296]?.[0] === 2 ? 'in' : tags[296]?.[0] === 3 ? 'cm' : null;
  return unit ? create(width / tags[282]?.[0] * factor(unit), height / tags[283]?.[0] * factor(unit), 'μm', 'TIFF 分辨率（可能为打印 DPI，请核对）') : null;
}
async function inputBytes(input, env) {
  if (input.path) { const result = await env.files.read(input.path); return Buffer.isBuffer(result) ? result : Buffer.from(result.data, 'base64'); }
  if (!input.data) throw new Error('缺少原始图片数据。');
  return Buffer.from(input.data.replace(/^data:[^,]+,/, ''), 'base64');
}
const encoded = async (pixels, width, height, extra = {}) => ({ data: (await sharp(pixels, { raw: { width, height, channels: 4 } }).png().toBuffer()).toString('base64'), mime: 'image/png', width, height, ...extra });
function vectorMetadata(bytes, name = '', mime = '') {
  const type = String(mime).toLowerCase();
  if (bytes.length >= 26 && bytes.subarray(0, 2).toString() === 'BM') {
    const dibSize = bytes.readUInt32LE(14), width = dibSize === 12 ? bytes.readUInt16LE(18) : bytes.readInt32LE(18), height = dibSize === 12 ? bytes.readUInt16LE(20) : Math.abs(bytes.readInt32LE(22));
    if (!positive(width) || !positive(height)) throw new Error('BMP 图片没有有效的尺寸。');
    return { width, height, mime: 'image/bmp' };
  }
  if (/\bemf\b/i.test(type) || /\.emf$/i.test(name) || bytes.length >= 44 && bytes.readUInt32LE(40) === 0x464d4520) {
    if (bytes.length < 44 || bytes.readUInt32LE(0) !== 1 || bytes.readUInt32LE(40) !== 0x464d4520) throw new Error('EMF 文件头无效。');
    const width = bytes.readInt32LE(16) - bytes.readInt32LE(8), height = bytes.readInt32LE(20) - bytes.readInt32LE(12);
    if (!positive(width) || !positive(height)) throw new Error('EMF 图片没有有效的尺寸。');
    return { width, height, mime: 'image/x-emf' };
  }
  if (/\bwmf\b/i.test(type) || /\.wmf$/i.test(name) || bytes.length >= 4 && bytes.readUInt32LE(0) === 0x9ac6cdd7) {
    let width, height;
    if (bytes.length >= 22 && bytes.readUInt32LE(0) === 0x9ac6cdd7) {
      const inch = bytes.readUInt16LE(14);
      width = (bytes.readInt16LE(10) - bytes.readInt16LE(6)) * 96 / inch;
      height = (bytes.readInt16LE(12) - bytes.readInt16LE(8)) * 96 / inch;
    } else {
      for (let offset = 18; offset + 10 <= bytes.length;) {
        const words = bytes.readUInt32LE(offset), fn = bytes.readUInt16LE(offset + 4);
        if (words < 3 || offset + words * 2 > bytes.length) throw new Error('WMF 记录不完整。');
        if (fn === 0x020c || fn === 0x020e) { height = Math.abs(bytes.readInt16LE(offset + 6)); width = Math.abs(bytes.readInt16LE(offset + 8)); break; }
        offset += words * 2;
      }
    }
    if (!positive(width) || !positive(height)) throw new Error('WMF 图片没有有效的尺寸。');
    return { width, height, mime: 'image/wmf' };
  }
  return null;
}
// Shared with the PPTX media parser: vector headers are authoritative, raster formats use libvips.
/** Resolve intrinsic {width,height,mime}; vector BMP/EMF/WMF never go through sharp. Throws for unreadable bytes. */
export async function readImageDimensions(bytes, mime = '') {
  const buffer = Buffer.isBuffer(bytes) ? bytes
    : ArrayBuffer.isView(bytes) ? Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength)
    : typeof bytes === 'string' ? Buffer.from(bytes.replace(/^data:[^,]+,/, ''), 'base64')
    : Buffer.from(bytes);
  const vector = vectorMetadata(buffer, '', mime);
  if (vector) return vector;
  const metadata = await sharp(buffer, { page: 0, pages: 1 }).metadata();
  if (!positive(metadata.width) || !positive(metadata.height)) throw new Error('无法识别图片格式或尺寸。');
  const format = ({ jpeg: 'image/jpeg', png: 'image/png', gif: 'image/gif', svg: 'image/svg+xml', webp: 'image/webp', heif: 'image/heic', avif: 'image/avif', tiff: 'image/tiff', bmp: 'image/bmp' })[metadata.format];
  return { width: metadata.width, height: metadata.height, mime: format || String(mime || '') || null };
}
async function normalize(input, env) {
  const bytes = await inputBytes(input, env), isTiff = ['II', 'MM'].includes(bytes.subarray(0, 2).toString());
  let fov = null, highBitDepth = false, width, height, output;
  const vector = vectorMetadata(bytes, input.name || input.path || '', input.mime);
  if (vector) return { ...vector, data: bytes.toString('base64'), source: input.path || input.source || null, name: input.name || path.basename(input.path || '图片.emf'), fov: null, highBitDepth: false };
  if (isTiff) {
    const directory = firstTiffTags(bytes), tags = directory.tags;
    fov = tiffFov(tags); width = tags[256]?.[0]; height = tags[257]?.[0];
    highBitDepth = tags[258]?.every(x => x === 16) && [0, 1, 2].includes(tags[262]?.[0]);
    if (highBitDepth && !directory.big && !tags[322] && tags[284]?.[0] !== 2 && [1, 5, 8, 32773].includes(tags[259]?.[0] || 1)) {
      const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength), ifd = UTIF.decode(buffer)[0];
      if (!ifd) throw new Error('TIFF 中没有图像。');
      UTIF.decodeImage(buffer, ifd);
      const channels = tags[277]?.[0] || 1, count = width * height;
      if (![1, 3, 4].includes(channels)) throw new Error('16 位 TIFF 需要灰度、RGB 或 RGBA 像素。');
      const pixels = new Uint16Array(count * channels), raw = Buffer.from(ifd.data);
      if (raw.length < pixels.length * 2) throw new Error('TIFF 首帧像素不完整。');
      for (let i = 0; i < count; i++) for (let c = 0; c < channels; c++) {
        const index = tags[284]?.[0] === 2 ? c * count + i : i * channels + c;
        // UTIF converts decompressed 16-bit samples to little endian.
        pixels[i * channels + c] = raw.readUInt16LE(index * 2);
      }
      if (channels === 1 && tags[262]?.[0] === 0) for (let i = 0; i < pixels.length; i++) pixels[i] = 65535 - pixels[i];
      if (channels === 4 && tags[338]?.[0] === 1) for (let i = 0; i < pixels.length; i += 4) for (let c = 0; c < 3; c++) pixels[i + c] = pixels[i + 3] ? Math.min(65535, Math.round(pixels[i + c] * 65535 / pixels[i + 3])) : 0;
      output = await encoded(display16(pixels, channels), width, height, { fov, highBitDepth });
    } else if (highBitDepth) {
      const raw = await sharp(bytes, { page: 0, pages: 1 }).toColourspace('rgb16').raw({ depth: 'ushort' }).toBuffer({ resolveWithObject: true });
      const pixels = new Uint16Array(raw.data.buffer, raw.data.byteOffset, raw.data.length / 2);
      output = await encoded(display16(pixels, raw.info.channels), raw.info.width, raw.info.height, { fov, highBitDepth });
    }
  }
  if (!output) {
    if (isTiff) {
      const result = await sharp(bytes, { page: 0, pages: 1 }).rotate().png().toBuffer({ resolveWithObject: true });
      output = { data: result.data.toString('base64'), mime: 'image/png', width: result.info.width, height: result.info.height, fov, highBitDepth };
    } else {
      const metadata = await sharp(bytes, { page: 0, pages: 1 }).metadata();
      const mime = ({ jpeg: 'image/jpeg', png: 'image/png', gif: 'image/gif', svg: 'image/svg+xml', webp: 'image/webp', heif: 'image/heic', avif: 'image/avif' })[metadata.format] || input.mime;
      if (!mime?.startsWith('image/') || !positive(metadata.width) || !positive(metadata.height)) throw new Error('无法识别图片格式或尺寸。');
      output = { data: bytes.toString('base64'), mime, width: metadata.width, height: metadata.pageHeight || metadata.height, fov, highBitDepth };
    }
  }
  return { ...output, source: input.path || input.source || null, name: input.name || (input.path ? path.basename(input.path) : '图片.png') };
}
// PowerPoint stores a picture shrunk inside its frame as negative crop; the uncovered area is transparent.
async function paddedCrop(bytes, metadata, crop) {
  const w = metadata.width, h = metadata.height;
  const left = Math.round((crop.left || 0) * w), top = Math.round((crop.top || 0) * h);
  const right = Math.round((crop.right || 0) * w), bottom = Math.round((crop.bottom || 0) * h);
  if (w - left - right < 1 || h - top - bottom < 1) throw new Error('图片裁剪范围无效。');
  const x = Math.max(0, left), y = Math.max(0, top), x2 = Math.min(w, w - right), y2 = Math.min(h, h - bottom);
  if (x2 <= x || y2 <= y) throw new Error('图片裁剪范围无效。');
  const { data, info } = await sharp(bytes, { page: 0, pages: 1 }).extract({ left: x, top: y, width: x2 - x, height: y2 - y }).ensureAlpha()
    .extend({ left: Math.max(0, -left), top: Math.max(0, -top), right: Math.max(0, -right), bottom: Math.max(0, -bottom), background: { r: 0, g: 0, b: 0, alpha: 0 } })
    .raw().toBuffer({ resolveWithObject: true });
  return sharp(data, { raw: { width: info.width, height: info.height, channels: info.channels } });
}
async function raster(input, env, target) {
  const bytes = await inputBytes(input, env);
  let image = sharp(bytes, { page: 0, pages: 1 });
  const metadata = await image.metadata(), crop = input.crop || {};
  if (['left', 'top', 'right', 'bottom'].some(side => (crop[side] || 0) < 0)) {
    image = await paddedCrop(bytes, metadata, crop);
    if (target) image.resize(target.width, target.height, { fit: 'fill' });
    else image.resize({ width: input.maximum || 4096, height: input.maximum || 4096, fit: 'inside', withoutEnlargement: true });
    return image.ensureAlpha().raw().toBuffer({ resolveWithObject: true });
  }
  const left = Math.round((crop.left || 0) * metadata.width), top = Math.round((crop.top || 0) * metadata.height);
  const width = Math.round(metadata.width * (1 - (crop.left || 0) - (crop.right || 0))), height = Math.round(metadata.height * (1 - (crop.top || 0) - (crop.bottom || 0)));
  if (width < 1 || height < 1 || left < 0 || top < 0 || left + width > metadata.width || top + height > metadata.height) throw new Error('图片裁剪范围无效。');
  image.extract({ left, top, width, height });
  if (target) image.resize(target.width, target.height, { fit: 'fill' });
  else image.resize({ width: input.maximum || 4096, height: input.maximum || 4096, fit: 'inside', withoutEnlargement: true });
  return image.ensureAlpha().raw().toBuffer({ resolveWithObject: true });
}
async function color(input, env) {
  if (!/^#[\da-f]{6}$/i.test(input.color)) throw new Error('伪彩颜色必须是 #RRGGBB。');
  const { data, info } = await raster(input, env), rgb = [1, 3, 5].map(i => parseInt(input.color.slice(i, i + 2), 16));
  const peak = input.sourcePeak > 0 && input.sourcePeak <= 255 ? input.sourcePeak : 255;
  // The source byte and the recovered intensity are both 0–255; tabulate both mappings instead of dividing per pixel.
  const intensity = new Uint8Array(256), masked = [new Uint8Array(256), new Uint8Array(256), new Uint8Array(256)];
  for (let value = 0; value < 256; value++) {
    intensity[value] = Math.min(255, Math.floor((value * 255 + peak / 2) / peak));
    for (let c = 0; c < 3; c++) masked[c][value] = Math.floor(value * rgb[c] / 255);
  }
  const red = masked[0], green = masked[1], blue = masked[2];
  for (let i = 0; i < data.length; i += 4) {
    const value = intensity[Math.max(data[i], data[i + 1], data[i + 2])];
    data[i] = red[value]; data[i + 1] = green[value]; data[i + 2] = blue[value];
  }
  return encoded(data, info.width, info.height);
}
async function merge(input, env) {
  if (!Array.isArray(input.images) || input.images.length < 2 || input.images.length > 7) throw new Error('合并通道需要 2–7 张图片。');
  // Source byte × alpha is the same 256×256 product for every channel of every image; tabulate it once.
  const premultiplied = new Uint8Array(65536);
  for (let value = 0; value < 256; value++) for (let alpha = 0; alpha < 256; alpha++) premultiplied[(value << 8) | alpha] = Math.floor((value * alpha + 127) / 255);
  let output, size;
  for (const image of input.images) {
    const result = await raster(image, env, size);
    if (!size) { size = { width: result.info.width, height: result.info.height }; output = Buffer.alloc(size.width * size.height * 4); for (let i = 3; i < output.length; i += 4) output[i] = 255; }
    for (let i = 0; i < output.length; i += 4) {
      const alpha = result.data[i + 3];
      for (let c = 0; c < 3; c++) output[i + c] = Math.min(255, output[i + c] + premultiplied[(result.data[i + c] << 8) | alpha]);
    }
  }
  return encoded(output, size.width, size.height);
}
async function crop(input, env) {
  const { data, info } = await raster(input, env), r = input.region;
  if (!r || ![r.x, r.y, r.width, r.height, r.rotation || 0].every(Number.isFinite) || r.width <= 0 || r.height <= 0) throw new Error('放大区域无效。');
  const aspect = positive(input.displayAspect) ? input.displayAspect : info.width / info.height;
  const canvasWidth = info.width, canvasHeight = Math.max(1, Math.round(canvasWidth / aspect));
  const source = canvasHeight === info.height ? data : await sharp(data, { raw: { width: info.width, height: info.height, channels: 4 } }).resize(canvasWidth, canvasHeight, { fit: 'fill' }).raw().toBuffer();
  const width = Math.max(1, Math.round(r.width * canvasWidth)), height = Math.max(1, Math.round(r.height * canvasHeight));
  if (width * height > 64 * 1024 * 1024) throw new Error('放大区域过大。');
  const output = Buffer.alloc(width * height * 4), angle = (r.rotation || 0) * Math.PI / 180, co = Math.cos(angle), si = Math.sin(angle);
  const cx = (r.x + r.width / 2) * canvasWidth, cy = (r.y + r.height / 2) * canvasHeight;
  let overlap = 0;
  for (let y = 0; y < height; y++) for (let x = 0; x < width; x++) {
    const dx = x + .5 - width / 2, dy = y + .5 - height / 2, sx = cx + dx * co - dy * si - .5, sy = cy + dx * si + dy * co - .5;
    if (sx < -.5 || sy < -.5 || sx >= canvasWidth - .5 || sy >= canvasHeight - .5) continue;
    overlap++;
    const x0 = Math.floor(sx), y0 = Math.floor(sy), fx = sx - x0, fy = sy - y0, index = (y * width + x) * 4;
    let alpha = 0; const rgb = [0, 0, 0];
    for (let oy = 0; oy < 2; oy++) for (let ox = 0; ox < 2; ox++) {
      const xx = Math.max(0, Math.min(canvasWidth - 1, x0 + ox)), yy = Math.max(0, Math.min(canvasHeight - 1, y0 + oy));
      const p = (yy * canvasWidth + xx) * 4, weight = (ox ? fx : 1 - fx) * (oy ? fy : 1 - fy), a = source[p + 3] / 255;
      alpha += a * weight; for (let c = 0; c < 3; c++) rgb[c] += source[p + c] * a * weight;
    }
    output[index + 3] = Math.round(alpha * 255); for (let c = 0; c < 3; c++) output[index + c] = alpha ? Math.round(rgb[c] / alpha) : 0;
  }
  if (!overlap) throw new Error('放大矩形与原图没有重叠。');
  return encoded(output, width, height);
}
// Settings writes are serialized and atomic: concurrent read-modify-write calls (pseudo-color,
// scale bar and zoom all save here) must not drop each other's values or leave a torn file.
let settingsQueue = Promise.resolve();
function settings(input, env) {
  const task = settingsQueue.then(async () => {
    await mkdir(env.storageDir, { recursive: true });
    const filename = path.join(env.storageDir, 'images-settings.json');
    let saved = {};
    try {
      const parsed = JSON.parse(await readFile(filename, 'utf8'));
      if (parsed && typeof parsed === 'object' && !Array.isArray(parsed)) saved = parsed;
    } catch (error) {
      if (error.code !== 'ENOENT') await rename(filename, `${filename}.corrupt-${Date.now()}`).catch(() => {});
    }
    if (input.values && typeof input.values === 'object') {
      saved = { ...saved, ...input.values };
      const temporary = `${filename}.${randomUUID()}.tmp`;
      try { await writeFile(temporary, JSON.stringify(saved), { mode: 0o600 }); await rename(temporary, filename); }
      finally { await rm(temporary, { force: true }); }
    }
    return saved;
  });
  settingsQueue = task.catch(() => {});
  return task;
}
// Only the calibration numbers leave the service, so a source path recorded in the document is read
// even when this service session has not granted it (for example after a restart).
async function fovBytes(input, env) {
  try { return await inputBytes(input, env); }
  catch (error) {
    const target = typeof input.path === 'string' && path.isAbsolute(input.path) && !/[\0\r\n]/.test(input.path) ? input.path : null;
    if (!target) throw error;
    const info = await stat(target).catch(() => null);
    if (!info?.isFile()) throw new Error('原始文件已不存在或已移动：' + target);
    return readFile(target);
  }
}
export const imageHandlers = { normalize, color, merge, crop, settings, fov: async (input, env) => { const bytes = await fovBytes(input, env); return tiffFov(firstTiffTags(bytes).tags); } };
