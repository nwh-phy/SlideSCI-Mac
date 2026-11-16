#!/usr/bin/env node
// Renders the macOS ribbon icon set from Phosphor Icons.
//
// Every <Control> and <Item> in office/manifest.xml gets a key derived from its
// id (SlideSCI.AlignHorizontalCenter -> align-horizontal-center). Each key is
// mapped to one Phosphor glyph, optionally with a small badge glyph for
// copy/paste style variants. Groups reference their primary button's assets.
// Output is transparent 16/32/80 px PNG, deterministic per run.
//
// Usage: npm run icons

import { mkdir, readFile, readdir, rm, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { DOMParser } from '@xmldom/xmldom';
import sharp from 'sharp';

const here = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(here, '..');
const manifestFile = path.join(root, 'office/manifest.xml');
const phosphorDir = path.join(root, 'node_modules', '@phosphor-icons', 'core');
const outputDir = path.join(root, 'assets', 'ribbon');
const licenseSource = path.join(phosphorDir, 'LICENSE');
const licenseTarget = path.join(outputDir, 'PHOSPHOR-LICENSE.txt');

const SIZES = [16, 32, 80];
const CANVAS = 256;
const SUPERSAMPLE = 4;

// Limited functional palette: one hue per ribbon category, chosen to stay
// legible on both light and dark Office ribbons.
const COLOR = {
  arrange: '#3B74F0',
  size: '#12949E',
  position: '#7C5CF0',
  format: '#D96B14',
  label: '#219A53',
  image: '#D2437F',
  science: '#A34BC7',
  content: '#4A6FCC',
  media: '#C08212',
  neutral: '#7C8899',
  ai: '#8060FF',
};

// Primary glyph plus optional badge. Badge is drawn as a white glyph on a
// darkened disc of the category colour, bottom-right.
const ICONS = {
  'arrange': { glyph: 'squares-four', color: 'arrange' },

  'align-menu': { glyph: 'arrows-horizontal', color: 'arrange' },
  'align-left': { glyph: 'align-left', color: 'arrange' },
  'align-horizontal-center': { glyph: 'align-center-horizontal', color: 'arrange' },
  'align-right': { glyph: 'align-right', color: 'arrange' },
  'align-top': { glyph: 'align-top', color: 'arrange' },
  'align-vertical-center': { glyph: 'align-center-vertical', color: 'arrange' },
  'align-bottom': { glyph: 'align-bottom', color: 'arrange' },

  'size-menu': { glyph: 'resize', color: 'size' },
  'spacing-horizontal': { glyph: 'arrows-out-line-horizontal', color: 'size' },
  'spacing-vertical': { glyph: 'arrows-out-line-vertical', color: 'size' },
  'distribute-horizontal': { glyph: 'columns', color: 'size' },
  'distribute-vertical': { glyph: 'rows', color: 'size' },
  'copy-dimensions': { glyph: 'frame-corners', color: 'size', badge: 'copy' },
  'paste-dimensions': { glyph: 'frame-corners', color: 'size', badge: 'clipboard' },
  'copy-width': { glyph: 'arrows-horizontal', color: 'size', badge: 'copy' },
  'paste-width': { glyph: 'arrows-horizontal', color: 'size', badge: 'clipboard' },
  'copy-height': { glyph: 'arrows-vertical', color: 'size', badge: 'copy' },
  'paste-height': { glyph: 'arrows-vertical', color: 'size', badge: 'clipboard' },

  'position-menu': { glyph: 'crosshair', color: 'position' },
  'copy-position': { glyph: 'crosshair', color: 'position', badge: 'copy' },
  'paste-position': { glyph: 'crosshair', color: 'position', badge: 'clipboard' },
  'swap-position': { glyph: 'arrows-left-right', color: 'position' },
  'copy-relative-position': { glyph: 'target', color: 'position', badge: 'copy' },
  'paste-relative-position': { glyph: 'target', color: 'position', badge: 'clipboard' },

  'format-menu': { glyph: 'paint-brush', color: 'format' },
  'copy-shape-format': { glyph: 'shapes', color: 'format', badge: 'copy' },
  'paste-shape-format': { glyph: 'shapes', color: 'format', badge: 'clipboard' },
  'copy-text-format': { glyph: 'text-t', color: 'format', badge: 'copy' },
  'paste-text-format': { glyph: 'text-t', color: 'format', badge: 'clipboard' },
  'copy-group-format': { glyph: 'bounding-box', color: 'format', badge: 'copy' },
  'paste-group-format': { glyph: 'bounding-box', color: 'format', badge: 'clipboard' },
  'copy-crop': { glyph: 'crop', color: 'format', badge: 'copy' },
  'paste-crop': { glyph: 'crop', color: 'format', badge: 'clipboard' },

  'titles-menu': { glyph: 'text-aa', color: 'label' },
  'titles-below': { glyph: 'text-t', color: 'label', badge: 'arrow-down' },
  'titles-above': { glyph: 'text-t', color: 'label', badge: 'arrow-up' },
  'labels': { glyph: 'tag', color: 'label' },
  'update-labels': { glyph: 'tag', color: 'label', badge: 'arrows-clockwise' },

  'insert': { glyph: 'image', color: 'image', badge: 'plus' },
  'pseudo-color': { glyph: 'palette', color: 'image' },
  'merge-channels': { glyph: 'stack-plus', color: 'image' },
  'reveal-source': { glyph: 'folder-open', color: 'image' },
  'paste-images': { glyph: 'images', color: 'image', badge: 'clipboard' },

  'scale-bar': { glyph: 'ruler', color: 'science' },
  'zoom-images': { glyph: 'magnifying-glass-plus', color: 'science' },
  'science-menu': { glyph: 'flask', color: 'science' },
  'set-fov': { glyph: 'ruler', color: 'science', badge: 'pencil-simple' },
  'read-fov': { glyph: 'ruler', color: 'science', badge: 'arrows-clockwise' },

  'markdown': { glyph: 'markdown-logo', color: 'content' },
  'equation': { glyph: 'math-operations', color: 'content' },
  'content-menu': { glyph: 'dots-three', color: 'content' },
  'code': { glyph: 'code', color: 'content' },
  'latex-svg': { glyph: 'function', color: 'content' },
  'svg': { glyph: 'bezier-curve', color: 'content' },
  'rich-text': { glyph: 'paragraph', color: 'content' },

  'library': { glyph: 'books', color: 'media' },
  'export': { glyph: 'export', color: 'media' },
  'export-menu': { glyph: 'export', color: 'media', badge: 'dots-three' },
  'export-original': { glyph: 'images', color: 'media', badge: 'export' },
  'copy-original': { glyph: 'image', color: 'media', badge: 'copy' },
  'copy-selection': { glyph: 'selection', color: 'media', badge: 'copy' },
  'library-save': { glyph: 'books', color: 'media', badge: 'plus' },

  'open': { glyph: 'sidebar-simple', color: 'neutral' },
  'assistant': { glyph: 'sparkle', color: 'ai' },
  'settings': { glyph: 'gear-six', color: 'neutral' },
};

const glyphCache = new Map();

function phosphorFile(name, weight) {
  const file = weight === 'fill' ? `${name}-fill.svg` : `${name}.svg`;
  return path.join(phosphorDir, 'assets', weight, file);
}

async function glyphPaths(name, weight) {
  const file = phosphorFile(name, weight);
  if (glyphCache.has(file)) return glyphCache.get(file);
  let svg;
  try {
    svg = await readFile(file, 'utf8');
  } catch {
    throw new Error(`缺少 Phosphor 图标 ${weight}/${name}：${path.relative(root, file)}`);
  }
  const paths = [...svg.matchAll(/<path\b[^>]*\bd="([^"]+)"/g)].map(match => match[1]);
  if (!paths.length) throw new Error(`Phosphor 图标没有路径数据：${path.relative(root, file)}`);
  glyphCache.set(file, paths);
  return paths;
}

function darken(hex, factor = 0.72) {
  const value = Number.parseInt(hex.slice(1), 16);
  const channel = shift => Math.round(((value >> shift) & 0xff) * factor);
  return `#${[channel(16), channel(8), channel(0)].map(n => n.toString(16).padStart(2, '0')).join('')}`;
}

async function glyphGroup(name, weight, box, x, y, color) {
  const paths = await glyphPaths(name, weight);
  const body = paths.map(d => `<path d="${d}"/>`).join('');
  return `<g transform="translate(${x} ${y}) scale(${box / CANVAS})" fill="${color}">${body}</g>`;
}

async function ribbonSvg(entry, pixelSize) {
  const color = COLOR[entry.color];
  if (!color) throw new Error(`未定义的图标颜色：${entry.color}`);

  const parts = [];
  if (entry.badge) {
    // Larger badge: the primary glyph stays clear of the seal and stays
    // readable at 16 px.
    parts.push(await glyphGroup(entry.glyph, 'regular', 184, 2, 2, color));
    parts.push(await glyphGroup('circle', 'fill', 84, 170, 170, darken(color)));
    parts.push(await glyphGroup(entry.badge, 'regular', 62, 181, 181, '#FFFFFF'));
  } else {
    parts.push(await glyphGroup(entry.glyph, 'regular', 208, 24, 24, color));
  }

  const renderSize = pixelSize * SUPERSAMPLE;
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${renderSize}" height="${renderSize}" viewBox="0 0 ${CANVAS} ${CANVAS}">${parts.join('')}</svg>`;
}

function keyFor(id) {
  return id
    .replace(/^SlideSCI\./, '')
    .replace(/([a-z0-9])([A-Z])/g, '$1-$2')
    .replace(/([A-Z]+)([A-Z][a-z])/g, '$1-$2')
    .toLowerCase();
}

async function manifestKeys() {
  const xml = await readFile(manifestFile, 'utf8');
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  const controls = [];

  const walk = node => {
    for (let child = node.firstChild; child; child = child.nextSibling) {
      if (child.nodeType !== 1) continue;
      if (child.localName === 'Control' || child.localName === 'Item') {
        const id = child.getAttribute('id');
        if (id && id.startsWith('SlideSCI.')) controls.push(keyFor(id));
      }
      walk(child);
    }
  };

  walk(doc);
  return [...new Set(controls)].sort();
}

async function main() {
  await mkdir(outputDir, { recursive: true });

  const controls = await manifestKeys();
  const missing = controls.filter(key => !ICONS[key]);
  if (missing.length) {
    throw new Error(`以下功能区控件缺少图标映射：\n  ${missing.join('\n  ')}`);
  }

  const unused = Object.keys(ICONS).filter(key => !controls.includes(key));
  if (unused.length) {
    console.warn(`ribbon-icons: 未使用的映射（已跳过）：${unused.join(', ')}`);
  }

  const jobs = controls.map(key => [key, ICONS[key]]);

  for (const [key, entry] of jobs) {
    for (const size of SIZES) {
      const svg = await ribbonSvg(entry, size);
      const file = path.join(outputDir, `${key}-${size}.png`);
      await sharp(Buffer.from(svg))
        .resize(size, size, { fit: 'fill', kernel: 'lanczos3' })
        .png({ compressionLevel: 9, adaptiveFiltering: false, palette: false })
        .toFile(file);
    }
  }

  const expected = new Set(jobs.flatMap(([key]) => SIZES.map(size => `${key}-${size}.png`)));
  for (const name of await readdir(outputDir)) {
    if (/-\d+\.png$/.test(name) && !expected.has(name)) {
      await rm(path.join(outputDir, name));
    }
  }

  await writeFile(licenseTarget, await readFile(licenseSource, 'utf8'));

  console.log(
    `ribbon-icons: ${controls.length} 个控件/项目，` +
      `生成 ${jobs.length * SIZES.length} 个 PNG（${SIZES.join('/')}）。`,
  );
}

await main();
