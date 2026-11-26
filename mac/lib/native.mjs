import { spawn } from 'node:child_process';
import { access, chmod, mkdir, readFile, realpath, stat, writeFile, rename, rm } from 'node:fs/promises';
import { constants } from 'node:fs';
import { homedir } from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createHash, randomUUID } from 'node:crypto';

const helperSource = fileURLToPath(new URL('../native/SlideSCI.swift', import.meta.url));
const issuedFiles = new Set();
const issuedFolders = new Set();
const reservedFiles = new Set();
const watches = new Map();
let compiling;
let officeQueue = Promise.resolve();
let clipboardOperationDepth = 0;
let ignoredGeneralChangeCount = -1;

// restoreClipboard/writeClipboard run in a separate helper process and bump NSPasteboard.general's
// changeCount. The watcher polls every 200ms, so it often sees that write after clipboardOperationDepth
// drops back to 0. Suppress general-pasteboard events up through the last internal write count.
function noteClipboardWrite(result) {
  if (Number.isInteger(result?.changeCount) && result.changeCount > ignoredGeneralChangeCount) {
    ignoredGeneralChangeCount = result.changeCount;
  }
  return result;
}
export function shouldSuppressWatchEvent(event, { depth = clipboardOperationDepth, ignoredCount = ignoredGeneralChangeCount } = {}) {
  if (!event || event.type === 'nativeError') return false;
  if (depth > 0) return true;
  if (event.type !== 'mediaSource' && event.type !== 'clipboardImage') return false;
  const general = event.reason === 'clipboard' || event.pasteboard === 'Apple CFPasteboard general';
  return general && Number.isInteger(event.changeCount) && event.changeCount <= ignoredCount;
}

function processCall(executable, args, input, { timeout = 120000, maxBytes = 256 * 1024 * 1024 } = {}) {
  return new Promise((resolve, reject) => {
    const child = spawn(executable, args, { stdio: ['pipe', 'pipe', 'pipe'] });
    const stdout = [], stderr = []; let size = 0, settled = false;
    const timer = setTimeout(() => { child.kill('SIGTERM'); finish(new Error('Native operation timed out.')); }, timeout);
    function finish(error, result) { if (settled) return; settled = true; clearTimeout(timer); error ? reject(error) : resolve(result); }
    child.on('error', (error) => finish(error));
    child.stdout.on('data', (chunk) => { size += chunk.length; if (size > maxBytes) { child.kill(); finish(new Error('Native response exceeds the size limit.')); } else stdout.push(chunk); });
    child.stderr.on('data', (chunk) => { if (Buffer.concat(stderr).length < 65536) stderr.push(chunk); });
    child.on('close', (code) => { if (code === 0) finish(null, Buffer.concat(stdout).toString('utf8').trim()); else finish(new Error(Buffer.concat(stderr).toString('utf8').trim() || `Native process exited (${code}).`)); });
    child.stdin.on('error', (error) => { if (error.code !== 'EPIPE') finish(error); });
    child.stdin.end(input);
  });
}
async function helperBinary() {
  if (process.platform === 'darwin' && process.arch === 'arm64') {
    const bundled = fileURLToPath(new URL('../bin/slidesci-native', import.meta.url));
    try { await access(bundled, constants.X_OK); return bundled; }
    catch (error) { if (error.code !== 'ENOENT') throw error; }
  }
  if (!compiling) compiling = (async () => {
    const source = await readFile(helperSource);
    const digest = createHash('sha256').update(source).digest('hex').slice(0, 20);
    const directory = path.join(homedir(), 'Library', 'Caches', 'SlideSCI-Mac', 'native');
    await mkdir(directory, { recursive: true, mode: 0o700 });
    const binary = path.join(directory, `slidesci-${digest}`);
    try { await access(binary, constants.X_OK); return binary; } catch (error) { if (error.code !== 'ENOENT' && error.code !== 'EACCES') throw error; }
    const temporary = `${binary}-${randomUUID()}`;
    try {
      await processCall('/usr/bin/xcrun', ['swiftc', '-O', '-swift-version', '5', helperSource, '-o', temporary, '-framework', 'AppKit', '-framework', 'Security', '-framework', 'PDFKit', '-framework', 'CoreImage', '-framework', 'ApplicationServices', '-framework', 'AVFoundation', '-framework', 'UniformTypeIdentifiers'], '', { timeout: 180000 });
      await chmod(temporary, 0o700); await rename(temporary, binary); return binary;
    } finally { await rm(temporary, { force: true }); }
  })().catch((error) => { compiling = undefined; throw error; });
  return compiling;
}
async function helper(action, input = {}) {
  const output = await processCall(await helperBinary(), [], JSON.stringify({ ...input, action }));
  return JSON.parse(output);
}
function hostContainer(host = 'powerpoint') {
  const id = host === 'wps' ? 'com.kingsoft.wpsoffice.mac' : host === 'powerpoint' ? 'com.microsoft.Powerpoint' : null;
  if (!id) throw new Error('Unknown presentation host.');
  return path.join(homedir(), 'Library', 'Containers', id, 'Data', 'tmp', 'SlideSCI-Mac');
}
const extensions = {
  'image/png': '.png', 'image/jpeg': '.jpg', 'image/tiff': '.tiff', 'image/bmp': '.bmp', 'image/gif': '.gif',
  'image/svg+xml': '.svg', 'image/emf': '.emf', 'image/x-emf': '.emf', 'application/x-emf': '.emf',
  'image/wmf': '.wmf', 'image/x-wmf': '.wmf', 'application/x-wmf': '.wmf',
  'application/pdf': '.pdf', 'application/vnd.openxmlformats-officedocument.presentationml.presentation': '.pptx',
  'video/mp4': '.mp4', 'video/quicktime': '.mov', 'video/webm': '.webm', 'text/plain': '.txt',
};
const mimeTypes = {
  '.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.tif': 'image/tiff', '.tiff': 'image/tiff',
  '.bmp': 'image/bmp', '.gif': 'image/gif', '.svg': 'image/svg+xml', '.emf': 'image/emf', '.wmf': 'image/wmf',
  '.pdf': 'application/pdf', '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  '.mp4': 'video/mp4', '.m4v': 'video/mp4', '.mov': 'video/quicktime', '.webm': 'video/webm',
  '.avi': 'video/x-msvideo', '.wmv': 'video/x-ms-wmv', '.mpg': 'video/mpeg', '.mpeg': 'video/mpeg',
};
// Linear checks only: a repeated-group regex over a multi-megabyte string overflows V8's regex
// stack (a 300 dpi slide export or a large picture is several megabytes of base64).
function decode(data) {
  if (typeof data !== 'string' || data.length % 4 || /[^A-Za-z0-9+/=]/.test(data) || /=[^=]|={3}/.test(data)) throw new Error('Invalid base64 file data.');
  return Buffer.from(data, 'base64');
}
// APFS limits a name to 255 UTF-8 bytes; cached copies carry a 37-byte UUID prefix.
function shortName(name, limit = 180) {
  if (Buffer.byteLength(name) <= limit) return name;
  const extension = path.extname(name).slice(0, 16), characters = [...name.slice(0, name.length - path.extname(name).length)];
  while (characters.length && Buffer.byteLength(characters.join('') + extension) > limit) characters.pop();
  return (characters.join('') || 'media') + extension;
}
function safeName(name) { if (typeof name !== 'string' || !name || name === '.' || name === '..' || /[\x00-\x1f/\\:]/.test(name) || path.basename(name) !== name) throw new Error('Invalid output filename.'); return name; }
// macOS asks before one app's process writes into another app's container; a denial must read as such.
function containerError(error, host) {
  if (!['EPERM', 'EACCES'].includes(error?.code)) return error;
  const name = host === 'wps' ? 'WPS' : 'PowerPoint';
  return new Error(`macOS 拒绝访问 ${name} 的数据目录。请在“访问其他 App 的数据”提示中选择允许，或在“系统设置 → 隐私与安全性”中为运行 SlideSCI 服务的终端开启相应权限，然后重新运行 Start.command。`);
}
async function materialize({ data, mime = 'application/octet-stream', name, host = 'powerpoint' }) {
  const directory = hostContainer(host); await mkdir(directory, { recursive: true, mode: 0o700 }).catch((error) => { throw containerError(error, host); });
  const base = shortName(name ? safeName(name) : 'media');
  const filename = `${randomUUID()}-${base}${path.extname(base) ? '' : extensions[mime] || '.bin'}`;
  const output = path.join(directory, filename); await writeFile(output, decode(data), { mode: 0o600, flag: 'wx' }).catch((error) => { throw containerError(error, host); }); issuedFiles.add(await realpath(output)); return { path: output, name: name || path.basename(output), mime };
}
async function authorizeFile(filename, env) {
  if (typeof filename !== 'string' || !path.isAbsolute(filename) || /[\x00-\x1f]/.test(filename)) throw new Error('A valid absolute path is required.');
  const actual = await realpath(filename);
  if (reservedFiles.has(path.resolve(filename)) && actual === path.resolve(filename)) { issuedFiles.add(actual); return actual; }
  if (issuedFiles.has(actual)) return actual;
  if (env?.files?.requireFile) { await env.files.requireFile(actual); return actual; }
  if (env?.files?.read) { await env.files.read(actual); return actual; }
  throw new Error('Choose this file in the native file picker before reading it.');
}
async function readIssued({ path: filename }, env) { const actual = await authorizeFile(filename, env); return { path: actual, name: path.basename(actual), mime: mimeTypes[path.extname(actual).toLowerCase()] || 'application/octet-stream', data: (await readFile(actual)).toString('base64') }; }
async function nativeAsset(input, env) {
  if (!input.path) return (await materialize({ ...input, host: 'powerpoint' })).path;
  const filename = await authorizeFile(input.path, env);
  await mkdir(hostContainer('powerpoint'), { recursive: true, mode: 0o700 });
  const directory = await realpath(hostContainer('powerpoint'));
  if (filename.startsWith(directory + path.sep)) return filename;
  return (await materialize({ data: (await readFile(filename)).toString('base64'), mime: input.mime || mimeTypes[path.extname(filename).toLowerCase()], name: path.basename(filename), host: 'powerpoint' })).path;
}
async function discardCached(filename) {
  const resolved = path.resolve(filename);
  reservedFiles.delete(resolved); issuedFiles.delete(resolved);
  await rm(resolved, { force: true });
}
async function pickFiles(input = {}) { const result = await helper('pickFiles', input); for (const file of result) issuedFiles.add(await realpath(file.path)); return result.map((file) => ({ ...file, mime: mimeTypes[path.extname(file.path).toLowerCase()] || 'application/octet-stream' })); }
async function chooseFolder() { const result = await helper('chooseFolder'); if (result) issuedFolders.add(await realpath(result)); return result; }
async function writeIssued({ folder, name, data }) { const actual = await realpath(folder); if (!issuedFolders.has(actual)) throw new Error('Choose the destination folder before writing files.'); const filename = path.join(actual, safeName(name)); await writeFile(filename, decode(data), { flag: 'wx', mode: 0o600 }); issuedFiles.add(await realpath(filename)); return { path: filename, name }; }
function asText(value) { return '"' + String(value).replaceAll('\\', '\\\\').replaceAll('"', '\\"').replaceAll('\r', '\\r').replaceAll('\n', '\\n') + '"'; }
function number(value, label = 'Number') { if (!Number.isFinite(Number(value))) throw new Error(`${label} must be finite.`); return Number(value); }
function index(value) { const n = number(value, 'Index'); if (!Number.isInteger(n) || n < 0) throw new Error('Slide indices must be nonnegative integers.'); return n; }
function refExpression(ref, presentation = 'targetPresentation') {
  if (!ref || !Array.isArray(ref.shapePath) || !ref.shapePath.length || ref.shapePath.some((n) => !Number.isInteger(n) || n < 1)) throw new Error('A resolved {slideIndex, shapePath:[1-based indices]} is required.');
  let result = `slide ${index(ref.slideIndex) + 1} of ${presentation}`;
  for (const part of ref.shapePath) result = `shape ${part} of (${result})`;
  return result;
}
function refsFor(input) { return (input.refs || []).map((entry) => entry.ref ? entry : { id: entry.id || JSON.stringify(entry), ref: entry }); }
const prelude = `use scripting additions\n`;
const foundationPrelude = `use framework "Foundation"\n${prelude}`;
async function appleScript(body, { json = false } = {}) {
  // Loading Foundation changes how PowerPoint receives file specifiers: `open` and `save` silently fail or
  // open modal panels. Only JSON-returning scripts need it, so keep file operations on the plain prelude.
  const script = `${json ? foundationPrelude : prelude}\nwith timeout of 180 seconds\n${body}\nend timeout`;
  try {
    const result = await processCall('/usr/bin/osascript', ['-'], script, { timeout: 190000 });
    return json ? JSON.parse(result) : result;
  } catch (error) {
    const text = String(error.message || '');
    // -1743: the user has not allowed (or refused) automation; -1712: an unanswered permission prompt or a modal dialog.
    if (/-1743|not authori[sz]ed to send Apple events|不允许发送 Apple 事件/i.test(text)) error.message = 'macOS 尚未允许 SlideSCI 服务控制 PowerPoint。请在“系统设置 → 隐私与安全性 → 自动化”中，为运行 SlideSCI 的终端打开 Microsoft PowerPoint，然后重试。';
    else if (/-1712|AppleEvent.*(timed out|已超时)/i.test(text)) error.message = 'PowerPoint 没有及时响应。如果屏幕上有“终端想要控制 Microsoft PowerPoint”的提示，请选择允许后重试；也请确认 PowerPoint 没有打开中的对话框。';
    if (process.env.SLIDESCI_NATIVE_DEBUG) error.message += `\n--- AppleScript ---\n${body}`;
    throw error;
  }
}
function jsonReturn(value) { return `set jsonData to current application's NSJSONSerialization's dataWithJSONObject:${value} options:0 |error|:(missing value)\nreturn (current application's NSString's alloc()'s initWithData:jsonData encoding:(current application's NSUTF8StringEncoding)) as text`; }
function office(body, expected) {
  const guard = expected ? `if (name of targetPresentation) is not ${asText(expected)} then error "The active PowerPoint presentation is not the requested document." number 1\n` : '';
  return `tell application "Microsoft PowerPoint"\nset targetPresentation to active presentation\n${guard}set targetWindow to active window\nset targetSlide to missing value\nif (count of slides of targetPresentation) > 0 then set targetSlide to slide of view of targetWindow\n${body}\nend tell`;
}

function openFile(filename) {
  const sourceName = path.basename(filename);
  return `set sourceName to ${asText(sourceName)}\nopen (POSIX file ${asText(filename)})\nset openWait to 0\nrepeat while (name of active presentation) is not sourceName and openWait < 200\ndelay 0.05\nset openWait to openWait + 1\nend repeat\nif (name of active presentation) is not sourceName then error ("PowerPoint did not open " & sourceName) number 1`;
}
function waitForPresentation(expression) {
  return `set presentationWait to 0\nrepeat while (name of active presentation) is not (${expression}) and presentationWait < 200\ndelay 0.05\nset presentationWait to presentationWait + 1\nend repeat\nif (name of active presentation) is not (${expression}) then error "PowerPoint did not return to the target presentation." number 1`;
}
async function preservingClipboard(operation) { clipboardOperationDepth++; let items; try { items = await helper('snapshotClipboard'); return await operation(); } finally { try { if (items) noteClipboardWrite(await helper('restoreClipboard', { items })); } finally { clipboardOperationDepth--; } } }
function rgb(value) { if (!/^#[0-9a-f]{6}$/i.test(value)) throw new Error('Colors must be #RRGGBB.'); return '{' + [1, 3, 5].map((offset) => parseInt(value.slice(offset, offset + 2), 16)).join(', ') + '}'; }
function enumValue(mapping, value, label) { if (!(value in mapping)) throw new Error(`Unknown ${label}: ${value}`); return mapping[value]; }
const strike = { true: 'single strike', false: 'no strike' };
function fontScript(font, target) {
  const result = [];
  const properties = { name: ['font name', 'text'], nameAscii: ['ASCII name', 'text'], nameFarEast: ['east asian name', 'text'], size: ['font size', 'number'], bold: ['bold', 'bool'], italic: ['italic', 'bool'], underline: ['underline', 'bool'], superscript: ['superscript', 'bool'], subscript: ['subscript', 'bool'], emboss: ['emboss', 'bool'], baselineOffset: ['base line offset', 'number'], kerning: ['kerning', 'number'], spacing: ['spacing', 'number'], transparency: ['transparency', 'number'] };
  for (const [key, [property, type]] of Object.entries(properties)) if (font[key] !== undefined) result.push(`set ${property} of (${target}) to ${type === 'bool' ? Boolean(font[key]) : type === 'number' ? number(font[key]) : asText(font[key])}`);
  if (font.color) result.push(`set font color of (${target}) to ${rgb(font.color)}`);
  if (font.strikethrough !== undefined) result.push(`set strike type of (${target}) to ${strike[Boolean(font.strikethrough)]}`);
  if (font.highlight) result.push(`set highlight color of (${target}) to ${rgb(typeof font.highlight === 'string' ? font.highlight : font.highlight.color)}`);
  if (font.underlineColor) result.push(`set underline color of (${target}) to ${rgb(font.underlineColor)}`);
  if (font.allCaps !== undefined || font.smallCaps !== undefined) result.push(`set caps type of (${target}) to ${font.allCaps ? 'all caps' : font.smallCaps ? 'small caps' : 'no caps'}`);
  if (font.outline) result.push(formattingScript({ line: font.outline }, target));
  if (font.shadow) {
    const shadow = font.shadow, s = `shadow format of (${target})`;
    for (const [key, property] of [['blur', 'blur'], ['offsetX', 'offset X'], ['offsetY', 'offset Y'], ['size', 'size'], ['transparency', 'transparency']]) if (shadow[key] !== undefined) result.push(`set ${property} of (${s}) to ${number(shadow[key])}`);
    if (shadow.visible !== undefined) result.push(`set visible of (${s}) to ${Boolean(shadow.visible)}`);
    if (shadow.color) result.push(`set fore color of (${s}) to ${rgb(shadow.color)}`);
  }
  if (font.glow) { const g = `glow format of (${target})`; if (font.glow.color) result.push(`set color of (${g}) to ${rgb(font.glow.color)}`); if (font.glow.radius !== undefined) result.push(`set radius of (${g}) to ${number(font.glow.radius)}`); }
  if (font.reflection?.type !== undefined) {
    const type = index(font.reflection.type); if (type > 9) throw new Error('Reflection preset must be 0 through 9.');
    result.push(`set reflection type of reflection format of (${target}) to ${type === 0 ? 'reflection type none' : `reflection type${type}`}`);
  }
  return result.join('\n');
}
function textScript(properties, target) {
  const result = [], textRange = `text range of text frame of ${target}`;
  let runs = properties.runs;
  if (properties.paragraphs?.some((paragraph) => paragraph.runs || paragraph.text !== undefined)) runs = properties.paragraphs.flatMap((paragraph, i) => [...(paragraph.runs || [{ text: paragraph.text || '', font: paragraph.font }]), ...(i < properties.paragraphs.length - 1 ? [{ text: '\r' }] : [])]);
  if (runs) {
    const text = runs.map((run) => run.text || '').join(''); result.push(`set content of (${textRange}) to ${asText(text)}`);
    let offset = 1;
    for (const run of runs) { const length = (run.text || '').length; if (length && run.font) result.push(fontScript(run.font, `font of (characters ${offset} thru ${offset + length - 1} of (${textRange}))`)); offset += length; }
  } else if (properties.text !== undefined) result.push(`set content of (${textRange}) to ${asText(properties.text)}`);
  if (properties.font) result.push(fontScript(properties.font, `font of (${textRange})`));
  if (properties.textRange) {
    const range = properties.textRange, start = index(range.start ?? 0) + 1, length = index(range.length);
    if (!length) throw new Error('Select at least one text character to apply text formatting.');
    const selection = `characters ${start} thru ${start + length - 1} of (${textRange})`;
    if (range.text !== undefined) result.push(`set content of (${selection}) to ${asText(range.text)}`);
    if (range.font) result.push(fontScript(range.font, `font of (${selection})`));
  }
  if (properties.textFrame) {
    const frame = properties.textFrame, f = `text frame of ${target}`;
    for (const [key, property] of [['marginLeft', 'margin left'], ['marginRight', 'margin right'], ['marginTop', 'margin top'], ['marginBottom', 'margin bottom']]) if (frame[key] !== undefined) result.push(`set ${property} of (${f}) to ${number(frame[key])}`);
    if (frame.wordWrap !== undefined) result.push(`set word wrap of (${f}) to ${Boolean(frame.wordWrap)}`);
    if (frame.autoSize) result.push(`set auto size of (${f}) to ${enumValue({ none: 'auto size none', shapeToFitText: 'shape to fit text', textToFitShape: 'text to fit shape' }, frame.autoSize, 'text auto size')}`);
    if (frame.verticalAnchor) result.push(`set vertical anchor of (${f}) to ${enumValue({ top: 'anchor top', middle: 'anchor middle', center: 'anchor middle', bottom: 'anchor bottom' }, frame.verticalAnchor, 'vertical anchor')}`);
    if (frame.horizontalAnchor) result.push(`set horizontal anchor of (${f}) to ${enumValue({ left: 'horizontal anchor none', center: 'horizontal anchor center' }, frame.horizontalAnchor, 'horizontal anchor')}`);
  }
  for (const [i, paragraph] of (properties.paragraphs || []).entries()) {
    const range = `paragraph ${i + 1} of (${textRange})`, format = `paragraph format of (${range})`;
    if (paragraph.level !== undefined) result.push(`set indent level of (${range}) to ${Math.max(1, index(paragraph.level) + 1)}`);
    for (const [key, property] of [['spaceBefore', 'space before'], ['spaceAfter', 'space after'], ['spaceWithin', 'space within']]) if (paragraph[key] !== undefined) result.push(`set ${property} of (${format}) to ${number(paragraph[key])}`);
    if (paragraph.alignment) result.push(`set alignment of (${format}) to ${enumValue({ left: 'paragraph align left', center: 'paragraph align center', right: 'paragraph align right', justify: 'paragraph align justify' }, paragraph.alignment, 'paragraph alignment')}`);
    if (paragraph.bullet !== undefined) { const bullet = paragraph.bullet; const type = !bullet || bullet.type === 'none' ? 'bullet type none' : bullet.type === 'numbered' ? 'bullet type numbered' : 'bullet type unnumbered'; result.push(`set bullet type of bullet format of (${format}) to ${type}`); if (bullet?.character) result.push(`set bullet character of bullet format of (${format}) to ${asText(bullet.character)}`); if (bullet?.startAt) result.push(`set bullet start value of bullet format of (${format}) to ${index(bullet.startAt)}`); }
  }
  return result.filter(Boolean).join('\n');
}
function formattingScript(properties, target = 'targetShape', { nameHandled = false, rotationFirst = false } = {}) {
  const result = [];
  if (rotationFirst && properties.rotation !== undefined) result.push(`set rotation of ${target} to ${number(properties.rotation)}`);
  if (properties.width !== undefined || properties.height !== undefined) result.push(`set lock aspect ratio of ${target} to false`);
  for (const [key, property] of [['left', 'left position'], ['top', 'top'], ['width', 'width'], ['height', 'height']]) if (properties[key] !== undefined) result.push(`set ${property} of ${target} to ${number(properties[key])}`);
  if (properties.fill) {
    const fill = properties.fill, f = `fill format of ${target}`;
    if (fill.type === 'none') result.push(`set visible of (${f}) to false`);
    else {
      result.push(`set visible of (${f}) to true`);
      if (fill.type === 'gradient') {
        const stops = fill.stops || []; if (stops.length < 2) throw new Error('A gradient requires at least two stops.');
        result.push(`set fore color of (${f}) to ${rgb(stops[0].color)}`, `set back color of (${f}) to ${rgb(stops.at(-1).color)}`, `two color gradient (${f}) style horizontal gradient variant 1`);
        result.push(`repeat while (count of gradient stops of (${f})) > 2\ndelete gradient stop (${f}) stop index 3\nend repeat`);
        for (const [i, stop] of stops.entries()) {
          if (i < 2) result.push(`set color of gradient stop ${i + 1} of (${f}) to ${rgb(stop.color)}`, `set position of gradient stop ${i + 1} of (${f}) to ${number(stop.position)}`, `set transparency of gradient stop ${i + 1} of (${f}) to ${number(stop.transparency || 0)}`);
          else result.push(`insert gradient stop (${f}) custom color ${rgb(stop.color)} position ${number(stop.position)} transparency ${number(stop.transparency || 0)}`);
        }
      } else if (fill.type === 'picture') { if (!fill.path) throw new Error('Picture fills require an issued local path.'); result.push(`user picture (${f}) picture file ${asText(fill.path)}`); }
      else { result.push(`solid (${f})`); if (fill.color) result.push(`set fore color of (${f}) to ${rgb(fill.color)}`); }
      if (fill.transparency !== undefined) result.push(`set transparency of (${f}) to ${number(fill.transparency)}`);
    }
  }
  if (properties.line) {
    const line = properties.line, f = `line format of ${target}`;
    // The dictionary has no line.visible. Zero transparency is opaque; one hides it.
    if (line.visible === false) result.push(`set transparency of (${f}) to 1`);
    else {
      // A shape created without an outline ignores fore color until a line style is active.
      if (line.visible === true || line.color || line.width !== undefined || line.dashStyle) result.push(`set line style of (${f}) to single line`);
      if (line.transparency !== undefined || line.visible === true) result.push(`set transparency of (${f}) to ${number(line.transparency || 0)}`);
    }
    if (line.color) result.push(`set fore color of (${f}) to ${rgb(line.color)}`);
    if (line.width !== undefined) result.push(`set line weight of (${f}) to ${number(line.width)}`);
    if (line.dashStyle) result.push(`set dash style of (${f}) to ${enumValue({ solid: 'line dash style solid', dash: 'line dash style dash', dot: 'line dash style round dot', roundDot: 'line dash style round dot', squareDot: 'line dash style square dot', dashDot: 'line dash style dash dot', dashDotDot: 'line dash style dash dot dot', longDash: 'line dash style long dash', longDashDot: 'line dash style long dash dot', longDashDotDot: 'line dash style long dash dot dot', sysDash: 'line dash style system dash', sysDot: 'line dash style system dot', sysDashDot: 'line dash style system dash dot' }, line.dashStyle, 'line dash style')}`);
  }
  result.push(textScript(properties, target));
  if (properties.crop) {
    const crop = properties.crop, l = number(crop.left || 0), r = number(crop.right || 0), t = number(crop.top || 0), b = number(crop.bottom || 0);
    if (l + r >= 1 || t + b >= 1) throw new Error('Crop fractions must retain a visible image area.');
    // CropLeft/Right/Top/Bottom use original-image points, independent of current scaling.
    result.push(`set savedWidth to width of ${target}\nset savedHeight to height of ${target}\nset savedLeft to left position of ${target}\nset savedTop to top of ${target}`);
    result.push(`set probeShape to duplicate ${target}\nset crop left of picture format of probeShape to 0\nset crop right of picture format of probeShape to 0\nset crop top of picture format of probeShape to 0\nset crop bottom of picture format of probeShape to 0\nscale width probeShape factor 1 relative to original size true scale scale from top left\nscale height probeShape factor 1 relative to original size true scale scale from top left\nset originalWidth to width of probeShape\nset originalHeight to height of probeShape\ndelete probeShape`);
    for (const [key, fraction, dimension] of [['left', l, 'originalWidth'], ['right', r, 'originalWidth'], ['top', t, 'originalHeight'], ['bottom', b, 'originalHeight']]) result.push(`set crop ${key} of picture format of ${target} to (${dimension} * ${fraction})`);
    result.push(`set width of ${target} to savedWidth\nset height of ${target} to savedHeight\nset left position of ${target} to savedLeft\nset top of ${target} to savedTop`);
  }
  // PowerPoint invalidates newly created shape references on rotation/rename, so set them last. Creation
  // commands pass the name in the make properties and skip it here. Existing shapes reset rotation first so
  // the bounding-box offset does not leak into the geometry update.
  if (!rotationFirst && properties.rotation !== undefined) result.push(`set rotation of ${target} to ${number(properties.rotation)}`);
  if (!nameHandled && properties.name !== undefined) result.push(`set name of ${target} to ${asText(properties.name)}`);
  return result.filter(Boolean).join('\n');
}
async function textSelection(input) {
  const script = `set selectedText to {}\nif selection type of selection of targetWindow is selection type text then\nset r to text range of selection of targetWindow\nset f to font of r\nset selectedText to {(offset of r) - 1, text length of r, content of r, font name of f, font size of f, font color of f, bold of f, italic of f, underline of f, superscript of f, subscript of f, ASCII name of f, east asian name of f, transparency of f, emboss of f, base line offset of f}\nend if`;
  const row = await appleScript(`set selectedText to {}\n${office(script, input.presentationName)}\n${jsonReturn('selectedText')}`, { json: true });
  if (!row.length) return null;
  const color = '#' + row[5].map((v) => Math.round(v > 255 ? v / 257 : v).toString(16).padStart(2, '0')).join('').toUpperCase();
  return { shapeId: input.context?.selectionIds?.[0] ?? input.shapeId, start: Math.max(0, row[0]), length: row[1], text: row[2], font: { name: row[3], size: row[4], color, bold: row[6], italic: row[7], underline: row[8], superscript: row[9], subscript: row[10], nameAscii: row[11], nameFarEast: row[12], transparency: row[13], emboss: row[14], baselineOffset: row[15] } };
}
async function context(input) {
  const refs = refsFor(input), rows = [];
  const body = refs.map(({ id, ref }) => `set s to ${refExpression(ref)}\nset textValue to ""\nset fontValue to {}\nif has text frame of s then\nset textValue to content of text range of text frame of s\nset f to font of text range of text frame of s\nset fontValue to {font name of f, font size of f, font color of f, bold of f, italic of f, underline of f, superscript of f, subscript of f}\nend if\nset fillValue to {visible of fill format of s, fore color of fill format of s, transparency of fill format of s, (fill format type of fill format of s) as text}\nset lineValue to {fore color of line format of s, line weight of line format of s, transparency of line format of s, (dash style of line format of s) as text}\nset cropValue to {}\nif (shape type of s is shape type picture) or (shape type of s is shape type linked picture) then\nset cropValue to {crop left of picture format of s, crop top of picture format of s, crop right of picture format of s, crop bottom of picture format of s}\nend if\nset end of output to {${asText(id)}, name of s, left position of s, top of s, width of s, height of s, rotation of s, (shape type of s) as text, textValue, fontValue, fillValue, lineValue, cropValue}`).join('\n');
  const result = await appleScript(`set output to {}\n${office(body, input.presentationName)}\n${jsonReturn('output')}`, { json: true });
  const hex = (values) => '#' + values.map((v) => Math.round(v > 255 ? v / 257 : v).toString(16).padStart(2, '0')).join('').toUpperCase();
  for (const [i, row] of result.entries()) {
    const [id, name, left, top, width, height, rotation, nativeType, text, font, fill, line, crop] = row;
    const item = { id, ref: refs[i].ref, name, left, top, width, height, rotation, nativeType, text, fill: { type: fill[0] ? fill[3].includes('gradient') ? 'gradient' : fill[3].includes('picture') ? 'picture' : 'solid' : 'none', color: hex(fill[1]), transparency: fill[2] }, line: { color: hex(line[0]), width: line[1], transparency: line[2], dashStyle: line[3], visible: line[2] < 1 } };
    if (font.length) item.font = { name: font[0], size: font[1], color: hex(font[2]), bold: font[3], italic: font[4], underline: font[5], superscript: font[6], subscript: font[7] };
    if (crop.length) item.cropPoints = { left: crop[0], top: crop[1], right: crop[2], bottom: crop[3] };
    rows.push(item);
  }
  const selectedText = await textSelection(input);
  if (!input.context) return { shapes: rows, textSelection: selectedText };
  const updates = new Map(rows.map((s) => [s.id, s]));
  const merge = (shapes) => shapes.map((s) => ({ ...s, ...updates.get(s.id), ...(s.children ? { children: merge(s.children) } : {}) }));
  return { ...input.context, shapes: merge(input.context.shapes || []), textSelection: selectedText };
}
async function styleApply(style, ref, env, expected) {
  if (!style?.fragment) throw new Error('Full shape formatting requires style.fragment with an editable source PPTX.');
  if (!ref) throw new Error('Style application requires the target shape ref.');
  const slideIndex = index(ref.slideIndex), slide = `slide ${slideIndex + 1} of targetPresentation`;
  const before = Number(await appleScript(office(`return count of shapes of (${slide})`, expected)));
  await insertFragment({ fragment: style.fragment, presentationName: expected, context: {} }, env);
  const after = Number(await appleScript(office(`return count of shapes of (${slide})`, expected)));
  const styleShape = before + 1;
  if (after <= before) throw new Error('The style source could not be inserted for pick up.');
  const cleanup = `repeat with styleNumber from ${after} to ${styleShape} by -1
try
delete shape styleNumber of (${slide})
end try
end repeat`;
  try {
    // Cross-presentation `apply` cannot consume a shape reference from another presentation. Insert the
    // style source into the target slide, pick up its formatting there, then delete it.
    await appleScript(office(`set targetShape to ${refExpression(ref)}\nset savedStyleGeometry to {left position of targetShape, top of targetShape, width of targetShape, height of targetShape, rotation of targetShape}\nset savedStyleLock to lock aspect ratio of targetShape\nset styleShape to shape ${styleShape} of (${slide})\npick up styleShape\napply targetShape\nset lock aspect ratio of targetShape to false\nset left position of targetShape to item 1 of savedStyleGeometry\nset top of targetShape to item 2 of savedStyleGeometry\nset width of targetShape to item 3 of savedStyleGeometry\nset height of targetShape to item 4 of savedStyleGeometry\nset lock aspect ratio of targetShape to savedStyleLock\nset rotation of targetShape to item 5 of savedStyleGeometry\n${cleanup}`, expected));
  } catch (error) {
    await appleScript(office(cleanup, expected)).catch(() => {});
    throw error;
  }
}
async function mediaBox(command, filename) {
  const hasWidth = command.width !== undefined && command.width !== null;
  const hasHeight = command.height !== undefined && command.height !== null;
  const boxWidth = hasWidth ? number(command.width, 'Media width') : null;
  const boxHeight = hasHeight ? number(command.height, 'Media height') : null;
  if (hasWidth && boxWidth <= 0 || hasHeight && boxHeight <= 0) throw new Error('Media width and height must be positive.');
  let info = null;
  try { info = await helper('mediaInfo', { path: filename }); } catch (error) { info = null; }
  const intrinsicWidth = Number(info?.width) > 0 ? Number(info.width) : 0;
  const intrinsicHeight = Number(info?.height) > 0 ? Number(info.height) : 0;
  let width, height;
  if (intrinsicWidth && intrinsicHeight) {
    const scale = hasWidth && hasHeight ? Math.min(boxWidth / intrinsicWidth, boxHeight / intrinsicHeight) : null;
    if (scale !== null) { width = intrinsicWidth * scale; height = intrinsicHeight * scale; }
    else if (hasWidth) { width = boxWidth; height = boxWidth * intrinsicHeight / intrinsicWidth; }
    else if (hasHeight) { height = boxHeight; width = boxHeight * intrinsicWidth / intrinsicHeight; }
    else { width = intrinsicWidth; height = intrinsicHeight; }
  } else {
    if (!hasWidth || !hasHeight) throw new Error('PowerPoint could not read the video dimensions; supply both width and height for the media box.');
    width = boxWidth; height = boxHeight;
  }
  return {
    left: number(command.left, 'Media left') + (hasWidth ? (boxWidth - width) / 2 : 0),
    top: number(command.top, 'Media top') + (hasHeight ? (boxHeight - height) / 2 : 0),
    width, height,
  };
}
async function insertImageCommand(command, slide, env, expected) {
  const filename = await nativeAsset(command, env);
  return preservingClipboard(async () => {
    noteClipboardWrite(await helper('writeFileClipboard', { path: filename }));
    await appleScript(office(`set index of targetWindow to 1\nset originalSlideNumber to slide index of slide of view of targetWindow\ngo to slide (view of targetWindow) number (slide index of ${slide})\nset beforeCount to count of shapes of ${slide}\npaste object (view of targetWindow)\nif (count of shapes of ${slide}) <= beforeCount then error "PowerPoint did not paste the image from the clipboard." number 1\nset targetShape to last shape of ${slide}\nif (shape type of targetShape) is not shape type picture then error ("PowerPoint pasted the file as " & ((shape type of targetShape) as text) & " instead of an embedded picture.") number 1\n${formattingScript(command)}\ngo to slide (view of targetWindow) number originalSlideNumber`, expected));
    env.events?.emit?.('changed', { host: 'powerpoint' });
  });
}
async function documentInfo(input = {}) {
  const [name, rawPath, fullName, slideNumber] = await appleScript(`tell application "Microsoft PowerPoint"\nset targetPresentation to active presentation\n${input.presentationName ? `if (name of targetPresentation) is not ${asText(input.presentationName)} then error "The active PowerPoint presentation is not the requested document." number 1\n` : ''}set documentName to name of targetPresentation\nset documentPath to ""\nset documentFullName to documentName\ntry\nset documentPath to path of targetPresentation\nend try\ntry\nset documentFullName to full name of targetPresentation\nend try\nif documentPath is not "" and documentPath does not start with "/" then\ntry\nset documentPath to POSIX path of (documentPath as alias)\nend try\nend if\nset slideNumber to 0\ntry\nset slideNumber to slide index of slide of view of active window\nend try\nend tell\n${jsonReturn('{documentName, documentPath, documentFullName, slideNumber}')}`, { json: true });
  return { name, path: rawPath || null, fullName, slideIndex: slideNumber > 0 ? slideNumber - 1 : 0 };
}
async function apply(input, env) {
  const scripts = [];
  for (const command of input.commands || []) {
    const target = command.ref ? refExpression(command.ref) : null;
    const slide = command.slideIndex === undefined ? 'targetSlide' : `slide ${index(command.slideIndex) + 1} of targetPresentation`;
    switch (command.op) {
      case 'update': {
        if (!target) throw new Error('Update requires command.ref.');
        const properties = { ...(command.properties || {}) };
        if (properties.fill?.path) { const fillMime = mimeTypes[path.extname(properties.fill.path).toLowerCase()] || 'image/png'; properties.fill = { ...properties.fill, path: await nativeAsset({ path: properties.fill.path, mime: fillMime }, env) }; }
        scripts.push(`set targetShape to ${target}`);
        if (properties.style) { await styleApply(properties.style, command.ref, env, input.presentationName); delete properties.style; }
        scripts.push(formattingScript(properties, 'targetShape', { rotationFirst: true })); break;
      }
      case 'insertText': scripts.push(`set targetShape to make new text box at end of (${slide}) with properties {left position:${number(command.left)}, top:${number(command.top)}, width:${number(command.width)}, height:${number(command.height)}${command.name ? `, name:${asText(command.name)}` : ''}}`, formattingScript(command, 'targetShape', { nameHandled: true })); break;
      case 'insertShape': {
        if (command.shapeType === 'line') scripts.push(`set targetShape to make new line shape at end of (${slide}) with properties {begin line X:${number(command.left)}, begin line Y:${number(command.top)}, end line X:${number(command.left) + number(command.width)}, end line Y:${number(command.top) + number(command.height)}${command.name ? `, name:${asText(command.name)}` : ''}}`);
        else { const types = { rectangle: 'autoshape rectangle', ellipse: 'autoshape oval', oval: 'autoshape oval', roundedRectangle: 'autoshape rounded rectangle', triangle: 'autoshape isosceles triangle', rightArrow: 'autoshape right arrow', diamond: 'autoshape diamond', hexagon: 'autoshape hexagon', pentagon: 'autoshape regular pentagon', chevron: 'autoshape chevron', arc: 'autoshape arc', bracket: 'autoshape left bracket' }; scripts.push(`set targetShape to make new shape at end of (${slide}) with properties {auto shape type:${enumValue(types, command.shapeType || 'rectangle', 'shape type')}, left position:${number(command.left)}, top:${number(command.top)}, width:${number(command.width)}, height:${number(command.height)}${command.name ? `, name:${asText(command.name)}` : ''}}`); }
        scripts.push(formattingScript(command, 'targetShape', { nameHandled: true })); break;
      }
      case 'insertImage': {
        if (scripts.length) { await preservingClipboard(() => appleScript(office(scripts.join('\n'), input.presentationName))); scripts.length = 0; }
        await insertImageCommand(command, slide, env, input.presentationName); break;
      }
      case 'insertMedia': {
        const filename = await nativeAsset(command, env);
        const geometry = await mediaBox(command, filename);
        scripts.push(`set targetShape to make new media2 object at end of (${slide}) with properties {file name:${asText(filename)}${command.name ? `, name:${asText(command.name)}` : ''}}`, formattingScript({ ...command, ...geometry, rotation: command.rotation ?? 0 }, 'targetShape', { nameHandled: true })); break;
      }
      case 'duplicate': if (!target) throw new Error('Duplicate requires command.ref.'); scripts.push(`set targetShape to duplicate (${target})`, formattingScript({ ...command.properties, ...(command.name ? { name: command.name } : {}) })); break;
      case 'delete': {
        const refs = command.refs || (command.ref ? [command.ref] : []); if (!refs.length) throw new Error('Delete requires command.refs.');
        const sorted = [...refs].sort((a, b) => b.slideIndex - a.slideIndex || b.shapePath.length - a.shapePath.length || b.shapePath.at(-1) - a.shapePath.at(-1));
        scripts.push(...sorted.map((ref) => `delete (${refExpression(ref)})`)); break;
      }
      case 'group': scripts.push('if (selection type of selection of targetWindow) is not selection type shapes then error "Select shapes before grouping." number 1\nset targetShape to group (shape range of selection of targetWindow)', formattingScript({ name: command.name })); break;
      case 'ungroup': scripts.push('if (selection type of selection of targetWindow) is not selection type shapes then error "Select a group before ungrouping." number 1\nungroup (shape range of selection of targetWindow)'); break;
      case 'zOrder': {
        if (!target) throw new Error('zOrder requires command.ref.');
        if (command.position === 'index') {
          const position = index(command.index) + 1;
          scripts.push(`set targetShape to ${target}\nz order targetShape z order position send shape to back\nrepeat ${position - 1} times\nz order targetShape z order position bring shape forward\nend repeat`);
        } else scripts.push(`z order (${target}) z order position ${enumValue({ front: 'bring shape to front', back: 'send shape to back', forward: 'bring shape forward', backward: 'send shape backward' }, command.position, 'z-order position')}`);
        break;
      }
      case 'insertSlide': {
        const at = index(command.index ?? command.slideIndex ?? 0) + 1;
        const layout = enumValue({ blank: 'slide layout blank', title: 'slide layout title slide', titleOnly: 'slide layout title only', text: 'slide layout text slide', twoColumn: 'slide layout two column text' }, command.layout || 'blank', 'slide layout');
        scripts.push(`set createdSlide to make new slide at end of targetPresentation with properties {layout:${layout}}`, `if ${at} < (count of slides of targetPresentation) then move createdSlide to before slide ${at} of targetPresentation`); break;
      }
      case 'deleteSlides': scripts.push(...[...new Set(command.indices || [])].map(index).sort((a, b) => b - a).map((i) => `delete slide ${i + 1} of targetPresentation`)); break;
      case 'moveSlide': { const from = index(command.index), to = index(command.toIndex); if (from !== to) scripts.push(`move slide ${from + 1} of targetPresentation to ${from < to ? 'after' : 'before'} slide ${to + 1} of targetPresentation`); break; }
      case 'insertFragment': {
        if (scripts.length) { await preservingClipboard(() => appleScript(office(scripts.join('\n'), input.presentationName))); scripts.length = 0; }
        await insertFragment({ ...command, context: input.context, presentationName: input.presentationName }, env); break;
      }
      case 'select': throw new Error('Select shapes with Office.js before native group or ungroup. PowerPoint does not expose an AppleScript select command.');
      default: throw new Error(`Unknown native command: ${command.op}`);
    }
  }
  if (scripts.length) await preservingClipboard(() => appleScript(office(scripts.join('\n'), input.presentationName)));
  env.events?.emit?.('changed', { host: 'powerpoint' });
  return input.context && input.refs ? context(input) : { applied: (input.commands || []).length };
}
async function waitForFile(filename, timeout = 10000) {
  const deadline = Date.now() + timeout;
  let previous = -1;
  while (Date.now() < deadline) {
    try { const { size } = await stat(filename); if (size > 0 && size === previous) return; previous = size; } catch (error) { if (error.code !== 'ENOENT') throw error; }
    await new Promise((resolve) => setTimeout(resolve, 100));
  }
  throw new Error('PowerPoint did not finish writing the exported file.');
}
async function temporaryPath(extension, host = 'powerpoint') { const directory = hostContainer(host); await mkdir(directory, { recursive: true, mode: 0o700 }); return path.join(directory, `${randomUUID()}.${extension}`); }
async function descriptor(filename, mime) { issuedFiles.add(await realpath(filename)); return { path: filename, name: path.basename(filename), mime, data: (await readFile(filename)).toString('base64') }; }
async function insertFragment(input, env) {
  const fragment = input.fragment || input;
  const filename = await nativeAsset({ data: fragment.pptx, path: fragment.path, mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', name: 'fragment.pptx' }, env);
  const bounds = fragment.bounds || input.bounds;
  const options = { ...(input.options || {}), ...Object.fromEntries(['left', 'top', 'width', 'height', 'center'].filter((key) => input[key] !== undefined).map((key) => [key, input[key]])) };
  const current = input.context || {};
  const sx = options.width !== undefined && bounds?.width ? number(options.width) / bounds.width : options.height !== undefined && bounds?.height ? number(options.height) / bounds.height : 1;
  const sy = options.height !== undefined && bounds?.height ? number(options.height) / bounds.height : sx;
  const left = options.left ?? (options.center && bounds && current.width ? (current.width - bounds.width * sx) / 2 : bounds?.left ?? 0);
  const top = options.top ?? (options.center && bounds && current.height ? (current.height - bounds.height * sy) / 2 : bounds?.top ?? 0);
  // Capture geometry before PowerPoint applies its duplicate-paste offset.
  const sourceGeometry = bounds ? 'set sourceLeft to left position of sourceShape\nset sourceTop to top of sourceShape\nset sourceWidth to width of sourceShape\nset sourceHeight to height of sourceShape' : '';
  const transform = bounds ? `set lock aspect ratio of insertedShape to false\nset width of insertedShape to sourceWidth * ${sx}\nset height of insertedShape to sourceHeight * ${sy}\nset left position of insertedShape to ${number(left)} + (sourceLeft - ${number(bounds.left)}) * ${sx}\nset top of insertedShape to ${number(top)} + (sourceTop - ${number(bounds.top)}) * ${sy}` : '';
  return preservingClipboard(async () => {
    // Copying from another presentation requires closing it before pasting: while the source is open,
    // PowerPoint keeps it as the active presentation even after its window is brought forward, so the
    // target slide cannot be addressed. The clipboard survives closing the source.
    const body = `set sourcePresentation to missing value\nset targetName to name of targetPresentation\nset targetSlideNumber to slide index of targetSlide\nset originalCount to count of shapes of targetSlide\nset sourceCount to 0\ntry\n${openFile(filename)}\nset sourcePresentation to active presentation\nset sourceCount to count of shapes of slide 1 of sourcePresentation\nclose sourcePresentation saving no\nset sourcePresentation to missing value\n${waitForPresentation('targetName')}\nrepeat with shapeNumber from 1 to sourceCount\n${openFile(filename)}\nset sourcePresentation to active presentation\nset sourceShape to shape shapeNumber of slide 1 of sourcePresentation\n${sourceGeometry}\ncopy shape sourceShape\nclose sourcePresentation saving no\nset sourcePresentation to missing value\n${waitForPresentation('targetName')}\ngo to slide (view of active window) number targetSlideNumber\npaste object (view of active window)\nset insertedShape to last shape of slide targetSlideNumber of presentation targetName\n${transform}\nend repeat\nset targetPresentation to presentation targetName\nset targetSlide to slide targetSlideNumber of targetPresentation\non error errorText number errorNumber\nif sourcePresentation is not missing value then close sourcePresentation saving no\nset targetPresentation to presentation targetName\nset targetSlide to slide targetSlideNumber of targetPresentation\nerror errorText number errorNumber\nend try\nreturn ((count of shapes of targetSlide) - originalCount)`;
    const count = Number(await appleScript(office(body, input.presentationName))); env.events?.emit?.('changed', { host: 'powerpoint' }); return { inserted: count };
  });
}
async function renderPresentationPDF(filename, expected) {
  const pdf = await temporaryPath('pdf');
  await appleScript(office(`set exportSource to missing value\nset targetName to name of targetPresentation\ntry\n${openFile(filename)}\nset exportSource to active presentation\nsave exportSource in (POSIX file ${asText(pdf)}) as save as PDF\nclose exportSource saving no\nset exportSource to missing value\n${waitForPresentation('targetName')}\nset index of targetWindow to 1\ngo to slide (view of targetWindow) number (slide index of targetSlide)\non error errorText number errorNumber\nif exportSource is not missing value then close exportSource saving no\nerror errorText number errorNumber\nend try`, expected));
  await waitForFile(pdf);
  return pdf;
}
async function independentPDF(filename) {
  const pdf = await temporaryPath('pdf');
  await appleScript(`tell application "Microsoft PowerPoint"\nset previousWindow to missing value\nset previousSlideNumber to 0\nset previousName to ""\ntry\nset previousWindow to active window\nset previousName to name of active presentation\nset previousSlideNumber to slide index of slide of view of previousWindow\non error\nset previousWindow to missing value\nend try\nset exportSource to missing value\ntry\n${openFile(filename)}\nset exportSource to active presentation\nsave exportSource in (POSIX file ${asText(pdf)}) as save as PDF\nclose exportSource saving no\nset exportSource to missing value\nif previousName is not "" then\n${waitForPresentation('previousName')}\nend if\nif previousWindow is not missing value then\nset index of previousWindow to 1\ngo to slide (view of previousWindow) number previousSlideNumber\nend if\non error errorText number errorNumber\nif exportSource is not missing value then close exportSource saving no\nerror errorText number errorNumber\nend try\nend tell`);
  await waitForFile(pdf);
  return pdf;
}
async function mediaInfo(input, env) {
  if (!input?.path) throw new Error('A media file path is required.');
  return helper('mediaInfo', { path: await authorizeFile(input.path, env) });
}
async function renderPicturePDF(filename) {
  const pdf = await temporaryPath('pdf');
  const geometry = await preservingClipboard(async () => {
    noteClipboardWrite(await helper('writeFileClipboard', { path: filename }));
    return appleScript(`tell application "Microsoft PowerPoint"\nset previousWindow to missing value\nset previousSlideNumber to 0\nset previousName to ""\ntry\nset previousWindow to active window\nset previousName to name of active presentation\nset previousSlideNumber to slide index of slide of view of previousWindow\non error\nset previousWindow to missing value\nend try\nset renderPresentation to missing value\nset renderWidth to 0\nset renderHeight to 0\ntry\nset renderPresentation to make new presentation\nrepeat while (count of slides of renderPresentation) > 0\ndelete slide 1 of renderPresentation\nend repeat\nmake new slide at end of renderPresentation with properties {layout:slide layout blank}\nset renderSlide to slide 1 of renderPresentation\nrepeat while (count of shapes of renderSlide) > 0\ndelete shape 1 of renderSlide\nend repeat\nset renderWindow to document window 1 of renderPresentation\nset renderView to view of renderWindow\nset index of renderWindow to 1\ngo to slide renderView number 1\npaste object renderView\nif (count of shapes of renderSlide) is 0 then error "PowerPoint did not paste the picture." number 1\nset renderPicture to last shape of renderSlide\nif (shape type of renderPicture) is not shape type picture then error ("PowerPoint pasted the file as " & ((shape type of renderPicture) as text) & " instead of a picture.") number 1\nset renderWidth to width of renderPicture\nset renderHeight to height of renderPicture\nset lock aspect ratio of renderPicture to false\nset renderFactor to 1\nif renderWidth > 684 then set renderFactor to 684 / renderWidth\nif renderHeight > 378 and (378 / renderHeight) < renderFactor then set renderFactor to 378 / renderHeight\nset renderWidth to renderWidth * renderFactor\nset renderHeight to renderHeight * renderFactor\nset left position of renderPicture to 0\nset top of renderPicture to 0\nset width of renderPicture to renderWidth\nset height of renderPicture to renderHeight\nsave renderPresentation in (POSIX file ${asText(pdf)}) as save as PDF\nclose renderPresentation saving no\nset renderPresentation to missing value\nif previousWindow is not missing value then\nset index of previousWindow to 1\ngo to slide (view of previousWindow) number previousSlideNumber\nend if\non error errorText number errorNumber\nif renderPresentation is not missing value then close renderPresentation saving no\nerror errorText number errorNumber\nend try\nend tell\nreturn ((renderWidth as text) & "," & (renderHeight as text))`);
  });
  const [width, height] = String(geometry).split(',').map(Number);
  if (!(width > 0) || !(height > 0)) throw new Error('PowerPoint could not render the source picture.');
  return { pdf, bounds: { left: 0, top: 0, width, height } };
}
async function rasterizeImage(input, env) {
  const requested = input.maximum ?? input.maxDimension ?? 4096;
  const maximum = number(requested, 'Maximum raster dimension');
  if (!Number.isInteger(maximum) || maximum < 1 || maximum > 32768) throw new Error('Maximum raster dimension must be an integer between 1 and 32768.');
  const fragment = input.fragment || (input.pptx ? input : {});
  const temporary = [];
  try {
    let pdf, bounds;
    if (fragment.pptx || fragment.path) {
      const data = fragment.pptx || (await readFile(await authorizeFile(fragment.path, env))).toString('base64');
      const stored = await materialize({ data, mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', name: 'raster-source.pptx', host: 'powerpoint' });
      temporary.push(stored.path);
      bounds = input.bounds || fragment.bounds;
      if (!bounds) throw new Error('The picture fragment must include bounds so the raster keeps only the native picture.');
      pdf = await independentPDF(stored.path);
      temporary.push(pdf);
    } else if (input.path) {
      const filename = await nativeAsset(input, env);
      const rendered = await renderPicturePDF(filename);
      pdf = rendered.pdf; bounds = rendered.bounds; temporary.push(pdf);
    } else {
      throw new Error('rasterizeImage requires a normalized picture fragment (input.fragment) or an issued image path.');
    }
    const left = number(bounds.left, 'Picture left');
    const top = number(bounds.top, 'Picture top');
    const width = number(bounds.width, 'Picture width');
    const height = number(bounds.height, 'Picture height');
    if (!(width > 0) || !(height > 0)) throw new Error('Picture bounds must have a positive size.');
    const [page] = await helper('rasterPDF', { path: pdf, indices: [0], maximum, format: 'png', bounds: { left, top, width, height } });
    return { ...page, bounds: { left, top, width, height }, rotation: 0 };
  } finally {
    for (const filename of temporary) await discardCached(filename);
  }
}
async function fragment(input, env = {}) {
  const refs = refsFor(input);
  if (input.fragment?.pptx || input.pptx) {
    const packaged = input.fragment?.pptx ? input.fragment : await env.pptx.extractFragment(input.pptx, { slideIndex: input.slideIndex ?? input.context?.slideIndex ?? refs[0]?.ref.slideIndex ?? 0, ids: input.ids || refs.map(({ id }) => id) });
    const filename = await nativeAsset({ data: packaged.pptx, mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', name: 'fragment.pptx' }, env);
    if (input.preview === false) return { ...packaged, path: filename };
    const pdf = await renderPresentationPDF(filename, input.presentationName);
    const [preview] = await helper('rasterPDF', { path: pdf, indices: [0], dpi: 100, format: 'png', bounds: packaged.bounds });
    return { ...packaged, path: filename, preview: preview.data };
  }
  if (!refs.length) throw new Error('Fragment capture requires resolved shape refs.');
  const filename = await temporaryPath('pptx'), pdf = await temporaryPath('pdf');
  await preservingClipboard(async () => {
    const copy = refs.map(({ ref }) => `copy shape (${refExpression(ref)})\npaste object (view of document window 1 of capturePresentation)`).join('\n');
    const body = `set capturePresentation to missing value\ntry\nset capturePresentation to make new presentation\nif (count of slides of capturePresentation) is 0 then make new slide at end of capturePresentation with properties {layout:slide layout blank}\nset captureSlide to slide 1 of capturePresentation\nrepeat while (count of shapes of captureSlide) > 0\ndelete shape 1 of captureSlide\nend repeat\n${copy}\nsave capturePresentation in (POSIX file ${asText(filename)}) as save as Open XML presentation\nsave capturePresentation in (POSIX file ${asText(pdf)}) as save as PDF\nclose capturePresentation saving no\nset capturePresentation to missing value\nset index of targetWindow to 1\ngo to slide (view of targetWindow) number (slide index of targetSlide)\non error errorText number errorNumber\nif capturePresentation is not missing value then close capturePresentation saving no\nerror errorText number errorNumber\nend try`;
    await appleScript(office(body, input.presentationName));
  });
  const result = await descriptor(filename, 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
  const bounds = input.bounds || boundsFromContext(input.context, refs.map(({ id }) => id));
  if (input.preview === false) return { pptx: result.data, path: filename, shapeIds: refs.map(({ id }) => id), bounds, pdfPath: pdf };
  const [preview] = await helper('rasterPDF', { path: pdf, indices: [0], dpi: 100, format: 'png', bounds });
  return { pptx: result.data, path: filename, shapeIds: refs.map(({ id }) => id), preview: preview.data, bounds };
}
function boundsFromContext(context, ids) { const selected = new Set(ids), shapes = []; const visit = (list) => { for (const shape of list || []) { if (selected.has(shape.id)) shapes.push(shape); if (shape.children) visit(shape.children); } }; visit(context?.shapes); if (!shapes.length) return undefined; const left = Math.min(...shapes.map((s) => s.left)), top = Math.min(...shapes.map((s) => s.top)); return { left, top, width: Math.max(...shapes.map((s) => s.left + s.width)) - left, height: Math.max(...shapes.map((s) => s.top + s.height)) - top }; }
async function exportPresentationFiles(input, env, scratch) {
  const format = input.format || 'png';
  if (input.kind === 'original-images') {
    if (!input.pptx) throw new Error('Original image export requires the active-slide PPTX snapshot.');
    const images = await env.pptx.extractImages(input.pptx, { slideIndex: input.slideIndex ?? 0, ids: input.ids, original: true });
    const result = [];
    for (const image of images) { const file = await materialize({ data: image.data, mime: image.mime, name: image.name || `image-${image.id}${extensions[image.mime] || '.bin'}` }); result.push({ ...image, ...file }); }
    return result;
  }
  if (format === 'svg') throw new Error('PowerPoint exposes SVG for embedded vector images only; export original-images with the active-slide PPTX snapshot.');
  let filename, selectionBounds;
  if (input.kind === 'selection') {
    const captured = await fragment({ ...input, preview: false }, env); selectionBounds = captured.bounds;
    scratch.push(captured.path, captured.pdfPath);
    if (format === 'pptx') return [{ name: input.filename || 'selection.pptx', mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', data: captured.pptx, path: captured.path }];
    filename = captured.pdfPath || await renderPresentationPDF(captured.path, input.presentationName);
    scratch.push(filename);
  } else if (input.pptx) {
    const data = input.indices?.length ? await env.pptx.selectSlides(input.pptx, input.indices) : input.pptx;
    const packaged = await materialize({ data, mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', name: 'slides.pptx' });
    scratch.push(packaged.path);
    if (format === 'pptx') return [{ ...await descriptor(packaged.path, 'application/vnd.openxmlformats-officedocument.presentationml.presentation'), name: input.filename || 'slides.pptx' }];
    filename = await renderPresentationPDF(packaged.path, input.presentationName);
    scratch.push(filename);
  } else {
    const temporary = await temporaryPath('pptx');
    scratch.push(temporary);
    await preservingClipboard(async () => {
      const body = `set exportCopy to missing value\ntry\nset exportCopy to make new presentation\nrepeat while (count of slides of exportCopy) > 0\ndelete slide 1 of exportCopy\nend repeat\n` + (input.indices?.length ? input.indices.map((i) => `copy object (slide ${index(i) + 1} of targetPresentation)\npaste object exportCopy`).join('\n') : `repeat with sourceSlide in slides of targetPresentation\ncopy object sourceSlide\npaste object exportCopy\nend repeat`) + `\nsave exportCopy in (POSIX file ${asText(temporary)}) as save as Open XML presentation\nclose exportCopy saving no\nset exportCopy to missing value\ngo to slide (view of targetWindow) number (slide index of targetSlide)\non error errorText number errorNumber\nif exportCopy is not missing value then close exportCopy saving no\nerror errorText number errorNumber\nend try`;
      await appleScript(office(body, input.presentationName));
    });
    if (format === 'pptx') return [{ ...await descriptor(temporary, 'application/vnd.openxmlformats-officedocument.presentationml.presentation'), name: input.filename || 'slides.pptx' }];
    filename = await renderPresentationPDF(temporary, input.presentationName);
    scratch.push(filename);
  }
  if (format === 'pdf') {
    if (selectionBounds) { const cropped = await helper('slicePDF', { path: filename, bounds: selectionBounds }); const file = await materialize({ ...cropped, name: input.filename || 'selection.pdf' }); return [{ ...cropped, ...file }]; }
    if (input.separatePdf) {
      const pages = await helper('slicePDF', { path: filename, separate: true });
      const stem = safeName(input.filename || 'slides').replace(/\.pdf$/i, '');
      const files = [];
      for (const [index, page] of pages.entries()) {
        const name = `${stem}_页面${(input.indices?.[index] ?? index) + 1}.pdf`;
        files.push({ ...page, ...await materialize({ ...page, name }), name });
      }
      return files;
    }
    return [{ ...await descriptor(filename, 'application/pdf'), name: input.filename || 'slides.pdf' }];
  }
  if (!['png', 'jpg', 'bmp'].includes(format)) throw new Error(`Unknown export format: ${format}`);
  const pages = await helper('rasterPDF', { path: filename, dpi: input.dpi || 300, format, bounds: selectionBounds });
  const result = [];
  for (const [i, page] of pages.entries()) { const name = input.filename && pages.length === 1 ? safeName(input.filename) : `slide-${(input.indices?.[i] ?? i) + 1}.${format}`; const materialized = await materialize({ ...page, name }); result.push({ ...page, ...materialized, name }); }
  return result;
}
// Intermediate copies (packaged deck, rendered PDF, captured fragment) are removed as soon as
// the files handed back to the caller exist; returned files stay until the cache sweep.
async function exportPresentation(input, env) {
  const scratch = [];
  let result;
  try { result = await exportPresentationFiles(input, env, scratch); return result; }
  finally {
    const kept = new Set((result || []).map(file => file?.path && path.resolve(file.path)).filter(Boolean));
    for (const file of new Set(scratch.filter(Boolean).map(item => path.resolve(item)))) if (!kept.has(file)) await discardCached(file).catch(() => {});
  }
}
async function finishExport(input, env) {
  const result = await exportPresentation(input, env);
  for (const file of result) if (input.folder) {
    const saved = env.files?.write ? await env.files.write({ folder: input.folder, name: file.name, data: file.data, mime: file.mime }) : await writeIssued({ folder: input.folder, name: file.name, data: file.data });
    file.path = typeof saved === 'string' ? saved : saved.path;
  }
  if (input.copyToClipboard && result.length) noteClipboardWrite(await helper('writeClipboard', result[0]));
  return result;
}
async function nativeOSAction(input, env) {
  switch (input.action) {
    case 'temporaryPath': {
      const extension = String(input.extension || 'bin').replace(/^\./, '');
      if (!/^[a-zA-Z0-9]{1,12}$/.test(extension)) throw new Error('Invalid temporary file extension.');
      const filename = await temporaryPath(extension, input.host || 'powerpoint'); reservedFiles.add(path.resolve(filename)); return { path: filename };
    }
    case 'releaseFiles': {
      // Compare both the spelled and the symlink-resolved forms: issued files are recorded by
      // realpath, reserved ones by their spelled path, and a home folder may sit behind a symlink.
      const spelled = path.resolve(hostContainer(input.host || 'powerpoint'));
      const cache = await realpath(spelled).catch(() => spelled);
      for (const filename of input.paths || []) {
        if (typeof filename !== 'string' || !path.isAbsolute(filename)) throw new Error('Only issued files in the host cache may be released.');
        const plain = path.resolve(filename), actual = await realpath(plain).catch(() => plain);
        const inside = [plain, actual].some(value => value.startsWith(cache + path.sep) || value.startsWith(spelled + path.sep));
        const owned = [plain, actual].some(value => reservedFiles.has(value) || issuedFiles.has(value));
        if (!inside || !owned) throw new Error('Only issued files in the host cache may be released.');
        await rm(actual, { recursive: true, force: true });
        for (const value of [plain, actual]) { reservedFiles.delete(value); issuedFiles.delete(value); }
      }
      return {};
    }
    case 'snapshotClipboard': clipboardOperationDepth++; try { return { items: await helper('snapshotClipboard') }; } catch (error) { clipboardOperationDepth--; throw error; }
    case 'restoreClipboard': try { return noteClipboardWrite(await helper('restoreClipboard', { items: input.items })); } finally { clipboardOperationDepth = Math.max(0, clipboardOperationDepth - 1); }
    case 'readFile': return readIssued(input, env);
    case 'mediaInfo': return mediaInfo(input, env);
    case 'rasterPDF': return helper('rasterPDF', { ...input, path: await authorizeFile(input.path, env) });
    case 'slicePDF': return helper('slicePDF', { ...input, path: await authorizeFile(input.path, env) });
    default: throw new Error(`Unknown native OS action: ${input.action}`);
  }
}
async function execute(input, env = {}) {
  if (['temporaryPath', 'releaseFiles', 'snapshotClipboard', 'restoreClipboard', 'readFile', 'rasterPDF', 'slicePDF', 'mediaInfo'].includes(input.action)) return nativeOSAction(input, env);
  if (input.host && input.host !== 'powerpoint' && input.action !== 'rasterizeImage') throw new Error('WPS document editing must execute in the WPS Application frontend. Native filesystem services are shared.');
  if (input.action === 'rasterizeImage') { const raster = officeQueue.then(() => rasterizeImage(input, env), () => rasterizeImage(input, env)); officeQueue = raster.catch(() => {}); return raster; }
  const task = async () => { switch (input.action) { case 'documentInfo': return documentInfo(input); case 'context': return context(input); case 'textSelection': return textSelection(input); case 'apply': return apply(input, env); case 'insertFragment': return insertFragment(input, env); case 'fragment': return fragment(input, env); case 'export': return finishExport(input, env); default: throw new Error(`Unknown PowerPoint native action: ${input.action}`); } };
  const result = officeQueue.then(task, task); officeQueue = result.catch(() => {}); return result;
}
async function watchStart(input = {}, env = {}) {
  const id = randomUUID(), child = spawn(await helperBinary(), [], { stdio: ['pipe', 'pipe', 'pipe'] });
  let buffer = '', errors = '', readySettled = false, readyResolve, readyReject;
  const ready = new Promise((resolve, reject) => { readyResolve = resolve; readyReject = reject; });
  const deadline = setTimeout(() => { child.kill(); settleReady(new Error('Native media watcher did not start.')); }, 15000);
  function settleReady(error, value) { if (readySettled) return; readySettled = true; clearTimeout(deadline); if (error) readyReject(error); else readyResolve(value); }
  child.stdin.end(JSON.stringify({ ...input, action: 'watch' }));
  child.stderr.on('data', (chunk) => { if (errors.length < 4096) errors += chunk.toString('utf8'); });
  child.on('error', (error) => settleReady(error));
  child.stdout.on('data', (chunk) => { buffer += chunk.toString('utf8'); let newline; while ((newline = buffer.indexOf('\n')) >= 0) { const line = buffer.slice(0, newline); buffer = buffer.slice(newline + 1); try { const event = JSON.parse(line); if (event.type === 'ready') settleReady(null, { id, ...event }); else if (!shouldSuppressWatchEvent(event)) { for (const file of event.files || []) issuedFiles.add(path.resolve(file.path)); env.events?.emit?.(event.type === 'nativeError' ? 'nativeError' : event.type === 'hostChanged' ? 'changed' : 'mediaSource', event); } } catch (error) { env.events?.emit?.('nativeError', { message: error.message }); } } });
  child.on('close', (code) => { watches.delete(id); if (!readySettled) settleReady(new Error(errors || 'Native watcher exited before start.')); else if (code && !child.slideSciStopping) env.events?.emit?.('nativeError', { message: errors || 'Native watcher exited.' }); });
  watches.set(id, child); return ready;
}
async function watchStop({ id } = {}) { const stop = (child) => { child.slideSciStopping = true; child.kill('SIGTERM'); }; if (id) { const child = watches.get(id); if (child) stop(child); watches.delete(id); } else { for (const child of watches.values()) stop(child); watches.clear(); } return {}; }
export const nativeHandlers = {
  execute, pickFiles, chooseFolder,
  // Selecting a file in Finder discloses nothing to the caller, so a source path recorded in an
  // older session (or on this Mac before a restart) can be shown without a fresh grant.
  reveal: async (input) => { if (typeof input.path !== 'string' || !path.isAbsolute(input.path) || /[\x00-\x1f]/.test(input.path)) throw new Error('没有可显示的本机文件位置。'); const filename = await realpath(input.path).catch(() => { throw new Error('原始文件已不存在或已移动：' + input.path); }); return helper('reveal', { path: filename }); },
  fonts: async () => helper('fonts'),
  readClipboardFiles: async () => { const result = await helper('readClipboardFiles'); for (const file of result) issuedFiles.add(await realpath(file.path)); return result; },
  writeClipboard: async (input) => noteClipboardWrite(await helper('writeClipboard', input)),
  keychainGet: async ({ account }) => helper('keychainGet', { account, service: 'SlideSCI.AI' }),
  keychainSet: async ({ account, secret }) => { if (typeof secret !== 'string') throw new Error('A keychain secret string is required.'); return helper('keychainSet', { account, secret, service: 'SlideSCI.AI' }); },
  keychainDelete: async ({ account }) => helper('keychainDelete', { account, service: 'SlideSCI.AI' }),
  materialize: input => materialize(input),
  read: async (input, env) => env?.files?.read ? env.files.read(input.path) : readIssued(input, env),
  write: async (input, env) => env?.files?.write ? env.files.write(input) : writeIssued(input),
  mediaInfo,
  normalizeTIFF: async (input, env) => helper('normalizeTIFF', { path: await authorizeFile(input.path, env) }),
  watchStart, watchStop,
};
