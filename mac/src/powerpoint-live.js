// Live PowerPoint data for shapes parsed from an export of the same deck.
// The exported package cannot see geometry/fonts inherited from a layout or
// master, so every value read here comes from the real Office.js object model.
// This module only reads: it queues loads and `getSubstring` font loads, then
// mutates the supplied normalized shape tree. It never changes the document.

const SHAPE_PROPERTIES = 'id,name,left,top,width,height,rotation,altTextDescription,type';
const FILL_PROPERTIES = 'type,foregroundColor,transparency';
const LINE_PROPERTIES = 'visible,color,weight,dashStyle,transparency';
const TAG_PROPERTIES = 'key,value';
const FONT_PROPERTIES = 'name,size,color,bold,italic,underline,strikethrough,subscript,superscript';
const TEXT_PROPERTIES = 'text,' + FONT_PROPERTIES.split(',').map(name => `font/${name}`).join(',');

const fillTypes = {
  NoFill: 'none', Solid: 'solid', Gradient: 'gradient', Pattern: 'pattern',
  PictureAndTexture: 'picture', SlideBackground: 'background',
};
const dashStyles = {
  Solid: 'solid', Dash: 'dash', DashDot: 'dashDot', DashDotDot: 'dashDotDot',
  LongDash: 'longDash', LongDashDot: 'longDashDot', LongDashDotDot: 'longDashDotDot',
  RoundDot: 'roundDot', SquareDot: 'squareDot',
  SystemDash: 'sysDash', SystemDot: 'sysDot', SystemDashDot: 'sysDashDot',
  SystemDashDotDot: 'sysDashDotDot',
};
const withoutTextFrame = new Set(['Group', 'Image', 'Media', 'Table']);
const fontFlags = ['bold', 'italic', 'strikethrough', 'subscript', 'superscript'];

function color(value) {
  const text = String(value);
  return /^#[0-9a-f]{6}$/i.test(text) ? text.toUpperCase() : text;
}
function font(native, inherited) {
  const result = { ...(inherited || {}) };
  if (!native) return result;
  if (typeof native.name === 'string' && native.name) result.name = native.name;
  if (Number.isFinite(native.size) && native.size > 0) result.size = native.size;
  if (typeof native.color === 'string' && native.color) result.color = color(native.color);
  for (const key of fontFlags) if (typeof native[key] === 'boolean') result[key] = native[key];
  if (native.subscript === false && native.superscript === false) delete result.baselineOffset;
  if (typeof native.underline === 'string') {
    result.underline = native.underline !== 'None';
    if (result.underline) result.underlineStyle = native.underline;
    else delete result.underlineStyle;
  }
  return result;
}
function fill(native, parsed) {
  const type = fillTypes[String(native?.type)];
  if (!type) return parsed && typeof parsed === 'object' ? parsed : { type: 'none', transparency: 1 };
  if (type === 'none') return { type, transparency: 1 };
  const result = { ...(parsed?.type === type ? parsed : {}), type };
  if ((type === 'solid' || type === 'pattern') && typeof native.foregroundColor === 'string' && native.foregroundColor) {
    result.color = color(native.foregroundColor);
  }
  if (Number.isFinite(native.transparency)) result.transparency = native.transparency;
  return result;
}
function line(native, parsed) {
  if (!native) return parsed && typeof parsed === 'object' ? parsed : { visible: false, width: 0 };
  const result = { ...(parsed && typeof parsed === 'object' ? parsed : {}) };
  if (typeof native.visible === 'boolean') result.visible = native.visible;
  if (typeof native.color === 'string' && native.color) result.color = color(native.color);
  if (Number.isFinite(native.weight)) result.width = native.weight;
  if (Number.isFinite(native.transparency)) result.transparency = native.transparency;
  if (typeof native.dashStyle === 'string' && dashStyles[native.dashStyle]) result.dashStyle = dashStyles[native.dashStyle];
  return result;
}
function tags(native) {
  return Object.fromEntries((native?.items || []).map(tag => [String(tag.key), String(tag.value ?? '')]));
}

function nativeOffsetMap(value) {
  const text = String(value ?? ''), offsets = [0];
  for (let index = 0; index < text.length;) {
    const length = text[index] === '\r' && text[index + 1] === '\n' ? 2 : 1;
    index += length;
    offsets.push(index);
  }
  return offsets;
}

function parsedText(value) {
  return String(value ?? '').replace(/\r\n?/g, '\n').replace(/\v/g, '\n');
}
function nativeTextCandidates(value) {
  const raw = String(value ?? '');
  // Office usually terminates the final paragraph with a carriage return, but
  // some builds return separators only. Try the conventional stripped form and
  // the raw form so run offsets stay aligned with whichever form matches.
  let stripped = raw;
  if (raw.endsWith('\r\n')) stripped = raw.slice(0, -2);
  else if (raw.endsWith('\r')) stripped = raw.slice(0, -1);
  const normalize = text => text.replace(/\r\n?/g, '\n').replace(/\v/g, '\n');
  const result = [normalize(stripped), normalize(raw)];
  return result[0] === result[1] ? result.slice(0, 1) : result;
}

function target(model, collection, depth, parentLiveId) {
  const path = model?.ref?.shapePath;
  if (!Array.isArray(path) || path.length !== depth + 1) throw new Error('PowerPoint 对象引用无效，请刷新后重试。');
  const index = Number(path[depth]);
  if (!Number.isInteger(index) || index < 1) throw new Error('PowerPoint 对象引用无效，请刷新后重试。');
  return { model, collection, depth, index: index - 1, parentLiveId };
}
function readShape(entry) {
  const { model, shape, parentLiveId } = entry;
  if (model.packageId == null) model.packageId = model.id == null ? '' : String(model.id);
  model.id = String(shape.id);
  if (parentLiveId !== undefined) model.parentId = parentLiveId;
  if (typeof shape.name === 'string') model.name = shape.name;
  for (const key of ['left', 'top', 'width', 'height', 'rotation']) {
    const value = shape[key];
    if (typeof value === 'number' && Number.isFinite(value)) model[key] = value;
  }
  model.alternativeText = shape.altTextDescription == null ? '' : String(shape.altTextDescription);
  model.tags = tags(shape.tags);
  model.fill = fill(shape.fill, model.fill);
  model.line = line(shape.lineFormat, model.line);
  if (model.style && typeof model.style === 'object') {
    model.style.fill = model.fill;
    model.style.line = model.line;
  }
}

/**
 * Reads the live Office.js values for parsed slides from the same presentation.
 * Shapes are located with `{slideIndex, shapePath}` 1-based refs. Each shape is
 * mutated in place: the original parsed id becomes `packageId`, `id` and
 * `parentId` become live Office ids, and geometry, name, alt text, tags,
 * fill/line, text/font, and per-run fonts are refreshed. Returns `slides`.
 */
export async function enrichPowerPointSlides(slides) {
  if (!Array.isArray(slides)) throw new Error('幻灯片数据无效。');
  if (!slides.length) return slides;
  const host = globalThis.PowerPoint;
  if (!host?.run) throw new Error('PowerPoint 接口不可用。');
  const jobs = [];
  for (const slide of slides) {
    const index = Number(slide?.index);
    if (!Number.isInteger(index) || index < 0) throw new Error('幻灯片索引无效。');
    const shapes = Array.isArray(slide.shapes) ? slide.shapes : [];
    if (shapes.length) jobs.push({ index, shapes });
  }
  if (!jobs.length) return slides;

  await host.run(async context => {
    const presentationSlides = context.presentation.slides;
    const resolved = [];
    let level = [];
    for (const job of jobs) {
      const collection = presentationSlides.getItemAt(job.index).shapes;
      for (const model of job.shapes) level.push(target(model, collection, 0));
    }
    while (level.length) {
      const pending = level.map(entry => {
        const shape = entry.collection.getItemAt(entry.index);
        shape.load(SHAPE_PROPERTIES);
        shape.fill.load(FILL_PROPERTIES);
        shape.lineFormat.load(LINE_PROPERTIES);
        shape.tags.load(TAG_PROPERTIES);
        return { ...entry, shape };
      });
      await context.sync();
      const next = [];
      for (const entry of pending) {
        readShape(entry);
        resolved.push(entry);
        const children = entry.model.children;
        if (!children?.length) continue;
        if (String(entry.shape.type) !== 'Group') throw new Error('PowerPoint 对象层级已改变，请刷新后重试。');
        const childCollection = entry.shape.group.shapes;
        const depth = entry.depth + 1;
        const parentLiveId = String(entry.shape.id);
        for (const child of children) next.push(target(child, childCollection, depth, parentLiveId));
      }
      level = next;
    }

    const ranges = [];
    for (const entry of resolved) {
      if (entry.model.text === undefined || withoutTextFrame.has(String(entry.shape.type))) continue;
      const range = entry.shape.textFrame.textRange;
      range.load(TEXT_PROPERTIES);
      ranges.push({ model: entry.model, range, text: parsedText(entry.model.text) });
    }
    if (!ranges.length) return;
    await context.sync();

    const runs = [];
    for (const entry of ranges) {
      const candidates = nativeTextCandidates(entry.range.text);
      const text = candidates.find(candidate => candidate === entry.text) ?? candidates[0];
      entry.model.text = text;
      entry.model.font = font(entry.range.font, entry.model.font);
      if (text !== entry.text) continue;
      const source = Array.isArray(entry.model.runs) ? entry.model.runs : [];
      if (!source.length) continue;
      let offset = 0, combined = '';
      const spans = [];
      for (const run of source) {
        const value = parsedText(run.text);
        combined += value;
        if (value.length) spans.push({ run, start: offset, length: value.length });
        offset += value.length;
      }
      if (combined !== entry.text) continue;
      const offsets = nativeOffsetMap(entry.range.text).slice(0, entry.text.length + 1);
      if (offsets.length !== entry.text.length + 1) continue;
      entry.pending = spans.map(item => {
        const start = offsets[item.start], end = offsets[item.start + item.length];
        const range = entry.range.getSubstring(start, end - start);
        range.font.load(FONT_PROPERTIES);
        return { ...item, range };
      });
      runs.push(entry);
    }
    if (!runs.length) return;
    await context.sync();
    for (const entry of runs) {
      const enriched = new Map();
      for (const item of entry.pending) {
        item.run.font = font(item.range.font, item.run.font);
        enriched.set(item.run, item.run.font);
      }
      for (const paragraph of entry.model.paragraphs || []) {
        for (const run of paragraph.runs || []) if (enriched.has(run)) run.font = enriched.get(run);
      }
    }
  });
  return slides;
}
