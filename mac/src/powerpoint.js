import { enrichPowerPointSlides } from './powerpoint-live.js';
import { alignment } from './features/editing-geometry.js';
import { ensureOfficeReady } from './ribbon.js';
const flatten = shapes => shapes.flatMap(shape => [shape, ...flatten(shape.children || [])]);
const directProperties = new Set(['left', 'top', 'width', 'height', 'rotation', 'name', 'tags', 'alternativeText']);
const geometryProperties = new Set(['left', 'top', 'width', 'height', 'rotation']);
const simpleFontKeys = new Set(['name', 'nameFarEast', 'nameAscii', 'size', 'bold', 'italic', 'color', 'underline', 'strikethrough', 'subscript', 'superscript', 'allCaps']);
const simpleTextFrameKeys = new Set(['wordWrap', 'autoSize', 'marginLeft', 'marginRight', 'marginTop', 'marginBottom']);
const autoSizeValues = { none: 'AutoSizeNone', shapeToFitText: 'AutoSizeShapeToFitText', textToFitShape: 'AutoSizeTextToFitShape' };
const paragraphAlignValues = { left: 'Left', center: 'Center', right: 'Right', justify: 'Justify' };
const geometryShapeLoad = 'items/id,items/name,items/left,items/top,items/width,items/height,items/rotation,items/type';
const shapeTypes = { Group: 'group', Image: 'picture', Media: 'media', Table: 'table', TextBox: 'text' };
const textTypes = new Set(['TextBox', 'Placeholder', 'GeometricShape', 'Callout', 'Freeform']);
// Office returns a path or URL; a file name such as "增长 20%.pptx" is not valid URI encoding.
// Office.js reports GeneralException while PowerPoint is briefly busy, for example right after an
// AppleScript paste. Reads are safe to repeat; writes are never retried.
export async function readPowerPoint(operation, attempts = 3) {
  for (let attempt = 1; ; attempt++) {
    try { return await PowerPoint.run(operation); }
    catch (error) {
      const busy = error?.code === 'GeneralException' || error?.message === 'GeneralException';
      if (!busy || attempt >= attempts) throw error;
      await new Promise(resolve => setTimeout(resolve, 250 * attempt));
    }
  }
}
function documentName(url) {
  const last = String(url || '').split('/').at(-1) || '';
  try { return decodeURIComponent(last); } catch { return last; }
}

export async function connectPowerPoint(services) {
  const info = await ensureOfficeReady();
  if (info.host !== Office.HostType.PowerPoint) throw new Error('请从 PowerPoint 加载此插件。');
  if (!Office.context.requirements.isSetSupported('PowerPointApi', '1.10')) {
    throw new Error('此版本 PowerPoint 缺少完整文档与图片接口，请更新 Microsoft 365 PowerPoint 后重试（需要 PowerPointApi 1.10）。');
  }
  const request = (route, input) => {
    if (!services?.request) throw new Error('此高级功能需要 SlideSCI 本机服务，请确认服务正在运行后重试。');
    return services.request(route, input);
  };
  let boundTitle = '';
  const native = (action, input = {}) => {
    const presentationName = Office.context.document.url ? documentName(Office.context.document.url) : boundTitle;
    return request('native/execute', { host: 'powerpoint', action, ...input,
      ...(presentationName && action !== 'rasterizeImage' ? { presentationName } : {}) });
  };
  let queue = Promise.resolve(), timer, stopped = false, surface, notifying = false, pending = 0;
  let selectionSubscribed = false, selectionEventSupported = false, visibilitySubscribed = false;
  const listeners = new Map();
  function serial(operation) {
    pending++;
    const run = async () => { try { return await operation(); } finally { pending--; } };
    const result = queue.then(run, run);
    queue = result.catch(() => {});
    return result;
  }
  function title() {
    const url = Office.context.document.url;
    return url ? documentName(url).replace(/\.pptx?$/i, '') : boundTitle;
  }
  function officeShape(context, shape) {
    let collection = context.presentation.slides.getItemAt(shape.ref.slideIndex).shapes;
    let result;
    for (const [depth, index] of shape.ref.shapePath.entries()) {
      result = collection.getItemAt(index - 1);
      if (depth + 1 < shape.ref.shapePath.length) collection = result.group.shapes;
    }
    return result;
  }
  function metadata(shape, values) {
    if (values.alternativeText !== undefined) shape.altTextDescription = String(values.alternativeText);
    for (const [key, value] of Object.entries(values.tags || {})) {
      if (value === '') shape.tags.delete(key);
      else if (value !== undefined && value !== null) shape.tags.add(key, String(value));
    }
  }
  // Several thumbnails may be selected while the edited shapes sit on one of them. Office.js names
  // that slide through the selected shapes; AppleScript is only a fallback, and core layout keeps
  // working on the first selected slide when the local service is unavailable.
  async function parentSlideIndex(context, slides, selection) {
    const first = selection?.items?.[0];
    if (typeof first?.getParentSlideOrNullObject !== 'function') return -1;
    const parent = first.getParentSlideOrNullObject();
    parent.load('id');
    await context.sync();
    return parent.isNullObject ? -1 : slides.items.findIndex(item => item.id === parent.id);
  }
  async function activeSlideIndex(context, slides, selected, selection) {
    const indices = selected.items.map(slide => slides.items.findIndex(item => item.id === slide.id));
    let slideIndex = indices[0] ?? 0;
    if (indices.length <= 1) return { indices, slideIndex };
    const index = await parentSlideIndex(context, slides, selection);
    if (index >= 0) return { indices, slideIndex: index };
    try {
      const active = await native('documentInfo');
      if (indices.includes(active.slideIndex)) slideIndex = active.slideIndex;
    } catch { /* The first selected slide stays the target. */ }
    return { indices, slideIndex };
  }
  function targetSlideIndex(presentation, slides, target) {
    const index = slides.items.findIndex(slide => String(slide.id) === String(target.slideId));
    if (String(presentation.id) !== String(target.presentationId) || index < 0) {
      throw new Error('找不到操作开始时的演示文稿或幻灯片。');
    }
    return index;
  }
  async function quick(options = {}) {
    return readPowerPoint(async context => {
      const presentation = context.presentation;
      const slides = presentation.slides;
      const selected = presentation.getSelectedSlides();
      const selection = presentation.getSelectedShapes();
      const count = slides.getCount();
      selection.load('items/id');
      presentation.load('id');
      presentation.pageSetup.load('slideWidth,slideHeight');
      let firstSlide;
      if (options.slides === false) {
        selected.load('items/id,items/index');
        firstSlide = slides.getItemAt(0);
        firstSlide.load('id');
      } else {
        slides.load('items/id');
        selected.load('items/id');
      }
      await context.sync();
      if (!count.value) throw new Error('请先在当前演示文稿中添加幻灯片。');
      let indices, slideIndex, slideId;
      if (options.slides === false) {
        indices = selected.items.map(slide => slide.index);
        slideIndex = indices[0] ?? 0;
        slideId = selected.items[0]?.id ?? firstSlide.id;
        // Agree with the geometry context, or the live panel would see a slide change on every event.
        const first = selection.items[0];
        if (indices.length > 1 && typeof first?.getParentSlideOrNullObject === 'function') {
          const parent = first.getParentSlideOrNullObject();
          parent.load('id,index');
          await context.sync();
          if (!parent.isNullObject) { slideIndex = parent.index; slideId = parent.id; }
        }
      } else {
        ({ indices, slideIndex } = await activeSlideIndex(context, slides, selected, selection));
        slideId = slides.items[slideIndex].id;
      }
      return { presentationId: presentation.id, title: title(), slideIndex, slideId,
        width: presentation.pageSetup.slideWidth, height: presentation.pageSetup.slideHeight,
        selectionIds: selection.items.map(shape => String(shape.id)), selectedSlideIndices: indices,
        slideCount: count.value, slides: options.slides === false ? [] : slides.items.map((slide, index) => ({ id: slide.id, index })) };
    });
  }
  async function readSlideTree(context, slideIndex, load = geometryShapeLoad, withZOrder = false, withTags = false, slide) {
    const shapes = [];
    const tagged = withTags ? [] : null;
    let level = [{ collection: (slide || context.presentation.slides.getItemAt(slideIndex)).shapes, path: [], output: shapes }];
    while (level.length) {
      for (const entry of level) entry.collection.load(load);
      await context.sync();
      const next = [];
      for (const entry of level) {
        for (let index = 0; index < entry.collection.items.length; index++) {
          const shape = entry.collection.items[index];
          const path = [...entry.path, index + 1];
          const nativeType = String(shape.type);
          const model = { id: String(shape.id), nativeType, ref: { slideIndex, shapePath: path },
            left: shape.left, top: shape.top, width: shape.width, height: shape.height,
            rotation: shape.rotation, name: shape.name, type: shapeTypes[nativeType] || 'shape' };
          if (withZOrder) model.zOrder = Number(shape.zOrderPosition);
          if (tagged) {
            const tags = shape.tags;
            tags.load('key,value');
            tagged.push([model, tags]);
          }
          if (entry.parentId !== undefined) model.parentId = entry.parentId;
          entry.output.push(model);
          if (nativeType === 'Group') {
            model.children = [];
            next.push({ collection: shape.group.shapes, path, parentId: model.id, output: model.children });
          }
        }
      }
      level = next;
    }
    if (tagged?.length) {
      await context.sync();
      for (const [model, tags] of tagged) model.tags = Object.fromEntries(tags.items.map(({ key, value }) => [key, value]));
    }
    return shapes;
  }
  // 选中对象的几何子树：包含选中对象、其祖先组与全部后代；整页只在 scope:'slide' 时返回。
  function selectSubtrees(shapes, wanted) {
    const walk = (list, inherited) => list.flatMap(shape => {
      const chosen = inherited || wanted.has(shape.id);
      const children = walk(shape.children || [], chosen);
      if (!chosen && !children.length) return [];
      const copy = { ...shape };
      if (children.length) copy.children = children;
      else delete copy.children;
      return [copy];
    });
    return walk(shapes, false);
  }
  async function geometryContext(options = {}, extraIds = []) {
    return readPowerPoint(async context => {
      const presentation = context.presentation;
      const slides = presentation.slides;
      const selectedSlides = presentation.getSelectedSlides();
      const selection = presentation.getSelectedShapes();
      slides.load('items/id');
      selectedSlides.load('items/id');
      selection.load('items/id');
      presentation.load('id');
      presentation.pageSetup.load('slideWidth,slideHeight');
      await context.sync();
      if (!slides.items.length) throw new Error('请先在当前演示文稿中添加幻灯片。');
      // A target names its slide, so the active slide is only looked up when none is given.
      const active = options.target
        ? { indices: selectedSlides.items.map(slide => slides.items.findIndex(item => item.id === slide.id)) }
        : await activeSlideIndex(context, slides, selectedSlides, selection);
      const { indices } = active;
      const slideIndex = options.target ? targetSlideIndex(presentation, slides, options.target) : active.slideIndex;
      const selectionIds = selection.items.map(shape => String(shape.id));
      const wanted = new Set([...selectionIds, ...extraIds.map(String)]);
      let shapes = [];
      if (options.scope === 'slide' || wanted.size) {
        const slide = options.target ? slides.getItem(String(options.target.slideId)) : undefined;
        const tree = await readSlideTree(context, slideIndex, geometryShapeLoad, false, options.tags === true, slide);
        shapes = options.scope === 'slide' ? tree : selectSubtrees(tree, wanted);
      }
      // Text boxes are typed 'text'; placeholders and autoshapes report whether they hold text, so
      // labels drawn as rectangles or placeholders still count without a full PPTX export.
      if (options.textMeta) {
        const candidates = flatten(shapes).filter(shape => textTypes.has(shape.nativeType) && shape.nativeType !== 'TextBox');
        const mark = (shape, frame) => {
          if (shape.nativeType === 'Placeholder') {
            shape.isPlaceholder = true;
            if (frame.hasText !== undefined) shape.hasTextFrame = true;
          }
          shape.hasText = frame.hasText === true;
        };
        const read = shape => { const frame = officeShape(context, shape).textFrame; frame.load('hasText'); return frame; };
        const frames = [];
        for (const shape of candidates) {
          try { frames.push([shape, read(shape)]); } catch { /* no text frame */ }
        }
        let batched = !frames.length;
        if (frames.length) {
          try { await context.sync(); batched = true; } catch { /* a picture/chart/table placeholder has no text frame */ }
        }
        if (batched) for (const [shape, frame] of frames) mark(shape, frame);
        else for (const [shape] of frames) {
          try { const frame = read(shape); await context.sync(); mark(shape, frame); }
          catch { /* this shape has no text frame */ }
        }
      }
      return { presentationId: presentation.id, title: title(), slideId: slides.items[slideIndex].id, slideIndex,
        width: presentation.pageSetup.slideWidth, height: presentation.pageSetup.slideHeight, geometry: true,
        selectionIds, selectedSlideIndices: indices, shapes };
    });
  }
  function isSimpleFont(font) {
    if (!font || typeof font !== 'object' || Array.isArray(font)) return false;
    return Object.keys(font).every(key => simpleFontKeys.has(key));
  }
  function isSimpleTextFrame(frame) {
    if (!frame || typeof frame !== 'object' || Array.isArray(frame)) return false;
    if (!Object.keys(frame).every(key => simpleTextFrameKeys.has(key))) return false;
    return frame.autoSize === undefined || Boolean(autoSizeValues[frame.autoSize]);
  }
  function isSimpleParagraphs(paragraphs, text) {
    if (!Array.isArray(paragraphs) || paragraphs.length !== 1) return false;
    const [first] = paragraphs;
    if (!first || typeof first !== 'object') return false;
    for (const key of Object.keys(first)) {
      if (!['runs', 'alignment'].includes(key)) return false;
    }
    if (first.alignment !== undefined && !paragraphAlignValues[first.alignment]) return false;
    if (first.runs !== undefined) {
      if (!Array.isArray(first.runs) || first.runs.some(run => run.font !== undefined)) return false;
      if (first.runs.map(run => String(run.text ?? '')).join('') !== String(text ?? '')) return false;
    }
    return true;
  }
  function isSimpleInsertText(command) {
    if (command.op !== 'insertText' || command.slideIndex !== undefined) return false;
    if (command.fill !== undefined || command.line !== undefined || command.runs !== undefined) return false;
    if (command.font !== undefined && !isSimpleFont(command.font)) return false;
    if (command.textFrame !== undefined && !isSimpleTextFrame(command.textFrame)) return false;
    if (command.paragraphs !== undefined && !isSimpleParagraphs(command.paragraphs, command.text)) return false;
    return true;
  }
  function isSimpleTextUpdate(command) {
    if (command.op !== 'update' || command.slideIndex !== undefined || !command.properties) return false;
    const keys = Object.keys(command.properties);
    if (!keys.length) return false;
    let hasText = false;
    for (const key of keys) {
      if (directProperties.has(key)) continue;
      if (key === 'text') { hasText = true; continue; }
      if (key === 'font' && isSimpleFont(command.properties.font)) { hasText = true; continue; }
      return false;
    }
    return hasText;
  }
  function isSimpleOfficeBatch(commands) {
    return commands.every(command => command.slideIndex === undefined && (
      geometryUpdate(command) || isSimpleTextUpdate(command) || isSimpleInsertText(command)
      || command.op === 'select' || command.op === 'group' || command.op === 'delete'
    ));
  }
  function applySimpleFont(target, font) {
    if (!font) return;
    const name = font.name ?? font.nameFarEast ?? font.nameAscii;
    if (name !== undefined) target.name = String(name);
    if (font.size !== undefined) target.size = Number(font.size);
    if (font.bold !== undefined) target.bold = Boolean(font.bold);
    if (font.italic !== undefined) target.italic = Boolean(font.italic);
    if (font.color !== undefined) target.color = String(font.color);
    if (font.underline !== undefined) target.underline = font.underline ? 'Single' : 'None';
    if (font.strikethrough !== undefined) target.strikethrough = Boolean(font.strikethrough);
    if (font.subscript !== undefined) target.subscript = Boolean(font.subscript);
    if (font.superscript !== undefined) target.superscript = Boolean(font.superscript);
    if (font.allCaps !== undefined) target.allCaps = Boolean(font.allCaps);
  }
  async function applySimpleOffice(commands, readback = true, target) {
    const applied = await PowerPoint.run(async context => {
      const slides = context.presentation.slides;
      const selectedSlides = context.presentation.getSelectedSlides();
      const selection = target ? null : context.presentation.getSelectedShapes();
      slides.load('items/id');
      selectedSlides.load('items/id');
      selection?.load('items/id');
      if (target) context.presentation.load('id');
      await context.sync();
      if (!slides.items.length) throw new Error('请先在当前演示文稿中添加幻灯片。');
      const indices = selectedSlides.items.map(slide => slides.items.findIndex(item => item.id === slide.id));
      let slideIndex = target ? targetSlideIndex(context.presentation, slides, target) : indices[0] ?? 0;
      if (!target && indices.length > 1) {
        // Several thumbnails selected: edit the slide holding the selected shapes, or fall back.
        slideIndex = await parentSlideIndex(context, slides, selection);
        if (slideIndex < 0) return false;
      }
      if (target && commands.some(command => ['insertText', 'group', 'select'].includes(command.op))
        && (indices.length !== 1 || indices[0] !== slideIndex)) throw new Error('请返回开始操作时的幻灯片。');
      const slide = target ? slides.getItem(String(target.slideId)) : slides.getItemAt(slideIndex);
      const needsLookup = commands.some(command => command.op === 'update' || command.op === 'delete');
      const tree = needsLookup ? await readSlideTree(context, slideIndex, geometryShapeLoad, false, false, slide) : [];
      const byId = new Map(flatten(tree).map(shape => [String(shape.id), shape]));
      function targetShape(model) {
        if (!target) return officeShape(context, model);
        if (!model.parentId) return slide.shapes.getItem(model.id);
        return targetShape(byId.get(String(model.parentId))).group.shapes.getItem(model.id);
      }
      for (const command of commands) {
        if (command.op === 'insertText') {
          const created = slide.shapes.addTextBox(String(command.text ?? ''), {
            left: Number(command.left ?? 0), top: Number(command.top ?? 0),
            width: Number(command.width ?? 100), height: Number(command.height ?? 24)
          });
          if (command.name !== undefined) created.name = String(command.name);
          if (command.rotation !== undefined) created.rotation = Number(command.rotation);
          metadata(created, command);
          const frame = created.textFrame;
          if (command.textFrame) {
            if (command.textFrame.wordWrap !== undefined) frame.wordWrap = Boolean(command.textFrame.wordWrap);
            if (command.textFrame.autoSize && autoSizeValues[command.textFrame.autoSize]) frame.autoSizeSetting = autoSizeValues[command.textFrame.autoSize];
            if (command.textFrame.marginLeft !== undefined) frame.leftMargin = Number(command.textFrame.marginLeft);
            if (command.textFrame.marginRight !== undefined) frame.rightMargin = Number(command.textFrame.marginRight);
            if (command.textFrame.marginTop !== undefined) frame.topMargin = Number(command.textFrame.marginTop);
            if (command.textFrame.marginBottom !== undefined) frame.bottomMargin = Number(command.textFrame.marginBottom);
          }
          const range = frame.textRange;
          applySimpleFont(range.font, command.font);
          const alignment = command.paragraphs?.[0]?.alignment;
          if (alignment && paragraphAlignValues[alignment]) range.paragraphFormat.horizontalAlignment = paragraphAlignValues[alignment];
        } else if (command.op === 'update') {
          const found = byId.get(String(command.id));
          if (!found) return false;
          const shape = officeShape(context, found);
          const values = command.properties || {};
          for (const key of ['left', 'top', 'width', 'height', 'rotation', 'name']) {
            if (values[key] !== undefined) shape[key] = values[key];
          }
          metadata(shape, values);
          if (values.text !== undefined) shape.textFrame.textRange.text = String(values.text);
          if (values.font) applySimpleFont(shape.textFrame.textRange.font, values.font);
        } else if (command.op === 'group') {
          const ids = (command.ids || []).map(String);
          slide.setSelectedShapes(ids);
          await context.sync();
          const selection = context.presentation.getSelectedShapes();
          if (target) {
            selectedSlides.load('items/id');
            selection.load('items/id');
            await context.sync();
            if (selectedSlides.items.length !== 1 || String(selectedSlides.items[0].id) !== String(target.slideId)
              || selection.items.length !== ids.length || selection.items.some(shape => !ids.includes(String(shape.id)))) {
              throw new Error('编组前选区已变化，请返回原幻灯片重试。');
            }
          }
          const grouped = selection.group();
          if (command.name !== undefined) grouped.name = String(command.name);
          metadata(grouped, command);
        } else if (command.op === 'delete') {
          const ids = new Set((command.ids || (command.id !== undefined ? [command.id] : [])).map(String));
          const insideSelected = shape => {
            for (let parent = shape.parentId; parent; parent = byId.get(String(parent))?.parentId) if (ids.has(String(parent))) return true;
            return false;
          };
          for (const id of ids) {
            const found = byId.get(id);
            if (!found) return false;
            if (!insideSelected(found)) targetShape(found).delete();
          }
        } else if (command.op === 'select') {
          const ids = (command.ids || (command.id !== undefined ? [command.id] : [])).map(String);
          slide.setSelectedShapes(ids);
        }
      }
      await context.sync();
      return true;
    });
    if (!applied) return null;
    if (!readback) { emit(); return { applied: commands.length }; }
    const model = await geometryContext({ scope: 'slide' });
    emit();
    return model;
  }
  function geometryUpdate(command) {
    if (command.op !== 'update' || !command.properties) return false;
    const keys = Object.keys(command.properties);
    return keys.length > 0 && keys.every(key => geometryProperties.has(key));
  }
  async function applyGeometry(commands, movedIds, readback = true) {
    const applied = await PowerPoint.run(async context => {
      const slides = context.presentation.slides;
      const selectedSlides = context.presentation.getSelectedSlides();
      const selection = context.presentation.getSelectedShapes();
      slides.load('items/id');
      selectedSlides.load('items/id');
      selection.load('items/id');
      await context.sync();
      if (!slides.items.length) throw new Error('请先在当前演示文稿中添加幻灯片。');
      const indices = selectedSlides.items.map(slide => slides.items.findIndex(item => item.id === slide.id));
      let slideIndex = indices[0] ?? 0;
      if (indices.length > 1) {
        slideIndex = await parentSlideIndex(context, slides, selection);
        if (slideIndex < 0) return false;
      }
      const targets = new Map(selection.items.map(item => [String(item.id), item]));
      if (!commands.every(command => command.op !== 'update' || targets.has(String(command.id)))) return false;
      let selectedIds;
      for (const command of commands) {
        if (command.op === 'select') { selectedIds = (command.ids || (command.id !== undefined ? [command.id] : [])).map(String); continue; }
        const shape = targets.get(String(command.id));
        for (const key of geometryProperties) if (command.properties[key] !== undefined) shape[key] = command.properties[key];
      }
      if (selectedIds) slides.getItemAt(slideIndex).setSelectedShapes(selectedIds);
      await context.sync();
      return true;
    });
    if (!applied) return null;
    if (!readback) { emit(); return { applied: commands.length }; }
    const model = await geometryContext({ scope: 'selection' }, [...movedIds]);
    emit();
    return model;
  }
  async function align(edge) {
    const count = await PowerPoint.run(async context => {
      const presentation = context.presentation;
      const selection = presentation.getSelectedShapes();
      selection.load('items/id,items/left,items/top,items/width,items/height');
      await context.sync();
      if (!selection.items.length) throw new Error('请先选择至少 1 个形状。');
      const shapes = selection.items.map(item => ({ id: String(item.id), left: item.left, top: item.top, width: item.width, height: item.height }));
      let page;
      if (shapes.length === 1) {
        presentation.pageSetup.load('slideWidth,slideHeight');
        await context.sync();
        page = { width: presentation.pageSetup.slideWidth, height: presentation.pageSetup.slideHeight };
      }
      const commands = alignment(page, shapes, edge);
      const targets = new Map(selection.items.map(item => [String(item.id), item]));
      for (const command of commands) {
        const shape = targets.get(String(command.id));
        if (command.properties.left !== undefined) shape.left = command.properties.left;
        if (command.properties.top !== undefined) shape.top = command.properties.top;
      }
      await context.sync();
      return shapes.length;
    });
    emit();
    return { message: `已对齐 ${count} 个对象。` };
  }
  async function snapshot(options = {}) {
    const selection = await quick();
    const all = options.scope === 'presentation';
    const data = await readPowerPoint(async context => {
      const slides = context.presentation.slides;
      const exported = slides.exportAsBase64Presentation(all ? selection.slides.map(slide => slide.id) : [selection.slideId]);
      await context.sync();
      return exported.value;
    });
    const parsed = await request('pptx/parse', { pptx: data });
    const packageSlideIndex = all ? selection.slideIndex : 0;
    if (!all) {
      parsed.slides[0].index = selection.slideIndex;
      for (const shape of flatten(parsed.slides[0].shapes)) shape.ref.slideIndex = selection.slideIndex;
    }
    await enrichPowerPointSlides(parsed.slides);
    const active = parsed.slides.find(slide => slide.index === selection.slideIndex);
    if (!active) throw new Error('当前幻灯片已改变，请重试。');
    const model = { ...selection, width: parsed.width, height: parsed.height, shapes: active.shapes,
      slides: all ? parsed.slides : selection.slides };
    // The change watch never needs the text cursor. Reading it runs AppleScript against the active
    // window, which fails for every other open deck and would spawn a process every few seconds.
    if (!options.watch) {
      model.textSelection = await native('textSelection', { context: { selectionIds: selection.selectionIds } });
      if (model.textSelection && !model.textSelection.shapeId) await PowerPoint.run(async context => {
        const shape = context.presentation.getSelectedTextRange().getParentTextFrame().getParentShape();
        shape.load('id'); await context.sync();
        model.textSelection.shapeId = shape.id;
      });
    }
    if (options.images) {
      const images = await request('pptx/images', { pptx: data, slideIndex: packageSlideIndex });
      const byId = new Map(images.map(image => [image.id, image]));
      for (const shape of flatten(model.shapes)) if (byId.has(shape.packageId || shape.id)) shape.image = { ...shape.image, ...byId.get(shape.packageId || shape.id) };
    }
    return { context: model, pptx: data, packageSlideIndex };
  }
  function find(model, id, slideIndex = model.slideIndex) {
    const shapes = slideIndex === model.slideIndex ? model.shapes : model.slides?.find(slide => slide.index === slideIndex)?.shapes;
    const result = flatten(shapes || []).find(shape => String(shape.id) === String(id));
    if (!result) throw new Error('对象已改变或被删除，请重新选择。');
    return result;
  }
  async function select(model, ids, slideIndex = model.slideIndex) {
    await PowerPoint.run(async context => {
      context.presentation.slides.getItemAt(slideIndex).setSelectedShapes(ids.map(String));
      await context.sync();
    });
  }
  async function applyDirectUpdates(model, commands) {
    const scripted = commands.filter(command => Object.keys(command.properties || {}).some(key => !directProperties.has(key)));
    if (scripted.length) await native('apply', { commands: scripted.map(command => ({ ...command,
      ref: find(model, command.id, command.slideIndex).ref })) });
    await PowerPoint.run(async context => {
      for (const command of commands) {
        const values = command.properties || {}, shape = officeShape(context, find(model, command.id, command.slideIndex));
        if (!scripted.includes(command)) for (const key of ['left', 'top', 'width', 'height', 'rotation', 'name']) {
          if (values[key] !== undefined) shape[key] = values[key];
        }
        metadata(shape, values);
      }
      await context.sync();
    });
  }
  async function applyUpdates(model, commands, aliases) {
    const richFields = new Set(['font', 'fontScope', 'runs', 'paragraphs', 'textRange']);
    const portable = [];
    const direct = [];
    for (const command of commands) {
      const ordinary = {}, rich = {};
      for (const [key, value] of Object.entries(command.properties || {})) {
        (richFields.has(key) ? rich : ordinary)[key] = value;
      }
      if (rich.paragraphs) ordinary.paragraphs = rich.paragraphs.map(({ runs, text, font, ...layout }) => layout);
      if (rich.font?.emboss !== undefined && !rich.font.raw) {
        ordinary.font = { emboss: rich.font.emboss };
        rich.font = { ...rich.font }; delete rich.font.emboss;
      }
      if (rich.textRange?.font?.emboss !== undefined && !rich.textRange.font.raw) {
        ordinary.textRange = { ...rich.textRange, font: { emboss: rich.textRange.font.emboss } };
        delete ordinary.textRange.text;
        rich.textRange = { ...rich.textRange, font: { ...rich.textRange.font } };
        delete rich.textRange.font.emboss;
      }
      if (rich.font?.reflection && !rich.font.raw && !rich.font.reflection.raw) {
        ordinary.font = { ...ordinary.font, reflection: rich.font.reflection };
        rich.font = { ...rich.font }; delete rich.font.reflection;
      }
      if (Object.keys(ordinary).length) direct.push({ ...command, properties: ordinary });
      if (Object.keys(rich).some(key => key !== 'fontScope')) portable.push({ ...command, properties: rich });
    }
    if (direct.length) await applyDirectUpdates(model, direct);
    for (const command of portable) {
      const slideIndex = command.slideIndex ?? model.slideIndex;
      let id = String(command.id);
      while (aliases.has(slideIndex + ':' + id)) id = aliases.get(slideIndex + ':' + id);
      await replaceObject(model, { ...command, id, updates: [{ id, properties: command.properties }] }, aliases, 'formatText');
    }
  }
  async function replaceObject(model, command, aliases, operation = 'replacePicture') {
    const slideIndex = command.slideIndex ?? model.slideIndex;
    const source = await snapshot({ scope: slideIndex === model.slideIndex ? 'slide' : 'presentation' });
    const original = find(source.context, command.id, slideIndex);
    let root = original;
    while (root.parentId) root = find(source.context, root.parentId, slideIndex);
    const marker = 'SlideSCI_replace_' + crypto.randomUUID();
    const fragment = await request('pptx/' + operation, { pptx: source.pptx,
      slideIndex: slideIndex === source.context.slideIndex ? source.packageSlideIndex : slideIndex,
      id: original.packageId || original.id, rootId: root.packageId || root.id, rootName: marker,
      data: command.data, path: command.path, mime: command.mime, preserveCrop: command.preserveCrop !== false,
      updates: command.updates?.map(update => {
        const shape = find(source.context, update.id, slideIndex);
        return { id: shape.packageId || shape.id, properties: update.properties };
      }) });
    fragment.bounds = { left: root.left, top: root.top, width: root.width, height: root.height };
    if (slideIndex !== source.context.slideIndex) await PowerPoint.run(async context => {
      context.presentation.setSelectedSlides([source.context.slides.find(slide => slide.index === slideIndex).id]);
      await context.sync();
    });
    let created;
    let committed = false;
    try {
      await native('insertFragment', { fragment, context: source.context });
      const after = (await snapshot()).context;
      created = after.shapes.find(shape => shape.name === marker);
      if (!created) throw new Error('替换图片后未找到完整的新对象。');
      const pairs = [];
      function pair(previous, next) {
        if ((previous.children?.length || 0) !== (next.children?.length || 0)) throw new Error('图片组合的成员发生变化，已保留原对象。');
        pairs.push([previous, next]);
        (previous.children || []).forEach((child, index) => pair(child, next.children[index]));
      }
      pair(root, created);
      await PowerPoint.run(async context => {
        for (const [previous, next] of pairs) {
          const shape = officeShape(context, next);
          shape.name = previous.name;
          metadata(shape, { alternativeText: previous.id === original.id ? command.alternativeText ?? previous.alternativeText : previous.alternativeText,
            tags: { ...previous.tags, ...(previous.id === original.id ? command.tags : {}) } });
        }
        await context.sync();
      });
      await native('apply', { commands: [{ op: 'zOrder', ref: created.ref, position: 'index', index: root.zOrder }] });
      await PowerPoint.run(async context => {
        context.presentation.slides.getItemAt(slideIndex).shapes.getItem(root.id).delete();
        await context.sync();
      });
      committed = true;
      for (const [previous, next] of pairs) if (previous.id !== next.id) aliases.set(slideIndex + ':' + previous.id, next.id);
    } catch (error) {
      if (!committed) {
        const current = (await snapshot()).context;
        const added = created ? current.shapes.find(shape => shape.id === created.id) : current.shapes.find(shape => shape.name === marker);
        if (added) await PowerPoint.run(async context => {
          officeShape(context, added).delete();
          await context.sync();
        });
      }
      throw error;
    }
  }
  async function mutate(model, command, aliases) {
    const ids = command.ids || (command.id !== undefined ? [command.id] : []);
    const slideIndex = command.slideIndex ?? model.slideIndex;
    if (command.op === 'replaceImage') { await replaceObject(model, command, aliases); return; }
    if (command.op === 'select') { await select(model, ids, slideIndex); return; }
    if (command.op === 'delete' || command.op === 'group' || command.op === 'ungroup') {
      await PowerPoint.run(async context => {
        if (command.op === 'delete') {
          const selected = new Set(ids.map(String));
          // Members of any selected ancestor group go with it; nested groups included.
          const insideSelected = shape => {
            for (let parent = shape.parentId; parent; parent = find(model, parent, slideIndex).parentId) if (selected.has(String(parent))) return true;
            return false;
          };
          for (const id of selected) {
            const shape = find(model, id, slideIndex);
            if (!insideSelected(shape)) officeShape(context, shape).delete();
          }
        } else if (command.op === 'ungroup') officeShape(context, find(model, command.id, slideIndex)).group.ungroup();
        else {
          const slide = context.presentation.slides.getItemAt(slideIndex);
          slide.setSelectedShapes(ids.map(String));
          await context.sync();
          const grouped = context.presentation.getSelectedShapes().group();
          if (command.name !== undefined) grouped.name = command.name;
          metadata(grouped, command);
        }
        await context.sync();
      });
      return;
    }
    if (command.op === 'insertFragment') {
      await native('insertFragment', { fragment: command.fragment, options: command, context: model });
      return;
    }
    const previousShapes = slideIndex === model.slideIndex ? model.shapes : model.slides.find(slide => slide.index === slideIndex).shapes;
    const before = new Set(flatten(previousShapes).map(shape => shape.id));
    const original = command.id !== undefined ? find(model, command.id, slideIndex) : null;
    await native('apply', { commands: [{ ...command, ...(original ? { ref: original.ref } : {}) }] });
    if (['insertText', 'insertShape', 'insertImage', 'insertMedia', 'duplicate', 'replaceImage'].includes(command.op)) {
      const after = (await snapshot({ scope: slideIndex === model.slideIndex ? 'slide' : 'presentation' })).context;
      const shapes = slideIndex === after.slideIndex ? after.shapes : after.slides.find(slide => slide.index === slideIndex).shapes;
      const created = flatten(shapes).filter(shape => !before.has(shape.id));
      const target = command.op === 'replaceImage' && !created.length ? find(after, command.id, slideIndex) : created.find(shape => shape.parentId === original?.parentId) || created[0];
      if (!target) throw new Error('原生操作后未找到新增对象。');
      const inherited = ['replaceImage', 'duplicate'].includes(command.op) ? original : {};
      const values = { alternativeText: command.alternativeText ?? inherited.alternativeText,
        tags: { ...inherited.tags, ...command.properties?.tags, ...command.tags } };
      await PowerPoint.run(async context => {
        metadata(officeShape(context, target), values);
        await context.sync();
      });
      // Reflect the just-written metadata in the reused model instead of snapshotting again.
      if (values.alternativeText !== undefined) target.alternativeText = String(values.alternativeText);
      target.tags = { ...target.tags };
      for (const [key, value] of Object.entries(values.tags)) {
        if (value === '') delete target.tags[key];
        else if (value !== undefined && value !== null) target.tags[key] = String(value);
      }
      return after;
    }
  }
  async function apply(commands) {
    if (!Array.isArray(commands)) throw new Error('编辑命令必须是数组。');
    const scope = commands.some(command => command.slideIndex !== undefined) ? 'presentation' : 'slide';
    let model = (await snapshot({ scope })).context;
    const defaultSlide = model.slideIndex, aliases = new Map();
    const resolve = (id, slideIndex) => {
      let current = String(id);
      while (aliases.has(slideIndex + ':' + current)) current = aliases.get(slideIndex + ':' + current);
      return current;
    };
    const remap = command => {
      const slideIndex = command.slideIndex ?? defaultSlide;
      return { ...command, slideIndex,
        ...(command.id !== undefined ? { id: resolve(command.id, slideIndex) } : {}),
        ...(command.ids ? { ids: command.ids.map(id => resolve(id, slideIndex)) } : {}) };
    };
    for (let index = 0; index < commands.length;) {
      let after;
      if (commands[index].op === 'update') {
        const updates = [];
        while (index < commands.length && commands[index].op === 'update') updates.push(remap(commands[index++]));
        after = await applyUpdates(model, updates, aliases);
      } else after = await mutate(model, remap(commands[index++]), aliases);
      model = after || (await snapshot({ scope })).context;
    }
    emit();
    return model;
  }
  // Listeners learn whether only the selection moved: the scientific sync then skips a full
  // slide export, which otherwise ran on every click once the watch works.
  let emittedKind = null;
  function emit(kind = 'content') {
    emittedKind = emittedKind === 'content' || kind === 'content' ? 'content' : 'selection';
    if (notifying) return;
    notifying = true;
    queueMicrotask(() => {
      notifying = false;
      const change = { kind: emittedKind };
      emittedKind = null;
      for (const callback of listeners.keys()) {
        try { callback(change); } catch (error) { services?.reportError?.(error, { operation: 'powerpoint-change' }); }
      }
    });
  }
  const isHidden = () => globalThis.document?.visibilityState === 'hidden';
  function hasContentListener() {
    for (const entry of listeners.values()) if (entry.mode !== 'selection') return true;
    return false;
  }
  function needsPolling() {
    return !stopped && listeners.size > 0 && !isHidden() && (hasContentListener() || !selectionEventSupported);
  }
  function schedulePoll(delay = 1500) {
    clearTimeout(timer);
    timer = undefined;
    if (needsPolling()) timer = setTimeout(poll, delay);
  }
  function reconcileWatch() {
    if (!needsPolling()) {
      clearTimeout(timer);
      timer = undefined;
      return;
    }
    if (!timer) timer = setTimeout(poll, 1500);
  }
  function onSelectionEvent() {
    if (stopped || !listeners.size || isHidden()) return;
    emit('selection');
    if (hasContentListener() && !pending) schedulePoll(50);
  }
  function removeSelectionEvent(force = false) {
    if (!force && !selectionSubscribed) return;
    selectionSubscribed = false;
    selectionEventSupported = false;
    const documentApi = Office?.context?.document;
    const eventType = Office?.EventType?.DocumentSelectionChanged;
    if (typeof documentApi?.removeHandlerAsync === 'function' && eventType) {
      try { documentApi.removeHandlerAsync(eventType, { handler: onSelectionEvent }, () => {}); } catch {}
    }
  }
  function ensureSelectionEvent() {
    if (selectionSubscribed) return;
    const documentApi = Office?.context?.document;
    const eventType = Office?.EventType?.DocumentSelectionChanged;
    if (typeof documentApi?.addHandlerAsync !== 'function' || !eventType) return;
    selectionSubscribed = true;
    try {
      documentApi.addHandlerAsync(eventType, onSelectionEvent, result => {
        const failedStatus = Office?.AsyncResultStatus?.Failed;
        const failed = Boolean(result && (result.status === 'failed' || (failedStatus !== undefined && result.status === failedStatus)));
        if (failed || stopped || !listeners.size) {
          selectionEventSupported = false;
          if ((stopped || !listeners.size) && !failed) removeSelectionEvent(true);
          else selectionSubscribed = false;
          reconcileWatch();
          return;
        }
        selectionEventSupported = true;
        reconcileWatch();
      });
    } catch {
      selectionSubscribed = false;
      selectionEventSupported = false;
    }
  }
  function onVisibilityChange() {
    if (isHidden()) {
      clearTimeout(timer);
      timer = undefined;
      return;
    }
    reconcileWatch();
    if (!stopped && listeners.size) emit('selection');
  }
  function detachDocumentWatchers() {
    clearTimeout(timer);
    timer = undefined;
    removeSelectionEvent();
    if (visibilitySubscribed) {
      visibilitySubscribed = false;
      globalThis.document?.removeEventListener?.('visibilitychange', onVisibilityChange);
    }
  }
  // Cheap watch fingerprint: live IDs, geometry, type, name, z-order and text only.
  async function surfaceFingerprint() {
    return readPowerPoint(async context => {
      const presentation = context.presentation;
      const slides = presentation.slides;
      const selectedSlides = presentation.getSelectedSlides();
      const selection = presentation.getSelectedShapes();
      slides.load('items/id');
      selectedSlides.load('items/id');
      selection.load('items/id');
      presentation.load('id');
      await context.sync();
      if (!slides.items.length) throw new Error('请先在当前演示文稿中添加幻灯片。');
      const indices = selectedSlides.items.map(slide => slides.items.findIndex(item => item.id === slide.id));
      const slideIndex = indices[0] ?? 0;
      const tree = await readSlideTree(context, slideIndex, geometryShapeLoad + ',items/zOrderPosition', true);
      const shapes = flatten(tree);
      const texts = new Map();
      for (const shape of shapes) {
        if (!textTypes.has(shape.nativeType)) continue;
        const range = officeShape(context, shape).textFrame.textRange;
        range.load('text');
        texts.set(shape.id, range);
      }
      // Placeholders holding a picture, chart or table have no text frame; their load fails the
      // whole batch, so fall back to a geometry-only fingerprint instead of losing change detection.
      if (texts.size) {
        try { await context.sync(); } catch { texts.clear(); }
      }
      const round = value => Number.isFinite(value) ? Math.round(value * 100) / 100 : value;
      const rows = shapes.map(shape => {
        const native = texts.get(shape.id);
        return [shape.id, shape.type, round(shape.left), round(shape.top), round(shape.width), round(shape.height),
          round(shape.rotation), shape.name, shape.zOrder, shape.parentId ?? '', native ? String(native.text ?? '') : ''];
      });
      return { content: JSON.stringify([presentation.id, slideIndex, slides.items.length, rows]),
        selection: selection.items.map(shape => String(shape.id)).join(',') };
    });
  }
  // Lightweight Office.js watch only: never export full PPTX packages on a background timer.
  async function poll() {
    timer = undefined;
    if (!needsPolling()) return;
    if (pending) { schedulePoll(1500); return; }
    try {
      const cheap = await serial(surfaceFingerprint);
      const previous = surface;
      surface = cheap;
      if (previous && cheap.content !== previous.content) emit('content');
      else if (previous && !selectionEventSupported && cheap.selection !== previous.selection) emit('selection');
    } catch (error) { await services?.reportError?.(error, { operation: 'powerpoint-watch' }); }
    schedulePoll(1500);
  }
  const host = {
    name: 'Microsoft PowerPoint', kind: 'powerpoint',
    context: options => serial(async () => options?.geometry ? geometryContext(options) : (await snapshot(options)).context),
    summary: (options = {}) => serial(async () => {
      const withSlides = options?.slides !== false;
      const current = await quick({ slides: withSlides });
      return { presentationId: current.presentationId, title: current.title, slideIndex: current.slideIndex,
        slideId: current.slideId, width: current.width, height: current.height, selectionIds: current.selectionIds,
        slides: withSlides ? current.slides : [], slideCount: current.slideCount,
        selectedSlideIndices: current.selectedSlideIndices };
    }),
    align: edge => serial(() => align(edge)),
    apply: (commands, options = {}) => serial(async () => {
      if (!Array.isArray(commands)) throw new Error('编辑命令必须是数组。');
      if (!commands.length) return (await snapshot({ scope: 'slide' })).context;
      let rich = false, moved = false;
      const movedIds = new Set();
      for (const command of commands) {
        if (!command || typeof command.op !== 'string') throw new Error('无效的原生命令。');
        if (geometryUpdate(command)) { moved = true; movedIds.add(String(command.id)); }
        else if (command.op !== 'select') rich = true;
        if (command.slideIndex !== undefined) rich = true;
      }
      // Roll back newly created top-level groups without exporting the slide or using AppleScript.
      if (options.readback === false && commands.every(command => command.op === 'ungroup' && command.slideIndex === undefined)) {
        await PowerPoint.run(async context => {
          const presentation = context.presentation;
          const slides = options.target ? presentation.slides : presentation.getSelectedSlides();
          slides.load('items/id');
          if (options.target) presentation.load('id');
          await context.sync();
          if (!options.target && slides.items.length !== 1) throw new Error('请先选中要取消编组的幻灯片。');
          if (options.target) targetSlideIndex(presentation, slides, options.target);
          const slide = options.target ? slides.getItem(String(options.target.slideId)) : slides.items[0];
          const shapes = slide.shapes;
          for (const { id } of commands) shapes.getItem(String(id)).group.ungroup();
          await context.sync();
        });
        emit();
        return { applied: commands.length };
      }
      if (options.target) {
        if (!isSimpleOfficeBatch(commands)) throw new Error('该命令不能绑定到指定幻灯片。');
        const model = await applySimpleOffice(commands, options.readback !== false, options.target);
        if (!model) throw new Error('原幻灯片上的对象已经变化，请重新选择后重试。');
        return model;
      }
      if (options.readback === false && commands.every(command => command.op === 'update'
        && command.slideIndex === undefined && Object.keys(command.properties || {}).every(key => directProperties.has(key)))) {
        const selected = await PowerPoint.run(async context => {
          const selection = context.presentation.getSelectedShapes();
          selection.load('items/id');
          await context.sync();
          const targets = new Map(selection.items.map(shape => [String(shape.id), shape]));
          if (!commands.every(command => targets.has(String(command.id)))) return false;
          for (const command of commands) {
            const shape = targets.get(String(command.id));
            for (const key of directProperties) {
              if (!['tags', 'alternativeText'].includes(key) && command.properties[key] !== undefined) shape[key] = command.properties[key];
            }
            metadata(shape, command.properties);
          }
          await context.sync();
          return true;
        });
        if (selected) { emit(); return { applied: commands.length }; }
        const model = await geometryContext({ scope: 'slide' });
        await applyDirectUpdates(model, commands);
        emit();
        return { applied: commands.length };
      }
      const lightweight = !rich && (options.geometry === true || !moved);
      if (lightweight) {
        const model = await applyGeometry(commands, movedIds, options.readback !== false);
        if (model) return model;
      }
      if (isSimpleOfficeBatch(commands)) {
        const model = await applySimpleOffice(commands, options.readback !== false);
        if (model) return model;
      }
      return apply(commands);
    }),
    images: (ids, options = {}) => serial(async () => {
      const source = await snapshot();
      const selected = ids.map(id => find(source.context, id));
      const images = await request('pptx/images', { ...options, ids: selected.map(shape => shape.packageId || shape.id), pptx: source.pptx, slideIndex: source.packageSlideIndex });
      const nativeIds = new Map(selected.map(shape => [shape.packageId || shape.id, shape.id]));
      return images.map(image => ({ ...image, id: nativeIds.get(image.id) }));
    }),
    fragment: ids => serial(async () => {
      const source = await snapshot();
      const selected = ids.map(id => find(source.context, id));
      const fragment = await request('pptx/fragment', { ids: selected.map(shape => shape.packageId || shape.id), pptx: source.pptx, slideIndex: source.packageSlideIndex });
      const left = Math.min(...selected.map(shape => shape.left)), top = Math.min(...selected.map(shape => shape.top));
      fragment.bounds = { left, top, width: Math.max(...selected.map(shape => shape.left + shape.width)) - left,
        height: Math.max(...selected.map(shape => shape.top + shape.height)) - top };
      return fragment;
    }),
    insertFragment: (fragment, options = {}) => serial(async () => {
      const source = options.target ? { context: await geometryContext({ target: options.target }) } : await snapshot();
      await native('insertFragment', { fragment, options, context: source.context,
        ...(options.target ? { slideId: source.context.slideId, documentUrl: Office.context.document.url } : {}) });
      const result = options.readback === false ? undefined : (await snapshot()).context;
      emit(); return result;
    }),
    rasterizeImage: (id, options = {}) => serial(async () => {
      const source = await snapshot();
      const shape = find(source.context, id);
      const fragment = await request('pptx/imageFragment', { pptx: source.pptx, id: shape.packageId || shape.id, slideIndex: source.packageSlideIndex, original: true });
      return native('rasterizeImage', { ...options, fragment });
    }),
    export: options => serial(async () => {
      const source = await snapshot({ scope: options.kind === 'slides' ? 'presentation' : 'slide' });
      let fragment;
      if (options.kind === 'selection') fragment = await request('pptx/fragment', {
        pptx: source.pptx, slideIndex: source.packageSlideIndex,
        ids: (options.ids || source.context.selectionIds).map(id => { const shape = find(source.context, id); return shape.packageId || shape.id; }) });
      return native('export', { ...options, pptx: source.pptx, slideIndex: source.packageSlideIndex, context: source.context, fragment });
    }),
    onChanged(callback, options = {}) {
      const mode = options?.mode === 'selection' ? 'selection' : 'all';
      const first = listeners.size === 0;
      listeners.set(callback, { mode });
      if (first && !stopped) {
        ensureSelectionEvent();
        if (!visibilitySubscribed && typeof globalThis.document?.addEventListener === 'function') {
          visibilitySubscribed = true;
          globalThis.document.addEventListener('visibilitychange', onVisibilityChange);
        }
      }
      reconcileWatch();
      return () => {
        listeners.delete(callback);
        if (!listeners.size) detachDocumentWatchers();
        else reconcileWatch();
      };
    },
    dispose() { stopped = true; detachDocumentWatchers(); listeners.clear(); }
  };
  return host;
}
