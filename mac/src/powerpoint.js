import { enrichPowerPointSlides } from './powerpoint-live.js';
import { alignment } from './features/editing-geometry.js';
import { ensureOfficeReady } from './ribbon.js';
const flatten = shapes => shapes.flatMap(shape => [shape, ...flatten(shape.children || [])]);
const directProperties = new Set(['left', 'top', 'width', 'height', 'rotation', 'name', 'tags', 'alternativeText']);
const geometryProperties = new Set(['left', 'top', 'width', 'height', 'rotation']);
const geometryShapeLoad = 'items/id,items/name,items/left,items/top,items/width,items/height,items/rotation,items/type';
const shapeTypes = { Group: 'group', Image: 'picture', Media: 'media', Table: 'table', TextBox: 'text' };
const textTypes = new Set(['TextBox', 'Placeholder', 'GeometricShape', 'Callout', 'Freeform']);
// Office returns a path or URL; a file name such as "增长 20%.pptx" is not valid URI encoding.
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
  if (!services?.request) throw new Error('缺少 SlideSCI 本地服务连接。');
  const request = (route, input) => services.request(route, input);
  const native = (action, input = {}) => request('native/execute', { host: 'powerpoint', action, ...input,
    ...(boundTitle && action !== 'rasterizeImage' ? { presentationName: Office.context.document.url ? documentName(Office.context.document.url) : boundTitle } : {}) });
  let queue = Promise.resolve(), timer, stopped = false, fingerprint, surface, deepPolls = 0, notifying = false, pending = 0;
  const listeners = new Set();
  let boundTitle = '';
  const documentInfo = await native('documentInfo');
  boundTitle = documentInfo.name || '';
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
  async function activeSlideIndex(slides, selected) {
    const indices = selected.items.map(slide => slides.items.findIndex(item => item.id === slide.id));
    let slideIndex = indices[0] ?? 0;
    if (indices.length > 1) {
      const active = await native('documentInfo');
      if (indices.includes(active.slideIndex)) slideIndex = active.slideIndex;
    }
    return { indices, slideIndex };
  }
  async function quick(options = {}) {
    return PowerPoint.run(async context => {
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
      } else {
        ({ indices, slideIndex } = await activeSlideIndex(slides, selected));
        slideId = slides.items[slideIndex].id;
      }
      return { presentationId: presentation.id, title: title(), slideIndex, slideId,
        width: presentation.pageSetup.slideWidth, height: presentation.pageSetup.slideHeight,
        selectionIds: selection.items.map(shape => String(shape.id)), selectedSlideIndices: indices,
        slideCount: count.value, slides: options.slides === false ? [] : slides.items.map((slide, index) => ({ id: slide.id, index })) };
    });
  }
  async function readSlideTree(context, slideIndex, load = geometryShapeLoad, withZOrder = false) {
    const shapes = [];
    let level = [{ collection: context.presentation.slides.getItemAt(slideIndex).shapes, path: [], output: shapes }];
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
    return PowerPoint.run(async context => {
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
      const { indices, slideIndex } = await activeSlideIndex(slides, selectedSlides);
      const selectionIds = selection.items.map(shape => String(shape.id));
      const wanted = new Set([...selectionIds, ...extraIds.map(String)]);
      let shapes = [];
      if (options.scope === 'slide') shapes = await readSlideTree(context, slideIndex);
      else if (wanted.size) shapes = selectSubtrees(await readSlideTree(context, slideIndex), wanted);
      return { presentationId: presentation.id, title: title(), slideId: slides.items[slideIndex].id, slideIndex,
        width: presentation.pageSetup.slideWidth, height: presentation.pageSetup.slideHeight, geometry: true,
        selectionIds, selectedSlideIndices: indices, shapes };
    });
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
      if (indices.length > 1) return false;
      const slideIndex = indices[0] ?? 0;
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
    const data = await PowerPoint.run(async context => {
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
    if (!all && !options.watch) fingerprint = contentFingerprint(model);
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
      for (const callback of listeners) {
        try { callback(change); } catch (error) { services.reportError(error, { operation: 'powerpoint-change' }); }
      }
    });
  }
  // Selection and text cursor are watched by the cheap pass; the deep pass compares content only.
  function contentFingerprint(model) { return JSON.stringify({ ...model, selectionIds: undefined, textSelection: undefined }); }
  // Cheap watch fingerprint: live IDs, geometry, type, name, z-order and text only.
  // Rich formats (fill/line/crop/image bytes) still come from the periodic pptx snapshot.
  async function surfaceFingerprint() {
    return PowerPoint.run(async context => {
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
  // Rich formats (fill, crop, image bytes, alt text) need a slide export; check them once every
  // few unchanged cheap passes, as the WPS watch does.
  const DEEP_POLL_INTERVAL = 4;
  async function poll() {
    if (stopped || !listeners.size) return;
    if (pending) { timer = setTimeout(poll, 1500); return; }
    try {
      const cheap = await serial(surfaceFingerprint);
      if (surface === undefined) surface = cheap;
      else if (cheap.content !== surface.content) { surface = cheap; deepPolls = 0; emit('content'); }
      else if (cheap.selection !== surface.selection) { surface = cheap; emit('selection'); }
      else if (++deepPolls >= DEEP_POLL_INTERVAL) {
        deepPolls = 0;
        const next = await serial(() => snapshot({ scope: 'slide', watch: true }));
        const current = contentFingerprint(next.context);
        if (fingerprint !== undefined && current !== fingerprint) emit('content');
        fingerprint = current;
      }
    } catch (error) { await services.reportError(error, { operation: 'powerpoint-watch' }); }
    if (!stopped && listeners.size) timer = setTimeout(poll, 1500);
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
      const source = await snapshot();
      await native('insertFragment', { fragment, options, context: source.context });
      const result = (await snapshot()).context;
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
    onChanged(callback) {
      listeners.add(callback);
      if (listeners.size === 1 && !stopped) timer = setTimeout(poll, 1500);
      return () => { listeners.delete(callback); if (!listeners.size) clearTimeout(timer); };
    },
    dispose() { stopped = true; clearTimeout(timer); listeners.clear(); }
  };
  return host;
}
