/*
 * SlideSCI native ribbon entry points.
 *
 * The PowerPoint manifest registers ExecuteFunction commands. Excel-style
 * function commands run in the shared runtime, which is the same JavaScript
 * context as the task pane, so a ribbon click forwards an activation target
 * to the pane and the pane can show itself. Registration must not wait for
 * the local SlideSCI service: a cold runtime has to answer ribbon clicks even
 * when the native helper is still starting.
 */

const OFFICE_JS_URL = 'https://appsforoffice.microsoft.com/lib/1/hosted/office.js';
const REQUEST_KEY = 'slidesci.pane.request';
const READY_TIMEOUT = 10000;
const COMPLETE_TIMEOUT = 4000;

const SECTION = Object.freeze({
  editing: '排版与编辑',
  images: '图片',
  science: '科学图片',
  content: '内容与公式',
  library: '素材库',
  ai: 'AI 助手',
  export: '导出',
  help: '设置与帮助'
});

function target(tab, action, execute = false) {
  const value = { tab };
  if (action) value.action = action;
  if (execute) value.execute = true;
  return value;
}

function panel(tab, panelId) {
  return { tab, panel: panelId };
}

/*
 * FunctionName in office/manifest.xml must match these keys exactly.
 * `execute: true` is only used for actions that take no fields; the pane runs
 * them once on activation. Everything else only navigates and focuses the
 * form so the user keeps control over slide mutations.
 */
export const RIBBON_COMMANDS = Object.freeze({
  SlideSCIArrangeImages: target(SECTION.editing, 'editing.grid'),
  SlideSCIAlignLeft: target(SECTION.editing, 'editing.align.left', true),
  SlideSCIAlignHorizontalCenter: target(SECTION.editing, 'editing.align.horizontalCenter', true),
  SlideSCIAlignRight: target(SECTION.editing, 'editing.align.right', true),
  SlideSCIAlignTop: target(SECTION.editing, 'editing.align.top', true),
  SlideSCIAlignVerticalCenter: target(SECTION.editing, 'editing.align.verticalCenter', true),
  SlideSCIAlignBottom: target(SECTION.editing, 'editing.align.bottom', true),
  SlideSCISpacingHorizontal: target(SECTION.editing, 'editing.spacing.horizontal'),
  SlideSCISpacingVertical: target(SECTION.editing, 'editing.spacing.vertical'),
  SlideSCIDistributeHorizontal: target(SECTION.editing, 'editing.distribute.horizontal', true),
  SlideSCIDistributeVertical: target(SECTION.editing, 'editing.distribute.vertical', true),
  SlideSCICopyDimensions: target(SECTION.editing, 'editing.copy.dimensions', true),
  SlideSCIPasteDimensions: target(SECTION.editing, 'editing.paste.dimensions', true),
  SlideSCICopyWidth: target(SECTION.editing, 'editing.copy.width', true),
  SlideSCIPasteWidth: target(SECTION.editing, 'editing.paste.width', true),
  SlideSCICopyHeight: target(SECTION.editing, 'editing.copy.height', true),
  SlideSCIPasteHeight: target(SECTION.editing, 'editing.paste.height', true),
  SlideSCICopyPosition: target(SECTION.editing, 'editing.copyPosition'),
  SlideSCIPastePosition: target(SECTION.editing, 'editing.pastePosition', true),
  SlideSCISwapPosition: target(SECTION.editing, 'editing.swapPosition'),
  SlideSCICopyRelativePosition: target(SECTION.editing, 'editing.copyRelativePosition'),
  SlideSCIPasteRelativePosition: target(SECTION.editing, 'editing.pasteRelativePosition', true),
  SlideSCICopyShapeFormat: target(SECTION.editing, 'editing.copyShapeFormat'),
  SlideSCIPasteShapeFormat: target(SECTION.editing, 'editing.pasteShapeFormat', true),
  SlideSCICopyTextFormat: target(SECTION.editing, 'editing.copyTextFormat'),
  SlideSCIPasteTextFormat: target(SECTION.editing, 'editing.pasteTextFormat', true),
  SlideSCICopyGroupFormat: target(SECTION.editing, 'editing.copyGroupFormat', true),
  SlideSCIPasteGroupFormat: target(SECTION.editing, 'editing.pasteGroupFormat', true),
  SlideSCICopyCrop: target(SECTION.editing, 'editing.copyCrop', true),
  SlideSCIPasteCrop: target(SECTION.editing, 'editing.pasteCrop', true),
  SlideSCITitlesBelow: target(SECTION.editing, 'editing.addTitlesBelow'),
  SlideSCITitlesAbove: target(SECTION.editing, 'editing.addTitlesAbove'),
  SlideSCILabels: target(SECTION.editing, 'editing.addLabels'),
  SlideSCIUpdateLabels: target(SECTION.editing, 'editing.updateLabels'),

  SlideSCIInsertImages: target(SECTION.images, 'images.insert', true),
  SlideSCIPasteImages: target(SECTION.images, 'images.pasteFiles', true),
  SlideSCIPseudoColor: panel(SECTION.images, 'images.pseudoColor'),
  SlideSCIMergeChannels: target(SECTION.images, 'images.merge', true),
  SlideSCIRevealSource: target(SECTION.images, 'images.reveal', true),

  SlideSCIScaleBar: panel(SECTION.science, 'science.scale'),
  SlideSCIZoomImages: panel(SECTION.science, 'science.zoom'),
  SlideSCISetFov: target(SECTION.science, 'science.setFov'),
  SlideSCIReadFov: target(SECTION.science, 'science.readFov', true),

  SlideSCIMarkdown: target(SECTION.content, 'content.markdown'),
  SlideSCIEquation: target(SECTION.content, 'content.latex'),
  SlideSCICode: target(SECTION.content, 'content.code'),
  SlideSCILatexSvg: target(SECTION.content, 'content.latexSvg'),
  SlideSCISvg: target(SECTION.content, 'content.svg'),
  SlideSCIRichText: target(SECTION.content, 'content.richText'),

  SlideSCILibrary: panel(SECTION.library, 'library.browser'),
  SlideSCILibrarySave: target(SECTION.library, 'library.save'),
  SlideSCIExport: target(SECTION.export, 'export.slides'),
  SlideSCIExportOriginal: target(SECTION.export, 'export.originalImages', true),
  SlideSCICopyOriginal: target(SECTION.export, 'export.copyOriginal', true),
  SlideSCICopySelection: target(SECTION.export, 'export.copySelection', true),

  SlideSCIAssistant: panel(SECTION.ai, 'ai.chat'),
  SlideSCISettings: { tab: SECTION.help }
});

let sdkPromise = null;
let readyPromise = null;
let registration = null;

function message(error) {
  return error && error.message ? error.message : String(error);
}

/* Load Office.js once for both the pane and the function-command runtime. */
export function ensureOfficeSdk() {
  if (globalThis.Office && typeof globalThis.Office.onReady === 'function') return Promise.resolve();
  if (sdkPromise) return sdkPromise;
  sdkPromise = new Promise((resolve, reject) => {
    const selector = 'script[data-slidesci-office], script[src="' + OFFICE_JS_URL + '"]';
    const existing = document.querySelector(selector);
    if (existing) {
      if (globalThis.Office && typeof globalThis.Office.onReady === 'function') {
        resolve();
        return;
      }
      existing.addEventListener('load', () => resolve(), { once: true });
      existing.addEventListener('error', () => reject(new Error('Office.js 加载失败，请检查网络连接。')), { once: true });
      return;
    }
    const script = document.createElement('script');
    script.src = OFFICE_JS_URL;
    script.async = true;
    script.dataset.slidesciOffice = 'true';
    script.addEventListener('load', () => resolve(), { once: true });
    script.addEventListener('error', () => {
      sdkPromise = null;
      script.remove();
      reject(new Error('Office.js 加载失败，请检查网络连接。'));
    }, { once: true });
    document.head.append(script);
  });
  return sdkPromise;
}

/*
 * Resolve the host SDK even when the local service is down or the pane is
 * hidden. The timeout keeps a non-host browser from hanging forever.
 */
export async function ensureOfficeReady() {
  if (readyPromise) return readyPromise;
  const attempt = (async () => {
    const loading = (async () => {
      await ensureOfficeSdk();
      if (!globalThis.Office || typeof Office.onReady !== 'function') {
        throw new Error('Office.js 不可用，请从 PowerPoint 打开 SlideSCI。');
      }
      return Office.onReady();
    })();
    loading.catch(() => {});
    let timer;
    try {
      const info = await Promise.race([
        loading,
        new Promise((resolve, reject) => {
          timer = setTimeout(() => reject(new Error('Office.js 未就绪，请从 PowerPoint 打开 SlideSCI。')), READY_TIMEOUT);
        })
      ]);
      return info || { host: null, platform: null };
    } finally {
      clearTimeout(timer);
    }
  })();
  readyPromise = attempt;
  attempt.catch(() => { if (readyPromise === attempt) readyPromise = null; });
  return attempt;
}

/* Showing a hidden shared-runtime pane must never block or fail a command. */
async function showPane() {
  const addin = globalThis.Office && Office.addin;
  if (!addin || typeof addin.showAsTaskpane !== 'function') return;
  let timer;
  try {
    // Navigation must still be forwarded if the pane takes long to appear.
    await Promise.race([
      Promise.resolve().then(() => addin.showAsTaskpane()),
      new Promise(resolve => { timer = setTimeout(resolve, 2000); })
    ]);
  } catch (error) {
    console.warn('SlideSCI 无法显示任务窗格：' + message(error));
  } finally {
    clearTimeout(timer);
  }
}

/*
 * Fallback transport for a runtime where the shell did not provide navigate.
 * The timestamp lets the pane ignore an activation it has already handled.
 */
function remember(activation) {
  let at = Date.now();
  try {
    const saved = JSON.parse(globalThis.localStorage?.getItem(REQUEST_KEY) || 'null');
    if (saved && Number.isFinite(saved.at)) at = Math.max(at, saved.at + 1);
  } catch {
    // A malformed request is replaced below.
  }
  const request = { ...activation, at };
  try {
    globalThis.localStorage?.setItem(REQUEST_KEY, JSON.stringify(request));
  } catch (error) {
    console.warn('SlideSCI 无法保存功能区请求：' + message(error));
  }
  try {
    window.dispatchEvent(new CustomEvent('slidesci:activate', { detail: request }));
  } catch (error) {
    console.warn('SlideSCI 无法派发功能区请求：' + message(error));
  }
}

function dispatch(name, activation, navigate, event) {
  let settled = false;
  const finish = () => {
    if (settled) return;
    settled = true;
    clearTimeout(guard);
    try {
      if (event && typeof event.completed === 'function') event.completed();
    } catch (error) {
      console.warn('SlideSCI 无法结束功能区事件：' + message(error));
    }
  };
  // A lost pane or a rejected navigate must not leave the ribbon button disabled.
  const guard = setTimeout(finish, COMPLETE_TIMEOUT);
  (async () => {
    try {
      const showing = showPane();
      try {
        if (typeof navigate === 'function') await navigate({ ...activation });
        else remember(activation);
      } finally {
        await showing;
      }
    } catch (error) {
      console.warn('SlideSCI 功能区命令 ' + name + ' 执行失败：' + message(error));
    } finally {
      finish();
    }
  })();
}

/*
 * Called by the pane shell once the sections exist, before the host connects.
 * Returns a descriptor instead of rejecting so boot never waits on Office.
 */
export async function registerOfficeCommands(navigate) {
  if (registration) return registration;
  const attempt = (async () => {
    const hostParameter = new URLSearchParams(window.location.search).get('host');
    if (hostParameter === 'wps') return { registered: false, reason: 'wps-host' };
    let info;
    try {
      info = await ensureOfficeReady();
    } catch (error) {
      console.warn('SlideSCI 功能区未注册：' + message(error));
      return { registered: false, reason: 'office-sdk' };
    }
    const actions = globalThis.Office && Office.actions;
    if (!actions || typeof actions.associate !== 'function') {
      return { registered: false, reason: 'no-actions' };
    }
    if (info.host && info.host !== Office.HostType.PowerPoint) {
      return { registered: false, reason: 'host:' + info.host };
    }
    let count = 0;
    for (const [name, activation] of Object.entries(RIBBON_COMMANDS)) {
      try {
        actions.associate(name, event => dispatch(name, activation, navigate, event));
        count++;
      } catch (error) {
        console.warn('SlideSCI 功能区命令注册失败：' + name + '：' + message(error));
      }
    }
    return { registered: count > 0, count };
  })();
  registration = attempt;
  attempt.catch(() => { if (registration === attempt) registration = null; });
  return attempt;
}
