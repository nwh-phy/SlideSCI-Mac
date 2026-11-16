(function () {
  'use strict';

  // WPS Mac needs native resource keys; many Office imageMso names resolve to question marks.
  var ribbonImages = {
    SlideSciAlignImages: 'AdpDiagramArrangeTables',
    SlideSciAlignMenu: 'align_draw_left_kd',
    SlideSciAlignLeft: 'align_draw_left_kd',
    SlideSciAlignHorizontalCenter: 'align_draw_center_kd',
    SlideSciAlignRight: 'align_draw_right_kd',
    SlideSciAlignTop: 'align_draw_top_kd',
    SlideSciAlignVerticalCenter: 'align_draw_vertical_kd',
    SlideSciAlignBottom: 'align_draw_bottom_kd',
    SlideSciSizeMenu: 'GroupSizeAndPosition',
    SlideSciSpacingHorizontal: 'ranks_distribute_columns_kd',
    SlideSciSpacingVertical: 'ranks_distribute_rows_kd',
    SlideSciDistributeHorizontal: 'align_distribute_horizontal_kd',
    SlideSciDistributeVertical: 'align_distribute_vertical_kd',
    SlideSciCopyDimensions: 'copy_kd',
    SlideSciPasteDimensions: 'paste_kd',
    SlideSciCopyWidth: 'SizeToWidest',
    SlideSciPasteWidth: 'paste_kd',
    SlideSciCopyHeight: 'SizeToTallest',
    SlideSciPasteHeight: 'paste_kd',
    SlideSciPositionMenu: 'PicturePositionGallery',
    SlideSciCopyPosition: 'copy_kd',
    SlideSciPastePosition: 'paste_kd',
    SlideSciSwapPosition: 'DiagramReverseClassic',
    SlideSciCopyRelativePosition: 'copy_kd',
    SlideSciPasteRelativePosition: 'paste_kd',
    SlideSciFormatMenu: 'brush_kd',
    SlideSciCopyShapeFormat: 'brush_kd',
    SlideSciPasteShapeFormat: 'form_paste_format_kd',
    SlideSciCopyTextFormat: 'symbol_a_kd',
    SlideSciPasteTextFormat: 'form_text_paste_kd',
    SlideSciCopyGroupFormat: 'group_kd',
    SlideSciPasteGroupFormat: 'form_paste_format_kd',
    SlideSciCopyCrop: 'crop_kd',
    SlideSciPasteCrop: 'form_paste_format_kd',
    SlideSciTitlesMenu: 'insert_caption_kd',
    SlideSciTitlesBelow: 'insert_caption_kd',
    SlideSciTitlesAbove: 'insert_caption_kd',
    SlideSciAddLabelsButton: 'text_frame_kd',
    SlideSciUpdateLabelsButton: 'text_frame_kd',
    SlideSciInsertImages: 'pic_kd',
    SlideSciPseudoColorButton: 'color_palette_kd',
    SlideSciMergeChannelsButton: 'group_kd',
    SlideSciRevealSource: 'FileOpen',
    SlideSciPasteImages: 'form_paste_pic_from_clipboard_kd',
    SlideSciScaleBarButton: 'RulerShowHide',
    SlideSciZoomButton: 'screen_zoom_kd',
    SlideSciScienceMenu: 'GroupSizeAndPosition',
    SlideSciSetFovButton: 'GroupSizeAndPosition',
    SlideSciReadFovButton: 'Refresh',
    SlideSciMarkdownButton: 'TextFromFileInsert',
    SlideSciEquationButton: 'math_latex_kd',
    SlideSciContentMenu: 'math_omega_kd',
    SlideSciCodeButton: 'GroupCode',
    SlideSciLatexSvgButton: 'math_equation_professional_format_kd',
    SlideSciSvgButton: 'code_xml_structure_kd',
    SlideSciRichTextButton: 'symbol_a_kd',
    SlideSciLibraryButton: 'docer_resource_img_kd',
    SlideSciExportSlidesButton: 'replace_pic_kd',
    SlideSciExportMenu: 'ExportMoreMenu',
    SlideSciExportOriginalButton: 'change_to_pic_two_kd',
    SlideSciCopyOriginalButton: 'copy_as_pic_kd',
    SlideSciCopySelectionButton: 'copy_kd',
    SlideSciLibrarySaveButton: 'save_kd',
    SlideSciOpenButton: 'SelectionPane',
    SlideSciAIShortcut: 'bubble_add_kd',
    SlideSciSettingsButton: 'setup_kd',
  };

  window.SlideSciGetImage = function (control) {
    return ribbonImages[control.Id];
  };

  var PANE_REQUEST_KEY = 'slidesci.pane.request';
  var pane = null;
  var pending = { tab: null, action: null, panel: null, execute: false };
  var lastRequestAt = 0;
  var backgroundPromise = null;
  var backgroundServices = null;
  var paneChannel = typeof BroadcastChannel === 'function' ? new BroadcastChannel(PANE_REQUEST_KEY) : null;

  function taskpaneUrl() {
    return new URL('/taskpane.html?host=wps', window.location.href).href;
  }

  function paneUrl() {
    var url = new URL(taskpaneUrl());
    if (pending.tab) url.searchParams.set('tab', pending.tab);
    if (pending.action) url.searchParams.set('action', pending.action);
    if (pending.panel) url.searchParams.set('panel', pending.panel);
    if (pending.execute) url.searchParams.set('execute', 'true');
    return url.href;
  }

  function rememberTab(tab, extra) {
    if (!tab) return;
    pending = {
      tab: tab,
      action: extra && extra.action ? extra.action : null,
      panel: extra && extra.panel ? extra.panel : null,
      execute: Boolean(extra && extra.execute)
    };
    var previous = 0;
    try {
      var saved = JSON.parse(window.localStorage.getItem(PANE_REQUEST_KEY) || 'null');
      if (saved && Number.isFinite(saved.at)) previous = saved.at;
    } catch (error) {
      previous = 0;
    }
    var request = { tab: pending.tab, at: Math.max(Date.now(), previous + 1, lastRequestAt + 1) };
    lastRequestAt = request.at;
    if (pending.action) request.action = pending.action;
    if (pending.panel) request.panel = pending.panel;
    if (pending.execute) request.execute = true;
    try {
      window.localStorage.setItem(PANE_REQUEST_KEY, JSON.stringify(request));
    } catch (error) {
      console.warn('SlideSCI 无法保存选项卡请求：' + error.message);
    }
    if (paneChannel) paneChannel.postMessage(request);
  }

  function createPane() {
    var application = window.Application;
    if (!application) throw new Error('WPS 演示接口不可用。请从 WPS 演示中打开 SlideSCI 加载项。');
    if (typeof application.CreateTaskPane === 'function') {
      pane = application.CreateTaskPane(paneUrl());
    } else if (application.TaskPanes && typeof application.TaskPanes.Add === 'function') {
      pane = application.TaskPanes.Add(paneUrl());
    } else {
      throw new Error('当前 WPS 版本不支持侧边栏面板，请升级 WPS 演示后重试。');
    }
    if (pane && 'Width' in pane) {
      try {
        pane.Width = 430;
      } catch (error) {
        console.warn('SlideSCI 无法设置面板宽度：' + error.message);
      }
    }
    return pane;
  }

  function openPane(tab, extra) {
    rememberTab(tab, extra);
    try {
      if (!pane) createPane();
      pane.Visible = true;
      return true;
    } catch (error) {
      pane = null;
      try {
        createPane();
        pane.Visible = true;
        return true;
      } catch (retryError) {
        pane = null;
        window.alert('SlideSCI 无法打开侧边栏：' + retryError.message);
        return false;
      }
    }
  }

  function openAt(tab, extra) {
    return function () {
      return openPane(tab, extra);
    };
  }

  function startBackground() {
    if (backgroundPromise) return backgroundPromise;
    backgroundPromise = Promise.all([
      import(new URL('/src/services.js', window.location.href).href),
      import(new URL('/src/background.js', window.location.href).href)
    ]).then(function (modules) {
      var createServices = modules[0].createServices;
      var startBackgroundTracking = modules[1].startBackgroundTracking;
      var services = createServices({ host: 'wps' });
      backgroundServices = services;
      return startBackgroundTracking({
        host: 'wps',
        services: services,
        role: 'wps-hidden'
      });
    }).catch(function (error) {
      backgroundPromise = null;
      console.error('SlideSCI 后台跟踪启动失败：' + (error && error.message ? error.message : error));
      if (backgroundServices) backgroundServices.reportError(error, { operation: 'wps-background-start' }).catch(function () {});
      throw error;
    });
    return backgroundPromise;
  }

  window.SlideSciOpen = function () {
    return openPane(pending.tab);
  };
  window.SlideSciOpenEditing = openAt('排版与编辑');
  window.SlideSciOpenImages = openAt('图片');
  window.SlideSciOpenScience = openAt('科学图片');
  window.SlideSciOpenContent = openAt('内容与公式');
  window.SlideSciOpenLibrary = openAt('素材库', { panel: 'library.browser' });
  window.SlideSciOpenAI = openAt('AI 助手', { panel: 'ai.chat' });
  window.SlideSciOpenExport = openAt('导出');
  window.SlideSciOpenHelp = openAt('设置与帮助');

  /* 排列与尺寸 */
  window.SlideSciAlignImages = openAt('排版与编辑', { action: 'editing.grid' });
  window.SlideSciAlignLeft = openAt('排版与编辑', { action: 'editing.align.left', execute: true });
  window.SlideSciAlignHorizontalCenter = openAt('排版与编辑', { action: 'editing.align.horizontalCenter', execute: true });
  window.SlideSciAlignRight = openAt('排版与编辑', { action: 'editing.align.right', execute: true });
  window.SlideSciAlignTop = openAt('排版与编辑', { action: 'editing.align.top', execute: true });
  window.SlideSciAlignVerticalCenter = openAt('排版与编辑', { action: 'editing.align.verticalCenter', execute: true });
  window.SlideSciAlignBottom = openAt('排版与编辑', { action: 'editing.align.bottom', execute: true });
  window.SlideSciSpacingHorizontal = openAt('排版与编辑', { action: 'editing.spacing.horizontal' });
  window.SlideSciSpacingVertical = openAt('排版与编辑', { action: 'editing.spacing.vertical' });
  window.SlideSciDistributeHorizontal = openAt('排版与编辑', { action: 'editing.distribute.horizontal', execute: true });
  window.SlideSciDistributeVertical = openAt('排版与编辑', { action: 'editing.distribute.vertical', execute: true });
  window.SlideSciCopyDimensions = openAt('排版与编辑', { action: 'editing.copy.dimensions', execute: true });
  window.SlideSciPasteDimensions = openAt('排版与编辑', { action: 'editing.paste.dimensions', execute: true });
  window.SlideSciCopyWidth = openAt('排版与编辑', { action: 'editing.copy.width', execute: true });
  window.SlideSciPasteWidth = openAt('排版与编辑', { action: 'editing.paste.width', execute: true });
  window.SlideSciCopyHeight = openAt('排版与编辑', { action: 'editing.copy.height', execute: true });
  window.SlideSciPasteHeight = openAt('排版与编辑', { action: 'editing.paste.height', execute: true });
  window.SlideSciCopyPosition = openAt('排版与编辑', { action: 'editing.copyPosition' });
  window.SlideSciPastePosition = openAt('排版与编辑', { action: 'editing.pastePosition', execute: true });
  window.SlideSciSwapPosition = openAt('排版与编辑', { action: 'editing.swapPosition' });
  window.SlideSciCopyRelativePosition = openAt('排版与编辑', { action: 'editing.copyRelativePosition' });
  window.SlideSciPasteRelativePosition = openAt('排版与编辑', { action: 'editing.pasteRelativePosition', execute: true });
  window.SlideSciCopyShapeFormat = openAt('排版与编辑', { action: 'editing.copyShapeFormat' });
  window.SlideSciPasteShapeFormat = openAt('排版与编辑', { action: 'editing.pasteShapeFormat', execute: true });
  window.SlideSciCopyTextFormat = openAt('排版与编辑', { action: 'editing.copyTextFormat' });
  window.SlideSciPasteTextFormat = openAt('排版与编辑', { action: 'editing.pasteTextFormat', execute: true });
  window.SlideSciCopyGroupFormat = openAt('排版与编辑', { action: 'editing.copyGroupFormat', execute: true });
  window.SlideSciPasteGroupFormat = openAt('排版与编辑', { action: 'editing.pasteGroupFormat', execute: true });
  window.SlideSciCopyCrop = openAt('排版与编辑', { action: 'editing.copyCrop', execute: true });
  window.SlideSciPasteCrop = openAt('排版与编辑', { action: 'editing.pasteCrop', execute: true });
  window.SlideSciTitlesBelow = openAt('排版与编辑', { action: 'editing.addTitlesBelow' });
  window.SlideSciTitlesAbove = openAt('排版与编辑', { action: 'editing.addTitlesAbove' });
  window.SlideSciTitles = window.SlideSciTitlesBelow;
  window.SlideSciLabels = openAt('排版与编辑', { action: 'editing.addLabels' });
  window.SlideSciUpdateLabels = openAt('排版与编辑', { action: 'editing.updateLabels' });

  /* 图片处理 */
  window.SlideSciInsertImages = openAt('图片', { action: 'images.insert', execute: true });
  window.SlideSciPasteImages = openAt('图片', { action: 'images.pasteFiles', execute: true });
  window.SlideSciPseudoColor = openAt('图片', { panel: 'images.pseudoColor' });
  window.SlideSciMergeChannels = openAt('图片', { action: 'images.merge', execute: true });
  window.SlideSciRevealSource = openAt('图片', { action: 'images.reveal', execute: true });

  /* 科研标注 */
  window.SlideSciScaleBar = openAt('科学图片', { panel: 'science.scale' });
  window.SlideSciZoomImages = openAt('科学图片', { panel: 'science.zoom' });
  window.SlideSciSetFov = openAt('科学图片', { action: 'science.setFov' });
  window.SlideSciReadFov = openAt('科学图片', { action: 'science.readFov', execute: true });

  /* 内容插入 */
  window.SlideSciMarkdown = openAt('内容与公式', { action: 'content.markdown' });
  window.SlideSciLatex = openAt('内容与公式', { action: 'content.latex' });
  window.SlideSciCode = openAt('内容与公式', { action: 'content.code' });
  window.SlideSciLatexSvg = openAt('内容与公式', { action: 'content.latexSvg' });
  window.SlideSciSvg = openAt('内容与公式', { action: 'content.svg' });
  window.SlideSciRichText = openAt('内容与公式', { action: 'content.richText' });

  /* 素材与导出 */
  window.SlideSciLibrarySave = openAt('素材库', { action: 'library.save' });
  window.SlideSciExportSlides = openAt('导出', { action: 'export.slides' });
  window.SlideSciExportOriginal = openAt('导出', { action: 'export.originalImages', execute: true });
  window.SlideSciCopyOriginal = openAt('导出', { action: 'export.copyOriginal', execute: true });
  window.SlideSciCopySelection = openAt('导出', { action: 'export.copySelection', execute: true });

  /* 助手与设置 */
  window.SlideSciAbout = openAt('设置与帮助');
  window.SlideSciStartBackground = startBackground;

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', startBackground, { once: true });
  } else {
    startBackground().catch(function () {});
  }
  window.addEventListener('load', function () {
    startBackground().catch(function () {});
  }, { once: true });
})();
