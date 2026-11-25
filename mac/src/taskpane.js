import { createServices } from './services.js';
import { connectPowerPoint } from './powerpoint.js';
import { connectWps } from './wps.js';
import * as editingFeature from './features/editing.js';
import * as imagesFeature from './features/images.js';
import * as scienceFeature from './features/science.js';
import * as contentFeature from './features/content.js';
import * as libraryFeature from './features/library.js';
import * as exportFeature from './features/export.js';
import { startBackgroundTracking } from './background.js';

const SECTION_ORDER = ['排版与编辑', '图片', '科学图片', '内容与公式', '素材库', '导出', '设置与帮助'];
const PANE_REQUEST_KEY = 'slidesci.pane.request';
const EVENT_NAME = 'slidesci:activate';
const PARAMETERS = new URLSearchParams(location.search);
const HOST_PARAMETER = PARAMETERS.get('host');
const TAB_PARAMETER = PARAMETERS.get('tab');
const ACTION_PARAMETER = PARAMETERS.get('action');
const PANEL_PARAMETER = PARAMETERS.get('panel');
const EXECUTE_PARAMETER = PARAMETERS.get('execute') === 'true';
const MODULES = [editingFeature, imagesFeature, scienceFeature, contentFeature, libraryFeature, exportFeature];
const DEFAULT_FEATURES = {
  '排版与编辑': 'action:editing.grid',
  '图片': 'action:images.insert',
  '科学图片': 'panel:science.scale',
  '内容与公式': 'action:content.markdown',
  '素材库': 'panel:library.browser',
  '导出': 'action:export.slides',
  '设置与帮助': 'panel:export.preferences'
};

const dom = {
  workspace: document.querySelector('#workspace'),
  panels: document.querySelector('#panels'),
  tabs: document.querySelector('#tabs'),
  hostLabel: document.querySelector('#host-label'),
  connection: document.querySelector('#connection'),
  connectionText: document.querySelector('#connection-text'),
  retryConnection: document.querySelector('#retry-connection'),
  selectionSummary: document.querySelector('#selection-summary'),
  refreshSelection: document.querySelector('#refresh-selection'),
  progressRegion: document.querySelector('#progress-region'),
  progressBar: document.querySelector('#progress-bar'),
  progressText: document.querySelector('#progress-text'),
  cancelAction: document.querySelector('#cancel-action'),
  notifications: document.querySelector('#notifications'),
  startupError: document.querySelector('#startup-error'),
  startupErrorText: document.querySelector('#startup-error-text'),
  retryStartup: document.querySelector('#retry-startup'),
  versionLabel: document.querySelector('#version-label'),
  fonts: document.querySelector('#slidesci-fonts')
};

const state = {
  target: null,
  services: null,
  host: null,
  connected: false,
  connecting: false,
  busy: false,
  activeRun: null,
  scienceReady: false,
  scienceCleanup: null,
  hostUnsubscribe: null,
  background: null,
  sections: new Map(),
  panelState: new Map(),
  cards: new Map(),
  currentTab: null,
  currentFeature: null,
  pendingRun: null,
  targetSeq: 0,
  lastTargetKey: '',
  lastTargetAt: 0,
  summaryTimer: null,
  summaryError: '',
  summaryPromise: null,
  fontsLoaded: false,
  fontsFailed: false,
  lastNotice: { text: '', at: 0 },
  lastPaneRequest: 0,
  panePoll: null,
  version: null
};

function element(tag, properties = {}, text) {
  const node = document.createElement(tag);
  for (const [name, value] of Object.entries(properties)) {
    if (name === 'className') node.className = value;
    else if (name === 'dataset') Object.assign(node.dataset, value);
    else if (name in node) node[name] = value;
    else node.setAttribute(name, String(value));
  }
  if (text !== undefined) node.textContent = text;
  return node;
}

function slug(value) {
  return String(value).replace(/[^A-Za-z0-9\u4e00-\u9fff]+/g, '-').replace(/^-+|-+$/g, '') || 'item';
}

function resolveTarget() {
  if (HOST_PARAMETER === 'wps' || HOST_PARAMETER === 'powerpoint') return HOST_PARAMETER;
  if (globalThis.Application) return 'wps';
  if (globalThis.Office) return 'powerpoint';
  return null;
}

function notify(message, kind = 'info') {
  const text = String(message ?? '').trim();
  if (!text) return;
  if (text === state.lastNotice.text && Date.now() - state.lastNotice.at < 2500) return;
  state.lastNotice = { text, at: Date.now() };
  const node = element('div', { className: `notice notice-${kind}` }, text);
  dom.notifications.append(node);
  while (dom.notifications.children.length > 4) dom.notifications.firstElementChild.remove();
  const timeout = kind === 'error' ? 15000 : kind === 'warning' ? 10000 : 7000;
  setTimeout(() => node.remove(), timeout);
}

function setConnection(status, text) {
  dom.connection.dataset.state = status;
  dom.connectionText.textContent = text;
  dom.retryConnection.hidden = status !== 'error';
}

function showProgress(text) {
  dom.progressRegion.hidden = false;
  dom.progressText.textContent = text;
  dom.progressBar.setAttribute('aria-valuetext', text);
  dom.cancelAction.disabled = false;
}

function updateProgress(text) {
  if (dom.progressRegion.hidden) return;
  dom.progressText.textContent = text;
  dom.progressBar.setAttribute('aria-valuetext', text);
}

function hideProgress() {
  dom.progressRegion.hidden = true;
  dom.cancelAction.disabled = false;
}

function showStartupError(message) {
  dom.startupError.hidden = false;
  dom.startupErrorText.textContent = message;
  dom.panels.hidden = true;
}

function clearStartupError() {
  dom.startupError.hidden = true;
  dom.startupErrorText.textContent = '';
  dom.panels.hidden = false;
}

/* ---------------------------------- Sections and tabs ---------------------------------- */

function collectDefinitions() {
  const actions = [];
  const panels = [];
  const actionIds = new Set();
  const panelIds = new Set();
  for (const module of MODULES) {
    for (const action of module.actions || []) {
      if (!action?.id || actionIds.has(action.id)) throw new Error('功能定义重复：' + action?.id);
      actionIds.add(action.id);
      actions.push(action);
    }
    for (const panel of module.panels || []) {
      if (!panel?.id || panelIds.has(panel.id)) throw new Error('面板定义重复：' + panel?.id);
      panelIds.add(panel.id);
      panels.push(panel);
    }
  }
  const sections = new Map();
  const ensure = name => {
    const label = String(name || '其他功能');
    if (!sections.has(label)) sections.set(label, { label, actions: [], panels: [] });
    return sections.get(label);
  };
  for (const action of actions) ensure(action.section).actions.push(action);
  for (const panel of panels) ensure(panel.section).panels.push(panel);
  const ordered = [...sections.keys()].sort((a, b) => {
    const ia = SECTION_ORDER.indexOf(a);
    const ib = SECTION_ORDER.indexOf(b);
    return (ia < 0 ? 999 : ia) - (ib < 0 ? 999 : ib);
  });
  state.sections = new Map(ordered.map(name => [name, sections.get(name)]));
}

function groupLabel(action) {
  if (action.section !== '排版与编辑') return action.section;
  const id = action.id;
  if (id === 'editing.grid') return '图片排列';
  if (/^editing\.(?:align|spacing|distribute)\./.test(id)) return '对齐、间距与分布';
  if (/Position|RelativePosition/.test(id)) return '位置';
  if (/\.(?:dimensions|width|height|crop)$/.test(id)) return '尺寸与裁剪';
  if (/Format/.test(id)) return '格式复制';
  if (/Titles|Labels/.test(id)) return '标题与标签';
  if (/scaleTogether|selectAllText/.test(id)) return '文字与全选';
  return '其他编辑操作';
}

function groupedActions(actions) {
  const groups = new Map();
  for (const action of actions) {
    const label = groupLabel(action);
    if (!groups.has(label)) groups.set(label, []);
    groups.get(label).push(action);
  }
  return groups;
}

function renderTabs() {
  dom.tabs.replaceChildren();
  dom.panels.replaceChildren();
  for (const section of state.sections.values()) {
    const tabId = `tab-${slug(section.label)}`;
    const panelId = `panel-${slug(section.label)}`;
    const tab = element('button', {
      type: 'button',
      role: 'tab',
      id: tabId,
      'aria-controls': panelId,
      'aria-selected': 'false',
      tabIndex: -1
    }, section.label);
    tab.addEventListener('click', () => {
      state.pendingRun = null;
      activateTab(section.label);
    });
    dom.tabs.append(tab);

    const panel = element('section', {
      role: 'tabpanel',
      id: panelId,
      'aria-labelledby': tabId,
      tabIndex: -1
    });
    panel.hidden = true;
    panel.append(element('h2', { className: 'visually-hidden' }, section.label));
    const view = element('div', { className: 'feature-view' });
    panel.append(buildFeatureSelector(section), view);
    section.view = view;
    dom.panels.append(panel);
  }
  dom.tabs.addEventListener('keydown', onTabKeydown);
}

function buildFeatureSelector(section) {
  const id = `feature-select-${slug(section.label)}`;
  const wrapper = element('div', { className: 'feature-select' });
  const label = element('label', { className: 'feature-select-label', for: id }, '功能');
  const select = element('select', { id, className: 'feature-select-control' });
  select.setAttribute('aria-label', `${section.label}功能`);
  for (const [group, actions] of groupedActions(section.actions)) {
    const options = actions.map(action => element('option', { value: featureValue('action', action.id) }, action.label));
    if (group === section.label) select.append(...options);
    else {
      const optgroup = element('optgroup', { label: group });
      optgroup.append(...options);
      select.append(optgroup);
    }
  }
  if (section.panels.length) {
    const optgroup = element('optgroup', { label: '面板与预览' });
    for (const descriptor of section.panels) {
      optgroup.append(element('option', { value: featureValue('panel', descriptor.id) }, descriptor.label));
    }
    select.append(optgroup);
  }
  select.addEventListener('change', () => {
    state.pendingRun = null;
    state.targetSeq += 1;
    selectFeature(section.label, select.value, { remember: true });
  });
  wrapper.append(label, select);
  section.select = select;
  return wrapper;
}

function featureValue(kind, id) {
  return `${kind}:${id}`;
}

function describeFeature(section, value) {
  const text = String(value ?? '');
  const separator = text.indexOf(':');
  if (separator < 0) return null;
  const kind = text.slice(0, separator);
  const id = text.slice(separator + 1);
  if (!id) return null;
  if (kind === 'action' && section.actions.some(action => action.id === id)) return { kind, id };
  if (kind === 'panel' && section.panels.some(panel => panel.id === id)) return { kind, id };
  return null;
}

function featureSettingKey(label) {
  return `ui.feature.${label}`;
}

function resolveFeature(section) {
  const candidates = [state.services?.settings.get(featureSettingKey(section.label), ''), DEFAULT_FEATURES[section.label]];
  for (const candidate of candidates) {
    if (describeFeature(section, candidate)) return candidate;
  }
  if (section.actions.length) return featureValue('action', section.actions[0].id);
  if (section.panels.length) return featureValue('panel', section.panels[0].id);
  return '';
}

function rememberFeature(label, value) {
  if (!state.services) return;
  try {
    state.services.settings.set(featureSettingKey(label), value);
  } catch (error) {
    notify('无法保存当前功能选择：' + error.message, 'warning');
  }
}

function selectFeature(sectionLabel, value, { focus = false, remember = false, highlight = false } = {}) {
  const section = state.sections.get(sectionLabel);
  if (!section?.view) return false;
  const resolved = describeFeature(section, value) ? value : resolveFeature(section);
  const feature = describeFeature(section, resolved);
  if (!feature) return false;
  state.currentFeature = { ...feature, section: sectionLabel };
  if (remember) rememberFeature(sectionLabel, resolved);
  if (section.select && section.select.value !== resolved) section.select.value = resolved;
  let node = null;
  if (feature.kind === 'action') {
    const action = section.actions.find(item => item.id === feature.id);
    let entry = state.cards.get(action.id);
    if (!entry) {
      buildActionCard(action);
      entry = state.cards.get(action.id);
    }
    node = entry.card;
    if (node.parentElement !== section.view) section.view.append(node);
  } else {
    const descriptor = section.panels.find(item => item.id === feature.id);
    node = panelEntry(descriptor).node;
    if (node.parentElement !== section.view) section.view.append(node);
    if (state.connected) mountPanel(descriptor);
  }
  for (const child of section.view.children) child.hidden = child !== node;
  if (focus) focusFeature(node);
  if (highlight) pulseTarget(node);
  return true;
}

function focusFeature(node) {
  const target = node.querySelector('input:not([type="checkbox"]):not([type="radio"]), select, textarea, button[type="submit"]') || node;
  target.focus?.({ preventScroll: true });
  node.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
}

function pulseTarget(node) {
  node.classList.add('ribbon-target');
  setTimeout(() => node.classList.remove('ribbon-target'), 1800);
}

function mountSelectedPanel() {
  const feature = state.currentFeature;
  if (!feature || feature.kind !== 'panel' || !state.connected) return;
  const section = state.sections.get(feature.section);
  const descriptor = section?.panels.find(item => item.id === feature.id);
  if (descriptor) mountPanel(descriptor);
}

function onTabKeydown(event) {
  const tabs = [...dom.tabs.querySelectorAll('[role="tab"]')];
  const index = tabs.indexOf(document.activeElement);
  if (index < 0) return;
  let next = null;
  if (event.key === 'ArrowRight') next = tabs[(index + 1) % tabs.length];
  else if (event.key === 'ArrowLeft') next = tabs[(index - 1 + tabs.length) % tabs.length];
  else if (event.key === 'Home') next = tabs[0];
  else if (event.key === 'End') next = tabs[tabs.length - 1];
  if (!next) return;
  event.preventDefault();
  next.focus();
  activateTab(next.textContent, { focusTab: false });
}

function activateTab(label, { focusTab = false, feature = null } = {}) {
  const section = state.sections.get(label);
  if (!section) return false;
  state.currentTab = label;
  state.pendingRun = null;
  for (const tab of dom.tabs.querySelectorAll('[role="tab"]')) {
    const selected = tab.textContent === label;
    tab.setAttribute('aria-selected', String(selected));
    tab.tabIndex = selected ? 0 : -1;
    if (selected && focusTab) tab.focus();
  }
  for (const panelNode of dom.panels.querySelectorAll('[role="tabpanel"]')) {
    panelNode.hidden = panelNode.getAttribute('aria-labelledby') !== `tab-${slug(label)}`;
  }
  if (state.services) {
    try {
      state.services.settings.set('ui.tab', label);
    } catch (error) {
      notify('无法保存当前分类：' + error.message, 'warning');
    }
  }
  selectFeature(label, feature || resolveFeature(section), { remember: false });
  scheduleSummary(200);
  window.dispatchEvent(new Event('resize'));
  return true;
}

/* ---------------------------------- Ribbon and host targets ---------------------------------- */

function normalizeTarget(raw) {
  if (typeof raw === 'string') {
    return state.sections.has(raw) ? { tab: raw, action: null, panel: null, execute: false, at: null } : null;
  }
  if (!raw || typeof raw !== 'object') return null;
  if (typeof raw.tab !== 'string' || !state.sections.has(raw.tab)) return null;
  return {
    tab: raw.tab,
    action: typeof raw.action === 'string' && raw.action ? raw.action : null,
    panel: typeof raw.panel === 'string' && raw.panel ? raw.panel : null,
    execute: raw.execute === true || raw.execute === 'true' || raw.execute === 1,
    at: Number.isFinite(raw.at) ? raw.at : null
  };
}

function applyTarget(raw) {
  const target = normalizeTarget(raw);
  if (!target) return false;
  const now = Date.now();
  const key = JSON.stringify([target.tab, target.action, target.panel, target.execute]);
  if (target.at !== null) {
    if (target.at <= state.lastTargetAt) return false;
    state.lastTargetAt = target.at;
  } else {
    if (key === state.lastTargetKey && now - state.lastTargetAt < (target.execute ? 150 : 400)) return false;
    state.lastTargetAt = now;
  }
  state.lastTargetKey = key;
  const section = state.sections.get(target.tab);
  const wanted = target.action ? featureValue('action', target.action)
    : target.panel ? featureValue('panel', target.panel) : null;
  const exact = describeFeature(section, wanted);
  activateTab(target.tab, { feature: exact ? wanted : null });
  state.targetSeq += 1;
  state.pendingRun = null;
  const feature = state.currentFeature;
  selectFeature(
    target.tab,
    feature ? featureValue(feature.kind, feature.id) : null,
    { focus: true, highlight: true, remember: Boolean(exact) }
  );
  if (target.execute && exact) requestFeatureRun(state.targetSeq);
  return true;
}

function requestFeatureRun(seq) {
  const feature = state.currentFeature;
  if (!feature || feature.kind !== 'action') return;
  const entry = state.cards.get(feature.id);
  if (!entry || entry.controls.length) return;
  state.pendingRun = { seq, actionId: feature.id };
  flushPendingRun();
}

function flushPendingRun() {
  const pending = state.pendingRun;
  if (!pending) return;
  const feature = state.currentFeature;
  if (pending.seq !== state.targetSeq || !feature || feature.kind !== 'action' || feature.id !== pending.actionId) {
    state.pendingRun = null;
    return;
  }
  if (!state.connected || !state.host || state.busy) return;
  const entry = state.cards.get(pending.actionId);
  state.pendingRun = null;
  if (!entry || entry.controls.length) return;
  runAction(entry.action, entry.card, entry.controls);
}

/* ---------------------------------- Action forms ---------------------------------- */

function buildActionCard(action) {
  const card = element('article', { className: 'action-card', dataset: { actionId: action.id } });
  const head = element('div', { className: 'action-head' });
  head.append(element('h4', {}, action.label));
  if (action.description) head.append(element('p', { className: 'action-description' }, action.description));
  card.append(head);

  const form = element('form', { className: 'action-form', noValidate: true });
  form.setAttribute('novalidate', '');
  const fieldsNode = element('div', { className: 'fields' });
  const saved = action.fields?.length ? readSaved(action) : {};
  const controls = (action.fields || []).map(field => buildField(action, field, saved));
  for (const control of controls) fieldsNode.append(control.wrapper);
  if (controls.length) form.append(fieldsNode);
  const run = element('button', { type: 'submit', className: 'primary' }, '执行');
  run.disabled = state.busy || !state.connected;
  const note = element('span', { className: 'action-note' }, '');
  const buttons = element('div', { className: 'action-buttons' });
  buttons.append(run);
  if (!controls.length) note.textContent = '选择对象后执行';
  form.append(buttons);
  if (note.textContent) form.append(note);
  form.addEventListener('submit', event => {
    event.preventDefault();
    runAction(action, card, controls);
  });
  card.append(form);
  state.cards.set(action.id, { action, card, controls });
  return card;
}

function readSaved(action) {
  const key = action.settingsKey || `action.${action.id}`;
  const value = state.services?.settings.get(key, {});
  return value && typeof value === 'object' && !Array.isArray(value) ? value : {};
}

function initialValue(field, saved) {
  const value = Object.prototype.hasOwnProperty.call(saved, field.name) ? saved[field.name] : field.default;
  if (field.type === 'select') {
    const allowed = (field.options || []).map(option => String(option.value));
    return allowed.includes(String(value)) ? String(value) : String(field.default ?? allowed[0] ?? '');
  }
  if (field.type === 'checkbox') return value === undefined ? Boolean(field.default) : Boolean(value);
  return value ?? '';
}

function buildField(action, field, saved) {
  const id = `${slug(action.id)}-${slug(field.name)}`;
  const wrapper = element('div', { className: 'field' });
  const error = element('p', { className: 'field-error', role: 'alert' });
  error.hidden = true;
  const errorId = `${id}-error`;
  error.id = errorId;
  const control = {
    field,
    wrapper,
    input: null,
    read: () => '',
    validate: () => null,
    clearError,
    setError
  };

  function describeId() {
    return error.hidden ? '' : errorId;
  }

  function clearError() {
    wrapper.dataset.invalid = 'false';
    error.hidden = true;
    error.textContent = '';
    control.input?.removeAttribute?.('aria-invalid');
    if (control.input?.setAttribute) control.input.setAttribute('aria-describedby', describeId());
  }

  function setError(message) {
    wrapper.dataset.invalid = 'true';
    error.textContent = message;
    error.hidden = false;
    control.input?.setAttribute?.('aria-invalid', 'true');
    control.input?.setAttribute?.('aria-describedby', errorId);
  }

  if (field.type === 'checkbox') {
    const label = element('label', { for: id });
    const input = element('input', { type: 'checkbox', id, name: field.name });
    input.checked = Boolean(initialValue(field, saved));
    label.append(input, element('span', {}, field.label));
    wrapper.append(label, error);
    control.input = input;
    control.read = () => input.checked;
    input.addEventListener('change', clearError);
    return control;
  }

  if (field.type === 'file' || field.type === 'files') {
    const label = element('span', { className: 'field-label', id: `${id}-label` }, field.label);
    if (field.required) label.append(element('span', { className: 'field-required' }, '必填'));
    const buttons = element('div', { className: 'file-picker' });
    const pick = element('button', { type: 'button', className: 'pick-file' }, '选择文件…');
    pick.disabled = state.busy || !state.connected;
    const clear = element('button', { type: 'button', className: 'quiet', hidden: true }, '清除');
    const list = element('ul', { className: 'file-list', 'aria-live': 'polite' });
    buttons.append(pick, clear);
    wrapper.setAttribute('role', 'group');
    wrapper.setAttribute('aria-labelledby', `${id}-label`);
    if (field.required) wrapper.setAttribute('aria-required', 'true');
    wrapper.append(label, buttons, list, error);
    let files = Array.isArray(saved[field.name]) ? saved[field.name] : [];
    renderFiles();
    pick.addEventListener('click', async () => {
      pick.disabled = true;
      try {
        const picked = await state.services.pickFiles({ kind: fieldKind(field), multiple: field.type === 'files' || Boolean(field.multiple) });
        if (!picked?.length) return;
        files = field.type === 'file' ? [picked[0]] : [...files, ...picked.filter(file => !files.some(item => item.path === file.path))];
        renderFiles();
        clearError();
      } catch (error) {
        notify('选择文件失败：' + error.message, 'error');
      } finally {
        pick.disabled = state.busy || !state.connected;
      }
    });
    clear.addEventListener('click', () => {
      files = [];
      renderFiles();
    });
    function renderFiles() {
      list.replaceChildren();
      for (const file of files) list.append(element('li', {}, file.name || '已选文件'));
      clear.hidden = files.length === 0;
    }
    control.input = pick;
    control.read = () => field.type === 'file' ? files[0] : files;
    control.validate = () => field.required && files.length === 0 ? `请选择${field.label}。` : null;
    control.clearFiles = () => { files = []; renderFiles(); };
    return control;
  }

  const label = element('label', { for: id });
  label.append(element('span', {}, field.label));
  if (field.required) label.append(element('span', { className: 'field-required' }, '必填'));
  let input;
  if (field.type === 'textarea') {
    input = element('textarea', { id, name: field.name, rows: 5 });
    input.value = String(initialValue(field, saved));
  } else if (field.type === 'select') {
    input = element('select', { id, name: field.name });
    for (const option of field.options || []) {
      input.append(element('option', { value: String(option.value) }, option.label ?? String(option.value)));
    }
    input.value = String(initialValue(field, saved));
  } else {
    const type = ['text', 'number', 'color'].includes(field.type) ? field.type : 'text';
    input = element('input', { id, name: field.name, type });
    for (const key of ['min', 'max', 'step']) if (field[key] !== undefined) input.setAttribute(key, String(field[key]));
    if (field.type === 'number') input.inputMode = field.step && !Number.isInteger(field.step) ? 'decimal' : 'numeric';
    input.value = String(initialValue(field, saved));
  }
  wrapper.append(label, input);
  if (field.required) input.setAttribute('aria-required', 'true');
  if (field.min !== undefined || field.max !== undefined) {
    const hint = element('p', { className: 'field-hint' }, `范围：${field.min ?? '不限'} 至 ${field.max ?? '不限'}`);
    wrapper.append(hint);
  }
  wrapper.append(error);
  control.input = input;
  control.read = () => field.type === 'checkbox' ? input.checked : input.value;
  control.validate = () => validateField(field, input);
  const clear = () => clearError();
  input.addEventListener('input', clear);
  input.addEventListener('change', clear);
  if (/字体|font/i.test(`${field.name} ${field.label}`) && input.tagName === 'INPUT') {
    input.setAttribute('list', 'slidesci-fonts');
    input.addEventListener('focus', loadFonts, { once: true });
  }
  return control;
}

function validateField(field, input) {
  const value = String(input.value ?? '');
  if (field.required && !value.trim()) return `请填写${field.label}。`;
  if (!value.trim()) return null;
  if (field.type === 'number') {
    const number = Number(value);
    if (!Number.isFinite(number)) return `请输入有效的${field.label}。`;
    if (field.min !== undefined && number < Number(field.min)) return `${field.label}不能小于 ${field.min}。`;
    if (field.max !== undefined && number > Number(field.max)) return `${field.label}不能大于 ${field.max}。`;
  }
  return null;
}

function fieldKind(field) {
  const accept = String(field.accept || '');
  if (/image|\.png|\.jpe?g|\.tiff?|\.bmp|\.gif|\.svg|\.emf|\.wmf/i.test(accept)) return 'images';
  if (/video|\.mp4|\.mov|\.m4v|\.avi|\.wmv/i.test(accept)) return 'video';
  if (/presentation|\.pptx?|\.potx/i.test(accept)) return 'presentations';
  return 'all';
}

async function loadFonts() {
  if (state.fontsLoaded || !state.connected) return;
  try {
    const fonts = await state.services.fonts();
    state.fontsLoaded = true;
    dom.fonts.replaceChildren(...fonts.map(font => element('option', { value: font })));
  } catch (error) {
    if (!state.fontsFailed) {
      state.fontsFailed = true;
      notify('无法读取系统字体列表，可直接输入字体名称。', 'warning');
      state.services.reportError(error, { operation: 'fonts' });
    }
  }
}

function collectValues(controls) {
  const values = {};
  for (const control of controls) values[control.field.name] = control.read();
  return values;
}

function validateControls(controls) {
  let first = null;
  for (const control of controls) {
    const message = control.validate();
    if (message) {
      control.setError(message);
      first ??= control;
    } else control.clearError();
  }
  if (first) {
    first.input?.focus?.();
    notify('请检查表单中标出的内容。', 'error');
    return false;
  }
  return true;
}

function persistAction(action, values) {
  if (!action.fields?.length) return;
  const key = action.settingsKey || `action.${action.id}`;
  const saved = {};
  for (const field of action.fields) {
    if (field.type === 'file' || field.type === 'files') continue;
    if (Object.prototype.hasOwnProperty.call(values, field.name)) saved[field.name] = values[field.name];
  }
  try {
    state.services.settings.set(key, saved);
  } catch (error) {
    notify('操作已完成，但设置未保存：' + error.message, 'warning');
  }
}

function setBusy(busy, activeCard = null) {
  state.busy = busy;
  dom.workspace.setAttribute('aria-busy', String(busy));
  for (const card of dom.panels.querySelectorAll('.action-card')) {
    const active = busy && card === activeCard;
    card.dataset.running = String(active);
    card.setAttribute('aria-busy', String(active));
    const submit = card.querySelector('button[type="submit"]');
    if (submit) submit.textContent = active ? '执行中…' : '执行';
  }
  for (const button of dom.panels.querySelectorAll('.action-form button[type="submit"]')) button.disabled = busy || !state.connected;
  for (const button of dom.panels.querySelectorAll('.pick-file')) button.disabled = busy || !state.connected;
  dom.refreshSelection.disabled = busy || !state.connected;
  dom.cancelAction.hidden = !busy;
}

async function runAction(action, card, controls) {
  if (!state.connected || !state.host) {
    notify('尚未连接演示软件，暂时无法执行。', 'error');
    return;
  }
  if (state.busy) {
    notify('请等待当前操作完成或取消后再执行。', 'warning');
    return;
  }
  if (!validateControls(controls)) return;
  const values = collectValues(controls);
  const controller = new AbortController();
  state.activeRun = { controller, action };
  setBusy(true, card);
  showProgress(`${action.label}：正在处理…`);

  const scoped = Object.create(state.services);
  scoped.request = (name, payload, options = {}) => state.services.request(name, payload, { ...options, signal: controller.signal });
  const ctx = makeContext(scoped, {
    signal: controller.signal,
    progress: message => updateProgress(`${action.label}：${message}`)
  });
  let notified = false;
  const notifyOnce = ctx.notify;
  ctx.notify = (message, kind = 'info') => {
    notified = true;
    notifyOnce(message, kind);
  };

  try {
    const result = await action.run(ctx, values);
    if (controller.signal.aborted) {
      notify('已取消等待；当前原生操作可能已经完成，请检查幻灯片。', 'warning');
      return;
    }
    persistAction(action, values);
    if (result?.formValues && typeof result.formValues === 'object') {
      for (const control of controls) {
        if (!Object.prototype.hasOwnProperty.call(result.formValues, control.field.name) || !control.input) continue;
        if (control.field.type === 'checkbox') control.input.checked = Boolean(result.formValues[control.field.name]);
        else control.input.value = String(result.formValues[control.field.name]);
      }
      persistAction(action, { ...values, ...result.formValues });
    }
    if (!notified) {
      const message = typeof result?.message === 'string' && result.message.trim()
        ? result.message
        : `${action.label}完成。`;
      notify(message, 'success');
    }
  } catch (error) {
    if (error?.name === 'AbortError' || controller.signal.aborted) notify('操作已取消。', 'warning');
    else {
      notify(`${action.label}失败：${error.message || error}`, 'error');
      await state.services.reportError(error, { action: action.id });
    }
  } finally {
    state.activeRun = null;
    setBusy(false);
    hideProgress();
    scheduleSummary(150);
    flushPendingRun();
  }
}

function makeContext(services, { signal, progress } = {}) {
  return {
    host: state.host,
    services,
    settings: state.services.settings,
    signal,
    notify,
    progress: progress || (() => {}),
    refresh: () => scheduleSummary(200)
  };
}

/* ---------------------------------- Panels ---------------------------------- */

function panelPlaceholder(text) {
  return element('p', { className: 'panel-placeholder' }, text);
}

function panelEntry(descriptor) {
  const existing = state.panelState.get(descriptor.id);
  if (existing) return existing;
  const node = element('section', {
    className: 'feature-panel',
    id: `panel-view-${slug(descriptor.id)}`,
    tabIndex: -1
  });
  node.setAttribute('aria-label', descriptor.label);
  const body = element('div', { className: 'panel-body' });
  body.append(panelPlaceholder(state.connected ? '正在加载…' : '连接演示软件后可加载。'));
  node.append(body);
  const entry = { descriptor, node, body, mounted: false, mounting: false, cleanup: null, controller: null };
  state.panelState.set(descriptor.id, entry);
  return entry;
}

async function mountPanel(descriptor, { force = false } = {}) {
  const entry = panelEntry(descriptor);
  if (!state.connected) return;
  if ((entry.mounted && !force) || entry.mounting) return;
  entry.mounting = true;
  entry.cleanup?.();
  entry.cleanup = null;
  entry.controller?.abort();
  const controller = new AbortController();
  entry.controller = controller;
  // Panels append their controls beside the placeholder, so it is removed once mounting finishes.
  const placeholder = panelPlaceholder('正在加载…');
  entry.body.replaceChildren(placeholder);
  const scoped = Object.create(state.services);
  scoped.request = (name, payload, options = {}) => state.services.request(name, payload, { ...options, signal: controller.signal });
  const ctx = makeContext(scoped, { signal: controller.signal, progress: message => updateProgress(message) });
  try {
    const cleanup = await descriptor.mount(entry.body, ctx);
    if (controller.signal.aborted || entry.controller !== controller) {
      if (typeof cleanup === 'function') cleanup();
      return;
    }
    placeholder.remove();
    entry.mounted = true;
    entry.cleanup = typeof cleanup === 'function' ? cleanup : null;
  } catch (error) {
    if (error?.name === 'AbortError' || entry.controller !== controller) return;
    entry.mounted = false;
    const box = element('div', { className: 'panel-error' });
    box.append(element('p', {}, `${descriptor.label}加载失败：${error.message}`));
    const retry = element('button', { type: 'button' }, '重新加载面板');
    retry.addEventListener('click', () => mountPanel(descriptor, { force: true }));
    box.append(retry);
    entry.body.replaceChildren(box);
    notify(`${descriptor.label}加载失败：${error.message}`, 'error');
    await state.services.reportError(error, { panel: descriptor.id });
  } finally {
    if (entry.controller === controller) entry.mounting = false;
  }
}

function cleanupPanels({ reset = false } = {}) {
  for (const entry of state.panelState.values()) {
    try {
      entry.cleanup?.();
    } catch (error) {
      console.warn('SlideSCI 面板清理失败：' + error.message);
    }
    entry.cleanup = null;
    entry.controller?.abort();
    entry.controller = null;
    entry.mounted = false;
    entry.mounting = false;
    if (reset && entry.body) entry.body.replaceChildren(panelPlaceholder('正在加载…'));
  }
}

/* ---------------------------------- Connection ---------------------------------- */

function validateHost(host) {
  if (!host || typeof host !== 'object') throw new Error('没有取得演示软件连接。');
  if (typeof host.name !== 'string' || !['powerpoint', 'wps'].includes(host.kind)) {
    throw new Error('演示软件连接组件版本不匹配，请更新 SlideSCI 后重试。');
  }
  for (const name of ['summary', 'align', 'context', 'apply', 'images', 'fragment', 'insertFragment', 'export', 'onChanged']) {
    if (typeof host[name] !== 'function') throw new Error('演示软件连接组件不完整，请更新 SlideSCI 后重试。');
  }
}

async function connect() {
  if (state.connecting) return;
  state.connecting = true;
  state.connected = false;
  setBusy(false);
  clearStartupError();
  setConnection('connecting', '正在连接演示软件…');
  try {
    if (!state.target) throw new Error('请从 PowerPoint 或 WPS 演示的 SlideSCI 选项卡打开此面板。');
    state.hostUnsubscribe?.();
    state.hostUnsubscribe = null;
    await state.background?.stop();
    state.background = null;
    state.scienceCleanup?.();
    state.scienceCleanup = null;
    state.scienceReady = false;
    cleanupPanels({ reset: true });
    state.host?.dispose?.();
    const host = state.target === 'wps'
      ? await Promise.resolve(connectWps(state.services))
      : await connectPowerPoint(state.services);
    validateHost(host);
    if (state.target === 'powerpoint' && Office.context.requirements.isSetSupported('SharedRuntime', '1.1')) {
      await Office.addin.setStartupBehavior(Office.StartupBehavior.load);
    }
    state.host = host;
    state.connected = true;
    dom.hostLabel.textContent = host.name;
    setConnection('ready', `已连接 ${host.name}`);
    if (typeof host.onChanged === 'function') {
      state.hostUnsubscribe = host.onChanged(() => {
        scheduleSummary(350);
      }, { mode: 'selection' });
    }
    mountSelectedPanel();
    setBusy(false);
    flushPendingRun();
    scheduleSummary(0);
    probeLocalService(host);
  } catch (error) {
    state.connected = false;
    state.host = null;
    setConnection('error', '连接失败');
    showStartupError(error.message || '无法连接演示软件，请重试。');
    notify(error.message || '无法连接演示软件，请重试。', 'error');
    await state.services.reportError(error, { operation: 'connect', target: state.target });
  } finally {
    state.connecting = false;
  }
}

async function probeLocalService(host) {
  try {
    const session = await state.services.ready();
    if (host !== state.host) return;
    for (const warning of session?.warnings || []) notify(warning, 'error');
    if (state.services.settings.get('background.autoPasteTracking', false) === true) startTracking();
    loadVersion();
    if (session?.acceptance) {
      const { startAcceptance } = await import('./acceptance.js');
      startAcceptance(makeContext(state.services, {}));
    }
  } catch {
    if (host !== state.host) return;
    setConnection('ready', `已连接 ${host.name}（仅基础组图）`);
  }
}

function startTracking() {
  if (!state.connected || !state.host) return;
  if (state.background && !state.background.stopped) return;
  startBackgroundTracking({
    host: state.target,
    services: state.services,
    hostAdapter: state.host,
    role: 'foreground',
    science: false,
    onStatus: event => {
      if (!event?.message) return;
      if (event.kind === 'success') notify(event.message, 'success');
      else if (event.kind === 'error') notify(event.message, 'error');
    }
  }).then(runtime => {
    state.background = runtime;
  }).catch(error => {
    notify('自动粘贴监听未启动：' + error.message, 'warning');
  });
}

async function stopTracking() {
  const current = state.background;
  state.background = null;
  if (current) await current.stop();
}

async function loadVersion() {
  try {
    const info = await state.services.request('library/info', {});
    state.version = info?.version || null;
    if (info?.version) {
      dom.versionLabel.textContent = `版本 ${info.version}${info.upstreamVersion ? ` · 原版 ${info.upstreamVersion}` : ''}`;
      dom.versionLabel.hidden = false;
    }
  } catch (error) {
    dom.versionLabel.hidden = true;
    state.services.reportError(error, { operation: 'version-info' });
  }
}

function scheduleSummary(delay = 300) {
  clearTimeout(state.summaryTimer);
  state.summaryTimer = setTimeout(() => {
    refreshSummary().catch(error => console.warn('SlideSCI 刷新选择失败：' + error.message));
  }, delay);
}

async function refreshSummary({ manual = false } = {}) {
  if (!state.connected || !state.host || state.busy) return;
  if (!manual && document.visibilityState === 'hidden') return;
  if (state.summaryPromise) return state.summaryPromise;
  const host = state.host;
  state.summaryPromise = (async () => {
    try {
      const context = await host.summary({ slides: false });
      if (host !== state.host) return;
      const selected = Array.isArray(context.selectionIds) ? context.selectionIds.length : 0;
      const slide = Number.isInteger(context.slideIndex) ? context.slideIndex + 1 : 1;
      const name = typeof context.title === 'string' && context.title.trim() ? `${context.title.trim()} · ` : '';
      const text = `${name}第 ${slide} 页 · ${selected ? `已选择 ${selected} 个对象` : '未选择对象'}`;
      if (dom.selectionSummary.textContent !== text) dom.selectionSummary.textContent = text;
      state.summaryError = '';
    } catch (error) {
      if (host !== state.host) return;
      const message = '无法读取当前选择：' + error.message;
      dom.selectionSummary.textContent = message;
      if (manual) notify(message, 'error');
      if (state.summaryError !== message) {
        state.summaryError = message;
        state.services.reportError(error, { operation: 'selection-summary' });
      }
    } finally {
      state.summaryPromise = null;
    }
  })();
  return state.summaryPromise;
}

/* ---------------------------------- Ribbon requests ---------------------------------- */

function readPaneRequest(value) {
  try {
    if (!value) {
      const raw = globalThis.localStorage?.getItem(PANE_REQUEST_KEY);
      if (!raw) return;
      value = JSON.parse(raw);
    }
    if (!value || typeof value.tab !== 'string' || !Number.isFinite(value.at)) return;
    if (value.at <= state.lastPaneRequest || Date.now() - value.at > 120000) return;
    state.lastPaneRequest = value.at;
    applyTarget(value);
    try {
      globalThis.localStorage?.removeItem(PANE_REQUEST_KEY);
    } catch {
      // A request that cannot be cleared is still ignored by its timestamp on this pane.
    }
  } catch {
    // A malformed ribbon request must not block the panel.
  }
}

function initPaneRequests() {
  if (typeof BroadcastChannel === 'function') {
    const channel = new BroadcastChannel(PANE_REQUEST_KEY);
    channel.addEventListener('message', event => readPaneRequest(event.data));
    window.addEventListener('pagehide', () => channel.close(), { once: true });
  }
  window.addEventListener('storage', event => {
    if (event.key === PANE_REQUEST_KEY) readPaneRequest();
  });
  window.addEventListener('message', event => {
    if (event.origin !== location.origin) return;
    const data = event.data;
    if (!data || data.type !== EVENT_NAME) return;
    applyTarget(data.target ?? data);
  });
  window.addEventListener(EVENT_NAME, event => {
    applyTarget(event.detail?.target ?? event.detail);
  });
  document.addEventListener(EVENT_NAME, event => {
    applyTarget(event.detail?.target ?? event.detail);
  });
  state.panePoll = setInterval(() => {
    if (document.visibilityState === 'hidden') return;
    readPaneRequest();
  }, state.target === 'wps' ? 100 : 1200);
  document.addEventListener('visibilitychange', () => {
    if (document.visibilityState === 'hidden') return;
    readPaneRequest();
    scheduleSummary(100);
  });
}

function consumeUrlTarget() {
  if (!ACTION_PARAMETER && !PANEL_PARAMETER && !EXECUTE_PARAMETER) return;
  const target = {
    tab: state.sections.has(TAB_PARAMETER) ? TAB_PARAMETER : state.currentTab,
    action: ACTION_PARAMETER,
    panel: PANEL_PARAMETER,
    execute: EXECUTE_PARAMETER
  };
  const url = new URL(location.href);
  for (const key of ['action', 'panel', 'execute']) url.searchParams.delete(key);
  try {
    history.replaceState(history.state, '', url);
  } catch {
    // A consumed target that cannot be stripped is collapsed by its dispatch dedupe.
  }
  if (target.tab) applyTarget(target);
}

function startRibbonCommands() {
  if (state.target !== 'powerpoint') return;
  const navigate = target => applyTarget(target);
  import('./ribbon.js')
    .then(module => {
      if (typeof module.registerOfficeCommands !== 'function') throw new Error('功能区模块缺少注册接口。');
      return module.registerOfficeCommands(navigate);
    })
    .then(result => {
      if (!result || result.registered !== false) return result;
      const reason = String(result.reason || '');
      if (reason === 'wps-host' || reason.startsWith('host:')) return result;
      notify(`功能区按钮未注册：${reason || '未知原因'}`, 'warning');
      return result;
    })
    .catch(error => {
      console.warn('SlideSCI 功能区注册失败：' + error.message);
      notify('功能区按钮注册失败，可从侧边栏继续使用。', 'warning');
      state.services?.reportError(error, { operation: 'ribbon-register' });
    });
}

/* ---------------------------------- Boot ---------------------------------- */

function wireShell() {
  dom.retryConnection.addEventListener('click', () => connect());
  dom.retryStartup.addEventListener('click', () => connect());
  dom.refreshSelection.addEventListener('click', () => {
    refreshSummary({ manual: true }).catch(error => notify(error.message, 'error'));
  });
  dom.cancelAction.addEventListener('click', () => {
    if (!state.activeRun) return;
    dom.cancelAction.disabled = true;
    updateProgress('正在取消…当前原生操作结束后生效。');
    state.activeRun.controller.abort();
  });
  globalThis.addEventListener?.('slidesci:tracking-toggle', event => {
    if (event?.detail?.enabled) startTracking();
    else stopTracking().catch(error => notify('停止自动粘贴监听失败：' + error.message, 'warning'));
  });
  globalThis.addEventListener?.('pagehide', () => {
    cleanupPanels();
    state.scienceCleanup?.();
    state.hostUnsubscribe?.();
    clearInterval(state.panePoll);
  }, { once: true });
}

async function boot() {
  wireShell();
  state.target = resolveTarget();
  state.services = createServices({ host: state.target || 'powerpoint' });
  try {
    collectDefinitions();
  } catch (error) {
    setConnection('error', '功能加载失败');
    showStartupError(error.message);
    await state.services.reportError(error, { operation: 'definitions' });
    return;
  }
  renderTabs();
  startRibbonCommands();
  initPaneRequests();
  const savedTab = state.services.settings.get('ui.tab', SECTION_ORDER[0]);
  const initial = state.sections.has(TAB_PARAMETER) ? TAB_PARAMETER
    : state.sections.has(savedTab) ? savedTab
      : state.sections.keys().next().value;
  if (initial) activateTab(initial);
  readPaneRequest();
  consumeUrlTarget();
  if (!state.target) {
    setConnection('error', '请从演示软件打开');
    showStartupError('请从 PowerPoint 或 WPS 演示的 SlideSCI 选项卡打开此面板。');
    dom.retryConnection.hidden = false;
    setBusy(false);
    return;
  }
  await connect();
  setBusy(false);
}

boot().catch(async error => {
  setConnection('error', 'SlideSCI 启动失败');
  showStartupError(error.message || 'SlideSCI 启动失败。');
  try {
    await state.services?.reportError(error, { operation: 'boot' });
  } catch (reportError) {
    console.error('SlideSCI 错误上报失败：' + reportError.message);
  }
});
