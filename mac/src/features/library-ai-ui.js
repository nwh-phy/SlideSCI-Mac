export function element(tag, text, attributes = {}) {
  const node = document.createElement(tag);
  if (text !== undefined && text !== null) node.textContent = text;
  for (const [name, value] of Object.entries(attributes)) {
    if (name === 'className') node.className = value;
    else if (name in node) node[name] = value;
    else node.setAttribute(name, String(value));
  }
  return node;
}
export function button(text, action, ctx, signal) {
  const node = element('button', text, { type: 'button' });
  node.addEventListener('click', async () => {
    node.disabled = true;
    try { await action(); } catch (error) { ctx.notify(error.message || String(error), 'error'); }
    finally { node.disabled = false; }
  }, { signal });
  return node;
}
export function input(label, type = 'text', value = '', attributes = {}) {
  const wrapper = element('label', label);
  const control = element(type === 'textarea' ? 'textarea' : 'input', undefined, { ...attributes, ...(type === 'textarea' ? {} : { type }), value });
  wrapper.append(control); return { wrapper, control };
}
export function select(label, options, value) {
  const wrapper = element('label', label); const control = element('select');
  for (const item of options) control.append(element('option', item.label, { value: item.value }));
  if (value !== undefined) control.value = value;
  wrapper.append(control); return { wrapper, control };
}
export function statusNode() { return element('p', '', { role: 'status', 'aria-live': 'polite' }); }
export async function descriptorData(file, services) {
  if (!file) throw new Error('宿主没有返回需要读取的文件。');
  if (file.data) return file;
  if (file.path) return { ...file, ...await services.request('files/read', { path: file.path }) };
  throw new Error('宿主没有返回文件数据或本地文件路径。');
}
export function filesOf(result) {
  if (Array.isArray(result)) return result;
  if (Array.isArray(result?.files)) return result.files;
  if (result?.data || result?.path) return [result];
  throw new Error('宿主没有返回导出文件。');
}
