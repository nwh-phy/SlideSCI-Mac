import { element, button, input, select, statusNode } from './library-ai-ui.js';

export const actions = [];
export const panels = [{ id: 'ai.chat', label: 'AI 助手', section: 'AI 助手', mount: async (container, ctx) => {
  const lifetime = new AbortController(); const { signal } = lifetime;
  const request = (name, value = {}) => ctx.services.request(`ai/${name}`, value);
  // Configuration and conversation list are independent; load them together so
  // opening the panel costs one round trip instead of two.
  const [config, initialHistory] = await Promise.all([request('config'), request('history')]);
  let configuration = config; let conversationId = null; let activeJob = null;
  let busy = false; let stopRequested = false; let applying = false; let mounted = true;
  const root = element('section', undefined, { className: 'ai-sidebar' });
  const heading = element('h3', 'AI 原生编辑助手');
  const status = statusNode(); const messages = element('div', undefined, { className: 'ai-messages', role: 'log', 'aria-live': 'polite', 'aria-label': 'AI对话' });
  const profiles = select('当前AI配置', [], undefined);
  const histories = select('会话历史', [{ value: '', label: '新会话' }], '');
  const prompt = input('消息（Enter发送，Shift+Enter换行）', 'textarea', '', { rows: 4, placeholder: '例如：把选中的文本改成深蓝色，再添加一页空白幻灯片。' });
  const settings = element('details'); settings.append(element('summary', 'AI服务与系统提示词设置'));
  const profileId = select('编辑配置', [], undefined);
  const name = input('配置名称'); const base = input('API Base URL', 'url'); const model = input('模型名称');
  const secret = input('API Key（留空保留钥匙串中的密钥）', 'password', '', { autocomplete: 'new-password', spellcheck: false });
  const showSecret = input('显示本次输入的密钥', 'checkbox');
  showSecret.control.addEventListener('change', () => { secret.control.type = showSecret.control.checked ? 'text' : 'password'; }, { signal });
  const system = input('系统提示词', 'textarea', '', { rows: 8 });
  const keyState = element('p');
  const activate = input('保存后设为当前配置', 'checkbox'); activate.control.checked = true;
  const confirmProfileDelete = input('确认删除此AI配置与对应密钥', 'checkbox');
  function bubble(text, role = 'assistant') {
    const article = element('article', undefined, { className: `ai-message ai-${role}` });
    article.append(element('strong', role === 'user' ? '我' : role === 'error' ? '错误' : role === 'tool' ? '原生编辑结果' : 'AI'));
    const body = element('div', text, { className: 'ai-message-text' }); article.append(body); messages.append(article);
    article.scrollIntoView?.({ block: 'nearest' }); return body;
  }
  function fillProfile() {
    const profile = configuration.profiles.find(item => item.id === profileId.control.value);
    name.control.value = profile?.name || ''; base.control.value = profile?.baseUrl || 'https://api.openai.com/v1'; model.control.value = profile?.model || 'gpt-4o';
    system.control.value = profile?.systemPrompt || configuration.defaultSystemPrompt;
    secret.control.value = ''; secret.control.type = 'password'; showSecret.control.checked = false; confirmProfileDelete.control.checked = false;
    keyState.textContent = profile?.hasKey ? 'API密钥已保存在macOS钥匙串中。' : '尚未配置API密钥。';
  }
  function renderProfiles() {
    const editing = profileId.control.value;
    profiles.control.replaceChildren(...configuration.profiles.map(item => element('option', `${item.name} · ${item.model}`, { value: item.id })));
    profiles.control.value = configuration.activeProfileId;
    profileId.control.replaceChildren(...configuration.profiles.map(item => element('option', item.name, { value: item.id })), element('option', '新增配置…', { value: '' }));
    profileId.control.value = configuration.profiles.some(item => item.id === editing) ? editing : configuration.activeProfileId;
    fillProfile();
  }
  async function renderHistory(prefetched) {
    const { conversations } = prefetched || await request('history');
    histories.control.replaceChildren(element('option', '新会话', { value: '' }), ...conversations.map(item => element('option', `${item.title} · ${new Date(item.updatedAt).toLocaleString()}`, { value: item.id })));
    histories.control.value = conversationId || '';
  }
  function renderConversation(history) {
    messages.replaceChildren();
    for (const message of history.messages) {
      if (message.role === 'user' || message.role === 'assistant') { if (message.content) bubble(message.content, message.role); }
      if (message.role === 'tool') {
        let result;
        try { result = JSON.parse(message.content); } catch { result = { message: message.content }; }
        if (result.error || result.applied !== undefined) bubble(result.error || `已执行${result.applied}条原生命令。`, result.error ? 'error' : 'tool');
      }
    }
  }
  function state(value) {
    busy = value; prompt.control.disabled = value; profiles.control.disabled = value; histories.control.disabled = value;
    send.textContent = value ? '停止' : '发送'; clear.disabled = value; deleteHistory.disabled = value;
    for (const control of settings.querySelectorAll('input,textarea,select,button')) control.disabled = value;
  }
  async function stop() {
    stopRequested = true;
    status.textContent = applying ? '当前原生编辑完成后停止。' : '正在停止…';
    if (activeJob && !applying) await request('cancel', { id: activeJob });
  }
  async function sendMessage() {
    if (busy) { await stop(); return; }
    const message = prompt.control.value.trim(); if (!message) return;
    stopRequested = false; state(true); let cursor = 0; let assistant = null;
    bubble(message, 'user'); prompt.control.value = '';
    try {
      const context = await ctx.host.context({ scope: 'presentation', images: false });
      const started = await request('start', { conversationId, profileId: profiles.control.value, message, context });
      conversationId = started.conversationId; activeJob = started.id; status.textContent = 'AI正在响应…';
      if (stopRequested) await request('cancel', { id: activeJob });
      let finished = false;
      while (!finished && mounted) {
        const polled = await request('poll', { id: activeJob, after: cursor });
        for (const event of polled.events) {
          cursor = event.seq;
          if (event.type === 'delta') { if (!assistant) assistant = bubble(''); assistant.textContent += event.text; }
          else if (event.type === 'tool-error' || event.type === 'error') bubble(event.message, 'error');
          else if (event.type === 'cancelled') status.textContent = event.message;
          else if (event.type === 'tools') {
            if (stopRequested) { await request('cancel', { id: activeJob }); continue; }
            applying = true; const results = [];
            // One full-deck snapshot serves every read-only tool call in this
            // batch; it is discarded only after native commands can mutate it.
            let synced = null;
            const snapshot = () => synced ??= ctx.host.context({ scope: 'presentation', images: false });
            try {
              for (const call of event.calls) {
                if (stopRequested) { results.push({ tool_call_id: call.tool_call_id, content: { ok: false, error: '用户停止，命令未执行。' } }); continue; }
                try {
                  if (call.name === 'edit_presentation') {
                    const validated = await request('validate', { commands: call.commands, context: await snapshot() });
                    synced = null;
                    await ctx.host.apply(validated.commands); await ctx.refresh();
                    results.push({ tool_call_id: call.tool_call_id, content: { ok: true, applied: validated.commands.length } });
                    bubble(`已执行${validated.commands.length}条原生命令。`, 'tool');
                  } else results.push({ tool_call_id: call.tool_call_id, content: { ok: true, context: await snapshot() } });
                } catch (error) {
                  results.push({ tool_call_id: call.tool_call_id, content: { ok: false, error: error.message } }); bubble(error.message, 'error');
                }
              }
              await request('respond', { id: activeJob, results, context: await snapshot() });
              if (stopRequested) await request('cancel', { id: activeJob });
            } finally { applying = false; }
            assistant = null;
          }
        }
        finished = ['complete', 'cancelled', 'error'].includes(polled.status);
        if (!finished) await new Promise(resolve => setTimeout(resolve, 150));
        else if (polled.status === 'complete') status.textContent = 'AI响应已完成。';
        else if (polled.status === 'error') status.textContent = 'AI请求失败，可修改配置或消息后重试。';
      }
    } catch (error) {
      bubble(error.message, 'error'); status.textContent = 'AI请求未完成。';
      if (activeJob) await request('cancel', { id: activeJob }).catch(() => {});
    } finally {
      activeJob = null; applying = false;
      if (mounted) { state(false); await renderHistory(); prompt.control.focus(); }
    }
  }
  const send = button('发送', sendMessage, ctx, signal);
  // Stop must remain clickable while the request is streaming.
  send.addEventListener('click', () => { if (busy) send.disabled = false; }, { signal });
  const clear = button('新会话', async () => { conversationId = null; messages.replaceChildren(); histories.control.value = ''; status.textContent = '已开始新会话。'; }, ctx, signal);
  const confirmHistoryDelete = input('确认删除当前会话历史', 'checkbox');
  const deleteHistory = button('删除当前会话', async () => {
    if (!conversationId) throw new Error('当前没有已保存会话。');
    if (!confirmHistoryDelete.control.checked) throw new Error('请勾选删除会话确认。');
    await request('deleteHistory', { id: conversationId }); conversationId = null; messages.replaceChildren(); confirmHistoryDelete.control.checked = false; await renderHistory();
  }, ctx, signal);
  profiles.control.addEventListener('change', async () => { try { configuration = await request('selectProfile', { id: profiles.control.value }); renderProfiles(); } catch (error) { ctx.notify(error.message, 'error'); } }, { signal });
  profileId.control.addEventListener('change', fillProfile, { signal });
  histories.control.addEventListener('change', async () => { try { conversationId = histories.control.value || null; confirmHistoryDelete.control.checked = false; if (conversationId) renderConversation(await request('history', { id: conversationId })); else messages.replaceChildren(); } catch (error) { ctx.notify(error.message, 'error'); } }, { signal });
  prompt.control.addEventListener('keydown', event => { if (event.key === 'Enter' && !event.shiftKey && !event.isComposing) { event.preventDefault(); sendMessage().catch(error => ctx.notify(error.message, 'error')); } }, { signal });
  settings.append(profileId.wrapper, name.wrapper, base.wrapper, model.wrapper, secret.wrapper, showSecret.wrapper, keyState, system.wrapper,
    button('恢复默认系统提示词', () => { system.control.value = configuration.defaultSystemPrompt; }, ctx, signal),
    button('撤销未保存的配置修改', fillProfile, ctx, signal), activate.wrapper,
    button('保存配置', async () => {
      const payload = { id: profileId.control.value || undefined, name: name.control.value, baseUrl: base.control.value, model: model.control.value, systemPrompt: system.control.value, activate: activate.control.checked };
      if (secret.control.value) payload.secret = secret.control.value;
      secret.control.value = '';
      try { configuration = await request('saveProfile', payload); renderProfiles(); status.textContent = 'AI配置已保存。'; }
      finally { delete payload.secret; }
    }, ctx, signal),
    button('删除钥匙串密钥', async () => { if (!profileId.control.value) throw new Error('请先选择已有配置。'); await request('clearKey', { id: profileId.control.value }); configuration = await request('config'); renderProfiles(); }, ctx, signal),
    confirmProfileDelete.wrapper,
    button('删除配置', async () => { if (!confirmProfileDelete.control.checked) throw new Error('请先勾选删除配置确认。'); configuration = await request('deleteProfile', { id: profileId.control.value }); renderProfiles(); }, ctx, signal)
  );
  const toolbar = element('div', undefined, { className: 'feature-toolbar' }); toolbar.append(send, clear, deleteHistory);
  root.append(heading, profiles.wrapper, settings, histories.wrapper, confirmHistoryDelete.wrapper, messages, prompt.wrapper, toolbar, status); container.append(root);
  renderProfiles(); await renderHistory(initialHistory);
  return () => {
    mounted = false; stopRequested = true; lifetime.abort(); root.remove();
    if (activeJob && !applying) request('cancel', { id: activeJob }).catch(error => ctx.notify(error.message, 'error'));
  };
} }];
