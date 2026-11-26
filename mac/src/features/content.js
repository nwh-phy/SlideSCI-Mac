const section = '内容与公式';
const languages = ['python', 'matlab', 'javascript', 'html', 'css', 'r', 'fortran', 'c', 'cpp', 'csharp', 'java', 'json', 'bash', 'sql', 'yaml', 'text'].map(value => ({ value, label: value === 'text' ? '纯文本' : value }));
const themes = [{ value: 'light', label: '浅色' }, { value: 'dark', label: '深色' }];
const source = label => ({ name: 'source', label, type: 'textarea', required: true });
const width = { name: 'width', label: '内容宽度（磅）', type: 'number', default: 500, min: 20, step: 1 };
const size = { name: 'fontSize', label: '字号（磅）', type: 'number', default: 18, min: 1, max: 400, step: 1 };
const family = { name: 'fontFamily', label: '字体', type: 'text', default: 'Arial' };
const theme = { name: 'theme', label: '代码主题', type: 'select', default: 'saved', options: [{ value: 'saved', label: '使用已保存的主题' }, ...themes] };
const preferredTheme = (ctx, value) => value && value !== 'saved' ? value : ctx.settings.get('content.codeTheme', 'light');

function warn(ctx, result) {
  for (const warning of result.warnings || []) ctx.notify(warning.message, 'warning');
}

// Renders only need the page size, which the host summary already carries, so
// no slide snapshot or full-page geometry walk is needed.
async function pageSize(ctx) {
  const { width, height } = await ctx.host.summary({ slides: false });
  return { slideWidth: width, slideHeight: height };
}

const CONVERSION_LIMIT = 4;
async function convertAll(items, convert) {
  const results = new Array(items.length);
  let cursor = 0;
  const worker = async () => {
    while (cursor < items.length) {
      const index = cursor++;
      results[index] = await convert(items[index]);
    }
  };
  await Promise.all(Array.from({ length: Math.min(CONVERSION_LIMIT, items.length) }, worker));
  return results;
}

async function render(ctx, kind, values, position = {}) {
  const { slideWidth, slideHeight } = await pageSize(ctx);
  const payload = {
    ...values, ...position, kind,
    theme: preferredTheme(ctx, values.theme),
    language: values.language && values.language !== 'saved' ? values.language : ctx.settings.get('content.codeLanguage', 'python'),
    slideWidth, slideHeight
  };
  const result = await ctx.services.request('content/render', payload);
  warn(ctx, result);
  if (!result.fragment?.pptx) throw new Error('内容转换没有返回可编辑的 PowerPoint 片段。');
  await ctx.host.insertFragment(result.fragment);
  if (kind === 'code') ctx.settings.set('content.codeLanguage', payload.language);
  ctx.settings.set('content.codeTheme', payload.theme);
  await ctx.refresh();
  return result;
}

function textShapes(context) {
  const selected = new Set(context.selectionIds);
  const result = [];
  function visit(shape) {
    if (selected.has(shape.id) && typeof shape.text === 'string' && shape.text.trim()) result.push(shape);
    for (const child of shape.children || []) visit(child);
  }
  for (const shape of context.shapes) visit(shape);
  return result;
}

export const actions = [
  {
    id: 'content.markdown', label: '插入 Markdown', section,
    description: '插入带标题、嵌套列表、任务列表、链接、表格、代码及可编辑原生公式的富文本。',
    fields: [source('Markdown 内容'), width, size, family, theme],
    run: (ctx, values) => render(ctx, 'markdown', values)
  },
  {
    id: 'content.richText', label: '文本框转富文本', section,
    description: '将选中的 Markdown 文本框转换为同位置、同宽度的富文本；成功插入后删除原文本框。',
    fields: [theme],
    async run(ctx, values) {
      const context = await ctx.host.context({ scope: 'selection' });
      const target = { presentationId: context.presentationId, slideId: context.slideId };
      const selected = textShapes(context);
      if (!selected.length) throw new Error('请先选择一个或多个含有 Markdown 内容的文本框，或将光标定位在文本框中。');
      const resolvedTheme = preferredTheme(ctx, values.theme);
      const prepared = await convertAll(selected, async shape => {
        const result = await ctx.services.request('content/render', {
          kind: 'markdown', source: shape.text,
          left: shape.left, top: shape.top, width: shape.width,
          fontSize: shape.font?.size || 18, fontFamily: shape.font?.name || 'Arial',
          theme: resolvedTheme,
          slideWidth: context.width, slideHeight: context.height
        });
        if (!result.fragment?.pptx) throw new Error('内容转换没有返回可编辑的 PowerPoint 片段，原文本框已保留。');
        return { shape, result };
      });
      for (const { shape, result } of prepared) {
        warn(ctx, result);
        await ctx.host.insertFragment(result.fragment, { target, readback: false });
        await ctx.host.apply([{ op: 'delete', ids: [shape.id] }], { target, readback: false });
      }
      await ctx.refresh();
      return { converted: prepared.length, warnings: prepared.flatMap(entry => entry.result.warnings || []) };
    }
  },
  {
    id: 'content.code', label: '插入高亮代码', section,
    fields: [source('代码'), { name: 'language', label: '语言', type: 'select', default: 'saved', options: [{ value: 'saved', label: '使用上次的语言' }, ...languages] }, theme, width, { ...size, default: 12 }],
    run: (ctx, values) => render(ctx, 'code', values)
  },
  {
    id: 'content.codeTheme', label: '切换选中文本背景', section,
    description: '更新选中文本框的背景和文字颜色，并保存后续代码块的默认主题。',
    fields: [{ ...theme, default: 'light', options: themes }],
    async run(ctx, values) {
      const dark = values.theme === 'dark';
      const context = await ctx.host.context({ scope: 'selection' });
      const shapes = textShapes(context);
      if (!shapes.length) throw new Error('请先选择含有文本的文本框。');
      await ctx.host.apply(shapes.map(shape => ({ op: 'update', id: shape.id, properties: {
        fill: { type: 'solid', color: dark ? '#1E1E1E' : '#FFFFFF' },
        font: { color: dark ? '#FFFFFF' : '#000000' }
      } })));
      ctx.settings.set('content.codeTheme', values.theme);
      await ctx.refresh();
      return { updated: shapes.length };
    }
  },
  {
    id: 'content.latex', label: '插入可编辑 LaTeX 公式', section,
    description: '将 LaTeX 转为原生 Office Math；不支持的表达式保留源码并报告转换错误。',
    fields: [source('LaTeX 公式'), width, { ...size, default: 24 }],
    run: (ctx, values) => render(ctx, 'latex', values)
  },
  {
    id: 'content.latexSvg', label: '插入 LaTeX SVG', section,
    description: '使用内置 MathJax 生成矢量公式，并将 LaTeX 源码保存在图片替代文字中。',
    fields: [source('LaTeX 公式'), width],
    run: (ctx, values) => render(ctx, 'latex-svg', values)
  },
  {
    id: 'content.svg', label: '插入 SVG', section,
    fields: [source('完整 SVG 源码'), width],
    run: (ctx, values) => render(ctx, 'svg', values)
  }
];
