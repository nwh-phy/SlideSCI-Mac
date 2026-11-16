import { buildContentDocument } from './content-document.mjs';
import { contentToFragment } from './content-pptx.mjs';

export const contentHandlers = {
  async render(input) {
    if (!input || typeof input !== 'object' || Array.isArray(input)) throw new Error('内容请求必须为对象。');
    if (typeof input.source !== 'string') throw new Error('内容 source 必须为文本。');
    if (/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/.test(input.source)) throw new Error('内容包含 Office XML 不允许的控制字符。');
    const document = buildContentDocument(input);
    const fragment = await contentToFragment(document);
    return {
      fragment,
      warnings: document.warnings,
      bounds: document.bounds,
      shapeCount: document.shapes.length,
      editable: !['latex-svg', 'svg'].includes(document.kind),
      source: input.source
    };
  },
  async languages() {
    const { default: hljs } = await import('highlight.js');
    return { languages: ['text', ...hljs.listLanguages()].sort().map(value => ({ value, label: value === 'text' ? '纯文本' : value })) };
  }
};
