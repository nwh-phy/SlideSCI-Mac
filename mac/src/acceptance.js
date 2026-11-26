import * as editing from './features/editing.js';
import * as images from './features/images.js';
import * as science from './features/science.js';
import * as content from './features/content.js';
import * as library from './features/library.js';
import * as exported from './features/export.js';
const features = [editing, images, science, content, library, exported];
import { findNative } from './wps-host.js';
export function startAcceptance(ctx) {
  let stopped = false;
  const title = () => ctx.host.kind === 'wps' ? String(window.Application.ActivePresentation.Name) : decodeURIComponent(Office.context.document.url.split('/').at(-1));
  async function execute(job) {
    switch (job.op) {
      case 'manifest': return { actions: features.flatMap(feature => feature.actions || []).map(({ id, fields }) => ({ id, fields })), panels: features.flatMap(feature => feature.panels || []).map(({ id }) => id) };
      case 'summary': return ctx.host.summary(job.options || {});
      case 'context': return ctx.host.context(job.options || {});
      case 'apply': return ctx.host.apply(job.commands);
      case 'sources': return images.insertSourceFiles(ctx, job.files, job.options || {});
      case 'images': return ctx.host.images(job.ids, job.options || {});
      case 'fragment': return ctx.host.fragment(job.ids);
      case 'insertFragment': return ctx.host.insertFragment(job.fragment, job.options || {});
      case 'raster': return ctx.host.rasterizeImage(job.shapeId, job.options || {});
      case 'export': return ctx.host.export(job.options);
      case 'scale': {
        const context = await ctx.host.context({ scope: 'slide' });
        const shape = context.shapes.find(shape => shape.id === job.shapeId);
        if (!shape) throw new Error('Test scale source no longer exists');
        return science.applyScaleBar(ctx, shape, job.fov, job.settings);
      }
      case 'zoom': {
        const context = await ctx.host.context({ scope: 'slide' });
        const shape = context.shapes.find(shape => shape.id === job.shapeId);
        if (!shape) throw new Error('Test zoom source no longer exists');
        return science.applyZoomEntries(ctx, [shape], job.entries);
      }
      case 'pseudoColor': {
        const context = await ctx.host.context({ scope: 'slide' });
        const shapes = context.shapes.filter(shape => job.ids.includes(shape.id));
        return images.applyPseudoColors(ctx, shapes, job.colors, job.keepOriginal);
      }
      case 'selectText': {
        if (ctx.host.kind === 'wps') {
          const shape = findNative(window.Application.ActivePresentation.Slides.Item((job.slideIndex || 0) + 1), job.shapeId);
          shape.TextFrame.TextRange.Characters(job.start + 1, job.length).Select();
        } else await PowerPoint.run(async context => {
          const shape = context.presentation.slides.getItemAt(job.slideIndex || 0).shapes.getItem(job.shapeId);
          shape.textFrame.textRange.getSubstring(job.start, job.length).setSelected();
          await context.sync();
        });
        return ctx.host.context();
      }
      case 'action': {
        const action = features.flatMap(feature => feature.actions || []).find(action => action.id === job.action);
        if (!action) throw new Error('Unknown real feature action: ' + job.action);
        const defaults = Object.fromEntries((action.fields || []).filter(field => field.default !== undefined).map(field => [field.name, field.default]));
        return action.run(ctx, { ...defaults, ...job.values });
      }
      case 'inspect': {
        const app = window.Application, pres = app.ActivePresentation;
        return { name: String(pres.Name), slides: Number(pres.Slides.Count), SaveCopyAs: typeof pres.SaveCopyAs, SaveAs: typeof pres.SaveAs,
          window: typeof app.ActiveWindow, contextUrl: location.href };
      }
      case 'stop': stopped = true; return { stopped: true };
      default: throw new Error('Unknown acceptance operation');
    }
  }
  async function poll() {
    if (stopped) return;
    try {
      const next = await ctx.services.request('acceptance/next', { host: ctx.host.kind, title: title() });
      if (next.job) {
        try {
          const start = performance.now();
          const value = await execute(next.job);
          const result = next.job.measure ? { value, durationMs: performance.now() - start } : value;
          await ctx.services.request('acceptance/result', { id: next.job.id, ok: true, result });
        } catch (error) {
          await ctx.services.request('acceptance/result', { id: next.job.id, ok: false, error: error.message });
        }
      }
    } catch (error) { console.error('Native acceptance: ' + error.message); }
    if (!stopped) setTimeout(poll, 300);
  }
  poll();
  return () => { stopped = true; };
}
