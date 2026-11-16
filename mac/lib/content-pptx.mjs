import JSZip from 'jszip';
import PptxGenJS from 'pptxgenjs';
import sharp from 'sharp';
import { xml } from './content-math.mjs';

const ns = {
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  a14: 'http://schemas.microsoft.com/office/drawing/2010/main',
  m: 'http://schemas.openxmlformats.org/officeDocument/2006/math',
  mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  asvg: 'http://schemas.microsoft.com/office/drawing/2016/SVG/main'
};
const emu = points => Math.round(points * 12700);
const solid = color => `<a:solidFill><a:srgbClr val="${xml((color || '#000000').replace('#', ''))}"/></a:solidFill>`;
const geometry = shape => `<a:xfrm><a:off x="${emu(shape.left)}" y="${emu(shape.top)}"/><a:ext cx="${emu(shape.width)}" cy="${emu(shape.height)}"/></a:xfrm>`;
const fill = shape => shape.fill ? solid(shape.fill) : '<a:noFill/>';
const line = shape => shape.line ? `<a:ln w="12700">${solid(shape.line)}<a:prstDash val="solid"/></a:ln>` : '<a:ln><a:noFill/></a:ln>';
const sourceText = shape => shape.source || shape.paragraphs?.map(paragraph => paragraph.runs.map(run => run.text).join('')).join('\n') || '';
const cNvPr = (shape, id) => `<p:cNvPr id="${id}" name="${xml(shape.name || `SlideSCI ${shape.kind} ${id}`)}" descr="${xml(shape.alternativeText ?? sourceText(shape))}"/>`;

// Reuse the empty layout package; each render refreshes its document timestamps.
const templates = new Map();
function baseTemplate(document) {
  const key = `${document.slideWidth}x${document.slideHeight}|${document.kind}`;
  let pending = templates.get(key);
  if (!pending) {
    pending = (async () => {
      const presentation = new PptxGenJS();
      presentation.defineLayout({ name: 'SlideSCI', width: document.slideWidth / 72, height: document.slideHeight / 72 });
      presentation.layout = 'SlideSCI';
      presentation.author = 'SlideSCI';
      presentation.subject = 'Editable SlideSCI content';
      presentation.title = `SlideSCI ${document.kind}`;
      presentation.lang = 'zh-CN';
      presentation.addSlide();
      return presentation.write({ outputType: 'nodebuffer' });
    })();
    templates.set(key, pending);
    if (templates.size > 8) templates.delete(templates.keys().next().value);
    pending.catch(() => { if (templates.get(key) === pending) templates.delete(key); });
  }
  return pending;
}

export async function contentToFragment(document) {
  const zip = await JSZip.loadAsync(await baseTemplate(document));
  const core = zip.file('docProps/core.xml');
  if (core) {
    const timestamp = new Date().toISOString();
    zip.file('docProps/core.xml', (await core.async('string')).replace(
      /(<dcterms:(?:created|modified)\b[^>]*>)[^<]*(<\/dcterms:(?:created|modified)>)/g,
      (_, opening, closing) => opening + timestamp + closing));
  }
  const relationships = [];
  let nextRelationship = 100;
  const hyperlinks = new Map();
  function relationship(type, target, external = false) {
    const id = `rIdSlideSCI${nextRelationship++}`;
    relationships.push(`<Relationship Id="${id}" Type="${ns.r}/${type}" Target="${xml(target)}"${external ? ' TargetMode="External"' : ''}/>`);
    return id;
  }
  function hyperlink(value) {
    if (!hyperlinks.has(value)) hyperlinks.set(value, relationship('hyperlink', value, true));
    return hyperlinks.get(value);
  }
  function fontProperties(font, size, family, link) {
    const attributes = [
      `lang="zh-CN"`, `sz="${Math.round((font.size || size) * (font.superscript || font.subscript ? 0.7 : 1) * 100)}"`,
      ...(font.bold ? ['b="1"'] : []), ...(font.italic ? ['i="1"'] : []),
      ...(font.underline ? ['u="sng"'] : []), ...(font.strikethrough ? ['strike="sngStrike"'] : []),
      ...(font.superscript ? ['baseline="30000"'] : font.subscript ? ['baseline="-25000"'] : [])
    ].join(' ');
    return `<a:rPr ${attributes}>${solid(font.color || '#000000')}${font.highlight ? '<a:highlight><a:srgbClr val="FFF2CC"/></a:highlight>' : ''}<a:latin typeface="${xml(font.name || family)}"/><a:ea typeface="${xml(font.name || family)}"/><a:cs typeface="${xml(font.name || family)}"/>${link ? `<a:hlinkClick r:id="${hyperlink(link)}"/>` : ''}</a:rPr>`;
  }
  function paragraphsXml(paragraphs, shape, fallback = false) {
    return paragraphs.map(paragraph => {
      const size = paragraph.fontSize || shape.fontSize || document.fontSize;
      const family = shape.fontFamily || document.fontFamily;
      const alignment = { left: 'l', center: 'ctr', right: 'r', justify: 'just' }[paragraph.alignment] || 'l';
      const bullet = paragraph.bullet;
      const indentation = paragraph.marginLeft ? ` marL="${emu(paragraph.marginLeft)}" indent="${-emu(paragraph.hanging || 0)}"` : '';
      const bulletXml = bullet ? `<a:buClr>${solid('#000000').replace(/<\/?a:solidFill>/g, '')}</a:buClr><a:buSzPct val="100000"/><a:buFont typeface="Arial"/>` + (bullet.type === 'ordered' ? `<a:buAutoNum type="arabicPeriod" startAt="${Math.max(1, Math.min(32767, bullet.start || 1))}"/>` : `<a:buChar char="${xml(bullet.marker)}"/>`) : '<a:buNone/>';
      const properties = `<a:pPr algn="${alignment}"${indentation}><a:lnSpc><a:spcPct val="110000"/></a:lnSpc><a:spcBef><a:spcPts val="${Math.round((paragraph.spaceBefore || 0) * 100)}"/></a:spcBef><a:spcAft><a:spcPts val="${Math.round((paragraph.spaceAfter ?? 6) * 100)}"/></a:spcAft>${bulletXml}<a:defRPr sz="${Math.round(size * 100)}"><a:latin typeface="${xml(family)}"/><a:ea typeface="${xml(family)}"/></a:defRPr></a:pPr>`;
      const runs = paragraph.runs.map(run => {
        if (run.omml && !fallback) return `<a14:m><m:oMathPara><m:oMathParaPr><m:jc m:val="${paragraph.alignment === 'center' ? 'center' : 'left'}"/></m:oMathParaPr>${run.omml}</m:oMathPara></a14:m>`;
        const properties = fontProperties(run.font || {}, size, family, run.hyperlink);
        return run.text.split('\n').map((part, index) => `${index ? `<a:br>${properties}</a:br>` : ''}${part ? `<a:r>${properties}<a:t xml:space="preserve">${xml(part)}</a:t></a:r>` : ''}`).join('');
      }).join('');
      return `<a:p>${properties}${runs}<a:endParaRPr lang="zh-CN" sz="${Math.round(size * 100)}"/></a:p>`;
    }).join('') || '<a:p><a:endParaRPr/></a:p>';
  }
  function textBody(shape, fallback = false, cell = false) {
    const padding = shape.padding ?? (cell ? 6 : 0);
    return `<${cell ? 'a' : 'p'}:txBody><a:bodyPr wrap="square" lIns="${emu(padding)}" rIns="${emu(padding)}" tIns="${emu(padding)}" bIns="${emu(padding)}" anchor="t"><a:spAutoFit/></a:bodyPr><a:lstStyle/>${paragraphsXml(shape.paragraphs, shape, fallback)}</${cell ? 'a' : 'p'}:txBody>`;
  }
  function textShape(shape, id, fallback = false) {
    return `<p:sp><p:nvSpPr>${cNvPr(shape, id)}<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr>${geometry(shape)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${fill(shape)}${line(shape)}</p:spPr>${textBody(shape, fallback)}</p:sp>`;
  }
  function tableShape(shape, id, fallback = false) {
    const count = shape.rows[0].length;
    const table = `<a:tbl><a:tblPr firstRow="1" bandRow="0"><a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId></a:tblPr><a:tblGrid>${Array.from({ length: count }, () => `<a:gridCol w="${emu(shape.width / count)}"/>`).join('')}</a:tblGrid>${shape.rows.map((row, rowIndex) => `<a:tr h="${emu(shape.rowHeights[rowIndex])}">${row.map(cell => `<a:tc>${textBody({ ...shape, ...cell }, fallback, true)}<a:tcPr marL="76200" marR="76200" marT="76200" marB="76200">${['L', 'R', 'T', 'B'].map(side => `<a:ln${side} w="12700">${solid('#000000')}<a:prstDash val="solid"/></a:ln${side}>`).join('')}${solid(cell.fill || '#FFFFFF')}</a:tcPr></a:tc>`).join('')}</a:tr>`).join('')}</a:tbl>`;
    return `<p:graphicFrame><p:nvGraphicFramePr>${cNvPr(shape, id)}<p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${emu(shape.left)}" y="${emu(shape.top)}"/><a:ext cx="${emu(shape.width)}" cy="${emu(shape.height)}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">${table}</a:graphicData></a:graphic></p:graphicFrame>`;
  }
  const shapesXml = [];
  const shapeIds = [];
  let hasSvg = false;
  let shapeId = 2;
  for (const shape of document.shapes) {
    const id = shapeId++;
    shapeIds.push(String(id));
    if (shape.kind === 'svg') {
      hasSvg = true;
      const svgName = `slidesci-${id}.svg`;
      const pngName = `slidesci-${id}.png`;
      zip.file(`ppt/media/${svgName}`, shape.source);
      const png = await sharp(Buffer.from(shape.source), { density: 144, limitInputPixels: 100000000 }).png().toBuffer();
      zip.file(`ppt/media/${pngName}`, png);
      const imageId = relationship('image', `../media/${pngName}`);
      const svgId = relationship('image', `../media/${svgName}`);
      shapesXml.push(`<p:pic><p:nvPicPr>${cNvPr(shape, id)}<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${imageId}"><a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip r:embed="${svgId}"/></a:ext></a:extLst></a:blip><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>${geometry(shape)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`);
    } else if (shape.kind === 'rect') {
      shapesXml.push(`<p:sp><p:nvSpPr>${cNvPr(shape, id)}<p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>${geometry(shape)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${fill(shape)}${line(shape)}</p:spPr></p:sp>`);
    } else {
      const build = shape.kind === 'table' ? tableShape : textShape;
      const paragraphs = shape.paragraphs || shape.rows.flatMap(row => row.flatMap(cell => cell.paragraphs));
      const math = paragraphs.some(paragraph => paragraph.runs.some(run => run.omml));
      const native = build(shape, id);
      shapesXml.push(math ? `<mc:AlternateContent><mc:Choice Requires="a14">${native}</mc:Choice><mc:Fallback>${build(shape, id, true)}</mc:Fallback></mc:AlternateContent>` : native);
    }
  }
  const namespaceAttributes = Object.entries(ns).map(([prefix, uri]) => `xmlns:${prefix}="${uri}"`).join(' ');
  const slide = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld ${namespaceAttributes} mc:Ignorable="a14 asvg"><p:cSld name="SlideSCI"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${shapesXml.join('')}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
  zip.file('ppt/slides/slide1.xml', slide);
  const relPath = 'ppt/slides/_rels/slide1.xml.rels';
  const rels = await zip.file(relPath).async('string');
  zip.file(relPath, rels.replace('</Relationships>', relationships.join('') + '</Relationships>'));
  if (hasSvg) {
    let types = await zip.file('[Content_Types].xml').async('string');
    for (const [extension, contentType] of [['svg', 'image/svg+xml'], ['png', 'image/png']]) {
      if (!new RegExp(`Extension="${extension}"`).test(types)) types = types.replace('</Types>', `<Default Extension="${extension}" ContentType="${contentType}"/></Types>`);
    }
    zip.file('[Content_Types].xml', types);
  }
  return { pptx: await zip.generateAsync({ type: 'base64', compression: 'DEFLATE' }), shapeIds, name: `SlideSCI ${document.kind}`, bounds: document.bounds };
}
