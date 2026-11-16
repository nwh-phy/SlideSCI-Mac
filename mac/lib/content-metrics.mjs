import { DOMParser } from '@xmldom/xmldom';

/** Conservative intrinsic Office Math dimensions in points, used to keep later blocks below tall formulas. */
export function measureNativeMath(omml, size) {
  const document = new DOMParser().parseFromString(omml, 'text/xml');
  const kids = node => Array.from(node.childNodes || []).filter(child => child.nodeType === 1);
  const zero = () => ({ width: 0, ascent: 0, descent: 0 });
  const row = values => ({ width: values.reduce((sum, value) => sum + value.width, 0), ascent: Math.max(0, ...values.map(value => value.ascent)), descent: Math.max(0, ...values.map(value => value.descent)) });
  function measure(node, scale = 1) {
    if (!node) return zero();
    const tag = node.localName;
    const c = kids(node);
    const s = size * scale;
    const child = name => c.find(value => value.localName === name);
    const value = name => measure(child(name), scale);
    const script = name => measure(child(name), scale * 0.67);
    if (/Pr$/.test(tag) || ['degHide', 'chr', 'mc', 'mcs'].includes(tag)) return zero();
    if (tag === 'r') {
      const text = c.find(value => value.localName === 't')?.textContent || '';
      return { width: Array.from(text).reduce((width, char) => width + s * (/[\u2e80-\uffff]|[∑∏∫√]/u.test(char) ? 0.9 : 0.57), 0), ascent: 0.9 * s, descent: 0.25 * s };
    }
    if (tag === 'f') {
      const numerator = measure(child('num'), scale * 0.9);
      const denominator = measure(child('den'), scale * 0.9);
      return { width: Math.max(numerator.width, denominator.width) + s * 0.4, ascent: numerator.ascent + numerator.descent + s * 0.35, descent: denominator.ascent + denominator.descent + s * 0.2 };
    }
    if (['sSub', 'sSup', 'sSubSup', 'sPre'].includes(tag)) {
      const base = value('e');
      const sub = script('sub');
      const sup = script('sup');
      return { width: base.width + Math.max(sub.width, sup.width), ascent: base.ascent + (sup.ascent + sup.descent) * 0.7, descent: base.descent + (sub.ascent + sub.descent) * 0.7 };
    }
    if (tag === 'rad') {
      const base = value('e');
      const degree = measure(child('deg'), scale * 0.6);
      return { width: base.width + s * 0.7 + degree.width * 0.6, ascent: Math.max(base.ascent + s * 0.3, degree.ascent + s * 0.6), descent: base.descent + s * 0.1 };
    }
    if (['limLow', 'limUpp'].includes(tag)) {
      const base = value('e');
      const limit = measure(child('lim'), scale * 0.75);
      return { width: Math.max(base.width, limit.width), ascent: base.ascent + (tag === 'limUpp' ? limit.ascent + limit.descent + s * 0.1 : 0), descent: base.descent + (tag === 'limLow' ? limit.ascent + limit.descent + s * 0.1 : 0) };
    }
    if (tag === 'nary') {
      const lower = script('sub');
      const upper = script('sup');
      const base = value('e');
      const symbolWidth = Math.max(s, lower.width, upper.width);
      return { width: symbolWidth + base.width, ascent: Math.max(base.ascent, s + upper.ascent + upper.descent), descent: Math.max(base.descent, s * 0.3 + lower.ascent + lower.descent) };
    }
    if (tag === 'm') {
      const rows = c.filter(value => value.localName === 'mr').map(value => kids(value).map(cell => measure(cell, scale)));
      const columns = Math.max(0, ...rows.map(value => value.length));
      const width = Array.from({ length: columns }, (_, index) => Math.max(0, ...rows.map(value => value[index]?.width || 0))).reduce((a, b) => a + b, 0) + Math.max(0, columns - 1) * s * 0.7;
      const height = rows.reduce((sum, cells) => sum + Math.max(s * 1.15, ...cells.map(cell => cell.ascent + cell.descent)) + s * 0.25, 0);
      return { width, ascent: height * 0.55, descent: height * 0.45 };
    }
    if (['acc', 'bar', 'groupChr', 'borderBox', 'd'].includes(tag)) {
      const base = row(c.filter(value => !/Pr$/.test(value.localName)).map(value => measure(value, scale)));
      return { width: base.width + (tag === 'd' ? s : s * 0.15), ascent: base.ascent + s * 0.3, descent: base.descent + s * 0.15 };
    }
    return row(c.filter(value => !/Pr$/.test(value.localName)).map(value => measure(value, scale)));
  }
  return measure(document.documentElement);
}
