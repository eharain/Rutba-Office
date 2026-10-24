// A chart's cached values, brought up to date with its cells.
//
// A chart part keeps the values it was last drawn from beside each series'
// reference (`<c:numRef><c:f>Sales!$B$2:$B$5</c:f><c:numCache>`), and every
// reader draws from that cache until it recalculates. So a chart that is to
// follow its cells must be drawn from the cells, and the file must keep what
// it was drawn from: both are this rewrite of the caches. A cache that
// already agrees with its cells is left byte for byte as it was, so a chart
// nobody's figures touched never shows as changed.

const XML_ESC = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' };
const esc = (s) => String(s).replace(/[&<>"]/g, (c) => XML_ESC[c]);
const unesc = (s) => String(s).replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');

/**
 * A chart reference as cells: `Sheet!$A$1:$A$5` or `'Its sheet'!B2:E2`, one
 * area only. Null for anything else (a name, several areas), which keeps
 * its cache.
 */
export function chartRefCells(text) {
  const m = /^\s*(?:'((?:[^']|'')+)'|([^'!(),]+))!\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?\s*$/i.exec(String(text ?? ''));
  if (!m) return null;
  const col = (letters) => [...letters.toUpperCase()].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  const sheet = m[1] !== undefined ? m[1].replace(/''/g, "'") : m[2];
  const r1 = Number(m[4]) - 1;
  const c1 = col(m[3]);
  const r2 = m[6] ? Number(m[6]) - 1 : r1;
  const c2 = m[5] ? col(m[5]) : c1;
  const cells = [];
  for (let r = Math.min(r1, r2); r <= Math.max(r1, r2); r++) {
    for (let c = Math.min(c1, c2); c <= Math.max(c1, c2); c++) cells.push({ row: r, col: c });
  }
  return { sheet, cells };
}

/** The points a cache holds, by index: the text of each `<c:v>`. */
function cachedPoints(cacheXml) {
  const out = new Map();
  for (const m of cacheXml.matchAll(/<(?:\w+:)?pt\b[^>]*\bidx="(\d+)"[^>]*>\s*<(?:\w+:)?v>([\s\S]*?)<\/(?:\w+:)?v>/g)) {
    out.set(Number(m[1]), unesc(m[2]));
  }
  return out;
}

/**
 * The chart part with each reference's cache made from `read(ref, kind)`:
 * the values the reference covers now, in order (`kind` is 'num' for a
 * numRef, 'str' for a strRef), or null to leave that cache alone. A number
 * cache keeps only numbers, as Excel's does; a text cache every value's text.
 */
export function refreshChartCaches(xml, read) {
  if (!xml) return xml;
  return String(xml).replace(/<(\w+:)?(numRef|strRef)>([\s\S]*?)<\/\1\2>/g, (whole, prefix = '', kind, inner) => {
    const p = prefix ?? '';
    const f = /<(?:\w+:)?f>([\s\S]*?)<\/(?:\w+:)?f>/.exec(inner);
    if (!f) return whole;
    const num = kind === 'numRef';
    const values = read(unesc(f[1]), num ? 'num' : 'str');
    if (!values) return whole;
    const cacheTag = num ? 'numCache' : 'strCache';
    const cacheMatch = new RegExp('<(?:\\w+:)?' + cacheTag + '\\b[^>]*>[\\s\\S]*?</(?:\\w+:)?' + cacheTag + '>|<(?:\\w+:)?' + cacheTag + '\\b[^>]*/>').exec(inner);
    const cacheXml = cacheMatch ? cacheMatch[0] : '';

    const points = new Map();
    values.forEach((v, i) => {
      if (num) {
        if (typeof v === 'number' && Number.isFinite(v)) points.set(i, v);
      } else if (v !== '' && v !== null && v !== undefined) {
        points.set(i, String(v));
      }
    });
    // Unchanged when the count and every point agree.
    const count = /<(?:\w+:)?ptCount\b[^>]*\bval="(\d+)"/.exec(cacheXml);
    const held = cachedPoints(cacheXml);
    const same = cacheXml && (count ? Number(count[1]) === values.length : true) && held.size === points.size
      && [...points].every(([i, v]) => held.has(i) && (num ? Number(held.get(i)) === v : held.get(i) === v));
    if (same) return whole;

    const formatCode = num ? (/<(?:\w+:)?formatCode>[\s\S]*?<\/(?:\w+:)?formatCode>/.exec(cacheXml)?.[0] ?? '<' + p + 'formatCode>General</' + p + 'formatCode>') : '';
    const built = '<' + p + cacheTag + '>' + formatCode + '<' + p + 'ptCount val="' + values.length + '"/>'
      + [...points].map(([i, v]) => '<' + p + 'pt idx="' + i + '"><' + p + 'v>' + esc(v) + '</' + p + 'v></' + p + 'pt>').join('')
      + '</' + p + cacheTag + '>';
    const next = cacheXml ? inner.replace(cacheXml, () => built) : inner.replace(f[0], () => f[0] + built);
    return '<' + p + kind + '>' + next + '</' + p + kind + '>';
  });
}
