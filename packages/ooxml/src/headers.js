/**
 * Headers and footers.
 *
 * They were preserved from the first commit and never painted, which was fine
 * while a document was one endless page — there was no top or bottom of a sheet
 * to put them on. Pagination gives them somewhere to go, and a letterhead that
 * does not appear on the page is the sort of omission a customer notices
 * immediately, because it is their company name missing.
 *
 * Three reference types, and the distinction matters on the first page of a
 * letter: `first` replaces the header on page one (a letterhead has the logo
 * there and nothing on later pages), `even` alternates for bound printing, and
 * `default` covers the rest. Ignoring the distinction puts the letterhead on
 * every page, which is wrong in a way that looks deliberate.
 */
import { OoxmlPackage, attrs } from './package.js';
import { textOf, parseRuns } from './runs.js';

const REFERENCE = /<w:(header|footer)Reference\b([^>]*)\/>/g;

const fieldName = (instr) => {
  const m = /^\s*([A-Z]+)/.exec(String(instr ?? '').trim());
  return m ? m[1] : null;
};

/**
 * Paragraphs of a header/footer part, with field runs marked.
 *
 * A field becomes a run with a `field` name rather than plain text, so the
 * renderer can substitute the page number per page without re-parsing the part
 * once for every sheet.
 */
function parseBand(xml) {
  const out = [];
  for (const m of String(xml).matchAll(/<w:p\b[^>]*?(?:\/>|>([\s\S]*?)<\/w:p>)/g)) {
    const inner = m[1] ?? '';
    const pPr = /<w:pPr\b[^>]*>[\s\S]*?<\/w:pPr>/.exec(inner);
    const jc = pPr ? /<w:jc\b[^>]*w:val="([^"]*)"/.exec(pPr[0]) : null;

    // Walk the paragraph in document order so a field keeps its place in the
    // sentence: "Page 2 of 7" is three runs and one of them is computed.
    const parts = [];
    const re = /<w:fldSimple\b([^>]*)>([\s\S]*?)<\/w:fldSimple>|<w:r\b(?![a-zA-Z])[^>]*>([\s\S]*?)<\/w:r>/g;
    let piece;
    while ((piece = re.exec(inner))) {
      if (piece[1] !== undefined) {
        const name = fieldName(attrs(piece[1])['w:instr']);
        parts.push({
          field: name,
          // The cached result, shown when we cannot compute the field.
          text: textOf(piece[2]),
          bold: false,
          italic: false,
          underline: false,
        });
        continue;
      }
      const runXml = piece[0];
      if (!/<w:t\b/.test(runXml)) continue;
      const [run] = parseRuns(runXml);
      if (run) parts.push({ field: null, text: run.text, bold: run.bold, italic: run.italic, underline: run.underline });
    }

    out.push({
      align: jc ? jc[1] : null,
      runs: parts,
      text: parts.map((p) => p.text).join(''),
    });
  }
  return out;
}

/**
 * Every header and footer the section references, keyed by type.
 *
 * @returns {{headers: Record<string, object[]>, footers: Record<string, object[]>}}
 */
export function readHeadersAndFooters(pkg, mainPart, sectPrXml) {
  const result = { headers: {}, footers: {} };
  if (!sectPrXml) return result;

  const rels = new Map(pkg.rels(mainPart).map((r) => [r.Id, r.Target]));
  for (const m of String(sectPrXml).matchAll(REFERENCE)) {
    const a = attrs(m[2]);
    const id = a['r:id'] ?? a['relationships:id'] ?? a.id;
    const type = a['w:type'] ?? 'default';
    const target = id ? rels.get(id) : null;
    if (!target) continue;
    const part = OoxmlPackage.resolveTarget(mainPart, target);
    if (!pkg.has(part)) continue;
    const bucket = m[1] === 'header' ? result.headers : result.footers;
    bucket[type] = { part, paragraphs: parseBand(pkg.text(part)) };
  }
  return result;
}

/*
 * Choosing WHICH band a page gets, and resolving its page-number field, used to
 * live here and now lives in `@rutba/doc-view/bands`. Neither knows what OOXML
 * is — they work on the plain data `parseBand` produces — and having them here
 * put a format dependency in the middle of a format-neutral editor.
 */

export { parseBand };
