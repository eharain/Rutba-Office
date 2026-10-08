// Building blocks: words kept to be put in again — Insert → Quick Parts.
//
// Word keeps a building block as paragraphs in a template's glossary, and
// puts them into whichever document asks. Here a block is kept the same way,
// as the paragraphs' own XML, outside any one document — so it has to be
// able to travel: on the way out it loses what only made sense in the
// document it came from (a picture's relationship, a bookmark's name, a
// comment, a note, a tracked change still open), and on the way in it
// loses what the document it lands in cannot resolve (a style or a list it
// has no definition for). What is left is the words and their look.
//
// Pure: XML in, XML out.

// Text as XML holds it, leaving out what XML 1.0 has no place for, as the
// package's own escaper does (this file is drawn in the window too, so it
// keeps its own rather than import the package's Node side).
const esc = (s) => String(s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g, '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/**
 * A paragraph as a building block keeps it. Tracked changes are taken as
 * they would be accepted; a link to the web becomes a HYPERLINK field, which
 * needs no relationship (`linkTarget(rId)` gives the address); pictures,
 * notes, comments and bookmarks are left out; a section break inside it
 * goes, as does Word's revision bookkeeping.
 */
export function portableParagraph(xml, { linkTarget = () => null } = {}) {
  let s = String(xml || '');
  if (!/^<w:p\b/.test(s)) return null;
  s = s
    // Accepted: deletions gone, insertions and moves kept as plain words.
    .replace(/<w:(del|moveFrom)\b(?![^>]*\/>)[^>]*>[\s\S]*?<\/w:\1>/g, '')
    .replace(/<w:(ins|del|moveFrom|moveTo)\b[^>]*\/>/g, '')
    .replace(/<\/?w:(ins|moveTo)\b[^>]*>/g, '')
    .replace(/<w:(pPrChange|rPrChange|sectPr)\b[\s\S]*?<\/w:\1>/g, '')
    // Places in this document, by name or by number.
    .replace(/<w:(bookmarkStart|bookmarkEnd|commentRangeStart|commentRangeEnd|proofErr|permStart|permEnd|moveFromRangeStart|moveFromRangeEnd|moveToRangeStart|moveToRangeEnd)\b[^>]*\/>/g, '')
    // Runs holding what lives in another part of this document.
    .replace(/<w:r\b[^>]*>(?:(?!<\/w:r>)[\s\S])*?<w:(drawing|pict|object|footnoteReference|endnoteReference|commentReference)\b[\s\S]*?<\/w:r>/g, '')
    .replace(/\s+w:rsid\w*="[^"]*"/g, '')
    .replace(/\s+w14:(paraId|textId)="[^"]*"/g, '');
  // A link: to the web as a field; to a place in this document, its words.
  s = s.replace(/<w:hyperlink\b([^>]*)>([\s\S]*?)<\/w:hyperlink>/g, (_, attrs, inner) => {
    const rId = /\br:id="([^"]*)"/.exec(attrs)?.[1];
    const target = rId ? linkTarget(rId) : null;
    return target ? `<w:fldSimple w:instr=" HYPERLINK &quot;${esc(target).replace(/&quot;/g, '')}&quot; ">${inner}</w:fldSimple>` : inner;
  });
  return s;
}

/**
 * A kept paragraph fitted to the document it is going into: a paragraph or
 * character style the document has no definition for is taken off (the
 * words keep their direct formatting), as is a list the document has no
 * numbering for.
 */
export function fitParagraph(xml, { styles = null, charStyles = null, numIds = null } = {}) {
  let s = String(xml || '');
  if (styles) s = s.replace(/<w:pStyle\b[^>]*\bw:val="([^"]*)"[^>]*\/>/g, (m, id) => (styles.has(id) ? m : ''));
  if (charStyles) s = s.replace(/<w:rStyle\b[^>]*\bw:val="([^"]*)"[^>]*\/>/g, (m, id) => (charStyles.has(id) ? m : ''));
  if (numIds) {
    s = s.replace(/<w:numPr\b[^>]*>[\s\S]*?<\/w:numPr>/g, (m) => {
      const id = /<w:numId\b[^>]*\bw:val="([^"]*)"/.exec(m)?.[1];
      return id && numIds.has(id) ? m : '';
    });
  }
  // A pPr or rPr emptied by the above is no use to anyone.
  return s.replace(/<w:pPr><\/w:pPr>/g, '').replace(/<w:rPr><\/w:rPr>/g, '');
}

/** The words of kept paragraphs, for a gallery's preview. */
export function blockText(paragraphs, max = 160) {
  const text = (paragraphs || [])
    .map((p) => [...String(p).matchAll(/<w:t\b[^>]*>([^<]*)<\/w:t>|<w:tab\/>/g)].map((m) => (m[1] !== undefined ? m[1] : '\t')).join(''))
    .join('\n')
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
  return text.length > max ? text.slice(0, max - 1) + '…' : text;
}
