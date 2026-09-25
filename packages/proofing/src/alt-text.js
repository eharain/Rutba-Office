// Alternative text, where each format keeps it.
//
// All three formats keep an object's description on its "non-visual drawing
// properties" element — `wp:docPr` in a document, `xdr:cNvPr` on a sheet's
// drawing, `p:cNvPr` on a slide — as a `descr` attribute. "Mark as
// decorative" is an extension Office 2019 added under that same element:
//
//   <a:extLst><a:ext uri="{C183D7F6-B498-43B3-948B-1728B52AA6E4}">
//     <adec:decorative xmlns:adec="http://schemas.microsoft.com/office/drawing/2017/decorative" val="1"/>
//   </a:ext></a:extLst>
//
// These helpers read and rewrite one such element as text, leaving every
// other attribute and child — a hyperlink on the picture, another
// extension — exactly as it was.

export const DECORATIVE_URI = '{C183D7F6-B498-43B3-948B-1728B52AA6E4}';
const ADEC_NS = 'http://schemas.microsoft.com/office/drawing/2017/decorative';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';

const escAttr = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/\r?\n/g, '&#10;');
const unescAttr = (s) => String(s).replace(/&#10;|&#xA;/gi, '\n').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');

/**
 * Every occurrence of an element in `xml`, self-closing or with children:
 * `{ start, end, xml, open }`. The elements this is used for never nest in
 * themselves.
 */
export function findElements(xml, tag) {
  const out = [];
  const re = new RegExp(`<${tag}\\b[^>]*?(/?)>`, 'g');
  let m;
  while ((m = re.exec(xml))) {
    if (m[1] === '/') {
      out.push({ start: m.index, end: m.index + m[0].length, xml: m[0], open: m[0] });
      continue;
    }
    const close = xml.indexOf(`</${tag}>`, m.index + m[0].length);
    if (close < 0) break;
    const end = close + tag.length + 3;
    out.push({ start: m.index, end, xml: xml.slice(m.index, end), open: m[0] });
    re.lastIndex = end;
  }
  return out;
}

const attrOf = (open, name) => {
  const m = new RegExp(`\\s${name}="([^"]*)"`).exec(open);
  return m ? unescAttr(m[1]) : null;
};

/** What an element says about itself: its name, description, title and whether it is decorative. */
export function readAltProps(elementXml) {
  const open = /^<[^>]*>/.exec(String(elementXml))?.[0] || '';
  const decorative = new RegExp(`<a:ext\\b[^>]*uri="${DECORATIVE_URI.replace(/[{}]/g, '\\$&')}"[\\s\\S]*?<adec:decorative\\b[^>]*\\bval="(1|true)"`).test(String(elementXml));
  return {
    id: attrOf(open, 'id'),
    name: attrOf(open, 'name') || '',
    descr: attrOf(open, 'descr') || '',
    title: attrOf(open, 'title') || '',
    hidden: attrOf(open, 'hidden') === '1',
    decorative,
  };
}

/**
 * The element with a new description and decorative flag. A decorative
 * object carries no description (Office clears it); a described one drops
 * the decorative extension. An empty extension list left behind is removed.
 */
export function writeAltProps(elementXml, { descr = '', decorative = false } = {}) {
  const xml = String(elementXml);
  const openMatch = /^<([\w:]+)\b([^>]*?)(\/?)>/.exec(xml);
  if (!openMatch) return xml;
  const [openTag, tag, attrsText, selfClosing] = openMatch;
  let attrs = attrsText.replace(/\sdescr="[^"]*"/, '');
  const text = String(descr || '').trim();
  if (!decorative && text) attrs += ` descr="${escAttr(text)}"`;
  let children = selfClosing ? '' : xml.slice(openTag.length, xml.length - (`</${tag}>`).length);

  // Out with any decorative extension, then back in if it is wanted.
  const extRe = new RegExp(`<a:ext\\b[^>]*uri="${DECORATIVE_URI.replace(/[{}]/g, '\\$&')}"[^>]*>[\\s\\S]*?<\\/a:ext>`, 'g');
  children = children.replace(extRe, '');
  children = children.replace(/<a:extLst\b[^>]*>\s*<\/a:extLst>/g, '').replace(/<a:extLst\b[^>]*\/>/g, '');
  if (decorative) {
    const ext = `<a:ext uri="${DECORATIVE_URI}"><adec:decorative xmlns:adec="${ADEC_NS}" val="1"/></a:ext>`;
    if (/<a:extLst\b[^>]*>/.test(children)) children = children.replace(/<\/a:extLst>/, `${ext}</a:extLst>`);
    else children += `<a:extLst xmlns:a="${A_NS}">${ext}</a:extLst>`;
  }
  return children ? `<${tag}${attrs}>${children}</${tag}>` : `<${tag}${attrs}/>`;
}
