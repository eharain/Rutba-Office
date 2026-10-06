/**
 * View → Notes Master and Handout Master: the parts PowerPoint keeps for
 * the printed notes page and the printed handout, as PowerPoint writes a
 * new one — a header and the date along the top, the footer and the page
 * number along the foot, and on the notes page the slide's picture and the
 * notes' text between — laid out for the notes page's own size, portrait or
 * landscape. Nothing here touches a package: the deck puts these where they
 * go (see deck.js).
 */

export const NOTES_MASTER_CT = 'application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml';
export const HANDOUT_MASTER_CT = 'application/vnd.openxmlformats-officedocument.presentationml.handoutMaster+xml';
export const NOTES_MASTER_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster';
export const HANDOUT_MASTER_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/handoutMaster';

const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const CLR_MAP = '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>';

/** The placeholders each master has, in PowerPoint's order, with the names and idx PowerPoint gives them. */
export const MASTER_PLACEHOLDERS = {
  notes: [
    { type: 'hdr', name: 'Header Placeholder', idx: null },
    { type: 'dt', name: 'Date Placeholder', idx: 1 },
    { type: 'sldImg', name: 'Slide Image Placeholder', idx: 2 },
    { type: 'body', name: 'Notes Placeholder', idx: 3 },
    { type: 'ftr', name: 'Footer Placeholder', idx: 4 },
    { type: 'sldNum', name: 'Slide Number Placeholder', idx: 5 },
  ],
  handout: [
    { type: 'hdr', name: 'Header Placeholder', idx: null },
    { type: 'dt', name: 'Date Placeholder', idx: 1 },
    { type: 'ftr', name: 'Footer Placeholder', idx: 2 },
    { type: 'sldNum', name: 'Slide Number Placeholder', idx: 3 },
  ],
};

/**
 * Where a placeholder goes on a page of `page` (`{ cx, cy }` in EMU), as
 * PowerPoint lays a new notes master out: the corners for the header, date,
 * footer and number; the slide's picture, the slide's own shape (`slide`,
 * `{ cx, cy }`), across the top half; the notes under it.
 */
export function placeholderBox(type, page, slide = { cx: 12192000, cy: 6858000 }) {
  const W = page.cx;
  const H = page.cy;
  const cornerW = Math.round(W * 0.4333);
  const cornerH = Math.round(H * 0.0502);
  switch (type) {
    case 'hdr': return { x: 0, y: 0, cx: cornerW, cy: cornerH };
    case 'dt': return { x: W - cornerW, y: 0, cx: cornerW, cy: cornerH };
    case 'ftr': return { x: 0, y: H - cornerH, cx: cornerW, cy: cornerH };
    case 'sldNum': return { x: W - cornerW, y: H - cornerH, cx: cornerW, cy: cornerH };
    case 'sldImg': {
      // The slide's shape, as wide as it can be in the band across the top,
      // standing on the band's foot.
      const left = W * 0.1;
      const top = H * 0.075;
      const bandW = W * 0.8;
      const bandH = H * 0.3875;
      const aspect = (slide.cx || 16) / (slide.cy || 9);
      let cx = bandW;
      let cy = cx / aspect;
      if (cy > bandH) { cy = bandH; cx = cy * aspect; }
      return { x: Math.round(left + (bandW - cx) / 2), y: Math.round(top + bandH - cy), cx: Math.round(cx), cy: Math.round(cy) };
    }
    case 'body': return { x: Math.round(W * 0.1), y: Math.round(H * 0.48125), cx: Math.round(W * 0.8), cy: Math.round(H * 0.39375) };
    default: return null;
  }
}

/** A field's id, a GUID as PowerPoint writes one. */
const guid = () => '{' + 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (c) => {
  const r = Math.floor(Math.random() * 16);
  return (c === 'x' ? r : (r % 4) + 8).toString(16).toUpperCase();
}) + '}';

/** One placeholder's `p:sp`, as PowerPoint writes it on a new master. */
export function placeholderXml(kind, type, id, page, slide) {
  const spec = MASTER_PLACEHOLDERS[kind].find((p) => p.type === type);
  if (!spec) throw new Error(`a ${kind} master has no ${type} placeholder`);
  const box = placeholderBox(type, page, slide);
  const n = MASTER_PLACEHOLDERS[kind].indexOf(spec) + 1;
  const quarter = type === 'hdr' || type === 'body' || type === 'ftr' || type === 'sldNum' ? ' sz="quarter"' : '';
  const ph = `<p:ph type="${type}"${quarter}${spec.idx != null ? ` idx="${spec.idx}"` : ''}/>`;
  const locks = type === 'sldImg' ? '<a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/>' : '<a:spLocks noGrp="1"/>';
  const xfrm = `<a:xfrm><a:off x="${box.x}" y="${box.y}"/><a:ext cx="${box.cx}" cy="${box.cy}"/></a:xfrm>`;
  const look = type === 'sldImg' ? '<a:noFill/><a:ln w="12700"><a:solidFill><a:prstClr val="black"/></a:solidFill></a:ln>' : '';
  const anchor = type === 'ftr' || type === 'sldNum' ? ' anchor="b"' : type === 'sldImg' ? ' anchor="ctr"' : '';
  const bodyPr = `<a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"${anchor}/>`;
  const right = type === 'dt' || type === 'sldNum';
  const lstStyle = type === 'body' || type === 'sldImg' ? '<a:lstStyle/>' : `<a:lstStyle><a:lvl1pPr algn="${right ? 'r' : 'l'}"><a:defRPr sz="1200"/></a:lvl1pPr></a:lstStyle>`;
  let paras = '<a:p><a:endParaRPr lang="en-US"/></a:p>';
  if (type === 'dt') paras = `<a:p><a:fld id="${guid()}" type="datetimeFigureOut"><a:rPr lang="en-US"/><a:t>${new Date().toLocaleDateString('en-US')}</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p>`;
  if (type === 'sldNum') paras = `<a:p><a:fld id="${guid()}" type="slidenum"><a:rPr lang="en-US"/><a:t>‹#›</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p>`;
  if (type === 'body') paras = ['Click to edit Master text styles', 'Second level', 'Third level', 'Fourth level', 'Fifth level']
    .map((t, lvl) => `<a:p><a:pPr lvl="${lvl}"/><a:r><a:rPr lang="en-US"/><a:t>${t}</a:t></a:r></a:p>`).join('');
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${spec.name} ${n}"/><p:cNvSpPr>${locks}</p:cNvSpPr><p:nvPr>${ph}</p:nvPr></p:nvSpPr>`
    + `<p:spPr>${xfrm}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${look}</p:spPr>`
    + `<p:txBody>${bodyPr}${lstStyle}${paras}</p:txBody></p:sp>`;
}

/** The notes' text styles, PowerPoint's own: 12 pt, the theme's body face, each level indented half an inch more. */
const NOTES_STYLE = '<p:notesStyle>' + Array.from({ length: 9 }, (_, i) => `<a:lvl${i + 1}pPr marL="${i * 457200}" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl${i + 1}pPr>`).join('') + '</p:notesStyle>';

/** A whole new master part — `kind` 'notes' or 'handout' — for a page of `page`, slides of `slide`. */
export function masterPartXml(kind, page, slide) {
  const root = kind === 'notes' ? 'p:notesMaster' : 'p:handoutMaster';
  const shapes = MASTER_PLACEHOLDERS[kind].map((p, i) => placeholderXml(kind, p.type, i + 2, page, slide)).join('');
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + `<${root} ${NS}><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree>`
    + '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
    + shapes + `</p:spTree></p:cSld>${CLR_MAP}${kind === 'notes' ? NOTES_STYLE : ''}</${root}>`;
}

/** Which placeholder types a master part has now. */
export function placeholderTypesIn(xml) {
  return new Set([...String(xml).matchAll(/<p:ph\b[^>]*\btype="([A-Za-z]+)"/g)].map((m) => m[1]));
}
