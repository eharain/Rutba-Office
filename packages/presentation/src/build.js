// Making a .pptx from nothing.
//
// The smallest presentation PowerPoint will open is not small: a content-type
// map, package rels, the presentation part and its rels, a theme, a slide
// master with a colour map and text styles, at least one layout, and the slide
// itself. Miss one relationship and PowerPoint offers to repair the file —
// which is the single most damaging thing a new office suite can show a user.
//
// So this builds a complete, valid deck, and `tools/verify-office.ps1` opens
// what it writes in real PowerPoint to prove it.

import { OoxmlPackage } from '@rutba/ooxml/package';
import { escapeXml } from '@rutba/office-formats/xml';
import { pxToEmu, ptToSz } from './units.js';

const XMLNS_P = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

/** 16:9 at 13.333 x 7.5 inches, which is what every deck made since 2013 uses. */
export const SLIDE_SIZE_16_9 = { cx: 12192000, cy: 6858000 };
export const SLIDE_SIZE_4_3 = { cx: 9144000, cy: 6858000 };

const THEME_COLORS = {
  dk1: '000000',
  lt1: 'FFFFFF',
  dk2: '1F3864',
  lt2: 'EEF2F8',
  accent1: '2B5FD9',
  accent2: '0F9D58',
  accent3: 'E08B2B',
  accent4: 'C2408F',
  accent5: '7B5CD6',
  accent6: '3AAFA9',
  hlink: '0563C1',
  folHlink: '954F72',
};

function themeXml(name = 'Rutba') {
  const scheme = Object.entries(THEME_COLORS)
    .map(([k, v]) =>
      k === 'dk1' || k === 'lt1'
        ? `<a:${k}><a:sysClr val="${k === 'dk1' ? 'windowText' : 'window'}" lastClr="${v}"/></a:${k}>`
        : `<a:${k}><a:srgbClr val="${v}"/></a:${k}>`
    )
    .join('');
  return `${DECL}<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="${escapeXml(name)}">` +
    `<a:themeElements><a:clrScheme name="${escapeXml(name)}">${scheme}</a:clrScheme>` +
    `<a:fontScheme name="${escapeXml(name)}">` +
    `<a:majorFont><a:latin typeface="Segoe UI Semibold"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>` +
    `<a:minorFont><a:latin typeface="Segoe UI"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>` +
    `</a:fontScheme>` +
    `<a:fmtScheme name="${escapeXml(name)}">` +
    `<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>` +
    `<a:solidFill><a:schemeClr val="phClr"><a:tint val="60000"/></a:schemeClr></a:solidFill>` +
    `<a:solidFill><a:schemeClr val="phClr"><a:shade val="80000"/></a:schemeClr></a:solidFill></a:fillStyleLst>` +
    `<a:lnStyleLst>${'<a:ln w="9525" cap="flat"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>'.repeat(3)}</a:lnStyleLst>` +
    `<a:effectStyleLst>${'<a:effectStyle><a:effectLst/></a:effectStyle>'.repeat(3)}</a:effectStyleLst>` +
    `<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>` +
    `<a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/></a:schemeClr></a:solidFill>` +
    `<a:solidFill><a:schemeClr val="phClr"><a:shade val="90000"/></a:schemeClr></a:solidFill></a:bgFillStyleLst>` +
    `</a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`;
}

function textStyleLevels(sizes, color = 'tx1') {
  return sizes
    .map((sz, i) =>
      `<a:lvl${i + 1}pPr marL="${i * 342900}" indent="-342900" algn="l"><a:buChar char="•"/>` +
      `<a:defRPr sz="${sz}"><a:solidFill><a:schemeClr val="${color}"/></a:solidFill>` +
      `<a:latin typeface="+mn-lt"/></a:defRPr></a:lvl${i + 1}pPr>`
    )
    .join('');
}

function masterXml() {
  const ph = (type, idx, x, y, cx, cy, anchor, sz, bold) =>
    `<p:sp><p:nvSpPr><p:cNvPr id="${idx + 1}" name="${type} Placeholder ${idx + 1}"/>` +
    `<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="${type}"${idx ? ` idx="${idx}"` : ''}/></p:nvPr></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
    `<p:txBody><a:bodyPr anchor="${anchor}"/><a:lstStyle/>` +
    `<a:p><a:pPr algn="l"/><a:endParaRPr lang="en-US" sz="${sz}"${bold ? ' b="1"' : ''}/></a:p></p:txBody></p:sp>`;

  return `${DECL}<p:sldMaster ${XMLNS_P}><p:cSld><p:bg><p:bgPr>` +
    `<a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>` +
    `<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
    `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>` +
    ph('title', 0, 838200, 731520, 10515600, 1325563, 'b', 4400, true) +
    ph('body', 1, 838200, 2202180, 10515600, 3684270, 't', 2000, false) +
    `</p:spTree></p:cSld>` +
    `<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>` +
    `<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/><p:sldLayoutId id="2147483650" r:id="rId2"/></p:sldLayoutIdLst>` +
    `<p:txStyles>` +
    `<p:titleStyle><a:lvl1pPr algn="l"><a:defRPr sz="4400" b="1"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mj-lt"/></a:defRPr></a:lvl1pPr></p:titleStyle>` +
    `<p:bodyStyle>${textStyleLevels([2000, 1800, 1600, 1400, 1400])}</p:bodyStyle>` +
    `<p:otherStyle><a:lvl1pPr><a:defRPr sz="1800"/></a:lvl1pPr></p:otherStyle>` +
    `</p:txStyles></p:sldMaster>`;
}

function layoutXml(type, name) {
  const title =
    `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>` +
    `<p:nvPr><p:ph type="${type === 'title' ? 'ctrTitle' : 'title'}"/></p:nvPr></p:nvSpPr>` +
    `<p:spPr>${type === 'title' ? '<a:xfrm><a:off x="1524000" y="1935480"/><a:ext cx="9144000" cy="1500187"/></a:xfrm>' : ''}</p:spPr>` +
    `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>`;
  const body =
    `<p:sp><p:nvSpPr><p:cNvPr id="3" name="Content 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>` +
    `<p:nvPr><p:ph type="${type === 'title' ? 'subTitle' : 'body'}" idx="1"/></p:nvPr></p:nvSpPr>` +
    `<p:spPr>${type === 'title' ? '<a:xfrm><a:off x="1524000" y="3602038"/><a:ext cx="9144000" cy="1500187"/></a:xfrm>' : ''}</p:spPr>` +
    `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>`;
  return `${DECL}<p:sldLayout ${XMLNS_P} type="${type === 'title' ? 'title' : 'obj'}" preserve="1"><p:cSld name="${escapeXml(name)}">` +
    `<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
    `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>` +
    title + body +
    `</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;
}

function runXml(run) {
  if (run.text === '\n') return '<a:br/>';
  const bits = ['lang="en-US"'];
  if (run.size) bits.push(`sz="${ptToSz(run.size)}"`);
  if (run.bold) bits.push('b="1"');
  if (run.italic) bits.push('i="1"');
  if (run.underline) bits.push('u="sng"');
  const fill = run.color ? `<a:solidFill><a:srgbClr val="${String(run.color).replace('#', '')}"/></a:solidFill>` : '';
  const font = run.font ? `<a:latin typeface="${escapeXml(run.font)}"/>` : '';
  const rPr = fill || font ? `<a:rPr ${bits.join(' ')}>${fill}${font}</a:rPr>` : `<a:rPr ${bits.join(' ')}/>`;
  return `<a:r>${rPr}<a:t xml:space="preserve">${escapeXml(run.text ?? '')}</a:t></a:r>`;
}

function paragraphXml(p) {
  const attrs = [];
  if (p.level) attrs.push(`lvl="${p.level}"`);
  if (p.align) attrs.push(`algn="${{ left: 'l', center: 'ctr', right: 'r', justify: 'just' }[p.align] || 'l'}"`);
  const bullet = p.bullet === false ? '<a:buNone/>' : '';
  const pPr = attrs.length || bullet ? `<a:pPr ${attrs.join(' ')}>${bullet}</a:pPr>` : '';
  const runs = (p.runs || []).map(runXml).join('');
  return `<a:p>${pPr}${runs || '<a:endParaRPr lang="en-US"/>'}</a:p>`;
}

function placeholderShape(id, name, phType, phIdx, paragraphs, geometry) {
  const xfrm = geometry
    ? `<a:xfrm><a:off x="${pxToEmu(geometry.x)}" y="${pxToEmu(geometry.y)}"/><a:ext cx="${pxToEmu(geometry.w)}" cy="${pxToEmu(geometry.h)}"/></a:xfrm>`
    : '';
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${escapeXml(name)}"/>` +
    `<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>` +
    `<p:nvPr><p:ph type="${phType}"${phIdx != null ? ` idx="${phIdx}"` : ''}/></p:nvPr></p:nvSpPr>` +
    `<p:spPr>${xfrm}</p:spPr>` +
    `<p:txBody><a:bodyPr/><a:lstStyle/>${(paragraphs || []).map(paragraphXml).join('') || '<a:p/>'}</p:txBody></p:sp>`;
}

/**
 * One slide, as the XML a .pptx part holds.
 *
 * Exported because the Deck builds a new slide with exactly the same shape
 * a generated deck does — a second implementation of a placeholder shape is
 * how two code paths quietly diverge on what PowerPoint will open.
 */
export function slideXml(slide) {
  const shapes = [];
  let id = 2;
  if (slide.title != null) {
    const paragraphs = Array.isArray(slide.title)
      ? slide.title
      : [{ runs: [{ text: String(slide.title) }] }];
    shapes.push(placeholderShape(id++, 'Title 1', slide.layout === 'title' ? 'ctrTitle' : 'title', null, paragraphs, slide.titleBox));
  }
  if (slide.body != null) {
    const paragraphs = Array.isArray(slide.body)
      ? slide.body.map((b) => (typeof b === 'string' ? { runs: [{ text: b }] } : b))
      : [{ runs: [{ text: String(slide.body) }] }];
    shapes.push(placeholderShape(id++, 'Content 2', slide.layout === 'title' ? 'subTitle' : 'body', 1, paragraphs, slide.bodyBox));
  }
  for (const box of slide.textBoxes || []) {
    shapes.push(
      `<p:sp><p:nvSpPr><p:cNvPr id="${id++}" name="TextBox"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr><a:xfrm><a:off x="${pxToEmu(box.x)}" y="${pxToEmu(box.y)}"/>` +
      `<a:ext cx="${pxToEmu(box.w)}" cy="${pxToEmu(box.h)}"/></a:xfrm>` +
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>` +
      `<p:txBody><a:bodyPr wrap="square"><a:spAutoFit/></a:bodyPr><a:lstStyle/>` +
      `${(box.paragraphs || []).map(paragraphXml).join('') || '<a:p/>'}</p:txBody></p:sp>`
    );
  }
  return `${DECL}<p:sld ${XMLNS_P}><p:cSld><p:spTree>` +
    `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
    `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>` +
    shapes.join('') +
    `</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
}

function relsXml(rels) {
  return `${DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    rels.map((r) => `<Relationship Id="${r.id}" Type="${r.type}" Target="${r.target}"/>`).join('') +
    `</Relationships>`;
}

const R = {
  officeDocument: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
  coreProps: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
  extProps: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties',
  slide: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
  slideMaster: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster',
  slideLayout: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
  theme: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
};

/**
 * Build a presentation.
 * @param {{ title?: string, author?: string, size?: {cx:number,cy:number}, slides: Array<object> }} spec
 * @returns {Buffer}
 */
export function buildPptx(spec = {}) {
  const size = spec.size || SLIDE_SIZE_16_9;
  const slides = spec.slides?.length ? spec.slides : [{ layout: 'title', title: spec.title || 'Presentation', body: '' }];
  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');

  const parts = [];
  const push = (name, data) => parts.push({ name, data });

  const slideNames = slides.map((_, i) => `ppt/slides/slide${i + 1}.xml`);

  push('[Content_Types].xml',
    `${DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
    `<Default Extension="xml" ContentType="application/xml"/>` +
    `<Default Extension="png" ContentType="image/png"/>` +
    `<Default Extension="jpeg" ContentType="image/jpeg"/>` +
    `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>` +
    `<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>` +
    `<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>` +
    `<Override PartName="/ppt/slideLayouts/slideLayout2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>` +
    `<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>` +
    slideNames.map((n) => `<Override PartName="/${n}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join('') +
    `<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>` +
    `<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>` +
    `</Types>`);

  push('_rels/.rels', relsXml([
    { id: 'rId1', type: R.officeDocument, target: 'ppt/presentation.xml' },
    { id: 'rId2', type: R.coreProps, target: 'docProps/core.xml' },
    { id: 'rId3', type: R.extProps, target: 'docProps/app.xml' },
  ]));

  push('docProps/core.xml',
    `${DECL}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ` +
    `xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" ` +
    `xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">` +
    `<dc:title>${escapeXml(spec.title || 'Presentation')}</dc:title>` +
    `<dc:creator>${escapeXml(spec.author || 'Rutba Office')}</dc:creator>` +
    `<cp:lastModifiedBy>${escapeXml(spec.author || 'Rutba Office')}</cp:lastModifiedBy>` +
    `<dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>` +
    `<dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>` +
    `</cp:coreProperties>`);

  push('docProps/app.xml',
    `${DECL}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" ` +
    `xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">` +
    `<Application>Rutba Office</Application><Slides>${slides.length}</Slides>` +
    `<PresentationFormat>Widescreen</PresentationFormat><Company></Company></Properties>`);

  const presRels = [
    { id: 'rId1', type: R.slideMaster, target: 'slideMasters/slideMaster1.xml' },
    ...slideNames.map((n, i) => ({ id: `rId${i + 2}`, type: R.slide, target: `slides/slide${i + 1}.xml` })),
    { id: `rId${slideNames.length + 2}`, type: R.theme, target: 'theme/theme1.xml' },
  ];

  push('ppt/presentation.xml',
    `${DECL}<p:presentation ${XMLNS_P} saveSubsetFonts="1">` +
    `<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>` +
    `<p:sldIdLst>${slideNames.map((_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 2}"/>`).join('')}</p:sldIdLst>` +
    `<p:sldSz cx="${size.cx}" cy="${size.cy}"/><p:notesSz cx="${size.cy}" cy="${size.cx}"/>` +
    `<p:defaultTextStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr></p:defaultTextStyle>` +
    `</p:presentation>`);
  push('ppt/_rels/presentation.xml.rels', relsXml(presRels));

  push('ppt/slideMasters/slideMaster1.xml', masterXml());
  push('ppt/slideMasters/_rels/slideMaster1.xml.rels', relsXml([
    { id: 'rId1', type: R.slideLayout, target: '../slideLayouts/slideLayout1.xml' },
    { id: 'rId2', type: R.slideLayout, target: '../slideLayouts/slideLayout2.xml' },
    { id: 'rId3', type: R.theme, target: '../theme/theme1.xml' },
  ]));

  push('ppt/slideLayouts/slideLayout1.xml', layoutXml('title', 'Title Slide'));
  push('ppt/slideLayouts/_rels/slideLayout1.xml.rels', relsXml([
    { id: 'rId1', type: R.slideMaster, target: '../slideMasters/slideMaster1.xml' },
  ]));
  push('ppt/slideLayouts/slideLayout2.xml', layoutXml('obj', 'Title and Content'));
  push('ppt/slideLayouts/_rels/slideLayout2.xml.rels', relsXml([
    { id: 'rId1', type: R.slideMaster, target: '../slideMasters/slideMaster1.xml' },
  ]));

  push('ppt/theme/theme1.xml', themeXml(spec.theme || 'Rutba'));

  slides.forEach((slide, i) => {
    push(slideNames[i], slideXml(slide));
    push(`ppt/slides/_rels/slide${i + 1}.xml.rels`, relsXml([
      {
        id: 'rId1',
        type: R.slideLayout,
        target: `../slideLayouts/slideLayout${slide.layout === 'title' ? 1 : 2}.xml`,
      },
    ]));
  });

  return OoxmlPackage.fromParts(parts).write();
}

/** The templates the New menu offers. */
export const TEMPLATES = {
  blank: () => buildPptx({ title: 'Presentation', slides: [{ layout: 'title', title: 'Presentation', body: 'Subtitle' }] }),
  pitch: () =>
    buildPptx({
      title: 'Pitch',
      slides: [
        { layout: 'title', title: 'Company', body: 'One line about what you do' },
        { layout: 'obj', title: 'The problem', body: ['Who has it', 'How often', 'What it costs them'] },
        { layout: 'obj', title: 'The product', body: ['What it does', 'Why it is different', 'What it replaces'] },
        { layout: 'obj', title: 'The market', body: ['Who buys', 'How many', 'What they pay today'] },
        { layout: 'obj', title: 'The ask', body: ['What you need', 'What it buys', 'What happens next'] },
      ],
    }),
};
