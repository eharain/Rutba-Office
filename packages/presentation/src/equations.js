// Equations on a slide — Insert → Equation.
//
// PowerPoint writes an equation as a text box whose paragraph holds Office
// Math (`a14:m`, the OMML Word writes) inside an `mc:AlternateContent`: the
// choice for readers that know `a14`, and a fallback — the same box filled
// with a picture of the equation — for those that do not. The OMML comes
// from Word's own linear-format builder; its run properties are Word's
// (`w:rPr`), which a slide does not know, so they are written as the
// DrawingML ones PowerPoint puts there, in Cambria Math.

import { escapeXml } from '@rutba/office-formats/xml';

export const MATH_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
const MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const A14_NS = 'http://schemas.microsoft.com/office/drawing/2010/main';
const CAMBRIA = '<a:latin typeface="Cambria Math" panose="02040503050406030204" pitchFamily="18" charset="0"/>';

/** OMML as a slide holds it: DrawingML run properties, and the math namespace declared on it. */
export function ommlForSlide(omml) {
  let xml = String(omml || '')
    .replace(/<w:rPr>([\s\S]*?)<\/w:rPr>/g, (m, inner) => `<a:rPr lang="en-US"${/<w:i\s*\/>|<w:i\b[^>]*w:val="(1|true|on)"/.test(inner) ? ' i="1"' : ''}>${CAMBRIA}</a:rPr>`)
    .replace(/<w:rPr\s*\/>/g, `<a:rPr lang="en-US">${CAMBRIA}</a:rPr>`);
  xml = xml.replace(/^\s*<(m:oMathPara|m:oMath)\b(?![^>]*\bxmlns:m=)/, `<$1 xmlns:m="${MATH_NS}"`);
  return xml;
}

/**
 * The pair PowerPoint writes for one equation: the text box with the math
 * (`size` its type size, in points) and the same box drawn as its picture
 * (`embed` the picture's relationship id, or none — then the fallback holds
 * the equation's linear form as words, which a reader without math shows).
 */
export function equationShapeXml({ id, name = null, x, y, cx, cy, omml, size = 28, embed = null, linear = '' }) {
  const label = escapeXml(name || `Equation ${id - 1}`);
  const xfrm = `<a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>`;
  const sz = Math.round(size * 100);
  const choice = `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${label}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr>${xfrm}<a:noFill/></p:spPr>` +
    `<p:txBody><a:bodyPr wrap="none" rtlCol="0" anchor="ctr"><a:spAutoFit/></a:bodyPr><a:lstStyle/>` +
    `<a:p><a:pPr algn="ctr"/><a14:m>${ommlForSlide(omml)}</a14:m><a:endParaRPr lang="en-US" sz="${sz}" dirty="0"/></a:p></p:txBody></p:sp>`;
  const locks = '<a:spLocks noRot="1" noChangeAspect="1" noMove="1" noResize="1" noEditPoints="1" noAdjustHandles="1" noChangeArrowheads="1" noChangeShapeType="1" noTextEdit="1"/>';
  const fallback = embed
    ? `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${label}"/><p:cNvSpPr txBox="1">${locks}</p:cNvSpPr><p:nvPr/></p:nvSpPr>` +
      `<p:spPr>${xfrm}<a:blipFill><a:blip r:embed="${embed}"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></p:spPr>` +
      `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"><a:noFill/></a:rPr><a:t> </a:t></a:r></a:p></p:txBody></p:sp>`
    : `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${label}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr>${xfrm}<a:noFill/></p:spPr>` +
      `<p:txBody><a:bodyPr wrap="none" rtlCol="0" anchor="ctr"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="${sz}" dirty="0"/><a:t>${escapeXml(linear)}</a:t></a:r></a:p></p:txBody></p:sp>`;
  return `<mc:AlternateContent xmlns:mc="${MC_NS}"><mc:Choice xmlns:a14="${A14_NS}" Requires="a14">${choice}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent>`;
}
