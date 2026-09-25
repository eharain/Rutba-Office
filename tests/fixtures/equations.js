// Equations as Word writes them — Office Math (OMML) with Word's own run
// properties: every m:r in Cambria Math, every structure's control
// properties italic Cambria Math. The quadratic formula below is the one
// Word 365's Insert → Equation → Quadratic Formula writes (the showcase
// document's own, less the highlight that document puts on it).
import { OoxmlPackage } from '@rutba/ooxml';

export const M_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
const RPR = '<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr>';
export const CTRL = '<m:ctrlPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:i/></w:rPr></m:ctrlPr>';

/** One math run, as Word writes it. */
export const mr = (t, mrPr = '') => `<m:r>${mrPr ? `<m:rPr>${mrPr}</m:rPr>` : ''}${RPR}<m:t>${t}</m:t></m:r>`;

export const QUADRATIC_BODY =
  mr('x=') +
  `<m:f><m:fPr>${CTRL}</m:fPr><m:num>${mr('-b±')}` +
  `<m:rad><m:radPr><m:degHide m:val="1"/>${CTRL}</m:radPr><m:deg/><m:e>` +
  `<m:sSup><m:sSupPr>${CTRL}</m:sSupPr><m:e>${mr('b')}</m:e><m:sup>${mr('2')}</m:sup></m:sSup>${mr('-4ac')}` +
  `</m:e></m:rad></m:num><m:den>${mr('2a')}</m:den></m:f>`;

export const display = (body) => `<m:oMathPara><m:oMath>${body}</m:oMath></m:oMathPara>`;
export const inline = (body) => `<m:oMath>${body}</m:oMath>`;
export const QUADRATIC = display(QUADRATIC_BODY);

/**
 * A .docx built by `buildDocx`, with each `@name@` in its paragraphs'
 * words replaced by the equation XML of that name — standing among the
 * runs where Word puts one, the text around it split into runs of its own —
 * and the math namespace declared on the document, as Word declares it.
 */
export function withEquations(bytes, equations) {
  const pkg = OoxmlPackage.read(bytes);
  let xml = pkg.text('word/document.xml');
  xml = xml.replace(/<w:r>(<w:rPr>[\s\S]*?<\/w:rPr>)?<w:t xml:space="preserve">([^<]*@[A-Za-z0-9]+@[^<]*)<\/w:t><\/w:r>/g, (whole, rPr, text) => {
    return text.split(/(@[A-Za-z0-9]+@)/).map((piece) => {
      const name = /^@([A-Za-z0-9]+)@$/.exec(piece)?.[1];
      if (name && equations[name] !== undefined) return equations[name];
      return piece ? `<w:r>${rPr || ''}<w:t xml:space="preserve">${piece}</w:t></w:r>` : '';
    }).join('');
  });
  xml = xml.replace('<w:document ', `<w:document xmlns:m="${M_NS}" `);
  pkg.write_('word/document.xml', Buffer.from(xml, 'utf8'));
  return pkg.write();
}
