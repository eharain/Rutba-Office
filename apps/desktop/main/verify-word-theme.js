// Documents: Design → Colours, Fonts and Themes, pressed on the ribbon — a
// heading Word colours from accent 1 and sets in the heading face is drawn
// in the new colours and face after each, and the saved file carries the
// theme.

import fs from 'node:fs';
import path from 'node:path';
import { OoxmlPackage } from '@rutba/ooxml';
import { buildDocx } from '@rutba/ooxml/build';

/** Heading 1 as Word writes it: accent 1, shaded, in the heading face. */
function wordLike() {
  const pkg = OoxmlPackage.read(buildDocx({ styles: true, paragraphs: [{ text: 'A heading', style: 'Heading1' }, { text: 'Body words' }] }));
  const styles = pkg.text('word/styles.xml').replace(/(<w:style w:type="paragraph" w:styleId="Heading1">[\s\S]*?<w:rPr>)([\s\S]*?)(<\/w:rPr>)/, (m, open, props, close) =>
    open + '<w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/>' + props.replace(/<w:color\b[^>]*\/>/, '') + '<w:color w:val="2F5496" w:themeColor="accent1" w:themeShade="BF"/>' + close);
  pkg.write_('word/styles.xml', Buffer.from(styles, 'utf8'));
  return pkg.write();
}

/**
 * @param {object} h the harness: open, check, until, wait
 */
export async function verifyWordTheme(h, { dir }) {
  const { open, check, until, wait, doc } = h;
  const file = path.join(dir, 'theme.docx');
  try {
    fs.writeFileSync(file, wordLike());
    const win = await open('word', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const heading = () => js(`(() => { const p = [...document.querySelectorAll('.wd-page .wd-block')].find((n) => n.textContent.includes('A heading')); if (!p) return null; const s = getComputedStyle(p.querySelector('span') || p); return { colour: s.color, font: s.fontFamily }; })()`);
    const pick = async (label, item) => {
      await js(`(async () => {
        [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Design')?.click();
        await new Promise((r) => setTimeout(r, 200));
        const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === ${JSON.stringify(label)});
        b?.dispatchEvent(new MouseEvent('click', { bubbles: true, clientX: 300, clientY: 110 }));
        return 1;
      })()`);
      await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((n) => n.textContent.trim().startsWith(${JSON.stringify(item)})))`), `the ${item} item`, 3000).catch(() => {});
      return js(`(() => { const b = [...document.querySelectorAll('.rw-menu button')].find((n) => n.textContent.trim().startsWith(${JSON.stringify(item)})); if (!b) return 'no item'; b.click(); return 'picked'; })()`);
    };

    await until(async () => Boolean(await heading()), 'the heading', 8000);
    const before = await heading();
    const colours = await pick('Colours', 'Berry');
    await until(async () => (await heading())?.colour !== before.colour, 'the heading recoloured', 4000).catch(() => {});
    const recoloured = await heading();
    const fonts = await pick('Fonts', 'Classic');
    await until(async () => /Georgia/.test((await heading())?.font || ''), 'the heading in Georgia', 4000).catch(() => {});
    const refaced = await heading();
    check('documents: Design → Colours and Fonts redraw a heading Word colours from the theme and sets in its heading face',
      colours === 'picked' && fonts === 'picked' && recoloured.colour !== before.colour && /Georgia/.test(refaced.font),
      JSON.stringify({ before, recoloured, refaced }));

    const themed = await pick('Themes', 'Harbour');
    await wait(300);
    await js(`(() => { [...document.querySelectorAll('.rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith('Save'))?.click(); return 1; })()`);
    const saved = await until(() => {
      try { return /<a:theme\b[^>]*name="Harbour"/.test(OoxmlPackage.read(fs.readFileSync(file)).text('word/theme/theme1.xml')); } catch { return false; }
    }, 'the saved theme', 8000).then(() => true).catch(() => false);
    check('documents: Design → Themes puts a whole theme on the document, and the saved file carries it', themed === 'picked' && saved, JSON.stringify({ themed, saved }));

    // Set as Default (the question answered Yes), then a new blank document.
    globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'message', answer: { response: 0, checked: false } });
    const pressed = await js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Set as Default'); if (!b) return 'no button'; b.click(); return 'clicked'; })()`);
    const kept = await until(() => js(`window.rutbaOffice.store.get({ key: 'documents.defaultDesign', fallback: null }).then((d) => Boolean(d && /name="Harbour"/.test(d.theme || '')))`), 'the default kept', 4000).then(() => true).catch(() => false);
    const known = new Set(doc.sessions().map((s) => s.id));
    await js(`window.rutbaOffice.win.create({ app: 'word' }), 1`);
    const fresh = await until(() => doc.sessions().some((s) => s.kind === 'doc' && !known.has(s.id)), 'the new document', 8000).then(() => doc.sessions().find((s) => s.kind === 'doc' && !known.has(s.id))).catch(() => null);
    const startsIn = fresh ? doc.model({ id: fresh.id }).design?.name : null;
    check('documents: Design → Set as Default, asked and answered, makes new blank documents start in this theme',
      pressed === 'clicked' && kept && startsIn === 'Harbour' && fresh?.dirty === false, JSON.stringify({ pressed, kept, startsIn, dirty: fresh?.dirty }));
  } catch (err) {
    check('documents: the Design theme checks ran', false, err.message);
  }
}
