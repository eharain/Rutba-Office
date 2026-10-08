// Word: a web page, opened as a document.
//
// A page with a heading, a run in its looks, a link, a bulleted list with
// a level inside it, a lettered list starting at c, a table with a cell
// spanning two rows and one spanning two columns, and an embedded picture
// opens as its pages — the lists labelled, the table a table, the picture
// drawn — where it came in as its words alone. Run alone with
// RUTBA_VERIFY_ONLY=html.
import fs from 'node:fs';
import path from 'node:path';

const PNG = 'iVBORw0KGgoAAAANSUhEUgAAABAAAAAICAIAAAB/FOjAAAAAE0lEQVR4nGOQizpBEmIY1UALDQAzrqABidudowAAAABJRU5ErkJggg==';
const PAGE = `<!DOCTYPE html><html><head><title>A page</title><style>body { font-family: serif }</style></head><body>
<h1>A page, opened as a document</h1>
<p>Plain, <b>bold</b> and a <a href="https://example.org/">link</a>.</p>
<ul><li>First bullet<ul><li>Inside it</ul><li>Second bullet</ul>
<ol type="a" start="3"><li>Third<li>Fourth</ol>
<table border="1"><tr><th>Region<th>Q1<th>Q2<tr><td rowspan="2">North and South<td colspan="2">both quarters<tr><td>140<td>150</table>
<p><img src="data:image/png;base64,${PNG}" width="64" height="32" alt="Dot"></p>
<p>After the picture.</p>
</body></html>`;

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, doc, sessionFor
 * @param {{ dir: string }} args where the file is written
 */
export async function verifyWordHtml(h, { dir }) {
  const { open, check, until, errorsIn, sessionFor } = h;
  const file = path.join(dir, 'page.html');
  try {
    fs.writeFileSync(file, PAGE);
    const win = await open('word', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`document.querySelectorAll('.wd-page table.wd-table').length > 0`), 'the table to be drawn', 8000).catch(() => {});
    if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'word-html.png'), (await win.webContents.capturePage()).toPNG());
    const shown = await js(`(() => {
      const page = document.querySelector('.wd-page');
      const table = page?.querySelector('table.wd-table');
      return {
        text: page ? page.innerText : '',
        heading: page ? [...page.querySelectorAll('h1, .wd-block')].some((n) => /A page, opened as a document/.test(n.textContent) && parseFloat(getComputedStyle(n).fontSize) > 18) : false,
        rows: table ? [...table.tBodies[0].rows].map((r) => [...r.cells].map((c) => c.innerText.trim() + (c.rowSpan > 1 ? '^' + c.rowSpan : '') + (c.colSpan > 1 ? '<' + c.colSpan : '')).join('|')) : [],
        pictures: page ? [...page.querySelectorAll('img')].filter((i) => i.naturalWidth > 0).length : 0,
        links: page ? page.querySelectorAll('a[href="https://example.org/"], [data-href="https://example.org/"]').length : 0,
      };
    })()`);
    const session = sessionFor('doc');
    check('word: a web page opens as a document, its heading a heading',
      session.converted?.from === 'html' && shown.heading && /After the picture\./.test(shown.text),
      `from ${session.converted?.from}; heading ${shown.heading}; ${shown.text.replace(/\s+/g, ' ').slice(0, 140)}`);
    const labels = /•\s*First bullet/.test(shown.text) && /◦\s*Inside it/.test(shown.text) && /c\.\s*Third/.test(shown.text) && /d\.\s*Fourth/.test(shown.text);
    check('word: the page\'s lists are labelled, a level inside and a list starting at c', labels, shown.text.replace(/\s+/g, ' ').slice(0, 200));
    check('word: the page\'s table is a table with its cell down two rows and its cell across two columns',
      // Drawn over both rows, or — where the window draws a merge's continuation as its own place — an empty cell under it.
      shown.rows[0] === 'Region|Q1|Q2' && /^North and South(\^2)?\|both quarters<2$/.test(shown.rows[1] || '') && /^\|?140\|150$/.test(shown.rows[2] || ''),
      shown.rows.join(' / ') || 'no table');
    check('word: the page\'s embedded picture is drawn', shown.pictures === 1, `${shown.pictures} picture(s)`);
    const complaints = await errorsIn(win);
    check('word: opening a web page reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('word: the web page checks ran', false, err.message);
  }
}
