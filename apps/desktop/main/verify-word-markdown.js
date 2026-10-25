// Word: a Markdown file opened and saved (2.18.0, A3.1). A README opened
// and saved with no edit comes back as it was written: the raw HTML above a
// code block keeps its blank line, the table its padding, the reference
// links and the definition list their form. A task's box is a check box on
// the page, its "- " not shown; a click ticks it, and the save writes [x].
// Run alone with RUTBA_VERIFY_ONLY=wordmarkdown.
import fs from 'node:fs';
import path from 'node:path';

const README = `# Larkspur stock

<details>
<summary>How the list is kept</summary>
</details>

\`\`\`js
const shelf = 'A3';
\`\`\`

| Item  | Qty |
|-------|----:|
| Tea   |   2 |

See [the supplier list][suppliers].

Reorder level
: The count below which an item is ordered again.

- [ ] Order milk
- [x] Pay the rent

[suppliers]: https://example.com/suppliers
`;

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn
 * @param {{ dir: string }} args where the fixture is written
 */
export async function verifyWordMarkdown({ open, check, until, wait, press, errorsIn }, { dir }) {
  const file = path.join(dir, 'markdown-kept.md');
  fs.writeFileSync(file, README);
  let win = null;
  try {
    win = await open('word', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    await until(() => js(`document.querySelectorAll('.wd-page .wd-cc[data-kind="checkBox"]').length === 2`), 'the task boxes', 10000).catch(() => false);
    const page = await js(`(() => {
      const boxes = [...document.querySelectorAll('.wd-page .wd-cc[data-kind="checkBox"]')];
      const line = boxes[0]?.closest('[data-block]');
      return { boxes: boxes.map((b) => b.getAttribute('aria-checked')), line: line?.textContent || '' };
    })()`);
    check('word: a Markdown task item shows a check box, ticked as the file says, with no "- " before it',
      JSON.stringify(page.boxes) === '["false","true"]' && /Order milk/.test(page.line) && !/^\s*-\s/.test(page.line), JSON.stringify(page));

    // Saved with no edit: the file as it was written.
    await js(`document.querySelector('.wd-page')?.focus(), 1`);
    await press(wc, 's', { modifiers: ['control'] });
    await wait(800);
    const same = await until(() => fs.readFileSync(file, 'utf8') === README, 'the save', 6000).then(() => true).catch(() => false);
    const first = fs.readFileSync(file, 'utf8');
    check('word: a README opened and saved with no edit is written back as it was, HTML, code, table, links and definitions',
      same, same ? 'unchanged' : JSON.stringify(first.slice(0, 400)));

    // A click on the first box ticks it; the save writes [x].
    await js(`document.querySelector('.wd-page .wd-cc[data-kind="checkBox"]')?.click(), 1`);
    await until(() => js(`document.querySelector('.wd-page .wd-cc[data-kind="checkBox"]')?.getAttribute('aria-checked') === 'true'`), 'the box ticked', 4000).catch(() => false);
    await js(`document.querySelector('.wd-page')?.focus(), 1`);
    await press(wc, 's', { modifiers: ['control'] });
    const ticked = await until(() => fs.readFileSync(file, 'utf8').includes('- [x] Order milk\n- [x] Pay the rent'), 'the tick saved', 6000).then(() => true).catch(() => false);
    const after = fs.readFileSync(file, 'utf8');
    check('word: a click on a task box ticks it, and the saved file says [x], the rest unchanged',
      ticked && after === README.replace('- [ ] Order milk', '- [x] Order milk'), JSON.stringify(after.slice(after.indexOf('- ['), after.indexOf('- [') + 60)));

    const complaints = await errorsIn(win);
    check('word: the Markdown round trip reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('word: the Markdown checks ran', false, err?.message || String(err));
  }
}
