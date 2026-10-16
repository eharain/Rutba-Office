// What a screen reader is told of each app's document, read from Chromium's
// own accessibility tree (the DevTools protocol's Accessibility domain, the
// tree Windows UI Automation is given), not from the DOM: Documents' page a
// text box of many lines named by the document, with its headings at their
// levels and its words; the Worksheets grid telling the active cell's
// address and value, and following the arrow keys; Presentations' strip
// naming each slide by its number and title and the stage the slide shown;
// Mail's list a grid of messages whose rows carry the sender and the
// subject, and the message open an article named by its subject; Calendar's
// days named by their dates and its events by their times and titles; and a
// PDF's pages named by their numbers with their words in them. Run alone
// with RUTBA_VERIFY_ONLY=screen-reader.

import { fixtures, message } from './verify-polish-pass.js';

/** The window's accessibility tree, flattened: role, name, value and level, with each node's ancestors' roles. */
async function axTree(win) {
  const dbg = win.webContents.debugger;
  const mine = !dbg.isAttached();
  if (mine) dbg.attach('1.3');
  try {
    await dbg.sendCommand('Accessibility.enable');
    const { nodes } = await dbg.sendCommand('Accessibility.getFullAXTree');
    const byId = new Map(nodes.map((n) => [n.nodeId, n]));
    const prop = (n, name) => n.properties?.find((p) => p.name === name)?.value?.value;
    const out = nodes.filter((n) => !n.ignored).map((n) => {
      const up = [];
      for (let p = byId.get(n.parentId); p; p = byId.get(p.parentId)) if (!p.ignored) up.push({ role: p.role?.value, name: p.name?.value || '' });
      return {
        role: n.role?.value || '', name: n.name?.value || '', value: n.value?.value ?? '', level: prop(n, 'level'),
        focused: prop(n, 'focused') === true, selected: prop(n, 'selected'), current: prop(n, 'current'), multiline: prop(n, 'multiline'), up,
      };
    });
    return out;
  } finally {
    await dbg.sendCommand('Accessibility.disable').catch(() => {});
    if (mine) dbg.detach();
  }
}

const within = (n, role) => n.up.some((u) => u.role === role);

export async function verifyScreenReader({ open, check, until, wait, press }, { dir }) {
  const fx = fixtures(dir, 'screen-reader-fixtures');
  const only = (process.env.RUTBA_A11Y_APPS || '').split(',').map((x) => x.trim()).filter(Boolean);
  const wanted = (app) => !only.length || only.includes(app);
  const ready = (win, sel, what) => until(() => win.webContents.executeJavaScript(`Boolean(document.querySelector(${JSON.stringify(sel)}))`), what, 15000).catch(() => {});
  const brief = (list) => list.slice(0, 6).map((n) => `${n.role}:${n.name.slice(0, 40)}`).join(' | ');
  /** Waits for the tree to show what `test` finds, the window's own answer read until it comes. */
  const treeUntil = async (win, test, what, ms = 6000) => {
    let last = [];
    let got = null;
    await until(async () => { last = await axTree(win); got = test(last); return got; }, what, ms).catch(() => {});
    return { got, tree: last };
  };

  if (wanted('word')) {
    try {
      const win = await open('word', fx.docx);
      await ready(win, '.wd-page [data-block="1"]', 'the page');
      const { got, tree } = await treeUntil(win, (t) => {
        const page = t.find((n) => n.role === 'textbox' && n.multiline === true && /polish\.docx/.test(n.name));
        const heads = t.filter((n) => n.role === 'heading' && within(n, 'textbox'));
        const words = t.some((n) => /Sales rose in every region/.test(n.name) && within(n, 'textbox'));
        return page && heads.length >= 3 && words ? { page: page.name, heads: heads.map((h) => `${h.level}:${h.name}`) } : null;
      }, 'the page in the tree');
      check('screen-reader: Documents: the page is a text box of many lines named by the document, its headings headings at their levels (Title 1, Heading 1, Heading 2) and its words in it',
        Boolean(got) && got.heads.join('|') === '1:Quarterly review|1:Where the year stands|2:What comes next', JSON.stringify(got || brief(tree.filter((n) => n.role === 'heading' || n.role === 'textbox'))));
    } catch (err) { check('screen-reader: the Documents pass ran', false, err.stack || err.message); }
  }

  if (wanted('sheets')) {
    try {
      const win = await open('sheets', fx.xlsx);
      const js = (code) => win.webContents.executeJavaScript(code);
      await ready(win, '.sh-cell', 'the grid');
      await js(`document.querySelector('.sh')?.focus(), 1`);
      const first = await treeUntil(win, (t) => t.find((n) => n.role === 'gridcell' && /^A1 Region$/.test(n.name) && n.up.some((u) => u.role === 'grid')), 'A1 in the tree');
      // The keyboard moves the selection; the tree follows it.
      await press(win.webContents, 'Right');
      await press(win.webContents, 'Down');
      const moved = await treeUntil(win, (t) => t.find((n) => n.role === 'gridcell' && /^B2 120$/.test(n.name)), 'B2 in the tree');
      const group = moved.tree.find((n) => n.role === 'group' && n.name === 'Sales' && n.focused);
      check("screen-reader: Worksheets: the sheet, named, tells the active cell's address and value (A1 Region) as a cell of a grid, and the arrow keys move it (B2 120)",
        Boolean(first.got) && Boolean(moved.got) && Boolean(group), JSON.stringify({ first: first.got?.name, moved: moved.got?.name, group: Boolean(group), cells: brief(moved.tree.filter((n) => n.role === 'gridcell')) }));
    } catch (err) { check('screen-reader: the Worksheets pass ran', false, err.stack || err.message); }
  }

  if (wanted('slides')) {
    try {
      const win = await open('slides', fx.pptx);
      await ready(win, '.sl-thumb', 'the strip');
      const { got, tree } = await treeUntil(win, (t) => {
        const one = t.find((n) => n.role === 'button' && n.name === 'Slide 1: Quarterly review');
        const two = t.find((n) => n.role === 'button' && n.name === 'Slide 2: Where we stand');
        const stage = t.find((n) => n.role === 'group' && n.name === 'Slide 1: Quarterly review');
        const pic = t.find((n) => n.role === 'image' && n.name === 'Quarterly review' && within(n, 'group'));
        return one && two && stage && pic ? { stage: stage.name, pic: pic.name } : null;
      }, 'the slides in the tree');
      // The protocol's tree carries no aria-current: read from the element the tree's button is (since 2.3.0 the button inside the slide's picture).
      const current = await win.webContents.executeJavaScript(`document.querySelector('.sl-thumb [aria-current="true"], .sl-thumb[aria-current="true"]')?.getAttribute('aria-label') || ''`);
      if (got) got.current = current;
      check('screen-reader: Presentations: each slide in the strip is named by its number and title, the current one said so, and the stage is the slide it shows, its picture named by the title',
        Boolean(got) && got.current === 'Slide 1: Quarterly review', JSON.stringify(got || brief(tree.filter((n) => /button|group|image/.test(n.role) && /Slide|review|stand/.test(n.name)))));
    } catch (err) { check('screen-reader: the Presentations pass ran', false, err.stack || err.message); }
  }

  if (wanted('calendar')) {
    try {
      const win = await open('calendar', fx.ics);
      await ready(win, '.cal-month-grid', 'the month');
      const today = new Date();
      const { got, tree } = await treeUntil(win, (t) => {
        const day = t.find((n) => n.role === 'group' && n.name === today.toLocaleDateString('en', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' }));
        const ev = t.find((n) => n.role === 'button' && /Stand-up/.test(n.name) && /9[:.]00/.test(n.name));
        return day && ev ? { day: day.name, event: ev.name } : null;
      }, "today's events in the tree");
      const current = await win.webContents.executeJavaScript(`document.querySelector('.cal-day[aria-current="date"]')?.getAttribute('aria-label') || ''`);
      if (got) got.current = current;
      check("screen-reader: Calendar: today's day is a group named by its date in full and said to be today, its events buttons named by their times and titles",
        Boolean(got) && got.current === got.day, JSON.stringify(got || brief(tree.filter((n) => (n.role === 'group' && /October/.test(n.name)) || /Stand/.test(n.name)))));
    } catch (err) { check('screen-reader: the Calendar pass ran', false, err.stack || err.message); }
  }

  if (wanted('pdf')) {
    try {
      const win = await open('pdf', fx.pdf);
      await ready(win, '.pdf-page .textLayer[data-ready]', 'the text layer');
      const { got, tree } = await treeUntil(win, (t) => {
        const page = t.find((n) => n.role === 'group' && n.name === 'Page 1');
        const words = t.find((n) => /Quarterly review/.test(n.name) && n.up.some((u) => u.role === 'group' && u.name === 'Page 1'));
        return page && words ? { page: page.name, words: words.name } : null;
      }, 'the first page in the tree');
      check('screen-reader: PDF: each page is a group named by its number, the words on it read from its text layer',
        Boolean(got), JSON.stringify(got || brief(tree.filter((n) => n.role === 'group' || /Quarterly/.test(n.name)))));
    } catch (err) { check('screen-reader: the PDF pass ran', false, err.stack || err.message); }
  }

  if (wanted('mail')) {
    let accountId = null;
    let js = null;
    try {
      const win = await open('mail');
      js = (code) => win.webContents.executeJavaScript(code);
      const mail = (method, args) => js(`window.rutbaOffice.mail.${method}(${JSON.stringify(args ?? {})})`);
      await until(() => js(`document.querySelectorAll('.ml-accounts button').length > 0 || Boolean(document.querySelector('.rw-empty'))`), 'the Mail window', 8000).catch(() => {});
      // The new account listed before it is chosen.
      accountId = (await mail('addAccount', { account: { email: 'reader@checks.example', name: 'Reader (checks)', local: true } })).id;
      await mail('deliverTest', { accountId, folder: 'Inbox', raw: message({ from: 'Lena Weiss <lena.weiss@example.com>', subject: 'The autumn letter', body: 'Three things happened this month.', date: new Date().toUTCString() }) });
      await until(() => js(`[...document.querySelectorAll('.ml-accounts button')].some((n) => /reader@checks|Reader \\(checks\\)/.test(n.textContent + ' ' + (n.title || '')))`), 'the account listed', 8000).catch(() => {});
      await js(`(() => { const b = [...document.querySelectorAll('.ml-accounts button')].find((n) => /reader@checks|Reader \\(checks\\)/.test(n.textContent + ' ' + (n.title || ''))); b?.click(); return 1; })()`);
      await until(() => js(`[...document.querySelectorAll('.ml-row')].some((r) => /autumn letter/.test(r.textContent))`), 'the message listed', 8000).catch(() => {});
      await until(() => js(`document.querySelectorAll('.ml-row').length >= 1`), 'the message', 8000).catch(() => {});
      await js(`[...document.querySelectorAll('.ml-row')].find((r) => /autumn letter/.test(r.textContent))?.click(), 1`);
      await until(() => js(`Boolean(document.querySelector('.ml-reader'))`), 'the message open', 6000).catch(() => {});
      const { got, tree } = await treeUntil(win, (t) => {
        const grid = t.find((n) => n.role === 'grid' && n.name === 'Messages');
        const row = t.find((n) => n.role === 'row' && /Lena Weiss/.test(n.name) && /The autumn letter/.test(n.name) && within(n, 'grid'));
        const article = t.find((n) => n.role === 'article' && n.name === 'The autumn letter');
        const sender = t.find((n) => /Lena Weiss/.test(n.name) && n.up.some((u) => u.role === 'article'));
        return grid && row && article && sender ? { row: row.name.slice(0, 60), selected: row.selected, article: article.name } : null;
      }, 'the list and the message in the tree');
      check("screen-reader: Mail: the list is a grid of messages whose rows carry the sender and the subject, the open one selected, and the message open an article named by its subject with its sender in it",
        Boolean(got) && got.selected === true, JSON.stringify(got || brief(tree.filter((n) => /grid|row|article/.test(n.role)))));
    } catch (err) {
      check('screen-reader: the Mail pass ran', false, err.stack || err.message);
    } finally {
      if (accountId && js) await js(`window.rutbaOffice.mail.removeAccount({ id: ${JSON.stringify(accountId)} })`).catch(() => {});
    }
  }
}
