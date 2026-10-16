// Mail: a link in a message opens in the browser, and junk's links are
// guarded (owner, 2026-10-16: "the link in email do not open it leads to a
// blank page . it should open browser instead while spam should have a
// protection on links").
//
// Four messages arrive, each a page whose first thing is a large link. A
// click on the ordinary link opens it at once, and the message stays in its
// frame rather than turning blank. A link whose words name another site asks
// first, and Cancel opens nothing. A message moved to Junk asks before its
// link opens, and Open in Browser opens it. A mailto: link starts a new
// message to that address. The links are noted rather than opened (ipc.js,
// under the checks). Run alone with RUTBA_VERIFY_ONLY=mail-links.

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn
 */
export async function verifyMailLinks(h) {
  const { open, check, until, wait, errorsIn } = h;
  const stamp = Date.now();
  const page = (n, subject, body) => [
    'From: Sam <sam@work.example>', 'To: You <you@example.com>', `Subject: ${subject}`, `Date: ${new Date(stamp + n * 60000).toUTCString()}`,
    `Message-ID: <links-${stamp}-${n}@example.org>`, 'MIME-Version: 1.0', 'Content-Type: text/html; charset=utf-8', '',
    `<html><body>${body}<p>The rest of the message.</p></body></html>`, '',
  ].join('\r\n');
  const big = (href, words, extra = '') => `<a href="${href}" ${extra} style="display:block;height:140px;font-size:28px">${words}</a>`;
  const MESSAGES = [
    ['Links: ordinary', big('https://news.example/today', 'Read today’s news', 'target="_blank"')],
    ['Links: misleading', big('https://evil.example/login', 'www.paypal.com')],
    ['Links: junk', big('https://offer.example/claim', 'Claim your prize')],
    ['Links: write to', big('mailto:ana@work.example?subject=Hello', 'Write to Ana')],
  ];
  const opened = () => [...(globalThis.__rutbaCheckOpened || [])];
  const answers = () => globalThis.__rutbaCheckDialogAnswers || [];
  let win = null;
  try {
    win = await open('mail');
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`document.querySelectorAll('.ml-accounts button').length > 0`), 'the seeded account in the sidebar', 8000);
    const accountId = await js(`(async () => (await window.rutbaOffice.mail.accounts())[0]?.id || null)()`);
    const mail = (method, args) => js(`window.rutbaOffice.mail.${method}(${JSON.stringify(args)})`);
    for (const [i, [subject, body]] of MESSAGES.entries()) await mail('deliverTest', { accountId, folder: 'Inbox', raw: page(i, subject, body) });
    const folders = await mail('folders', { accountId });
    const inbox = folders.find((f) => /^inbox$/i.test(f.name || f.path))?.path || 'Inbox';

    const openFolder = (name) => js(`(() => { const f = [...document.querySelectorAll('[title="${name}"]')].find((n) => n.textContent.trim().startsWith(${JSON.stringify(name)})); if (!f) return 'no folder'; f.click(); return 'opened'; })()`);
    const openRow = (subject) => js(`(() => { const r = [...document.querySelectorAll('.ml-row')].find((row) => row.querySelector('.ml-subject')?.textContent.includes(${JSON.stringify(subject)})); if (!r) return 'no row'; r.click(); return 'opened'; })()`);
    const read = async (folder, subject) => {
      await until(async () => (await openFolder(folder)) === 'opened', `${folder} in the sidebar`, 5000).catch(() => {});
      await until(async () => (await openRow(subject)) === 'opened', `"${subject}" in the list`, 6000).catch(() => {});
      await until(() => js(`(document.querySelector('.ml-head h2')?.textContent || '').includes(${JSON.stringify(subject)})`), `"${subject}" in the reading pane`, 5000).catch(() => {});
      await wait(1200);
    };
    /**
     * A real click on the message's first link. It goes through the DevTools
     * protocol, which hit-tests into the frame the way a mouse does. The frame
     * runs in a process of its own, and sendInputEvent reaches only the
     * window's own page.
     */
    const dbg = win.webContents.debugger;
    if (!dbg.isAttached()) dbg.attach('1.3');
    const clickLink = async () => {
      const at = await js(`(() => { const f = document.querySelector('.ml-body iframe'); if (!f) return null; const r = f.getBoundingClientRect(); return { x: Math.round(r.left + 50), y: Math.round(r.top + 30) }; })()`);
      if (!at) return 'no frame';
      await dbg.sendCommand('Input.dispatchMouseEvent', { type: 'mouseMoved', x: at.x, y: at.y });
      await dbg.sendCommand('Input.dispatchMouseEvent', { type: 'mousePressed', x: at.x, y: at.y, button: 'left', clickCount: 1 });
      await dbg.sendCommand('Input.dispatchMouseEvent', { type: 'mouseReleased', x: at.x, y: at.y, button: 'left', clickCount: 1 });
      return 'clicked';
    };
    const frameUrl = () => win.webContents.mainFrame.frames.map((f) => f.url).find((u) => u) || '';

    // An ordinary link: opened in the browser at once, and the message still in its frame.
    await read(inbox, 'Links: ordinary');
    const before = opened().length;
    const clicked = await clickLink();
    const went = await until(() => opened().length > before, 'the link handed to the browser', 4000).then(() => true).catch(() => false);
    await wait(400);
    const stays = frameUrl();
    check('mail: a link in a message opens in the browser, and the message stays in the reading pane rather than turning blank',
      clicked === 'clicked' && went && opened().at(-1) === 'https://news.example/today' && /^about:srcdoc/.test(stays),
      `${clicked}; opened ${JSON.stringify(opened().slice(before))}; the frame at ${stays}`);

    // A link whose words name another site: asked first, Cancel opens nothing.
    await read(inbox, 'Links: misleading');
    const asked = answers().length;
    answers().push({ kind: 'message', answer: { response: 2 } });
    const before2 = opened().length;
    await clickLink();
    const askedIt = await until(() => answers().length === asked, 'the warning', 4000).then(() => true).catch(() => false);
    await wait(500);
    check('mail: a link whose words name another site than it goes to asks first, and Cancel opens nothing',
      askedIt && opened().length === before2, `asked ${askedIt}; opened ${JSON.stringify(opened().slice(before2))}`);
    if (!askedIt) globalThis.__rutbaCheckDialogAnswers = answers().filter((a) => a.kind !== 'message');

    // A message in Junk: asked first, Open in Browser opens it. It is put in
    // Junk as a person does it (Home, Junk, Mark as junk), so the folder shows.
    await read(inbox, 'Links: junk');
    await js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Junk'); if (!b) return 'no button'; b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true })); b.click(); return 'opened'; })()`);
    await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === 'Mark as junk'))`), 'Mark as junk', 3000).catch(() => {});
    await js(`[...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === 'Mark as junk')?.click(), 1`);
    await until(async () => !(await mail('messages', { accountId, folder: inbox, limit: 500 })).rows.some((r) => r.subject === 'Links: junk'), 'the message in Junk', 5000).catch(() => {});
    const junkPath = (await mail('folders', { accountId })).find((f) => f.role === 'junk')?.path || 'Junk';
    await read(junkPath, 'Links: junk');
    const asked3 = answers().length;
    answers().push({ kind: 'message', answer: { response: 0 } });
    const before3 = opened().length;
    await clickLink();
    const askedJunk = await until(() => answers().length === asked3, 'the warning', 4000).then(() => true).catch(() => false);
    const openedJunk = await until(() => opened().length > before3, 'the link opened', 4000).then(() => true).catch(() => false);
    check('mail: a link in a message in Junk asks first, and Open in Browser opens it',
      askedJunk && openedJunk && opened().at(-1) === 'https://offer.example/claim', `asked ${askedJunk}; opened ${JSON.stringify(opened().slice(before3))}`);
    if (!askedJunk) globalThis.__rutbaCheckDialogAnswers = answers().filter((a) => a.kind !== 'message');

    // A mailto: link: a new message to that address.
    await read(inbox, 'Links: write to');
    const before4 = opened().length;
    await clickLink();
    const compose = await until(() => js(`(() => { const row = document.querySelector('.ml-to-row'); if (!row) return null; const to = [...row.querySelectorAll('input')].map((n) => n.value).join(' ') + ' ' + row.textContent; return /ana@work\\.example/.test(to) ? to.trim() : null; })()`), 'the new message', 5000).catch(() => null);
    check('mail: a mailto: link starts a new message to that address, and nothing is handed to the browser',
      Boolean(compose) && opened().length === before4, `${compose || 'no new message'}; opened ${JSON.stringify(opened().slice(before4))}`);

    try { dbg.detach(); } catch { /* not attached */ }
    const complaints = await errorsIn(win);
    check('mail links: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('mail links: the pass ran', false, err.stack || err.message);
  }
}
