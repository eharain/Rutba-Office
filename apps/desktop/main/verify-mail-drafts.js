// Mail: a draft opens to be finished, not read (owner, 2026-10-20: "a draft
// email is not editable at all it opens it in read view only").
//
// A draft is saved with To, Cc, Bcc, a subject, formatted words and an
// attachment. Opened from Drafts, it is in Compose with all of them, not in
// the reading pane. Its subject is changed and it is saved: Drafts still has
// one draft, with the new subject and the attachment, not two. Opened again
// and sent, it leaves Drafts. Compose docks in the reading pane, goes full
// and docks again, and pops out to a window of its own (owner, 2026-10-21).
// Run alone with RUTBA_VERIFY_ONLY=mail-drafts.

import fs from 'node:fs';
import path from 'node:path';
import { BrowserWindow } from 'electron';

/**
 * How Compose is laid out where it is: the address boxes beside Subject, the
 * buttons along the top, what runs off the side, and the room under the
 * editor (owner, 2026-10-27: "fix the email response screen layout ... when
 * its popped out when its dialog and when its docked in").
 */
const LAYOUT = `(() => {
  const root = document.querySelector('.ml-compose');
  if (!root) return null;
  const box = (n) => (n ? n.getBoundingClientRect() : null);
  const control = (name) => [...root.querySelectorAll('.rw-field')].find((f) => f.querySelector('label')?.textContent === name)?.querySelector('input, select');
  const wide = (n) => Math.round(box(n)?.width || 0);
  const head = root.querySelector('.ml-compose-page-head');
  const body = root.querySelector('.ml-compose-page-body');
  const editor = root.querySelector('.ml-rich, .ml-compose-body');
  const form = root.querySelector('.ml-form');
  const buttons = [...head.querySelectorAll('button')].filter((b) => b.offsetParent);
  const h = box(head);
  return {
    width: Math.round(box(root).width),
    to: wide(control('To')), cc: wide(control('Cc')), bcc: wide(control('Bcc')), subject: wide(control('Subject')),
    headOut: buttons.filter((b) => box(b).right > h.right + 1 || box(b).left < h.left - 1 || box(b).bottom > h.bottom + 1).map((b) => b.textContent.trim()),
    overflowX: body.scrollWidth - body.clientWidth,
    overflowY: body.scrollHeight - body.clientHeight,
    gap: Math.round(box(body).bottom - box(form).bottom),
    editor: Math.round(box(editor).height),
    title: Math.round(box(head.querySelector('h2')).left - box(form).left),
  };
})()`;

/**
 * Laid out as a person would want it: To, Cc and Bcc as wide as Subject (To
 * less the Cc / Bcc switch beside it), every button in sight, nothing off the
 * side, and the editor filling the height down to the foot, with no page to
 * scroll. In full view and a window of its own the title starts where the
 * fields do.
 */
function laidOut(m, { column = false } = {}) {
  if (!m) return false;
  const along = (w) => w === 0 || Math.abs(w - m.subject) <= 2;
  return m.to >= m.subject - 120 && along(m.cc) && along(m.bcc) && m.headOut.length === 0 && m.overflowX <= 0 && m.overflowY <= 1 && m.gap <= 30 && m.editor >= 160 && (!column || Math.abs(m.title) <= 2);
}

/** A picture of the window for whoever is working on the layout, when RUTBA_VERIFY_SHOTS names a folder. */
async function shot(win, name) {
  const dir = process.env.RUTBA_VERIFY_SHOTS;
  if (!dir) return;
  fs.mkdirSync(dir, { recursive: true });
  fs.writeFileSync(path.join(dir, `${name}.png`), (await win.webContents.capturePage()).toPNG());
}

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn
 */
export async function verifyMailDrafts(h) {
  const { open, check, until, wait, errorsIn } = h;
  const stamp = Date.now();
  const subject = `Draft: the quarter's numbers ${stamp}`;
  let win = null;
  try {
    win = await open('mail');
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`document.querySelectorAll('.ml-accounts button').length > 0`), 'the seeded account', 8000);
    const accountId = await js(`(async () => (await window.rutbaOffice.mail.accounts())[0]?.id || null)()`);
    // The draft, as Save draft keeps one: everything a person wrote.
    const file = await js(`window.rutbaOffice.fs.temp({ ext: 'txt', bytes: new TextEncoder().encode('Figures attached.') }).then((r) => r.path)`);
    const draft = { to: 'Amina Yusuf <amina@work.example>', cc: 'tomas@berg-partners.example', bcc: 'files@work.example', subject, text: 'The board pack goes out Friday.', html: '<p>The board pack goes out <b>Friday</b>.</p>', attachments: [{ filename: 'figures.txt', path: file, size: 17 }] };
    await js(`window.rutbaOffice.mail.saveDraft({ accountId: ${JSON.stringify(accountId)}, draft: ${JSON.stringify(draft)} })`);
    // The window reads its folders again, as it does when mail arrives.
    await win.webContents.reload();
    await until(() => js(`document.querySelectorAll('.ml-accounts button').length > 0`), 'the account again', 8000);
    const drafts = () => js(`window.rutbaOffice.mail.messages({ accountId: ${JSON.stringify(accountId)}, folder: 'Drafts', query: '', limit: 200 }).then((r) => (r.rows || r).filter((m) => (m.subject || '').includes(${JSON.stringify(String(stamp))})).map((m) => ({ id: m.id, subject: m.subject })))`);
    const openDrafts = () => js(`(() => { const f = [...document.querySelectorAll('[title="Drafts"]')].find((n) => n.textContent.trim().startsWith('Drafts')); f?.click(); return Boolean(f); })()`);
    const openRow = (words) => js(`(() => { const r = [...document.querySelectorAll('.ml-row')].find((row) => row.querySelector('.ml-subject')?.textContent.includes(${JSON.stringify(String(stamp))}) && row.querySelector('.ml-subject')?.textContent.includes(${JSON.stringify(words)})); if (!r) return false; r.click(); return true; })()`);
    const field = (name) => `[...document.querySelectorAll('.ml-compose .rw-field')].find((f) => f.querySelector('label')?.textContent === ${JSON.stringify(name)})?.querySelector('input')?.value ?? null`;

    await until(async () => (await openDrafts()) && (await wait(400), await openRow('Draft:')), 'the draft in Drafts', 10000).catch(() => {});
    const composing = await until(() => js(`Boolean(document.querySelector('.ml-compose-docked .ml-form'))`), 'Compose docked in the reading pane', 8000).then(() => true).catch(() => false);
    const shown = await js(`({ to: ${field('To')}, cc: ${field('Cc')}, bcc: ${field('Bcc')}, subject: ${field('Subject')}, files: [...document.querySelectorAll('.ml-compose .ml-attachment')].map((n) => n.textContent), noDialog: !document.querySelector('.rw-dialog .ml-form'), list: Boolean(document.querySelector('.ml-list')) })`);
    check('mail drafts: a draft opened from Drafts is in Compose, docked in the reading pane beside the list (owner, 2026-10-21), with its To, Cc, Bcc, subject and attachment',
      composing && shown.noDialog && shown.list && /amina@work\.example/.test(shown.to || '') && /tomas@/.test(shown.cc || '') && /files@/.test(shown.bcc || '') && shown.subject === subject && shown.files.some((x) => x.includes('figures.txt')),
      JSON.stringify({ composing, shown }));
    const laidDocked = await js(LAYOUT);
    await shot(win, 'compose-docked');
    // A narrow window: the reading pane is narrow too, and nothing may go out of sight sideways.
    const [w0, h0] = win.getSize();
    win.setSize(1000, 760);
    await wait(500);
    const laidNarrow = await js(LAYOUT);
    await shot(win, 'compose-docked-narrow');
    win.setSize(w0, h0);
    await wait(300);

    // Full view (owner, 2026-10-20: "it should have conversion to full view as normal email"):
    // the message fills the window with what was written, Smaller docks it again, and the choice is kept.
    await js(`document.querySelector('.ml-compose-full')?.click(), 1`);
    const fullShown = await until(() => js(`Boolean(document.querySelector('.ml-compose-page .ml-form')) && !document.querySelector('.ml-compose-docked')`), 'the full view', 4000).then(() => true).catch(() => false);
    const inFull = await js(`({ subject: ${field('Subject')}, cc: ${field('Cc')}, body: document.querySelector('.ml-compose-page .ml-rich')?.textContent || '', files: document.querySelectorAll('.ml-compose-page .ml-attachment').length, wide: Math.round(document.querySelector('.ml-compose-page')?.getBoundingClientRect().width || 0), window: window.innerWidth })`);
    const keptFull = await js(`window.rutbaOffice.store.get({ key: 'mail.composePlace', fallback: null })`);
    const laidFull = await js(LAYOUT);
    await shot(win, 'compose-full');
    await js(`document.querySelector('.ml-compose-small')?.click(), 1`);
    const docked = await until(() => js(`Boolean(document.querySelector('.ml-compose-docked .ml-form')) && !document.querySelector('.ml-compose-page')`), 'docked again', 4000).then(() => true).catch(() => false);
    const body = await js(`document.querySelector('.ml-compose-docked .ml-rich')?.textContent || ''`);
    const keptDocked = await js(`window.rutbaOffice.store.get({ key: 'mail.composePlace', fallback: null })`);
    check('mail drafts: Full view puts the message in the whole window with everything written in it, Smaller docks it in the reading pane again, and the choice is kept for the next message',
      fullShown && inFull.subject === subject && /tomas@/.test(inFull.cc || '') && /board pack/.test(inFull.body) && inFull.files === 1 && inFull.wide >= inFull.window - 2 && keptFull === 'full' && docked && /board pack/.test(body) && keptDocked === 'docked',
      JSON.stringify({ fullShown, inFull, keptFull, docked, body: body.slice(0, 40), keptDocked }));

    // A new subject, saved: one draft, the new subject, the attachment kept.
    await js(`(() => { const n = [...document.querySelectorAll('.ml-compose .rw-field')].find((f) => f.querySelector('label')?.textContent === 'Subject')?.querySelector('input'); if (!n) return 0; Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(n, ${JSON.stringify(`Draft, edited: the quarter's numbers ${stamp}`)}); n.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
    await wait(200);
    await js(`[...document.querySelectorAll('.ml-compose .rw-btn, .ml-compose button')].find((b) => b.textContent.trim() === 'Save draft')?.click(), 1`);
    await until(() => js(`!document.querySelector('.ml-compose .ml-form')`), 'Compose closed', 6000).catch(() => {});
    await wait(500);
    const after = await drafts();
    const kept = after.length === 1 ? await js(`window.rutbaOffice.mail.message({ accountId: ${JSON.stringify(accountId)}, folder: 'Drafts', id: ${JSON.stringify(after[0].id)} })`) : null;
    check('mail drafts: saved again it replaces itself: one draft, with the new subject, its Cc and its attachment',
      after.length === 1 && /edited/.test(after[0].subject) && (kept?.cc || []).length === 1 && (kept?.attachments || []).some((a) => a.filename === 'figures.txt'),
      JSON.stringify({ after, cc: kept?.cc, files: (kept?.attachments || []).map((a) => a.filename) }));

    // Opened again and sent: it leaves Drafts.
    await openDrafts();
    await wait(400);
    await until(() => openRow('edited'), 'the edited draft', 8000).catch(() => {});
    await until(() => js(`Boolean(document.querySelector('.ml-compose .ml-form'))`), 'Compose again', 8000).catch(() => {});
    await js(`[...document.querySelectorAll('.ml-compose .rw-btn, .ml-compose button')].find((b) => b.textContent.trim() === 'Send' && !b.disabled)?.click(), 1`);
    const gone = await until(async () => (await drafts()).length === 0, 'the draft gone', 8000).then(() => true).catch(() => false);
    check('mail drafts: a draft sent leaves Drafts', gone, JSON.stringify(await drafts()));

    // Pop out (owner, 2026-10-21: "should also expand to a separate window as the read email
    // window does"): the message as written goes to a Mail window of its own, the composer
    // here closes, and Discard there closes that window.
    const before = new Set(BrowserWindow.getAllWindows().map((w) => w.id));
    await js(`document.querySelector('.ml-compose-cta button')?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.ml-compose .ml-form'))`), 'a new message', 6000).catch(() => {});
    await js(`(() => { const n = [...document.querySelectorAll('.ml-compose .rw-field')].find((f) => f.querySelector('label')?.textContent === 'Subject')?.querySelector('input'); if (!n) return 0; Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(n, ${JSON.stringify(`Popped out ${stamp}`)}); n.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
    await wait(200);
    await js(`document.querySelector('.ml-compose-popout')?.click(), 1`);
    const findPopped = () => BrowserWindow.getAllWindows().find((w) => !before.has(w.id) && !w.isDestroyed() && /compose=/.test(w.webContents.getURL())) || null;
    await until(findPopped, 'the window of its own', 8000).catch(() => {});
    const popped = findPopped();
    const pjs = (code) => popped.webContents.executeJavaScript(code);
    let there = null;
    let laidAlone = null;
    if (popped) {
      await until(() => pjs(`Boolean(document.querySelector('.ml-compose-page.alone .ml-form'))`), 'Compose in its window', 8000).catch(() => {});
      await wait(300);
      laidAlone = await pjs(LAYOUT);
      await shot(popped, 'compose-window');
      there = await pjs(`({ subject: [...document.querySelectorAll('.ml-compose .rw-field')].find((f) => f.querySelector('label')?.textContent === 'Subject')?.querySelector('input')?.value || '', popOut: Boolean(document.querySelector('.ml-compose-popout')) })`);
    }
    const leftHere = await until(() => js(`!document.querySelector('.ml-compose')`), 'the composer here to close', 4000).then(() => true, () => false);
    if (popped) await pjs(`[...document.querySelectorAll('.ml-compose .rw-btn')].find((b) => b.textContent.trim() === 'Discard')?.click(), 1`);
    const closed = popped ? await until(() => popped.isDestroyed() || null, 'the window to close', 6000).then(() => true, () => false) : false;
    check('mail drafts: Pop out puts the message as written in a Mail window of its own, closes it here, and Discard there closes that window',
      Boolean(popped) && there?.subject === `Popped out ${stamp}` && there?.popOut === false && leftHere && closed,
      JSON.stringify({ popped: Boolean(popped), there, leftHere, closed }));

    check('mail drafts: Compose is laid out docked, in full view and in a window of its own: To, Cc and Bcc as wide as Subject, every button in sight, and the editor filling the height (owner, 2026-10-27)',
      laidOut(laidDocked) && laidOut(laidFull, { column: true }) && laidOut(laidAlone, { column: true }) && laidOut(laidNarrow ? { ...laidNarrow, overflowY: 0, gap: 0 } : null),
      JSON.stringify({ docked: laidDocked, narrow: laidNarrow, full: laidFull, window: laidAlone }));

    const complaints = await errorsIn(win);
    check('mail drafts: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('mail drafts: the check ran to the end', false, err?.message || String(err));
  }
}
