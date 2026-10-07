// Mail: junk.
//
// Home → Junk marks a message junk and the filter learns from it; Junk
// Email Options shows what it has learned, teaches it from the folders
// already there, and keeps a Blocked Senders list; mail that arrives is
// then filed by it — junk to Junk, with the reading pane saying why, good
// mail left in the Inbox — and Not junk puts a message back and teaches
// the filter it was good. The options are the account's own, as Outlook
// keeps them; Never block this group or mailing list keeps a list's mail
// out of Junk. Run alone with RUTBA_VERIFY_ONLY=junk.
import fs from 'node:fs';
import path from 'node:path';

const JUNK = [
  ['prize@claims.example', 'You have WON a cash prize', 'Claim your lottery prize now, send your bank details to receive the transfer of one million dollars.'],
  ['offers@pills.example', 'Cheap pills online', 'Buy cheap pills online with no prescription, discount offer, click here http://pills.example/buy'],
  ['agent@claims.example', 'Urgent transfer pending', 'Urgent: your transfer of funds is pending, send bank details and a processing fee today.'],
  ['deals@shop.example', 'Limited offer just for you', 'Limited time offer, click here to claim your free gift card, the discount ends today.'],
  ['coins@moon.example', 'Double your money', 'Invest now and double your money, guaranteed returns, click here http://moon.example'],
  ['lotto@claims.example', 'Lottery winner notification', 'Congratulations winner, your address won the international lottery, claim your prize with your bank details.'],
];
const GOOD = [
  ['sam@work.example', 'Minutes from Tuesday', 'Here are the minutes from the planning meeting on Tuesday. The next review is on the 14th.'],
  ['ana@work.example', 'Draft report', 'The draft of the quarterly report is in the shared folder, could you read section two before Friday?'],
  ['sam@work.example', 'Lunch on Thursday?', 'Are you free for lunch on Thursday? The usual place near the office.'],
  ['it@work.example', 'Laptop upgrade schedule', 'Your laptop upgrade is scheduled for Monday morning, please save your work beforehand.'],
  ['ana@work.example', 'Re: the budget review', 'Thanks for the numbers, the budget review meeting moves to Wednesday afternoon.'],
  ['sam@work.example', 'Room booking', 'I booked the large meeting room for the review on Wednesday afternoon.'],
];

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn
 */
export async function verifyMailJunk(h) {
  const { open, check, until, wait, errorsIn } = h;
  const capture = async (name) => {
    if (!process.env.RUTBA_VERIFY_CAPTURE || !win) return;
    win.webContents.invalidate();
    await wait(400);
    fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, name), (await win.webContents.capturePage()).toPNG());
  };
  let win = null;
  let js = null;
  let accountId = null;
  const stamp = Date.now();
  // A minute apart, counting on from now, so the list shows them in order.
  const raw = ([from, subject, body], n) => [
    `From: <${from}>`, 'To: You <you@example.com>', `Subject: ${subject}`, `Date: ${new Date(stamp + n * 60000).toUTCString()}`,
    `Message-ID: <junk-${stamp}-${n}@example.org>`, 'MIME-Version: 1.0', 'Content-Type: text/plain; charset=utf-8', '', body, '',
  ].join('\r\n');
  try {
    win = await open('mail');
    js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`document.querySelectorAll('.ml-accounts button').length > 0`), 'the seeded account in the sidebar', 8000);
    accountId = await js(`(async () => (await window.rutbaOffice.mail.accounts())[0]?.id || null)()`);
    const A = JSON.stringify(accountId);
    const mail = (method, args) => js(`window.rutbaOffice.mail.${method}(${JSON.stringify(args)})`);
    const deliver = (spec, n) => mail('deliverTest', { accountId, folder: 'Inbox', raw: raw(spec, n) });
    const folders = await mail('folders', { accountId });
    const inbox = folders.find((f) => /^inbox$/i.test(f.name || f.path))?.path || 'Inbox';
    const subjectsIn = async (folder) => (await mail('messages', { accountId, folder, limit: 500 })).rows.map((r) => r.subject);
    await mail('setJunk', { patch: { level: 'low', safe: [], blocked: [] } });

    // Before it has learned anything, the filter leaves everything where it lands.
    for (const [i, spec] of [...JUNK, ...GOOD].entries()) await deliver(spec, i);
    const untaught = await subjectsIn(inbox);
    check('mail: a junk filter that has learned nothing files nothing',
      [...JUNK, ...GOOD].every(([, s]) => untaught.includes(s)), `${untaught.length} in the Inbox`);

    // Home → Junk → Mark as junk, on the first junk message: moved, and learned.
    // The seeded mailbox is a folder of its own; the mail arrived in the Inbox.
    const openFolder = (name) => js(`(() => { const f = [...document.querySelectorAll('[title="${name}"]')].find((n) => n.textContent.trim().startsWith(${JSON.stringify(name)})); if (!f) return 'no folder'; f.click(); return 'opened'; })()`);
    await until(async () => (await openFolder(inbox)) === 'opened', 'the Inbox in the sidebar', 5000);
    await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Home')?.click(), 'tab'`);
    const openRow = (subject) => js(`(() => { const r = [...document.querySelectorAll('.ml-row')].find((row) => row.querySelector('.ml-subject')?.textContent.includes(${JSON.stringify(subject)})); if (!r) return 'no row among: ' + [...document.querySelectorAll('.ml-row .ml-subject')].slice(0, 6).map((n) => n.textContent.trim()).join(' | ') + ' — ' + document.title; r.click(); return 'opened'; })()`);
    const firstOpen = await until(async () => (await openRow(JUNK[0][1])) === 'opened', 'the first junk message in the list', 6000).catch(async () => openRow(JUNK[0][1]));
    if (firstOpen !== true) throw new Error(firstOpen);
    await until(() => js(`(document.querySelector('.ml-head h2')?.textContent || '').includes(${JSON.stringify(JUNK[0][1])})`), 'it in the reading pane', 5000);
    const junkMenu = async (label) => {
      await js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Junk'); if (!b) return 'no button'; b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true })); b.click(); return 'clicked'; })()`);
      await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(label)}))`), `"${label}" in the Junk menu`, 3000);
      return js(`(() => { const b = [...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(label)}); if (!b || b.disabled) return 'missing'; b.click(); return 'clicked'; })()`);
    };
    const marked = await junkMenu('Mark as junk');
    const moved = await until(async () => !(await subjectsIn(inbox)).includes(JUNK[0][1]), 'the message out of the Inbox', 5000).catch(() => false);
    const afterOne = await mail('junk', { accountId });
    check('mail: Home → Junk → Mark as junk moves the message to Junk and the filter learns from it',
      marked === 'clicked' && moved === true && afterOne.learned.junk === 1, `${marked}; learned ${JSON.stringify(afterOne.learned)}`);

    // The rest of the junk put in Junk the way a drag does: each one learned.
    const rows = (await mail('messages', { accountId, folder: inbox, limit: 500 })).rows;
    const junkIds = rows.filter((r) => JUNK.slice(1).some(([, s]) => s === r.subject)).map((r) => r.id);
    await mail('move', { accountId, folder: inbox, ids: junkIds, to: 'Junk' });

    // Junk email options: what it has learned, and learning from the folders.
    await junkMenu('Junk email options…');
    await until(() => js(`Boolean(document.querySelector('.rw-dialog[aria-label="Junk email options"] .ml-junk-learned'))`), 'Junk email options', 5000);
    const scope = `document.querySelector('.rw-dialog[aria-label="Junk email options"]')`;
    const before = await js(`${scope}.querySelector('.ml-junk-learned span').textContent`);
    await js(`[...${scope}.querySelectorAll('button')].find((b) => b.textContent.trim() === 'Learn from my folders')?.click(), 'learn'`);
    const learned = await until(() => js(`/has learned from \\d+ junk and \\d+ good messages\\.$/.test(${scope}.querySelector('.ml-junk-learned span').textContent.trim())`), 'the filter ready', 6000).catch(() => false);
    const after = await js(`${scope}.querySelector('.ml-junk-learned span').textContent`);
    await capture('mail-junk-options.png');
    check('mail: Junk email options says what the filter has learned, and Learn from my folders teaches it the Inbox and Junk',
      /judges nothing until it has seen 5 of each/.test(before) && learned === true, `${before} → ${after}`);
    // A blocked domain, typed into the list.
    await js(`(() => { const el = ${scope}.querySelector('.ml-junk-blocked'); const setter = Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, 'value').set; setter.call(el, '@blocked.example'); el.dispatchEvent(new Event('input', { bubbles: true })); return true; })()`);
    await js(`[...${scope}.querySelectorAll('.rw-dialog-foot button')].find((b) => b.textContent.trim() === 'Save')?.click(), 'saved'`);
    await until(() => js(`!document.querySelector('.rw-dialog[aria-label="Junk email options"]')`), 'the dialog to close', 5000);
    const saved = await mail('junk', { accountId });
    const shared = await mail('junk', {});
    check('mail: the Blocked Senders list is kept as typed, as the account\'s own options', JSON.stringify(saved.blocked) === '["@blocked.example"]' && saved.own === true && shared.blocked.length === 0,
      `account ${JSON.stringify(saved.blocked)} (own ${saved.own}); shared ${JSON.stringify(shared.blocked)}`);

    // Now it files what arrives.
    const NEW_JUNK = ['winner@claims.example', 'Claim your prize today', 'You are a winner, send your bank details to claim the cash prize transfer.'];
    const NEW_GOOD = ['newcolleague@work.example', 'Question about the report', 'A question about section two of the quarterly report before the review meeting.'];
    const BLOCKED = ['hello@blocked.example', 'Hello there', 'Just saying hello.'];
    await deliver(NEW_JUNK, 100);
    await deliver(NEW_GOOD, 101);
    await deliver(BLOCKED, 102);
    const inInbox = await subjectsIn(inbox);
    const inJunk = await subjectsIn('Junk');
    check('mail: arriving junk goes to Junk, good mail stays in the Inbox, and a blocked sender goes to Junk',
      inJunk.includes(NEW_JUNK[1]) && !inInbox.includes(NEW_JUNK[1]) && inInbox.includes(NEW_GOOD[1]) && inJunk.includes(BLOCKED[1]),
      `Junk: ${inJunk.filter((s) => [NEW_JUNK[1], BLOCKED[1], NEW_GOOD[1]].includes(s)).join(', ')}; Inbox has the good one: ${inInbox.includes(NEW_GOOD[1])}`);

    // In Junk, the reading pane says why, and Not junk puts it back.
    await until(async () => (await openFolder('Junk')) === 'opened', 'Junk in the sidebar', 5000);
    await until(async () => (await openRow(NEW_JUNK[1])) === 'opened', 'the filed message in Junk', 6000);
    await until(() => js(`Boolean(document.querySelector('.ml-junk-note'))`), 'the note on why it is junk', 5000);
    const note = await js(`document.querySelector('.ml-junk-note')?.textContent || ''`);
    await capture('mail-junk-note.png');
    await js(`[...document.querySelectorAll('.ml-junk-note button')].find((b) => b.textContent.trim() === 'Not junk')?.click(), 'not junk'`);
    const back = await until(async () => (await subjectsIn(inbox)).includes(NEW_JUNK[1]), 'the message back in the Inbox', 5000).catch(() => false);
    const final = await mail('junk', { accountId });
    check('mail: a filed message says why it is in Junk, and Not junk puts it back in the Inbox and teaches the filter',
      /junk filter judged it junk/.test(note) && back === true && final.learned.good > saved.learned.good,
      `${note.trim()}; back ${back}; good ${saved.learned.good} → ${final.learned.good}`);

    // Never block this group or mailing list: with Safe Lists Only, a
    // list's post goes to Junk; pressed on it, the list is a Safe Recipient,
    // the post goes back, and the next one stays in the Inbox.
    await mail('setJunk', { accountId, patch: { level: 'safeOnly' } });
    const post = (subject, n) => [
      'From: <someone@example.org>', 'To: <dev@lists.example.org>', `Subject: ${subject}`, `Date: ${new Date(stamp + n * 60000).toUTCString()}`,
      `Message-ID: <list-${stamp}-${n}@example.org>`, 'List-Id: Developers <dev.lists.example.org>', 'List-Post: <mailto:dev@lists.example.org>',
      'MIME-Version: 1.0', 'Content-Type: text/plain; charset=utf-8', '', 'A post to the list.', '',
    ].join('\r\n');
    await mail('deliverTest', { accountId, folder: 'Inbox', raw: post('Build broken on main', 200) });
    const listFiled = (await subjectsIn('Junk')).includes('Build broken on main');
    await until(async () => (await openFolder('Junk')) === 'opened', 'Junk in the sidebar', 5000);
    await until(async () => (await openRow('Build broken on main')) === 'opened', 'the list post in Junk', 6000);
    await until(() => js(`(document.querySelector('.ml-head h2')?.textContent || '').includes('Build broken on main')`), 'it in the reading pane', 5000);
    const kept = await junkMenu('Never block this group or mailing list');
    const listBack = await until(async () => (await subjectsIn(inbox)).includes('Build broken on main'), 'the post back in the Inbox', 5000).then(() => true, () => false);
    const afterList = await mail('junk', { accountId });
    await mail('deliverTest', { accountId, folder: 'Inbox', raw: post('Build fixed on main', 201) });
    const nextStays = (await subjectsIn(inbox)).includes('Build fixed on main');
    check('mail: Junk → Never block this group or mailing list keeps the list\'s mail out of Junk',
      listFiled && kept === 'clicked' && listBack && JSON.stringify(afterList.safeRecipients) === '["dev@lists.example.org"]' && nextStays,
      `filed ${listFiled}; ${kept}; back ${listBack}; safe recipients ${JSON.stringify(afterList.safeRecipients)}; the next stays ${nextStays}`);

    // The options show it among the Safe Recipients, with the International lists beside.
    await junkMenu('Junk email options…');
    await until(() => js(`Boolean(document.querySelector('.rw-dialog[aria-label="Junk email options"] .ml-junk-recipients'))`), 'Junk email options', 5000);
    const shown = await js(`(() => { const d = document.querySelector('.rw-dialog[aria-label="Junk email options"]'); return { recipients: d.querySelector('.ml-junk-recipients').value, tlds: Boolean(d.querySelector('.ml-junk-tlds')), encodings: d.querySelectorAll('.ml-junk-encoding input').length }; })()`);
    await capture('mail-junk-options-2.png');
    await js(`[...document.querySelectorAll('.rw-dialog[aria-label="Junk email options"] .rw-dialog-foot button')].find((b) => b.textContent.trim() === 'Cancel')?.click(), 'closed'`);
    check('mail: Junk email options lists the Safe Recipients, the blocked top-level domains and the encodings',
      shown.recipients === 'dev@lists.example.org' && shown.tlds && shown.encodings >= 10, JSON.stringify(shown));

    const complaints = await errorsIn(win);
    check('mail: the junk checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('mail: the junk check ran', false, err.message);
  } finally {
    // The account is a fixture the other mail checks share: no filter left on for them, its own options or the shared.
    const off = { level: 'off', safe: [], blocked: [], safeRecipients: [], blockedTlds: [], blockedEncodings: [] };
    if (js) await js(`window.rutbaOffice.mail.setJunk(${JSON.stringify({ patch: off })})`).catch(() => {});
    if (js && accountId) await js(`window.rutbaOffice.mail.setJunk(${JSON.stringify({ accountId, patch: off })})`).catch(() => {});
  }
}
