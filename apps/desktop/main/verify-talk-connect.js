// Talk, Connect (Office 2.25), on the stand-in introducer: in Invite by
// email, a person from Contacts who is a contact already shows as one; one
// seen using Talk before (a contact once, taken off) gets Connect, which asks
// at once in Talk with no email; one who asked first gets Connect too, which
// makes them a contact on both sides. Another app opening Talk on a person
// (Contacts' Invite to Talk) opens the dialog with them in it. Run alone with
// RUTBA_VERIFY_ONLY=talk-connect.

import { BrowserWindow } from 'electron';
import { standInIntroducer } from './verify-talk.js';
import { labelKeyed } from './verify-keys.js';

const PAGE = (url) => `<!doctype html><meta charset="utf-8"><body><script>
const log = []; window.peer = { log };
const ws = new WebSocket(${JSON.stringify(url)});
ws.onopen = () => ws.send(JSON.stringify({ type: 'hello', token: 'ben' }));
ws.onmessage = (e) => log.push(JSON.parse(e.data));
window.peer.send = (m) => ws.send(JSON.stringify(m));
</script></body>`;

export async function verifyTalkConnect({ open, check, until, errorsIn }) {
  const ok = (p) => p.then(() => true, () => false);
  const intro = standInIntroducer({
    page: PAGE,
    people: { kai: { id: 'u-kai', name: 'Kai Lin', email: 'kai@example.com' }, mo: { id: 'u-mo', name: 'Mo Diaz', email: 'mo@example.com' } },
    contacts: [['u-amira', 'u-kai'], ['u-amira', 'u-mo']],
  });
  let win = null;
  let ben = null;
  let opened = null;
  const cards = [];
  try {
    const port = await intro.listen();
    win = await open('talk');
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('.tk-signin'))`), 'the sign-in', 8000);
    await js(`window.rutbaOffice.talk.signIn({ token: 'amira', server: ${JSON.stringify(`ws://127.0.0.1:${port}/ws`)} }).then(() => 1)`);
    await until(() => js(`Boolean(document.querySelector('.tk-contact[data-contact="u-kai"]'))`), 'signed in with her contacts', 8000);
    for (const [full, email] of [['Kai Lin', 'kai@example.com'], ['Ben Ortiz', 'ben@example.com'], ['Mo Diaz', 'mo@example.com'], ['Lee Park', 'lee@example.com']]) {
      cards.push(await js(`window.rutbaOffice.contacts.save({ contact: { name: { full: ${JSON.stringify(full)} }, emails: [{ value: ${JSON.stringify(email)}, type: 'home' }] }, book: 'local' })`));
    }
    // Kai was a contact and is taken off: still known to use Talk. Ben asks her first.
    await js(`window.rutbaOffice.talk.removeContact({ contactId: 'u-kai' })`);
    ben = new BrowserWindow({ show: false, webPreferences: { contextIsolation: true, nodeIntegration: false, partition: 'talk-connect-ben' } });
    await ben.loadURL(`http://127.0.0.1:${port}/ben`);
    const bjs = (code) => ben.webContents.executeJavaScript(code);
    await until(() => bjs(`window.peer.log.some((m) => m.type === 'welcome')`), 'Ben signed in', 6000);
    await bjs(`window.peer.send({ type: 'contact.request', email: 'amira@example.com', note: '' }), 1`);
    await until(() => js(`Boolean(document.querySelector('.tk-request'))`), 'Ben\'s request', 6000);

    await js(`document.querySelector('.tk-ie-btn').click(), 1`);
    const pick = (email) => `document.querySelector('.tk-ie-pick[data-email="${email}"]')`;
    const marked = await ok(until(() => js(`${pick('kai@example.com')}?.dataset.reach === 'uses-talk' && ${pick('ben@example.com')}?.dataset.reach === 'asked-you' && ${pick('mo@example.com')}?.dataset.reach === 'contact'`), 'each marked', 8000));
    const rows = await js(`[...document.querySelectorAll('.tk-ie-pick')].map((n) => ({ email: n.dataset.email, reach: n.dataset.reach, connect: Boolean(n.querySelector('.tk-connect-now')), disabled: n.querySelector('input').disabled, words: n.textContent }))`);
    check('talk: Invite by email marks each person from Contacts by what Talk has seen: a contact already (not invited again), someone who asked first and someone seen using Talk (both with Connect), and nobody known (invited by email)',
      marked && rows.find((r) => r.email === 'mo@example.com')?.disabled && !rows.find((r) => r.email === 'mo@example.com')?.connect
        && rows.find((r) => r.email === 'kai@example.com')?.connect && /Uses Talk/.test(rows.find((r) => r.email === 'kai@example.com')?.words)
        && rows.find((r) => r.email === 'ben@example.com')?.connect && rows.find((r) => r.email === 'lee@example.com')?.reach === '' && !rows.find((r) => r.email === 'lee@example.com')?.connect,
      JSON.stringify(rows));

    // Connect on Kai: asked at once, in Talk, no email.
    await js(`${pick('kai@example.com')}.querySelector('.tk-connect-now').click(), 1`);
    const kaiAsked = await ok(until(() => js(`document.querySelector('[data-asked="kai@example.com"]')?.dataset.told === 'talk:sent' && ${pick('kai@example.com')}?.dataset.reach === 'asked'`), 'Kai asked in Talk', 6000));
    const kaiWords = await js(`document.querySelector('[data-asked="kai@example.com"]')?.textContent || ''`);
    check('talk: Connect on someone who uses Talk asks them at once, in Talk, with no email, and the list of people asked says so',
      kaiAsked && /Asked in Talk/.test(kaiWords) && !intro.mailed().some((m) => m.to === 'kai@example.com'), JSON.stringify({ kaiAsked, kaiWords, mailed: intro.mailed() }));

    // Connect on Ben, who asked first: contacts on both sides.
    await js(`${pick('ben@example.com')}.querySelector('.tk-connect-now').click(), 1`);
    const benBoth = await ok(until(() => js(`Boolean(document.querySelector('.tk-contact[data-contact="u-ben"]')) && !document.querySelector('.tk-request')`), 'Ben a contact', 6000));
    const benTold = await ok(until(() => bjs(`window.peer.log.some((m) => m.type === 'contact.added' && m.contact.id === 'u-amira')`), 'Ben told', 5000));
    check('talk: Connect on someone who asked first makes them a contact on both sides at once', benBoth && benTold, JSON.stringify({ benBoth, benTold }));
    await js(`((list) => list.find((b) => ${labelKeyed("Cancel")}(b)) || list.find((b) => b.textContent.trim() === 'Cancel'))([...document.querySelectorAll('.rw-dialog button')])?.click(), 1`);

    // Another app opens Talk on a person: the dialog opens with them in it.
    opened = await open('talk', null, { inviteEmail: 'lee@example.com', inviteName: 'Lee Park' });
    const ojs = (code) => opened.webContents.executeJavaScript(code);
    const chip = await ok(until(() => ojs(`Boolean(document.querySelector('.tk-ie .tk-ie-chip[data-email="lee@example.com"]'))`), 'the dialog on Lee', 8000));
    const chipWords = chip ? await ojs(`document.querySelector('.tk-ie-chip[data-email="lee@example.com"]').textContent`) : '';
    check('talk: another app opening Talk on a person (Contacts\' Invite to Talk) opens Invite by email with that person in it', chip && /Lee Park/.test(chipWords), JSON.stringify({ chip, chipWords }));

    const complaints = [...await errorsIn(win), ...await errorsIn(opened)];
    check('talk: Connect reported no error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('talk: the Connect checks ran to the end', false, err?.message || String(err));
  } finally {
    const js = (code) => (win ? win.webContents.executeJavaScript(code).catch(() => {}) : Promise.resolve());
    try { opened?.close(); } catch { /* gone */ }
    await js(`window.rutbaOffice.talk.signOut().then(() => 1)`);
    if (cards.length) await js(`window.rutbaOffice.contacts.remove({ ids: ${JSON.stringify(cards.map((c) => c?.id).filter(Boolean))} }).then(() => 1)`);
    try { ben?.destroy(); } catch { /* gone */ }
    await intro.close().catch(() => {});
  }
}
