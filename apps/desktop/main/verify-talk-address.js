// Talk, My address (Office 2.25, D-2026-11-11-02), on two stand-in
// introducers. One keeps only the sign-in address, as the live introducer
// does: My address shows that one as the one shown and the one that finds
// you, says more come with a later Talk server, and offers nothing it cannot
// do. The other keeps several (the protocol the client speaks for phase 2):
// another address is proved with the code emailed to it, can be the one
// shown, finds you when someone asks by it, and stops finding you when told.
// Run alone with RUTBA_VERIFY_ONLY=talk-address.

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

export async function verifyTalkAddress({ open, check, until, wait, errorsIn }) {
  const ok = (p) => p.then(() => true, () => false);
  let win = null;
  let ben = null;
  let intro = null;
  try {
    win = await open('talk');
    const js = (code) => win.webContents.executeJavaScript(code);
    const set = (sel, value) => js(`(() => { const n = document.querySelector(${JSON.stringify(sel)}); Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(n, ${JSON.stringify(value)}); n.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
    const signIn = async (port) => {
      await until(() => js(`Boolean(document.querySelector('.tk-signin'))`), 'the sign-in', 8000);
      await js(`window.rutbaOffice.talk.signIn({ token: 'amira', server: ${JSON.stringify(`ws://127.0.0.1:${port}/ws`)} }).then(() => 1)`);
      await until(() => js(`Boolean(document.querySelector('.tk-addr-btn'))`), 'signed in', 8000);
    };
    const row = (address) => `document.querySelector('.tk-addr-row[data-address="${address}"]')`;

    // As the live introducer is: the sign-in address alone.
    intro = standInIntroducer({});
    await signIn(await intro.listen());
    await js(`document.querySelector('.tk-addr-btn').click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.tk-addr-list'))`), 'My address', 8000);
    const alone = await js(`(() => { const l = document.querySelector('.tk-addr-list'); return { supported: l.dataset.supported, rows: [...l.querySelectorAll('.tk-addr-row')].map((r) => ({ address: r.dataset.address, shown: r.dataset.shown, findable: r.dataset.findable, locked: [...r.querySelectorAll('input')].every((i) => i.disabled) })), later: document.querySelector('.tk-addr-later')?.textContent || '', add: Boolean(document.querySelector('.tk-addr-new')), sender: document.querySelector('.tk-addr-sender')?.textContent || '' }; })()`);
    check('talk: My address on a Talk server that keeps only the sign-in address shows that one as shown and finding you, says more come with a later Talk server, offers no way to add one, and says where invitations by email go from',
      alone.supported === 'no' && alone.rows.length === 1 && alone.rows[0].address === 'amira@example.com' && alone.rows[0].shown === 'yes' && alone.rows[0].findable === 'yes' && alone.rows[0].locked
        && /later Talk server/.test(alone.later) && !alone.add && /invitations/i.test(alone.sender),
      JSON.stringify(alone));
    await js(`((list) => list.find((b) => ${labelKeyed("Done")}(b)) || list.find((b) => b.textContent.trim() === 'Done'))([...document.querySelectorAll('.rw-dialog button')])?.click(), 1`);
    await js(`window.rutbaOffice.talk.signOut().then(() => 1)`);
    await intro.close();

    // A Talk server that keeps several.
    intro = standInIntroducer({ page: PAGE, addresses: true });
    const port = await intro.listen();
    await signIn(port);
    await js(`document.querySelector('.tk-addr-btn').click(), 1`);
    await until(() => js(`document.querySelector('.tk-addr-list')?.dataset.supported === 'yes'`), 'My address, more kept', 8000);
    await set('.tk-addr-new', 'Amira@Work.example');
    await js(`document.querySelector('.tk-addr-send').click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.tk-addr-code'))`), 'asked for the code', 6000);
    const code = intro.codes()['amira@work.example'];
    await set('.tk-addr-code', '000000');
    await js(`document.querySelector('.tk-addr-confirm').click(), 1`);
    await wait(300);
    const wrongKept = await js(`Boolean(document.querySelector('.tk-addr-code')) && !${row('amira@work.example')}`);
    await set('.tk-addr-code', code);
    await js(`document.querySelector('.tk-addr-confirm').click(), 1`);
    const proved = await ok(until(() => js(`${row('amira@work.example')}?.dataset.findable === 'yes' && !document.querySelector('.tk-addr-code')`), 'the address proved', 6000));
    check('talk: another address of one\'s own is proved with the code emailed to it (a wrong code is refused and can be typed again), and then finds you',
      Boolean(code) && wrongKept && proved, JSON.stringify({ code: Boolean(code), wrongKept, proved }));

    // Shown: Talk shows it as hers; Ben, asking by it, reaches her.
    await js(`${row('amira@work.example')}.querySelector('.tk-addr-shown input').click(), 1`);
    const shown = await ok(until(() => js(`${row('amira@work.example')}?.dataset.shown === 'yes' && ${row('amira@example.com')}?.dataset.shown === 'no'`), 'shown', 5000));
    const me = await js(`window.rutbaOffice.talk.state().then((s) => s.me.email)`);
    ben = new BrowserWindow({ show: false, webPreferences: { contextIsolation: true, nodeIntegration: false, partition: 'talk-address-ben' } });
    await ben.loadURL(`http://127.0.0.1:${port}/ben`);
    const bjs = (c) => ben.webContents.executeJavaScript(c);
    await until(() => bjs(`window.peer.log.some((m) => m.type === 'welcome')`), 'Ben signed in', 6000);
    await bjs(`window.peer.send({ type: 'contact.request', email: 'amira@work.example', note: 'By your work address' }), 1`);
    const found = await ok(until(() => js(`/By your work address/.test(document.querySelector('.tk-request')?.textContent || '')`), 'Ben found her', 6000));
    check('talk: an address proved can be the one shown, and someone asking by it reaches the person', shown && me === 'amira@work.example' && found, JSON.stringify({ shown, me, found }));

    // Kept from finding her: Ben's next ask by it reaches nobody. Then taken off; the sign-in address stays and is shown again.
    await js(`${row('amira@work.example')}.querySelector('.tk-addr-find input').click(), 1`);
    const unfound = await ok(until(() => js(`${row('amira@work.example')}?.dataset.findable === 'no'`), 'kept from finding her', 5000));
    const before = await js(`document.querySelectorAll('.tk-request').length`);
    await bjs(`window.peer.send({ type: 'contact.request', email: 'amira@work.example', note: 'Again' }), 1`);
    await wait(600);
    const after = await js(`document.querySelectorAll('.tk-request').length`);
    const signInLocked = await js(`!${row('amira@example.com')}.querySelector('.tk-addr-remove')`);
    await js(`${row('amira@work.example')}.querySelector('.tk-addr-remove').click(), 1`);
    const removed = await ok(until(() => js(`!${row('amira@work.example')} && ${row('amira@example.com')}?.dataset.shown === 'yes'`), 'taken off', 5000));
    check('talk: an address kept from finding you finds you no longer; it can be taken off, and the sign-in address, which cannot, is shown again',
      unfound && before === after && signInLocked && removed, JSON.stringify({ unfound, before, after, signInLocked, removed }));
    await js(`((list) => list.find((b) => ${labelKeyed("Done")}(b)) || list.find((b) => b.textContent.trim() === 'Done'))([...document.querySelectorAll('.rw-dialog button')])?.click(), 1`);

    // The wrong code typed on purpose above is the one refusal expected.
    const complaints = (await errorsIn(win)).filter((c) => !/^That code is not right\. 4 tries left\.$/.test(String(c).trim()));
    check('talk: My address reported no error but the wrong code it was given', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('talk: the My address checks ran to the end', false, err?.message || String(err));
  } finally {
    if (win) await win.webContents.executeJavaScript(`window.rutbaOffice.talk.signOut().then(() => 1)`).catch(() => {});
    try { ben?.destroy(); } catch { /* gone */ }
    await intro?.close().catch(() => {});
  }
}
