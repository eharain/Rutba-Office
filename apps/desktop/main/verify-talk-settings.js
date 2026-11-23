// Talk's settings, seen in its window (Office 2.27.0): the two settings as a
// person ticks them, and the people blocked, each by name with Unblock.
//
// Talk is the stand-in introducer of verify-talk.js. It takes someone off the
// contacts when they are blocked; letting them back (contact.unblock) is this
// check's own piece of the protocol, handed in as `more`. A check run makes
// no sound, so what the sound setting does to a ring is the unit tests'
// (tests/talk-ring.test.js); here it is the setting that is seen kept.

import { BrowserWindow } from 'electron';
import { standInIntroducer } from './verify-talk.js';

const windowsOf = (app) => BrowserWindow.getAllWindows().filter((w) => {
  if (w.isDestroyed()) return false;
  // A window with no address yet (one Office is still opening, or a hidden one of its own) is no app's window.
  try { return new URL(w.webContents.getURL()).searchParams.get('app') === app; } catch { return false; }
});

export async function verifyTalkSettings(h) {
  const { open, check, until, wait, errorsIn } = h;
  const ok = (p) => p.then(() => true).catch(() => false);
  const unblocked = []; // every contact.unblock as it arrived: { by, id }
  const intro = standInIntroducer({
    people: { chen: { id: 'u-chen', name: 'Chen Wei', email: 'chen@example.com' } },
    contacts: [['u-amira', 'u-ben'], ['u-amira', 'u-chen']],
    more: (msg, me, t) => {
      if (msg.type !== 'contact.unblock') return false;
      const id = String(msg.contactId || '');
      unblocked.push({ by: me.id, id });
      t.to(me.id, { type: 'contact.unblocked', id });
      return true;
    },
  });
  let spare = null;
  let js = null;
  let was = null;
  try {
    const port = await intro.listen();
    const url = `ws://127.0.0.1:${port}/ws`;

    // Talk windows left by the checks before this one are done with; another window keeps Office open meanwhile.
    const left = windowsOf('talk');
    spare = await open('word');
    for (const w of left) w.destroy();

    const win = await open('talk');
    js = (code) => win.webContents.executeJavaScript(code);
    const q = (sel) => `document.querySelector(${JSON.stringify(sel)})`;
    const click = (sel) => js(`(${q(sel)}?.click(), 1)`);
    const seen = (sel, what, ms = 8000) => ok(until(() => js(`Boolean(${q(sel)})`), what, ms));
    const gone = (sel, what, ms = 8000) => ok(until(() => js(`!${q(sel)}`), what, ms));
    const type = (sel, value) => js(`(() => { const n = ${q(sel)}; Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(n, ${JSON.stringify(value)}); n.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
    const talk = () => js(`window.rutbaOffice.talk.state()`);
    await js(`window.rutbaOffice.talk.signOut().then(() => 1, () => 1)`);
    await until(() => js(`Boolean(${q('.tk-signin')})`), 'the sign-in', 8000).catch(() => {});
    await js(`${q('.tk-test')}?.setAttribute('open', ''), 1`);
    await type('.tk-server', url);
    await type('.tk-token', 'amira');
    await click('.tk-connect');
    const signedIn = await ok(until(() => js(`Boolean(${q('.tk-contact[data-contact="u-chen"]')})`), 'signed in, Chen among the contacts', 10000));
    const before = await talk();
    was = { startWithOffice: before.startWithOffice !== false, ring: before.ring !== false };
    // The check starts from both settings on, as a new copy has them.
    await js(`window.rutbaOffice.talk.setOptions({ startWithOffice: true, ring: true }).then(() => 1)`);

    // Settings: the two boxes, each ticked as the setting stands, and nobody blocked.
    const boxes = () => js(`({ start: ${q('.tk-settings .tk-start-with-office')}?.checked, ring: ${q('.tk-settings .tk-ring-sound')}?.checked })`);
    const openSettings = async () => { await click('.tk-settings-btn'); return seen('.tk-settings', 'Talk\'s settings'); };
    const done = async () => { await click('.tk-settings .rw-dialog-foot button'); return gone('.tk-settings', 'the settings closed'); };
    const shown = await openSettings();
    const asNew = await boxes();
    const nobody = await js(`Boolean(${q('.tk-settings .tk-none')}) && !${q('.tk-settings .tk-blocked-list')}`);
    const noServer = await js(`!${q('.tk-settings input[type="text"], .tk-settings input[type="url"], .tk-settings input:not([type])')}`);
    // Each box unticked is the setting changed at once, and kept: closed and opened again the page shows it.
    await click('.tk-settings .tk-ring-sound');
    const quiet = await ok(until(() => talk().then((s) => s.ring === false && s.startWithOffice !== false), 'sounds off', 5000));
    await click('.tk-settings .tk-start-with-office');
    const notAtStart = await ok(until(() => talk().then((s) => s.ring === false && s.startWithOffice === false), 'not connecting at the start', 5000));
    const shut = await done();
    await openSettings();
    const keptOff = await boxes();
    await click('.tk-settings .tk-ring-sound');
    await click('.tk-settings .tk-start-with-office');
    const backOn = await ok(until(() => talk().then((s) => s.ring !== false && s.startWithOffice !== false), 'both on again', 5000));
    const ticked = await boxes();
    await done();
    check('talk settings: Settings shows the two settings ticked as they stand and nobody blocked, with no place to type a server; each box changes its setting at once, and the page shows it the same when opened again',
      signedIn && shown && asNew.start === true && asNew.ring === true && nobody && noServer && quiet && notAtStart && shut && keptOff.start === false && keptOff.ring === false && backOn && ticked.start === true && ticked.ring === true,
      JSON.stringify({ signedIn, shown, asNew, nobody, noServer, quiet, notAtStart, shut, keptOff, backOn, ticked }));

    // Chen is blocked from the contact's menu, asked first: off the contacts, and in Settings by name and address.
    globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'message', answer: { response: 0 } });
    await js(`${q('.tk-contact[data-contact="u-chen"]')}.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 120, clientY: 160 })), 1`);
    await seen('.rw-menu [data-key="block-2"]', 'Block in the menu', 5000);
    const menu = await js(`(() => { const b = ${q('.rw-menu [data-key="block-2"]')}; if (!b) return [...document.querySelectorAll('.rw-menu button')].map((n) => n.dataset.key || '').join(' '); b.click(); return 'pressed'; })()`);
    const offContacts = await gone('.tk-contact[data-contact="u-chen"]', 'Chen off the contacts');
    await openSettings();
    const ROW = '.tk-settings .tk-blocked-list .tk-member[data-blocked="u-chen"]';
    const listed = await seen(ROW, 'Chen among the people blocked');
    const row = await js(`(() => { const n = ${q(ROW)}; return n ? { says: n.querySelector('.tk-who')?.textContent || '', unblock: n.querySelector('.tk-unblock')?.disabled === false, rows: document.querySelectorAll('.tk-settings .tk-blocked-list .tk-member').length } : null; })()`);
    check('talk settings: someone blocked from the contact\'s menu is off the contacts and listed in Settings by name and address, with Unblock',
      menu === 'pressed' && offContacts && listed && /Chen Wei/.test(row?.says || '') && /chen@example\.com/.test(row?.says || '') && row.unblock && row.rows === 1,
      JSON.stringify({ menu, offContacts, listed, row }));

    // Unblock: the introducer is asked, the person is off the list, and the page says nobody is blocked. They are
    // not a contact again by that alone.
    await click(`${ROW} .tk-unblock`);
    const asked = await ok(until(() => unblocked.some((u) => u.by === 'u-amira' && u.id === 'u-chen') || null, 'the introducer asked', 6000));
    const offList = await gone(ROW, 'Chen off the list');
    const empty = await ok(until(() => js(`Boolean(${q('.tk-settings .tk-none')})`), 'nobody blocked', 5000));
    const state = await talk();
    await done();
    const notBack = await js(`!${q('.tk-contact[data-contact="u-chen"]')}`);
    check('talk settings: Unblock asks the introducer, takes the person off the list and leaves nobody blocked; they are not put back among the contacts by that alone',
      asked && offList && empty && !(state.blocked || []).includes('u-chen') && unblocked.length === 1 && notBack, JSON.stringify({ asked, offList, empty, blocked: state.blocked, asks: unblocked.length, notBack }));

    await wait(300);
    const complaints = await errorsIn(win);
    check('talk settings: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('talk settings: the check ran to the end', false, err?.message || String(err));
  } finally {
    if (js) {
      if (was) await js(`window.rutbaOffice.talk.setOptions(${JSON.stringify(was)}).then(() => 1)`).catch(() => {});
      await js(`window.rutbaOffice.talk.signOut().then(() => 1, () => 1)`).catch(() => {});
    }
    try { spare?.destroy(); } catch { /* gone */ }
    await intro.close().catch(() => {});
  }
}
