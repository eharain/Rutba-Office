// Profiles (D-2026-10-26-02; the owner, 2026-10-26: "we should offer
// multiple profiles in office so user can switch between them and manage
// them as well", side by side, with an optional password).
//
// In the run's own temporary base folder (never the person's): the
// launcher's badge and its menu; Manage Profiles makes a second profile; the
// second opens beside this one, as its own Rutba Office (verify-profiles-
// peer.js), its launcher wearing its own name and colour and not seeing a
// setting made here, nor this one its; a second launch of it goes to the one
// running; neither the first profile nor an open one can be removed, and the
// second, once closed, can, its folder gone. A password set here from the
// dialog seals this profile's files and comes off again. A profile with a
// password asks for it before it opens: a wrong one is refused, the right
// one opens it, and what it saved is there. Run alone with
// RUTBA_VERIFY_ONLY=profiles.

import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { app } from 'electron';
import { isSealed, readProfiles } from '@rutba/office-shell/profiles';

/** The environment a profile started by the checks runs in: a check run of its own, its windows off the desktop. */
function peerEnv(report, plan) {
  const env = { ...process.env };
  for (const k of Object.keys(env)) {
    if (/^RUTBA_(OFFICE_VERIFY|VERIFY_|SMOKE|SCREENS|PROFILE_PEER)/.test(k)) delete env[k];
  }
  delete env.ELECTRON_RUN_AS_NODE;
  return { ...env, RUTBA_PROFILE_PEER: report, RUTBA_PROFILE_PEER_PLAN: JSON.stringify(plan), RUTBA_WINDOW_DISPLAY: 'offscreen' };
}

const readReport = (file) => {
  try {
    return fs.readFileSync(file, 'utf8').split('\n').filter(Boolean).map((l) => JSON.parse(l));
  } catch {
    return [];
  }
};

/**
 * @param {object} h the harness: open, check, until, wait, capture
 */
export async function verifyProfiles(h) {
  const { open, check, wait, capture } = h;
  // The harness's until, answering with what was waited for rather than true.
  const until = async (fn, what, ms) => {
    let got = null;
    await h.until(async () => { got = await fn(); return got; }, what, ms);
    return got;
  };
  const base = app.getPath('userData');
  const scratch = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-profiles-check-'));
  const reports = [];
  const startPeer = (name, plan) => {
    const report = path.join(scratch, `${name}-${reports.length}.jsonl`);
    reports.push(report);
    globalThis.__rutbaCheckProfileEnv = () => peerEnv(report, plan);
    return report;
  };
  const event = (report, kind, test = () => true) => readReport(report).find((e) => e.event === kind && test(e)) || null;
  const win = await open('home');
  const js = (code) => win.webContents.executeJavaScript(code);
  const typeInto = (sel, text) => js(`(() => { const el = document.querySelector(${JSON.stringify(sel)}); if (!el) return false; Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(el, ${JSON.stringify(text)}); el.dispatchEvent(new Event('input', { bubbles: true })); return true; })()`);
  const click = (sel) => js(`(() => { const el = document.querySelector(${JSON.stringify(sel)}); if (!el) return false; el.click(); return true; })()`);
  const profiles = () => js(`window.rutbaOffice.profiles.list()`);
  // A picture once the dialog has finished coming in (an off-screen window shows its last painted frame).
  const shot = async (name) => {
    await wait(700);
    win.webContents.invalidate();
    await js(`new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(() => r(1))))`).catch(() => {});
    await wait(400);
    await capture?.(win, name);
  };
  let work = null;
  let secret = null;
  let contactId = null;
  try {
    // One profile: the launcher's badge says which, and its menu lists it.
    const first = await until(() => js(`(() => { const b = document.querySelector('.pf-badge'); return b ? { name: b.querySelector('.pf-badge-name')?.textContent, profile: b.dataset.profile } : null; })()`), 'the launcher\'s profile badge', 8000).catch(() => null);
    const current = await js(`window.rutbaOffice.profiles.current()`);
    check('profiles: with one profile, the launcher\'s title bar names it, Default, and the program runs in the base folder',
      first?.name === 'Default' && current.id === 'default' && current.count === 1 && path.resolve(base) === path.resolve(app.commandLine.getSwitchValue('user-data-dir')), JSON.stringify({ first, current }));
    await click('.pf-badge');
    const menu = await until(() => js(`(() => { const m = document.querySelector('.rw-menu'); return m ? [...m.querySelectorAll('button, .rw-menu-head')].map((b) => b.textContent.trim()) : null; })()`), 'the profile menu', 4000).catch(() => null);
    check('profiles: the badge opens the profile menu: the profiles, New Profile and Manage Profiles',
      Boolean(menu) && menu.includes('Profiles') && menu.some((m) => /^Default \(this window\)$/.test(m)) && menu.includes('New Profile…') && menu.includes('Manage Profiles…'), JSON.stringify(menu));

    // Manage Profiles: a second profile, named and coloured.
    await js(`[...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === 'Manage Profiles…')?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.pf-dialog .pf-new-name'))`), 'Manage Profiles', 6000).catch(() => {});
    await typeInto('.pf-new-name', 'Work');
    await click('.pf-new-row .pf-swatch[data-colour="#e08b2b"]');
    await wait(150);
    await click('.pf-make');
    const rows = await until(async () => { const n = await js(`document.querySelectorAll('.pf-dialog .pf-row').length`); return n === 2 ? n : null; }, 'the second profile in the list', 6000).catch(() => 0);
    const listed = readProfiles(base);
    work = listed.profiles.find((p) => p.name === 'Work') || null;
    const workFolder = work ? path.join(base, 'Profiles', work.id) : '';
    check('profiles: Manage Profiles makes a second profile with the name and colour chosen, in a folder of its own under the base folder\'s Profiles',
      rows === 2 && Boolean(work) && work.colour === '#e08b2b' && fs.existsSync(workFolder) && fs.existsSync(path.join(base, 'settings.json')), JSON.stringify({ rows, work }));
    const titled = await until(() => (/ · Default$/.test(win.getTitle()) ? win.getTitle() : null), 'the profile in the window title', 5000).catch(() => win.getTitle());
    check('profiles: with two profiles, each window\'s title names its profile', / · Default$/.test(titled), titled);

    // Side by side: the second profile opens as its own Rutba Office.
    await js(`window.rutbaOffice.store.set({ key: 'profiles.check', value: 'default' })`);
    const report = startPeer('work', { setting: ['profiles.check', 'work'], captureHome: process.env.RUTBA_VERIFY_CAPTURE ? path.join(process.env.RUTBA_VERIFY_CAPTURE, 'profiles-second.png') : null });
    await click(`.pf-row[data-profile="${work?.id}"] .pf-open`);
    const ready = await until(() => event(report, 'ready'), 'the second profile\'s launcher', 60000).catch(() => null);
    const started = event(report, 'start');
    check('profiles: Open starts the second profile beside this one, in its own folder, its launcher\'s badge in its own name and colour',
      Boolean(ready) && path.resolve(started?.userData || '') === path.resolve(workFolder) && ready.badge?.text === 'Work' && ready.badge?.colour === '#e08b2b' && ready.badge?.id === work.id && / · Work$/.test(ready.title || ''),
      JSON.stringify({ started, ready }));
    check('profiles: a setting made in one profile is not in the other, either way',
      ready?.setting === null && (await js(`window.rutbaOffice.store.get({ key: 'profiles.check', fallback: null })`)) === 'default'
        && (await until(() => { try { return JSON.parse(fs.readFileSync(path.join(workFolder, 'settings.json'), 'utf8'))['profiles.check'] === 'work' || null; } catch { return null; } }, 'the second profile\'s setting', 5000).catch(() => false)),
      `seen there: ${ready?.setting}`);
    const both = await until(async () => { const l = await profiles(); return l.profiles.find((p) => p.id === work.id)?.running ? l : null; }, 'the second profile running', 6000).catch(() => null);
    check('profiles: both run at once: this one says the other is open', Boolean(both), JSON.stringify(both?.profiles?.map((p) => [p.name, p.running])));
    await shot('profiles-manage.png');

    // A second launch of the open profile goes to it, not to a third copy.
    await click(`.pf-row[data-profile="${work?.id}"] .pf-open`);
    const handed = await until(() => event(report, 'second-instance'), 'the second launch handed over', 30000).catch(() => null);
    check('profiles: opening the second profile again goes to its running copy',
      Boolean(handed) && handed.pid === ready?.pid && readReport(report).filter((e) => e.event === 'start').length === 1, JSON.stringify(handed));

    // What cannot be removed, and why.
    const openOne = await js(`window.rutbaOffice.profiles.removal({ id: ${JSON.stringify(work.id)} })`);
    const firstOne = await js(`window.rutbaOffice.profiles.removal({ id: 'default' })`);
    check('profiles: neither an open profile nor the first one can be removed, each saying why',
      /Close its windows first/.test(openOne.refused || '') && /first profile keeps what every profile shares/.test(firstOne.refused || '')
        && !(await js(`Boolean(document.querySelector('.pf-row[data-profile="default"] .pf-remove-btn'))`)),
      `${openOne.refused} | ${firstOne.refused}`);

    // Closed, the second profile can go, and its folder with it.
    fs.writeFileSync(`${report}.quit`, '');
    const closed = await until(async () => { const l = await profiles(); return l.profiles.find((p) => p.id === work.id)?.running === false ? true : null; }, 'the second profile closed', 30000).catch(() => false);
    await click(`.pf-row[data-profile="${work.id}"] .pf-remove-btn`);
    const asked = await until(() => js(`document.querySelector('.pf-remove-dialog')?.textContent || null`), 'the question before removing', 5000).catch(() => '');
    await click('.pf-remove-go');
    const removed = await until(async () => ((await js(`document.querySelectorAll('.pf-dialog .pf-row').length`)) === 1 && !fs.existsSync(workFolder) ? true : null), 'the profile removed', 8000).catch(() => false);
    check('profiles: once closed, a profile is removed after the question, which says what goes, and its folder goes with it',
      closed && /Recycle Bin/.test(asked) && /stay where they are/.test(asked) && removed && readProfiles(base).profiles.length === 1, JSON.stringify({ closed, asked: asked.slice(0, 120), removed }));
    const untitled = await until(() => (!/ · /.test(win.getTitle()) ? win.getTitle() : null), 'the title without a profile', 5000).catch(() => win.getTitle());
    check('profiles: with one profile again, the window title names none', !/ · /.test(untitled), untitled);

    // A password set here, from the dialog: this profile's files sealed with it, then plain again once it is taken off.
    const contact = await js(`window.rutbaOffice.contacts.save({ contact: { name: { full: 'Sealed Person' }, emails: [{ value: 'sealed@example.com', type: 'home' }] } })`);
    contactId = contact?.id || null;
    await js(`window.rutbaOffice.app.addRecent({ path: ${JSON.stringify(path.join(scratch, 'Sealed report.docx'))}, app: 'word' })`).catch(() => {});
    const contactsFile = path.join(base, 'contacts.json');
    await until(() => { try { return fs.readFileSync(contactsFile, 'utf8').includes('Sealed Person') || null; } catch { return null; } }, 'the card written', 5000).catch(() => {});
    await click('.pf-set-password');
    await until(() => js(`Boolean(document.querySelector('.pf-password-dialog .pf-new'))`), 'Set a Password', 5000).catch(() => {});
    await typeInto('.pf-password-dialog .pf-new', 'check-pass-1');
    await typeInto('.pf-password-dialog .pf-again', 'check-pass-2');
    await click('.pf-password-go');
    const mismatch = await until(() => js(`document.querySelector('.pf-password-dialog .pf-error')?.textContent || null`), 'the two not the same', 4000).catch(() => '');
    const warned = await js(`document.querySelector('.pf-password-dialog .pf-warn')?.textContent || ''`);
    await shot('profiles-password.png');
    await typeInto('.pf-password-dialog .pf-again', 'check-pass-1');
    await click('.pf-password-go');
    const locked = await until(async () => ((await js(`window.rutbaOffice.profiles.current()`)).protected ? true : null), 'the password set', 10000).catch(() => false);
    const lock = readProfiles(base).profiles[0].lock;
    const listText = fs.readFileSync(path.join(base, 'profiles.json'), 'utf8');
    const sealedNow = await until(() => (isSealed(fs.readFileSync(contactsFile, 'utf8')) ? true : null), 'the card sealed', 5000).catch(() => false);
    check('profiles: Set a Password refuses two that differ, says a forgotten one cannot be recovered, then seals the address book and recent files, keeping a salt and a verifier, never the password',
      /not the same/.test(mismatch) && /cannot be recovered/.test(warned) && locked && Boolean(lock?.salt && lock?.verifier) && !listText.includes('check-pass-1')
        && sealedNow && !fs.readFileSync(contactsFile, 'utf8').includes('Sealed Person') && isSealed(fs.readFileSync(path.join(base, 'recent.json'), 'utf8')),
      JSON.stringify({ mismatch, locked, lock: Boolean(lock) }));
    const stillThere = await js(`window.rutbaOffice.contacts.list({ query: 'Sealed Person' }).then((l) => l.length)`);
    await until(() => js(`Boolean(document.querySelector('.pf-remove-password'))`), 'Remove Password', 4000).catch(() => {});
    await click('.pf-remove-password');
    await until(() => js(`Boolean(document.querySelector('.pf-password-dialog .pf-old'))`), 'Remove the Password', 5000).catch(() => {});
    await typeInto('.pf-password-dialog .pf-old', 'not-it');
    await click('.pf-password-go');
    const wrongOld = await until(() => js(`document.querySelector('.pf-password-dialog .pf-error')?.textContent || null`), 'the wrong current password', 6000).catch(() => '');
    await typeInto('.pf-password-dialog .pf-old', 'check-pass-1');
    await click('.pf-password-go');
    const unlocked = await until(async () => (!(await js(`window.rutbaOffice.profiles.current()`)).protected ? true : null), 'the password removed', 10000).catch(() => false);
    check('profiles: Remove Password asks for the current one, refuses a wrong one, then leaves the files plain and readable again',
      stillThere === 1 && /not right/.test(wrongOld) && unlocked && !readProfiles(base).profiles[0].lock && fs.readFileSync(contactsFile, 'utf8').includes('Sealed Person'),
      JSON.stringify({ stillThere, wrongOld, unlocked }));
    await click('.pf-dialog .rw-dialog-foot button');

    // A profile with a password asks for it before it opens.
    const made = await js(`window.rutbaOffice.profiles.make({ name: 'Private', colour: '#7b5cd6' })`);
    secret = made.made;
    const privateFolder = path.join(base, 'Profiles', secret);
    const setUp = startPeer('private', { secret: ['profiles.check', 'kept sealed'], contact: 'Private Person', setPassword: 'private-pass', quitAfter: true });
    await js(`window.rutbaOffice.profiles.open({ id: ${JSON.stringify(secret)} })`);
    const first2 = await until(() => event(setUp, 'leaving'), 'the private profile set up', 60000).catch(() => null);
    const setUpReady = event(setUp, 'ready');
    await until(async () => { const l = await profiles(); return l.profiles.find((p) => p.id === secret)?.running === false ? true : null; }, 'the private profile closed', 20000).catch(() => {});
    const privSecrets = (() => { try { return fs.readFileSync(path.join(privateFolder, 'secrets.json'), 'utf8'); } catch { return ''; } })();
    const privContacts = (() => { try { return fs.readFileSync(path.join(privateFolder, 'contacts.json'), 'utf8'); } catch { return ''; } })();
    check('profiles: a profile with a password set keeps its saved passwords and address book sealed on disk',
      Boolean(first2) && setUpReady?.passwordSet === true && setUpReady?.secretSaved === true && isSealed(privSecrets) && isSealed(privContacts) && !privContacts.includes('Private Person') && Boolean(readProfiles(base).profiles.find((p) => p.id === secret)?.lock),
      JSON.stringify({ setUpReady, sealed: [isSealed(privSecrets), isSealed(privContacts)] }));

    const lockShot = process.env.RUTBA_VERIFY_CAPTURE ? path.join(process.env.RUTBA_VERIFY_CAPTURE, 'profiles-lock.png') : null;
    const again = startPeer('private', { passwords: ['wrong-guess', 'private-pass'], readSecret: ['profiles.check', 'kept sealed'], findContact: 'Private Person', captureLock: lockShot, quitAfter: true });
    await js(`window.rutbaOffice.profiles.open({ id: ${JSON.stringify(secret)} })`);
    const opened = await until(() => event(again, 'ready'), 'the private profile opened', 60000).catch(() => null);
    const tries = readReport(again).filter((e) => e.event === 'password');
    check('profiles: the protected profile asks before it opens, refuses a wrong password, and opens with the right one',
      tries.length === 2 && tries[0].outcome === 'refused' && /not right/.test(tries[0].error) && tries[1].outcome === 'opened' && Boolean(opened), JSON.stringify(tries));
    check('profiles: opened with its password, the profile reads what it sealed: its saved password and its address book',
      opened?.secretMatches === true && opened?.contactFound === true && opened?.badge?.text === 'Private', JSON.stringify(opened));
    await until(() => event(again, 'leaving'), 'the private profile closed', 20000).catch(() => {});
    await until(async () => { const l = await profiles(); return l.profiles.find((p) => p.id === secret)?.running === false ? true : null; }, 'the private profile closed', 20000).catch(() => {});
  } finally {
    // Left as it was found: Default alone, with no password and no card of the check's.
    globalThis.__rutbaCheckProfileEnv = undefined;
    for (const r of reports) {
      try { fs.writeFileSync(`${r}.quit`, ''); } catch { /* gone */ }
    }
    if (contactId) await js(`window.rutbaOffice.contacts.remove({ id: ${JSON.stringify(contactId)} })`).catch(() => {});
    for (const id of [work?.id, secret].filter(Boolean)) {
      for (let i = 0; i < 40; i++) {
        const l = await profiles().catch(() => null);
        const p = l?.profiles?.find((x) => x.id === id);
        if (!p) break;
        if (!p.running) { await js(`window.rutbaOffice.profiles.remove({ id: ${JSON.stringify(id)} })`).catch(() => {}); break; }
        await wait(500);
      }
    }
    if ((await js(`window.rutbaOffice.profiles.current()`).catch(() => ({}))).protected) {
      await js(`window.rutbaOffice.profiles.removePassword({ old: 'check-pass-1' })`).catch(() => {});
    }
    await js(`document.querySelectorAll('.rw-dialog-foot button').forEach((b) => { if (/Close|Cancel/.test(b.textContent)) b.click(); }), 1`).catch(() => {});
    fs.rmSync(scratch, { recursive: true, force: true });
  }
}
