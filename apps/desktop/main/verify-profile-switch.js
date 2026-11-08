// Switching profile keeps you in the same app (D-2026-11-06-01; the owner, 6
// Nov: "switching profile on a window of an app should open the same app
// window of other profile. opening all app view should be an addition menu
// item there."). In the run's own temporary base folder, from a Mail window
// here: the other profile's Mail when that profile is not running and when
// it is, a second switch reusing the window, All apps opening its Home, Home's
// own menu with no extra items, an app turned off there, and a profile with a
// password, asked for both before it is open and while it is still asking.
// Run alone with RUTBA_VERIFY_ONLY=profile-switch.

import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

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

export async function verifyProfileSwitch(h) {
  const { open, check, wait, capture, errorsIn } = h;
  const until = async (fn, what, ms) => {
    let got = null;
    await h.until(async () => { got = await fn(); return got; }, what, ms);
    return got;
  };
  const scratch = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-switch-check-'));
  const shots = process.env.RUTBA_VERIFY_CAPTURE || null;
  let peers = 0;
  const startPeer = (name, plan) => {
    const report = path.join(scratch, `${name}-${peers++}.jsonl`);
    globalThis.__rutbaCheckProfileEnv = () => peerEnv(report, { watch: true, ...plan });
    return report;
  };
  const events = (report, kind) => readReport(report).filter((e) => e.event === kind);
  const latest = (report) => events(report, 'windows').at(-1) || null;
  const appsIn = (w) => (w?.windows || []).map((x) => x.app).sort();
  const quitPeer = async (report, id, js) => {
    fs.writeFileSync(`${report}.quit`, '');
    await until(async () => ((await js(`window.rutbaOffice.profiles.list()`)).profiles.find((p) => p.id === id)?.running === false ? true : null), 'the other profile closed', 30000).catch(() => {});
  };
  const home = await open('home');
  const hjs = (code) => home.webContents.executeJavaScript(code);
  let mail = null;
  let work = null;
  let secret = null;
  let reports = [];
  const tracked = (r) => { reports.push(r); return r; };
  try {
    const made = await hjs(`window.rutbaOffice.profiles.make({ name: 'Work', colour: '#0f7d6b' })`);
    work = made.made;
    mail = await open('mail');
    const mjs = (code) => mail.webContents.executeJavaScript(code);
    await until(() => mjs(`Boolean(document.querySelector('.pf-badge'))`), 'the badge in Mail', 10000);

    // The menu, opened afresh and read; a row chosen by its name.
    const openMenu = async (js) => {
      await js(`document.querySelector('.rw-menu') ? window.dispatchEvent(new MouseEvent('mousedown')) : 0, 1`);
      await wait(150);
      await js(`document.querySelector('.pf-badge').click(), 1`);
      return until(() => js(`(() => { const m = document.querySelector('.rw-menu'); if (!m) return null; return [...m.querySelectorAll(':scope > button')].map((b) => ({ label: b.querySelector('.rw-menu-label').textContent.trim(), hint: b.querySelector('.rw-menu-hint')?.textContent.trim() || null, mark: Boolean(b.querySelector('.pf-mark')) })); })()`), 'the profile menu', 5000);
    };
    const choose = (js, label) => js(`[...document.querySelectorAll('.rw-menu > button')].find((b) => b.querySelector('.rw-menu-label').textContent.trim() === ${JSON.stringify(label)})?.click(), 1`);
    const row = (rows, label) => rows.find((r) => r.label === label) || {};

    // Not running: its Mail opens, on the start screen, and no Home with it.
    const rows1 = await openMenu(mjs);
    check('profile switch: in Mail, the menu has each profile saying what it will do, then All apps in Work, then New Profile and Manage Profiles',
      row(rows1, 'Default').hint === 'This window' && row(rows1, 'Work').hint === 'Mail' && row(rows1, 'Work').mark
        && JSON.stringify(rows1.map((r) => r.label)) === JSON.stringify(['Default', 'Work', 'All apps in Work', 'New Profile…', 'Manage Profiles…']), JSON.stringify(rows1));
    if (shots) { await wait(500); mail.webContents.invalidate(); await wait(400); await capture?.(mail, 'profile-menu-mail.png'); }
    const report = tracked(startPeer('work', { appsOff: ['video'] }));
    await choose(mjs, 'Work');
    const first = await until(() => { const w = latest(report); return appsIn(w).includes('mail') ? w : null; }, 'Mail in the other profile', 60000).catch(() => null);
    check('profile switch: when the other profile is not running, its Mail opens on the start screen, and this window stays open',
      JSON.stringify(appsIn(first)) === '["mail"]' && /[?&]start=1/.test(first?.windows?.[0]?.url || '') && !mail.isDestroyed(), JSON.stringify(first));

    // Running, with Mail open: the row says so, and that window comes forward.
    const rows2 = await until(async () => { const r = await openMenu(mjs); return row(r, 'Work').hint === 'Mail, open' ? r : null; }, 'the row saying Mail is open', 10000).catch(() => null);
    check('profile switch: the row now says Mail is open there', row(rows2 || [], 'Work').hint === 'Mail, open', JSON.stringify(rows2));
    if (shots) { await wait(500); mail.webContents.invalidate(); await wait(400); await capture?.(mail, 'profile-menu-mail-open.png'); }
    const focusedBefore = events(report, 'focused').length;
    await choose(mjs, 'Work');
    const again = await until(() => events(report, 'second-instance').find((e) => e.app === '--app=mail') || null, 'the second switch handed over', 30000).catch(() => null);
    const forward = await until(() => events(report, 'focused').slice(focusedBefore).find((e) => e.app === 'mail') || null, 'its Mail brought forward', 8000).catch(() => null);
    await wait(700);
    check('profile switch: a second switch goes to the same running copy and brings its Mail forward, opening no other window',
      Boolean(again) && Boolean(forward) && JSON.stringify(appsIn(latest(report))) === '["mail"]' && events(report, 'start').length === 1, JSON.stringify({ again, forward, now: latest(report) }));

    // All apps: its Home, beside its Mail.
    await openMenu(mjs);
    await choose(mjs, 'All apps in Work');
    const allApps = await until(() => { const w = latest(report); return appsIn(w).includes('home') ? w : null; }, 'its Home', 15000).catch(() => null);
    check('profile switch: All apps in Work opens its Home, and its Mail stays', JSON.stringify(appsIn(allApps)) === '["home","mail"]', JSON.stringify(allApps));

    // Home's own menu: a profile opens Home, and there are no further items.
    await until(() => hjs(`Boolean(document.querySelector('.pf-badge'))`), 'the badge in Home', 8000);
    const homeRows = await until(async () => { const r = await openMenu(hjs); return row(r, 'Work').hint === 'Home, open' ? r : null; }, 'the row in Home', 10000).catch(() => null);
    check('profile switch: in Home a profile opens Home, and All apps is not offered',
      row(homeRows || [], 'Work').hint === 'Home, open' && JSON.stringify((homeRows || []).map((r) => r.label)) === JSON.stringify(['Default', 'Work', 'New Profile…', 'Manage Profiles…']), JSON.stringify(homeRows));
    const homeFocused = events(report, 'focused').length;
    await choose(hjs, 'Work');
    const homeForward = await until(() => events(report, 'focused').slice(homeFocused).find((e) => e.app === 'home') || null, 'its Home brought forward', 15000).catch(() => null);
    await wait(500);
    check('profile switch: from Home, the other profile\'s Home comes forward, with no second one', Boolean(homeForward) && JSON.stringify(appsIn(latest(report))) === '["home","mail"]', JSON.stringify({ homeForward, now: latest(report) }));

    // An app turned off there, with its Home open: that Home says so. (Started with appsOff: a check run otherwise has every app on.)
    const video = await open('video');
    const vjs = (code) => video.webContents.executeJavaScript(code);
    await until(() => vjs(`Boolean(document.querySelector('.pf-badge'))`), 'the badge in Video', 10000);
    await openMenu(vjs);
    await choose(vjs, 'Work');
    const offRunning = await until(() => { const w = latest(report); return w?.notice ? w : null; }, 'its Home saying Video is off', 20000).catch(() => null);
    check('profile switch: an app turned off in a running profile opens no window of it; its Home comes forward saying so',
      offRunning?.notice === 'Video is off in Work' && JSON.stringify(appsIn(offRunning)) === '["home","mail"]', JSON.stringify(offRunning));
    await quitPeer(report, work, hjs);

    // And not running: its Home opens, saying so.
    const coldShot = shots ? path.join(shots, 'profile-app-off.png') : null;
    const coldReport = tracked(startPeer('work', { captureOff: coldShot, appsOff: ['video'] }));
    await openMenu(vjs);
    await choose(vjs, 'Work');
    const offCold = await until(() => { const w = latest(coldReport); return w?.notice ? w : null; }, 'its Home saying Video is off', 60000).catch(() => null);
    check('profile switch: an app turned off in a profile not running opens its Home, saying so, with Choose apps there',
      offCold?.notice === 'Video is off in Work' && JSON.stringify(appsIn(offCold)) === '["home"]' && /[?&]off=video/.test(offCold?.windows?.[0]?.url || ''), JSON.stringify(offCold));
    await wait(1200);
    await quitPeer(coldReport, work, hjs);
    video.destroy();

    // A profile with a password: the password window first, then the app.
    const privMade = await hjs(`window.rutbaOffice.profiles.make({ name: 'Private', colour: '#7b5cd6' })`);
    secret = privMade.made;
    const setUp = tracked(path.join(scratch, 'private-setup.jsonl'));
    globalThis.__rutbaCheckProfileEnv = () => peerEnv(setUp, { setPassword: 'private-pass', quitAfter: true });
    await hjs(`window.rutbaOffice.profiles.open({ id: ${JSON.stringify(secret)} })`);
    await until(() => readReport(setUp).find((e) => e.event === 'leaving') || null, 'the private profile set up', 60000).catch(() => {});
    await until(async () => ((await hjs(`window.rutbaOffice.profiles.list()`)).profiles.find((p) => p.id === secret)?.running === false ? true : null), 'the private profile closed', 20000).catch(() => {});

    const coldLocked = tracked(startPeer('private', { passwords: ['private-pass'] }));
    await openMenu(mjs);
    await choose(mjs, 'Private');
    const opened = await until(() => { const w = latest(coldLocked); return appsIn(w).includes('mail') ? w : null; }, 'Mail after the password', 60000).catch(() => null);
    const told = events(coldLocked, 'password');
    check('profile switch: a profile with a password asks first, then opens the same app, and no Home',
      told.length === 1 && told[0].outcome === 'opened' && JSON.stringify(appsIn(opened)) === '["mail"]', JSON.stringify({ told, opened }));
    await quitPeer(coldLocked, secret, hjs);

    // Asked while it is still asking for its password: the app is kept.
    const waiting = tracked(startPeer('private', { passwords: ['private-pass'], passwordWait: true }));
    await hjs(`window.rutbaOffice.profiles.open({ id: ${JSON.stringify(secret)} })`);
    await until(async () => ((await hjs(`window.rutbaOffice.profiles.list()`)).profiles.find((p) => p.id === secret)?.running ? true : null), 'the private profile asking', 30000).catch(() => {});
    await wait(1500);
    await openMenu(mjs);
    await choose(mjs, 'Private');
    const held = await until(() => events(waiting, 'second-instance').find((e) => e.app === '--app=mail' && e.locked) || null, 'the switch while it asks', 30000).catch(() => null);
    fs.writeFileSync(`${waiting}.go`, '');
    const kept = await until(() => { const w = latest(waiting); return appsIn(w).includes('mail') ? w : null; }, 'Mail after the password, asked for while it asked', 60000).catch(() => null);
    check('profile switch: asked while the password window is up, the app is kept and opens after the password, and no Home with it',
      Boolean(held) && JSON.stringify(appsIn(kept)) === '["mail"]', JSON.stringify({ held, kept }));
    await quitPeer(waiting, secret, hjs);

    check('profile switch: this profile\'s Mail stayed open through all of it', !mail.isDestroyed() && /app=mail/.test(mail.webContents.getURL()), mail.webContents.getURL());
    const complaints = [...await errorsIn(mail), ...await errorsIn(home)];
    check('profile switch: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('profile switch: the check ran to the end', false, err?.message || String(err));
  } finally {
    globalThis.__rutbaCheckProfileEnv = undefined;
    for (const r of reports) {
      try { fs.writeFileSync(`${r}.quit`, ''); } catch { /* gone */ }
    }
    for (const id of [work, secret].filter(Boolean)) {
      for (let i = 0; i < 40; i++) {
        const l = await hjs(`window.rutbaOffice.profiles.list()`).catch(() => null);
        const p = l?.profiles?.find((x) => x.id === id);
        if (!p) break;
        if (!p.running) { await hjs(`window.rutbaOffice.profiles.remove({ id: ${JSON.stringify(id)} })`).catch(() => {}); break; }
        await wait(500);
      }
    }
    if (mail && !mail.isDestroyed()) mail.destroy();
    fs.rmSync(scratch, { recursive: true, force: true });
  }
}
