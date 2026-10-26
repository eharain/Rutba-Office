// The second profile the profile checks start beside the first
// (verify-profiles.js): a Rutba Office of its own, in a profile of its own,
// that writes what it sees to a report file the checks read. Started with
// RUTBA_PROFILE_PEER=<report> and a plan (RUTBA_PROFILE_PEER_PLAN, JSON):
//
//   passwords   typed into its own password window in turn, each outcome reported
//   setting     [key, value] written to its settings once its launcher is up
//   secret      [key, value] saved among its passwords
//   readSecret  [key, value] whether the saved password reads back as that value
//   contact     a card's name, added to its address book
//   findContact a card's name, looked for in its address book
//   setPassword a password set on itself, from its own window
//   captureLock where to write a picture of its password window
//   captureHome where to write a picture of its launcher
//   quitAfter   leave once all of it is done
//
// It leaves at once when a file `<report>.quit` appears, and after two
// minutes whatever happens, so no check leaves it behind.

import fs from 'node:fs';
import { app, BrowserWindow } from 'electron';

const wait = (ms) => new Promise((r) => setTimeout(r, ms));

export function startPeer({ report, profile }) {
  let plan = {};
  try {
    plan = JSON.parse(process.env.RUTBA_PROFILE_PEER_PLAN || '{}');
  } catch {
    plan = {};
  }
  const log = (entry) => {
    try {
      fs.appendFileSync(report, `${JSON.stringify({ at: Date.now(), pid: process.pid, ...entry })}\n`);
    } catch {
      /* the checks see nothing, and say so */
    }
  };
  app.on('second-instance', (_e, argv) => log({ event: 'second-instance', profile: argv.find((a) => String(a).startsWith('--profile')) || null }));
  const leave = (why) => {
    log({ event: 'leaving', why });
    app.quit();
  };
  setTimeout(() => leave('two minutes'), 120000).unref?.();
  const quitFile = `${report}.quit`;
  const poll = setInterval(() => {
    if (fs.existsSync(quitFile)) {
      clearInterval(poll);
      leave('asked');
    }
  }, 300);
  poll.unref?.();

  // Its own password window, answered as a person would: typed, Open pressed, the window read.
  const passwords = Array.isArray(plan.passwords) ? plan.passwords.slice() : [];
  if (profile.lock && passwords.length) {
    (async () => {
      let win = null;
      for (let i = 0; i < 300 && !win; i++) {
        win = BrowserWindow.getAllWindows().find((w) => !w.isDestroyed() && w.webContents.getURL().startsWith('data:text/html'));
        if (!win) await wait(100);
      }
      if (!win) return log({ event: 'password', outcome: 'no window' });
      const js = (code) => win.webContents.executeJavaScript(code);
      for (let i = 0; i < 100; i++) {
        if (await js(`Boolean(document.getElementById('pw') && window.rutbaLock)`).catch(() => false)) break;
        await wait(100);
      }
      if (plan.captureLock) {
        await wait(400);
        try {
          win.webContents.invalidate();
          await wait(300);
          fs.writeFileSync(plan.captureLock, (await win.webContents.capturePage()).toPNG());
        } catch {
          /* a picture is a nicety */
        }
      }
      for (const [i, password] of passwords.entries()) {
        await js(`(() => { const pw = document.getElementById('pw'); pw.value = ${JSON.stringify(password)}; document.getElementById('err').textContent = ''; document.getElementById('f').requestSubmit(); return 1; })()`).catch(() => {});
        let outcome = null;
        let error = '';
        for (let n = 0; n < 150 && !outcome; n++) {
          await wait(100);
          if (win.isDestroyed()) outcome = 'opened';
          else {
            error = await js(`document.getElementById('err').textContent`).catch(() => '');
            if (error) outcome = 'refused';
          }
        }
        log({ event: 'password', tried: i, outcome: outcome || 'nothing happened', error });
        if (outcome === 'opened') break;
      }
    })();
  }

  return {
    /** Once its stores and services are up: the plan, when its launcher has drawn. */
    ready({ stores, services }) {
      // Said here, by the copy that holds the profile's lock: a second launch of the profile never gets this far.
      log({ event: 'start', id: profile.id, folder: profile.folder, userData: app.getPath('userData'), protected: Boolean(profile.lock) });
      (async () => {
        let home = null;
        for (let i = 0; i < 300; i++) {
          home = BrowserWindow.getAllWindows().find((w) => !w.isDestroyed() && /[?&]app=home/.test(w.webContents.getURL()));
          if (home && !home.webContents.isLoading() && (await home.webContents.executeJavaScript(`Boolean(document.querySelector('.pf-badge'))`).catch(() => false))) break;
          home = null;
          await wait(100);
        }
        if (!home) return log({ event: 'ready', badge: null });
        const js = (code) => home.webContents.executeJavaScript(code);
        await wait(400);
        const badge = await js(`(() => { const b = document.querySelector('.pf-badge'); return b ? { text: b.querySelector('.pf-badge-name')?.textContent || '', colour: b.style.getPropertyValue('--pf-colour'), id: b.dataset.profile } : null; })()`).catch(() => null);
        const out = { event: 'ready', badge, title: home.getTitle(), setting: plan.setting ? stores.settings.get(plan.setting[0], null) : null };
        if (plan.captureHome) {
          try {
            home.webContents.invalidate();
            await wait(500);
            fs.writeFileSync(plan.captureHome, (await home.webContents.capturePage()).toPNG());
          } catch {
            /* a picture is a nicety */
          }
        }
        if (plan.readSecret) out.secretMatches = stores.secrets.get(plan.readSecret[0]) === plan.readSecret[1];
        if (plan.findContact) out.contactFound = (services?.contacts?.list?.({ query: plan.findContact }) || []).some((c) => String(c.display || '').includes(plan.findContact));
        if (plan.setting) stores.settings.set(plan.setting[0], plan.setting[1]);
        if (plan.secret) {
          try {
            stores.secrets.set(plan.secret[0], plan.secret[1]);
            out.secretSaved = true;
          } catch (err) {
            out.secretSaved = err.message;
          }
        }
        if (plan.contact) {
          await js(`window.rutbaOffice.contacts.save({ contact: { name: { full: ${JSON.stringify(plan.contact)} }, emails: [{ value: 'peer@example.com', type: 'home' }] } }).then(() => 1)`).catch((err) => { out.contactError = String(err); });
          await wait(700);
        }
        if (plan.setPassword) {
          out.passwordSet = await js(`window.rutbaOffice.profiles.setPassword({ password: ${JSON.stringify(plan.setPassword)} }).then((c) => c.protected)`).catch((err) => String(err));
        }
        log(out);
        if (plan.quitAfter) {
          await wait(800);
          leave('done');
        }
      })().catch((err) => log({ event: 'ready', error: String(err?.stack || err) }));
    },
  };
}
