// How a mail account signs in, and in which browser.
//
// In Add account an address alone (from the built-in table, offline) says
// how the account signs in: a Gmail address with Google's own page in a
// browser (OAuth 2.0), a Yahoo one with an app password, an address at a
// domain the table does not know with nothing said until Add looks it up.
// The Google sign-in offers the browsers on this computer, as Windows lists
// them, the default one marked; the one chosen is remembered for the next
// sign-in. Nothing is signed in to and no browser is opened by the check.
// Run alone with RUTBA_VERIFY_ONLY=mail-signin.

/**
 * How long the browsers may take to be listed. Windows is asked with reg.exe,
 * four at once, which answer in well under a second; on a computer busy with
 * other work (the gate's other lanes) starting them took from 6 to 30 seconds.
 */
const BROWSERS_MS = 60000;

/**
 * @param {object} h the harness: open, check, until, wait, capture, errorsIn
 */
export async function verifyMailSignIn(h) {
  const { open, check, until, wait, capture, errorsIn } = h;
  let js = null;
  let saved = null;
  try {
    const win = await open('mail');
    js = (code) => win.webContents.executeJavaScript(code);
    saved = await js(`window.rutbaOffice.store.get({ key: 'mail.signInBrowser', fallback: null })`);
    const waitFor = async (fn, what, ms) => {
      let got = null;
      await until(async () => {
        got = await fn();
        return Boolean(got);
      }, what, ms).catch(() => {});
      return got || null;
    };
    const setAddress = (value) => js(`(() => { const el = document.querySelector('#ml-address'); if (!el) return false; Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(el, ${JSON.stringify(value)}); el.dispatchEvent(new Event('input', { bubbles: true })); return true; })()`);
    const openDialog = async () => {
      await until(() => js(`Boolean([...document.querySelectorAll('button')].find((b) => /^Add account$/.test(b.textContent.trim())))`), 'an Add account button', 10000);
      await js(`[...document.querySelectorAll('button')].find((b) => /^Add account$/.test(b.textContent.trim()))?.click(), 1`);
      await until(() => js(`Boolean(document.querySelector('#ml-address'))`), 'the account dialog', 6000);
    };
    const close = () => js(`[...document.querySelectorAll('.rw-dialog-foot button')].find((b) => b.textContent.trim() === 'Cancel')?.click(), 1`);
    const authLine = () => js(`(() => { const a = document.querySelector('.ml-auth-kind'); return a ? { kind: a.dataset.kind, text: a.textContent.trim() } : null; })()`);

    await openDialog();
    await setAddress('someone.signin@gmail.com');
    const google = await waitFor(async () => {
      const a = await authLine();
      const b = await js(`(() => { const s = document.querySelector('.ml-browser select'); return s ? { options: [...s.options].map((o) => ({ value: o.value, text: o.textContent })), value: s.value } : null; })()`);
      return a?.kind === 'oauth' && b ? { auth: a, browsers: b } : null;
    }, 'Google\'s sign-in with its browsers', BROWSERS_MS);
    check('mail sign-in: a Gmail address says it signs in with Google\'s own page (OAuth 2.0), and the sign-in offers the browsers on this computer, the default marked',
      Boolean(google) && /Signs in with Google's own page \(OAuth 2\.0\), in the browser you choose/.test(google.auth.text) && google.browsers.options[0].value === '' && google.browsers.options.length >= 2 && google.browsers.options.some((o) => /\(the default\)$/.test(o.text)),
      JSON.stringify(google));
    win.webContents.invalidate();
    await wait(700);
    await capture(win, 'mail-signin-1-google.png');

    // A browser chosen is remembered for the next sign-in.
    const other = google?.browsers.options.find((o) => o.value && !/\(the default\)$/.test(o.text)) || google?.browsers.options.find((o) => o.value);
    if (other) await js(`(() => { const s = document.querySelector('.ml-browser select'); Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set.call(s, ${JSON.stringify(other.value)}); s.dispatchEvent(new Event('change', { bubbles: true })); return 1; })()`);
    await wait(300);
    await close();
    await until(() => js(`!document.querySelector('#ml-address')`), 'the dialog closed', 4000).catch(() => {});
    await openDialog();
    await setAddress('someone.signin@gmail.com');
    const kept = await waitFor(() => js(`document.querySelector('.ml-browser select')?.value || null`), 'the browser remembered', BROWSERS_MS);
    check('mail sign-in: the browser chosen is remembered for the next sign-in', Boolean(other) && kept === other.value, JSON.stringify({ chosen: other, kept }));

    await setAddress('someone.signin@yahoo.com');
    const yahoo = await waitFor(async () => {
      const a = await authLine();
      return a?.kind === 'app-password' ? a : null;
    }, 'Yahoo\'s app password said', 8000);
    check('mail sign-in: a Yahoo address says it takes an app password rather than the account\'s own', Boolean(yahoo) && /Yahoo takes an app password/.test(yahoo.text), JSON.stringify(yahoo));

    await setAddress('someone.signin@nowhere-signin.example');
    const unknown = await waitFor(async () => {
      const s = await js(`(() => ({ state: document.querySelector('.ml-search')?.dataset.state, text: document.querySelector('.ml-search')?.textContent || '' }))()`);
      return s.state === 'nothing' ? s : null;
    }, 'an unknown domain said so', 8000);
    check('mail sign-in: an address at a domain the built-in list does not know says that Add will look up its own records, and nothing is looked up while it is typed',
      Boolean(unknown) && /Add looks up its own records: its DNS \(SRV and MX\), autoconfig\.nowhere-signin\.example and autodiscover\.nowhere-signin\.example/.test(unknown.text),
      JSON.stringify(unknown));
    await close();
    const complaints = await errorsIn(win);
    check('mail sign-in: the window reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('mail sign-in: the checks ran', false, err.stack || err.message);
  } finally {
    if (js) await js(`window.rutbaOffice.store.set({ key: 'mail.signInBrowser', value: ${JSON.stringify(saved ?? '')} }).then(() => 1)`).catch(() => {});
  }
}
