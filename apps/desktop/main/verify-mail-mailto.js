// Mail: Office Mail as the Windows mail app (Mail 2.25). A mailto: link
// Windows opens Office with opens Mail on a message to write, filled from
// the link (To, Cc, Bcc, Subject and Body); a link given while Mail is open
// goes to that window; and Mail app on the Tools tab says which program
// opens mail links now and, outside a check run, opens Settings at Rutba
// Office's own page of Default apps. A check run opens no Settings.
// Run alone with RUTBA_VERIFY_ONLY=mail-mailto.

import { parseLink } from '@rutba/office-shell/links';
import { sendEvent } from '@rutba/office-shell/electron/main';
import { labelKeyed } from './verify-keys.js';

export async function verifyMailMailto({ open, check, until, wait, errorsIn }) {
  const stamp = Date.now();
  const link = `mailto:amina@work.example,lee@work.example?cc=sam@checks.example&BCC=kit@checks.example&subject=${encodeURIComponent(`From a link ${stamp}`)}&body=${encodeURIComponent('Hello there.')}`;
  const field = (name) => `[...document.querySelectorAll('.ml-compose .rw-field')].find((f) => f.querySelector('label')?.textContent === ${JSON.stringify(name)})?.querySelector('input')?.value ?? null`;
  const fields = `({ to: ${field('To')}, cc: ${field('Cc')}, bcc: ${field('Bcc')}, subject: ${field('Subject')}, body: document.querySelector('.ml-compose .ml-rich, .ml-compose .ml-compose-body')?.textContent || '' })`;
  const discard = `(((list) => list.find((b) => ${labelKeyed("Discard")}(b)) || list.find((b) => b.textContent.trim() === 'Discard'))([...document.querySelectorAll('.ml-compose .rw-btn')])?.click(), 1)`;
  let linked = null;
  let open2 = null;
  try {
    // What Windows hands Office for a mailto: link is the link; the shell takes it to Mail.
    const asked = parseLink(link);
    check('mail: a mailto: link Windows opens Office with is taken to Mail, as it is', asked?.app === 'mail' && asked?.query?.mailto === link, JSON.stringify(asked));

    linked = await open('mail', null, asked.query);
    const ljs = (code) => linked.webContents.executeJavaScript(code);
    await until(() => ljs(`Boolean(document.querySelector('.ml-compose .ml-form'))`), 'a message to write', 10000).catch(() => {});
    const shown = await ljs(fields);
    check('mail: Mail opened on a mailto: link has a message to write, filled from the link, and nothing sent',
      shown.to === 'amina@work.example, lee@work.example' && shown.cc === 'sam@checks.example' && (shown.bcc === null || shown.bcc === 'kit@checks.example')
        && shown.subject === `From a link ${stamp}` && shown.body.includes('Hello there.'),
      JSON.stringify(shown));
    await ljs(discard);

    // A link given while a Mail window is open goes to that window.
    open2 = await open('mail');
    const ojs = (code) => open2.webContents.executeJavaScript(code);
    await until(() => ojs(`document.querySelectorAll('.ml-accounts button').length > 0`), 'the Mail window', 8000).catch(() => {});
    sendEvent(open2, 'app:link', { mailto: `mailto:bo@work.example?subject=${encodeURIComponent(`Second link ${stamp}`)}` });
    await until(() => ojs(`(${field('Subject')}) === ${JSON.stringify(`Second link ${stamp}`)}`), 'the second link written to', 6000).catch(() => {});
    const second = await ojs(fields);
    check('mail: a mailto: link given while Mail is open goes to the open window', second.to === 'bo@work.example' && second.subject === `Second link ${stamp}`, JSON.stringify(second));
    await ojs(discard);
    await wait(300);

    // Mail app, on the Tools tab: says which program opens mail links; a check run opens no Settings.
    await ojs(`((list) => list.find((b) => ${labelKeyed("Tools")}(b)) || list.find((b) => b.textContent.trim() === 'Tools'))([...document.querySelectorAll('.rw-tab')])?.click(), 1`);
    await until(() => ojs(`Boolean(document.querySelector('.ml-mail-app'))`), 'the Mail app button', 5000).catch(() => {});
    await ojs(`document.querySelector('.ml-mail-app')?.click(), 1`);
    await until(() => ojs(`[...document.querySelectorAll('.rw-toast')].some((n) => /mail (app|links) on this computer|system's own settings/.test(n.textContent))`), 'what opens mail links', 5000).catch(() => {});
    const said = await ojs(`[...document.querySelectorAll('.rw-toast')].map((n) => n.textContent).find((x) => /mail (app|links) on this computer|system's own settings/.test(x)) || ''`);
    const asked2 = await ojs(`window.rutbaOffice.defaults.mailApp({ open: false })`);
    const windowsPage = process.platform !== 'win32' || /^ms-settings:defaultapps\?registeredApp(User|Machine)=Rutba%20Office$/.test(asked2.page || '');
    check('mail: Mail app says which program opens mail links, and names Rutba Office\'s own page of Default apps, opening nothing in a check run',
      Boolean(said) && asked2.opened === false && windowsPage, JSON.stringify({ said, ...asked2 }));

    const complaints = [...(await errorsIn(linked)), ...(await errorsIn(open2))];
    check('mail: Office Mail as the mail app reported no error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('mail: the mailto: checks ran to the end', false, err?.message || String(err));
  } finally {
    for (const w of [linked, open2]) {
      try {
        if (w && !w.isDestroyed()) {
          await w.webContents.executeJavaScript(discard).catch(() => {});
          w.close();
        }
      } catch { /* gone */ }
    }
  }
}
