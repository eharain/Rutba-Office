// Mail: a rich signature, and the composer's rich text.
//
// The signature editor makes words bold and links them; what it keeps is
// cut down to what a signature can carry — an event handler or a script
// pasted in with a web page goes — and its plain words are kept beside it.
// A new message written as rich text carries it as one marked block. In the
// composer, Insert link makes a link (Electron has no window.prompt, which
// it used to ask with), and switching to Plain text and back keeps what was
// typed. Run alone with RUTBA_VERIFY_ONLY=richsignature.

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn
 */
export async function verifyMailRichSignature(h) {
  const { open, check, until, wait, errorsIn } = h;
  let win = null;
  let accountId = null;
  try {
    win = await open('mail');
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`document.querySelectorAll('.ml-accounts button').length > 0`), 'the seeded account in the sidebar', 8000);
    accountId = await js(`(async () => (await window.rutbaOffice.mail.accounts())[0]?.id || null)()`);
    const press = (scope, title) => js(`(() => { const b = [...document.querySelectorAll(${JSON.stringify(scope)} + ' button')].find((n) => (n.dataset.tip || n.title || n.textContent.trim()) === ${JSON.stringify(title)}); if (!b) return 'no ' + ${JSON.stringify(title)}; b.click(); return 'pressed'; })()`);
    const typeLink = async (url) => {
      await until(() => js(`Boolean(document.querySelector('.ml-linkrow input'))`), 'the link row', 3000);
      await js(`(() => { const el = document.querySelector('.ml-linkrow input'); const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set; setter.call(el, ${JSON.stringify(url)}); el.dispatchEvent(new Event('input', { bubbles: true })); return true; })()`);
      await wait(100);
      return js(`(() => { const b = [...document.querySelectorAll('.ml-linkrow button')].find((n) => n.textContent.trim() === 'Add link'); if (!b || b.disabled) return 'no Add link'; b.click(); return 'added'; })()`);
    };

    // The signature: "Jane Doe" in bold, then "Northwind" linked.
    await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Folder')?.click(), 'tab'`);
    await until(() => js(`Boolean([...document.querySelectorAll('button')].find((b) => /^Signature$/.test(b.textContent.trim()) && !b.disabled))`), "the account's Signature button", 5000);
    await js(`[...document.querySelectorAll('button')].find((b) => /^Signature$/.test(b.textContent.trim()) && !b.disabled)?.click(), 'clicked'`);
    await until(() => js(`Boolean(document.querySelector('.ml-signature-editor'))`), 'the signature editor', 5000);
    await js(`(() => {
      const el = document.querySelector('.ml-signature-editor');
      el.focus();
      document.execCommand('insertText', false, 'Jane Doe');
      const range = document.createRange();
      range.selectNodeContents(el);
      const sel = window.getSelection(); sel.removeAllRanges(); sel.addRange(range);
      return 'typed';
    })()`);
    const bold = await press('.ml-signature-tools', 'Bold');
    await js(`(() => {
      const el = document.querySelector('.ml-signature-editor');
      const sel = window.getSelection(); sel.selectAllChildren(el); sel.collapseToEnd();
      document.execCommand('insertLineBreak');
      document.execCommand('insertText', false, 'Northwind');
      const text = [...el.childNodes].reverse().find((n) => n.nodeType === 3 && n.nodeValue.includes('Northwind'))
        || [...el.querySelectorAll('*')].flatMap((n) => [...n.childNodes]).find((n) => n.nodeType === 3 && n.nodeValue.includes('Northwind'));
      const range = document.createRange();
      range.setStart(text, text.nodeValue.indexOf('Northwind'));
      range.setEnd(text, text.nodeValue.indexOf('Northwind') + 'Northwind'.length);
      sel.removeAllRanges(); sel.addRange(range);
      return 'selected';
    })()`);
    await press('.ml-signature-tools', 'Link');
    const linked = await typeLink('https://northwind.example');
    // Something pasted from a web page, with a handler and a script in it.
    await js(`(() => { document.querySelector('.ml-signature-editor').insertAdjacentHTML('beforeend', '<span onclick="window.__clicked = 1">Phone 555</span><script>window.__ran = 1<\\/script>'); return true; })()`);
    await js(`[...document.querySelectorAll('.rw-dialog button')].find((b) => /^Save$/.test(b.textContent.trim()))?.click(), 'saved'`);
    await until(() => js(`!document.querySelector('.ml-signature-editor')`), 'the signature editor to close', 5000);
    const account = await js(`(async () => (await window.rutbaOffice.mail.accounts()).find((a) => a.id === ${JSON.stringify(accountId)}))()`);
    const html = account?.signatureHtml || '';
    check('mail: the signature editor makes words bold and links them, and keeps their plain words beside them',
      bold === 'pressed' && linked === 'added' && /<b>Jane Doe(<br>|<\/b>)/.test(html) && /<a href="https:\/\/northwind\.example">Northwind<\/a>/.test(html)
        && /^Jane Doe\nNorthwind <https:\/\/northwind\.example>/.test(account?.signature || ''),
      `${bold}/${linked}; ${html.slice(0, 160)}; plain ${JSON.stringify(account?.signature)}`);
    check('mail: what a signature cannot carry — an event handler, a script — is not kept',
      /Phone 555/.test(html) && !/onclick|script|__ran/i.test(html), html.slice(-120));

    // A new message, as rich text: the signature as one marked block.
    await js(`document.querySelector('.ml-compose-cta button')?.click(), 'clicked'`);
    await until(() => js(`Boolean(document.querySelector('.ml-rich [data-rutba-signature]'))`), 'the signature in the rich body', 5000);
    const block = await js(`document.querySelector('.ml-rich [data-rutba-signature]').innerHTML`);
    check('mail: a new rich message carries the signature, bold and link and all, as one marked block',
      /^-- <br>/.test(block) && /<b>Jane Doe(<br>|<\/b>)/.test(block) && /href="https:\/\/northwind\.example"/.test(block), block.slice(0, 160));

    // Insert link, in the composer.
    await js(`(() => {
      const el = document.querySelector('.ml-rich');
      el.focus();
      const sel = window.getSelection(); sel.selectAllChildren(el.firstChild || el); sel.collapseToStart();
      document.execCommand('insertText', false, 'See the docs');
      const text = [...el.querySelectorAll('*'), el].flatMap((n) => [...n.childNodes]).find((n) => n.nodeType === 3 && n.nodeValue.includes('See the docs'));
      const range = document.createRange();
      range.setStart(text, text.nodeValue.indexOf('docs'));
      range.setEnd(text, text.nodeValue.indexOf('docs') + 4);
      sel.removeAllRanges(); sel.addRange(range);
      return 'selected';
    })()`);
    await press('.ml-toolbar', 'Insert link');
    const composed = await typeLink('https://docs.example');
    const docsLink = await until(() => js(`Boolean(document.querySelector('.ml-rich a[href="https://docs.example"]'))`), 'the link in the message', 3000).catch(() => false);
    check('mail: Insert link in the composer makes a link of the selected words',
      composed === 'added' && docsLink === true && (await js(`document.querySelector('.ml-rich a[href="https://docs.example"]')?.textContent`)) === 'docs', composed);

    // Plain text and back: what was typed stays.
    await press('.ml-toolbar', 'Plain text');
    await until(() => js(`Boolean(document.querySelector('.ml-compose-body'))`), 'the plain body', 3000);
    const plain = await js(`document.querySelector('.ml-compose-body').value`);
    await js(`[...document.querySelectorAll('.ml-toolbar button')].find((b) => b.textContent.trim() === 'Rich text')?.click(), 'rich'`);
    await until(() => js(`Boolean(document.querySelector('.ml-rich'))`), 'the rich editor again', 3000);
    const again = await js(`document.querySelector('.ml-rich').innerText`);
    check('mail: switching to Plain text and back keeps what was typed, the signature with it',
      /^See the docs <https:\/\/docs\.example>/.test(plain) && /-- \nJane Doe\nNorthwind/.test(plain) && /See the docs/.test(again) && /Jane Doe/.test(again),
      JSON.stringify(plain.slice(0, 120)));
    await js(`[...document.querySelectorAll('button')].find((b) => /^Discard$/.test(b.textContent.trim()))?.click(), 'closed'`);

    const complaints = await errorsIn(win);
    check('mail: the rich signature checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('mail: the rich signature check ran', false, err.message);
  } finally {
    // The account is a fixture the other mail checks share.
    if (win && accountId) {
      await win.webContents.executeJavaScript(`window.rutbaOffice.mail.updateAccount({ id: ${JSON.stringify(accountId)}, patch: { signature: '', signatureHtml: null } })`).catch(() => {});
    }
  }
}
