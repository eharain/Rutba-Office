// Mail's way out in a window check run, and the test server it goes to.
//
// A check run never opens a connection to a real mail server (mail.js gives
// it a transport that sends nothing). A check that must see a message
// really leave (verify-ai-tools.js: a message the AI tools asked to send
// goes only when the person presses Send) starts the small SMTP server here
// on this computer and names its port in `globalThis.__rutbaCheckSmtp`; the
// transport then hands the message to it, in plain SMTP on 127.0.0.1, and
// to nothing else. Without a port named, it sends nothing, as before.

import net from 'node:net';
import crypto from 'node:crypto';

/** The transport mail.js is given in a check run (its `smtp`). */
export function checkTransport() {
  return {
    async send({ account, draft }) {
      const to = globalThis.__rutbaCheckSmtp;
      if (!to?.port) return { messageId: `check-${crypto.randomUUID()}` };
      const nodemailer = (await import('nodemailer')).default;
      const transport = nodemailer.createTransport({ host: '127.0.0.1', port: to.port, secure: false, ignoreTLS: true });
      if (draft.raw) return transport.sendMail({ envelope: draft.envelope, raw: draft.raw });
      return transport.sendMail({
        from: { name: account.name, address: account.email },
        to: draft.to,
        cc: draft.cc || undefined,
        bcc: draft.bcc || undefined,
        subject: draft.subject || '',
        text: draft.text || '',
        attachments: (draft.attachments || []).map((a) => ({ filename: a.filename, path: a.path })),
      });
    },
  };
}

/**
 * A test SMTP server on 127.0.0.1, at a port the system gives: every
 * message handed to it is kept in `received` ({ from, to, data }). Speaks
 * only what a sender needs: EHLO or HELO, MAIL, RCPT, DATA, RSET, NOOP, QUIT.
 */
export async function startTestSmtp() {
  const received = [];
  const server = net.createServer((socket) => {
    socket.setEncoding('utf8');
    let buffered = '';
    let inData = false;
    let mail = { from: null, to: [], data: '' };
    const say = (line) => socket.write(`${line}\r\n`);
    say('220 rutba-check ESMTP ready');
    socket.on('data', (chunk) => {
      buffered += chunk;
      for (;;) {
        if (inData) {
          const end = buffered.indexOf('\r\n.\r\n');
          if (end < 0) return;
          mail.data = buffered.slice(0, end).replace(/\r\n\.\./g, '\r\n.');
          buffered = buffered.slice(end + 5);
          inData = false;
          received.push(mail);
          mail = { from: null, to: [], data: '' };
          say('250 kept');
          continue;
        }
        const at = buffered.indexOf('\r\n');
        if (at < 0) return;
        const line = buffered.slice(0, at);
        buffered = buffered.slice(at + 2);
        const verb = line.slice(0, 4).toUpperCase();
        if (verb === 'EHLO') { say('250-rutba-check'); say('250 8BITMIME'); }
        else if (verb === 'HELO') say('250 rutba-check');
        else if (verb === 'MAIL') { mail.from = /<([^>]*)>/.exec(line)?.[1] || ''; say('250 ok'); }
        else if (verb === 'RCPT') { mail.to.push(/<([^>]*)>/.exec(line)?.[1] || ''); say('250 ok'); }
        else if (verb === 'DATA') { inData = true; say('354 go on'); }
        else if (verb === 'RSET') { mail = { from: null, to: [], data: '' }; say('250 ok'); }
        else if (verb === 'NOOP') say('250 ok');
        else if (verb === 'QUIT') { say('221 bye'); socket.end(); return; }
        else say('502 not here');
      }
    });
    socket.on('error', () => {});
  });
  await new Promise((resolve) => server.listen(0, '127.0.0.1', resolve));
  return {
    port: server.address().port,
    received,
    close: () => new Promise((resolve) => server.close(() => resolve())),
  };
}
