// Mail as Talk sends through it (Office 2.25, D-2026-11-11-01 and -02): an
// invitation goes from the person's own sending account, the one Mail
// marks "Send new mail from", as a personal note.
//
// Mail's own contract for this is `defaultAccount()`, which gives
// { id, address, name } or null, and `sendFrom({ account, to, subject,
// text, html })`, which sends through that account, files the copy in Sent
// and answers { messageId, refused, accountId, from }. It throws with
// err.code 'no-account' (no account to send from: Rutba sends instead) or
// 'cannot-send' (the account could not send: the person is told), and
// nothing is in Sent then. This is the one place Talk meets it. A Mail from before that
// contract is still reached the older way: the first account with a server
// is the default, and `send({ accountId, draft })` sends.

/** An account as Talk shows it: { id, address, name }, or null for one with no server to send through. */
const shaped = (a) => {
  if (!a || a.local) return null;
  const address = String(a.address || a.email || '').trim();
  return a.id != null && address ? { id: a.id, address, name: String(a.name || '') } : null;
};

/**
 * @param {object|null} mail what the Talk service was given: Mail itself, or a
 *   stand-in with accounts() and send(), or nothing when Mail is absent
 */
export function talkMail(mail) {
  const accounts = async () => (await Promise.resolve(mail?.accounts?.() || [])).map(shaped).filter(Boolean);
  return {
    /** The accounts an invitation can go from. */
    accounts,

    /** The sending account: Mail's choice, else the first that can send, else null. */
    async defaultAccount() {
      if (typeof mail?.defaultAccount === 'function') {
        const chosen = shaped(await mail.defaultAccount());
        if (chosen) return chosen;
      }
      return (await accounts())[0] || null;
    },

    /** Sends from `account` (an id), or from the sending account when none is named. Answers { from } or throws. */
    async sendFrom({ account = null, to, subject, text, html }) {
      const list = await accounts();
      const from = account != null ? list.find((a) => a.id === account) : await this.defaultAccount();
      if (!from) throw Object.assign(new Error('There is no Mail account to send from.'), { code: 'no-account' });
      let sent;
      if (typeof mail?.sendFrom === 'function') sent = await mail.sendFrom({ account: from.id, to, subject, text, html });
      else if (typeof mail?.send === 'function') sent = await mail.send({ accountId: from.id, draft: { to, subject, text, html } });
      else throw Object.assign(new Error('Mail is not installed on this computer.'), { code: 'no-account' });
      // The server took the message but turned this person down: not sent to them.
      const refused = (sent?.refused || []).map((a) => String(a).trim().toLowerCase());
      if (refused.includes(String(to).trim().toLowerCase())) throw Object.assign(new Error(`The mail server turned down ${to}.`), { code: 'refused' });
      return { from, messageId: sent?.messageId ?? null };
    },
  };
}
