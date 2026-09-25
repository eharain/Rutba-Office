// Automatic replies (out of office), decided the way rules are: as data in,
// a plan out, nothing touched until the caller acts on it.
//
// RFC 3834 exists because two out-of-office repliers can otherwise mail each
// other forever, and because a mailing list does not want to hear from every
// member's away message at once. Its rule is simple — say what you are, and
// do not answer anything that already said the same — and it is the whole of
// what makes an automatic reply a good citizen rather than a nuisance. None
// of it needs the network or the store: it is a function of a message and an
// account's settings, which is what lets a unit test see the whole decision.

/** One header's value off a stored message, whatever case the account sent it in. */
function header(message, name) {
  const key = String(name).toLowerCase();
  const hit = (message.headers || []).find((h) => (h.key || h.name || '').toLowerCase() === key);
  return hit ? hit.value : null;
}

const addressOf = (person) => String(person?.address || person || '').trim().toLowerCase();

/** Senders nothing ever auto-replies to, whatever the account's settings say. */
const NEVER_REPLY = /^(no-?reply|mailer-daemon|postmaster)@/i;

/**
 * Why this message gets no automatic reply, or null when one is owed.
 *
 * The RFC 3834 checks first — a message that already says it is automatic
 * (`Auto-Submitted` other than "no"), one addressed to a list rather than a
 * person (`Precedence: bulk/list/junk`, or a `List-Id`/`List-Unsubscribe`
 * header), one from an address that exists only to bounce mail — then the
 * two this account adds itself: its own address, so a copy to yourself or a
 * forwarding loop never answers itself, and a sender already answered this
 * period, so the same person does not get a second "I'm away" for a second
 * message.
 */
export function skipReason(message, account, { sentTo = [] } = {}) {
  const from = addressOf(message.from?.[0]);
  if (!from) return 'no sender';
  if (from === addressOf(account?.email)) return 'from this account itself';
  if (NEVER_REPLY.test(from)) return 'a no-reply address';

  const autoSubmitted = header(message, 'auto-submitted');
  if (autoSubmitted && !/^no$/i.test(autoSubmitted.trim())) return 'already an automatic message';

  const precedence = header(message, 'precedence');
  if (precedence && /^(bulk|list|junk)$/i.test(precedence.trim())) return 'bulk mail';

  if (header(message, 'list-id') || header(message, 'list-unsubscribe')) return 'a mailing list';

  if (sentTo.includes(from)) return 'already answered this period';

  return null;
}

/** Whether `atMs` (the moment mail arrived) falls inside the settings' start/end window. */
export function withinWindow(atMs, settings) {
  if (settings?.start && atMs < new Date(settings.start).getTime()) return false;
  if (settings?.end && atMs > new Date(settings.end).getTime()) return false;
  return true;
}

/**
 * Which of these newly-arrived messages get an automatic reply, in order.
 *
 * `isContact(address)` is asked only when the account restricts replies to
 * people in Contacts; passing none is the same as the option being off. A
 * sender answered once by this call is not answered again by it — the same
 * "once per period" rule that keeps `sentTo` from growing applies within a
 * single batch of arrivals, not just across two.
 */
export function planAutoReplies(messages, account, atMs, { sentTo = [], isContact = null } = {}) {
  const settings = account?.autoReply;
  if (!settings?.enabled) return [];
  if (!withinWindow(atMs, settings)) return [];

  const answered = new Set(sentTo);
  const out = [];
  for (const message of messages) {
    const reason = skipReason(message, account, { sentTo: [...answered] });
    if (reason) continue;
    const from = addressOf(message.from?.[0]);
    if (settings.contactsOnly && isContact && !isContact(from)) continue;
    out.push({ message, to: from });
    answered.add(from);
  }
  return out;
}

/**
 * The reply itself: RFC 3834's `Auto-Submitted: auto-replied`, the
 * `In-Reply-To`/`References` that make it thread under the message it
 * answers, and a subject built from the account's own template — "Automatic
 * reply: <original subject>" unless the account wrote its own, in which
 * case `<original subject>` in that text is filled in the same way.
 */
export function buildReply(message, account, settings) {
  const original = message.subject || '(no subject)';
  const template = settings?.subject?.trim() || 'Automatic reply: <original subject>';
  const subject = template.includes('<original subject>') ? template.replace(/<original subject>/g, original) : template;
  const references = [...(message.references || []), message.messageId].filter(Boolean);
  return {
    to: message.from?.[0]?.address || message.from?.[0] || '',
    subject,
    text: settings?.message || '',
    inReplyTo: message.messageId || undefined,
    references,
    headers: { 'Auto-Submitted': 'auto-replied' },
  };
}
