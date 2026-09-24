// A signature block — the plain-text kind, not a cryptographic one.
//
// Nothing here touches Electron, a window or a store: it is handed a body and
// a signature and hands back a body, which is what lets the composer and the
// engine tests agree on exactly what "insert" and "swap" mean without either
// one standing up a mail account to find out.

/** The line mail readers use to tell a signature from the message above it. */
const DASH_LINE = '-- ';

/**
 * The block a signature becomes once it is inserted: a blank line, the dash
 * line, then the signature itself, trailing blank lines trimmed off so two
 * accounts with differently-tidy signatures still produce a block that
 * compares equal once inserted and unedited.
 */
function block(signature) {
  const text = String(signature || '').replace(/\s+$/, '');
  if (!text) return '';
  return `\n\n${DASH_LINE}\n${text}`;
}

/**
 * Add a signature to a body. An empty signature is a no-op — nothing is
 * inserted, and the body comes back exactly as it was handed in.
 *
 * A new message gets the block at the end, after the rest of what was typed.
 * A reply or forward gets it at the front instead: the body passed in for
 * those is already the quoted or forwarded message, the cursor sits above all
 * of it, and the signature belongs between the cursor and the quote — not
 * after it, where nobody would see it without scrolling past their own words.
 */
export function withSignature(body, signature, { reply = false } = {}) {
  const sig = block(signature);
  const text = body || '';
  if (!sig) return text;
  return reply ? `${sig}\n\n${text}` : `${text}${sig}`;
}

/**
 * Swap one account's signature for another's, in a body that may already
 * hold other edits.
 *
 * The only thing this is willing to touch is the exact block `withSignature`
 * would have put there for the old signature, found at the very start (the
 * reply/forward shape) or the very end (the new-message shape) of the body.
 * Found there unaltered, it is replaced with the new account's block, same
 * position. Found anywhere else, trimmed, merged into the person's own
 * sentence, anything but untouched — the body comes back exactly as handed
 * in, because a signature swap that overwrites something a person wrote is
 * worse than one that does nothing.
 */
export function swapSignature(body, oldSignature, newSignature) {
  const text = body || '';
  const oldBlock = block(oldSignature);
  const newBlock = block(newSignature);
  if (!oldBlock) return text;
  if (text.startsWith(oldBlock)) return newBlock + text.slice(oldBlock.length);
  if (text.endsWith(oldBlock)) return text.slice(0, text.length - oldBlock.length) + newBlock;
  return text;
}
