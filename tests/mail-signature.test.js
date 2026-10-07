// The signature block: where it lands, and when it is left alone.

import test from 'node:test';
import assert from 'node:assert/strict';
import { withSignature, swapSignature } from '../packages/mailbox/src/signature.js';

test('a new message gets the signature at the end, after a blank line and the dash line', () => {
  const body = withSignature('Thanks for the update.', 'Jane Doe\nNorthwind Traders');
  assert.equal(body, 'Thanks for the update.\n\n-- \nJane Doe\nNorthwind Traders');
});

test('a reply puts the signature above the quoted text, below the cursor', () => {
  const quote = '\n\nOn Tuesday, Amina Yusuf wrote:\n> the numbers are in';
  const body = withSignature(quote, 'Jane Doe', { reply: true });
  assert.equal(body, '\n\n-- \nJane Doe\n\n\n\nOn Tuesday, Amina Yusuf wrote:\n> the numbers are in');
  // The signature reads before the quote, whatever else is above the block.
  assert.ok(body.indexOf('Jane Doe') < body.indexOf('the numbers are in'));
});

test('an account with no signature inserts nothing, for a new message or a reply', () => {
  assert.equal(withSignature('Hello', ''), 'Hello');
  assert.equal(withSignature('', null), '');
  assert.equal(withSignature('> quoted', undefined, { reply: true }), '> quoted');
});

test('switching accounts swaps the block when it is still exactly what was inserted', () => {
  const drafted = withSignature('', 'Jane Doe');
  const swapped = swapSignature(drafted, 'Jane Doe', 'J. Doe, Northwind Traders');
  assert.equal(swapped, withSignature('', 'J. Doe, Northwind Traders'));
  assert.ok(!swapped.includes('Jane Doe') || swapped.includes('J. Doe'));
});

test('a reply-shaped signature block swaps too, leaving the quote untouched', () => {
  const quote = '\n\nOn Tuesday, Amina Yusuf wrote:\n> hello';
  const drafted = withSignature(quote, 'Jane Doe', { reply: true });
  const swapped = swapSignature(drafted, 'Jane Doe', 'Jane D.');
  assert.equal(swapped, withSignature(quote, 'Jane D.', { reply: true }));
  assert.ok(swapped.endsWith(quote));
});

test('an edited signature is left alone on the next account switch', () => {
  const drafted = withSignature('Hi there', 'Jane Doe');
  const edited = `${drafted}\nP.S. one more thing`; // the block is no longer at the very end
  const swapped = swapSignature(edited, 'Jane Doe', 'Someone Else');
  assert.equal(swapped, edited);
  assert.ok(swapped.includes('Jane Doe'));
});

test('there is nothing to swap when the old account had no signature', () => {
  const body = 'Hello, no signature was ever inserted here.';
  assert.equal(swapSignature(body, '', 'New Signature'), body);
});

/* ── the rich kind ──────────────────────────────────────────────────────── */

import { withSignatureHtml, signatureBlockHtml, signatureTextToHtml, SIGNATURE_ATTR } from '../packages/mailbox/src/signature.js';

test('a rich signature goes in as one marked block, after a line to type on, or above a quote', () => {
  const sig = '<b>Jane Doe</b><br><a href="https://northwind.example">Northwind</a>';
  const fresh = withSignatureHtml('', sig);
  assert.equal(fresh, `<div><br></div><div><br></div><div ${SIGNATURE_ATTR}="">-- <br>${sig}</div>`);
  const quote = '<div>On Tuesday, Amina wrote:</div><blockquote>the numbers are in</blockquote>';
  const reply = withSignatureHtml(quote, sig, { reply: true });
  assert.ok(reply.indexOf('Jane Doe') < reply.indexOf('the numbers are in'), 'above the quote');
  assert.ok(reply.endsWith(quote));
  assert.equal(withSignatureHtml('<div>Hi</div>', ''), '<div>Hi</div>', 'no signature, no change');
  assert.equal(signatureBlockHtml('   '), '');
});

test('a plain signature becomes HTML with its lines and nothing it could be mistaken for', () => {
  assert.equal(signatureTextToHtml('Jane Doe\nNorthwind <Traders> & Co\n\n'), 'Jane Doe<br>Northwind &lt;Traders&gt; &amp; Co');
  assert.equal(signatureTextToHtml(''), '');
});
