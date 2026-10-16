// A link followed in a message (owner, 2026-10-16: "the link in email do not
// open it leads to a blank page . it should open browser instead while spam
// should have a protection on links"): every link asks for a window of its own,
// which the window refuses and hands to Mail; which links are warned about first, and why.
import test from 'node:test';
import assert from 'node:assert/strict';
import { linkWords, linkRisks, linkKind, mailtoFields, namedHost } from '../apps/desktop/renderer/apps/mail/links.js';
import { bodyDocument } from '../apps/desktop/renderer/apps/mail/parts.js';

test('the body sends every link to a window of its own, which the window refuses and hands to Mail: none is followed in the frame', () => {
  const doc = bodyDocument({ html: '<p><a href="https://example.com/" target="_blank" rel="noopener">Example</a> and <a target=\'_top\' href="https://b.example/">B</a></p>' });
  assert.ok(!/<a\b[^>]*\btarget\s*=/i.test(doc), 'the links’ own targets are gone');
  assert.match(doc, /<base target="_blank">/);
  assert.match(doc, /href="https:\/\/example\.com\/"/, 'the link is kept');
});

test('the words of each link, by where it goes', () => {
  const words = linkWords('<a href="https://x.example/a?b=1&amp;c=2"><b>Sign</b> in</a><a href="https://x.example/a?b=1&amp;c=2">again</a>');
  assert.deepEqual(words.get('https://x.example/a?b=1&c=2'), ['Sign in', 'again']);
  assert.equal(namedHost('www.PayPal.com'), 'paypal.com');
  assert.equal(namedHost('https://bank.example/login'), 'bank.example');
  assert.equal(namedHost('Click here'), null);
});

test('an ordinary link in an ordinary message opens without a word', () => {
  assert.deepEqual(linkRisks('https://news.example/today', { html: '<a href="https://news.example/today">Read today’s news</a>' }), []);
  assert.deepEqual(linkRisks('https://www.news.example/', { html: '<a href="https://www.news.example/">news.example</a>' }), [], 'www is the same site');
  assert.deepEqual(linkRisks('https://mail.news.example/', { html: '<a href="https://mail.news.example/">news.example</a>' }), [], 'a site’s own part is the same site');
});

test('junk, a failed sender check, and a link whose words name another site are each warned about', () => {
  const html = '<a href="https://evil.example/login">www.paypal.com</a>';
  const r = linkRisks('https://evil.example/login', { html, inJunk: true, authFailed: true });
  assert.equal(r.length, 3, r.join(' | '));
  assert.match(r[0], /Junk/);
  assert.match(r[1], /SPF, DKIM, DMARC/);
  assert.match(r[2], /reads paypal\.com but goes to evil\.example/);
});

test('an address that is itself a trick is warned about', () => {
  assert.match(linkRisks('https://paypal.com@evil.example/')[0], /really goes to evil\.example/);
  assert.match(linkRisks('http://192.168.4.7/pay')[0], /bare number \(192\.168\.4\.7\)/);
  assert.match(linkRisks('https://xn--pypal-4ve.com/')[0], /look like others/);
});

test('what the window does with a link: the web to the browser, mailto to a new message, anything else refused', () => {
  assert.equal(linkKind('https://a.example/'), 'web');
  assert.equal(linkKind('HTTP://a.example/'), 'web');
  assert.equal(linkKind('mailto:someone@example.com'), 'mail');
  for (const bad of ['javascript:alert(1)', 'file:///C:/Windows/system32/calc.exe', 'ms-settings:', 'data:text/html,hi', '']) assert.equal(linkKind(bad), 'refused', bad);
  assert.deepEqual(mailtoFields('mailto:ana%40example.com?subject=Hello%20there&body=Hi'), { to: 'ana@example.com', subject: 'Hello there', text: 'Hi' });
});
