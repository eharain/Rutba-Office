// Live Broadcast → Add → Web page (plan A10; owner, 5 Nov: "user able to flip next
// and previous product from web CMS group as the display live").
//
// A server on this computer plays the store: one page speaks the show-page
// contract (consumer's show-deck.ts) in three colours, one reports nothing and
// flips on the arrow keys, one answers 503 and then recovers, and one tells the
// server what it could reach. The colour is read back from the program's own
// canvas, so a flip is seen to arrive where the viewers would see it. Nothing
// leaves this computer's loopback.
// Run alone with RUTBA_VERIFY_ONLY=web-source.

import fs from 'node:fs';
import path from 'node:path';
import http from 'node:http';
import { emptyStack } from '@rutba/studio/stack';
import { seedSetup, restoreSetup, PANEL, IN_PANEL } from './verify-studio-seed.js';

const COLOURS = [[200, 30, 30], [30, 180, 60], [30, 60, 200]];
const css = (c) => `rgb(${c.join(',')})`;

const deckPage = `<!doctype html><html><body style="margin:0;height:100vh;font:64px sans-serif;color:#fff">
<div id="t"></div><script>
const items = ['Blue lawn kurta', 'Red scarf', 'Green shawl'];
const colours = ${JSON.stringify(COLOURS.map(css))};
let i = 0;
const send = () => {
  document.getElementById('t').textContent = items[i];
  document.body.style.background = colours[i];
  window.rutbaShow?.state({ rutbaShow: 'state', index: i, count: items.length, title: items[i] });
};
window.addEventListener('message', (e) => {
  const c = e.data?.rutbaShow;
  if (c === 'next') i = Math.min(items.length - 1, i + 1);
  else if (c === 'previous') i = Math.max(0, i - 1);
  else if (c === 'first') i = 0;
  else if (c === 'last') i = items.length - 1;
  else if (c === 'go') i = e.data.index;
  else if (c !== 'report') return;
  send();
});
send();
</script></body></html>`;

const plainPage = `<!doctype html><html><body style="margin:0;height:100vh"><script>
const colours = ${JSON.stringify(COLOURS.map(css))};
let i = 0;
document.body.style.background = colours[0];
window.addEventListener('keydown', (e) => {
  if (e.key === 'ArrowRight') i = Math.min(2, i + 1);
  if (e.key === 'ArrowLeft') i = Math.max(0, i - 1);
  document.body.style.background = colours[i];
});
</script></body></html>`;

const lockedPage = `<!doctype html><html><body style="margin:0;background:#333"><script>
(async () => {
  let mic = 'allowed';
  try { await navigator.mediaDevices.getUserMedia({ audio: true }); } catch (e) { mic = 'refused'; }
  let opened = 'opened';
  try { opened = window.open('about:blank') ? 'opened' : 'refused'; } catch { opened = 'refused'; }
  const report = {
    require: typeof require, process: typeof process, office: typeof window.rutbaOffice,
    bridge: Object.keys(window.rutbaShow || {}).join(','), mic, opened,
  };
  fetch('/report?' + encodeURIComponent(JSON.stringify(report)));
})();
</script></body></html>`;

function startStore() {
  const store = { flaky: 503, report: null, server: null, port: 0, asked: [] };
  store.server = http.createServer((req, res) => {
    const u = new URL(req.url, 'http://127.0.0.1');
    store.asked.push(u.pathname);
    const page = (code, body) => { res.writeHead(code, { 'content-type': 'text/html' }); res.end(body); };
    if (u.pathname === '/show/products/deck') page(200, deckPage);
    else if (u.pathname === '/plain') page(200, plainPage);
    else if (u.pathname === '/locked') page(200, lockedPage);
    else if (u.pathname === '/flaky') page(store.flaky, store.flaky === 200 ? deckPage : '<h1>The store is down</h1>');
    else if (u.pathname === '/missing') page(404, '<h1>No such group</h1>');
    else if (u.pathname === '/report') { store.report = JSON.parse(decodeURIComponent(u.search.slice(1))); page(200, 'ok'); }
    else page(404, '');
  });
  return new Promise((resolve) => store.server.listen(0, '127.0.0.1', () => { store.port = store.server.address().port; resolve(store); }));
}

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn
 */
export async function verifyWebSource({ open, check, until, wait, errorsIn }) {
  const ok = (p) => p.then(() => true).catch(() => false);
  let store = null;
  let js = null;
  let kept = null;
  try {
    store = await startStore();
    const base = `http://127.0.0.1:${store.port}`;
    const win = await open('broadcast');
    js = (code) => win.webContents.executeJavaScript(code);
    const studio = { size: 720, fps: 30, countdown: 0, liveCopy: false, minimise: false, shot: 'program', shotTimer: 0 };
    kept = await seedSetup(js, { app: 'broadcast', stack: emptyStack(), studio });
    await win.webContents.reload();
    await until(() => js(`Boolean(document.querySelector('.st-canvas')) && Boolean(document.querySelector('.st-livebar'))`), 'the Live Broadcast window', 10000);
    await js(`window.__rutbaLocalPages = true, 1`);
    const toPanel = async (p, shows) => { await js(PANEL(p)); await until(() => js(IN_PANEL(shows)), `the ${shows} panel`, 5000); };

    const click = (sel) => js(`(() => { const n = document.querySelector(${JSON.stringify(sel)}); if (!n) return 'no ' + ${JSON.stringify(sel)}; n.click(); return 1; })()`);
    const setUrl = (value) => js(`(() => { const n = document.querySelector('.st-web-dialog input'); Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(n, ${JSON.stringify(value)}); n.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
    const openAsk = async () => {
      await toPanel('setup', 'setup');
      await click('.st-add');
      // 5 s was not enough beside another lane (gate 8, 11 Nov); alone it opens at once.
      await until(() => js(`Boolean(document.querySelector('.rw-menu'))`), 'the Add menu', 12000);
      await js(`[...document.querySelectorAll('.rw-menu button')].find((b) => /Web page/.test(b.textContent))?.click(), 1`);
      await until(() => js(`Boolean(document.querySelector('.st-web-dialog input'))`), 'the address box', 5000);
    };
    const addPage = async (url) => {
      await openAsk();
      await setUrl(url);
      await wait(150);
      await click('.st-web-add');
      await until(() => js(`!document.querySelector('.st-web-dialog')`), 'the box closed', 5000);
      await toPanel('main', 'live');
    };
    // What the program's canvas shows at a point, as an index into the colours, or -1.
    const shownColour = () => js(`(() => {
      const c = document.querySelector('.st-air-canvas');
      const d = c.getContext('2d').getImageData(c.width - 20, c.height - 20, 1, 1).data;
      const near = ${JSON.stringify(COLOURS)}.findIndex((k) => Math.abs(k[0] - d[0]) + Math.abs(k[1] - d[1]) + Math.abs(k[2] - d[2]) < 40);
      return near;
    })()`);
    const strip = () => js(`(() => { const s = document.querySelector('.st-flip'); if (!s) return null; return { status: s.dataset.status, reported: s.dataset.reported, text: s.querySelector('.st-flip-now')?.textContent || '', prev: s.querySelector('.st-flip-prev')?.disabled, next: s.querySelector('.st-flip-next')?.disabled }; })()`);
    const pages = () => js(`window.rutbaOffice.capture.webPages()`);

    const capture = async (name) => {
      if (!process.env.RUTBA_VERIFY_CAPTURE) return;
      await wait(900);
      win.webContents.invalidate();
      await wait(700);
      fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, name), (await win.webContents.capturePage()).toPNG());
    };

    // https only.
    await openAsk();
    await setUrl('http://shop.example/show/products/x');
    await wait(150);
    const refused = await js(`document.querySelector('.st-web-add')?.disabled`);
    await setUrl('https://shop.example/show/products/x');
    await wait(150);
    const taken = await js(`!document.querySelector('.st-web-add')?.disabled`);
    await capture('web-add-page.png');
    const direct = await js(`window.rutbaOffice.capture.webOpen({ id: 'nope', url: 'http://shop.example/' }).then(() => 'opened', (e) => e.message)`);
    check('web source: the address box takes only an https address, and the main process refuses anything else whatever the window asks', refused === true && taken && /https/.test(direct), JSON.stringify({ refused, taken, direct }));
    await js(`[...document.querySelectorAll('.st-web-dialog button')].find((b) => b.textContent.trim() === 'Cancel')?.click(), 1`);
    await until(() => js(`!document.querySelector('.st-web-dialog')`), 'the box closed', 5000).catch(() => {});

    // A page that speaks the contract: added by address, its state read back, flipped from the strip.
    await addPage(`${base}/show/products/deck`);
    const reported = await ok(until(async () => (await strip())?.reported === '1', 'the page\'s report', 20000));
    await until(async () => (await shownColour()) === 0, 'the first colour on the program', 15000).catch(() => {});
    const first = await strip();
    const c0 = await shownColour();
    const layers = await js(`document.querySelector('.st-air-canvas').dataset.layers`);
    const page0 = (await pages())[0];
    check('web source: a page added by its address is drawn on the program; the strip reads what it reports (1 of 3, Blue lawn kurta) and Previous is grey at the first item; it is laid out at 1920 x 1080',
      reported && c0 === 0 && layers === 'web' && /1 of 3/.test(first.text) && /Blue lawn kurta/.test(first.text) && first.prev === true && first.next === false && page0?.width === 1920 && page0?.height === 1080,
      JSON.stringify({ reported, c0, layers, first, page0 }));

    await click('.st-flip-next');
    await capture('web-strip.png');
    const flipped = await ok(until(async () => /2 of 3/.test((await strip()).text) && (await shownColour()) === 1, 'the flip on the program', 8000));
    const second = await strip();
    await js(`window.dispatchEvent(new KeyboardEvent('keydown', { key: 'PageDown', bubbles: true })), 1`);
    const keyed = await ok(until(async () => /3 of 3/.test((await strip()).text) && (await shownColour()) === 2, 'Page Down on the program', 8000));
    const last = await strip();
    await js(`window.dispatchEvent(new KeyboardEvent('keydown', { key: 'PageUp', bubbles: true })), 1`);
    const keyedBack = await ok(until(async () => /2 of 3/.test((await strip()).text) && (await shownColour()) === 1, 'Page Up on the program', 8000));
    check('web source: Next on the strip flips the page and the new picture reaches the program; Page Down and Page Up do the same, so a clicker works; Next is grey at the last item',
      flipped && /Red scarf/.test(second.text) && keyed && /Green shawl/.test(last.text) && last.next === true && keyedBack,
      JSON.stringify({ flipped, second, keyed, last, keyedBack }));

    const clearStack = async () => {
      await seedSetup(js, { app: 'broadcast', stack: emptyStack(), studio });
      await win.webContents.reload();
      await until(() => js(`Boolean(document.querySelector('.st-canvas')) && Boolean(document.querySelector('.st-livebar'))`), 'the Live Broadcast window', 10000);
      await js(`window.__rutbaLocalPages = true, 1`);
      await until(async () => (await pages()).length === 0, 'the page closed', 8000).catch(() => {});
    };

    // A page that reports nothing: arrow keys, and only Previous and Next.
    await clearStack();
    const closedOnReload = (await pages()).length === 0;
    await addPage(`${base}/plain`);
    await until(async () => (await shownColour()) === 0, 'the plain page', 15000).catch(() => {});
    await wait(500);
    const plain = await strip();
    await click('.st-flip-next');
    const arrowed = await ok(until(async () => (await shownColour()) === 1, 'the arrow key on the program', 8000));
    check('web source: a page that reports nothing has only Previous and Next, is sent the arrow keys, and the flip reaches the program; a closed window lets its page go',
      plain?.reported === '0' && plain.text === '' && plain.next === false && arrowed && closedOnReload,
      JSON.stringify({ plain, arrowed, closedOnReload }));

    // A page that does not answer: Not available, its error page never drawn, and tried again.
    await clearStack();
    await addPage(`${base}/flaky`);
    const down = await ok(until(async () => (await strip())?.status === 'unavailable', 'Not available', 15000));
    await wait(600);
    const downText = await strip();
    const blank = await shownColour();
    const layersDown = await js(`document.querySelector('.st-air-canvas').dataset.layers`);
    store.flaky = 200;
    const up = await ok(until(async () => (await strip())?.reported === '1', 'the store back', 12000));
    const c1 = await ok(until(async () => (await shownColour()) === 0, 'the page drawn', 8000));
    check('web source: a page answering 503 shows Not available, its error page is never on the program, and it is tried again until it answers',
      down && /Not available/.test(downText.text) && blank === -1 && layersDown === '' && up && c1,
      JSON.stringify({ down, downText, blank, layersDown, up, c1 }));

    await clearStack();
    await addPage(`${base}/missing`);
    const gone = await ok(until(async () => (await strip())?.status === 'unavailable', '404 Not available', 15000));
    check('web source: an address the store does not know (404) shows Not available', gone, JSON.stringify(await strip()));

    // What a page can reach.
    await clearStack();
    await addPage(`${base}/locked`);
    const told = await ok(until(() => store.report, 'the page\'s report of what it could reach', 15000));
    const r = store.report || {};
    check('web source: a page has no Node, no Office and only the one bridge; it cannot have the microphone or open a window',
      told && r.require === 'undefined' && r.process === 'undefined' && r.office === 'undefined' && r.bridge === 'state' && r.mic === 'refused' && r.opened === 'refused',
      JSON.stringify(r));

    const complaints = await errorsIn(win);
    check('web source: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('web source: the check ran to the end', false, err?.message || String(err));
  } finally {
    if (js && kept) await restoreSetup(js, kept).catch(() => {});
    store?.server.closeAllConnections?.();
    store?.server.close();
  }
}
