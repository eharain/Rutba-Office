// Zoom by Ctrl and the wheel, and by Ctrl+Plus, Ctrl+Minus and Ctrl+0, in
// Documents, Worksheets, Presentations and Mail's reading pane.
//
// Each app's own level moves (the one its status bar's slider or View's
// buttons set), 10% a notch or a press, and the window's own zoom stays at
// 100%, so the ribbon keeps its height. Ctrl+= and Ctrl+Shift+= are both
// Ctrl+Plus, each press is one step (never two: the application menu's
// accelerator must not step again), and Ctrl+0 is 100%. Ctrl and the wheel
// keeps the point under the pointer where it was, to within a few pixels,
// on whichever axis can scroll; a touchpad's pinch (fractional wheel
// amounts) scales smoothly rather than in steps; and over Mail's message,
// a sandboxed frame the page cannot hear, the wheel still zooms the message.
// Pictures keeps its own steps, and its Ctrl+0 that fits the window.
// The zoom is a view setting: nothing is saved, and every app is put back to
// 100% before the check ends. Run alone with RUTBA_VERIFY_ONLY=zoom.

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, capture
 * @param {{ files: object }} args the run's fixtures (docx, xlsx, pptx)
 */
export async function verifyZoom(h, { files }) {
  const { open, check, until, wait, press, errorsIn, capture } = h;

  const APPS = [
    { app: 'word', name: 'documents', file: files.docx, area: '.wd-scroll', target: '.wd-page', ready: '.wd-page .wd-block', presses: 10 },
    { app: 'sheets', name: 'worksheets', file: files.xlsx, area: '.sh-grid', target: '.sh-canvas', ready: '.sh-cell', presses: 5 },
    { app: 'slides', name: 'presentations', file: files.pptx, area: '.sl-stage', target: '.sl-fit', ready: '.sl-fit .sl-svg', presses: 10 },
    { app: 'mail', name: 'mail', file: null, area: '.ml-reader', target: null, ready: '.ml-row', presses: 0 },
  ];

  for (const a of APPS) {
    try {
      const win = await open(a.app, a.file);
      const wc = win.webContents;
      const js = (code) => wc.executeJavaScript(code);
      await until(() => js(`Boolean(document.querySelector(${JSON.stringify(a.ready)}))`), `the ${a.name} window`, 10000);
      if (a.app === 'mail') {
        await js(`document.querySelector('.ml-row')?.click(), 'opened'`);
        await until(() => js(`Boolean(document.querySelector('.ml-reader iframe'))`), 'the reading pane', 8000);
        await wait(400);
      }
      // What the app shows as its level: the status bar's percentage, or
      // for Mail the reading pane's own zoom.
      const level = () => js(a.app === 'mail'
        ? `(() => { const z = document.querySelector('.ml-reader')?.style.zoom; return z ? Math.round(parseFloat(z) * 100) : 100; })()`
        : `(() => { const p = document.querySelector('.rw-zoom-pct')?.textContent || ''; return parseInt(p, 10) || null; })()`);
      const ribbon = () => js(`(() => { const r = document.querySelector('.rw-ribbon'); return r ? Math.round(r.getBoundingClientRect().height * 10) / 10 : 0; })()`);
      // A key press, then the level it should come to, then a beat and the
      // level read again: a second step from the menu would show by then.
      const keyTo = async (keyCode, modifiers, want) => {
        await press(wc, keyCode, { modifiers });
        await until(async () => (await level()) === want, `${a.name} at ${want}%`, 4000).catch(() => {});
        await wait(250);
        return level();
      };
      const ribbonBefore = await ribbon();
      // The page focused, as a person would have it, without moving anything.
      await js(`(() => { const el = document.querySelector(${JSON.stringify(a.area)}); el?.focus?.({ preventScroll: true }); return 1; })()`);

      const steps = [];
      steps.push(['Ctrl+0', await keyTo('0', ['control'], 100)]);
      steps.push(['Ctrl+=', await keyTo('=', ['control'], 110)]);
      steps.push(['Ctrl+Shift+=', await keyTo('Plus', ['control', 'shift'], 120)]);
      steps.push(['Ctrl+-', await keyTo('-', ['control'], 110)]);
      steps.push(['Ctrl+0', await keyTo('0', ['control'], 100)]);
      const wanted = [100, 110, 120, 110, 100];
      check(`${a.name}: Ctrl+Plus (Ctrl+= or Ctrl+Shift+=) and Ctrl+Minus step the ${a.app === 'mail' ? 'message' : 'page'} by 10% a press, and Ctrl+0 puts it at 100%`,
        steps.every(([, got], i) => got === wanted[i]),
        steps.map(([k, got]) => `${k} ${got}%`).join(', '));

      // Ctrl and the wheel, a notch towards the person and one away.
      if (a.target) {
        for (let i = 0; i < a.presses; i++) await press(wc, '=', { modifiers: ['control'] });
        const base = 100 + a.presses * 10;
        await until(async () => (await level()) === base, `${a.name} at ${base}%`, 5000).catch(() => {});
        // Scrolled into the middle, so there is room to keep the point both ways.
        await js(`(() => { const el = document.querySelector(${JSON.stringify(a.area)}); el.scrollTop = Math.min(320, (el.scrollHeight - el.clientHeight) / 2); el.scrollLeft = (el.scrollWidth - el.clientWidth) / 2; return 1; })()`);
        await wait(400);
      }
      const at = await js(`(() => {
        const area = document.querySelector(${JSON.stringify(a.area)});
        const r = area.getBoundingClientRect();
        const head = document.querySelector('.ml-head');
        const h = head ? head.getBoundingClientRect() : null;
        return h ? { x: Math.round(h.left + h.width * 0.4), y: Math.round(h.top + Math.min(40, h.height / 2)) } : { x: Math.round(r.left + r.width * 0.4), y: Math.round(r.top + r.height * 0.45) };
      })()`);
      // Where the point under the pointer sits on the page: as fractions of
      // the page's box, read again once the level has moved.
      const mark = () => js(`(() => {
        const t = ${a.target ? `document.querySelector(${JSON.stringify(a.area)} + ' ' + ${JSON.stringify(a.target)})` : 'null'};
        if (!t) return null;
        const r = t.getBoundingClientRect();
        window.__zoomMark = { fx: (${at.x} - r.left) / r.width, fy: (${at.y} - r.top) / r.height };
        return window.__zoomMark;
      })()`);
      const drift = () => js(`(() => {
        const area = document.querySelector(${JSON.stringify(a.area)});
        const t = ${a.target ? `area.querySelector(${JSON.stringify(a.target)})` : 'null'};
        const m = window.__zoomMark;
        if (!t || !m) return null;
        const r = t.getBoundingClientRect();
        const room = (size, client, pos) => size - client > 1 && pos > 0.5 && pos < size - client - 0.5;
        return {
          dx: Math.round((r.left + m.fx * r.width - ${at.x}) * 10) / 10,
          dy: Math.round((r.top + m.fy * r.height - ${at.y}) * 10) / 10,
          // Whether each axis can still scroll both ways: a page narrower
          // than the window stays centred, as Word's does.
          canX: room(area.scrollWidth, area.clientWidth, area.scrollLeft),
          canY: room(area.scrollHeight, area.clientHeight, area.scrollTop),
        };
      })()`);
      const wheel = (dy, ticks) => {
        wc.sendInputEvent({ type: 'mouseMove', x: at.x, y: at.y });
        wc.sendInputEvent({ type: 'mouseWheel', x: at.x, y: at.y, deltaX: 0, deltaY: dy, wheelTicksX: 0, wheelTicksY: ticks, canScroll: true, modifiers: ['control'] });
      };
      const start = await level();
      await mark();
      // Chromium asks the shell to zoom as well, even for a wheel the page
      // took: counted, to show that a notch is still one step.
      const asked = [];
      const onAsk = (_e, d) => asked.push(d);
      wc.on('zoom-changed', onAsk);
      wheel(100, 1);
      await until(async () => (await level()) === start + 10, `${a.name} a notch in`, 4000).catch(() => {});
      await wait(300);
      const inLevel = await level();
      const inDrift = await drift();
      if (process.env.RUTBA_VERIFY_CAPTURE) { wc.invalidate(); await wait(700); }
      await capture(win, `zoom-${a.app}.png`);
      await mark();
      wheel(-100, -1);
      await until(async () => (await level()) === start, `${a.name} a notch out`, 4000).catch(() => {});
      await wait(300);
      const outLevel = await level();
      const outDrift = await drift();
      const kept = (d) => !d || ((Math.abs(d.dx) <= 3 || !d.canX) && (Math.abs(d.dy) <= 3 || !d.canY));
      check(`${a.name}: Ctrl and the wheel zooms 10% a notch, in and back out`,
        inLevel === start + 10 && outLevel === start,
        `${start}% -> ${inLevel}% -> ${outLevel}%; the shell was asked ${asked.join(', ') || 'nothing'}`);
      if (a.target) {
        check(`${a.name}: the point under the pointer stays where it was as the wheel zooms`,
          kept(inDrift) && kept(outDrift) && Boolean(inDrift) && (inDrift.canY || inDrift.canX),
          `in ${JSON.stringify(inDrift)}; out ${JSON.stringify(outDrift)}; pointer ${at.x},${at.y}`);
      }

      // A touchpad's pinch: small fractional amounts, no notches.
      // Smooth means levels between the 10% steps along the way: the end
      // alone can land on one by chance (150% pinched to exactly 170%).
      const before = await level();
      const seen = [];
      for (let i = 0; i < 12; i++) { wheel(1.25, 0); await wait(20); seen.push(await level()); }
      await until(async () => (await level()) !== before, `${a.name} to follow the pinch`, 3000).catch(() => {});
      await wait(200);
      const pinched = await level();
      seen.push(pinched);
      check(`${a.name}: a touchpad pinch zooms smoothly, not in 10% steps`,
        pinched > before && pinched < Math.round(before * 1.4) && seen.some((l) => l !== before && l % 10 !== 0),
        `${before}% -> ${pinched}% (along the way ${[...new Set(seen)].join(', ')})`);

      if (a.app === 'mail') {
        // Over the message itself, a sandboxed frame the page cannot hear.
        await keyTo('0', ['control'], 100);
        const frame = await js(`(() => {
          const f = document.querySelector('.ml-reader iframe').getBoundingClientRect();
          const x = Math.round(f.left + f.width / 2);
          const y = Math.round(f.top + Math.min(120, f.height / 2));
          return { x, y, under: document.elementFromPoint(x, y)?.tagName || '' };
        })()`);
        wc.sendInputEvent({ type: 'mouseMove', x: frame.x, y: frame.y });
        wc.sendInputEvent({ type: 'mouseWheel', x: frame.x, y: frame.y, deltaX: 0, deltaY: 100, wheelTicksX: 0, wheelTicksY: 1, canScroll: true, modifiers: ['control'] });
        await until(async () => (await level()) === 110, 'the message a notch in from over its frame', 4000).catch(() => {});
        await wait(400);
        const overFrame = await level();
        check('mail: Ctrl and the wheel over the message itself zooms the message, one step a notch', overFrame === 110, `${overFrame}% (under the pointer: ${frame.under})`);
        // Where the frame keeps the wheel to itself, Chromium's own request
        // is all there is: the shell hands it on as Zoom In, the same
        // command the application menu sends.
        await wait(500);
        wc.emit('zoom-changed', { preventDefault() {} }, 'in');
        await until(async () => (await level()) === 120, 'the message a step in from the shell', 4000).catch(() => {});
        await wait(300);
        const fromShell = await level();
        check('mail: Zoom In from the shell (the menu, or a wheel the page never heard) zooms the message one step', fromShell === 120, `${fromShell}%`);
      }

      wc.removeListener('zoom-changed', onAsk);
      const ribbonAfter = await ribbon();
      const factor = wc.getZoomFactor();
      check(`${a.name}: the ribbon keeps its height and the window its 100% through all of it`,
        ribbonBefore > 0 && Math.abs(ribbonAfter - ribbonBefore) < 1 && Math.abs(factor - 1) < 0.001,
        `ribbon ${ribbonBefore} -> ${ribbonAfter}; window factor ${factor}`);

      // Back to 100% (Mail keeps its level between runs).
      const reset = await keyTo('0', ['control'], 100);
      const complaints = await errorsIn(win);
      check(`${a.name}: the zoom keys and wheel report nothing, and Ctrl+0 leaves it at 100%`, complaints.length === 0 && reset === 100, `${reset}%; ${complaints.join(' | ') || 'nothing reported'}`);
    } catch (err) {
      check(`${a.name}: the Ctrl and wheel zoom checks ran`, false, err.message);
    }
  }

  // Pictures keeps its own: Ctrl and the wheel over the picture, Ctrl+Plus
  // and Ctrl+Minus a quarter at a time, Ctrl+0 back to fitting the window.
  if (files.png) {
    try {
      const win = await open('pictures', files.png);
      const wc = win.webContents;
      const js = (code) => wc.executeJavaScript(code);
      await until(() => js(`Boolean(document.querySelector('.pv-stage .pv-image'))`), 'the picture', 8000);
      const zoomed = () => js(`Boolean(document.querySelector('.pv-stage.zoomed'))`);
      const at = await js(`(() => { const r = document.querySelector('.pv-stage').getBoundingClientRect(); return { x: Math.round(r.left + r.width / 2), y: Math.round(r.top + r.height / 2) }; })()`);
      const seen = [];
      wc.sendInputEvent({ type: 'mouseMove', x: at.x, y: at.y });
      wc.sendInputEvent({ type: 'mouseWheel', x: at.x, y: at.y, deltaX: 0, deltaY: 100, wheelTicksX: 0, wheelTicksY: 1, canScroll: true, modifiers: ['control'] });
      seen.push(['wheel', await until(zoomed, 'the picture zoomed by the wheel', 3000).catch(() => false)]);
      await press(wc, '0', { modifiers: ['control'] });
      seen.push(['Ctrl+0', await until(async () => !(await zoomed()), 'the picture fitted again', 3000).catch(() => false)]);
      await press(wc, '=', { modifiers: ['control'] });
      seen.push(['Ctrl+=', await until(zoomed, 'the picture zoomed by Ctrl+=', 3000).catch(() => false)]);
      await press(wc, '0', { modifiers: ['control'] });
      seen.push(['Ctrl+0', await until(async () => !(await zoomed()), 'the picture fitted again', 3000).catch(() => false)]);
      check('pictures: Ctrl and the wheel, and Ctrl+=, still zoom the picture, and Ctrl+0 fits it to the window again',
        seen.every(([, ok]) => ok === true) && Math.abs(wc.getZoomFactor() - 1) < 0.001,
        seen.map(([k, ok]) => `${k} ${ok === true ? 'yes' : 'no'}`).join(', ') + `; window factor ${wc.getZoomFactor()}`);
    } catch (err) {
      check('pictures: the Ctrl and wheel zoom checks ran', false, err.message);
    }
  }
}
