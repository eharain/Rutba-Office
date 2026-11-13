// A document that opened as it was saved, because the unsaved work sealed for
// it would not open here and was kept aside for later (documents.js,
// `setAside`): the window says so. The words are in locked-work.js, which only
// a build made with the switch below imports.
const says = process.env.RUTBA_LOCKER === '1' ? (name, how) => import('./locked-work.js').then((m) => m.setAsideLine(name, how)) : null;

/** Told after an open: a toast when `opened` says its earlier work was set aside. */
export function sayIfSetAside(opened, toast) {
  if (!opened?.setAside || !says) return;
  const name = opened.name || String(opened.path || '').split(/[\\/]/).pop();
  says(name, opened.setAside).then((line) => toast(line, { ms: 9000 })).catch(() => {});
}
