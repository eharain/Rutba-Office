// Automate → scripts: the worker a script runs in. It is a bundle of its
// own, served under a policy that lets it work the script out and reach
// nothing — no file, no network (protocol.js) — and it has no bridge to
// the shell. It is handed the workbook's snapshot and hands back the edits.
import { runScript } from '@rutba/sheet-view/scripts';

self.onmessage = (e) => {
  const { code, snapshot } = e.data || {};
  self.postMessage(runScript(String(code || ''), snapshot || { sheets: [] }));
};
