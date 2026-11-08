// The Recorder opened for another window (A7): Record Screen and Record Voice in
// Video, an introduction's scene, a recording for a slide. The window that asked
// and this one talk over a BroadcastChannel named in the window's address; the
// other half is recording/open-recorder.js.

/** What this window was opened for, or null: { into: 'video'|'intro'|'slide', channel, title, kind }. */
export function readInto() {
  const q = new URLSearchParams(location.search);
  const into = q.get('into');
  if (!['video', 'intro', 'slide'].includes(into) || !q.get('channel')) return null;
  return { into, channel: q.get('channel'), title: q.get('title') || '', kind: q.get('kind') || '' };
}

/** @returns {{ post: (message: object) => void, close: () => void }} */
export function openLink(channel, onMessage) {
  const bc = new BroadcastChannel(channel);
  bc.onmessage = (e) => onMessage(e.data || {});
  return { post: (m) => { try { bc.postMessage(m); } catch { /* the other side has gone */ } }, close: () => bc.close() };
}
