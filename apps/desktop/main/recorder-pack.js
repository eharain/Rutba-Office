// Recorder's pack: what the core calls to register it (docs/APP-PACKS.md 5.1.2).
// It registers no service of its own: recording the screen is the core's (the capture
// namespace), shared with Video and Presentations, and going live is Live Broadcast's.

export function register() {
  return {};
}
