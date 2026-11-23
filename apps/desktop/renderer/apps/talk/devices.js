// The microphone and the camera as a meeting takes them: why one could not
// be used, and what a person chooses before coming in. Nothing of a screen
// here, so it is tested without a window.
//
// Before coming in, the microphone is only chosen, never opened: nothing
// listens until the person has joined. The camera opens when they turn it
// on, so they see their own picture first, and that same camera goes into
// the meeting with them.

/**
 * Why a microphone or a camera could not be used, from what the browser
 * threw: 'none' when the computer has none, 'blocked' for anything else
 * (the system's privacy settings, or another app holding it).
 */
export function deviceFault(err) {
  const name = String(err?.name || '');
  return name === 'NotFoundError' || name === 'DevicesNotFoundError' || name === 'OverconstrainedError' ? 'none' : 'blocked';
}

const browserMedia = () => ({ getUserMedia: (c) => navigator.mediaDevices.getUserMedia(c) });

/** What a person chooses before coming into a meeting: the microphone on or off, and the camera, with its picture once it is on. */
export class Preview {
  /**
   * @param {object} [o]
   * @param {{ getUserMedia: (c: object) => Promise<object> }} [o.rtc] the browser's, or a stand-in (the checks)
   * @param {(p: Preview) => void} [o.onChange]
   */
  constructor({ rtc = null, onChange = null } = {}) {
    this.rtc = rtc || browserMedia();
    this.onChange = onChange;
    this.micOn = true;
    /** The camera's track while it is on, or null. */
    this.camera = null;
    /** 'none' or 'blocked' when the camera was asked for and could not be used. */
    this.cameraFault = null;
    /** The camera is being opened. */
    this.opening = false;
    // Each turning on or off of the camera: an opening that is no longer the latest lets its camera go.
    this.turn = 0;
  }

  changed() { this.onChange?.(this); }

  setMic(on) {
    this.micOn = Boolean(on);
    this.changed();
  }

  /** The camera on, with its picture, or off. */
  async setCamera(on) {
    const turn = ++this.turn;
    this.cameraFault = null;
    if (!on) {
      this.stop();
      this.opening = false;
      this.changed();
      return;
    }
    if (this.camera) return;
    this.opening = true;
    this.changed();
    let track = null;
    let fault = null;
    try {
      track = (await this.rtc.getUserMedia({ video: { width: { ideal: 960 }, height: { ideal: 540 } }, audio: false })).getVideoTracks()[0] || null;
      if (!track) fault = 'none';
    } catch (err) {
      fault = deviceFault(err);
    }
    // Turned off, or the dialog closed, while it was opening.
    if (turn !== this.turn) { try { track?.stop(); } catch { /* stopped */ } return; }
    this.opening = false;
    this.camera = track;
    this.cameraFault = fault;
    this.changed();
  }

  stop() {
    try { this.camera?.stop(); } catch { /* stopped */ }
    this.camera = null;
  }

  /**
   * What the meeting is joined with: { video, muted, camera }. The camera
   * goes with it, and is the meeting's to stop from then on.
   */
  take() {
    this.turn += 1;
    const camera = this.camera;
    this.camera = null;
    this.opening = false;
    return { video: Boolean(camera), muted: !this.micOn, camera };
  }

  /** The dialog closed without joining: the camera is let go. */
  close() {
    this.turn += 1;
    this.opening = false;
    this.stop();
  }
}
