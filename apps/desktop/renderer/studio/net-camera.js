// The window's half of a network camera source (main/net-cameras.js reads the
// camera): a canvas per camera, painted from the H.264 frames the main process
// sends, which the program draws from as it draws a web page or a camera.
//
// The frames are decoded here by the browser's own decoder (WebCodecs), which
// uses the computer's hardware where it can. They come as Annex B and every
// key frame carries its parameter sets, so a decoder can be started at any key
// frame: on the first one, after the camera comes back, after this window
// fell behind. Nothing is shown from a frame that leans on one never decoded.
//
// A camera is held open only for a source that is started in the scene this
// window has open; when the last such source goes, the camera is let go and
// the main process stops reading it.
//
// This file takes nothing from React or the labels: what a camera's trouble
// is called is net-camera-words.js.

/** How often the list's small picture of a live camera is made again, in milliseconds. */
const THUMB_MS = 5000;
/** Frames waiting for the decoder beyond which it has fallen behind and starts again at the next key frame. */
const MAX_WAITING = 30;

export class NetCameras {
  /**
   * @param {{ shell, onChange?: () => void, makeCanvas?: () => HTMLCanvasElement, Decoder?: typeof VideoDecoder, Chunk?: typeof EncodedVideoChunk, now?: () => number }} o
   * The last four are for the tests, which have no browser.
   */
  constructor({ shell, onChange = null, makeCanvas = () => document.createElement('canvas'), Decoder = globalThis.VideoDecoder, Chunk = globalThis.EncodedVideoChunk, now = () => performance.now() }) {
    this.shell = shell;
    this.onChange = onChange;
    this.makeCanvas = makeCanvas;
    this.Decoder = Decoder;
    this.Chunk = Chunk;
    this.now = now;
    // camera id -> { id, canvas, ctx, decoder, codec, waitKey, painted, thumbAt, told, trouble, failures, state }
    this.cams = new Map();
    this.offFrame = shell.on('camera:frame', (f) => this.frame(f));
    this.offState = shell.on('camera:state', (s) => this.state(s));
  }

  changed() { this.onChange?.(); }

  /** Opens the cameras the started items need and lets go of the rest. */
  sync(items) {
    const want = new Set();
    for (const it of items) if (it.kind === 'netcam' && it.live && it.ref) want.add(it.ref);
    let moved = false;
    for (const [id, cam] of this.cams) {
      if (want.has(id)) continue;
      this.drop(cam);
      moved = true;
    }
    for (const id of want) {
      if (this.cams.has(id)) continue;
      const cam = { id, canvas: null, ctx: null, decoder: null, codec: '', waitKey: true, painted: false, thumbAt: 0, told: false, trouble: '', failures: 0, state: { status: 'connecting', code: '' } };
      this.cams.set(id, cam);
      moved = true;
      this.shell.capture.cameraOpen({ id }).then((s) => {
        // What the camera has said since is newer than what opening it answered.
        if (this.cams.get(id) !== cam || cam.told || !s) return;
        cam.state = { status: s.status, code: s.code || '' };
        this.changed();
      }, () => {
        if (this.cams.get(id) !== cam) return;
        cam.state = { status: 'failed', code: 'closed' };
        this.changed();
      });
    }
    if (moved) this.changed();
  }

  drop(cam) {
    this.cams.delete(cam.id);
    this.closeDecoder(cam);
    this.shell.capture.cameraClose({ id: cam.id }).catch(() => {});
  }

  closeDecoder(cam) {
    const d = cam.decoder;
    cam.decoder = null;
    cam.codec = '';
    cam.waitKey = true;
    try {
      if (d && d.state !== 'closed') d.close();
    } catch {
      // Already closed.
    }
  }

  /** The decoder could not be made or would not take the picture: said once a camera has shown nothing at all. */
  failed(cam) {
    cam.failures += 1;
    if (cam.painted || cam.trouble) return;
    cam.trouble = 'decode';
    this.changed();
  }

  configure(cam, codec) {
    this.closeDecoder(cam);
    if (!this.Decoder || !this.Chunk) {
      this.failed(cam);
      return;
    }
    let decoder;
    try {
      decoder = new this.Decoder({
        output: (frame) => {
          if (cam.decoder !== decoder || this.cams.get(cam.id) !== cam) {
            frame.close();
            return;
          }
          this.paint(cam, frame);
        },
        error: () => {
          if (cam.decoder !== decoder) return;
          this.closeDecoder(cam);
          this.failed(cam);
        },
      });
      decoder.configure({ codec, optimizeForLatency: true });
    } catch {
      this.failed(cam);
      return;
    }
    cam.decoder = decoder;
    cam.codec = codec;
    cam.waitKey = true;
  }

  frame({ id, key, codec, timestamp, bytes }) {
    const cam = this.cams.get(id);
    if (!cam || !bytes?.length) return;
    if (key && codec && (!cam.decoder || cam.decoder.state !== 'configured' || cam.codec !== codec)) this.configure(cam, codec);
    const decoder = cam.decoder;
    if (!decoder || decoder.state !== 'configured') return;
    if (cam.waitKey && !key) return;
    if (decoder.decodeQueueSize > MAX_WAITING) {
      // Behind: what is waiting is let go, and the picture starts again at the next key frame.
      this.closeDecoder(cam);
      return;
    }
    cam.waitKey = false;
    try {
      decoder.decode(new this.Chunk({ type: key ? 'key' : 'delta', timestamp, data: bytes }));
    } catch {
      this.closeDecoder(cam);
      this.failed(cam);
    }
  }

  paint(cam, frame) {
    const w = frame.displayWidth;
    const h = frame.displayHeight;
    if (!w || !h) {
      frame.close();
      return;
    }
    if (!cam.canvas) {
      cam.canvas = this.makeCanvas();
      cam.ctx = cam.canvas.getContext('2d');
    }
    if (cam.canvas.width !== w || cam.canvas.height !== h) {
      cam.canvas.width = w;
      cam.canvas.height = h;
    }
    cam.ctx.drawImage(frame, 0, 0, w, h);
    frame.close();
    cam.failures = 0;
    const at = this.now();
    // The list's small picture is made when the window draws: on the first frame, then now and again.
    if (!cam.painted || cam.trouble || at - cam.thumbAt > THUMB_MS) {
      cam.painted = true;
      cam.trouble = '';
      cam.thumbAt = at;
      cam.canvas.dataUrl = null;
      this.changed();
    }
  }

  state({ id, status, code }) {
    const cam = this.cams.get(id);
    if (!cam) return;
    cam.told = true;
    cam.state = { status, code: code || '' };
    // Read afresh, the camera starts at a key frame, and so does a new decoder; its sign-in or address may have been put right.
    if (status === 'connecting') {
      this.closeDecoder(cam);
      cam.trouble = '';
      cam.failures = 0;
    }
    this.changed();
  }

  /** The canvas an item is drawn from: the camera's last picture, or null before its first. */
  drawable(item) {
    const cam = this.cams.get(item.ref);
    return cam?.painted ? cam.canvas : null;
  }

  /** connecting, live, waiting (dropped, being tried again), failed, or off for an item stopped. */
  statusOf(item) {
    const cam = this.cams.get(item.ref);
    if (!cam) return 'off';
    return cam.trouble ? 'failed' : cam.state.status;
  }

  /** Why a camera is waiting or failed: the reader's code, 'decode' when this computer could not show its picture, or ''. */
  codeOf(item) {
    const cam = this.cams.get(item.ref);
    if (!cam) return '';
    return cam.trouble || (cam.state.status === 'waiting' || cam.state.status === 'failed' ? cam.state.code || 'closed' : '');
  }

  /** The source list's word for it: live, opening, error or off. */
  rowState(item) {
    const cam = this.cams.get(item.ref);
    if (!cam) return 'off';
    const status = this.statusOf(item);
    if (status === 'waiting' || status === 'failed') return 'error';
    return status === 'live' && cam.painted ? 'live' : 'opening';
  }

  dispose() {
    this.offFrame?.();
    this.offState?.();
    for (const cam of [...this.cams.values()]) this.drop(cam);
  }
}
