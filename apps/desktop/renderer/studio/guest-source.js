// Guests from Talk as sources of the studio (Office 2.27.0, the Talk, Meet
// and Live plan's phase 3): each person who sits in a seat gives a picture a
// scene can place and a sound the mix can carry.
//
// The pictures and sounds come over the meeting the studio holds (Talk's
// group call, between the computers); this file only draws and mixes them.
// Who sits in which seat is told from outside (guests.js) and kept here for
// as long as the window is open, never in the stack.
//
// Sound: the person at this computer always hears every guest, on air or
// not, so they can talk before a guest is shown. A guest's sound goes into
// what is recorded and streamed only while their seat is shown in the scene
// on air and Sound is on. What this computer sends back into the meeting is
// the programme's picture and its sound without the guests', who hear each
// other directly, so nobody hears themselves late.
//
// A guest with no camera is drawn as a card with their name; a seat nobody
// sits in draws nothing.

const CARD_W = 640;
const CARD_H = 360;
const FONT = '"Segoe UI", "Segoe UI Variable", system-ui, sans-serif';

const initials = (name) => String(name || '').trim().split(/\s+/).slice(0, 2).map((w) => w[0] || '').join('').toUpperCase() || '?';

export class GuestSources {
  /**
   * @param {object} o
   * @param {() => AudioContext} o.audio  the studio's one sound graph, made when first needed
   * @param {() => AudioNode} o.local  everything this computer makes itself: the microphones, the computer's sound, music, spoken lines
   * @param {() => AudioNode} o.air  what is recorded and streamed
   * @param {() => void} [o.onChange]
   * @param {Document} [o.doc]
   */
  constructor({ audio, local, air, onChange = null, doc = globalThis.document }) {
    Object.assign(this, { audio, local, air, onChange, doc });
    /** { [seat]: { id, name } } */
    this.seats = {};
    /** person id -> the person as the meeting has them: { id, name, stream, connected, camera, sharing, muted } */
    this.people = new Map();
    /** person id -> { stream, video, track, src, hear, air } */
    this.held = new Map();
    /** The seats shown in the scene on air, and the Sound switch there. */
    this.shown = new Set();
    this.sound = 'on';
    this.cards = new Map();
    this.back = null;
    this.disposed = false;
  }

  changed() { if (!this.disposed) this.onChange?.(); }

  /** Who sits where, and the people as the meeting has them now. */
  set(seats = {}, people = []) {
    if (this.disposed) return;
    this.seats = seats || {};
    this.people = new Map((people || []).filter((p) => p?.id).map((p) => [p.id, p]));
    const seated = new Set(Object.values(this.seats).map((w) => w.id));
    for (const [id, h] of [...this.held]) {
      const p = this.people.get(id);
      if (!seated.has(id) || !p?.stream || p.stream !== h.stream) this.drop(id);
    }
    for (const id of seated) {
      const p = this.people.get(id);
      if (p?.stream) this.hold(p);
    }
    this.route();
  }

  /** What is on air: the seats its scene shows, and whether its Sound is on. */
  sync(stack) {
    this.shown = new Set((stack?.items || []).filter((it) => it.kind === 'guest' && it.on).map((it) => it.ref));
    this.sound = stack?.sound || 'on';
    this.route();
  }

  seatOf(id) { return Object.keys(this.seats).find((ref) => this.seats[ref]?.id === id) || null; }
  personOf(item) { return this.seats[item?.ref] || null; }
  nameOf(item) { return this.personOf(item)?.name || ''; }
  /** Whether a person's sound is going into what is recorded and streamed just now. */
  onAir(id) { return this.sound === 'on' && this.shown.has(this.seatOf(id)); }

  hold(p) {
    let h = this.held.get(p.id);
    if (!h) {
      h = { stream: p.stream, video: this.play(p.stream), track: null, src: null, hear: null, air: null };
      this.held.set(p.id, h);
    }
    const track = p.stream.getAudioTracks?.()[0] || null;
    if (track === h.track) return;
    this.unwire(h);
    if (track) this.wire(h, track);
  }

  wire(h, track) {
    const ctx = this.audio();
    if (!ctx) return;
    h.track = track;
    h.src = ctx.createMediaStreamSource(new MediaStream([track]));
    // To this computer's speakers, always.
    h.hear = ctx.createGain();
    h.src.connect(h.hear);
    h.hear.connect(ctx.destination);
    // Into what is recorded and streamed, closed until the seat is shown on air.
    h.air = ctx.createGain();
    h.air.gain.value = 0;
    h.src.connect(h.air);
    h.air.connect(this.air());
  }

  unwire(h) {
    for (const n of [h.src, h.hear, h.air]) { try { n?.disconnect(); } catch { /* already */ } }
    h.track = null;
    h.src = null;
    h.hear = null;
    h.air = null;
  }

  route() {
    for (const [id, h] of this.held) if (h.air) h.air.gain.value = this.onAir(id) ? 1 : 0;
  }

  /** Whether a guest's sound is in what is recorded and streamed just now. */
  heard() {
    for (const [id, h] of this.held) if (h.src && this.onAir(id) && !this.people.get(id)?.muted) return true;
    return false;
  }

  /** A stream playing in a video kept in the page, unseen and silent here: its sound goes through the graph. */
  play(stream) {
    const doc = this.doc;
    if (!doc?.createElement) return null;
    let shelf = doc.getElementById?.('st-sources-shelf');
    if (!shelf) {
      shelf = doc.createElement('div');
      shelf.id = 'st-sources-shelf';
      shelf.setAttribute?.('aria-hidden', 'true');
      if (shelf.style) shelf.style.cssText = 'position:fixed;left:0;top:0;width:1px;height:1px;overflow:hidden;opacity:0;pointer-events:none;';
      doc.body?.appendChild?.(shelf);
    }
    const v = doc.createElement('video');
    v.muted = true;
    v.playsInline = true;
    v.autoplay = true;
    if (v.style) v.style.cssText = 'width:1px;height:1px;';
    v.srcObject = stream;
    shelf.appendChild?.(v);
    v.play?.()?.catch?.(() => {});
    v.addEventListener?.('loadeddata', () => this.changed());
    v.addEventListener?.('resize', () => this.changed());
    return v;
  }

  drop(id) {
    const h = this.held.get(id);
    if (!h) return;
    this.held.delete(id);
    this.unwire(h);
    if (h.video) { h.video.srcObject = null; h.video.remove?.(); }
  }

  /** Whether a seat's person is giving a picture just now. */
  picture(item) {
    const who = this.personOf(item);
    if (!who) return null;
    const p = this.people.get(who.id);
    const v = this.held.get(who.id)?.video;
    return p && (p.camera || p.sharing) && v?.videoWidth ? v : null;
  }

  /** What a seat is drawn from: its person's picture, their name on a card when they give none, nothing when nobody sits in it. */
  drawable(item) {
    const who = this.personOf(item);
    if (!who) return null;
    return this.picture(item) || this.card(who.name);
  }

  /** The live stream of a seat's picture, for its preview in the list, or null. */
  streamOf(item) { return this.picture(item) ? this.held.get(this.personOf(item).id).stream : null; }

  /** 'off' with nobody in the seat, 'opening' while its person is still connecting, 'live' after. */
  stateOf(item) {
    const who = this.personOf(item);
    if (!who) return 'off';
    return this.people.get(who.id)?.connected ? 'live' : 'opening';
  }

  card(name) {
    const key = String(name || '');
    if (this.cards.has(key)) return this.cards.get(key);
    const c = this.doc?.createElement?.('canvas');
    if (!c) return null;
    c.width = CARD_W;
    c.height = CARD_H;
    const g = c.getContext?.('2d');
    if (g) {
      g.fillStyle = '#1f2933';
      g.fillRect(0, 0, CARD_W, CARD_H);
      g.fillStyle = '#3e4c59';
      g.beginPath();
      g.arc(CARD_W / 2, CARD_H * 0.42, CARD_H * 0.22, 0, Math.PI * 2);
      g.fill();
      g.fillStyle = '#fff';
      g.textAlign = 'center';
      g.textBaseline = 'middle';
      g.font = `600 ${Math.round(CARD_H * 0.17)}px ${FONT}`;
      g.fillText(initials(key), CARD_W / 2, CARD_H * 0.43);
      g.font = `600 ${Math.round(CARD_H * 0.085)}px ${FONT}`;
      g.fillText(key, CARD_W / 2, CARD_H * 0.8, CARD_W * 0.9);
    }
    this.cards.set(key, c);
    return c;
  }

  /**
   * What this computer's seat in the meeting sends: the programme's picture,
   * and its sound without the guests'. Fresh tracks each time; whoever takes
   * them stops them.
   */
  feed(canvas, { fps = 15 } = {}) {
    let audio = null;
    try {
      const ctx = this.audio();
      if (!this.back) {
        this.back = ctx.createMediaStreamDestination();
        this.local().connect(this.back);
      }
      audio = this.back.stream.getAudioTracks()[0]?.clone() || null;
    } catch { /* no sound to send: the picture still goes */ }
    let video = null;
    try { video = canvas?.captureStream?.(fps)?.getVideoTracks()[0] || null; } catch { /* no picture to send */ }
    return { audio, video };
  }

  dispose() {
    this.disposed = true;
    for (const id of [...this.held.keys()]) this.drop(id);
    this.cards.clear();
    this.seats = {};
    this.people.clear();
    this.back = null;
  }
}
