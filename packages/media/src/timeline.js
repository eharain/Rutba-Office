// The edit timeline.
//
// A video edit is a list of clips taken from sources, not a re-encoded file.
// Trimming changes two numbers; splitting makes two clips out of one; nothing
// is written until an export asks for it. That is what lets a two-hour
// recording be trimmed instantly on a laptop.
//
// No DOM and no codecs here — this is the arithmetic of an edit, which is where
// the off-by-one errors live and where they can be tested.

let seq = 0;

export function clip(source, { start = 0, end = null, speed = 1, volume = 1, name = '' } = {}) {
  const duration = end == null ? source.duration : end;
  return {
    id: `c${++seq}`,
    sourceId: source.id,
    name: name || source.name,
    start: Math.max(0, start),
    end: Math.min(source.duration, Math.max(start, duration)),
    speed,
    volume,
  };
}

/** How long a clip lasts once its speed is taken into account. */
export const clipDuration = (c) => Math.max(0, (c.end - c.start) / (c.speed || 1));

export class Timeline {
  constructor(sources = [], clips = []) {
    this.sources = sources;
    this.clips = clips;
  }

  static of(source) {
    return new Timeline([source], [clip(source)]);
  }

  get duration() {
    return this.clips.reduce((n, c) => n + clipDuration(c), 0);
  }

  sourceOf(c) {
    return this.sources.find((s) => s.id === c.sourceId) || null;
  }

  /** Timeline position → which clip, and where inside its source. */
  locate(time) {
    let at = 0;
    for (const c of this.clips) {
      const length = clipDuration(c);
      if (time < at + length || c === this.clips[this.clips.length - 1]) {
        return { clip: c, offset: time - at, sourceTime: c.start + (time - at) * (c.speed || 1) };
      }
      at += length;
    }
    return null;
  }

  /** Where a clip begins on the timeline. */
  startOf(clipId) {
    let at = 0;
    for (const c of this.clips) {
      if (c.id === clipId) return at;
      at += clipDuration(c);
    }
    return 0;
  }

  /** Cut a clip in two at a timeline position. */
  split(time) {
    const found = this.locate(time);
    if (!found || found.offset <= 0.02 || found.offset >= clipDuration(found.clip) - 0.02) return this;
    const index = this.clips.indexOf(found.clip);
    const left = { ...found.clip, id: `c${++seq}`, end: found.sourceTime };
    const right = { ...found.clip, id: `c${++seq}`, start: found.sourceTime };
    const clips = [...this.clips];
    clips.splice(index, 1, left, right);
    return new Timeline(this.sources, clips);
  }

  trim(clipId, { start, end }) {
    return new Timeline(
      this.sources,
      this.clips.map((c) => {
        if (c.id !== clipId) return c;
        const source = this.sourceOf(c);
        const nextStart = start == null ? c.start : Math.max(0, Math.min(start, (end ?? c.end) - 0.05));
        const nextEnd = end == null ? c.end : Math.min(source?.duration ?? end, Math.max(end, nextStart + 0.05));
        return { ...c, start: nextStart, end: nextEnd };
      })
    );
  }

  remove(clipId) {
    return new Timeline(this.sources, this.clips.filter((c) => c.id !== clipId));
  }

  move(clipId, toIndex) {
    const from = this.clips.findIndex((c) => c.id === clipId);
    if (from < 0) return this;
    const clips = [...this.clips];
    const [moved] = clips.splice(from, 1);
    clips.splice(Math.max(0, Math.min(toIndex, clips.length)), 0, moved);
    return new Timeline(this.sources, clips);
  }

  set(clipId, patch) {
    return new Timeline(this.sources, this.clips.map((c) => (c.id === clipId ? { ...c, ...patch } : c)));
  }

  append(source) {
    return new Timeline(
      this.sources.some((s) => s.id === source.id) ? this.sources : [...this.sources, source],
      [...this.clips, clip(source)]
    );
  }

  /** Clips with their timeline positions, which is what a track draws. */
  layout() {
    let at = 0;
    return this.clips.map((c) => {
      const length = clipDuration(c);
      const row = { ...c, from: at, to: at + length, length };
      at += length;
      return row;
    });
  }
}

/** mm:ss.t — the way a video tool writes a position. */
export function timecode(seconds, { frames = 0 } = {}) {
  if (!Number.isFinite(seconds) || seconds < 0) return '0:00';
  const whole = Math.floor(seconds);
  const h = Math.floor(whole / 3600);
  const m = Math.floor((whole % 3600) / 60);
  const s = whole % 60;
  const base = h ? `${h}:${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}` : `${m}:${String(s).padStart(2, '0')}`;
  if (!frames) return `${base}.${Math.floor((seconds - whole) * 10)}`;
  return `${base}:${String(Math.floor((seconds - whole) * frames)).padStart(2, '0')}`;
}
