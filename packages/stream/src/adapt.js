// Go Live keeps up with the upload (owner, 2026-10-20: "how do the system
// cater for slower connections and drop frames of voice and videos and lower
// the quality to keep it upto speed").
//
// Every few seconds the destinations say how they are doing: frames dropped
// since the last look, and how full their line's queue is. One that dropped
// frames, or is more than half full, brings the picture's bitrate down a
// quarter, not below a floor. When every destination has kept up for a while,
// it goes back up a tenth at a time, never above what was asked for. The sound
// is never touched: a stream is still watchable on a poor picture, not on a
// broken voice. One encoder feeds every destination, so the slowest sets it.

export class BitrateGovernor {
  /**
   * @param {object} o
   * @param {number} o.target the bitrate asked for, kb/s
   * @param {number} [o.floor] the lowest it goes, kb/s (a quarter of the target, at least 300)
   * @param {number} [o.recoverMs] how long every destination keeps up before it goes up again
   */
  constructor({ target, floor = Math.max(300, Math.round(target * 0.25)), recoverMs = 10000 }) {
    this.target = target;
    this.floor = Math.min(floor, target);
    this.recoverMs = recoverMs;
    this.kbps = target;
    this.calmSince = null;
  }

  /**
   * One look at the destinations: [{ dropped, queued }] where `dropped` is
   * frames dropped since the last look and `queued` how full the queue is
   * (0 to 1). Answers the bitrate to encode at now.
   */
  tick(destinations, now = Date.now()) {
    const struggling = destinations.some((d) => (d.dropped || 0) > 0 || (d.queued || 0) > 0.5);
    if (struggling) {
      this.kbps = Math.max(this.floor, Math.round(this.kbps * 0.75));
      this.calmSince = now;
      return this.kbps;
    }
    if (this.calmSince == null) this.calmSince = now;
    if (this.kbps < this.target && now - this.calmSince >= this.recoverMs) {
      this.kbps = Math.min(this.target, Math.round(this.kbps * 1.1));
      this.calmSince = now;
    }
    return this.kbps;
  }

  /** Whether the picture is below what was asked for, and so says so to the person streaming. */
  get lowered() {
    return this.kbps < this.target;
  }
}
