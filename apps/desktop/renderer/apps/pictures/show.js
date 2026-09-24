// The slideshow's arithmetic: how far the picture on screen is from the
// next one, how a length reads on a tile, and how long the show waits on
// what it is showing before it moves on. No window in any of it, which is
// what lets it be tested without opening one.

/**
 * The index the show lands on `delta` steps from `index`, looping either
 * way round a folder of `count` items. `-1` when there is nothing to show.
 */
export function nextIndex(count, index, delta) {
  if (!count || count < 1) return -1;
  return ((index + delta) % count + count) % count;
}

/**
 * A length in seconds, the way a clip's tile badge and the details panel
 * read one: `m:ss`, or `h:mm:ss` once it runs past an hour. `--:--` for
 * anything that is not a real, finite length — a clip whose duration has
 * not been read yet, or one nothing could make sense of.
 */
export function formatLength(seconds) {
  if (!Number.isFinite(seconds) || seconds < 0) return '--:--';
  const whole = Math.floor(seconds);
  const h = Math.floor(whole / 3600);
  const m = Math.floor((whole % 3600) / 60);
  const s = whole % 60;
  const mm = h ? String(m).padStart(2, '0') : String(m);
  const ss = String(s).padStart(2, '0');
  return h ? `${h}:${mm}:${ss}` : `${mm}:${ss}`;
}

/**
 * How long, in seconds, the show waits on `item` before moving on. A
 * picture gets the chosen `interval`; a clip is left to play — its own
 * `duration` when that is known and shorter than a minute, a minute
 * otherwise. In the window this is the ceiling under the clip's own
 * `ended` event, which usually moves the show on first.
 */
export function advanceAfter(item, interval) {
  const isClip = item?.kind === 'video' || item?.kind === 'audio';
  if (!isClip) return interval;
  const duration = item?.duration;
  return Number.isFinite(duration) && duration > 0 ? Math.min(duration, 60) : 60;
}
