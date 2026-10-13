// Pixels on their way to the recogniser: any picture made grey, and the
// grey picture written as the file the engine reads.
//
// A picture is { width, height, data } with one, two, three or four
// channels a pixel (grey, grey and alpha, RGB, RGBA), eight bits each, as
// a canvas's ImageData or decodePng gives it. Transparency is read over
// white paper.

/** The channels a pixel of this picture carries. */
export function channelsOf({ width, height, data, channels }) {
  if (channels) return channels;
  const n = data.length / (width * height);
  if (![1, 2, 3, 4].includes(n)) throw new Error(`A ${width} by ${height} picture cannot hold ${data.length} bytes.`);
  return n;
}

/** The picture as one byte a pixel, 0 black to 255 white (ITU-R BT.601 weights). */
export function toGrey(image) {
  const { width, height, data } = image;
  const channels = channelsOf(image);
  if (channels === 1) return { width, height, data: data instanceof Uint8Array ? data : Uint8Array.from(data) };
  const out = new Uint8Array(width * height);
  for (let i = 0, o = 0; o < out.length; i += channels, o++) {
    let g = channels >= 3 ? (data[i] * 299 + data[i + 1] * 587 + data[i + 2] * 114 + 500) / 1000 : data[i];
    const a = channels === 2 ? data[i + 1] : channels === 4 ? data[i + 3] : 255;
    if (a < 255) g = (g * a + 255 * (255 - a)) / 255;
    out[o] = g;
  }
  return { width, height, data: out };
}

/** A grey picture as a binary PGM (P5), which the engine's image reader takes as it is. */
export function pgm({ width, height, data }) {
  const head = Buffer.from(`P5\n${width} ${height}\n255\n`, 'latin1');
  const out = new Uint8Array(head.length + width * height);
  out.set(head, 0);
  out.set(data.subarray(0, width * height), head.length);
  return out;
}
