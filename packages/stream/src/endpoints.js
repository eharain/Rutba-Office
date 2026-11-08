// Where each platform takes a stream, and an RTMP address taken apart.
//
// A person pastes their stream key. YouTube, Facebook and Twitch have one
// address known here and can be changed (a platform's backup server, a
// region's); Facebook takes RTMPS only, YouTube and Twitch take both, and the
// encrypted one is chosen. The rest give the address beside the key, on their
// own pages, so none has an address filled in here that would need keeping
// right (plan A11). `mark` and `colour` draw its icon; `tall` means the
// platform is made for a tall picture, so Tall is the one suggested; `maxKbps`
// is a ceiling the platform states on its own help page.

export const PLATFORMS = {
  youtube: { name: 'YouTube', url: 'rtmps://a.rtmps.youtube.com:443/live2', keyHelp: 'YouTube Studio, Create, Go live, Stream: the stream key', mark: 'Yt', colour: '#c8372d' },
  facebook: { name: 'Facebook', url: 'rtmps://live-api-s.facebook.com:443/rtmp/', keyHelp: 'Facebook, Live, Go live, Streaming software: the stream key', mark: 'Fb', colour: '#1d5bbf' },
  twitch: { name: 'Twitch', url: 'rtmps://ingest.global-contribute.live-video.net:443/app', keyHelp: 'Twitch Creator Dashboard, Settings, Stream: the stream key', mark: 'Tw', colour: '#6b4fd1' },
  tiktok: { name: 'TikTok', url: '', keyHelp: 'The server address and the stream key TikTok gives you', mark: 'Tk', colour: '#16181c', tall: true },
  instagram: { name: 'Instagram', url: '', keyHelp: 'instagram.com, Add post, Live: the address and the stream key', mark: 'In', colour: '#a8336f', tall: true },
  linkedin: { name: 'LinkedIn', url: '', keyHelp: 'LinkedIn Live Studio, Manage streams, Prepare to go live, Get URL', mark: 'Li', colour: '#0e7fa8' },
  x: { name: 'X', url: '', keyHelp: 'X Live Studio: create a source, then copy its address and stream key', mark: 'X', colour: '#2a2e36' },
  kick: { name: 'Kick', url: '', keyHelp: 'Kick creator dashboard, Channel, Stream URL and Key', mark: 'Ki', colour: '#0f8a4a', maxKbps: 8000 },
  rumble: { name: 'Rumble', url: '', keyHelp: 'The server address and the stream key Rumble gives you', mark: 'Ru', colour: '#5b8a2b' },
  custom: { name: 'RTMP server', url: '', keyHelp: 'The server address and the stream key your service gives you', mark: 'RT', colour: '#545a65' },
};

/** The sizes a stream is sent at, and the video bitrate each wants, in kbit/s (the platforms' own advice). */
export const LIVE_BITRATES = {
  '720-30': 4000,
  '720-60': 6000,
  '1080-30': 6000,
  '1080-60': 9000,
};

/** The video bitrate for a size and a frame rate. */
export function liveBitrate(size, fps) {
  return LIVE_BITRATES[`${size}-${fps >= 50 ? 60 : 30}`] || 4000;
}

/**
 * An RTMP address taken apart: { secure, host, port, app, tcUrl }, or null
 * for one that is not an RTMP address. The app is everything after the host,
 * as the servers want it in connect.
 */
export function parseRtmpUrl(url) {
  const m = /^(rtmps?):\/\/([^/:?#]+)(?::(\d+))?(\/[^?#]*)?/i.exec(String(url || '').trim());
  if (!m) return null;
  const secure = m[1].toLowerCase() === 'rtmps';
  const port = m[3] ? Number(m[3]) : secure ? 443 : 1935;
  const app = (m[4] || '/').replace(/^\/+/, '').replace(/\/+$/, '');
  if (!app) return null;
  return { secure, host: m[2], port, app, tcUrl: `${m[1].toLowerCase()}://${m[2]}${m[3] ? `:${m[3]}` : ''}/${app}` };
}

/** A key shown without giving it away: its last four characters. */
export function maskKey(key) {
  const k = String(key || '');
  return k.length > 4 ? `••••${k.slice(-4)}` : k ? '••••' : '';
}
