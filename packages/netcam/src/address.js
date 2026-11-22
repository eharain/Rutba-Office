// A camera's address, as a person types it and as the reader uses it.
//
// People paste what the camera's manual or its app gives them, and that is
// often the address with the sign-in inside it (rtsp://admin:secret@10.0.0.9/
// stream1). The sign-in is taken out here: the address that is kept, shown
// and written into a setup never holds a password, and the user name and
// password travel beside it (the password to the keystore).

const PORTS = { 'rtsp:': 554, 'rtsps:': 322 };

const decoded = (s) => {
  try {
    return decodeURIComponent(s);
  } catch {
    return s;
  }
};

/**
 * What was typed, as an address the reader takes: { url, secure, host, port,
 * path, user, password }, or null when it is not a camera's address. `url`
 * is whole and carries no sign-in. An address typed without rtsp:// is taken
 * as one.
 */
export function parseCameraAddress(text) {
  let typed = String(text ?? '').trim();
  if (!typed || /\s/.test(typed) || typed.length > 2000) return null;
  if (!/^[a-z][a-z0-9+.-]*:\/\//i.test(typed)) typed = `rtsp://${typed}`;
  let u;
  try {
    u = new URL(typed);
  } catch {
    return null;
  }
  const scheme = u.protocol.toLowerCase();
  if (!PORTS[scheme] || !u.hostname) return null;
  const port = u.port ? Number(u.port) : PORTS[scheme];
  if (!Number.isInteger(port) || port < 1 || port > 65535) return null;
  const path = `${u.pathname || '/'}${u.search || ''}`;
  const shownPort = port === PORTS[scheme] ? '' : `:${port}`;
  return {
    url: `${scheme}//${u.hostname}${shownPort}${path}`,
    secure: scheme === 'rtsps:',
    host: u.hostname.replace(/^\[|\]$/g, ''),
    port,
    path,
    user: decoded(u.username),
    password: decoded(u.password),
  };
}

/** The camera's host as a name for it, when the person gives none: "10.0.0.9" or "porch.example". */
export function cameraNameFrom(url) {
  return parseCameraAddress(url)?.host || '';
}

/** Whether two addresses are the same camera stream. */
export function sameCamera(a, b) {
  const x = parseCameraAddress(a);
  const y = parseCameraAddress(b);
  return Boolean(x && y && x.url === y.url);
}
