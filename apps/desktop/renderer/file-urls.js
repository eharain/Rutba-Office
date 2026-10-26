// The addresses the windows show a file and its thumbnail by: rutba://file and
// rutba://thumb, the path in base64url. Pictures shows its library with them,
// and Worksheets the picture it reads data or a bill from.

const b64url = (s) => {
  const bytes = new TextEncoder().encode(s);
  let bin = '';
  for (const b of bytes) bin += String.fromCharCode(b);
  return btoa(bin).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
};

/** The file itself, decoded by Chromium straight from disk. */
export const fileUrl = (p) => `rutba://file/${b64url(p)}`;

/** The file's thumbnail, made once by the platform and kept. */
export const thumbUrl = (p, size = 256) => `rutba://thumb/${b64url(p)}?s=${size}`;
