// The icon set.
//
// Drawn here rather than fetched: an office suite that will not open a document
// until a font has downloaded is not one anybody can rely on, and the whole
// point of this build is that it works with the network off.
//
// One grid (24), one stroke weight, one join style — which is what makes a set
// of icons look like a set rather than a collection.

import React from 'react';

const P = (d, extra) => ({ d, ...extra });

const PATHS = {
  // apps
  mail: [P('M3 7a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2v10a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z'), P('m3.5 7.5 8.5 6 8.5-6')],
  calendar: [P('M4 6a2 2 0 0 1 2-2h12a2 2 0 0 1 2 2v13a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2z'), P('M4 10h16'), P('M8 2v4M16 2v4'), P('M8 14h3M13 14h3M8 17h3')],
  contacts: [P('M5 4a1 1 0 0 1 1-1h12a1 1 0 0 1 1 1v16a1 1 0 0 1-1 1H6a1 1 0 0 1-1-1z'), P('M12 12a3 3 0 1 0 0-6 3 3 0 0 0 0 6z'), P('M7 18c1-2.5 3-3.5 5-3.5s4 1 5 3.5')],
  word: [P('M6 3h8l4 4v14a1 1 0 0 1-1 1H6a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1z'), P('M14 3v5h4'), P('M8 13h7M8 17h5')],
  sheets: [P('M4 5a1 1 0 0 1 1-1h14a1 1 0 0 1 1 1v14a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1z'), P('M4 9h16M4 14h16M10 4v16M15 4v16')],
  slides: [P('M3 5a1 1 0 0 1 1-1h16a1 1 0 0 1 1 1v9a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1z'), P('M12 15v5M8 20h8')],
  pictures: [P('M4 5a1 1 0 0 1 1-1h14a1 1 0 0 1 1 1v14a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1z'), P('m5 16 4-4 3 3 3.5-3.5L20 16'), P('M15 8.5h.01', { strokeWidth: 2.4 })],
  image: [P('M4 5a1 1 0 0 1 1-1h14a1 1 0 0 1 1 1v14a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1z'), P('M9 4v16M4 9h16'), P('m14 15 3-3 3 3')],
  video: [P('M3 7a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v10a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z'), P('m16 10 5-3v10l-5-3z')],
  home: [P('m4 11 8-7 8 7'), P('M6 10v9a1 1 0 0 0 1 1h10a1 1 0 0 0 1-1v-9')],

  // file and document
  file: [P('M6 3h8l4 4v14a1 1 0 0 1-1 1H6a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1z'), P('M14 3v5h4')],
  folder: [P('M3 7a1 1 0 0 1 1-1h5l2 2h9a1 1 0 0 1 1 1v9a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1z')],
  folderOpen: [P('M3 7a1 1 0 0 1 1-1h5l2 2h8a1 1 0 0 1 1 1v1'), P('M3 9h18l-2.2 8.3a1 1 0 0 1-1 .7H4a1 1 0 0 1-1-1z')],
  save: [P('M5 4h11l3 3v13a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1V5a1 1 0 0 1 1-1z'), P('M8 4v5h7V4M8 20v-6h8v6')],
  open: [P('M4 6a1 1 0 0 1 1-1h5l2 2h7a1 1 0 0 1 1 1v3'), P('M3 12h18l-2 7H5z')],
  new: [P('M6 3h8l4 4v14a1 1 0 0 1-1 1H6a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1z'), P('M14 3v5h4'), P('M12 11v6M9 14h6')],
  print: [P('M7 9V4h10v5'), P('M5 9h14a1 1 0 0 1 1 1v6h-4v4H8v-4H4v-6a1 1 0 0 1 1-1z')],
  pdf: [P('M6 3h8l4 4v14a1 1 0 0 1-1 1H6a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1z'), P('M14 3v5h4'), P('M8 17v-4h1.5a1.2 1.2 0 0 1 0 2.4H8M13 17v-4h1.6c1 0 1.6.8 1.6 2s-.6 2-1.6 2z')],
  export: [P('M12 3v11'), P('m8 7 4-4 4 4'), P('M5 15v4a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1v-4')],
  import: [P('M12 14V3'), P('m8 10 4 4 4-4'), P('M5 15v4a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1v-4')],

  // editing
  undo: [P('M9 10H5V6'), P('M5 10a8 8 0 1 1 2 8')],
  redo: [P('M15 10h4V6'), P('M19 10a8 8 0 1 0-2 8')],
  cut: [P('M6 4l12 14M18 4L6 18'), P('M7 20a2 2 0 1 0 0-4 2 2 0 0 0 0 4zM17 20a2 2 0 1 0 0-4 2 2 0 0 0 0 4z')],
  copy: [P('M9 9h10a1 1 0 0 1 1 1v10a1 1 0 0 1-1 1H9a1 1 0 0 1-1-1V10a1 1 0 0 1 1-1z'), P('M5 15H4a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1h10a1 1 0 0 1 1 1v1')],
  paste: [P('M9 4h6v3H9z'), P('M8 5H6a1 1 0 0 0-1 1v14a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1V6a1 1 0 0 0-1-1h-2')],
  bold: [P('M7 4h6a4 4 0 0 1 0 8H7zM7 12h7a4 4 0 0 1 0 8H7z')],
  italic: [P('M15 4h-5M14 20H9M13 4 10 20')],
  underline: [P('M7 4v7a5 5 0 0 0 10 0V4M6 20h12')],
  strike: [P('M5 12h14'), P('M8 8a3.5 3.5 0 0 1 3.5-3h1A3.5 3.5 0 0 1 16 8M8 16a3.5 3.5 0 0 0 3.5 3h1a3.5 3.5 0 0 0 3.5-3')],
  alignLeft: [P('M4 6h16M4 11h10M4 16h14M4 21h8')],
  alignCenter: [P('M4 6h16M7 11h10M5 16h14M8 21h8')],
  alignRight: [P('M4 6h16M10 11h10M6 16h14M12 21h8')],
  alignJustify: [P('M4 6h16M4 11h16M4 16h16M4 21h16')],
  listBullet: [P('M9 6h11M9 12h11M9 18h11'), P('M5 6h.01M5 12h.01M5 18h.01', { strokeWidth: 2.4 })],
  listNumber: [P('M10 6h10M10 12h10M10 18h10'), P('M4 8V4l-1 .7M3.5 12h2l-2 3h2')],
  link: [P('M10 13a4 4 0 0 0 5.7 0l2.6-2.6a4 4 0 1 0-5.7-5.7L11.4 6'), P('M14 11a4 4 0 0 0-5.7 0l-2.6 2.6a4 4 0 0 0 5.7 5.7L12.6 18')],
  table: [P('M4 5h16v14H4z'), P('M4 10h16M4 15h16M10 5v14')],
  picture: [P('M4 5h16v14H4z'), P('m5 16 4-4 3 3 3.5-3.5L20 16')],
  shape: [P('M12 3 3 20h18z')],
  textbox: [P('M4 6h16v12H4z'), P('M9 9h6M12 9v6')],
  find: [P('M11 18a7 7 0 1 0 0-14 7 7 0 0 0 0 14z'), P('m20 20-4-4')],
  formula: [P('M6 20c2 0 2.5-1 3-4l2-8c.5-3 1-4 3-4'), P('M6 8h9')],
  sum: [P('M17 5H7l6 7-6 7h10')],
  filter: [P('M4 5h16l-6 7v6l-4 2v-8z')],
  sort: [P('M7 5v14M4 16l3 3 3-3'), P('M17 19V5M14 8l3-3 3 3')],
  chart: [P('M4 20V9M10 20V4M16 20v-7M22 20H2')],
  freeze: [P('M4 5h16v14H4z'), P('M4 10h16M9 5v14', { strokeDasharray: '3 2' })],

  // media
  crop: [P('M6 2v14a1 1 0 0 0 1 1h14'), P('M2 6h14a1 1 0 0 1 1 1v14')],
  rotate: [P('M20 11a8 8 0 1 0-2 6'), P('M20 5v6h-6')],
  flip: [P('M12 3v18'), P('M8 7 4 12l4 5zM16 7l4 5-4 5z')],
  zoomIn: [P('M11 18a7 7 0 1 0 0-14 7 7 0 0 0 0 14z'), P('m20 20-4-4M11 8v6M8 11h6')],
  zoomOut: [P('M11 18a7 7 0 1 0 0-14 7 7 0 0 0 0 14z'), P('m20 20-4-4M8 11h6')],
  play: [P('m7 4 13 8-13 8z')],
  pause: [P('M8 4h3v16H8zM13 4h3v16h-3z')],
  stop: [P('M6 6h12v12H6z')],
  skipBack: [P('M18 5v14L8 12z'), P('M6 5v14')],
  skipForward: [P('M6 5v14l10-7z'), P('M18 5v14')],
  scissors: [P('m8 8 10 10M18 8 8 18'), P('M6 8a2 2 0 1 0 0-4 2 2 0 0 0 0 4zM6 20a2 2 0 1 0 0-4 2 2 0 0 0 0 4z')],
  volume: [P('M5 9v6h4l5 4V5L9 9z'), P('M17 9a4 4 0 0 1 0 6')],
  sliders: [P('M4 7h10M18 7h2M4 17h4M12 17h8'), P('M16 7a2 2 0 1 0 0-.01M10 17a2 2 0 1 0 0-.01')],
  wand: [P('m5 19 9-9'), P('M14 5V3M17.5 6.5 19 5M19 10h2M6 5 4.5 3.5M9 3v2'), P('m13 11 3-3 5 5-3 3z')],
  contrast: [P('M12 21a9 9 0 1 0 0-18 9 9 0 0 0 0 18z'), P('M12 3v18a9 9 0 0 0 0-18z', { fill: 'currentColor', stroke: 'none' })],

  // mail
  send: [P('m21 3-9 18-2.5-7.5L2 11z'), P('M21 3 9.5 13.5')],
  reply: [P('M9 8V4L2 11l7 7v-4c6 0 9 2 11 5-1-7-4-11-11-11z')],
  replyAll: [P('M12 8V4l-7 7 7 7v-4c5 0 8 2 10 5-1-7-4-11-10-11z'), P('M6 4 1 11l5 7')],
  forward: [P('M15 8V4l7 7-7 7v-4c-6 0-9 2-11 5 1-7 4-11 11-11z')],
  inbox: [P('M4 13h5l1 3h4l1-3h5'), P('M5 5h14l2 8v6a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1v-6z')],
  trash: [P('M4 7h16M10 4h4M9 7v12M15 7v12'), P('M6 7l1 13a1 1 0 0 0 1 1h8a1 1 0 0 0 1-1l1-13')],
  archive: [P('M3 5h18v4H3z'), P('M5 9v10a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1V9M10 13h4')],
  flag: [P('M5 21V4'), P('M5 5h13l-2.5 4L18 13H5z')],
  attach: [P('M20 11l-8.5 8.5a5 5 0 0 1-7-7L13 4a3.5 3.5 0 0 1 5 5l-8.5 8.5a2 2 0 0 1-3-3L15 6')],
  refresh: [P('M20 12a8 8 0 1 1-2.3-5.6'), P('M20 4v5h-5')],
  spam: [P('M12 3 3 20h18z'), P('M12 9v5M12 17h.01', { strokeWidth: 2.2 })],

  // ui
  close: [P('M6 6l12 12M18 6 6 18')],
  minimize: [P('M5 12h14')],
  maximize: [P('M5 5h14v14H5z')],
  restore: [P('M8 8h11v11H8z'), P('M5 16V5h11')],
  chevronDown: [P('m6 9 6 6 6-6')],
  chevronRight: [P('m9 6 6 6-6 6')],
  chevronLeft: [P('m15 6-6 6 6 6')],
  chevronUp: [P('m6 15 6-6 6 6')],
  more: [P('M6 12h.01M12 12h.01M18 12h.01', { strokeWidth: 2.4 })],
  plus: [P('M12 5v14M5 12h14')],
  minus: [P('M5 12h14')],
  check: [P('m5 13 4 4L19 7')],
  settings: [P('M12 15a3 3 0 1 0 0-6 3 3 0 0 0 0 6z'), P('M19.4 15a1.6 1.6 0 0 0 .3 1.8l.1.1a2 2 0 1 1-2.8 2.8l-.1-.1a1.6 1.6 0 0 0-2.7 1.1V21a2 2 0 1 1-4 0v-.1A1.6 1.6 0 0 0 7.5 19l-.1.1a2 2 0 1 1-2.8-2.8l.1-.1a1.6 1.6 0 0 0-1.1-2.7H3a2 2 0 1 1 0-4h.1A1.6 1.6 0 0 0 4.6 7.5l-.1-.1a2 2 0 1 1 2.8-2.8l.1.1a1.6 1.6 0 0 0 1.8.3H9a1.6 1.6 0 0 0 1-1.5V3a2 2 0 1 1 4 0v.1a1.6 1.6 0 0 0 2.7 1.1l.1-.1a2 2 0 1 1 2.8 2.8l-.1.1a1.6 1.6 0 0 0-.3 1.8V9a1.6 1.6 0 0 0 1.5 1H21a2 2 0 1 1 0 4h-.1a1.6 1.6 0 0 0-1.5 1z')],
  info: [P('M12 21a9 9 0 1 0 0-18 9 9 0 0 0 0 18z'), P('M12 11v5M12 8h.01', { strokeWidth: 2.2 })],
  sun: [P('M12 17a5 5 0 1 0 0-10 5 5 0 0 0 0 10z'), P('M12 2v2M12 20v2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M2 12h2M20 12h2M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4')],
  moon: [P('M20 14.5A8.5 8.5 0 0 1 9.5 4a8.5 8.5 0 1 0 10.5 10.5z')],
  grid: [P('M4 4h7v7H4zM13 4h7v7h-7zM4 13h7v7H4zM13 13h7v7h-7z')],
  list: [P('M8 6h12M8 12h12M8 18h12'), P('M4 6h.01M4 12h.01M4 18h.01', { strokeWidth: 2.4 })],
  clock: [P('M12 21a9 9 0 1 0 0-18 9 9 0 0 0 0 18z'), P('M12 7v5l3 2')],
  star: [P('m12 3 2.6 5.6 6 .8-4.4 4.2 1.1 6L12 16.8 6.7 19.6l1.1-6L3.4 9.4l6-.8z')],
  lock: [P('M6 11h12v9a1 1 0 0 1-1 1H7a1 1 0 0 1-1-1z'), P('M8.5 11V8a3.5 3.5 0 0 1 7 0v3')],
  eye: [P('M2 12s3.5-6 10-6 10 6 10 6-3.5 6-10 6-10-6-10-6z'), P('M12 15a3 3 0 1 0 0-6 3 3 0 0 0 0 6z')],
  download: [P('M12 3v12'), P('m8 11 4 4 4-4'), P('M5 19h14')],
  globe: [P('M12 21a9 9 0 1 0 0-18 9 9 0 0 0 0 18z'), P('M3 12h18'), P('M12 3a14 14 0 0 1 0 18 14 14 0 0 1 0-18z')],
  shield: [P('M12 3 5 6v6c0 4.5 3 7.5 7 9 4-1.5 7-4.5 7-9V6z'), P('m9 12 2 2 4-4')],
  heart: [P('M12 20s-7-4.4-7-9.5A4 4 0 0 1 12 8a4 4 0 0 1 7 2.5C19 15.6 12 20 12 20z')],
};

/**
 * @param {{ name: string, size?: number, className?: string, style?: object, title?: string }} props
 */
export function Icon({ name, size = 16, className, style, title, strokeWidth = 1.7 }) {
  const paths = PATHS[name];
  if (!paths) return null;
  return React.createElement(
    'svg',
    {
      viewBox: '0 0 24 24',
      width: size,
      height: size,
      className,
      style,
      fill: 'none',
      stroke: 'currentColor',
      strokeWidth,
      strokeLinecap: 'round',
      strokeLinejoin: 'round',
      'aria-hidden': title ? undefined : 'true',
      role: title ? 'img' : undefined,
      focusable: 'false',
    },
    title ? React.createElement('title', null, title) : null,
    paths.map((p, i) => React.createElement('path', { key: i, ...p }))
  );
}

export const iconNames = Object.keys(PATHS);
export default Icon;
