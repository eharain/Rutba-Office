// The studio toolset's names as the Image window shows them. The engine keeps
// its vocabularies in English (packages/studio: geometries, word art presets,
// adjustment presets, icons, charts); each is written out here through t() so
// the catalogue carries it in every language, and a name the engine adds later
// shows in English until it is listed here.

import { t } from '@rutba/office-ui';

export const SHAPE_GROUP_NAMES = {
  Boxes: t('Boxes'), Rounds: t('Rounds'), Points: t('Points'), Pointers: t('Pointers'), Labels: t('Labels'),
};

export const SHAPE_NAMES = {
  rect: t('Rectangle'), pill: t('Pill'), parallelogram: t('Parallelogram'), trapezoid: t('Trapezoid'),
  diamond: t('Diamond'), frame: t('Frame'), ellipse: t('Ellipse'), semicircle: t('Semicircle'),
  ring: t('Ring'), pie: t('Pie slice'), teardrop: t('Teardrop'), blob: t('Blob'), cloud: t('Cloud'),
  triangle: t('Triangle'), polygon: t('Polygon'), star: t('Star'), burst: t('Burst'), gear: t('Gear'),
  plus: t('Cross'), bolt: t('Lightning'), arrow: t('Arrow'), doubleArrow: t('Double arrow'),
  line: t('Rule'), wave: t('Wave'), banner: t('Banner'), chevron: t('Chevron'), bubble: t('Speech bubble'),
  tag: t('Price tag'), bookmark: t('Bookmark'), shield: t('Shield'),
};

export const WORDART_NAMES = {
  plain: t('Plain'), outline: t('Outline'), sticker: t('Sticker'), gold: t('Gold'), neon: t('Neon'), extrude: t('3D'), arch: t('Arch'),
};

export const ADJUST_NAMES = {
  none: t('None'), punch: t('Punch'), matte: t('Matte'), warm: t('Warm'), cool: t('Cool'), mono: t('Mono'), fade: t('Fade'), noir: t('Noir'),
};

export const ICON_GROUP_NAMES = {
  Marks: t('Marks'), Arrows: t('Arrows'), Media: t('Media'), People: t('People'), 'Place & time': t('Place and time'),
  Commerce: t('Commerce'), Attention: t('Attention'), Things: t('Things'),
};

export const ICON_NAMES_SHOWN = {
  check: t('Check'), close: t('Close'), plus: t('Plus'), minus: t('Minus'), star: t('Star'),
  sparkle: t('Sparkle'), heart: t('Heart'), circle: t('Circle'), square: t('Square'), ring: t('Ring'),
  arrowRight: t('Arrow right'), arrowLeft: t('Arrow left'), arrowUp: t('Arrow up'),
  arrowDown: t('Arrow down'), chevronRight: t('Chevron'), chevronDown: t('Chevron down'),
  play: t('Play'), pause: t('Pause'), stop: t('Stop'), camera: t('Camera'), image: t('Picture'),
  music: t('Music'), mic: t('Microphone'), user: t('Person'), users: t('People'), chat: t('Chat'),
  bell: t('Bell'), share: t('Share'), thumbsUp: t('Thumbs up'), home: t('Home'), pin: t('Location'),
  clock: t('Clock'), calendar: t('Calendar'), cart: t('Cart'), tag: t('Tag'), gift: t('Gift'),
  percent: t('Percent'), card: t('Card'), truck: t('Delivery'), store: t('Store'), fire: t('Fire'),
  bolt: t('Lightning'), crown: t('Crown'), trophy: t('Trophy'), rocket: t('Rocket'), bulb: t('Idea'),
  leaf: t('Leaf'), search: t('Search'), lock: t('Lock'), key: t('Key'), eye: t('Eye'), shield: t('Shield'),
  download: t('Download'), upload: t('Upload'), flag: t('Flag'), sun: t('Sun'), moon: t('Moon'),
  droplet: t('Droplet'),
};

export const CHART_NAMES = {
  bar: t('Bars'), column: t('Columns'), line: t('Line'), area: t('Area'), pie: t('Pie'), donut: t('Donut'),
};

export const BLEND_NAMES = {
  normal: t('Normal'), multiply: t('Multiply'), screen: t('Screen'), overlay: t('Overlay'), darken: t('Darken'),
  lighten: t('Lighten'), 'color-dodge': t('Colour dodge'), 'color-burn': t('Colour burn'), 'hard-light': t('Hard light'),
  'soft-light': t('Soft light'), difference: t('Difference'), exclusion: t('Exclusion'), hue: t('Hue'),
  saturation: t('Saturation'), color: t('Colour'), luminosity: t('Luminosity'),
};

export const ANCHOR_NAMES = {
  nw: t('Top left'), n: t('Top'), ne: t('Top right'), w: t('Left'), c: t('Centre'), e: t('Right'),
  sw: t('Bottom left'), s: t('Bottom'), se: t('Bottom right'),
};

/** What a layer of each type is called before it has a name of its own. */
export const TYPE_NAMES = {
  image: t('Picture'), textblock: t('Text'), text: t('Text'), wordart: t('WordArt'), shape: t('Shape'),
  icon: t('Icon'), chart: t('Chart'), adjust: t('Adjustment'), group: t('Group'),
};

/** A layer's name as the panel shows it: its own, else its words, else its kind. */
export function nameOf(p) {
  if (!p) return '';
  if (p.label) return p.label;
  if (p.id === 'background') return t('Background');
  // The advanced editor's own pixel layers, named in the window's language.
  const n = String(p.id || '').match(/-(\d+)$/)?.[1] || '';
  if (p.kind === 'paint') return t('Paint {number}', { number: n }).trim();
  if (p.kind === 'adjusted') return t('Adjusted selection {number}', { number: n }).trim();
  if (p.kind === 'copied') return t('Copied selection {number}', { number: n }).trim();
  if (p.type === 'textblock' || p.type === 'text' || p.type === 'wordart') {
    const line = String(p.text || '').split('\n')[0].trim();
    if (line) return line.length > 32 ? `${line.slice(0, 31)}…` : line;
  }
  if (p.type === 'shape') return SHAPE_NAMES[p.geometry] || TYPE_NAMES.shape;
  if (p.type === 'icon') return ICON_NAMES_SHOWN[p.name] || TYPE_NAMES.icon;
  if (p.type === 'chart') return t('{kind} chart', { kind: CHART_NAMES[p.chart] || '' }).trim();
  if (p.type === 'adjust') return t('{look} adjustment', { look: ADJUST_NAMES[p.preset] || '' }).trim();
  return TYPE_NAMES[p.type] || p.id;
}
