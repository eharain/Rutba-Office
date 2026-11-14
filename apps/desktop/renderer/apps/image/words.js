// The studio toolset's names as the Image window shows them. The engine keeps
// its vocabularies in English (packages/studio: geometries, word art presets,
// adjustment presets, icons, charts); each is written out here through t() so
// the catalogue carries it in every language, and a name the engine adds later
// shows in English until it is listed here.

import { t } from '@rutba/office-ui';

export const SHAPE_GROUP_NAMES = {
  Boxes: t('common.boxes'), Rounds: t('common.rounds'), Points: t('common.points'), Pointers: t('common.pointers'), Labels: t('common.labels'),
};

export const SHAPE_NAMES = {
  rect: t('common.rectangle'), pill: t('common.pill'), parallelogram: t('common.parallelogram'), trapezoid: t('common.trapezoid'),
  diamond: t('common.diamond'), frame: t('common.frame'), ellipse: t('common.ellipse'), semicircle: t('common.semicircle'),
  ring: t('common.ring'), pie: t('common.pie-slice'), teardrop: t('common.teardrop'), blob: t('common.blob'), cloud: t('common.cloud'),
  triangle: t('common.triangle'), polygon: t('common.polygon'), star: t('common.star'), burst: t('common.burst'), gear: t('common.gear'),
  plus: t('common.cross'), bolt: t('common.lightning'), arrow: t('common.arrow'), doubleArrow: t('common.double-arrow'),
  line: t('common.rule'), wave: t('common.wave'), banner: t('common.banner'), chevron: t('common.chevron'), bubble: t('common.speech-bubble'),
  tag: t('common.price-tag'), bookmark: t('common.bookmark'), shield: t('common.shield'),
};

export const WORDART_NAMES = {
  plain: t('common.plain'), outline: t('common.outline'), sticker: t('common.sticker'), gold: t('common.gold'), neon: t('common.neon'), extrude: t('common.3d'), arch: t('common.arch'),
};

export const ADJUST_NAMES = {
  none: t('common.none'), punch: t('image.words.punch'), matte: t('image.words.matte'), warm: t('image.words.warm'), cool: t('image.words.cool'), mono: t('image.words.mono'), fade: t('common.fade'), noir: t('image.words.noir'),
};

export const ICON_GROUP_NAMES = {
  Marks: t('common.marks'), Arrows: t('common.arrows'), Media: t('common.media'), People: t('common.people'), 'Place & time': t('image.words.place-and-time'),
  Commerce: t('common.commerce'), Attention: t('common.attention'), Things: t('common.things'),
};

export const ICON_NAMES_SHOWN = {
  check: t('common.check'), close: t('common.close'), plus: t('common.plus'), minus: t('common.minus'), star: t('common.star'),
  sparkle: t('common.sparkle'), heart: t('common.heart'), circle: t('common.circle'), square: t('common.square'), ring: t('common.ring'),
  arrowRight: t('common.arrow-right'), arrowLeft: t('common.arrow-left'), arrowUp: t('common.arrow-up'),
  arrowDown: t('common.arrow-down'), chevronRight: t('common.chevron'), chevronDown: t('common.chevron-down'),
  play: t('common.play'), pause: t('common.pause'), stop: t('common.stop'), camera: t('common.camera'), image: t('common.picture'),
  music: t('common.music'), mic: t('common.microphone'), user: t('common.person'), users: t('common.people'), chat: t('common.chat'),
  bell: t('common.bell'), share: t('common.share'), thumbsUp: t('common.thumbs-up'), home: t('common.home'), pin: t('common.location'),
  clock: t('common.clock'), calendar: t('common.calendar'), cart: t('common.cart'), tag: t('common.tag'), gift: t('common.gift'),
  percent: t('common.percent'), card: t('common.card'), truck: t('common.delivery'), store: t('common.store'), fire: t('common.fire'),
  bolt: t('common.lightning'), crown: t('common.crown'), trophy: t('common.trophy'), rocket: t('common.rocket'), bulb: t('common.idea'),
  leaf: t('common.leaf'), search: t('common.search'), lock: t('common.lock'), key: t('common.key'), eye: t('common.eye'), shield: t('common.shield'),
  download: t('common.download'), upload: t('common.upload'), flag: t('common.flag'), sun: t('common.sun'), moon: t('common.moon'),
  droplet: t('common.droplet'),
};

export const CHART_NAMES = {
  bar: t('image.words.bars'), column: t('common.columns'), line: t('common.line'), area: t('common.area'), pie: t('common.pie'), donut: t('image.words.donut'),
};

export const BLEND_NAMES = {
  normal: t('common.normal'), multiply: t('common.multiply'), screen: t('common.screen'), overlay: t('common.overlay'), darken: t('common.darken'),
  lighten: t('common.lighten'), 'color-dodge': t('image.words.colour-dodge'), 'color-burn': t('image.words.colour-burn'), 'hard-light': t('image.words.hard-light'),
  'soft-light': t('image.words.soft-light'), difference: t('common.difference'), exclusion: t('image.words.exclusion'), hue: t('common.hue'),
  saturation: t('image.saturation'), color: t('common.colour'), luminosity: t('image.words.luminosity'),
};

export const ANCHOR_NAMES = {
  nw: t('common.top-left'), n: t('common.top'), ne: t('common.top-right'), w: t('common.left'), c: t('common.centre'), e: t('common.right'),
  sw: t('common.bottom-left'), s: t('common.bottom'), se: t('common.bottom-right'),
};

/** What a layer of each type is called before it has a name of its own. */
export const TYPE_NAMES = {
  image: t('common.picture'), textblock: t('common.text'), text: t('common.text'), wordart: t('common.word-art'), shape: t('common.shape'),
  icon: t('common.icon'), chart: t('common.chart'), adjust: t('image.adjustment'), group: t('common.group'),
};

/** A layer's name as the panel shows it: its own, else its words, else its kind. */
export function nameOf(p) {
  if (!p) return '';
  if (p.label) return p.label;
  if (p.id === 'background') return t('common.background');
  // The advanced editor's own pixel layers, named in the window's language.
  const n = String(p.id || '').match(/-(\d+)$/)?.[1] || '';
  if (p.kind === 'paint') return t('image.words.paint-number', { number: n }).trim();
  if (p.kind === 'adjusted') return t('image.words.adjusted-selection', { number: n }).trim();
  if (p.kind === 'copied') return t('image.words.copied-selection-number', { number: n }).trim();
  if (p.type === 'textblock' || p.type === 'text' || p.type === 'wordart') {
    const line = String(p.text || '').split('\n')[0].trim();
    if (line) return line.length > 32 ? `${line.slice(0, 31)}…` : line;
  }
  if (p.type === 'shape') return SHAPE_NAMES[p.geometry] || TYPE_NAMES.shape;
  if (p.type === 'icon') return ICON_NAMES_SHOWN[p.name] || TYPE_NAMES.icon;
  if (p.type === 'chart') return t('image.words.kind-chart', { kind: CHART_NAMES[p.chart] || '' }).trim();
  if (p.type === 'adjust') return t('image.words.look-adjustment', { look: ADJUST_NAMES[p.preset] || '' }).trim();
  return TYPE_NAMES[p.type] || p.id;
}
