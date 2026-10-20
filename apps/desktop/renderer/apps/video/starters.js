// Video starters in the window: their names and titles in the window's
// language, and the stand-in pictures each one starts on, drawn here (a sky,
// a sun and two rows of hills in a palette of their own, no two alike) so a
// starter plays as soon as it opens and nothing is fetched for it.

import { t } from '@rutba/office-ui';

export const STARTER_NAMES = () => ({
  slideshow: t('Slideshow'),
  tutorial: t('Tutorial'),
  social: t('Vertical Social'),
  montage: t('Montage'),
});

export const STARTER_LINES = () => ({
  slideshow: t('pictures one after another, crossfading, with a title and an end card'),
  tutorial: t('a title, two steps named on the picture, a tip and an end card'),
  social: t('upright 9:16 for Reels, Shorts and TikTok, quick slides and a sticker'),
  montage: t('quick cuts of ten pictures between a title and an end card'),
});

/** The starters as the app menu offers them. */
export const STARTER_MENU = () => ({
  slideshow: t('New Slideshow'),
  tutorial: t('New Tutorial'),
  social: t('New Vertical Social Video'),
  montage: t('New Montage'),
});

export const STARTER_ICONS = { slideshow: 'picture', tutorial: 'textbox', social: 'phone', montage: 'layers' };

/** The titles a starter is laid out with. */
export const STARTER_WORDS = (key) => ({
  slideshow: { title: t('Our story'), end: t('Thank you for watching') },
  tutorial: { title: t('How to get started'), step1: t('Step 1: open the file'), step2: t('Step 2: make the change'), tip: t('Tip: press Ctrl+S to save'), end: t('Thank you for watching') },
  social: { title: t('Something new is here'), tag: t('NEW'), end: t('Follow for more') },
  montage: { title: t('The year in pictures'), end: t('Thank you for watching') },
}[key] || {});

const PALETTES = [
  ['#1b2a4e', '#f6a15a', '#ffe9b0', '#6a4c7a', '#2e2a44'],
  ['#5aa9e6', '#cdebfa', '#ffffff', '#6fae8c', '#2f6b4f'],
  ['#2b1e3f', '#e06c75', '#ffd38a', '#8e4b6e', '#3a2440'],
  ['#7fb7d9', '#f2e6c9', '#fff6d6', '#b79b6c', '#6b5434'],
  ['#0f2a3a', '#3e8a9a', '#d9f3f0', '#1e5560', '#0e2f36'],
  ['#f7c59f', '#fce8d5', '#ffffff', '#d08c60', '#7a4a2a'],
];

/** Stand-in picture `i` at `w` x `h`, as PNG bytes. */
export async function standInPicture(i, w, h) {
  const c = document.createElement('canvas');
  c.width = w;
  c.height = h;
  const g = c.getContext('2d');
  const [top, horizon, sun, far, near] = PALETTES[i % PALETTES.length];
  const sky = g.createLinearGradient(0, 0, 0, h * 0.7);
  sky.addColorStop(0, top);
  sky.addColorStop(1, horizon);
  g.fillStyle = sky;
  g.fillRect(0, 0, w, h);
  const sx = w * (0.22 + ((i * 37) % 60) / 100);
  const glow = g.createRadialGradient(sx, h * 0.36, 0, sx, h * 0.36, Math.min(w, h) * 0.22);
  glow.addColorStop(0, sun);
  glow.addColorStop(0.42, sun);
  glow.addColorStop(1, 'rgba(255,255,255,0)');
  g.fillStyle = glow;
  g.fillRect(0, 0, w, h);
  const ridge = (base, amp, colour, k) => {
    g.beginPath();
    g.moveTo(0, h);
    for (let x = 0; x <= w; x += 8) {
      const y = base + Math.sin(x / (w / (2.2 + k)) + i) * amp + Math.sin(x / (w / (5.1 + k)) + i * 2.3) * amp * 0.4;
      g.lineTo(x, y);
    }
    g.lineTo(w, h);
    g.closePath();
    g.fillStyle = colour;
    g.fill();
  };
  ridge(h * 0.62, h * 0.05, far, 0);
  ridge(h * 0.78, h * 0.07, near, 1.7);
  const blob = await new Promise((resolve) => c.toBlob(resolve, 'image/png'));
  return new Uint8Array(await blob.arrayBuffer());
}
