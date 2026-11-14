// The installer's Choose apps page in the suite's languages.
//
//   node tools/make-installer-words.mjs <rutba-office-packages>/languages/packs
//                                     writes apps/desktop/resources/installer-words.nsh
//
// The page's words are the window's own (renderer/apps-choice.js marks them
// for the catalogue), taken from the Urdu, Arabic and Hindi packs (lang-ur,
// lang-ar, lang-hi), each word found by its label's hash as a window finds it,
// so the installer and Choose apps in Office say the same thing. Office
// carries no words of its own for them since D-2026-11-14-16: the file is
// made at the release from the packs that release ships (RELEASING.md), as
// the English copy is taken. The installer
// itself is in Arabic on an Arabic system; NSIS has no Urdu and its Hindi is
// broken (electron-builder.config.cjs), so on those systems the installer is
// in English and this page alone follows the system's own language.
// tests/installer-words.test.js holds the file to what this makes from the
// packs when RUTBA_PACKS names them, and to its shape always.

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { keyOf, hashOf } from '../packages/office-shell/src/electron/language-packs.js';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
export const OUT = path.join(ROOT, 'apps', 'desktop', 'resources', 'installer-words.nsh');

/** Each word the page shows: its variable, and the English the catalogues are keyed by. */
export const WORDS = [
  ['title', 'Choose apps'],
  ['sub', 'The apps you want to see and to open your files.'],
  ['lead', 'Tick the apps and add-ons you want. Any of them can be added or taken away later, from Choose apps in Rutba Office.'],
  ['office', 'Office'],
  ['studio', 'Studio'],
  ['addons', 'Add-ons'],
  ['ai', 'AI Tools'],
  ['none', 'Choose one app at least.'],
  ['mail', 'Mail'],
  ['calendar', 'Calendar'],
  ['contacts', 'Contacts'],
  ['word', 'Documents'],
  ['sheets', 'Worksheets'],
  ['slides', 'Presentations'],
  ['pdf', 'PDF'],
  ['pictures', 'Pictures'],
  ['image', 'Image'],
  ['video', 'Video'],
  ['recorder', 'Recorder'],
  ['broadcast', 'Live Broadcast'],
  ['talk', 'Talk'],
  ['base', 'Data Files'],
  // The page that says which apps are put on the computer at all (docs/APP-PACKS.md 5.3).
  ['ocr', 'Text recognition'],
  ['ptitle', 'Apps to install'],
  ['psub', 'Only the apps you tick are put on this computer.'],
  ['plead', 'Tick the apps to install. An app not ticked is not on this computer at all; add it later from Add or remove apps in Rutba Office.'],
  // The Languages page (L5, renderer/voices-dialog.js INSTALLER_LANGUAGE_WORDS).
  ['ltitle', 'Languages'],
  ['lsub', 'Rutba Office in the languages this computer uses.'],
  ['llead', 'Tick the languages to add. They are added the first time Rutba Office starts while this computer is online. Any language can be added or removed later, in Voices and languages.'],
  ['ldraft', '{name} (draft)'],
];

/** What the uninstaller asks: whether the voices and languages added go too (renderer/voices-dialog.js). */
export const UNINSTALLER_WORDS = [
  ['udrop', 'Remove the voices and languages added to Rutba Office as well?'],
  ['ukeep', 'They take {size} on this computer. Keep them if you will install Rutba Office again.'],
];

/** The languages, by the primary language of a Windows language id. */
const LANGUAGES = [
  ['ar', 1, 'Arabic'],
  ['ur', 32, 'Urdu'],
  ['hi', 57, 'Hindi'],
];

/** A string as NSIS reads it: $ doubled, quotes and line ends escaped. */
const nsis = (s) => String(s).replace(/\$/g, '$$$$').replace(/"/g, '$\\"').replace(/\r?\n/g, '$\\r$\\n');

/**
 * Each language's words for the page, by English: from its pack's
 * messages.json in `packs` (lang-<tag>/), an entry taken by its `source`,
 * the hash of the label's English (language-packs.js), never by its id.
 * A word the pack lacks is left out, and the page shows it in English.
 */
export function packWords(packs) {
  const out = {};
  for (const [code] of LANGUAGES) {
    const file = path.join(packs, `lang-${code}`, 'messages.json');
    const bySource = new Map(Object.values(JSON.parse(fs.readFileSync(file, 'utf8'))).filter((e) => e && typeof e.text === 'string').map((e) => [e.source, e.text]));
    out[code] = {};
    for (const [, en] of [...WORDS, ...UNINSTALLER_WORDS]) {
      const word = bySource.get(hashOf(keyOf({ message: en })));
      if (word) out[code][en] = word;
    }
  }
  return out;
}

/** The file's text, from the packs in `packs` (or each language's words, by English, as packWords gives them). */
export async function makeInstallerWords(packs) {
  const cats = typeof packs === 'string' ? packWords(packs) : packs;
  return [
    '; The Choose apps page in the suite\'s languages. Made by tools/make-installer-words.mjs',
    '; from the language packs: make this again from the packs a release ships.',
    '',
    ...pickWords(cats, 'rutbaInstallerWords', 'rutbaPickWords', WORDS),
    '',
    '; What the uninstaller asks, in a function of its own, as NSIS names them un.',
    ...pickWords(cats, 'rutbaUninstallerWords', 'un.rutbaPickUninstallerWords', UNINSTALLER_WORDS),
    '',
  ].join('\n');
}

/** A macro that declares the words, with a function that picks them in the language. */
function pickWords(cats, macro, fn, words) {
  const lines = [
    `!macro ${macro}`,
    ...words.map(([v]) => `  Var rutbaW_${v}`),
    '',
    `Function ${fn}`,
    ...words.map(([v, en]) => `  StrCpy $rutbaW_${v} "${nsis(en)}"`),
    '  ; The installer\'s own language, or, where it fell back to English, the system\'s.',
    '  IntOp $0 $LANGUAGE & 0x3FF',
    '  ${If} $0 == 9',
    '    System::Call \'kernel32::GetUserDefaultUILanguage() i .r1\'',
    '    IntOp $0 $1 & 0x3FF',
    '  ${EndIf}',
  ];
  LANGUAGES.forEach(([code, id, name], i) => {
    lines.push(`  ${i === 0 ? '${If}' : '${ElseIf}'} $0 == ${id} ; ${name}`);
    for (const [v, en] of words) {
      const word = cats[code]?.[en];
      if (word && word !== en) lines.push(`    StrCpy $rutbaW_${v} "${nsis(word)}"`);
    }
  });
  lines.push('  ${EndIf}', 'FunctionEnd', '!macroend');
  return lines;
}

if (import.meta.url === pathToFileURL(process.argv[1]).href) {
  const packs = process.argv[2];
  if (!packs || !fs.existsSync(path.join(packs, 'lang-ur', 'messages.json'))) {
    console.error('node tools/make-installer-words.mjs <rutba-office-packages>/languages/packs');
    process.exit(1);
  }
  fs.writeFileSync(OUT, await makeInstallerWords(packs));
  console.log(path.relative(ROOT, OUT));
}
