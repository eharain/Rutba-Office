// The installer's Choose apps page in the suite's languages.
//
//   node tools/make-installer-words.mjs     writes apps/desktop/resources/installer-words.nsh
//
// The page's words are the window's own (renderer/apps-choice.js marks them
// for the catalogue), taken from the Urdu, Arabic and Hindi catalogues, so
// the installer and Choose apps in Office say the same thing. The installer
// itself is in Arabic on an Arabic system; NSIS has no Urdu and its Hindi is
// broken (electron-builder.config.cjs), so on those systems the installer is
// in English and this page alone follows the system's own language.
// tests/installer-words.test.js holds the file to what this makes.

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

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

/** The file's text, from the catalogues as they are now. */
export async function makeInstallerWords() {
  const cats = {};
  for (const [code] of LANGUAGES) {
    cats[code] = (await import(pathToFileURL(path.join(ROOT, 'packages', 'office-ui', 'src', 'catalogues', `${code}.js`)).href)).default;
  }
  return [
    '; The Choose apps page in the suite\'s languages. Made by tools/make-installer-words.mjs',
    '; from the catalogues: change a word there and make this again.',
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
      const word = cats[code][en];
      if (word && word !== en) lines.push(`    StrCpy $rutbaW_${v} "${nsis(word)}"`);
    }
  });
  lines.push('  ${EndIf}', 'FunctionEnd', '!macroend');
  return lines;
}

if (import.meta.url === pathToFileURL(process.argv[1]).href) {
  fs.writeFileSync(OUT, await makeInstallerWords());
  console.log(path.relative(ROOT, OUT));
}
