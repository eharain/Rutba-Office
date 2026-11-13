// The installer's Languages page: what it matches the computer's languages
// with, made from the list Office uses (L5).
//
//   node tools/make-installer-languages.mjs   writes apps/desktop/resources/installer-languages.nsh
//
// The page reads the computer's languages (a key for each under
// HKCU\Control Panel\International\User Profile, and the language Windows
// shows itself in), matches each with the same keys language-choice.js uses,
// longest first, and shows a box for each language found, ticked. English is
// never among them, nor a language whose pack is not yet released (a row
// with no status: withPacks), nor one Office carries built in (notBuiltIn).
// A language whose pack is a draft (its row's `status`, from
// the pack's manifest, never from its tag) is named "{name} (draft)", in the
// page's words. tests/installer-languages.test.js holds the file to what this
// makes.
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { LANGUAGE_LIST } from '../packages/office-shell/src/electron/language-packs.js';
import { choiceKeys, regionalBases, withPacks, notBuiltIn } from '../packages/office-shell/src/electron/language-choice.js';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
export const OUT = path.join(ROOT, 'apps', 'desktop', 'resources', 'installer-languages.nsh');

/** The most boxes the page shows: the page's height at 12u each, under its lead. */
export const MOST = 6;

/** A string as NSIS reads it: $ doubled, quotes escaped. */
const nsis = (s) => String(s).replace(/\$/g, '$$$$').replace(/"/g, '$\\"');

/** How the page names a language, as Voices and languages does: its own name, and the English one where they differ. */
export function labelOf(row) {
  const own = row.nativeName || row.name;
  return row.name && row.name !== own ? `${own} · ${row.name}` : own;
}

export function makeInstallerLanguages(list = notBuiltIn(withPacks(LANGUAGE_LIST))) {
  const keys = choiceKeys(list);
  const rows = list.filter((r) => r.pack !== 'lang-en');
  const out = [
    '; The installer\'s Languages page: the keys it matches the computer\'s languages with, and',
    '; the names it shows. Made by tools/make-installer-languages.mjs from languages.json,',
    '; with the keys language-choice.js uses: change the list there and make this again.',
    '',
    '!macro rutbaLanguageFunctions',
    '; $R1 a language tag, as Windows writes it; $R2 the pack for it, or "".',
    'Function rutbaPackForTag',
    '  StrCpy $R2 ""',
  ];
  for (const [key, pack] of keys) {
    out.push(
      `  StrCpy $R4 $R1 ${key.length + 1}`,
      `  \${If} $R1 == "${nsis(key)}"`,
      `  \${OrIf} $R4 == "${nsis(key)}-"`,
      `    StrCpy $R2 "${nsis(pack)}"`,
      '    Return',
      '  ${EndIf}',
    );
  }
  out.push('FunctionEnd', '', '; $R2 a pack; $R5 its name on the page; $R8 "draft" when its pack is a draft.', 'Function rutbaLanguageLabel', '  StrCpy $R5 $R2', '  StrCpy $R8 ""');
  for (const r of rows) out.push(`  \${If} $R2 == "${nsis(r.pack)}"`, `    StrCpy $R5 "${nsis(labelOf(r))}"`, ...(r.status === 'draft' ? ['    StrCpy $R8 "draft"'] : []), '  ${EndIf}');
  out.push('FunctionEnd', '!macroend', '', '; A regional pack brings its base: the base is not offered beside it.', '!macro rutbaDropRegionalBases VAR');
  for (const [pack, base] of regionalBases(list)) {
    out.push(
      `  \${WordFind} "\${VAR}" ",${nsis(pack)}," "E*" $R6`,
      '  ${IfNot} ${Errors}',
      `    \${WordReplace} "\${VAR}" ",${nsis(base)}," "," "+" \${VAR}`,
      '  ${EndIf}',
      '  ClearErrors',
    );
  }
  out.push('!macroend', '');
  return out.join('\n');
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  fs.writeFileSync(OUT, makeInstallerLanguages());
  console.log(path.relative(process.cwd(), OUT));
}
