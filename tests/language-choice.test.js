// Choosing languages at install and at a first start (L5): the rule that
// matches the computer's languages to packs, the installer's page made from
// the same rule, and what the first start adds or offers.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { choiceKeys, packForTag, matchLanguages, regionalBases, withPacks, notBuiltIn } from '../packages/office-shell/src/electron/language-choice.js';
import { LANGUAGE_LIST } from '../packages/office-shell/src/electron/language-packs.js';
import { makeInstallerLanguages, labelOf, OUT } from '../tools/make-installer-languages.mjs';
import { createFirstLanguages, FIRST_KEY, PENDING_KEY, OFFER_KEY } from '../apps/desktop/main/first-languages.js';

const INSTALLER = new URL('../apps/desktop/resources/installer.nsh', import.meta.url);

test('the exact tag, then a shorter form, then its base; English never', () => {
  const cases = {
    'pt-PT': 'lang-pt-pt', 'pt-BR': 'lang-pt-br', pt: 'lang-pt-br', 'pt-AO': 'lang-pt-pt',
    'es-MX': 'lang-es-419', 'es-419': 'lang-es-419', 'es-ES': 'lang-es', es: 'lang-es',
    'fr-CA': 'lang-fr-ca', 'fr-FR': 'lang-fr', 'fr-BE': 'lang-fr',
    'en-US': 'lang-en-us', 'en-GB': null, en: null, 'en-AU': null,
    'de-AT': 'lang-de', 'de-CH': 'lang-de', 'ur-PK': 'lang-ur', 'ur-IN': 'lang-ur', 'ar-SA': 'lang-ar', 'hi-IN': 'lang-hi',
    'zh-CN': 'lang-zh-hans', 'zh-SG': 'lang-zh-hans', 'zh-TW': 'lang-zh-hant', 'zh-HK': 'lang-zh-hant', 'zh-Hans-CN': 'lang-zh-hans', 'zh-Hant-TW': 'lang-zh-hant', zh: 'lang-zh-hans',
    'ga-IE': 'lang-ga', 'ja-JP': null, 'sr-Latn-RS': null, '': null,
    // As Linux writes them.
    'de_DE.UTF-8': 'lang-de', 'pt_BR': 'lang-pt-br', 'fr_CA@euro': 'lang-fr-ca',
  };
  for (const [tag, pack] of Object.entries(cases)) assert.equal(packForTag(tag), pack, tag);
  // A key never matches a longer word it begins: "ga" is not "gal".
  assert.equal(packForTag('gal-ES'), null);
  for (const [key, pack] of choiceKeys()) {
    assert.notEqual(pack, 'lang-en');
    assert.ok(LANGUAGE_LIST.some((r) => r.pack === pack), `${key} -> ${pack} is in the list`);
  }
});

test('each language once, in the computer\'s order, a regional pack\'s base left out beside it', () => {
  assert.deepEqual(matchLanguages(['ur-PK', 'en-GB', 'ur-IN']), ['lang-ur']);
  assert.deepEqual(matchLanguages(['fr-CA', 'fr-FR']), ['lang-fr-ca']);
  assert.deepEqual(matchLanguages(['fr-FR', 'fr-CA']), ['lang-fr-ca']);
  assert.deepEqual(matchLanguages(['pt-BR', 'pt-PT']), ['lang-pt-pt']);
  assert.deepEqual(matchLanguages(['de-DE', 'ar-EG', 'es-MX', 'es-ES']), ['lang-de', 'lang-ar', 'lang-es-419']);
  assert.deepEqual(matchLanguages(['en-GB', 'ja-JP']), []);
  assert.deepEqual(matchLanguages(null), []);
  assert.deepEqual(Object.fromEntries(regionalBases()), { 'lang-pt-pt': 'lang-pt-br', 'lang-es-419': 'lang-es', 'lang-fr-ca': 'lang-fr' });
});

test('the installer\'s keys and names are made from the same list, and are current', () => {
  const made = makeInstallerLanguages();
  assert.equal(fs.readFileSync(OUT, 'utf8').replace(/\r\n/g, '\n'), made, 'run node tools/make-installer-languages.mjs');
  // Every key, longest first, as the rule above has them.
  const keys = [...made.matchAll(/\$\{If\} \$R1 == "([^"]+)"\n {2}\$\{OrIf\} \$R4 == "\1-"\n {4}StrCpy \$R2 "([^"]+)"/g)].map((m) => [m[1], m[2]]);
  assert.deepEqual(keys, choiceKeys(notBuiltIn(withPacks())));
  for (const [key] of keys) assert.match(made, new RegExp(`StrCpy \\$R4 \\$R1 ${key.length + 1}\\n  \\$\\{If\\} \\$R1 == "${key}"`));
  // A name for every language but English, in its own words first.
  for (const r of notBuiltIn(withPacks()).filter((x) => x.pack !== 'lang-en')) assert.ok(made.includes(`StrCpy $R5 "${labelOf(r)}"`), r.pack);
  // Urdu, Arabic and Hindi are offered as packs like the rest: Office carries none built in since D-2026-11-14-16.
  for (const pack of ['lang-ur', 'lang-ar', 'lang-hi']) assert.ok(made.includes(`StrCpy $R2 "${pack}"`), pack);
  // A language whose pack is not yet released is not offered.
  for (const r of LANGUAGE_LIST.filter((x) => !x.status)) assert.ok(!made.includes(`"${r.pack}"`), r.pack);
  assert.equal(labelOf({ name: 'Urdu', nativeName: 'اردو' }), 'اردو · Urdu');
  assert.equal(labelOf({ name: 'Deutsch', nativeName: 'Deutsch' }), 'Deutsch');
  assert.doesNotMatch(made, /lang-en"/);
  for (const [pack, base] of regionalBases(notBuiltIn(withPacks()))) assert.ok(made.includes(`",${pack},"`) && made.includes(`",${base},"`), pack);
  // A draft is marked by its row's status alone, never its tag.
  const drafted = makeInstallerLanguages([{ tag: 'de', pack: 'lang-de', name: 'German', nativeName: 'Deutsch', direction: 'ltr', status: 'draft' }, { tag: 'fr', pack: 'lang-fr', name: 'French', nativeName: 'Français', direction: 'ltr', status: 'reviewed' }]);
  assert.ok(drafted.includes('$R2 == "lang-de"\n    StrCpy $R5 "Deutsch · German"\n    StrCpy $R8 "draft"\n'), 'German marked a draft');
  assert.ok(drafted.includes('$R2 == "lang-fr"\n    StrCpy $R5 "Français · French"\n  ${EndIf}'), 'French, reviewed, not marked');
});

test('the page is shown on a first install only, with the computer\'s languages ticked, and its choice written for Office', () => {
  const script = fs.readFileSync(INSTALLER, 'utf8').replace(/\r\n/g, '\n');
  assert.match(script, /Page custom rutbaAppsPageShow rutbaAppsPageLeave\n {2}Page custom rutbaLanguagesPageShow rutbaLanguagesPageLeave\n/);
  const show = script.slice(script.indexOf('Function rutbaLanguagesPageShow'), script.indexOf('FunctionEnd', script.indexOf('Function rutbaLanguagesPageShow')));
  assert.ok(show.indexOf('${isUpdated}') < show.indexOf('nsDialogs::Create'), 'not on an update');
  assert.match(show, /\$\{If\} \$rutbaLang1 == ""\n {4}Abort/, 'no page when no language matches');
  const find = script.slice(script.indexOf('Function rutbaFindLanguages'), script.indexOf('FunctionEnd', script.indexOf('Function rutbaFindLanguages')));
  assert.match(find, /EnumRegKey \$R1 HKCU "Control Panel\\International\\User Profile" \$R0/);
  assert.match(find, /GetUserDefaultUILanguage/);
  assert.ok(find.indexOf('rutbaDropRegionalBases') < find.indexOf('rutbaLangSlot 1'), 'bases left out before the boxes');
  assert.match(script, /\$\{NSD_Check\} \$rutbaLangBox\$\{N\}/, 'ticked');
  assert.match(script, /\$\{If\} \$rutbaLangShown == "1"\n {4}WriteRegStr HKCU "Software\\Rutba Office" "LanguagesChosen" "\$rutbaLangChosen"/);
  assert.match(script, /!include "\$\{__FILEDIR__\}\\installer-languages.nsh"/);
});

/* ── the first start ───────────────────────────────────────────────── */

const settings = (init = {}) => {
  const m = new Map(Object.entries(init));
  return { get: (k, d) => (m.has(k) ? m.get(k) : d), set: (k, v) => m.set(k, v), all: () => Object.fromEntries(m) };
};

/** An add-ons service as first-languages.js uses it: a feed that offers some packs, online or not. */
function fakeAddons({ offered = [], online = true, installed = [], drafts = [] } = {}) {
  const added = [];
  const here = [...installed];
  return {
    added,
    online,
    installed: () => here.map((id) => ({ id })),
    list: () => ({ staged: [], jobs: [] }),
    async feed() {
      // As the service lists them: `status` is whether it is added, `packStatus` the pack's own; offline, no feed and the languages known.
      const packStatus = (id) => (drafts.includes(id) ? 'draft' : 'reviewed');
      if (!this.online) return { feed: null, offered: [], languages: offered.map((id) => ({ tag: id.replace(/^lang-/, ''), status: packStatus(id) })) };
      return { feed: { from: 'network' }, offered: offered.map((id) => ({ id, kind: 'language', tag: id.replace(/^lang-/, ''), status: 'add', packStatus: packStatus(id) })), languages: [] };
    },
    async add({ id }) {
      if (!this.online) throw new Error('offline');
      added.push(id);
      here.push(id);
      return {};
    },
  };
}

test('Windows: the installer\'s choice is added at the first start, and Home offers nothing', async () => {
  const stores = { settings: settings() };
  const addons = fakeAddons({ offered: ['lang-ur', 'lang-de'] });
  const first = createFirstLanguages({ stores, addons, office: '2.26.0', chosen: async () => ['lang-ur', 'lang-en', 'not a pack'], preferred: () => ['de-DE'] });
  assert.deepEqual(await first.atStart(), ['lang-ur']);
  assert.deepEqual(addons.added, ['lang-ur']);
  assert.deepEqual(stores.settings.get(FIRST_KEY), { office: '2.26.0', from: 'installer' });
  assert.deepEqual((await first.offer()).offer, []);
  // None ticked is a choice too: nothing added, nothing offered.
  const none = createFirstLanguages({ stores: { settings: settings() }, addons: fakeAddons({ offered: ['lang-de'] }), office: '2.26.0', chosen: async () => [], preferred: () => ['de-DE'] });
  assert.deepEqual(await none.atStart(), []);
  assert.deepEqual((await none.offer()).offer, []);
});

test('with no installer\'s choice, the computer\'s languages are offered once; Add or Not now', async () => {
  const stores = { settings: settings() };
  const addons = fakeAddons({ offered: ['lang-de', 'lang-fr-ca'], installed: ['lang-ar'], drafts: ['lang-de'] });
  const first = createFirstLanguages({ stores, addons, office: '2.26.0', preferred: () => ['de-AT', 'ar-SA', 'fr-CA', 'en-GB'] });
  assert.deepEqual(await first.atStart(), [], 'nothing added unasked');
  assert.deepEqual(await first.offer(), { offer: ['lang-de', 'lang-fr-ca'], waiting: [], online: true, moved: null, drafts: ['lang-de'] }, 'Arabic is here already; German offered as a draft, by its status');
  // Offline, the draft mark stays: from the languages the list knows.
  addons.online = false;
  assert.deepEqual((await first.offer()).drafts, ['lang-de']);
  addons.online = true;
  const took = await first.takeOffer({ ids: ['lang-fr-ca'] });
  assert.deepEqual(took.added, ['lang-fr-ca']);
  assert.deepEqual(took.offer, ['lang-de'], 'the one not named stays offered');
  // Once: a second start offers nothing new; Not now puts the rest away.
  await first.atStart();
  assert.deepEqual((await first.offer()).offer, ['lang-de']);
  first.dismissOffer();
  await first.atStart();
  assert.deepEqual((await first.offer()).offer, []);

  // A computer in Urdu, Arabic or Hindi is offered their packs like any other language's:
  // Office carries none of them built in since D-2026-11-14-16 (the contract's builtIn is empty).
  const builtIn = createFirstLanguages({ stores: { settings: settings() }, addons: fakeAddons({ offered: ['lang-ur', 'lang-ar', 'lang-hi', 'lang-de'] }), office: '2.26.0', preferred: () => ['ur-PK', 'ar-SA', 'hi-IN', 'de-DE'] });
  await builtIn.atStart();
  assert.deepEqual((await builtIn.offer()).offer, ['lang-ur', 'lang-ar', 'lang-hi', 'lang-de']);
  // A language a release still carried built in would be offered nothing.
  assert.deepEqual(notBuiltIn([{ tag: 'ur-IN' }, { tag: 'de' }], ['ur']), [{ tag: 'de' }]);

  const other = createFirstLanguages({ stores: { settings: settings() }, addons: fakeAddons({ offered: ['lang-de'] }), office: '2.26.0', preferred: () => ['de-DE'] });
  await other.atStart();
  assert.deepEqual(other.dismissOffer().offer, []);
  assert.deepEqual((await other.offer()).offer, []);
});

test('offline, what was chosen waits and is added at a later start; a pack not offered yet stays waiting', async () => {
  const stores = { settings: settings() };
  const addons = fakeAddons({ offered: ['lang-ur'], online: false });
  const first = createFirstLanguages({ stores, addons, office: '2.26.0', chosen: async () => ['lang-ur', 'lang-mt'] });
  assert.deepEqual(await first.atStart(), []);
  assert.deepEqual(await first.offer(), { offer: [], waiting: ['lang-ur', 'lang-mt'], online: false, moved: null, drafts: [] });
  addons.online = true;
  const again = createFirstLanguages({ stores, addons, office: '2.26.0', chosen: async () => null });
  assert.deepEqual(await again.atStart(), ['lang-ur']);
  assert.deepEqual(stores.settings.get(PENDING_KEY), ['lang-mt'], 'kept for a list that offers it');
  // A later installer's choice is taken too, beside what waits.
  const later = createFirstLanguages({ stores, addons, office: '2.26.0', chosen: async () => ['lang-de'] });
  await later.atStart();
  assert.deepEqual(stores.settings.get(PENDING_KEY), ['lang-mt', 'lang-de']);
  assert.deepEqual(stores.settings.get(OFFER_KEY, []), []);
});

test('after a release that leaves a language out of the core, Home says so once, and where to add it when it was not brought', async () => {
  const stores = { settings: settings() };
  const addons = fakeAddons({ installed: ['lang-ur'] });
  const first = createFirstLanguages({ stores, addons, office: '2.27.0', windowLanguage: () => 'ur-PK', leftCore: ['ur'] });
  assert.deepEqual((await first.offer()).moved, { tag: 'ur', added: true, waiting: false });
  await first.dismissOffer({ moved: true });
  assert.equal((await first.offer()).moved, null, 'once');
  const missing = createFirstLanguages({ stores: { settings: settings() }, addons: fakeAddons(), office: '2.27.0', windowLanguage: () => 'ur', leftCore: ['ur'] });
  assert.deepEqual((await missing.offer()).moved, { tag: 'ur', added: false, waiting: false });
  // A window in another language, or nothing left the core: nothing said.
  assert.equal((await createFirstLanguages({ stores: { settings: settings() }, addons: fakeAddons(), office: '2.27.0', windowLanguage: () => 'de', leftCore: ['ur'] }).offer()).moved, null);
  assert.equal((await createFirstLanguages({ stores: { settings: settings() }, addons: fakeAddons(), office: '2.26.0', windowLanguage: () => 'ur', leftCore: [] }).offer()).moved, null);
  // This release's own list, from the contract: Urdu, Arabic and Hindi left the core (D-2026-11-14-16).
  assert.deepEqual((await createFirstLanguages({ stores: { settings: settings() }, addons: fakeAddons(), office: '2.26.0', windowLanguage: () => 'ur' }).offer()).moved, { tag: 'ur', added: false, waiting: false });
});

test('the first start after an update that left the window\'s language out of the core brings its pack: added for a person who chose it, offered in one click to anyone else', async () => {
  // From an Office whose updater did not wait for packs (2.25.0): Urdu left the core, and lang-ur is not here.
  const update = (o) => createFirstLanguages({ office: '2.27.0', leftCore: ['ur'], windowLanguage: () => 'ur-PK', ...o });
  const started = { languages: { firstStart: { office: '2.25.0', from: 'system' } } };
  const settingsFrom = () => {
    const s = settings();
    s.set('languages.firstStart', started.languages.firstStart);
    return s;
  };

  // Chosen in Office, the list to be had: added at the start, and Home says it was added.
  const chose = { settings: settingsFrom() };
  const online = fakeAddons({ offered: ['lang-ur'] });
  const a = update({ stores: chose, addons: online, chosenLanguage: () => 'ur' });
  await a.atStart();
  assert.deepEqual(online.added, ['lang-ur']);
  assert.deepEqual((await a.offer()).moved, { tag: 'ur', added: true, waiting: false });

  // Chosen, offline: asked for, waiting, said on Home; added at a later start once online, and asked for once only.
  const waits = { settings: settingsFrom() };
  const away = fakeAddons({ offered: ['lang-ur'], online: false });
  const b = update({ stores: waits, addons: away, chosenLanguage: () => 'ur-PK' });
  await b.atStart();
  assert.deepEqual(waits.settings.get(PENDING_KEY), ['lang-ur']);
  const shown = await b.offer();
  assert.deepEqual(shown.moved, { tag: 'ur', added: false, waiting: true });
  assert.equal(shown.online, false);
  await b.atStart();
  assert.deepEqual(waits.settings.get(PENDING_KEY), ['lang-ur'], 'not asked for twice');
  away.online = true;
  await b.atStart();
  assert.deepEqual(away.added, ['lang-ur']);
  assert.deepEqual(waits.settings.get(PENDING_KEY), []);

  // In Urdu by the computer's language, not chosen in Office: offered on Home, and one Add brings it.
  const system = { settings: settingsFrom() };
  const offerOnline = fakeAddons({ offered: ['lang-ur'] });
  const c = update({ stores: system, addons: offerOnline });
  await c.atStart();
  assert.deepEqual(offerOnline.added, [], 'nothing added unasked');
  const home = await c.offer();
  assert.deepEqual(home.moved, { tag: 'ur', added: false, waiting: false });
  assert.deepEqual(home.offer, ['lang-ur']);
  assert.equal(home.online, true);
  const took = await c.takeOffer({ ids: ['lang-ur'] });
  assert.deepEqual(offerOnline.added, ['lang-ur']);
  assert.deepEqual(took.moved, { tag: 'ur', added: true, waiting: false });

  // The one Add takes only its pack: another first-start offer stays.
  const both = { settings: settingsFrom() };
  both.settings.set(OFFER_KEY, ['lang-de']);
  const two = fakeAddons({ offered: ['lang-ur', 'lang-de'] });
  const e = update({ stores: both, addons: two });
  await e.atStart();
  assert.deepEqual(both.settings.get(OFFER_KEY), ['lang-de', 'lang-ur']);
  await e.takeOffer({ ids: ['lang-ur'] });
  assert.deepEqual(two.added, ['lang-ur']);
  assert.deepEqual(both.settings.get(OFFER_KEY), ['lang-de']);
  assert.deepEqual((await e.offer()).offer, ['lang-de']);

  // Closed instead: the line and its offer go together, and a later start offers nothing again.
  const declined = { settings: settingsFrom() };
  const d = update({ stores: declined, addons: fakeAddons({ offered: ['lang-ur'] }) });
  await d.atStart();
  const after = await d.dismissOffer({ moved: true });
  assert.equal(after.moved, null);
  assert.deepEqual(declined.settings.get(OFFER_KEY), []);
  await d.atStart();
  assert.deepEqual(declined.settings.get(OFFER_KEY), [], 'once');

  // Still built in (no leftCore), or the pack already here: nothing brought.
  const kept = { settings: settingsFrom() };
  await createFirstLanguages({ stores: kept, addons: fakeAddons({ offered: ['lang-ur'] }), office: '2.26.0', leftCore: [], windowLanguage: () => 'ur', chosenLanguage: () => 'ur' }).atStart();
  assert.deepEqual([kept.settings.get(PENDING_KEY, []), kept.settings.get(OFFER_KEY, [])], [[], []]);
  const had = { settings: settingsFrom() };
  await update({ stores: had, addons: fakeAddons({ installed: ['lang-ur'] }), chosenLanguage: () => 'ur' }).atStart();
  assert.deepEqual([had.settings.get(PENDING_KEY, []), had.settings.get(OFFER_KEY, [])], [[], []]);
});

test('a first start offers only languages with a pack: the list, and those the feed already offers', async () => {
  const fresh = (offered, online = true) => createFirstLanguages({ stores: { settings: settings() }, addons: fakeAddons({ offered, online }), office: '2.26.0', preferred: () => ['fr-CA', 'sv-SE'] });
  // No lang-fr-ca released and Swedish not yet a pack: French is offered for fr-CA, Swedish not at all.
  const noRegion = fresh(['lang-fr']);
  await noRegion.atStart();
  assert.deepEqual((await noRegion.offer()).offer, ['lang-fr']);
  // The feed offers lang-fr-ca: it is offered instead of its base.
  const region = fresh(['lang-fr', 'lang-fr-ca']);
  await region.atStart();
  assert.deepEqual((await region.offer()).offer, ['lang-fr-ca']);
});
