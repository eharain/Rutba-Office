// Review → Check Accessibility and Review → Spelling, on the main process.
//
// The engines live here, so this is where a document is read for the
// accessibility rules and walked for the spelling pass. It hands the
// document service two things: operations for its DOC_OPS / SHEET_OPS /
// DECK_OPS tables (every fix and every change is an op, so it is one undo
// step and the window's model follows as it does after any edit), and
// `act`, behind `doc.proof`, for what reads rather than writes — the
// findings, the next misspelt word, a word's suggestions, the dictionary.
//
// The person's own dictionary is kept in the suite's settings store under
// `proofing.dictionary`, shared by the three apps; Ignore All lasts as long
// as the document is open. The dictionaries themselves are read on a worker
// thread the first time a pass or a right-click asks (speller-host.js).

import fs from 'node:fs';
import {
  RULES, wordDrawings, sheetDrawings, readAltProps, findElements, toDic, fromDic, checkAccessibility,
  describeWord, describeSheet, describeDeck,
  setWordAltText, setWordTableHeader, removeWordParagraphs, setWordTitle,
  setSheetAltText, setSheetTableHeader,
  setDeckAltText, setDeckTableHeader, setDeckTextColour, setDeckSlideTitle,
  wordSegments, wordStart, wordLanguage, replaceWordText,
  sheetSegments, sheetStart, replaceSheetText,
  deckSegments, deckLanguage, replaceDeckText,
  createSpellerHost, chooseLanguage, LANGUAGE_NAMES, nextMisspelling, changeAllEdits, acceptedBy, DEFAULT_OPTIONS,
} from '@rutba/proofing';

const DICTIONARY_KEY = 'proofing.dictionary';
const OPTIONS_KEY = 'proofing.options';

function systemLocale() {
  try {
    return Intl.DateTimeFormat().resolvedOptions().locale || 'en-GB';
  } catch {
    return 'en-GB';
  }
}

export function createProofing({ stores = null, locale = systemLocale, worker = true } = {}) {
  const host = createSpellerHost({ worker });
  // Without a settings store (a test) the dictionary lives here.
  let memoryWords = [];
  const readWords = () => {
    const list = stores?.settings ? stores.settings.get(DICTIONARY_KEY, []) : memoryWords;
    return Array.isArray(list) ? list.filter((w) => typeof w === 'string' && w.trim()) : [];
  };
  const writeWords = (list) => {
    const clean = [...new Set(list.map((w) => String(w).trim()).filter(Boolean))].sort((a, b) => a.localeCompare(b));
    if (stores?.settings) stores.settings.set(DICTIONARY_KEY, clean);
    else memoryWords = clean;
    return clean;
  };
  const readOptions = () => ({ ...DEFAULT_OPTIONS, ...(stores?.settings?.get(OPTIONS_KEY, null) || {}) });

  /** Chromium's own underline learns what the person adds, and forgets what they take out. */
  const teachChromium = (win, word, add) => {
    try {
      const ses = win?.webContents?.session;
      if (!ses) return;
      if (add) ses.addWordToSpellCheckerDictionary?.(word);
      else ses.removeWordFromSpellCheckerDictionary?.(word);
    } catch {
      /* a window on its way out */
    }
  };

  const stateOf = (session) => session.proofing || (session.proofing = { ignored: new Set() });

  const languageOf = (session) => {
    const own = session.kind === 'doc' ? wordLanguage(session.engine) : session.kind === 'deck' ? deckLanguage(session.engine) : null;
    return chooseLanguage(own, locale());
  };

  const segmentsOf = (session, { sheets } = {}) => {
    if (session.kind === 'doc') return wordSegments(session.engine);
    if (session.kind === 'sheet') return sheetSegments(session.engine, { sheets: Array.isArray(sheets) && sheets.length ? sheets : [session.engine.activeSheet] });
    return deckSegments(session.engine);
  };

  /** Where a pass starts: the caret, the active cell, the slide on screen. */
  const startOf = (session, a, segments) => {
    if (session.kind === 'doc') return wordStart(session.engine);
    if (session.kind === 'sheet') {
      const s = sheetStart(session.engine);
      const after = segments.find((g) => g.where.sheet === s.sheet && (g.where.row > s.row || (g.where.row === s.row && g.where.col >= s.col)));
      return after ? { key: after.key, offset: 0 } : { key: segments[0]?.key ?? null, offset: 0 };
    }
    const slide = Number(a.slide) || 0;
    const first = segments.find((g) => g.where.slide >= slide);
    return first ? { key: first.key, offset: 0 } : { key: segments[0]?.key ?? null, offset: 0 };
  };

  const accepted = (session) => {
    const personal = new Set(readWords());
    const ignored = stateOf(session).ignored;
    return (word) => acceptedBy(personal, word) || acceptedBy(ignored, word);
  };

  /* ── operations, for the document service's tables ─────────────────── */

  const ops = {
    doc: {
      // By its place among the drawings, or — from a picture picked on the
      // page — by its paragraph and its place among that paragraph's pictures.
      setAltText: (v, a) => setWordAltText(v, { drawing: a.drawing ?? wordDrawingAt(v, a)?.ordinal, descr: a.descr, decorative: Boolean(a.decorative) }),
      setHeaderRow: (v, a) => setWordTableHeader(v, { table: a.table }),
      removeEmptyParagraphs: (v, a) => removeWordParagraphs(v, { blocks: a.blocks }),
      setDocTitle: (v, a) => setWordTitle(v, { title: a.title }),
      spellReplace: (v, a) => { replaceWordText(v, a.edits || []); },
      spellChangeAll: (v, a) => {
        const edits = changeAllEdits(wordSegments(v), a.word, a.replacement, readOptions());
        if (edits.length) replaceWordText(v, edits);
        return edits.length;
      },
    },
    sheet: {
      setAltText: (v, a) => { setSheetAltText(v, { sheet: a.sheet, anchor: a.anchor, descr: a.descr, decorative: Boolean(a.decorative) }); },
      setHeaderRow: (v, a) => { setSheetTableHeader(v, { part: a.part }); },
      spellReplace: (v, a) => replaceSheetText(v, a.edits || []),
      spellChangeAll: (v, a) => {
        const edits = changeAllEdits(sheetSegments(v, { sheets: Array.isArray(a.sheets) && a.sheets.length ? a.sheets : [v.activeSheet] }), a.word, a.replacement, readOptions());
        return edits.length ? replaceSheetText(v, edits) : 0;
      },
    },
    deck: {
      setAltText: (d, a) => { setDeckAltText(d, { slide: a.slide, shape: a.shape, descr: a.descr, decorative: Boolean(a.decorative) }); },
      setHeaderRow: (d, a) => { setDeckTableHeader(d, { slide: a.slide, shape: a.shape }); },
      setShapeTextColour: (d, a) => { setDeckTextColour(d, { slide: a.slide, shape: a.shape, colour: a.colour }); },
      setSlideTitle: (d, a) => { setDeckSlideTitle(d, { slide: a.slide, shape: a.shape, title: a.title }); },
      spellReplace: (d, a) => replaceDeckText(d, a.edits || []),
      spellChangeAll: (d, a) => {
        const edits = changeAllEdits(deckSegments(d), a.word, a.replacement, readOptions());
        return edits.length ? replaceDeckText(d, edits) : 0;
      },
    },
  };

  /** The Word drawing a picked picture is: its paragraph, and its place among that paragraph's pictures. */
  function wordDrawingAt(view, { block, image = 0 }) {
    return wordDrawings(view).find((d) => d.block === Number(block) && d.kind === 'picture' && d.image === Number(image)) || null;
  }

  /** What an object's Alt Text dialog opens with: its name, its description, whether it is decorative. */
  function altTextOf(session, a) {
    if (session.kind === 'doc') {
      const d = a.drawing != null ? wordDrawings(session.engine).find((x) => x.ordinal === Number(a.drawing)) : wordDrawingAt(session.engine, a);
      return d ? { name: d.name, descr: d.descr, decorative: d.decorative, target: { drawing: d.ordinal } } : null;
    }
    if (session.kind === 'sheet') {
      const d = sheetDrawings(session.engine, a.sheet || session.engine.activeSheet)[Number(a.anchor) || 0];
      return d ? { name: d.name, descr: d.descr, decorative: d.decorative, target: { sheet: a.sheet || session.engine.activeSheet, anchor: d.anchor } } : null;
    }
    const part = session.engine.slideParts[Number(a.slide)]?.part;
    if (!part) return null;
    const hit = findElements(session.engine.pkg.text(part), 'p:cNvPr').find((e) => / id="([^"]*)"/.exec(e.open)?.[1] === String(a.shape));
    if (!hit) return null;
    const p = readAltProps(hit.xml);
    return { name: p.name, descr: p.descr, decorative: p.decorative, target: { slide: Number(a.slide), shape: a.shape } };
  }

  /* ── doc.proof: what reads ─────────────────────────────────────────── */

  async function act(session, action, a = {}, win = null) {
    switch (action) {
      case 'accessibility': {
        const engine = session.engine;
        const model = session.kind === 'doc' ? describeWord(engine) : session.kind === 'sheet' ? describeSheet(engine) : describeDeck(engine);
        const result = checkAccessibility(model);
        return { ...result, rules: RULES };
      }

      case 'spellStart': {
        const { lang, from, documentLang } = languageOf(session);
        const began = Date.now();
        const loaded = await host.load(lang);
        const segments = segmentsOf(session, a);
        return { lang, languageName: LANGUAGE_NAMES[lang], from, documentLang: documentLang || null, start: startOf(session, a, segments), segments: segments.length, loadMs: Date.now() - began, builtMs: loaded.builtMs, options: readOptions() };
      }

      case 'spellNext': {
        const { lang } = languageOf(session);
        const options = { ...readOptions(), ...(a.options || {}) };
        const segments = segmentsOf(session, a);
        const step = await nextMisspelling({
          segments, from: a.from, stop: a.stop, wrapped: Boolean(a.wrapped), options,
          misspelt: (words) => host.misspelt(lang, words),
          accepted: accepted(session),
        });
        // The word is answered at once; its suggestions are asked for next
        // (spellSuggest), on a thread of their own — an unusual name can
        // take two seconds, and the pane shows the word meanwhile.
        return step;
      }

      case 'spellSuggest': {
        const { lang } = languageOf(session);
        const word = String(a.word || '').trim();
        return { word, suggestions: word ? await host.suggest(lang, word, 6) : [] };
      }

      case 'spellCheckWord': {
        // A right-click: is the word under the pointer spelt right, and if
        // not, what might it be?
        const word = String(a.word || '').replace(/’/g, "'").trim();
        if (!word) return { misspelt: false, suggestions: [] };
        const { lang } = languageOf(session);
        if (accepted(session)(word)) return { misspelt: false, suggestions: [] };
        const wrong = await host.misspelt(lang, [word]);
        if (!wrong.has(word)) return { misspelt: false, suggestions: [] };
        return { misspelt: true, suggestions: await host.suggest(lang, word, 5) };
      }

      case 'ignoreAll': {
        const word = String(a.word || '').trim();
        if (word) stateOf(session).ignored.add(word);
        return { ignored: [...stateOf(session).ignored] };
      }

      case 'addWord': {
        const word = String(a.word || '').replace(/’/g, "'").trim();
        if (!word) return { words: readWords() };
        const words = writeWords([...readWords(), word]);
        teachChromium(win, word, true);
        return { words };
      }

      case 'dictionary': {
        // The Dictionary dialog: the list, or the list as it now is.
        if (Array.isArray(a.words)) {
          const before = new Set(readWords());
          const words = writeWords(a.words);
          const after = new Set(words);
          for (const w of after) if (!before.has(w)) teachChromium(win, w, true);
          for (const w of before) if (!after.has(w)) teachChromium(win, w, false);
          return { words };
        }
        return { words: readWords() };
      }

      case 'altTextOf':
        return altTextOf(session, a);

      case 'dictionaryImport': {
        // A .dic word list (or any list of words, one to a line) added to the person's own.
        const words = fromDic(fs.readFileSync(String(a.path), 'utf8'));
        return { added: words, words: readWords() };
      }

      case 'dictionaryExport': {
        const list = Array.isArray(a.words) ? a.words : readWords();
        fs.writeFileSync(String(a.path), toDic(list), 'utf8');
        return { count: list.length };
      }

      case 'options': {
        if (a.options && stores?.settings) stores.settings.set(OPTIONS_KEY, { ...readOptions(), ...a.options });
        return { options: readOptions() };
      }

      case 'timings':
        return { ...host.timings };

      default:
        throw new Error(`no proofing action "${action}"`);
    }
  }

  return { ops, act, close: () => host.close() };
}
