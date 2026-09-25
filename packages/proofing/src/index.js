// @rutba/proofing — Review's tools for Word, Worksheets and Presentation.
//
//   Check Accessibility: accessibility.js holds the rules; word.js, sheet.js
//   and deck.js describe each kind of document for them and write the fixes.
//   Spelling: tokenize.js finds the words, speller.js and speller-host.js
//   hold the dictionaries, spelling.js walks a document's text from the
//   caret, and the same three app modules list the text and change it.

export { SEVERITIES, RULES, checkAccessibility, groupIssues, unclearLinkText, isDefaultSheetName, visualOrder } from './accessibility.js';
export { normaliseColour, luminance, contrastRatio, requiredRatio, isLargeText, readableOn, HIGHLIGHT_HEX } from './colour.js';
export { DECORATIVE_URI, findElements, readAltProps, writeAltProps } from './alt-text.js';
export { readTitle, writeTitle } from './core-props.js';
export { paragraphText, paragraphsIn, replaceInParagraph, textBoxesIn } from './wordml.js';
export {
  describeWord, wordDrawings, wordLanguage, setWordAltText, setWordTableHeader, removeWordParagraphs, setWordTitle,
  wordSegments, wordStart, replaceWordText,
} from './word.js';
export {
  describeSheet, sheetDrawings, sheetCells, setSheetAltText, setSheetTableHeader, sheetSegments, sheetStart, replaceSheetText,
} from './sheet.js';
export {
  describeDeck, deckLanguage, setDeckAltText, setDeckTableHeader, setDeckTextColour, setDeckSlideTitle, deckSegments, deckStart, replaceDeckText,
} from './deck.js';
export { tokenize, matchCase, DEFAULT_OPTIONS } from './tokenize.js';
export { dictionaryFor, chooseLanguage, loadDictionary, checkWords, suggestWords, LANGUAGE_NAMES } from './speller.js';
export { createSpellerHost } from './speller-host.js';
export { nextMisspelling, changeAllEdits, contextOf, acceptedBy, toDic, fromDic } from './spelling.js';
