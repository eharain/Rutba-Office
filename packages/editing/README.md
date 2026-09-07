# @rutba/editing

Undo, and the rules that make it feel right.

Undo is not a feature users ask for; it is one they assume, and its absence is
felt in the first minute. Three implementations of it would drift into three
different feels, so it lives here — knowing nothing about documents, sheets or
mail. A `state` is opaque: the caller decides what is worth saving and how to put
it back.

```js
import { History } from '@rutba/editing';

const history = new History();               // 100 entries, 1200ms coalesce window
history.record({ state, label: 'typing', group: 'type:' + block, meta: selection });

const step = history.undo(currentState, currentSelection);
if (step) restore(step.state, step.meta);

history.describe();  // { canUndo, canRedo, undoLabel, redoLabel, depth, redoDepth }
```

## The interesting part is granularity, not mechanism

An undo that steps back one character at a time is technically correct and
unusable. The rules:

- Consecutive records sharing a non-null `group` merge into one step. Pass
  `'type:' + paragraphIndex` for a document, `'cell:' + ref` for a sheet.
- A pause longer than the window ends the run even if nothing else changed —
  walking away mid-sentence should not make the whole paragraph one undo.
- `break()` ends the run immediately. Call it on a caret move: "type, click away,
  type" is two undos, and every editor behaves that way.
- Nested edits are the caller's to collapse. Keep an edit-depth counter so
  `insertText` delegating to `deleteSelection` records once, and a paste spanning
  three paragraphs is one step.

## Snapshots, not inverse operations

Inverse operations are the more elegant design and the wrong one here. Where the
model is a document held verbatim, a snapshot is a string you already have —
whereas an inverse-operation log needs every edit to know how to undo itself,
which is exactly the code that goes subtly wrong and corrupts a customer file.

What a snapshot *contains* is the caller's choice, and two surfaces already
choose differently for good reason: a document snapshots its one XML part, while
a sheet snapshots only the inputs of the cells an edit will touch, because
copying a 5MB workbook per keystroke would make typing quadratic.

The property that matters, and that both surfaces test: **typing and then undoing
leaves the file byte-identical.**

## Status

Plain ESM JavaScript, no build step, no dependencies. Lives in Rutba-Workspace
for now; belongs in Rutba-Platform once a second consumer arrives.
