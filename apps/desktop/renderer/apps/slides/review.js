// Presentations' side of Review → Check Accessibility and Spelling: a
// finding or a misspelt word is a shape on a slide (or the slide's notes);
// going to it turns to the slide and selects the shape. The reading-order
// fix opens the Layers pane beside the findings, as PowerPoint opens its
// Selection Pane. The panes, the pass and the dialogs are renderer/review.js.

import { useMemo, useRef } from 'react';
import { t } from '@rutba/office-ui';
import { useReview, wordAround } from '../../review.js';

export function useSlidesReview({ shell, doc, model, apply, toast, index, setIndex, setSelected, patchView, caret = null, commitText = null }) {
  const indexRef = useRef(index);
  indexRef.current = index;

  const turnTo = (where) => {
    if (where?.slide != null && where.slide !== indexRef.current) setIndex(where.slide);
    setSelected(where?.shape ?? null);
  };

  const adapter = useMemo(() => ({
    goTo: (where) => turnTo(where),
    async altText(target, props) {
      await apply({ op: 'setAltText', slide: target.slide, shape: target.shape, descr: props.descr, decorative: props.decorative });
    },
    async fix(issue, f) {
      switch (f.kind) {
        case 'tableHeader':
          return apply({ op: 'setHeaderRow', slide: f.target.slide, shape: f.target.shape });
        case 'textColour':
          return apply({ op: 'setShapeTextColour', slide: f.target.slide, shape: f.target.shape, colour: f.colour });
        default:
          return null;
      }
    },
    slideTitle: (target, title) => apply({ op: 'setSlideTitle', slide: target.slide, shape: target.shape, title }),
    openLayers(target) {
      turnTo({ slide: target.slide });
      patchView((v) => ({ pane: 'layers' }));
    },
    spellArgs: () => ({ slide: indexRef.current }),
    showWord: (found) => turnTo(found.where?.notes ? { slide: found.where.slide } : found.where),
    whereLabel: (found) => (found.where ? (found.where.notes ? t('Slide {number} notes', { number: found.where.slide + 1 }) : t('Slide {number}', { number: found.where.slide + 1 })) : ''),
    // The word where the caret was in a shape's text last edited on this slide.
    thesaurus: caret ? {
      word() {
        const at = caret.current;
        if (!at || at.row != null || at.slide !== indexRef.current) return null;
        const found = wordAround(at.text, at.start, at.end);
        return found ? { ...found, id: at.id, text: at.text } : null;
      },
      async replace(found, word) {
        const text = found.text.slice(0, found.start) + word + found.text.slice(found.end);
        await commitText?.(found.id, text);
        if (caret.current && caret.current.id === found.id) caret.current = { ...caret.current, text, start: found.start, end: found.start + word.length };
      },
    } : null,
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }), [apply, setIndex, setSelected, patchView]);

  return useReview({ shell, doc, model, apply, toast, adapter });
}
