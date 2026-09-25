// A slide's SVG, put on the page once per drawing.
//
// React 19 writes `innerHTML` again whenever a `dangerouslySetInnerHTML`
// object is new — and `{ __html: svg }` written inline is new on every
// render. For a slide that meant the whole drawing re-parsed each time the
// window re-rendered (every pointer move of a drag, every thumbnail that
// arrived, the presenter's clock each second), and anything done to the
// drawing's own elements — a shape the show has hidden, an animation
// running on it — thrown away with it. Memoised on the markup string, the
// element is written again only when the drawing itself changes.

import React from 'react';

export const Markup = React.memo(function Markup({ as: Tag = 'div', html, className, style }) {
  return <Tag className={className} style={style} dangerouslySetInnerHTML={{ __html: html || '' }} />;
});

/** A full-size box for a drawing inside a layer that is itself sized and moved. */
export const FILL = { width: '100%', height: '100%' };
