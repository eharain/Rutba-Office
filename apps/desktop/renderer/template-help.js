// A document, workbook or presentation made from one of the gallery's
// templates has a small link to that template's page online, in its status
// bar (owner, 2026-10-22, D-2026-10-22-03: "the template them self should have
// little help link to go online and see the documentations"). The page's
// address comes with the template set, as the gallery has it.

import React, { useEffect, useState } from 'react';
import { Icon, t } from '@rutba/office-ui';

/** The status bar's link to the template this window was made from, or nothing. */
export function TemplateHelp({ shell }) {
  const [tpl, setTpl] = useState(null);
  useEffect(() => {
    const id = new URLSearchParams(location.search).get('template');
    if (!id || id === 'blank') return undefined;
    let live = true;
    shell.doc.gallery({}).then((g) => {
      const found = (g.templates || []).find((x) => x.id === id);
      if (live && found?.page) setTpl(found);
    }).catch(() => {});
    return () => { live = false; };
  }, [shell]);
  if (!tpl) return null;
  const label = t('About the {name} template online', { name: t(tpl.name) });
  return (
    <button type="button" className="chip rw-template-help" data-page={tpl.page} title={label} aria-label={label}
      onClick={() => shell.shell.openExternal({ url: tpl.page }).catch(() => {})}>
      <Icon name="globe" size={12} />
      <span>{t('{name} template', { name: t(tpl.name) })}</span>
    </button>
  );
}
