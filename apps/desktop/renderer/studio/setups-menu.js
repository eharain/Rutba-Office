// The setups (A4): the button beside the panels and its menu, the setups newest
// used first with the one in use ticked, then New setup, Duplicate, Rename…,
// Delete…, then Save to a file…, Open a file…. And the dialog that says what
// was found when a setup file is opened on another computer.

import React, { useState } from 'react';
import { Button, Dialog, Icon, Menu, t } from '@rutba/office-ui';
import { menuOrder } from '@rutba/studio/setup';
import { NameDialog, ConfirmDialog } from './parts.js';
import { kindName } from './list.js';

const appWord = (app) => (app === 'broadcast' ? t('studio.live-broadcast') : t('common.recorder'));

/** @param {{ s: object }} p  the studio (window.js) */
export function SetupsButton({ s }) {
  const [menu, setMenu] = useState(null);
  const [ask, setAsk] = useState(null); // 'new' | 'rename' | 'delete' | 'duplicate'
  const { setup, index, busy } = s;
  if (!setup) return null;
  if (busy) {
    return <span className="st-setup-btn locked" data-tip={t('studio.setups-menu.setups-changed-not')}><span className="st-setup-word">{t('studio.setups-menu.setup')}</span> <b>{setup.name}</b></span>;
  }
  const list = index ? menuOrder(index) : [];
  const items = [
    ...list.map((x) => ({ label: x.name, key: appWord(x.app), checked: x.id === setup.id, run: () => s.openSetup(x.id) })),
    '-',
    { label: t('studio.new-setup'), icon: 'new', run: () => setAsk('new') },
    { label: t('common.duplicate'), icon: 'duplicate', run: () => setAsk('duplicate') },
    { label: t('common.rename-2'), icon: 'textbox', run: () => setAsk('rename') },
    { label: t('studio.setups-menu.delete'), icon: 'trash', disabled: list.length < 2, run: () => setAsk('delete') },
    '-',
    { label: t('studio.setups-menu.save-file'), icon: 'save', key: '.rstudio', run: s.saveToFile },
    { label: t('studio.setups-menu.open-a-file'), icon: 'open', run: s.openFromFile },
  ];
  return (
    <>
      <Button className="st-setup-btn" dropdown aria-haspopup="menu" onClick={(e) => { const r = e.currentTarget.getBoundingClientRect(); setMenu({ x: r.left, y: r.bottom + 4 }); }}>
        <span className="st-setup-word">{t('studio.setups-menu.setup')}</span> <b>{setup.name}</b>
      </Button>
      {menu ? <Menu x={menu.x} y={menu.y} items={items} label={t('studio.setups-menu.setups')} onClose={() => setMenu(null)} /> : null}
      {ask === 'new' ? <NameDialog title={t('studio.new-setup')} label={t('common.name')} value="" okLabel={t('common.create')} onOk={(name) => { setAsk(null); s.newSetupNamed(name); }} onClose={() => setAsk(null)} /> : null}
      {ask === 'duplicate' ? <NameDialog title={t('common.duplicate')} label={t('common.name')} value={t('studio.name-copy', { name: setup.name })} okLabel={t('common.duplicate')} onOk={(name) => { setAsk(null); s.duplicateNamed(name); }} onClose={() => setAsk(null)} /> : null}
      {ask === 'rename' ? <NameDialog title={t('common.rename')} label={t('common.name')} value={setup.name} okLabel={t('common.rename')} onOk={(name) => { setAsk(null); s.renameSetup(name); }} onClose={() => setAsk(null)} /> : null}
      {ask === 'delete' ? <ConfirmDialog title={t('studio.delete-name', { name: setup.name })} body={t('studio.setups-menu.setup-holds-go')} okLabel={t('common.delete')} onOk={s.deleteSetup} onClose={() => setAsk(null)} /> : null}
    </>
  );
}

/** Opening a setup file (the artboard's second half): what was found here, and what was not. */
export function OpenSetupDialog({ s }) {
  const { fileOpening } = s;
  if (!fileOpening) return null;
  const { setup, found, missing, mediaCount, streams } = fileOpening;
  const seen = (rec) => (rec.how === 'size' ? t('studio.setups-menu.found-screen-same') : rec.how === 'title' ? t('studio.setups-menu.found-title') : rec.how === 'program' ? t('studio.setups-menu.found-window-same') : rec.how === 'name' ? t('studio.setups-menu.found-name') : rec.how === 'main' ? t('studio.setups-menu.main-screen-stands') : t('studio.setups-menu.found'));
  return (
    <Dialog title={t('studio.setups-menu.open-a-setup')} width={560} className="st-open-setup" onClose={() => s.setFileOpening(null)} actions={<>
      <Button label={t('common.cancel')} onClick={() => s.setFileOpening(null)} />
      <Button primary label={t('studio.setups-menu.add-setups')} className="st-open-add" onClick={s.addOpenedFile} />
    </>}>
      <h1 className="st-open-title">{setup.name}</h1>
      <p className="st-open-sub">{missing.length ? t('studio.setups-menu.most-found-check') : t('studio.setups-menu.everything-needs-found')}</p>
      <div className="st-open-rows">
        {found.map((rec) => <div className="st-open-row" key={`f-${rec.id}`}><span className="st-open-ok"><Icon name="check" size={13} /></span><span className="st-open-name">{rec.name || kindName(rec.kind)}</span><span className="st-open-how">{seen(rec)}</span></div>)}
        {mediaCount ? <div className="st-open-row"><span className="st-open-ok"><Icon name="check" size={13} /></span><span className="st-open-name">{t('studio.setups-menu.pictures-and-music')}</span><span className="st-open-how">{t('studio.setups-menu.n-files-all', { n: mediaCount })}</span></div> : null}
        {missing.map((rec) => <div className="st-open-row lost" key={`m-${rec.id}`}><span className="st-open-no"><Icon name="info" size={13} /></span><span className="st-open-name">{rec.name || kindName(rec.kind)}<small>{t('studio.setups-menu.not-found-kept')}</small></span></div>)}
        {(streams || []).map((d) => <div className="st-open-row lost" key={`k-${d.id}`}><span className="st-open-no"><Icon name="lock" size={13} /></span><span className="st-open-name">{d.name}<small>{t('studio.setups-menu.key-needed-stream')}</small></span>{s.isLive ? <Button label={t('studio.setups-menu.add-key')} onClick={() => s.setLiveOpen(true)} /> : null}</div>)}
      </div>
    </Dialog>
  );
}
