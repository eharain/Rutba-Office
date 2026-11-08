// The setups (A4): the button beside the panels and its menu, the setups newest
// used first with the one in use ticked, then New setup, Duplicate, Rename…,
// Delete…, then Save to a file…, Open a file…. And the dialog that says what
// was found when a setup file is opened on another computer.

import React, { useState } from 'react';
import { Button, Dialog, Icon, Menu, t } from '@rutba/office-ui';
import { menuOrder } from '@rutba/studio/setup';
import { NameDialog, ConfirmDialog } from './parts.js';
import { kindName } from './list.js';

const appWord = (app) => (app === 'broadcast' ? t('Live Broadcast') : t('Recorder'));

/** @param {{ s: object }} p  the studio (window.js) */
export function SetupsButton({ s }) {
  const [menu, setMenu] = useState(null);
  const [ask, setAsk] = useState(null); // 'new' | 'rename' | 'delete' | 'duplicate'
  const { setup, index, busy } = s;
  if (!setup) return null;
  if (busy) {
    return <span className="st-setup-btn locked" data-tip={t('Setups can be changed when not recording or live')}><span className="st-setup-word">{t('Setup:')}</span> <b>{setup.tools ? t('For the tools') : setup.name}</b></span>;
  }
  const list = index ? menuOrder(index) : [];
  const items = [
    ...list.map((x) => ({ label: x.name, key: appWord(x.app), checked: x.id === setup.id, run: () => s.openSetup(x.id) })),
    '-',
    { label: t('New setup'), icon: 'new', run: () => setAsk('new') },
    { label: t('Duplicate'), icon: 'duplicate', run: () => setAsk('duplicate') },
    { label: t('Rename…'), icon: 'textbox', run: () => setAsk('rename') },
    { label: t('Delete…'), icon: 'trash', disabled: list.length < 2, run: () => setAsk('delete') },
    '-',
    { label: t('Save to a file…'), icon: 'save', key: '.rstudio', run: s.saveToFile },
    { label: t('Open a file…'), icon: 'open', run: s.openFromFile },
  ];
  return (
    <>
      <Button className="st-setup-btn" dropdown aria-haspopup="menu" onClick={(e) => { const r = e.currentTarget.getBoundingClientRect(); setMenu({ x: r.left, y: r.bottom + 4 }); }}>
        <span className="st-setup-word">{t('Setup:')}</span> <b>{setup.name}</b>
      </Button>
      {menu ? <Menu x={menu.x} y={menu.y} items={items} label={t('Setups')} onClose={() => setMenu(null)} /> : null}
      {ask === 'new' ? <NameDialog title={t('New setup')} label={t('Name')} value="" okLabel={t('Create')} onOk={(name) => { setAsk(null); s.newSetupNamed(name); }} onClose={() => setAsk(null)} /> : null}
      {ask === 'duplicate' ? <NameDialog title={t('Duplicate')} label={t('Name')} value={t('{name} copy', { name: setup.name })} okLabel={t('Duplicate')} onOk={(name) => { setAsk(null); s.duplicateNamed(name); }} onClose={() => setAsk(null)} /> : null}
      {ask === 'rename' ? <NameDialog title={t('Rename')} label={t('Name')} value={setup.name} okLabel={t('Rename')} onOk={(name) => { setAsk(null); s.renameSetup(name); }} onClose={() => setAsk(null)} /> : null}
      {ask === 'delete' ? <ConfirmDialog title={t('Delete {name}?', { name: setup.name })} body={t('The setup and what it holds will go from this computer. This cannot be undone.')} okLabel={t('Delete')} onOk={s.deleteSetup} onClose={() => setAsk(null)} /> : null}
    </>
  );
}

/** Opening a setup file (the artboard's second half): what was found here, and what was not. */
export function OpenSetupDialog({ s }) {
  const { fileOpening } = s;
  if (!fileOpening) return null;
  const { setup, found, missing, mediaCount, streams } = fileOpening;
  const seen = (rec) => (rec.how === 'size' ? t('Found: the screen of the same size') : rec.how === 'title' ? t('Found by its title') : rec.how === 'program' ? t('Found: a window of the same program') : rec.how === 'name' ? t('Found by its name') : rec.how === 'main' ? t('The main screen stands in until you choose one') : t('Found'));
  return (
    <Dialog title={t('Open a setup')} width={560} className="st-open-setup" onClose={() => s.setFileOpening(null)} actions={<>
      <Button label={t('Cancel')} onClick={() => s.setFileOpening(null)} />
      <Button primary label={t('Add to my setups')} className="st-open-add" onClick={s.addOpenedFile} />
    </>}>
      <h1 className="st-open-title">{setup.name}</h1>
      <p className="st-open-sub">{missing.length ? t('Most of it was found here; check the ones below.') : t('Everything it needs was found here.')}</p>
      <div className="st-open-rows">
        {found.map((rec) => <div className="st-open-row" key={`f-${rec.id}`}><span className="st-open-ok"><Icon name="check" size={13} /></span><span className="st-open-name">{rec.name || kindName(rec.kind)}</span><span className="st-open-how">{seen(rec)}</span></div>)}
        {mediaCount ? <div className="st-open-row"><span className="st-open-ok"><Icon name="check" size={13} /></span><span className="st-open-name">{t('Pictures and music')}</span><span className="st-open-how">{t('{n} files, all in the setup file', { n: mediaCount })}</span></div> : null}
        {missing.map((rec) => <div className="st-open-row lost" key={`m-${rec.id}`}><span className="st-open-no"><Icon name="info" size={13} /></span><span className="st-open-name">{rec.name || kindName(rec.kind)}<small>{t('Not found here. Kept in its place until you choose one.')}</small></span></div>)}
        {(streams || []).map((d) => <div className="st-open-row lost" key={`k-${d.id}`}><span className="st-open-no"><Icon name="lock" size={13} /></span><span className="st-open-name">{d.name}<small>{t('Key needed. Stream keys never travel in a setup file.')}</small></span>{s.isLive ? <Button label={t('Add key…')} onClick={() => s.setLiveOpen(true)} /> : null}</div>)}
      </div>
    </Dialog>
  );
}
