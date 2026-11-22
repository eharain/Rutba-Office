// Reporting someone, beside blocking them (Office 2.28). A report says why,
// in the person's own words if they want, and carries the last messages that
// someone sent only when the person ticks so: Talk's messages are not read by
// anyone otherwise. The person reported is never told. Blocking is a second
// tick, since stopping someone and telling Rutba about them are two things.
//
// Nothing is chosen for the person (the screen review of 2026-11-22): no
// reason is ticked until they tick one. The tick for messages says how many
// at most and written where. A block that was asked for and did not happen
// is said, never passed over with a thank-you.

import React, { useState } from 'react';
import { Button, Dialog } from '@rutba/office-ui';
import { t } from '../talk-words.js';
import { refusalText, reportMessagesText, reportReasons } from './lines.js';

/** `requestId` when the person reported is someone asking to be a contact: there are no messages from them to send, and a block puts their request away. */
export function Report({ shell, person, groupId = null, requestId = null, toast, onClose, onBlocked }) {
  const name = person.name || person.email || '';
  const [reason, setReason] = useState(null);
  const [note, setNote] = useState('');
  const [withMessages, setWithMessages] = useState(false);
  const [block, setBlock] = useState(false);
  const [busy, setBusy] = useState(false);
  const send = async () => {
    setBusy(true);
    try {
      const r = await shell.talk.report({ about: person.id, reason, note: note.trim(), groupId, withMessages: withMessages && !requestId });
      if (r.error) { toast(r.error === 'too-many' ? t('talk.report.sent-many-reports') : refusalText(r.error), { tone: 'bad', ms: 5000 }); return; }
      let blocked = false;
      if (block) blocked = !(await shell.talk.block({ contactId: person.id, ...(requestId ? { requestId } : {}) }).catch(() => ({ error: 'server' }))).error;
      // The report went; a block that was asked for and did not happen is said, so nobody thinks they are rid of someone they are not.
      if (block && !blocked) toast(t('talk.report.reported-name-could', { name }), { tone: 'warn', ms: 7000 });
      else toast(blocked ? t('talk.report.reported-name-blocked', { name }) : t('talk.report.reported-thank-you'), { tone: 'good', ms: 4000 });
      if (blocked) onBlocked?.(person);
      onClose();
    } catch (err) {
      toast(err?.message || String(err), { tone: 'bad', ms: 5000 });
    } finally {
      setBusy(false);
    }
  };
  return (
    <Dialog
      title={t('talk.report.report-name', { name })}
      width={460}
      className="tk-report"
      onClose={onClose}
      actions={<><Button label={t('common.cancel')} onClick={onClose} /><Button primary className="tk-report-send" label={t('talk.report.send-report')} disabled={busy || !reason} onClick={send} /></>}
    >
      <p className="tk-lead">{t('talk.report.report-goes-rutba', { name })}</p>
      <fieldset className="tk-how">
        <legend>{t('talk.report.what-is-wrong')}</legend>
        {reportReasons().map(([key, label]) => (
          <label key={key} className="tk-relay" data-reason={key}><input type="radio" name="reason" checked={reason === key} onChange={() => setReason(key)} /><span><strong>{label}</strong></span></label>
        ))}
      </fieldset>
      <label className="tk-field"><span>{t('talk.report.happened-want-say')}</span>
        <textarea className="tk-ie-note tk-report-note" rows={3} maxLength={1000} value={note} onChange={(e) => setNote(e.target.value)} />
      </label>
      {requestId ? null : <label className="tk-tick"><input type="checkbox" className="tk-report-messages" checked={withMessages} onChange={(e) => setWithMessages(e.target.checked)} /><span>{reportMessagesText(name, { group: Boolean(groupId) })}</span></label>}
      <label className="tk-tick"><input type="checkbox" className="tk-report-block" checked={block} onChange={(e) => setBlock(e.target.checked)} /><span>{t('talk.report.block-name-well', { name })}</span></label>
    </Dialog>
  );
}
