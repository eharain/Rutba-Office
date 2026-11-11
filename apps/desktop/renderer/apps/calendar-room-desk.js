// "Room mailboxes": which of Mail's accounts this copy of Office answers for
// as a room (calendar.setRoomMailbox, main/calendar-desk.js). A room's
// mailbox marked here books an invitation with no clash, declines one that
// clashes with the times free that day, and answers an enquiry; it does so
// only while Office is open on this computer, and the dialog says so plainly.

import React, { useCallback, useEffect, useState } from 'react';
import { Button, Dialog, Empty, tn, t } from '@rutba/office-ui';
import { mailWords } from './calendar-words.js';

export function RoomDeskDialog({ shell, toast, onClose }) {
  const [rows, setRows] = useState(null);
  const load = useCallback(() => shell.calendar.roomMailboxes().then(setRows).catch((err) => { setRows([]); toast(err.message, { tone: 'bad' }); }), [shell, toast]);
  useEffect(() => {
    load();
  }, [load]);

  const flip = async (row, room) => {
    try {
      await shell.calendar.setRoomMailbox({ accountId: row.id, room, words: mailWords() });
      // A room's card in Contacts, for "Add a room", when there is none for this address yet.
      if (room && shell.contacts?.bookable) {
        const known = await shell.contacts.bookable().catch(() => []);
        if (!known.some((r) => String(r.email).toLowerCase() === String(row.email).toLowerCase())) {
          const name = row.name || row.email;
          await shell.contacts.save({ contact: { name: { full: name, given: name, family: '', middle: '', prefix: '', suffix: '' }, emails: [{ value: row.email, type: 'work', pref: true, label: null }], phones: [], addresses: [], urls: [], categories: [], kind: 'room', room: { capacity: null, where: null, equipment: [] }, extra: [] } }).catch(() => null);
        }
      }
      toast(room ? t('{address} is answered as a room while Office is open.', { address: row.email }) : t('{address} is an ordinary mailbox again.', { address: row.email }), { tone: 'good' });
      load();
    } catch (err) {
      toast(err.message, { tone: 'bad' });
    }
  };

  return (
    <Dialog title={t('Room mailboxes')} width={520} className="cal-room-desk" onClose={onClose} actions={<Button primary label={t('Done')} onClick={onClose} />}>
      <style>{CSS}</style>
      <p className="cal-desk-note">{t('A Mail account marked as a room is answered by Office for the room: an invitation that clashes with nothing is accepted and booked, one that clashes is declined with the times free that day, and a message asking when it is free gets the free times.')}</p>
      <p className="cal-desk-warn">{t("This works only while this copy of Office is open on this computer. While it is closed, the mail waits in the room's Inbox and is answered when Office next opens.")}</p>
      {rows === null ? null : !rows.length ? (
        <Empty icon="mail" title={t('No Mail accounts')}>{t("Add the room's mailbox as an account in Mail, then mark it here.")}</Empty>
      ) : (
        <div className="cal-desk-list" role="list">
          {rows.map((row) => (
            <label key={row.id} role="listitem" className="cal-desk-row" data-account={row.id}>
              <input type="checkbox" checked={row.room} onChange={(e) => flip(row, e.target.checked)} />
              <span className="cal-desk-text">
                <b>{row.name || row.email}</b>
                <span>{row.email}</span>
              </span>
              <span className="cal-desk-state">{row.room ? tn(row.bookings, 'A room: {count} booking', 'A room: {count} bookings') : t('This mailbox is a room')}</span>
            </label>
          ))}
        </div>
      )}
    </Dialog>
  );
}

const CSS = `
.cal-desk-note { margin: 0 0 8px; font-size: 13px; color: var(--ink-2); }
.cal-desk-warn { margin: 0 0 12px; font-size: 13px; padding: 8px 10px; border-radius: var(--r-2); background: var(--sunken); color: var(--ink); }
.cal-desk-list { display: flex; flex-direction: column; gap: 2px; }
.cal-desk-row { display: flex; align-items: center; gap: 10px; padding: 6px 8px; border-radius: var(--r-2); cursor: pointer; }
.cal-desk-row:hover { background: var(--hover); }
.cal-desk-text { flex: 1; min-width: 0; display: flex; flex-direction: column; font-size: 13px; }
.cal-desk-text span { color: var(--ink-3); font-size: 12.5px; }
.cal-desk-state { font-size: 12px; color: var(--ink-3); white-space: nowrap; }
`;

export default RoomDeskDialog;
