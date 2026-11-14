// Calendar and Contacts: rooms and resources (2.25). Contacts' New room makes
// a card marked as a room, with its mailbox, seats, place and equipment, and
// the card says so. The event editor's "Add a room" lists the rooms for the
// time chosen, a room already booked then marked Busy and listed after one
// that is not; Add puts the room's mailbox on the Invite line and its name in
// Location; and the meeting saved has the room on it as CUTYPE=ROOM. Contacts'
// Invite to Talk is greyed on a room and ready on a person with an address. Run alone
// with RUTBA_VERIFY_ONLY=calendar-rooms.

import { labelKeyed } from './verify-keys.js';

export async function verifyCalendarRooms({ open, check, until, wait, errorsIn }) {
  const made = [];
  let booking = null;
  let saved = null;
  let people = null;
  let win = null;
  try {
    people = await open('contacts');
    const pjs = (code) => people.webContents.executeJavaScript(code);
    const ptype = (sel, value) => pjs(`(() => { const i = document.querySelector(${JSON.stringify(sel)}); if (!i) return false; Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(i, ${JSON.stringify(value)}); i.dispatchEvent(new Event('input', { bubbles: true })); return true; })()`);
    const pclick = (sel, text) => pjs(`(() => { const b = [...document.querySelectorAll(${JSON.stringify(sel)})].find((x) => x.textContent.trim() === ${JSON.stringify(text)}); b?.click(); return Boolean(b); })()`);

    // New room in Contacts: a card that is a room, with its mailbox, seats, place and kit.
    const ribbon = await pclick('.rw-ribbon .rw-btn', 'New room');
    await until(() => pjs(`Boolean(document.querySelector('.ct-editor .ct-booking'))`), 'the room editor', 5000).catch(() => {});
    const kind = await pjs(`document.querySelector('.ct-editor .ct-kind')?.value ?? null`);
    await ptype('.ct-editor .ct-name', 'Check board room');
    await ptype('.ct-editor .ct-capacity', '10');
    await ptype('.ct-editor .ct-where', 'Floor 3');
    await ptype('.ct-editor .ct-equipment-input', 'Screen, Whiteboard');
    await ptype('.ct-editor input[aria-label="Email"]', 'board-check@rooms.example');
    await wait(150);
    await pclick('.ct-editor .ct-actions .rw-btn', 'Save');
    await until(() => pjs(`Boolean(document.querySelector('.ct-room-line'))`), 'the room card', 5000).catch(() => {});
    const card = await pjs(`({ line: document.querySelector('.ct-room-line')?.textContent || '', kit: [...document.querySelectorAll('.ct-equipment .chip')].map((c) => c.textContent) })`);
    const listed = await pjs(`window.rutbaOffice.contacts.bookable()`);
    const board = listed.find((r) => r.email === 'board-check@rooms.example');
    if (board) made.push(board.id);
    check('calendar rooms: New room in Contacts makes a card marked as a room, with its mailbox, seats, place and equipment, and the card says so',
      ribbon && kind === 'room' && board?.kind === 'room' && board.capacity === 10 && board.where === 'Floor 3' && board.equipment.join('|') === 'Screen|Whiteboard'
        && /Room/.test(card.line) && /10 seats/.test(card.line) && /Floor 3/.test(card.line) && card.kit.join('|') === 'Screen|Whiteboard',
      JSON.stringify({ ribbon, kind, board, card }));

    // Invite to Talk: offered with Talk installed, greyed on a room, live on a person with an address.
    const talkButton = () => pjs(`(() => { const b = document.querySelector('.ct-talk-invite'); return b ? { disabled: b.disabled, tip: b.dataset.tip || b.title || '', label: b.textContent.trim() } : null; })()`);
    const talkHere = await pjs(`Boolean(window.rutbaOffice.talk)`);
    const onRoom = await talkButton();
    const person = await pjs(`window.rutbaOffice.contacts.save({ contact: { name: { full: 'Check Talk person', given: 'Check', family: 'Talk person', middle: '', prefix: '', suffix: '' }, emails: [{ value: 'talk-check@people.example', type: 'work', pref: true, label: null }], phones: [], addresses: [], urls: [], categories: [], extra: [] } })`);
    if (person?.id) made.push(person.id);
    await until(() => pjs(`[...document.querySelectorAll('.ct-item .name')].some((n) => n.textContent === 'Check Talk person')`), 'the person in the list', 5000).catch(() => {});
    await pjs(`[...document.querySelectorAll('.ct-item')].find((b) => b.querySelector('.name')?.textContent === 'Check Talk person')?.click(), 1`);
    await wait(200);
    const onPerson = await talkButton();
    check('contacts talk: Invite to Talk is on the card\'s ribbon with Talk installed, greyed on a room and saying why, and ready on a person with an address',
      talkHere ? onRoom?.label === 'Invite to Talk' && onRoom.disabled && /booked, not asked into Talk/.test(onRoom.tip) && onPerson && !onPerson.disabled && /Asks Check Talk person to connect in Talk/.test(onPerson.tip)
        : onRoom === null && onPerson === null,
      JSON.stringify({ talkHere, onRoom, onPerson }));
    // The list's right-click menu offers it too, for the card right-clicked: ready on the person, greyed on the room.
    const menuFor = async (name) => {
      await pjs(`(() => { const row = [...document.querySelectorAll('.ct-item')].find((b) => b.querySelector('.name')?.textContent === ${JSON.stringify(name)}); if (!row) return 0; const r = row.getBoundingClientRect(); row.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true, button: 2, clientX: r.left + 30, clientY: r.top + 10 })); return 1; })()`);
      await until(() => pjs(`Boolean(document.querySelector('.rw-menu'))`), 'the menu', 3000).catch(() => {});
      const item = await pjs(`(() => { const b = ((list) => list.find((x) => ${labelKeyed("Invite to Talk")}(x)) || list.find((x) => x.textContent.trim() === 'Invite to Talk'))([...document.querySelectorAll('.rw-menu button')]); return { items: [...document.querySelectorAll('.rw-menu button')].map((x) => x.textContent.trim()), talk: b ? { disabled: b.disabled || b.getAttribute('aria-disabled') === 'true' } : null }; })()`);
      await pjs(`document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true })), 1`);
      await until(() => pjs(`!document.querySelector('.rw-menu')`), 'the menu to close', 2000).catch(() => {});
      return item;
    };
    const personMenu = await menuFor('Check Talk person');
    const roomMenu = await menuFor('Check board room');
    check('contacts talk: the list\'s right-click menu offers Invite to Talk for the card right-clicked, ready on a person and greyed on a room',
      talkHere ? personMenu.talk && !personMenu.talk.disabled && roomMenu.talk?.disabled === true && personMenu.items.includes('Send mail')
        : !personMenu.talk && !roomMenu.talk,
      JSON.stringify({ talkHere, personMenu, roomMenu }));

    // A second room, made straight through Contacts, free all day.
    const quiet = await pjs(`window.rutbaOffice.contacts.save({ contact: { name: { full: 'Check quiet room', given: 'Check quiet room', family: '', middle: '', prefix: '', suffix: '' }, emails: [{ value: 'quiet-check@rooms.example', type: 'work', pref: true, label: null }], phones: [], addresses: [], urls: [], categories: [], kind: 'room', room: { capacity: 4, where: null, equipment: [] }, extra: [] } })`);
    made.push(quiet.id);

    win = await open('calendar');
    const js = (code) => win.webContents.executeJavaScript(code);
    win.setContentSize(1456, 910);
    const click = (sel, text) => js(`(() => { const b = [...document.querySelectorAll(${JSON.stringify(sel)})].find((x) => x.textContent.trim() === ${JSON.stringify(text)}); b?.click(); return Boolean(b); })()`);
    const type = (sel, value) => js(`(() => { const i = document.querySelector(${JSON.stringify(sel)}); if (!i) return false; Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(i, ${JSON.stringify(value)}); i.dispatchEvent(new Event('input', { bubbles: true })); return true; })()`);

    await click('.rw-ribbon .rw-btn', 'New event');
    await until(() => js(`Boolean(document.querySelector('.cal-invitees'))`), 'the editor', 5000).catch(() => {});
    await type('.cal-form .cal-title', 'Room check meeting');
    // The board room already booked at the editor's own time, by a meeting here with it on.
    booking = await js(`(async () => {
      const f = document.querySelector('.cal-form');
      // The editor's own start and end, each on its own day: after 23:00 the hour it offers ends tomorrow.
      const [date, endDate = date] = [...f.querySelectorAll('input[type=date]')].map((i) => i.value);
      const [from, to] = [...f.querySelectorAll('input[type=time]')].map((i) => i.value);
      const tz = Intl.DateTimeFormat().resolvedOptions().timeZone;
      const at = (d, t) => new Date(d + 'T' + t + ':00').getTime();
      const cal = (await window.rutbaOffice.calendar.calendars())[0];
      const e = await window.rutbaOffice.calendar.save({ calendarId: cal.id, event: { summary: 'Room check booking', start: { date, time: from + ':00', tzid: tz, utc: false, allDay: false, at: at(date, from) }, end: { date: endDate, time: to + ':00', tzid: tz, utc: false, allDay: false, at: at(endDate, to) }, attendees: [{ email: 'board-check@rooms.example', name: 'Check board room', type: 'ROOM', role: 'NON-PARTICIPANT', partstat: 'ACCEPTED' }] } });
      return e?.id || null;
    })()`);
    await click('.rw-dialog .cal-add-room', 'Add a room');
    await until(() => js(`document.querySelectorAll('.cal-rooms .cal-room').length >= 2`), 'the rooms', 6000).catch(() => {});
    const rows = await js(`[...document.querySelectorAll('.cal-rooms .cal-room')].map((r) => ({ email: r.dataset.email, state: r.querySelector('.cal-room-state')?.textContent || '', busy: r.classList.contains('busy') }))`);
    const ours = rows.filter((r) => /-check@rooms\.example$/.test(r.email));
    check('calendar rooms: Add a room lists the rooms for the time chosen, the one already booked then marked Busy and listed after the one that is not',
      ours.length === 2 && ours[0].email === 'quiet-check@rooms.example' && !ours[0].busy && ours[1].email === 'board-check@rooms.example' && ours[1].busy && ours[1].state === 'Busy'
        && rows.indexOf(ours[0]) < rows.indexOf(ours[1]),
      JSON.stringify(rows));

    // Add the quiet room: its mailbox on the Invite line, its name the place.
    await js(`[...document.querySelectorAll('.cal-rooms .cal-room')].find((r) => r.dataset.email === 'quiet-check@rooms.example')?.querySelector('.cal-room-add')?.click(), 1`);
    await wait(200);
    const form = await js(`({ invite: document.querySelector('.cal-invitees')?.value || '', where: document.querySelector('.cal-form input[placeholder="Where"]')?.value || '', added: [...document.querySelectorAll('.cal-rooms .cal-room')].find((r) => r.dataset.email === 'quiet-check@rooms.example')?.querySelector('.cal-room-add')?.textContent || '' })`);
    check('calendar rooms: Add puts the room\'s mailbox on the Invite line and its name in Location, and marks it Added',
      /quiet-check@rooms\.example/.test(form.invite) && form.where === 'Check quiet room' && form.added === 'Added', JSON.stringify(form));

    // Saved (no account here to send from): the meeting has the room on it as a room.
    await click('.rw-dialog .rw-btn', 'Save');
    await wait(400);
    saved = await js(`(async () => { const all = await window.rutbaOffice.calendar.events({ from: Date.now() - 86400000, to: Date.now() + 40 * 86400000 }); const s = all.find((e) => e.summary === 'Room check meeting'); return s ? window.rutbaOffice.calendar.get({ id: s.id }) : null; })()`);
    const line = saved?.attendees?.find((a) => a.email === 'quiet-check@rooms.example');
    check('calendar rooms: the meeting saved has the room on it as a room (CUTYPE=ROOM), not a person to answer for',
      line?.type === 'ROOM' && line.role === 'NON-PARTICIPANT' && line.name === 'Check quiet room', JSON.stringify(saved?.attendees));
    const complaints = [...(await errorsIn(win)), ...(await errorsIn(people))];
    check('calendar rooms: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('calendar rooms: the checks ran', false, err.stack || err.message);
  } finally {
    const api = win || people;
    if (api) {
      const run = (code) => (api.isDestroyed() ? null : api.webContents.executeJavaScript(`${code}.then(() => 1, () => 1)`).catch(() => null));
      for (const id of [booking, saved?.id].filter(Boolean)) await run(`window.rutbaOffice.calendar.remove({ id: ${JSON.stringify(id)}, scope: 'all' })`);
      for (const id of made) await run(`window.rutbaOffice.contacts.remove({ id: ${JSON.stringify(id)} })`);
    }
    for (const w of [win, people]) if (w && !w.isDestroyed()) w.destroy();
  }
}
