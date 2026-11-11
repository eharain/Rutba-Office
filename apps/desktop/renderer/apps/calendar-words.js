// The words of the messages the calendar sends, in the window's language: an
// invitation, a change, a cancellation, an answer, a room's reply. The main
// process writes the messages (main/calendar-text.js, whose English these
// are, key for key); a window hands these over with each message it asks
// for, and the last ones handed over are what a room's mailbox answers in.

import { t, language } from '@rutba/office-ui';

export function mailWords() {
  return {
    language: language(),
    inviteSubject: t('Invitation: {summary}'),
    updateSubject: t('Updated invitation: {summary}'),
    cancelSubject: t('Cancelled: {summary}'),
    acceptedSubject: t('Accepted: {summary}'),
    tentativeSubject: t('Tentative: {summary}'),
    declinedSubject: t('Declined: {summary}'),
    when: t('When: {when}'),
    where: t('Where: {where}'),
    organizer: t('Organizer: {name}'),
    invited: t('Invited: {names}'),
    allDay: t('{date}, all day'),
    attached: t('The invitation is attached: a calendar program shows it with Accept, Tentative and Decline.'),
    updated: t('This meeting has changed. The new details are above and attached.'),
    cancelled: t('This meeting is cancelled. A calendar program takes it off the calendar with the file attached.'),
    acceptedBody: t('{name} accepted.'),
    tentativeBody: t('{name} answered tentative.'),
    declinedBody: t('{name} declined.'),
    roomBooked: t('{room} is booked for {when}.'),
    roomClash: t('{room} is not free for {when}.'),
    roomFreeThatDay: t('Free that day: {times}.'),
    roomNoneFree: t('It has no free time left that day.'),
    roomFreeSubject: t('Free times: {room}'),
    roomFreeOn: t('{room}, {date}: free {times}.'),
    roomBusyAllDay: t('{room}, {date}: no free time.'),
    roomAsk: t('To ask when {room} is free, write to it with a date, such as "free on 12 March?".'),
    roomOffice: t('This answer comes from Rutba Office, which looks after this room\'s mailbox while it is open.'),
    freeBusySubject: t('When are you free? {from} to {to}'),
    freeBusyBody: t('This asks your calendar when you are busy between {from} and {to}. A calendar program answers with busy times only, never what fills them.'),
    sendNoAccount: t('Not sent: there is no Mail account to send it from. Add one in Mail, under Accounts.'),
    sendCannot: t('Not sent: Mail cannot send from {address}, as that account has no outgoing server set up.'),
    sendFailed: t('Not sent: {reason}'),
  };
}
