// The words of the messages the calendar sends, in the window's language: an
// invitation, a change, a cancellation, an answer, a room's reply. The main
// process writes the messages (main/calendar-text.js, whose English these
// are, key for key); a window hands these over with each message it asks
// for, and the last ones handed over are what a room's mailbox answers in.

import { t, language } from '@rutba/office-ui';

export function mailWords() {
  return {
    language: language(),
    inviteSubject: t('calendar.words.invitation-summary'),
    updateSubject: t('calendar.words.updated-invitation'),
    cancelSubject: t('calendar.cancelled-summary'),
    acceptedSubject: t('calendar.words.accepted-summary'),
    tentativeSubject: t('calendar.words.tentative-summary'),
    declinedSubject: t('calendar.words.declined-summary'),
    when: t('calendar.words.when-when'),
    where: t('calendar.words.where-where'),
    organizer: t('calendar.words.organizer-name'),
    invited: t('calendar.words.invited-names'),
    allDay: t('calendar.date-all-day'),
    attached: t('calendar.words.invitation-attached'),
    updated: t('calendar.words.meeting-changed-new'),
    cancelled: t('calendar.words.meeting-cancelled'),
    acceptedBody: t('calendar.words.name-accepted'),
    tentativeBody: t('calendar.words.name-answered-tentative'),
    declinedBody: t('calendar.words.name-declined'),
    roomBooked: t('calendar.words.room-booked'),
    roomClash: t('calendar.words.room-not-free'),
    roomFreeThatDay: t('calendar.words.free-day-times'),
    roomNoneFree: t('calendar.words.no-free-time'),
    roomFreeSubject: t('calendar.words.free-times-room'),
    roomFreeOn: t('calendar.words.room-date-free'),
    roomBusyAllDay: t('calendar.words.room-date-no'),
    roomAsk: t('calendar.words.ask-room-free'),
    roomOffice: t('calendar.words.answer-comes-rutba'),
    freeBusySubject: t('calendar.words.free'),
    freeBusyBody: t('calendar.words.asks-calendar-busy'),
    sendNoAccount: t('calendar.words.not-sent-no'),
    sendCannot: t('calendar.words.not-sent-mail'),
    sendFailed: t('calendar.words.not-sent-reason'),
  };
}
