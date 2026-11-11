// The calendar's card in Mail's reading pane: a message that carries a
// calendar part (an invitation, a change to one, a cancellation, an answer)
// shows what it is and what to do with it.
//
//   import { InviteCard } from '../calendar-invite-card.js';
//   <InviteCard shell={shell} message={message} where={{ accountId, folder, id }} onCompose={(draft) => ...} />
//
// It draws nothing for a message without one. An invitation offers Accept,
// Tentative and Decline: the answer goes into the calendar and is mailed to
// the organizer from the address invited (calendar.answer), or, with no
// account of that address here, `onCompose` opens it as a message with the
// reply attached, ready to send. A change or a cancellation from the
// organizer, or an answer to this person's own meeting, is taken into the
// calendar as the card opens (calendar.takeMail, which takes each once), and
// the card says what it did. The class names are Mail's own card's (ml-invite),
// which Mail's styles draw.

import React, { useCallback, useEffect, useState } from 'react';
import { Button, Icon, Chip, useToast, t, tn, language } from '@rutba/office-ui';
import { mailWords } from './calendar-words.js';

const ANSWERED = () => ({ ACCEPTED: t('accepted'), TENTATIVE: t('tentative'), DECLINED: t('declined') });

/** Which of a message's attachments is its calendar part, or -1. */
export function calendarPartOf(message) {
  return (message?.attachments || []).findIndex((a) => /text\/calendar|application\/ics/i.test(a.type || '') || /\.ics$/i.test(a.filename || ''));
}

/** An event's time, as the card says it. */
function whenOf(first) {
  if (!first?.start) return '';
  if (first.allDay) return t('{date}, all day', { date: new Date(first.start.at).toLocaleDateString(language(), { weekday: 'short', day: 'numeric', month: 'short', timeZone: 'UTC' }) });
  const day = { weekday: 'short', day: 'numeric', month: 'short' };
  const time = { hour: '2-digit', minute: '2-digit' };
  const a = new Date(first.start.at);
  const b = new Date(first.end?.at ?? first.start.at);
  const sameDay = a.toDateString() === b.toDateString();
  return `${a.toLocaleString(language(), { ...day, ...time })} – ${sameDay ? b.toLocaleTimeString(language(), time) : b.toLocaleString(language(), { ...day, ...time })}`;
}

export function InviteCard({ shell, message, where, onCompose = null }) {
  const toast = useToast();
  const [card, setCard] = useState(null);
  const index = calendarPartOf(message);
  const key = where ? `${where.accountId}|${where.folder}|${where.id}` : null;

  useEffect(() => {
    let live = true;
    setCard(null);
    if (!message || index < 0 || !where || !shell?.calendar) return undefined;
    (async () => {
      try {
        const part = await shell.mail.attachmentText({ ...where, index });
        if (!part?.text) return;
        const seen = await shell.calendar.openFile({ text: part.text, name: part.name });
        const first = seen.first;
        // A change, a cancellation or an answer: into the calendar now, once.
        const taken = seen.method && seen.method !== 'PUBLISH' ? await shell.calendar.takeMail?.({ text: part.text, from: message.from?.[0]?.address || null }).catch(() => null) : null;
        // What else is on at that time, for an invitation still to answer.
        let clash = [];
        if (seen.method === 'REQUEST' && first?.start && !first.allDay && first.end) {
          const on = await shell.calendar.events({ from: first.start.at, to: first.end.at }).catch(() => []);
          clash = on.filter((e) => !e.allDay && e.start < first.end.at && e.end > first.start.at && e.summary !== first.summary).map((e) => e.summary);
        }
        if (!live) return;
        const me = seen.invitation?.me;
        const who = first?.attendees?.[0];
        setCard({
          text: part.text,
          method: seen.method,
          summary: first?.summary || part.name,
          when: whenOf(first),
          location: first?.location || '',
          organizer: first?.organizer?.name || first?.organizer?.email || '',
          answered: me?.partstat && me.partstat !== 'NEEDS-ACTION' ? me.partstat : null,
          reply: seen.method === 'REPLY' && who ? t('{name} {answer}', { name: who.name || who.email, answer: ANSWERED()[who.partstat] || String(who.partstat || t('replied')).toLowerCase() }) : null,
          taken,
          clash: [...new Set(clash)].slice(0, 3),
          removed: false,
        });
      } catch (err) {
        // A calendar part that will not read is an attachment like any other; say why on the console.
        console.error(`invitation: ${err.message}`);
      }
    })();
    return () => {
      live = false;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [message, key, shell]);

  const answer = useCallback(async (partstat) => {
    if (!card) return;
    try {
      const r = await shell.calendar.answer({ text: card.text, partstat, words: mailWords() });
      setCard((c) => (c ? { ...c, answered: partstat } : c));
      const to = r.organizer?.name || r.organizer?.email;
      if (r.sent) {
        toast(r.kept ? t('Added to your calendar, and your answer is sent to {name}.', { name: to }) : t('Declined, and your answer is sent to {name}.', { name: to }), { tone: 'good', ms: 6000 });
        return;
      }
      if (r.replyPath && r.organizer?.email) onCompose?.({ to: r.organizer.email, subject: r.subject, text: '', attachments: [{ filename: 'reply.ics', path: r.replyPath, size: r.replySize || 0 }] });
      toast(r.kept ? t('Added to your calendar; the reply is ready to send.') : t('Declined; the reply is ready to send.'), { tone: 'good', ms: 6000 });
    } catch (err) {
      toast(err.message, { tone: 'bad', ms: 8000 });
    }
  }, [card, shell, toast, onCompose]);

  const keep = useCallback(async () => {
    if (!card) return;
    try {
      const r = await shell.calendar.importFile({ text: card.text });
      toast(tn(r.added, '{count} added to {calendar}.', '{count} added to {calendar}.', { calendar: r.calendar }), { tone: 'good' });
    } catch (err) {
      toast(err.message, { tone: 'bad' });
    }
  }, [card, shell, toast]);

  const removeCancelled = useCallback(async () => {
    const id = card?.taken?.eventId;
    if (!id) return;
    try {
      await shell.calendar.remove({ id, scope: 'all' });
      setCard((c) => (c ? { ...c, removed: true } : c));
    } catch (err) {
      toast(err.message, { tone: 'bad' });
    }
  }, [card, shell, toast]);

  if (!card) return null;
  const answers = ANSWERED();
  const taken = card.taken;
  // What the card did with a change, a cancellation or an answer, said in a line.
  const note = !taken ? null
    : taken.kind === 'reply' ? t('The answer is on the meeting in your calendar.')
      : taken.kind === 'cancel' ? (card.removed ? t('Taken off your calendar.') : t('Marked cancelled in your calendar.'))
        : taken.kind === 'update' ? t('Your calendar has the new details.') : null;
  return (
    <div className={`ml-invite cal-card${card.method === 'CANCEL' ? ' cancelled' : ''}`} data-method={card.method || ''}>
      <style>{CSS}</style>
      <Icon name="calendar" size={18} />
      <div className="ml-invite-text">
        <div className="ml-invite-title">{card.method === 'CANCEL' ? t('Cancelled: {summary}', { summary: card.summary }) : card.summary}</div>
        <div className="ml-invite-when">{card.when}{card.location ? ` · ${card.location}` : ''}</div>
        <div className="ml-invite-who">{card.organizer ? t('From {organizer}', { organizer: card.organizer }) : ''}{card.answered ? ` · ${t('you answered {answer}', { answer: answers[card.answered] || card.answered.toLowerCase() })}` : ''}</div>
        {card.clash.length ? <div className="cal-card-clash">{t('At the same time: {events}', { events: card.clash.join(', ') })}</div> : null}
        {note ? <div className="cal-card-note">{note}</div> : null}
      </div>
      {card.method === 'REQUEST' ? (
        <div className="ml-invite-actions">
          <Button primary label={t('Accept')} onClick={() => answer('ACCEPTED')} />
          <Button label={t('Tentative')} onClick={() => answer('TENTATIVE')} />
          <Button label={t('Decline')} onClick={() => answer('DECLINED')} />
        </div>
      ) : card.method === 'REPLY' ? (
        <Chip>{card.reply || t('a reply')}</Chip>
      ) : card.method === 'CANCEL' ? (
        taken?.eventId && !card.removed ? <Button label={t('Remove from calendar')} onClick={removeCancelled} /> : null
      ) : (
        <Button label={t('Add to calendar')} onClick={keep} />
      )}
    </div>
  );
}

const CSS = `
.cal-card .cal-card-clash { font-size: 12.5px; color: var(--bad, #b3261e); }
.cal-card .cal-card-note { font-size: 12.5px; color: var(--ink-3); }
.cal-card.cancelled { border-left-color: var(--ink-3); }
`;

export default InviteCard;
