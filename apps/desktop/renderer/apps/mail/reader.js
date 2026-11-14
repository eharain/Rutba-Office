// The reading pane.
//
// The body is drawn in a frame with an empty sandbox attribute — no scripts, no
// forms, no same-origin access, no navigation — under a policy that blocks every
// remote fetch. That is where most clients stop. This one also says what it
// blocked and who it belonged to, because "images blocked" is a fact about the
// software and "Mailchimp tried to record that you opened this" is a fact about
// the message, and only the second one is worth a person's attention.

import React, { useEffect, useMemo, useRef, useState } from 'react';
import { Button, Icon, Chip, formatWhen, formatBytes, t, tn } from '@rutba/office-ui';
import { bodyDocument, avatarFor, displayName } from './parts.js';
import { SmimeStrip, unsafeDecrypted } from './smime.js';
import { SafetySwitch } from './safety.js';

export default function Reader({
  message,
  zoom = 1,
  insight,
  remote,
  dark,
  plain,
  onRemote,
  onSaveAttachment,
  onOpenAttachment,
  onReply,
  onReplyAll,
  onForward,
  onUnsubscribe,
  onSender,
  // The calendar's card for a message with a calendar part (calendar-invite-card.js), drawn above the attachments.
  inviteCard = null,
  onKeepSender,
  inJunk = false,
  onNotJunk,
  onSmimeImport,
  onCertificates,
  notes = null,
  // The Safety switch (safety.js): on, copying and saving wait until it is turned off.
  safety = true,
  onSafety = null,
  onSource = null,
  onBlocked = null,
  onClose = null,
}) {
  const frameRef = useRef(null);
  const [height, setHeight] = useState(600);
  const [showTrackers, setShowTrackers] = useState(false);
  // An encrypted message not signed by someone trusted is drawn as plain text, nothing fetched (EFAIL).
  const guarded = unsafeDecrypted(message.smime);
  const doc = useMemo(() => bodyDocument(message, { remote: remote && !guarded, dark, plain: plain || guarded }), [message, remote, dark, plain, guarded]);
  const blocked = !remote && /data-blocked-remote=/.test(doc);
  const from = message.from?.[0];
  const avatar = avatarFor(from);

  // The frame has no scripts, so its height is measured from outside once it
  // has laid out. A same-origin read is impossible under the sandbox, so the
  // frame is simply given room and allowed to scroll itself.
  useEffect(() => {
    setHeight(600);
    const el = frameRef.current;
    if (!el?.parentElement) return undefined;
    const ro = new ResizeObserver(() => setHeight(el.parentElement?.clientHeight || 600));
    ro.observe(el.parentElement);
    return () => ro.disconnect();
  }, [message]);

  useEffect(() => setShowTrackers(false), [message]);

  const attachments = (message.attachments || []).filter((a) => !a.inline);
  const trackers = insight?.trackers;
  const auth = insight?.auth;

  return (
    <div className="ml-reader" role="article" aria-label={message.subject || t('mail.no-subject')} style={{ zoom: Math.abs(zoom - 1) > 0.001 ? zoom : undefined }}>
      <header className="ml-head">
        <h2>
          {message.flagged ? <Icon name="star" size={16} /> : null}
          <span>{message.subject || t('mail.no-subject')}</span>
        </h2>

        <div className="ml-meta">
          <span className="ml-avatar" style={{ background: avatar.colour }}>{avatar.initial}</span>
          <div className="ml-meta-text">
            <div>
              <button type="button" className="ml-linkchip" style={{ fontSize: 'inherit', padding: 0 }} onClick={() => onSender?.(from)}>
                <strong>{displayName(from)}</strong>
              </button>
              {from?.name && from?.address ? <span className="ml-addr"> &lt;{from.address}&gt;</span> : null}
              {from?.address && onKeepSender ? (
                <button type="button" className="ml-keep" title={t('mail.reader.keep-sender-contacts')} onClick={() => onKeepSender(from)}>
                  <Icon name="contacts" size={13} /> {t('mail.reader.keep')}
                </button>
              ) : null}
            </div>
            <div className="ml-to">
              {t('mail.reader.to-recipients', { recipients: (message.to || []).map((t) => t.name || t.address).join(', ') || t('mail.reader.undisclosed') })}
              {message.cc?.length ? ` · ${t('mail.reader.cc-recipients', { recipients: message.cc.map((c) => c.name || c.address).join(', ') })}` : ''}
            </div>
          </div>
          <span className="ml-date">{formatWhen(message.date, { long: true })}</span>
          {onSafety ? <SafetySwitch on={safety} onChange={onSafety} /> : null}
          <Button icon="reply" title={t('common.reply')} onClick={onReply} />
          <Button icon="replyAll" title={t('mail.reply-all')} onClick={onReplyAll} />
          <Button icon="forward" title={t('common.forward')} onClick={onForward} />
          {onSource ? <Button icon="eye" title={t('mail.reader.view-source-whole')} className="ml-source-btn" onClick={onSource} /> : null}
          {onClose ? <Button icon="close" title={t('common.close')} className="ml-reader-close" onClick={onClose} /> : null}
        </div>

        {inJunk ? (
          <div className="ml-junk-note">
            <Icon name="spam" size={15} />
            <span>{message.junk?.reason ? t('mail.reader.message-junk-because', { reason: message.junk.reason }) : t('mail.message-junk')}</span>
            <Button label={t('mail.not-junk')} onClick={onNotJunk} />
          </div>
        ) : null}
        {inviteCard}
        {attachments.length ? (
          <div className="ml-attachments">
            {attachments.map((a, i) => {
              const index = message.attachments.indexOf(a);
              return (
                <button
                  key={`${a.filename}-${i}`}
                  type="button"
                  className="ml-attachment"
                  onClick={() => onOpenAttachment?.(index, a)}
                  onContextMenu={(e) => {
                    e.preventDefault();
                    if (safety && onBlocked) onBlocked();
                    else onSaveAttachment(index);
                  }}
                  title={safety && onBlocked ? t('mail.reader.name-click-open-turn', { name: a.filename }) : t('mail.reader.name-click-open', { name: a.filename })}
                >
                  <Icon name="attach" size={14} />
                  <span className="name">{a.filename}</span>
                  <span className="size">{formatBytes(a.size)}</span>
                </button>
              );
            })}
            {attachments.length > 1 ? (
              <button type="button" className="ml-attachment ml-save-all" onClick={() => (safety && onBlocked ? onBlocked() : onSaveAttachment('all'))}>
                <Icon name="download" size={13} />
                <span className="name">{t('mail.reader.save-all')}</span>
              </button>
            ) : null}
          </div>
        ) : null}
        {/* An attached invoice's or receipt's details, read when the message is opened (docnote.js). */}
        {notes}

        {/*
          One strip, three facts: whether anyone is watching, whether the sender
          checks out, and whether there is a way off the list. Each is only shown
          when there is something to say about it.
        */}
        <div className="ml-strip">
          {blocked ? (
            <button type="button" className="ml-badge act warn" onClick={onRemote}>
              <Icon name="shield" size={13} />
              <span>
                {trackers?.watching
                  ? tn(trackers.watching, 'mail.reader.count-tracker-blocked')
                  : t('mail.reader.remote-images')}
              </span>
              <span style={{ opacity: 0.7 }}>· {t('mail.reader.show')}</span>
            </button>
          ) : remote ? (
            <span className="ml-badge">
              <Icon name="eye" size={13} />
              <span>{t('mail.reader.images-loaded')}</span>
            </span>
          ) : null}

          {trackers?.hosts?.length ? (
            <button type="button" className="ml-badge act" onClick={() => setShowTrackers((s) => !s)}>
              <Icon name="globe" size={13} />
              <span>
                {trackers.networks.length
                  ? trackers.networks.slice(0, 2).join(', ') + (trackers.networks.length > 2 ? ` +${trackers.networks.length - 2}` : '')
                  : tn(trackers.hosts.length, 'mail.reader.count-remote-host')}
              </span>
              <Icon name={showTrackers ? 'chevronUp' : 'chevronDown'} size={12} />
            </button>
          ) : null}

          {auth?.known ? (
            <span className={`ml-badge ${auth.ok ? 'good' : 'bad'}`} title={`SPF ${auth.spf || '—'} · DKIM ${auth.dkim || '—'} · DMARC ${auth.dmarc || '—'}`}>
              <Icon name={auth.ok ? 'lock' : 'info'} size={13} />
              <span>{auth.summary}</span>
            </span>
          ) : null}

          {insight?.unsubscribe ? (
            <button type="button" className="ml-badge act" onClick={() => onUnsubscribe?.(insight.unsubscribe)}>
              <Icon name="close" size={13} />
              <span>{t('mail.unsubscribe')}</span>
            </button>
          ) : null}

          {insight?.bulk && !insight?.unsubscribe ? (
            <span className="ml-badge">
              <Icon name="inbox" size={13} />
              <span>{t('mail.reader.bulk-mail')}</span>
            </span>
          ) : null}
        </div>

        {showTrackers && trackers?.hosts?.length ? (
          <div className="ml-trackers">
            {trackers.hosts.map((h) => (
              <div key={h.host} className="ml-tracker">
                <Icon name={h.pixels ? 'eye' : 'image'} size={13} />
                <span className="host">{h.host}</span>
                {h.network ? <span className="net">{h.network}</span> : null}
                {h.sameAsSender ? <Chip>{t('mail.reader.senders-own')}</Chip> : null}
                {h.pixels ? <span className="pix">{tn(h.pixels, 'mail.reader.count-tracking-pixel')}</span> : null}
              </div>
            ))}
          </div>
        ) : null}
      </header>

      {message.smime ? <SmimeStrip smime={message.smime} from={from?.address || ''} onImport={onSmimeImport} onCertificates={onCertificates} /> : null}

      <div className="ml-body">
        {/*
          sandbox with one token: no scripts, no forms, no same-origin, no
          top-level navigation. The frame draws, and a person's click on a link
          may ask for a window (allow-popups). The window refuses every one
          and hands the link to Mail, which opens it in the browser or warns
          first. With no scripts, nothing but a click can ask.
        */}
        <iframe ref={frameRef} title={t('mail.message')} sandbox="allow-popups" referrerPolicy="no-referrer" srcDoc={doc} style={{ height }} />
      </div>

      {message.internetHeaders || message.headers?.length ? (
        <details className="ml-headers">
          <summary>{t('mail.reader.message-headers')}</summary>
          <pre className="selectable">
            {message.internetHeaders || (message.headers || []).map((h) => `${h.name}: ${h.value}`).join('\n')}
          </pre>
        </details>
      ) : null}
    </div>
  );
}
