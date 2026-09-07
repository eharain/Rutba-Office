// The reading pane.
//
// The body is drawn in a frame with an empty sandbox attribute — no scripts, no
// forms, no same-origin access, no navigation — under a policy that blocks every
// remote fetch. That is where most clients stop. This one also says what it
// blocked and who it belonged to, because "images blocked" is a fact about the
// software and "Mailchimp tried to record that you opened this" is a fact about
// the message, and only the second one is worth a person's attention.

import React, { useEffect, useMemo, useRef, useState } from 'react';
import { Button, Icon, Chip, formatWhen, formatBytes } from '@rutba/office-ui';
import { bodyDocument, avatarFor, displayName } from './parts.js';

export default function Reader({
  message,
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
}) {
  const frameRef = useRef(null);
  const [height, setHeight] = useState(600);
  const [showTrackers, setShowTrackers] = useState(false);
  const doc = useMemo(() => bodyDocument(message, { remote, dark, plain }), [message, remote, dark, plain]);
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
    <div className="ml-reader">
      <header className="ml-head">
        <h2>
          {message.flagged ? <Icon name="star" size={16} /> : null}
          <span>{message.subject || '(no subject)'}</span>
        </h2>

        <div className="ml-meta">
          <span className="ml-avatar" style={{ background: avatar.colour }}>{avatar.initial}</span>
          <div className="ml-meta-text">
            <div>
              <button type="button" className="ml-linkchip" style={{ fontSize: 'inherit', padding: 0 }} onClick={() => onSender?.(from)}>
                <strong>{displayName(from)}</strong>
              </button>
              {from?.name && from?.address ? <span className="ml-addr"> &lt;{from.address}&gt;</span> : null}
            </div>
            <div className="ml-to">
              to {(message.to || []).map((t) => t.name || t.address).join(', ') || 'undisclosed recipients'}
              {message.cc?.length ? ` · cc ${message.cc.map((c) => c.name || c.address).join(', ')}` : ''}
            </div>
          </div>
          <span className="ml-date">{formatWhen(message.date, { long: true })}</span>
          <Button icon="reply" title="Reply" onClick={onReply} />
          <Button icon="replyAll" title="Reply all" onClick={onReplyAll} />
          <Button icon="forward" title="Forward" onClick={onForward} />
        </div>

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
                    onSaveAttachment(index);
                  }}
                  title={`${a.filename} — click to open, right-click to save`}
                >
                  <Icon name="attach" size={14} />
                  <span className="name">{a.filename}</span>
                  <span className="size">{formatBytes(a.size)}</span>
                </button>
              );
            })}
            {attachments.length > 1 ? (
              <button type="button" className="ml-attachment" onClick={() => onSaveAttachment('all')}>
                <Icon name="download" size={13} />
                <span className="name">Save all</span>
              </button>
            ) : null}
          </div>
        ) : null}

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
                  ? `${trackers.watching} tracker${trackers.watching === 1 ? '' : 's'} blocked`
                  : 'Remote images blocked'}
              </span>
              <span style={{ opacity: 0.7 }}>· show</span>
            </button>
          ) : remote ? (
            <span className="ml-badge">
              <Icon name="eye" size={13} />
              <span>Images loaded for this message</span>
            </span>
          ) : null}

          {trackers?.hosts?.length ? (
            <button type="button" className="ml-badge act" onClick={() => setShowTrackers((s) => !s)}>
              <Icon name="globe" size={13} />
              <span>
                {trackers.networks.length
                  ? trackers.networks.slice(0, 2).join(', ') + (trackers.networks.length > 2 ? ` +${trackers.networks.length - 2}` : '')
                  : `${trackers.hosts.length} remote host${trackers.hosts.length === 1 ? '' : 's'}`}
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
              <span>Unsubscribe</span>
            </button>
          ) : null}

          {insight?.bulk && !insight?.unsubscribe ? (
            <span className="ml-badge">
              <Icon name="inbox" size={13} />
              <span>Bulk mail</span>
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
                {h.sameAsSender ? <Chip>sender's own</Chip> : null}
                {h.pixels ? <span className="pix">{h.pixels} tracking pixel{h.pixels === 1 ? '' : 's'}</span> : null}
              </div>
            ))}
          </div>
        ) : null}
      </header>

      <div className="ml-body">
        {/*
          sandbox with no tokens: no scripts, no forms, no same-origin, no
          top-level navigation. The only thing this frame can do is draw.
        */}
        <iframe ref={frameRef} title="Message" sandbox="" referrerPolicy="no-referrer" srcDoc={doc} style={{ height }} />
      </div>

      {message.internetHeaders || message.headers?.length ? (
        <details className="ml-headers">
          <summary>Message headers</summary>
          <pre className="selectable">
            {message.internetHeaders || (message.headers || []).map((h) => `${h.name}: ${h.value}`).join('\n')}
          </pre>
        </details>
      ) : null}
    </div>
  );
}
