// "Open these files with Rutba Office."
//
// The honest version of a button every free suite ships and most of them lie
// about. On Linux this really does set the handlers. On Windows and macOS the
// operating system reserves the choice for the person — deliberately, because a
// program that could seize file types silently would be malware — so the button
// takes them to the exact place where the choice is made and says so plainly
// rather than flashing a tick and changing nothing.
//
// The list underneath is read back from the system, not from what we wish were
// true, so it shows which formats actually open with us right now.

import React, { useCallback, useEffect, useState } from 'react';
import { Button, Dialog, Icon, Chip, Spinner } from '@rutba/office-ui';

export function useDefaults(shell) {
  const [status, setStatus] = useState(null);

  const refresh = useCallback(() => {
    let live = true;
    setStatus(null);
    shell.defaults
      .status()
      .then((s) => live && setStatus(s))
      .catch(() => live && setStatus({ formats: [], ours: 0, total: 0, canSet: false, instructions: '' }));
    return () => {
      live = false;
    };
  }, [shell]);

  useEffect(() => refresh(), [refresh]);
  return { status, refresh };
}

export function DefaultsDialog({ shell, onClose, toast }) {
  const { status, refresh } = useDefaults(shell);
  const [working, setWorking] = useState(false);

  const apply = useCallback(async () => {
    setWorking(true);
    try {
      const result = await shell.defaults.set({});
      toast?.(result.message, { tone: result.changed ? 'good' : 'plain', ms: 7000 });
      if (result.changed) refresh();
    } catch (err) {
      toast?.(err.message, { tone: 'bad' });
    } finally {
      setWorking(false);
    }
  }, [shell, toast, refresh]);

  return (
    <Dialog
      title="Open files with Rutba Office"
      width={560}
      onClose={onClose}
      actions={
        <>
          <Button label="Close" onClick={onClose} />
          <Button
            primary
            icon="check"
            label={status?.canSet ? 'Set them all' : 'Open system settings'}
            disabled={working || !status}
            onClick={apply}
          />
        </>
      }
    >
      {!status ? (
        <div style={{ padding: 30, display: 'grid', placeItems: 'center' }}>
          <Spinner />
        </div>
      ) : (
        <>
          <div className="ml-import-summary">
            <Chip>{status.ours} of {status.total} formats</Chip>
            <Chip>{status.platform}</Chip>
            {status.packaged ? null : <Chip>development build</Chip>}
          </div>

          <div className="ml-note" style={{ marginBottom: 12 }}>
            <Icon name="info" size={14} />
            <span>{status.instructions}</span>
          </div>

          {!status.packaged ? (
            <p className="rw-hint" style={{ marginTop: 0 }}>
              File types are registered by the installer. A build run from the source tree has not registered any, so
              this list will read empty until Rutba Office is installed.
            </p>
          ) : null}

          <div className="ml-import-folders" style={{ maxHeight: 260 }}>
            {status.formats.map((f) => (
              <div key={f.ext} className="ml-import-folder">
                <Icon name={f.isDefault ? 'check' : 'file'} size={14} />
                <span className="name">
                  .{f.ext} — {f.name}
                </span>
                <span className="count" title={f.handler || 'no handler recorded'}>
                  {f.isDefault ? 'Rutba Office' : f.handler ? shorten(f.handler) : '—'}
                </span>
              </div>
            ))}
          </div>
        </>
      )}
    </Dialog>
  );
}

/** A registry ProgId or a .desktop name, cut to something a column can hold. */
function shorten(handler) {
  const text = String(handler).replace(/\.desktop$/, '');
  return text.length > 26 ? `${text.slice(0, 24)}…` : text;
}
