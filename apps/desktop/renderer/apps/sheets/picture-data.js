// Data from Picture: a table in a picture read into cells, as Excel's Data
// from Picture does, on this computer.
//
// The picture (a file, or the clipboard's) is read on the main process's
// threads with the table pass of @rutba/ocr (its rows and columns from its
// ruling lines or its spacing, our own layout code), and shown here beside
// a preview grid that can be corrected cell by cell. Cells the recogniser
// was unsure of are marked to check. Insert Data puts the grid in at the
// selected cell as one step, which Undo takes back.

import React, { useCallback, useEffect, useRef, useState } from 'react';
import { Button, Dialog, Spinner, t, tn } from '@rutba/office-ui';
import { OcrLanguagePicker, chosenOcrLanguages, rememberOcrLanguages, ocrLanguageNames, pixelsOf, loadPicture } from '../../ocr-ui.js';
import { fileUrl } from '../../file-urls.js';

/** Below this, a cell's least sure word is worth a second look. */
const CHECK_BELOW = 75;

/**
 * @param {{ shell, source: { path?: string, clipboard?: boolean }, at: string,
 *   onInsert: (rows: string[][]) => void, onClose: () => void }} props
 */
export function DataFromPictureDialog({ shell, source, at, onInsert, onClose }) {
  const [url, setUrl] = useState(null);
  const [rows, setRows] = useState(null);
  const [sure, setSure] = useState([]);
  const [error, setError] = useState(null);
  const [reading, setReading] = useState(true);
  const [languages, setLanguages] = useState(chosenOcrLanguages);
  const [choosing, setChoosing] = useState(false);
  const picture = useRef(null);
  const job = useRef(0);

  // The picture: a file's, or the clipboard's as a PNG held in this window.
  useEffect(() => {
    let made = null;
    let live = true;
    (async () => {
      try {
        let href;
        if (source.clipboard) {
          const clip = await shell.clipboard.readImage();
          if (!clip?.bytes) throw new Error(t('There is no picture on the clipboard.'));
          made = URL.createObjectURL(new Blob([clip.bytes], { type: 'image/png' }));
          href = made;
        } else {
          href = fileUrl(source.path);
        }
        const img = await loadPicture(href);
        if (!live) return;
        picture.current = img;
        setUrl(href);
      } catch (err) {
        if (live) { setError(err.message); setReading(false); }
      }
    })();
    return () => { live = false; if (made) URL.revokeObjectURL(made); };
  }, [shell, source]);

  const read = useCallback(async (langs) => {
    if (!picture.current) return;
    const n = ++job.current;
    setReading(true);
    setError(null);
    try {
      const pixels = pixelsOf(picture.current);
      const result = await shell.ocr.recognize({ image: { width: pixels.width, height: pixels.height, data: pixels.data }, languages: langs, table: true, job: `picture-data-${n}` });
      if (n !== job.current) return;
      const cells = result.table?.cells || [];
      if (!cells.length) {
        setRows(null);
        setError(t('No table was found in this picture.'));
      } else {
        setRows(cells.map((r) => [...r]));
        setSure(result.table.confidence || []);
      }
    } catch (err) {
      if (n === job.current) setError(t('The table could not be read: {reason}', { reason: err.message }));
    } finally {
      if (n === job.current) setReading(false);
    }
  }, [shell]);

  useEffect(() => { if (url) read(languages); }, [url]); // eslint-disable-line react-hooks/exhaustive-deps

  const cancel = () => {
    job.current += 1;
    shell.ocr.cancel({ job: `picture-data-${job.current - 1}` }).catch(() => {});
    onClose();
  };

  const toCheck = rows ? rows.reduce((n, row, r) => n + row.filter((v, c) => v && (sure[r]?.[c] ?? 100) < CHECK_BELOW).length, 0) : 0;
  const names = ocrLanguageNames();
  const width = rows ? Math.max(...rows.map((r) => r.length)) : 0;

  return (
    <Dialog
      title={t('Data from Picture')}
      onClose={cancel}
      width={900}
      actions={(
        <>
          <Button label={t('Cancel')} onClick={cancel} />
          <Button primary label={t('Insert Data')} data-role="pd-insert" disabled={reading || !rows} onClick={() => { onInsert(rows); onClose(); }} />
        </>
      )}
    >
      <div className="pd-body" data-role="pd-dialog">
        <div className="pd-picture">
          {url ? <img src={url} alt="" draggable={false} /> : <Spinner />}
        </div>
        <div className="pd-side">
          {reading ? (
            <div className="pd-reading" data-role="pd-reading"><Spinner /> {t('Reading the table in the picture…')}</div>
          ) : error ? (
            <div className="pd-error" data-role="pd-error">{error}</div>
          ) : rows ? (
            <>
              <div className="pd-note" data-role="pd-note" data-check={toCheck}>
                {toCheck
                  ? tn(toCheck, 'Check the data before it goes in: {count} cell the reading was unsure of is marked.', 'Check the data before it goes in: {count} cells the reading was unsure of are marked.')
                  : t('Check the data before it goes in: words read from a picture can be wrong.')}
              </div>
              <div className="pd-grid-wrap">
                <table className="pd-grid" data-role="pd-grid" data-rows={rows.length} data-cols={width}>
                  <tbody>
                    {rows.map((row, r) => (
                      <tr key={r}>
                        {Array.from({ length: width }, (_, c) => (
                          <td key={c} data-check={row[c] && (sure[r]?.[c] ?? 100) < CHECK_BELOW ? 'yes' : undefined}>
                            <input
                              value={row[c] ?? ''}
                              aria-label={t('Row {row}, column {column}', { row: r + 1, column: c + 1 })}
                              onChange={(e) => {
                                const v = e.target.value;
                                setRows((all) => all.map((x, i) => (i === r ? Object.assign([...x], { [c]: v }) : x)));
                                setSure((all) => all.map((x, i) => (i === r ? Object.assign([...(x || [])], { [c]: 100 }) : x)));
                              }}
                            />
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <div className="pd-where">{t('The data goes in from the selected cell, {cell}.', { cell: at })}</div>
            </>
          ) : null}
          <div className="pd-langs">
            {choosing ? (
              <>
                <OcrLanguagePicker value={languages} onChange={setLanguages} />
                <div className="pd-langs-row">
                  <Button label={t('Read Again')} data-role="pd-read-again" disabled={reading} onClick={() => { rememberOcrLanguages(languages); setChoosing(false); read(languages); }} />
                </div>
              </>
            ) : (
              <div className="pd-langs-row">
                <span className="pd-langs-label">{t('Languages: {list}', { list: languages.map((c) => names[c]).join(', ') })}</span>
                <Button ghost label={t('Change…')} data-role="pd-languages" onClick={() => setChoosing(true)} />
              </div>
            )}
          </div>
        </div>
      </div>
    </Dialog>
  );
}

export const PICTURE_DATA_CSS = `
.pd-body { display: grid; grid-template-columns: minmax(220px, 300px) minmax(380px, 1fr); gap: 18px; min-height: 280px; max-height: 62vh; }
.pd-picture { display: grid; place-items: center; background: var(--sunken); border: 1px solid var(--line-soft); border-radius: var(--r-2); padding: 8px; overflow: auto; min-height: 0; }
.pd-picture img { max-width: 100%; max-height: 100%; object-fit: contain; display: block; box-shadow: 0 0 0 1px rgba(0,0,0,0.08); background: #fff; }
.pd-side { display: flex; flex-direction: column; gap: 10px; min-width: 0; min-height: 0; }
.pd-reading { display: flex; align-items: center; gap: 10px; font-size: 13px; color: var(--ink-2); padding: 20px 4px; }
.pd-reading .rw-spinner { width: 18px; height: 18px; }
.pd-error { font-size: 13px; color: var(--bad); padding: 12px 4px; }
.pd-note { font-size: 12.5px; line-height: 1.45; color: var(--ink-2); }
.pd-note[data-check]:not([data-check='0']) { color: var(--ink); }
.pd-grid-wrap { flex: 1; min-height: 0; overflow: auto; border: 1px solid var(--line); border-radius: var(--r-2); background: var(--surface); }
.pd-grid { border-collapse: collapse; font-size: 12.5px; min-width: 100%; }
.pd-grid td { border: 1px solid var(--line-soft); padding: 0; min-width: 72px; }
.pd-grid tr:first-child td { background: color-mix(in srgb, var(--ink) 4%, var(--surface)); font-weight: 600; }
.pd-grid input { width: 100%; box-sizing: border-box; border: 0; background: transparent; padding: 5px 7px; font: inherit; color: var(--ink); outline: none; }
.pd-grid input:focus { box-shadow: inset 0 0 0 2px var(--accent); background: var(--surface); }
.pd-grid td[data-check='yes'] { background: color-mix(in srgb, var(--warn) 22%, var(--surface)); }
.pd-where { font-size: 12px; color: var(--ink-3); }
.pd-langs { border-top: 1px solid var(--line-soft); padding-top: 8px; display: flex; flex-direction: column; gap: 8px; }
.pd-langs-row { display: flex; align-items: center; gap: 8px; font-size: 12.5px; color: var(--ink-2); }
.pd-langs-label { flex: 1; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
`;
