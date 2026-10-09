// Merging and splitting table cells is not recorded by Track Changes, in Word
// or here, and Word says so before it does either while recording. So does
// this: the person is asked, and nothing happens on Cancel.

import { t } from '@rutba/office-ui';

/** The table operations Track Changes does not record. */
export const UNTRACKED_TABLE_OPS = new Set(['mergeCells', 'mergeRight', 'splitCell']);

/** Whether to go on with table operation `kind`: asked only while recording, and only for one that is not recorded. */
export async function untrackedOk(shell, model, kind) {
  if (!model?.trackRevisions || !UNTRACKED_TABLE_OPS.has(kind)) return true;
  const { response } = await shell.dialog.message({
    type: 'warning',
    message: t('This action will not be marked as a change.'),
    detail: t('Merging and splitting cells is not recorded by Track Changes. Do you want to continue?'),
    buttons: [t('OK'), t('Cancel')],
    defaultId: 0,
    cancelId: 1,
  });
  return response === 0;
}
