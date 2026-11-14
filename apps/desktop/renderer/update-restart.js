// Restart and update, from any of the places that offer it.
import { t } from '@rutba/office-ui/messages';

/** Installs the update, or says what has to be saved first and leaves everything as it is. */
export async function restartToUpdate(shell) {
  const result = await shell.update.install();
  if (result?.unsaved?.length) {
    const names = result.unsaved.filter(Boolean);
    await shell.dialog.message({
      type: 'info',
      message: t('update-restart.save-work-before'),
      detail: `${names.length ? `${names.join('\n')}\n\n` : ''}${t('update-restart.updating-closes-every')}`,
      buttons: [t('common.ok')],
    });
  }
  return result;
}
