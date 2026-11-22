// What a network camera's trouble is called, for the source list.

import { t } from '@rutba/office-ui';

/** The sentence for a camera that is waiting or has failed, by the reader's code; null when there is nothing to say. */
export function netCameraTrouble(code) {
  switch (code) {
    case '':
    case null:
    case undefined:
      return null;
    case 'sign-in-needed':
      return t('studio.net-camera-words.camera-asks-user');
    case 'sign-in-refused':
      return t('studio.net-camera-words.camera-did-not');
    case 'sign-in-unsupported':
      return t('studio.net-camera-words.camera-asks-kind');
    case 'not-found':
      return t('studio.net-camera-words.camera-no-stream');
    case 'refused':
      return t('studio.net-camera-words.camera-refused-give');
    case 'redirect':
      return t('studio.net-camera-words.camera-sent-office');
    case 'h265':
      return t('studio.net-camera-words.address-sends-h');
    case 'codec':
    case 'no-video':
      return t('studio.net-camera-words.camera-sends-no');
    case 'no-tcp':
      return t('studio.net-camera-words.camera-does-not');
    case 'certificate':
      return t('studio.net-camera-words.cameras-certificate');
    case 'missing':
      return t('studio.net-camera-words.camera-not-set');
    case 'address':
      return t('studio.net-camera-words.not-camera-address');
    case 'decode':
      return t('studio.net-camera-words.computer-could-not');
    case 'protocol':
      return t('studio.net-camera-words.answers-address-not');
    default:
      return t('studio.net-camera-words.camera-not-answering');
  }
}
