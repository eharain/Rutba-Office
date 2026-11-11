// Talk's checks from Office 2.25 on, in the order they build on each other.
// They run at the end of the talk-manage check, inside Talk's section of the
// window checks, until the key list gives each its own key (talk-email, talk-connect).

import { verifyTalkEmail } from './verify-talk-email.js';
import { verifyTalkConnect } from './verify-talk-connect.js';

export async function verifyTalkNext(ctx) {
  await verifyTalkEmail(ctx);
  await verifyTalkConnect(ctx);
}
