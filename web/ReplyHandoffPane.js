// ReplyHandoffPane.js
// Dedicated minimal pane for the large-message reply-handoff (#22). Opened
// automatically via an InsightMessage notification's ShowTaskPane action
// (see Functions/ReplyHandoffRuntime.classic.js) on hosts where
// BroadcastChannel isn't available in the OnNewMessageCompose event
// handler's own runtime -- confirmed true for classic Outlook on Windows.
// A single click on that notification opens this pane, which arms the same
// listener MessageCompose.js's own pane uses (see
// web/js/pgp/reply-handoff-runtime-core.js), then closes itself once the
// handoff settles (success or timeout) via Office.context.ui.closeContainer()
// -- a plain Mailbox 1.5 API, not the SharedRuntime-gated Office.addin.hide()
// (which Outlook doesn't support at all -- see that module's docblock and
// the #22 plan for why).
//
// Deliberately minimal: no recipient list, no sign toggle, nothing the full
// Encrypt/Decrypt pane shows -- this pane's only job is to load the full
// browser-runtime JS environment (with a real BroadcastChannel) long enough
// to complete the splice, then get out of the way.

import { armReplyHandoffListener } from './js/pgp/reply-handoff-runtime-core.js';

const CLOSE_DELAY_MS = 900; // brief pause so the user can see the outcome before the pane closes

// Matches the id ReplyHandoffRuntime.classic.js posts its InsightMessage
// notification under -- same mail item, so this pane can dismiss it as soon
// as it opens (clicking the notification's action doesn't clear it by
// itself; without this it stays on screen indefinitely).
const INSIGHT_NOTIFICATION_ID = 'pgp_reply_handoff_insight';

function showStatus(message, type = 'info') {
  const bar = document.getElementById('status-bar');
  bar.className = `pgp-alert pgp-alert--${type}`;
  bar.textContent = message;
}

Office.onReady(async () => {
  Office.context.mailbox.item.notificationMessages.removeAsync(INSIGHT_NOTIFICATION_ID);

  const has110 = Office.context.requirements.isSetSupported('Mailbox', '1.10');
  const has114 = Office.context.requirements.isSetSupported('Mailbox', '1.14');

  await armReplyHandoffListener({
    has110,
    has114,
    onStatus: showStatus,
    onSettled: (result) => {
      if (result.success) {
        // Close automatically -- the reply window itself now shows the
        // spliced content, which is the user's real confirmation; no need
        // to make them close this pane by hand too.
        setTimeout(() => Office.context.ui.closeContainer(), CLOSE_DELAY_MS);
        return;
      }
      // On failure (timeout, splice miss, write failure, or an immediate
      // skip like "not a reply"/"no scoping ID"): stay open and show what
      // happened, rather than leaving the generic "Inserting..." message up
      // forever -- auto-closing here would also hide the one signal telling
      // the user something needs manual attention.
      showStatus(result.message, 'warning');
    },
  });
});
