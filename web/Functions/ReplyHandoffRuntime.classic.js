// Event-based-activation runtime for OnNewMessageCompose (manifest.xml's
// Runtimes/LaunchEvent elements, #22), CLASSIC Outlook on Windows variant.
//
// This file runs in Outlook's stripped-down "JavaScript-only runtime" and
// must not use anything the other (web/ReplyHandoffRuntime.js) runtime can:
//   - no async/await (times the add-in out on builds before ~April 2024 /
//     Build 17425.20000)
//   - no ternary operator (prevents the add-in from loading on those builds)
//   - no import/ES modules (must be a single flat file; nothing bundled)
// Confirmed live (2026-09-02): this runtime has neither `document` nor
// `BroadcastChannel` -- so unlike web/ReplyHandoffRuntime.js (full browser
// runtime), this handler cannot receive or splice the decrypted payload
// itself. Its only job is to prompt the user to do it themselves, in one
// click: post an InsightMessage notification whose action opens
// web/ReplyHandoffPane.js -- a dedicated minimal pane that has a real
// browser runtime (and thus a real BroadcastChannel), completes the splice,
// and closes itself. See that file and CLAUDE.md's #22 section for the full
// picture, including why sessionData/showAsTaskpane/temp files were all
// considered and ruled out as alternatives.
//
// Office.onReady()/Office.initialize do NOT run in this context -- there is
// no shared bootstrap moment, every invocation is a cold start.

function onNewMessageComposeHandler(event) {
  function finish() {
    // Only called once we're done -- code after event.completed() is "not
    // guaranteed to run", and Outlook may tear the runtime down once it's
    // called.
    event.completed();
  }

  try {
    Office.context.mailbox.item.getComposeTypeAsync(function (result) {
      var composeType = null;
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        composeType = result.value.composeType;
      }

      // Office.MailboxEnums.ComposeType has exactly three values -- Reply,
      // NewMail, Forward -- Reply All is NOT a distinct value; it also
      // reports 'reply'. A new message or forward has no prior encrypted
      // content to hand off, so there's nothing to prompt for.
      if (composeType !== Office.MailboxEnums.ComposeType.Reply) {
        finish();
        return;
      }

      // Heuristic only, not the full scoping-ID fallback chain
      // web/js/pgp/reply-handoff-runtime-core.js's armReplyHandoffListener
      // itself does (conversationId, then item.inReplyTo on Mailbox 1.14+):
      // if there's no conversationId here, showing the notification would
      // just lead to a pane that finds nothing to listen on. The pane does
      // the real, complete check; this is only deciding whether it's worth
      // prompting at all.
      var conversationId = Office.context.mailbox.item.conversationId;
      if (!conversationId) {
        finish();
        return;
      }

      Office.context.mailbox.item.notificationMessages.replaceAsync('pgp_reply_handoff_insight', {
        type: Office.MailboxEnums.ItemNotificationMessageType.InsightMessage,
        message: 'A large decrypted reply is ready to be inserted into this message.',
        icon: 'icon16',
        actions: [
          {
            actionText: 'Insert decrypted reply',
            actionType: Office.MailboxEnums.ActionType.ShowTaskPane,
            commandId: 'msgComposeReplyHandoffButton',
          },
        ],
      }, function () {
        finish();
      });
    });
  } catch (e) {
    console.error('ReplyHandoffRuntime.classic: handler failed', e);
    finish();
  }
}

Office.actions.associate('onNewMessageComposeHandler', onNewMessageComposeHandler);
