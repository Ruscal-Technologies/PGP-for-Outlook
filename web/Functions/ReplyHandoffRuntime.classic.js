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
//
// actionType is passed as the plain string 'showTaskPane' rather than
// Office.MailboxEnums.ActionType.ShowTaskPane -- the interface explicitly
// accepts `string | MailboxEnums.ActionType`, and this restricted runtime's
// Office.MailboxEnums object may not populate every enum the full browser
// runtime does (confirmed working here: ItemNotificationMessageType; NOT
// confirmed: ActionType) -- an undefined access there would throw inside
// this async callback, which the outer try/catch can't catch (it's already
// returned by the time the callback runs), silently killing the handler
// with no notification and no error surfaced anywhere. Every step below is
// wrapped so a failure is at minimum visible via a fallback notification,
// using only the InformationalMessage type already confirmed to work here.
//
// HANDOFF_PENDING_MARKER: this file can't import
// web/js/pgp/reply-handoff-channel.js (no ES modules here), so its exact
// literal value is duplicated below -- it MUST match that module's exported
// constant exactly, or this gate silently never fires. Checked in the body
// before posting the notification so it only appears on a reply this add-in
// actually opened for a handoff (via MessageRead.js's
// openNativeReplyWithHandoff), not on every reply to every encrypted
// message -- e.g. one the user replied to with Outlook's own Reply button,
// or one whose handoff already completed.
var HANDOFF_PENDING_MARKER = '=== PGP: Click Insert Decrypted Reply above to complete this reply ===';

function onNewMessageComposeHandler(event) {
  function finish() {
    // Only called once we're done -- code after event.completed() is "not
    // guaranteed to run", and Outlook may tear the runtime down once it's
    // called.
    event.completed();
  }

  function showFallbackInfo(message) {
    try {
      Office.context.mailbox.item.notificationMessages.replaceAsync('pgp_reply_handoff_insight', {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        icon: 'icon16',
        message: message,
        persistent: false,
      });
    } catch (fallbackErr) {
      console.error('ReplyHandoffRuntime.classic: fallback notification also failed', fallbackErr);
    }
  }

  try {
    Office.context.mailbox.item.getComposeTypeAsync(function (result) {
      try {
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
        // itself does (conversationId, then item.inReplyTo on Mailbox
        // 1.14+): if there's no conversationId here, showing the
        // notification would just lead to a pane that finds nothing to
        // listen on. The pane does the real, complete check; this is only
        // deciding whether it's worth prompting at all.
        var conversationId = Office.context.mailbox.item.conversationId;
        if (!conversationId) {
          finish();
          return;
        }

        // The real gate: only prompt on a reply this add-in actually opened
        // for a handoff (see HANDOFF_PENDING_MARKER above).
        Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, function (bodyResult) {
          try {
            if (bodyResult.status !== Office.AsyncResultStatus.Succeeded ||
                bodyResult.value.indexOf(HANDOFF_PENDING_MARKER) === -1) {
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
                  actionType: 'showTaskPane',
                  commandId: 'msgComposeReplyHandoffButton',
                  contextData: {},
                },
              ],
            }, function (notifyResult) {
              if (notifyResult.status === Office.AsyncResultStatus.Failed) {
                var errMessage = notifyResult.error && notifyResult.error.message;
                console.error('ReplyHandoffRuntime.classic: InsightMessage notification failed', notifyResult.error);
                showFallbackInfo('Reply handoff ready, but the notification failed: ' + errMessage);
              }
              finish();
            });
          } catch (bodyErr) {
            console.error('ReplyHandoffRuntime.classic: handler failed inside body.getAsync callback', bodyErr);
            showFallbackInfo('ReplyHandoffRuntime error: ' + bodyErr.message);
            finish();
          }
        });
      } catch (innerErr) {
        console.error('ReplyHandoffRuntime.classic: handler failed inside getComposeTypeAsync callback', innerErr);
        showFallbackInfo('ReplyHandoffRuntime error: ' + innerErr.message);
        finish();
      }
    });
  } catch (e) {
    console.error('ReplyHandoffRuntime.classic: handler failed', e);
    finish();
  }
}

Office.actions.associate('onNewMessageComposeHandler', onNewMessageComposeHandler);
