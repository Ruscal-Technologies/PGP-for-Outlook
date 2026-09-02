// Event-based-activation runtime for OnNewMessageCompose (manifest.xml's
// Runtimes/LaunchEvent elements, #22), CLASSIC Outlook on Windows variant.
//
// This file runs in Outlook's stripped-down "JavaScript-only runtime" and
// must not use anything the other (web/ReplyHandoffRuntime.js) runtime can:
//   - no async/await (times the add-in out on builds before ~April 2024 /
//     Build 17425.20000)
//   - no ternary operator (prevents the add-in from loading on those builds)
//   - no import/ES modules (must be a single flat file; nothing bundled)
// See web/ReplyHandoffRuntime.js's own comments for what this mirrors on
// other platforms, and CLAUDE.md / the #22 plan for the full rationale.
//
// Office.onReady()/Office.initialize do NOT run in this context -- there is
// no shared bootstrap moment, every invocation is a cold start.
//
// STEP 1 (current): no-op. Only confirms the event fires at all here, and
// (just as importantly) whether `document`/BroadcastChannel even exist in
// this runtime -- neither is guaranteed, unlike the browser-runtime file.
// Check via Windows Event Viewer (Windows Logs > Application, Event ID 63
// on failure) and the downloaded-handler-file location under
// %LOCALAPPDATA%\Microsoft\Office\16.0\Wef\...\Javascript\.

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
      console.log('ReplyHandoffRuntime.classic: OnNewMessageCompose fired', {
        composeType: composeType,
        conversationId: Office.context.mailbox.item.conversationId,
        hasDocument: typeof document !== 'undefined',
        hasBroadcastChannel: typeof BroadcastChannel !== 'undefined',
      });
      finish();
    });
  } catch (e) {
    console.error('ReplyHandoffRuntime.classic: handler failed', e);
    finish();
  }
}

Office.actions.associate('onNewMessageComposeHandler', onNewMessageComposeHandler);
