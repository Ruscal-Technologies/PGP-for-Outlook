// Event-based-activation runtime for the OnNewMessageCompose LaunchEvent
// (manifest.xml's Runtimes/LaunchEvent elements, #22). Runs automatically
// the instant ANY new compose window opens (New Mail, Reply, Reply All, or
// Forward), on Outlook on the web, new Outlook on Windows, and new Mac UI --
// full browser runtime, modern JS/ES modules/BroadcastChannel all fine here.
// Classic Outlook on Windows instead runs Functions/ReplyHandoffRuntime.classic.js
// (the manifest's Override child), which has none of those guarantees.
//
// Office.onReady()/Office.initialize do NOT run in this context -- there is
// no shared bootstrap moment, every invocation is a cold start. Setup must
// live inside the handler itself.
//
// STEP 1 (current): no-op. Only confirms the event fires at all, on which
// platforms, for which compose types, before any real splice logic is
// written -- see the plan for #22. Check this via
// WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS=--auto-open-devtools-for-tabs (or
// the browser's own F12 devtools when testing Outlook on the web).

async function onNewMessageComposeHandler(event) {
  try {
    const composeType = await getComposeTypeAsync();
    console.log('ReplyHandoffRuntime: OnNewMessageCompose fired', {
      composeType,
      conversationId: Office.context.mailbox.item.conversationId,
    });
  } catch (e) {
    console.error('ReplyHandoffRuntime: handler failed', e);
  } finally {
    // Only call this once we're done with everything this handler needs to
    // do -- code after event.completed() is "not guaranteed to run", and
    // Outlook may tear the runtime down once it's called.
    event.completed();
  }
}

function getComposeTypeAsync() {
  return new Promise((resolve) => {
    Office.context.mailbox.item.getComposeTypeAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value.composeType);
      } else {
        resolve(null);
      }
    });
  });
}

Office.actions.associate('onNewMessageComposeHandler', onNewMessageComposeHandler);
