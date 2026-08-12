'use strict';
/**
 * DecryptedPopup.js
 * Dialog-page counterpart to MessageRead.js's openDecryptedPopupDialog().
 * Receives decrypted content over a per-token BroadcastChannel and renders it.
 *
 * Carried-forward platform knowledge (see openDecryptedPopup's docblock in
 * MessageRead.js for the original context this was migrated away from):
 *  - The blob: URL WebView2 block does not apply here — this page never
 *    constructs a blob: URL; content arrives via srcdoc/textContent, not
 *    navigation. Don't reintroduce a blob: URL "optimization" without
 *    knowing it breaks on Outlook Desktop's WebView2 host.
 *  - The UTF-8 mojibake fix is moot here: BroadcastChannel carries JS
 *    strings (UTF-16 internally) end-to-end — there's no document.write()
 *    of a manually byte-assembled HTML string and no intermediate
 *    re-encoding step, which is what caused the historical apostrophe
 *    mojibake in the window.open path.
 *  - The focus-stacking caveat is the entire reason this file exists —
 *    Outlook raises the dialog itself, so there's no win.focus()-equivalent
 *    call needed or available here, and none should be added.
 */

const PGP_POPOUT_HANDSHAKE_TIMEOUT_MS = 10000;

function showError(message) {
  const errorEl = document.getElementById('popout-error');
  errorEl.textContent = message;
  errorEl.classList.remove('pgp-hidden');
}

function renderPayload({ text, isHtml, title }) {
  document.title = title || 'PGP Decrypted';
  if (isHtml) {
    document.getElementById('popout-html-frame').srcdoc = text;
    document.getElementById('popout-html-wrapper').classList.remove('pgp-hidden');
  } else {
    document.getElementById('popout-text').textContent = text;
    document.getElementById('popout-text').classList.remove('pgp-hidden');
  }
}

Office.onReady(() => {
  const token = new URLSearchParams(location.search).get('token');
  if (!token) {
    showError('This pop-out window was opened without a valid token.');
    return;
  }

  let channel;
  try {
    channel = new BroadcastChannel('pgp_popout_' + token);
  } catch {
    showError('This pop-out window could not connect to receive the decrypted content.');
    try {
      Office.context.ui.messageParent(JSON.stringify({ type: 'popout-error', reason: 'broadcast-channel-unavailable' }));
    } catch {
      // Not fatal — this dialog may not be running inside a real Office dialog context.
    }
    return;
  }

  let timeoutId = setTimeout(() => {
    channel.close();
    showError('The decrypted content did not arrive in time. Please try again.');
    try {
      Office.context.ui.messageParent(JSON.stringify({ type: 'popout-error', reason: 'timeout' }));
    } catch {
      // Not fatal — this dialog may not be running inside a real Office dialog context.
    }
  }, PGP_POPOUT_HANDSHAKE_TIMEOUT_MS);

  channel.onmessage = (event) => {
    if (event.data?.type === 'payload') {
      clearTimeout(timeoutId);
      channel.close();
      renderPayload(event.data);
    }
  };

  channel.postMessage({ type: 'dialog-listening' });
});
