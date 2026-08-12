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
 *
 * Deliberately shorter than MessageRead.js's PGP_POPOUT_HANDSHAKE_TIMEOUT_MS
 * (10s): this side reports a specific error via messageParent and the parent
 * falls back to the legacy popup as soon as that arrives, so this timeout
 * should win the race against the parent's own generic backstop timer in the
 * normal case, giving the user the more specific message.
 */
const PGP_POPOUT_HANDSHAKE_TIMEOUT_MS = 8000;

function showError(message) {
  const errorEl = document.getElementById('popout-error');
  errorEl.textContent = message;
  errorEl.classList.remove('pgp-hidden');
}

export function renderPayload({ text, isHtml, title }) {
  document.title = title || 'PGP Decrypted';
  if (isHtml) {
    document.getElementById('popout-html-frame').srcdoc = text;
    document.getElementById('popout-html-wrapper').classList.remove('pgp-hidden');
  } else {
    document.getElementById('popout-text').textContent = text;
    document.getElementById('popout-text').classList.remove('pgp-hidden');
  }
}

/**
 * Best-effort notification to the parent pane. office.js may still be
 * loading (or fail to load) when this fires, so it deliberately does not
 * gate the render/handshake logic below — only this call waits on
 * Office.onReady, and only if the Office global exists at all.
 */
function notifyParent(payload) {
  if (typeof Office === 'undefined' || typeof Office.onReady !== 'function') return;
  Office.onReady()
    .then(() => {
      try {
        Office.context.ui.messageParent(JSON.stringify(payload));
      } catch {
        // Not fatal — this dialog may not be running inside a real Office dialog context.
      }
    })
    .catch(() => {
      // Not fatal — same as above.
    });
}

(function init() {
  const token = new URLSearchParams(location.search).get('token');
  if (!token) {
    console.error('Pop-out dialog opened without a token in the URL.');
    showError('This pop-out window was opened without a valid token.');
    return;
  }

  let channel;
  try {
    channel = new BroadcastChannel('pgp_popout_' + token);
  } catch (err) {
    console.error('Pop-out dialog: BroadcastChannel unavailable', err);
    showError('This pop-out window could not connect to receive the decrypted content.');
    notifyParent({ type: 'popout-error', reason: 'broadcast-channel-unavailable' });
    return;
  }

  const timeoutId = setTimeout(() => {
    channel.close();
    console.error('Pop-out dialog: handshake timed out waiting for the parent to deliver a payload.');
    showError('The decrypted content did not arrive in time. Please try again.');
    notifyParent({ type: 'popout-error', reason: 'timeout' });
  }, PGP_POPOUT_HANDSHAKE_TIMEOUT_MS);

  channel.onmessage = (event) => {
    if (event.data?.type !== 'payload') return;
    clearTimeout(timeoutId);
    channel.close();
    if (typeof event.data.text !== 'string') {
      console.error('Pop-out dialog: received a payload message with a missing/invalid text field', event.data);
      showError('The decrypted content could not be displayed. Please try again.');
      return;
    }
    renderPayload(event.data);
  };

  channel.postMessage({ type: 'dialog-listening' });
})();
