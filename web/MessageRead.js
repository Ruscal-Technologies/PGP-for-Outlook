'use strict';
/**
 * MessageRead.js
 * Task pane for reading PGP-encrypted or PGP-signed messages.
 *
 * Capabilities:
 *  - Detect PGP content in the message body (encrypted / signed)
 *  - Decrypt the body using the user's private key
 *  - Verify signatures against keys in the local keyring or via WKD/VKS
 *  - List .pgp attachments and allow individual decryption + download
 */

import {
  unlockPrivateKey,
  decryptMessage, decryptAttachment, verifyCleartextMessage,
  detectPgpContent, stripPgpExtension, applyDecryptedExtensionPrefix,
  encryptMessage, readPublicKey,
} from './js/pgp/pgp-core.js';
import { hasKeyPair, getPrivateKey, getPublicKey, getSignDefault, getKeyMetadata } from './js/pgp/key-storage.js';
import { getContactKeyObject } from './js/pgp/keyring.js';
import { discoverKey, KeyStatus } from './js/pgp/key-discovery.js';
import { loadOrgConfig, getDecryptedExtensionPrefix } from './js/pgp/org-config.js';
import {
  cacheSessionKey, getSessionKey, clearSessionKey,
  getSessionEmail, getSessionShortId, onSessionCleared,
} from './js/pgp/session-cache.js';
import { formatDecryptedContentAsHtml, formatDecryptedContentAsPlainTextHtml } from './js/pgp/quoted-content.js';
import { getReplyHandoffChannelName, HANDOFF_PENDING_MARKER } from './js/pgp/reply-handoff-channel.js';

// ── Module state ──────────────────────────────────────────────────────────────

/** Decrypted payload, stored so reply handlers can quote it. */
let _decryptedText = null;
let _decryptedIsHtml = false;

/** True when running inside Outlook on iOS or Android. */
let _isMobile = false;

// True while a large-message native-reply handoff (openNativeReplyWithHandoff)
// is in flight from THIS reading pane, i.e. between displayReplyForm/
// displayReplyAllForm succeeding and the handoff settling (ack, timeout, or
// channel failure). The Reply/Reply All buttons are disabled for the
// duration so a second click from the same pane can't start a second
// concurrent handoff that would share (and cross-wire) the same
// conversation-scoped BroadcastChannel -- see issue #17. This only guards
// against two attempts from the same pane; it can't prevent two separate
// Outlook windows on the same conversation from racing each other, since
// each has its own independent module state.
let _nativeReplyHandoffInFlight = false;

/**
 * True when the host meets Mailbox 1.7 (Outlook 2021+).
 * Required for item.from (sender email/name in read mode).
 * @type {boolean}
 */
let _has17 = false;

/**
 * True when the host meets Mailbox 1.8 (Outlook 2021+ / Microsoft 365 / OWA).
 * Required for getAttachmentContentAsync (attachment decryption).
 * @type {boolean}
 */
let _has18 = false;

/**
 * True when the host meets Mailbox 1.4 (Outlook 2016+).
 * Required for Office.context.ui.displayDialogAsync / messageParent, used by
 * the dialog-based "Pop Out" implementation (see openDecryptedPopupDialog).
 * @type {boolean}
 */
let _has14 = false;

/** Tracks the single open pop-out dialog, if any (the Dialog API only supports one per host window). */
let _popoutDialog = null;

/**
 * Timer + channel for the pop-out dialog's BroadcastChannel handshake.
 * Tracked at module scope (rather than only as locals inside
 * openDecryptedPopupDialog) so onPopoutDialogClosed()/onPopoutDialogMessage()
 * can settle a handshake that's still in flight — or, per #12, still
 * servicing a reload — regardless of which event arrives.
 *
 * _popoutHandshakeChannel is deliberately kept open past the first payload
 * delivery (see openDecryptedPopupDialog) rather than closed immediately: if
 * the dialog reloads, it re-broadcasts "dialog-listening" and needs a live
 * channel to receive the payload again. It's closed only once the dialog
 * itself closes, or the handshake fails outright.
 */
let _popoutHandshakeTimer = null;
let _popoutHandshakeChannel = null;

/**
 * Args to replay through the legacy openDecryptedPopup() fallback if the
 * dialog-based path fails after successfully opening. Set for the duration
 * of one openDecryptedPopupDialog() call; cleared once the fallback fires
 * (or never used, if the dialog succeeds and the user closes it normally).
 */
let _popoutFallbackArgs = null;

/**
 * Guards triggerPopoutFallback() so exactly one failure signal — the
 * parent's own handshake timeout, a relayed error from the dialog, or an
 * unexpected DialogEventReceived — opens the legacy popup, even though more
 * than one of those can fire for the same underlying failure.
 */
let _popoutFallbackTriggered = false;

/** Handshake budget for the pop-out dialog's BroadcastChannel readiness signal. */
const PGP_POPOUT_HANDSHAKE_TIMEOUT_MS = 10000;


// ── Helpers ───────────────────────────────────────────────────────────────────

function el(id) { return document.getElementById(id); }

/**
 * Enables/disables the Reply and Reply All buttons -- used to block a second
 * concurrent large-message handoff from this same pane while one is already
 * in flight (see _nativeReplyHandoffInFlight). No-op for any button not
 * present in the current DOM (e.g. mobile layout).
 */
function setReplyButtonsDisabled(disabled) {
  const replyBtn = el('btn-reply-encrypted');
  const replyAllBtn = el('btn-reply-all-encrypted');
  if (replyBtn) replyBtn.disabled = disabled;
  if (replyAllBtn) replyAllBtn.disabled = disabled;
}

function escHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function showSection(id) { el(id).classList.remove('pgp-hidden'); }
function hideSection(id) { el(id).classList.add('pgp-hidden'); }

function showStatus(message, type = 'info') {
  const bar = el('status-bar');
  bar.className = `pgp-alert pgp-alert--${type}`;
  bar.textContent = message;
  bar.classList.remove('pgp-hidden');
}

function showStatusReplyOpened() {
  const bar = el('status-bar');
  bar.className = 'pgp-alert pgp-alert--info';
  bar.textContent = 'Reply opened — click ';
  const strong = document.createElement('strong');
  strong.textContent = 'Encrypt';
  bar.appendChild(strong);
  bar.appendChild(document.createTextNode(' in the ribbon to encrypt before sending.'));
  bar.classList.remove('pgp-hidden');
}

// ── Session status ────────────────────────────────────────────────────────────

function updateSessionStatus() {
  const bar   = el('session-status');
  const label = el('session-status-text');
  const email   = getSessionEmail();
  const shortId = getSessionShortId();
  if (email) {
    label.textContent = `Key unlocked: ${email}${shortId ? ' · …' + shortId : ''}`;
    bar.classList.remove('pgp-hidden');
  } else {
    bar.classList.add('pgp-hidden');
  }
}

// ── Passphrase modal ──────────────────────────────────────────────────────────

function promptPassphrase(message = 'Enter your passphrase to decrypt.') {
  return new Promise((resolve, reject) => {
    const modal = el('passphrase-modal');
    const input = el('passphrase-input');
    const errEl = el('passphrase-error');
    const msgEl = el('passphrase-modal-msg');

    msgEl.textContent = message;
    input.value = '';
    errEl.classList.add('pgp-hidden');
    modal.style.display = 'flex';
    modal.classList.remove('pgp-hidden');
    input.focus();

    function cleanup() {
      modal.style.display = '';
      modal.classList.add('pgp-hidden');
      el('btn-passphrase-ok').removeEventListener('click', onOk);
      el('btn-passphrase-cancel').removeEventListener('click', onCancel);
      input.removeEventListener('keydown', onKeydown);
    }

    function onOk() {
      const val = input.value;
      if (!val) {
        errEl.textContent = 'Passphrase is required.';
        errEl.classList.remove('pgp-hidden');
        return;
      }
      cleanup();
      resolve(val);
    }

    function onCancel() { cleanup(); reject(new Error('Cancelled.')); }
    function onKeydown(e) {
      if (e.key === 'Enter')  onOk();
      if (e.key === 'Escape') onCancel();
    }

    el('btn-passphrase-ok').addEventListener('click', onOk);
    el('btn-passphrase-cancel').addEventListener('click', onCancel);
    input.addEventListener('keydown', onKeydown);
  });
}

// ── Message body detection ────────────────────────────────────────────────────

/**
 * Normalise an armoured PGP text block that may have been mangled by Outlook
 * on Android / iOS or by HTML-to-text conversion.
 *
 * Known mutations introduced by Outlook mobile:
 *  - Non-breaking hyphens (U+2011), en/em dashes (U+2013/2014), figure dash
 *    (U+2012), horizontal bar (U+2015), minus sign (U+2212) replacing the five
 *    ASCII hyphens that delimit -----BEGIN/END PGP …----- headers.
 *  - Soft hyphens (U+00AD) inserted invisibly inside lines.
 *  - Non-breaking spaces (U+00A0) replacing ordinary spaces in header lines.
 *  - Zero-width characters (U+200B–U+200F, U+FEFF, U+2028, U+2029) injected
 *    between lines or at line boundaries.
 *  - Windows line endings (CRLF) or bare CR mixed with LF.
 *  - Trailing whitespace left on individual lines.
 */
function sanitizeArmoredText(text) {
  if (!text) return text;

  // Replace every visually-similar dash/hyphen with ASCII hyphen-minus U+002D.
  text = text
    .replace(/\u00AD/g, '')    // soft hyphen — invisible, just remove it
    .replace(/[\u2011\u2012\u2013\u2014\u2015\u2212]/g, '-'); // dashes → -

  // Non-breaking space → regular space (appears in armor header lines).
  text = text.replace(/\u00A0/g, ' ');

  // Strip zero-width and Unicode line/paragraph separators that Outlook mobile
  // injects and that OpenPGP.js rejects as malformed armor.
  text = text.replace(/[\u200B-\u200F\uFEFF\u2028\u2029]/g, '');

  // Normalise line endings to LF, then trim whitespace from both ends of each
  // line.  PGP armor never has leading or trailing whitespace on any line, so
  // trimming both ends is safe and removes indentation that Outlook Desktop
  // classic sometimes adds to the HTML representation of the body.
  text = text
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .split('\n')
    .map(line => line.trim())
    .join('\n');

  // PGP armor is defined as ASCII-only (RFC 4880 §6).  Strip any remaining
  // non-ASCII characters that Outlook Desktop classic's Word-based HTML renderer
  // may have injected (e.g. typographic substitutions, Windows code-page
  // artifacts) and that the targeted replacements above did not already convert.
  // This is a catch-all safety net; the atob() call inside OpenPGP.js throws on
  // any character that is not in [A-Za-z0-9+/=] and not ASCII whitespace.
  text = text.replace(/[^\x00-\x7F]/g, '');

  return text;
}

/**
 * Extract text from an HTML body string for PGP armor detection.
 *
 * Handles two storage formats:
 *
 *   <pre>-wrapped  — set by this add-in's setBodyAsync().  Uses node.textContent
 *                    which is CSS-independent and preserves all original newlines
 *                    regardless of white-space:pre-wrap on mobile.
 *
 *   <div>-per-line — pasted plain text in any Outlook client (each line is a
 *                    separate <div>).  The recursive walk appends \n after each
 *                    block-level element to reconstruct line structure.
 *
 * Does NOT use innerText.  innerText requires a rendered (attached) element to
 * compute layout; on a detached div it silently collapses all whitespace in
 * non-pre contexts, corrupting the base64 payload.  The recursive walk below
 * handles both <pre> and <div>-per-line content correctly without CSS layout.
 */
function extractArmorFromHtml(html) {
  const div = document.createElement('div');
  div.innerHTML = html;

  const BLOCK = new Set([
    'div', 'p', 'li', 'blockquote', 'tr',
    'h1', 'h2', 'h3', 'h4', 'h5', 'h6',
    'article', 'section', 'header', 'footer', 'html', 'body',
  ]);

  // Elements whose text content must never be extracted.  When Outlook Desktop
  // returns a full HTML document from getBodyAsync(Html), the browser's fragment
  // parser places <head> children (<style>, <title>, <meta>) as siblings of the
  // body content inside our temporary <div>.  Without this skip list, walk()
  // would include raw CSS and document title text in the extracted string.
  const SKIP = new Set(['style', 'script', 'head', 'title', 'noscript']);

  function walk(node) {
    if (node.nodeType === Node.TEXT_NODE) return node.textContent;
    if (node.nodeType !== Node.ELEMENT_NODE) return '';
    const tag = node.tagName.toLowerCase();
    if (SKIP.has(tag)) return '';
    if (tag === 'br') return '\n';
    if (tag === 'pre') {
      // textContent is CSS-independent: gives the raw text with original \n
      // characters regardless of white-space:pre-wrap on mobile WebViews.
      return '\n' + node.textContent + '\n';
    }
    let text = '';
    for (const child of node.childNodes) text += walk(child);
    if (BLOCK.has(tag)) text += '\n';
    return text;
  }

  return walk(div);
}

/**
 * Strip informational armor headers from PGP MESSAGE blocks.
 *
 * RFC 4880 §6.2 allows any number of key-value header lines (Version:,
 * Comment:, Charset:, MessageID:, etc.) between the -----BEGIN PGP MESSAGE-----
 * line and the blank line that separates headers from the base64 payload.
 * For encrypted messages these headers are purely informational and play no
 * role in decryption.  However, email clients (notably Outlook Desktop) can
 * mangle the header values — injecting non-ASCII characters, adding extra
 * whitespace, or altering encoding — in ways that confuse OpenPGP.js's armor
 * reader even after our character-level sanitization.
 *
 * This function removes all such headers, leaving only:
 *   -----BEGIN PGP MESSAGE-----
 *   [blank line]
 *   [base64 payload + =checksum]
 *   -----END PGP MESSAGE-----
 *
 * Only -----BEGIN PGP MESSAGE----- blocks are affected.  SIGNED MESSAGE blocks
 * are intentionally left alone because their Hash: header tells the verifier
 * which digest algorithm was used and is required for signature verification.
 *
 * Input must already have LF-only line endings (post-sanitizeArmoredText).
 */
function stripArmorHeaders(text) {
  // Process line-by-line rather than with a regex so we handle all edge
  // cases: missing blank separator, blank lines interspersed between
  // header lines (produced when sanitizeArmoredText trims a whitespace-only
  // header value to empty), or any other structural variation a third-party
  // client may produce.
  //
  // RFC 4880 §6.2 header lines always contain ':' (key-value format).
  // Base64 characters are [A-Za-z0-9+/=] — none of which is ':' — so
  // testing for ':' reliably distinguishes header lines from body lines.
  const lines = text.split('\n');
  const out = [];
  let i = 0;
  while (i < lines.length) {
    if (lines[i] === '-----BEGIN PGP MESSAGE-----') {
      out.push(lines[i]);
      i++;
      // Skip header lines (contain ':') and blank lines that follow BEGIN.
      while (i < lines.length && (lines[i] === '' || lines[i].includes(':'))) {
        i++;
      }
      out.push(''); // blank separator before base64 payload
    } else {
      out.push(lines[i]);
      i++;
    }
  }
  return out.join('\n');
}

/**
 * Extract the first complete PGP armor block from a body string.
 *
 * In reply threads the full body contains:
 *  - The reply's armor (pasted by the user, at the top)
 *  - Outlook-added separators like "-----Original Message-----"
 *  - The quoted original message (may contain a second PGP armor block)
 *
 * "-----Original Message-----" has the same -----…----- format as a PGP
 * armor header.  If OpenPGP.js sees it while scanning for the END marker it
 * throws "Unknown ASCII armor type" (or tries to parse a second block).
 *
 * This function isolates just the first BEGIN…END block so that
 * openpgp.readMessage() receives a clean, unambiguous input.
 *
 * PGP SIGNED MESSAGE is handled as a special case: its structure is
 *   -----BEGIN PGP SIGNED MESSAGE-----
 *   …plaintext…
 *   -----BEGIN PGP SIGNATURE-----
 *   …
 *   -----END PGP SIGNATURE-----
 * (there is no -----END PGP SIGNED MESSAGE-----)
 *
 * Returns the original text unchanged when no complete armor block is found
 * so the caller can still attempt decryption with whatever it has.
 */
function extractFirstArmorBlock(text) {
  const beginMatch = text.match(/-----BEGIN PGP ([A-Z ]+?)-----/);
  if (!beginMatch) return text;

  const type   = beginMatch[1]; // e.g. "MESSAGE", "SIGNED MESSAGE"
  const endStr = type === 'SIGNED MESSAGE'
    ? '-----END PGP SIGNATURE-----'
    : `-----END PGP ${type}-----`;

  const startIdx = text.indexOf(beginMatch[0]);
  const endIdx   = text.indexOf(endStr, startIdx);
  if (endIdx === -1) return text; // incomplete armor — let OpenPGP.js report the error

  return text.slice(startIdx, endIdx + endStr.length);
}

async function detectAndRenderBody() {
  let body = null;
  let pgpType = null;

  // Always prefer the HTML body path on all platforms.
  //
  // extractArmorFromHtml() handles both formats:
  //   - <pre>-wrapped armor (set by this add-in's setBodyAsync)
  //   - <div>-per-line armor (pasted plain text in any Outlook client)
  //
  // The HTML path correctly preserves the blank-line separator between armor
  // headers and base64 payload because it inserts \n at every </div> boundary
  // and replaces <pre> elements with their textContent (CSS-independent).
  //
  // CoercionType.Text is unreliable as a primary source:
  //   - On Outlook Web it can collapse the required blank line in pasted armor,
  //     causing openpgp.readMessage() to throw "Misformed armored text".
  //   - On Outlook Android it can return raw HTML with tags still present, or
  //     inject visual line-wrap newlines into base64 lines.
  try {
    const htmlBody = await getBodyAsync(Office.CoercionType.Html);
    const extracted = sanitizeArmoredText(extractArmorFromHtml(htmlBody));
    const t = detectPgpContent(extracted);
    if (t) { body = extracted; pgpType = t; }
  } catch { /* HTML unavailable — fall through to text */ }

  // Text fallback: catches plain-text-only messages and edge cases where the
  // HTML body is unavailable or extractArmorFromHtml misses the armor.
  if (!pgpType) {
    try {
      const textBody = sanitizeArmoredText(await getBodyAsync(Office.CoercionType.Text));
      const t = detectPgpContent(textBody);
      if (t) { body = textBody; pgpType = t; }
    } catch { /* body completely unavailable */ }
  }

  // Strip reply-thread noise (quoted originals, "-----Original Message-----"
  // separators, etc.) so OpenPGP.js sees a single clean armor block.
  if (body && pgpType) {
    body = extractFirstArmorBlock(body);
    body = stripArmorHeaders(body);
  }

  el('detection-loading').classList.add('pgp-hidden');

  const result = el('detection-result');
  result.classList.remove('pgp-hidden');

  if (!pgpType) {
    result.innerHTML = `<div class="pgp-alert pgp-alert--info">
      This message does not appear to contain PGP content.
    </div>`;
    renderPgpAttachments(); // still look for .pgp attachments
    return;
  }

  if (pgpType === 'encrypted') {
    result.innerHTML = `<div class="pgp-alert pgp-alert--info">
      <strong>Encrypted message</strong> — PGP-encrypted content detected.
    </div>`;
    showSection('section-decrypt');
    el('btn-decrypt').addEventListener('click', () => handleDecryptBody(body), { once: true });
    handleDecryptBody(body); // auto-start: warm session is instant; cold session shows passphrase modal
  }

  if (pgpType === 'signed') {
    result.innerHTML = `<div class="pgp-alert pgp-alert--info">
      <strong>Signed message</strong> — PGP-signed content detected.
    </div>`;
    showSection('section-signed-only');
    await handleVerifySignedBody(body);
  }

  if (pgpType === 'public-key') {
    result.innerHTML = `<div class="pgp-alert pgp-alert--warning">
      This message contains a <strong>PGP public key</strong>.
      You can copy it and import it via <em>Manage Keys</em>.
    </div>`;
  }

  if (pgpType === 'private-key') {
    result.innerHTML = `<div class="pgp-alert pgp-alert--error">
      ⚠ This message contains what appears to be a <strong>private key</strong>.
      Do not share or import private keys.
    </div>`;
  }

  renderPgpAttachments();
}

// ── Decrypt body ──────────────────────────────────────────────────────────────

async function handleDecryptBody(encryptedBody) {
  const btn = el('btn-decrypt');
  const spinner = el('decrypt-spinner');
  btn.disabled = true;
  spinner.classList.remove('pgp-hidden');

  try {
    // Check the session cache before prompting — avoids repeated passphrase
    // entry when the user decrypts several messages or attachments in one session.
    let privateKey = getSessionKey();

    if (!privateKey) {
      const passphrase = await promptPassphrase('Enter your passphrase to decrypt this message.');
      privateKey = await unlockPrivateKey(getPrivateKey(), passphrase);

      // Cache for 15 minutes of inactivity.
      const userEmail = Office.context.mailbox.userProfile?.emailAddress || '';
      const meta = getKeyMetadata();
      cacheSessionKey(privateKey, userEmail, meta?.keyId?.slice(-8) || '');
      updateSessionStatus();
    }

    // Attempt to get the sender's public key for signature verification
    const senderEmail = Office.context.mailbox.item.from?.emailAddress;
    const verificationKeys = await resolveVerificationKeys(senderEmail);

    const { data, signatureResult } = await decryptMessage(
      encryptedBody, privateKey, verificationKeys
    );

    renderDecryptedBody(data, signatureResult, senderEmail);
    hideSection('section-decrypt');

  } catch (e) {
    if (e.message === 'Cancelled.') {
      /* user cancelled — silently re-enable */
    } else if (e.message?.includes('Error decrypting') || e.message?.includes('Decryption error')) {
      showStatus('Decryption failed — wrong passphrase or key?', 'error');
    } else {
      showStatus(`Decryption failed: ${e.message}`, 'error');
    }
  } finally {
    btn.disabled = false;
    spinner.classList.add('pgp-hidden');
  }
}

function renderDecryptedBody(text, signatureResult, senderEmail) {
  _decryptedText = text;
  _decryptedIsHtml = /^\s*<[a-zA-Z!]/.test(text);

  showSection('section-decrypted');

  // Render signature badge
  const sigBadge = el('signature-badge');
  const sigDetails = el('signature-details');

  if (signatureResult.valid === null) {
    sigBadge.innerHTML = `<span class="pgp-badge pgp-badge--neutral">No signature</span>`;
  } else if (signatureResult.valid) {
    sigBadge.innerHTML = `<span class="pgp-badge pgp-badge--success">✓ Valid signature</span>`;
    if (senderEmail) {
      sigDetails.textContent = `Signed by ${senderEmail} · Key ID: ${signatureResult.signedByKeyId || 'unknown'}`;
      sigDetails.classList.remove('pgp-hidden');
    } else if (!_has17) {
      sigDetails.textContent = 'Sender information unavailable on this Outlook version — upgrade to Outlook 2021 to see who signed this message.';
      sigDetails.classList.remove('pgp-hidden');
    }
  } else {
    sigBadge.innerHTML = `<span class="pgp-badge pgp-badge--error">✗ Invalid signature</span>`;
    sigDetails.textContent = `Signature could not be verified. The message may have been tampered with.`;
    sigDetails.classList.remove('pgp-hidden');
  }

  // Detect whether the decrypted payload is HTML.  Outlook's getBodyAsync(Html)
  // can return content starting with <div>, <body>, or <html> depending on the
  // client, so we check for any leading HTML tag rather than just <html>.
  const isHtml = /^\s*<[a-zA-Z!]/.test(text);

  if (isHtml) {
    // Render in a sandboxed iframe.  'allow-same-origin' lets the iframe read
    // its srcdoc content but scripts are NOT allowed (no 'allow-scripts').
    // This prevents any JavaScript inside the decrypted HTML from running.
    const frame = el('decrypted-html-frame');
    frame.srcdoc = text;
    el('decrypted-html-wrapper').classList.remove('pgp-hidden');

    // Resize iframe to fit content once it loads
    frame.addEventListener('load', () => {
      try {
        frame.style.height = frame.contentDocument.body.scrollHeight + 32 + 'px';
      } catch { /* cross-origin guard — shouldn't fire with srcdoc + allow-same-origin */ }
    }, { once: true });
  } else {
    el('decrypted-body').textContent = text;
    el('decrypted-body').classList.remove('pgp-hidden');
  }

  el('btn-copy-decrypted').addEventListener('click', async () => {
    try {
      await navigator.clipboard.writeText(text);
      showStatus('Decrypted content copied to clipboard.', 'success');
    } catch {
      window.prompt('Copy the decrypted content:', text);
    }
  });

  el('btn-popout-decrypted').addEventListener('click', () => {
    const subject = Office.context.mailbox.item?.subject || '';
    if (_has14 && !_isMobile) {
      openDecryptedPopupDialog(text, isHtml, subject);
    } else {
      openDecryptedPopup(text, isHtml, subject);
    }
  });
}

// ── Pop-out window (dialog-based, and legacy window.open fallback) ─────────────

/**
 * Open decrypted content in a larger, resizable browser window.
 *
 * @param {string}  text     - Decrypted payload
 * @param {boolean} isHtml   - True when the payload is HTML
 * @param {string}  subject  - Original message subject (used as window title)
 *
 * For HTML payloads a CSP meta tag is injected to block script execution,
 * mirroring the sandbox= restriction used by the in-pane iframe.
 * The window title is set to "PGP Decrypted : <subject>" and browser chrome
 * (address bar, toolbar, menu bar) is suppressed via window.open features.
 * Note: modern browsers may still show the address bar for security reasons,
 * but Outlook's embedded WebView typically honours these flags.
 *
 * We write the HTML directly via document.write() rather than a Blob URL.
 * Blob URL navigation is blocked in Outlook Desktop's WebView2 host: Windows
 * intercepts the blob: protocol at the OS level and shows "Get an app to open
 * this 'blob' link".  Writing to a blank window bypasses that restriction and
 * also avoids the UTF-8 encoding ambiguity that caused apostrophes and other
 * non-ASCII characters to render as mojibake (â€™ etc.) in OWA.
 *
 * We also call win.focus() after writing — Outlook Classic's WebView2 host
 * doesn't always raise the new window above the Outlook window on its own.
 * See the inline comment near that call for the focus-stacking caveats.
 */
function openDecryptedPopup(text, isHtml, subject = '') {
  const pageTitle = subject ? `PGP Decrypted : ${subject}` : 'PGP Decrypted';
  // Escape the title for safe insertion into HTML.
  const safeTitle = pageTitle.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

  let html;
  if (isHtml) {
    // Inject charset declaration, CSP (blocks scripts), and our window title
    // into <head>.  charset must come first so the parser uses UTF-8 from the
    // very beginning of the document.  Any existing <title> is removed so the
    // window caption always shows "PGP Decrypted : …" rather than the email's
    // own title.
    const inject = `<meta charset="UTF-8">` +
                   `<meta http-equiv="Content-Security-Policy" ` +
                   `content="script-src 'none'; object-src 'none';">` +
                   `<title>${safeTitle}</title>`;
    if (/<head[\s>]/i.test(text)) {
      // Remove any pre-existing <title> then prepend our tags after <head …>
      const noTitle = text.replace(/<title[^>]*>[\s\S]*?<\/title>/gi, '');
      html = noTitle.replace(/(<head[\s>][^>]*>)/i, `$1${inject}`);
    } else {
      html = `<!DOCTYPE html><html><head>${inject}</head><body>${text}</body></html>`;
    }
  } else {
    const safe = text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    html = `<!DOCTYPE html><html><head><meta charset="UTF-8"><title>${safeTitle}</title>` +
      `<style>body{font-family:Calibri,Arial,sans-serif;font-size:14px;` +
      `line-height:1.6;padding:24px;white-space:pre-wrap;word-break:break-word;}</style>` +
      `</head><body>${safe}</body></html>`;
  }

  const win = window.open(
    '', '_blank',
    'resizable=yes,width=840,height=680,scrollbars=yes,' +
    'location=no,toolbar=no,menubar=no,status=no'
  );
  if (win) {
    win.document.open();
    win.document.write(html);
    win.document.close();
    // Outlook Classic's WebView2 host doesn't always raise the new window
    // above the Outlook window on its own; focus() is the one thing content
    // scripts can do to ask for it. It's a request, not a guarantee — once
    // the user clicks back into Outlook, normal OS window stacking applies
    // and there's no API to pin it "always on top" beyond this point.
    win.focus();
  } else {
    showStatus('Pop-out window was blocked. Please allow pop-ups for this site and try again.', 'error');
  }
}

/**
 * Generates an unguessable per-session token to gate/correlate a
 * BroadcastChannel handoff — used both for the pop-out dialog's payload
 * channel (kept open for the life of the dialog, to survive a reload) and
 * the reply native-quote handoff to MessageCompose.js (see
 * handleReplyEncrypted). crypto.randomUUID() when available, otherwise
 * crypto.getRandomValues() (available anywhere BroadcastChannel is; unlike
 * randomUUID, it doesn't require a secure context) — never Date.now()/
 * Math.random(), which are guessable.
 */
function generateChannelToken() {
  if (crypto.randomUUID) return crypto.randomUUID();
  const bytes = new Uint8Array(16);
  crypto.getRandomValues(bytes);
  return Array.from(bytes, (b) => b.toString(16).padStart(2, '0')).join('');
}

/**
 * Open decrypted content in an Office-managed dialog box.
 *
 * openDecryptedPopup()'s window.open() + focus() approach regressed: Windows'
 * OS-level foreground-lock / focus-stealing prevention increasingly blocks a
 * background WebView2 process from raising a window it didn't create as the
 * foreground process — a content-script fix can't reliably override that OS
 * policy decision, and it's expected to keep regressing as Windows/Edge/WebView2
 * tighten enforcement further.
 *
 * displayDialogAsync sidesteps this: Outlook itself — already the foreground
 * process, responding directly to the user's click — creates and raises the
 * dialog, so it isn't subject to the same restriction.
 *
 * The decrypted payload is handed to the dialog page over a same-origin
 * BroadcastChannel (in-memory only, never touches disk) rather than Office.js's
 * own Dialog.messageChild, which requires Mailbox 1.9 — a full tier above this
 * add-in's 1.5 floor, and would silently be unavailable on Outlook 2019-era
 * desktop clients. See web/DecryptedPopup.js for the receiving side of the
 * handshake (it signals "dialog-listening" before we post the payload, so the
 * message can't be sent before anything is there to receive it).
 *
 * Falls back to openDecryptedPopup() (window.open) if BroadcastChannel is
 * unavailable, if displayDialogAsync itself fails to open a dialog, or if a
 * dialog that DID open never completes its handshake / reports its own
 * failure / closes unexpectedly — see triggerPopoutFallback(), the single
 * chokepoint all of those failure signals funnel through.
 *
 * @param {string}  text     - Decrypted payload
 * @param {boolean} isHtml   - True when the payload is HTML
 * @param {string}  subject  - Original message subject (used as dialog title)
 */
export function openDecryptedPopupDialog(text, isHtml, subject = '') {
  const pageTitle = subject ? `PGP Decrypted : ${subject}` : 'PGP Decrypted';

  if (typeof BroadcastChannel !== 'function') {
    openDecryptedPopup(text, isHtml, subject);
    return;
  }

  // Defensively close any dialog we're still tracking before opening a new
  // one — if our reference ever went stale (missed a close event), this is
  // what keeps a fresh attempt from immediately hitting DialogAlreadyOpened.
  closePopoutDialogQuietly();

  const token = generateChannelToken();
  let channel;
  try {
    channel = new BroadcastChannel('pgp_popout_' + token);
  } catch (err) {
    console.error('Pop-out dialog: BroadcastChannel construction failed', err);
    openDecryptedPopup(text, isHtml, subject);
    return;
  }
  _popoutHandshakeChannel = channel;
  _popoutFallbackTriggered = false;
  _popoutFallbackArgs = { text, isHtml, subject };

  _popoutHandshakeTimer = setTimeout(() => {
    console.error('Pop-out dialog: handshake timed out waiting for the dialog to signal readiness.');
    triggerPopoutFallback('Pop-out window failed to load. Please try again.');
  }, PGP_POPOUT_HANDSHAKE_TIMEOUT_MS);

  // The dialog signals readiness first; only then do we hand it the payload,
  // so there's no window where the message could be sent before the dialog's
  // own listener exists. The channel is deliberately left open afterward
  // (see the module-state docblock above) rather than closed here.
  channel.onmessage = (event) => {
    if (event.data?.type !== 'dialog-listening') return;
    clearPopoutHandshakeTimer();
    channel.postMessage({ type: 'payload', text, isHtml, title: pageTitle });
  };

  const dialogUrl = new URL(`DecryptedPopup.html?token=${encodeURIComponent(token)}`, window.location.href).href;

  Office.context.ui.displayDialogAsync(dialogUrl, { height: 70, width: 60 }, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.error('Pop-out dialog: displayDialogAsync failed to open', asyncResult.error);
      clearPopoutHandshakeTimer();
      closePopoutHandshakeChannel();
      handleDialogOpenFailure(asyncResult.error, text, isHtml, subject);
      return;
    }

    // Captured locally (in addition to the module-level _popoutDialog) so
    // the two handlers below can tell a stale event from a dialog that's
    // since been superseded — e.g. one that's still in flight from before
    // closePopoutDialogQuietly()'s dialog.close() took effect — apart from a
    // real event for the dialog that's actually current. See the
    // "_popoutDialog !== dialog" guard in both handlers.
    const dialog = asyncResult.value;
    _popoutDialog = dialog;
    dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => onPopoutDialogMessage(dialog, arg));
    dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => onPopoutDialogClosed(dialog, arg));
  });
}

/** Cancels the pending pop-out handshake timeout, if one is still running. */
function clearPopoutHandshakeTimer() {
  if (_popoutHandshakeTimer) {
    clearTimeout(_popoutHandshakeTimer);
    _popoutHandshakeTimer = null;
  }
}

/** Closes the pending pop-out handshake's BroadcastChannel, if still open. */
function closePopoutHandshakeChannel() {
  if (_popoutHandshakeChannel) {
    _popoutHandshakeChannel.close();
    _popoutHandshakeChannel = null;
  }
}

/**
 * Ends the in-flight pop-out handshake exactly once and falls back to the
 * legacy openDecryptedPopup() (window.open) path, so the user is never left
 * staring at a stranded/blank dialog. Safe to call more than once for the
 * same failure — e.g. a relayed error from the dialog and this window's own
 * backstop timeout can both fire for the same underlying cause, but only
 * the first call does anything; see _popoutFallbackTriggered.
 */
function triggerPopoutFallback(statusMessage) {
  if (_popoutFallbackTriggered) return;
  _popoutFallbackTriggered = true;

  clearPopoutHandshakeTimer();
  closePopoutHandshakeChannel();

  if (_popoutDialog) {
    try { _popoutDialog.close(); } catch { /* dialog may already be gone */ }
    _popoutDialog = null;
  }

  if (statusMessage) showStatus(statusMessage, 'error');

  const args = _popoutFallbackArgs;
  _popoutFallbackArgs = null;
  if (args) openDecryptedPopup(args.text, args.isHtml, args.subject);
}

/**
 * Closes any dialog we're currently tracking without falling back to the
 * legacy popup — used when the user explicitly wants decrypted content to
 * stop being shown (Lock), not replaced with another window.
 */
export function closePopoutDialogQuietly() {
  clearPopoutHandshakeTimer();
  closePopoutHandshakeChannel();
  _popoutFallbackTriggered = true;
  _popoutFallbackArgs = null;
  if (_popoutDialog) {
    try { _popoutDialog.close(); } catch { /* already closed */ }
    _popoutDialog = null;
  }
}

/**
 * Handle a displayDialogAsync() open failure.
 *
 * DialogAlreadyOpened (12007) means the user already has a pop-out open —
 * surfaced as-is rather than falling back, since opening a second window via
 * window.open() underneath an already-open dialog would be confusing rather
 * than helpful. Any other failure falls back to the legacy window.open() path
 * so the user isn't left with a dead button.
 *
 * Exact numeric codes are from the documented Dialog API error surface, not
 * re-verified against the live docs this session — confirm against "Handle
 * errors and events in the Office dialog box" if this ever needs updating.
 */
export function handleDialogOpenFailure(error, text, isHtml, subject) {
  if (error?.code === 12007) {
    // The *existing* open dialog (the one causing this 12007) is untouched —
    // only this just-failed attempt's own fallback state needs clearing, so
    // its plaintext doesn't sit parked in module scope until the next click.
    _popoutFallbackTriggered = true;
    _popoutFallbackArgs = null;
    showStatus('A pop-out window is already open.', 'error');
    return;
  }
  showStatus('Could not open the pop-out window as a dialog; opening a regular window instead.', 'warning');
  openDecryptedPopup(text, isHtml, subject);
}

/**
 * Relays a message the dialog reported about itself — currently only its own
 * handshake-timeout / BroadcastChannel-unavailable errors. Falls back to the
 * legacy popup rather than just showing a toast, consistent with this file's
 * "the user should never be left with just a dead dialog" goal.
 */
function onPopoutDialogMessage(dialog, arg) {
  if (_popoutDialog !== dialog) return; // stale event from a superseded pop-out session

  let msg;
  try {
    msg = JSON.parse(arg.message);
  } catch (err) {
    console.warn('Pop-out dialog: received a malformed relay message', arg.message, err);
    return;
  }
  if (msg.type === 'popout-error') {
    console.error('Pop-out dialog reported an error', msg.reason);
    triggerPopoutFallback('The pop-out window could not display the decrypted content. Opening a regular window instead.');
  }
}

/**
 * Fires when the dialog closes — including an ordinary user-initiated close
 * (error code 12006), which is expected UX and shows no error message.
 *
 * The timer/channel cleanup runs unconditionally, before the 12006 check —
 * this used to run only on the non-12006 path (via triggerPopoutFallback),
 * which left two bugs on an ordinary close: a still-pending handshake timer
 * would fire ~10s later and pop an unprompted plaintext window via the
 * fallback, and the handshake channel — still holding the plaintext text in
 * its onmessage closure — was never closed at all.
 *
 * Any other code covers cases like the dialog failing to load (12002) or
 * other Dialog API runtime errors — those are real failures the previous
 * version of this handler silently discarded (it null'd out _popoutDialog
 * and nothing else), so it's treated as a signal to fall back.
 */
function onPopoutDialogClosed(dialog, arg) {
  if (_popoutDialog !== dialog) return; // stale event from a superseded pop-out session

  _popoutDialog = null;
  clearPopoutHandshakeTimer();
  closePopoutHandshakeChannel();
  if (arg?.error === 12006) return;
  console.error('Pop-out dialog closed unexpectedly, error code', arg?.error);
  triggerPopoutFallback('The pop-out window closed unexpectedly. Opening a regular window instead.');
}

// ── Verify signed-only body ───────────────────────────────────────────────────

async function handleVerifySignedBody(signedBody) {
  const statusEl = el('signed-body-status');
  const bodyEl = el('signed-body');

  // Extract the text between "-----BEGIN PGP SIGNED MESSAGE-----" and the signature
  const textMatch = signedBody.match(
    /-----BEGIN PGP SIGNED MESSAGE-----[\s\S]*?\n\n([\s\S]*?)\n-----BEGIN PGP SIGNATURE-----/
  );
  const plainText = textMatch ? textMatch[1] : '';

  if (plainText) {
    bodyEl.textContent = plainText;
    bodyEl.classList.remove('pgp-hidden');
  }

  statusEl.innerHTML = `<div class="pgp-alert pgp-alert--info"><span class="pgp-spinner"></span> Verifying signature…</div>`;

  try {
    const senderEmail = Office.context.mailbox.item.from?.emailAddress;
    const verificationKeys = await resolveVerificationKeys(senderEmail);

    if (verificationKeys.length === 0) {
      const noSenderDueToVersion = !senderEmail && !_has17;
      const hint = noSenderDueToVersion
        ? 'Sender information is unavailable on this Outlook version. Upgrade to Outlook 2021 to verify signatures.'
        : `No public key found for <strong>${escHtml(senderEmail || 'sender')}</strong>. Import their key via Manage Keys to verify future messages.`;
      statusEl.innerHTML = `<div class="pgp-alert pgp-alert--warning">${hint}</div>`;
      return;
    }

    const verifyResult = await verifyCleartextMessage(signedBody, verificationKeys);
    const sig = verifyResult.signatures[0];

    try {
      await sig.verified;
      statusEl.innerHTML = `<div class="pgp-alert pgp-alert--success">
        ✓ Valid signature from <strong>${escHtml(senderEmail || 'sender')}</strong>
      </div>`;
    } catch {
      statusEl.innerHTML = `<div class="pgp-alert pgp-alert--error">
        ✗ Invalid signature — this message may have been modified after signing.
      </div>`;
    }
  } catch (e) {
    statusEl.innerHTML = `<div class="pgp-alert pgp-alert--warning">Signature verification failed: ${escHtml(e.message)}</div>`;
  }
}

// ── Resolve sender's verification key ────────────────────────────────────────

async function resolveVerificationKeys(senderEmail) {
  if (!senderEmail) return [];
  try {
    // Try local keyring first (fast), then skip network lookup for read pane UX
    const localKey = await getContactKeyObject(senderEmail);
    if (localKey) return [localKey];

    // Try WKD/VKS silently
    const result = await discoverKey(senderEmail);
    return result.key ? [result.key] : [];
  } catch {
    return [];
  }
}

// ── Encrypted attachments ─────────────────────────────────────────────────────

// This add-in only ever produces .pgp attachments (see MessageCompose.js), but
// recognizes attachments encrypted by other PGP tools (GPG Suite, gpg4win, etc.)
// too, so they can be decrypted here as well.
const PGP_ATTACHMENT_EXTENSIONS = ['.pgp', '.gpg', '.asc'];

function renderPgpAttachments() {
  const item = Office.context.mailbox.item;
  const attachments = item.attachments || [];
  const pgpAttachments = attachments.filter(a =>
    !a.isInline && PGP_ATTACHMENT_EXTENSIONS.some(ext => a.name.toLowerCase().endsWith(ext))
  );

  if (pgpAttachments.length === 0) return;

  showSection('section-attachments');
  const list = el('attachment-list');
  list.innerHTML = '';

  pgpAttachments.forEach(att => {
    const li = document.createElement('li');
    li.className = 'pgp-attachment-item';

    const nameSpan = document.createElement('span');
    nameSpan.className = 'pgp-attachment-item__name';
    nameSpan.title = att.name;
    nameSpan.textContent = att.name;
    li.appendChild(nameSpan);

    // Attachment decryption requires Mailbox 1.8. On older clients render a
    // static note instead of a button that would fail at runtime.
    if (_has18) {
      const btn = document.createElement('button');
      btn.className = 'pgp-btn pgp-btn--secondary pgp-btn--sm btn-decrypt-att';
      btn.dataset.id   = att.id;
      btn.dataset.name = att.name;
      btn.textContent  = 'Decrypt & Download';
      li.appendChild(btn);
    } else {
      const note = document.createElement('span');
      note.style.fontSize = '11px';
      note.style.color    = '#797775';
      note.textContent    = 'Requires Outlook 2021 or Microsoft 365';
      li.appendChild(note);
    }

    list.appendChild(li);
  });

  list.addEventListener('click', async (e) => {
    const btn = e.target.closest('.btn-decrypt-att');
    if (!btn) return;

    btn.disabled = true;
    btn.textContent = '…';

    const attachmentId = btn.dataset.id;
    const attachmentName = btn.dataset.name;

    try {
      const savedName = await decryptAndDownloadAttachment(item, attachmentId, attachmentName);
      showStatus(`"${savedName}" decrypted and downloaded.`, 'success');
    } catch (e) {
      if (e.message !== 'Cancelled.') {
        showStatus(`Could not decrypt ${attachmentName}: ${e.message}`, 'error');
      }
    } finally {
      btn.disabled = false;
      btn.textContent = 'Decrypt & Download';
    }
  });

  const saveAllBtn = el('btn-save-all-attachments');
  if (_has18) {
    saveAllBtn.classList.remove('pgp-hidden');
    saveAllBtn.onclick = () => saveAllAttachments(item, pgpAttachments);
  } else {
    saveAllBtn.classList.add('pgp-hidden');
  }
}

// Decrypts a single PGP attachment and triggers its download. Shared by the
// per-item "Decrypt & Download" buttons and the "Save All" batch handler so
// both go through the same passphrase-unlock/decrypt/download path.
async function decryptAndDownloadAttachment(item, attachmentId, attachmentName) {
  let privateKey = getSessionKey();
  if (!privateKey) {
    const passphrase = await promptPassphrase(`Enter your passphrase to decrypt ${attachmentName}.`);
    privateKey = await unlockPrivateKey(getPrivateKey(), passphrase);
    const userEmail = Office.context.mailbox.userProfile?.emailAddress || '';
    const meta = getKeyMetadata();
    cacheSessionKey(privateKey, userEmail, meta?.keyId?.slice(-8) || '');
    updateSessionStatus();
  }

  const contentResult = await getAttachmentContentAsync(item, attachmentId);
  // Strip any non-ASCII characters before decoding — Office.js should return
  // clean RFC 4648 base64, but some Outlook Desktop builds append whitespace
  // or platform-specific characters to the attachment content string.
  const armoredMessage = atob(contentResult.content.replace(/[^\x00-\x7F]/g, ''));

  const { data: decryptedBytes, filename } = await decryptAttachment(armoredMessage, privateKey);

  const fallbackName = stripPgpExtension(attachmentName);
  const decryptedName = filename || fallbackName;
  const finalName = applyDecryptedExtensionPrefix(decryptedName, getDecryptedExtensionPrefix());
  downloadBytes(decryptedBytes, finalName);
  return finalName;
}

async function saveAllAttachments(item, pgpAttachments) {
  const saveAllBtn = el('btn-save-all-attachments');
  saveAllBtn.disabled = true;
  saveAllBtn.textContent = '…';

  let successCount = 0;
  const failures = [];

  for (const att of pgpAttachments) {
    try {
      await decryptAndDownloadAttachment(item, att.id, att.name);
      successCount++;
    } catch (e) {
      if (e.message === 'Cancelled.') break;
      failures.push(att.name);
    }
  }

  saveAllBtn.disabled = false;
  saveAllBtn.textContent = 'Save All';

  if (failures.length === 0 && successCount > 0) {
    showStatus(`${successCount} attachment(s) decrypted and downloaded.`, 'success');
  } else if (successCount > 0) {
    showStatus(`${successCount} saved, failed: ${failures.join(', ')}`, 'error');
  } else if (failures.length > 0) {
    showStatus(`Could not decrypt: ${failures.join(', ')}`, 'error');
  }
}

function getAttachmentContentAsync(item, attachmentId) {
  return new Promise((resolve, reject) => {
    item.getAttachmentContentAsync(attachmentId, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
      else reject(new Error(result.error.message));
    });
  });
}

function downloadBytes(bytes, filename) {
  const blob = new Blob([bytes]);
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 5000);
}

// ── Reply encrypted ───────────────────────────────────────────────────────────

// Office.js caps `htmlBody` for displayNewMessageForm/Async at 32 KB (32,768)
// characters and Outlook Classic enforces it synchronously, throwing
// Sys.ArgumentOutOfRangeException if exceeded. Stay comfortably under that so
// a large quoted message degrades gracefully instead of crashing the reply.
const MAX_REPLY_HTML_BODY_LENGTH = 31000;

// Appended to the quote when it had to be truncated to fit maxLength.
// Exported for tests only — handleReplyEncrypted detects truncation via
// buildQuotedReplyHtml()'s own returned `truncated` flag, not by
// string-searching the output for this text: the decrypted message could
// legitimately *contain* this exact string (e.g. quoting an earlier reply
// that itself got truncated), which would make an .includes() check
// false-positive on an otherwise normal-sized message.
export const REPLY_TRUNCATION_NOTICE = '<br><em>[Original message truncated — too large to quote in full]</em>';

/**
 * Builds the quoted-reply HTML block from a decrypted message, capped to fit
 * under Outlook's htmlBody size limit (see MAX_REPLY_HTML_BODY_LENGTH).
 *
 * If the fully-formatted HTML quote would exceed the limit, falls back to a
 * plain-text quote (never truncates raw HTML, which could leave unbalanced
 * tags). If even that is too large, truncates the text and appends a visible
 * truncation notice.
 *
 * @param {string} decryptedText
 * @param {boolean} decryptedIsHtml
 * @param {string} senderName
 * @param {string} sentDate
 * @param {number} [maxLength]
 * @returns {{html: string, truncated: boolean}} `html` is wrapped, size-capped
 *   HTML ready to use as formData.htmlBody; `truncated` is true whenever this
 *   call had to degrade the content to fit — either falling back from
 *   formatted HTML to plain text (losing all formatting, even if the
 *   resulting text itself ends up fitting), or actually cutting text and
 *   appending the truncation notice.
 */
export function buildQuotedReplyHtml(decryptedText, decryptedIsHtml, senderName, sentDate, maxLength = MAX_REPLY_HTML_BODY_LENGTH) {
  const quoteHeader = `<br>--- Original message${senderName ? ` from ${escHtml(senderName)}` : ''}${sentDate ? ` on ${escHtml(sentDate)}` : ''} ---<br>`;

  const wrapHtml = (innerHtml) =>
    `<br><div style="border-left:2px solid #888;padding-left:8px;margin-left:4px;">` + quoteHeader + innerHtml + `</div>`;
  const wrapText = (innerHtml) =>
    `<br><blockquote style="border-left:2px solid #888;padding-left:8px;margin-left:4px;">` + quoteHeader + innerHtml + `</blockquote>`;

  let bodyContent;
  let wrap;
  let truncated = false;
  if (decryptedIsHtml) {
    bodyContent = formatDecryptedContentAsHtml(decryptedText, true);
    wrap = wrapHtml;

    if (wrap(bodyContent).length > maxLength) {
      // Doesn't fit even with formatting — fall back to plain text so any
      // further truncation below can't leave unbalanced/broken HTML tags.
      // This is itself content degradation (all HTML formatting lost), not
      // just a possible follow-on truncation below, so it counts as
      // `truncated` even if the resulting plain text ends up fitting fine.
      truncated = true;
      bodyContent = formatDecryptedContentAsPlainTextHtml(decryptedText, true);
      wrap = wrapText;
    }
  } else {
    bodyContent = formatDecryptedContentAsHtml(decryptedText, false);
    wrap = wrapText;
  }

  if (wrap(bodyContent).length > maxLength) {
    truncated = true;
    const overhead = wrap('').length + REPLY_TRUNCATION_NOTICE.length;
    bodyContent = bodyContent.slice(0, Math.max(0, maxLength - overhead)) + REPLY_TRUNCATION_NOTICE;
  }

  const result = wrap(bodyContent);
  // Hard backstop: if maxLength is smaller than the fixed wrap+notice
  // overhead itself, the slice above still leaves `result` over budget.
  // Guarantee the size contract always holds, even for a degenerate
  // maxLength — this matters more than well-formed HTML in that case.
  const html = result.length > maxLength ? result.slice(0, maxLength) : result;
  return { html, truncated };
}

// How long to keep re-broadcasting the handoff before giving up and falling
// back to the original displayNewMessageForm + buildQuotedReplyHtml path.
const REPLY_HANDOFF_TIMEOUT_MS = 10000;
// How often to re-broadcast while waiting for MessageCompose.js's ack — a
// compose window that takes a couple of seconds to load still catches one of
// these, since BroadcastChannel drops messages posted before a listener
// exists with no queueing for late subscribers.
const REPLY_HANDOFF_BROADCAST_INTERVAL_MS = 400;

/**
 * Entry point for both reply buttons.
 *
 * Desktop / OWA: opens a compose window pre-filled with the quoted decrypted
 * body (if available) and asks the user to click Encrypt in the ribbon.
 * `displayReplyForm` / `displayReplyAllForm` are available on these platforms.
 *
 * Mobile: `displayReplyForm` and `displayReplyAllForm` are explicitly listed as
 * unsupported on Outlook iOS/Android in the Office.js docs.  Instead, the
 * in-pane compose section (`section-mobile-compose`) is shown.  The user types
 * their reply, taps "Encrypt Reply", and the armor is placed in a read-only
 * textarea and auto-copied to the clipboard.  They then start a normal reply in
 * Outlook and paste the armor as the message body.
 *
 * Note: on mobile, Reply and Reply All produce the same in-pane compose
 * experience — the distinction is not meaningful without a real compose window.
 *
 * @param {boolean} replyAll  - true → displayReplyAllForm on desktop (ignored on mobile)
 */
function handleReplyEncrypted(replyAll) {
  if (_isMobile) {
    openMobileCompose();
    return;
  }

  // Refuse to start a second large-message handoff while one from this same
  // pane is still in flight -- see _nativeReplyHandoffInFlight's docblock
  // and issue #17. The buttons are also disabled for the duration (below),
  // this is the functional backstop in case a click still gets through.
  if (_nativeReplyHandoffInFlight) {
    showStatus('A reply is already being set up — please wait for it to finish.', 'warning');
    return;
  }

  // ── Desktop flow ──────────────────────────────────────────────────────────
  const item    = Office.context.mailbox.item;
  const myEmail = (Office.context.mailbox.userProfile?.emailAddress || '').toLowerCase();

  // ── Subject ───────────────────────────────────────────────────────────────
  const origSubject = item.subject || '';
  const subject = /^re:\s/i.test(origSubject) ? origSubject : `Re: ${origSubject}`;

  // ── Recipients ────────────────────────────────────────────────────────────
  const toRecipients = [item.from?.emailAddress].filter(Boolean);

  let ccRecipients = [];
  if (replyAll) {
    const allOthers = [...(item.to || []), ...(item.cc || [])];
    ccRecipients = allOthers
      .map(r => r.emailAddress)
      .filter(addr => addr && addr.toLowerCase() !== myEmail);
  }

  // ── Quoted body ───────────────────────────────────────────────────────────
  let htmlBody = '';
  let wouldTruncate = false;
  if (_decryptedText) {
    const senderName = item.from?.displayName || item.from?.emailAddress || '';
    const sentDate    = item.dateTimeCreated
      ? item.dateTimeCreated.toLocaleString(undefined, {
          dateStyle: 'medium', timeStyle: 'short',
        })
      : '';
    ({ html: htmlBody, truncated: wouldTruncate } = buildQuotedReplyHtml(_decryptedText, _decryptedIsHtml, senderName, sentDate));
  }

  // buildQuotedReplyHtml() already had to truncate — the message is too big
  // for displayNewMessageForm/displayReplyForm's shared 32 KB htmlBody cap.
  // Route it through the native-reply + handoff path instead, which splices
  // the full decrypted content in from inside the new compose window via
  // Body.setAsync (1 MB limit) — see openNativeReplyWithHandoff. Everything
  // else (the common case: messages that fit) keeps today's exact behavior.
  if (wouldTruncate && typeof BroadcastChannel === 'function') {
    openNativeReplyWithHandoff(replyAll, toRecipients, ccRecipients, subject, htmlBody);
    return;
  }

  openReplyComposeForm(toRecipients, ccRecipients, subject, htmlBody, null);
}

/**
 * Reason a handoff attempt fell back to openReplyComposeForm(), used to pick
 * an accurate warning. 'timeout' and 'channel-failed' both happen *after*
 * displayReplyForm/displayReplyAllForm already succeeded, so a second
 * (still-blank/still-armored) native reply window is left open alongside
 * this one. 'no-scoping-id' and 'display-reply-failed' both happen *before*
 * any native reply was opened, so there is no second window to mention.
 * @enum {string}
 */
const HandoffFallbackReason = {
  TIMEOUT: 'timeout',
  CHANNEL_FAILED: 'channel-failed',
  NO_SCOPING_ID: 'no-scoping-id',
  DISPLAY_REPLY_FAILED: 'display-reply-failed',
};

/**
 * Opens a new-message compose form pre-filled with recipients/subject/quoted
 * body — the original Reply/Reply All mechanism. Used directly for
 * normal-sized messages, and as the safety-net fallback when the native-reply
 * handoff (openNativeReplyWithHandoff) can't be confirmed to have worked for
 * a large one.
 *
 * @param {?string} handoffFallbackReason - null for the normal-sized-message
 *   path; otherwise one of HandoffFallbackReason, used to show an accurate
 *   warning instead of the normal success message.
 */
function openReplyComposeForm(toRecipients, ccRecipients, subject, htmlBody, handoffFallbackReason) {
  const formData = { toRecipients, ccRecipients, subject, ...(htmlBody ? { htmlBody } : {}) };

  const onSuccess = () => {
    if (handoffFallbackReason === HandoffFallbackReason.TIMEOUT) {
      showStatus(
        "Reply setup didn't finish in time, so a backup reply window was opened with the message content — " +
        "formatting may have been shortened to fit Outlook's size limit. You can close the other blank reply " +
        'window that also opened.',
        'warning'
      );
    } else if (handoffFallbackReason === HandoffFallbackReason.CHANNEL_FAILED) {
      showStatus(
        "Reply setup couldn't be completed, so a backup reply window was opened with the message content — " +
        "formatting may have been shortened to fit Outlook's size limit. You can close the other blank reply " +
        'window that also opened.',
        'warning'
      );
    } else if (
      handoffFallbackReason === HandoffFallbackReason.NO_SCOPING_ID ||
      handoffFallbackReason === HandoffFallbackReason.DISPLAY_REPLY_FAILED
    ) {
      showStatus(
        "Reply setup couldn't be completed, so this reply window was opened with the message content — " +
        "formatting may have been shortened to fit Outlook's size limit.",
        'warning'
      );
    } else {
      showStatusReplyOpened();
    }
  };

  const onResult = r => {
    if (r && r.status === Office.AsyncResultStatus.Failed) {
      showStatus(`Could not open reply: ${r.error.message}`, 'error');
    } else {
      onSuccess();
    }
  };

  try {
    const mailbox = Office.context.mailbox;
    if (typeof mailbox.displayNewMessageFormAsync === 'function') {
      mailbox.displayNewMessageFormAsync(formData, onResult);
    } else {
      mailbox.displayNewMessageForm(formData);
      onSuccess();
    }
  } catch (e) {
    showStatus(`Could not open reply: ${e.message}`, 'error');
  }
}

/**
 * Large-message Reply/Reply All path.
 *
 * Opens Outlook's NATIVE reply (proper In-Reply-To/References threading,
 * recipients, and "Re:" subject — all handled by Outlook itself; self is
 * excluded from Reply All automatically) with no custom body, so it quotes
 * the original message as Outlook has it: still PGP-armored, since Outlook
 * has no notion of decryption. The decrypted plaintext is then handed to
 * MessageCompose.js over a BroadcastChannel so it can splice it in from
 * inside the new compose window's own script context, via Body.setAsync
 * (1 MB limit) — the only body-write path not bound by the 32 KB htmlBody
 * cap this whole path exists to avoid. See MessageCompose.js for the
 * receiving side (armor stripping + splice).
 *
 * BroadcastChannel has an inherent timing race: a message posted before the
 * new compose window's listener is ready is silently dropped, with no
 * queueing for late subscribers. This re-broadcasts on an interval until an
 * ack arrives; if none arrives within REPLY_HANDOFF_TIMEOUT_MS, it falls back
 * to openReplyComposeForm() (today's original path, truncation and all) —
 * opening a SECOND window, since Office.js gives no way to retract the
 * native reply already opened, or get a live handle to it.
 */
export function openNativeReplyWithHandoff(replyAll, toRecipients, ccRecipients, subject, htmlBody) {
  const item = Office.context.mailbox.item;
  const fallBack = (reason) => openReplyComposeForm(toRecipients, ccRecipients, subject, htmlBody, reason);

  // Prefer conversationId, but fall back to internetMessageId (available
  // since Mailbox 1.1, broader than conversationId) when it's missing --
  // MessageCompose.js derives the matching scoping ID the same way, from
  // its own item.conversationId / item.inReplyTo (the internet message ID
  // of the message it's replying to). See setupReplyHandoffListener.
  const scopingId = item.conversationId || item.internetMessageId;
  if (!scopingId) {
    // No way to scope the handoff channel to this specific conversation/
    // message at all -- falling back to the shared base channel name would
    // let any same-origin page listen for this (and every other) large
    // reply's decrypted plaintext. Treat this as "handoff unavailable" and
    // skip straight to the existing path, rather than opening a native
    // reply we can't safely hand plaintext to.
    console.error('Native reply: no conversationId or internetMessageId available, skipping handoff');
    fallBack(HandoffFallbackReason.NO_SCOPING_ID);
    return;
  }

  // Captured now, not read live from the broadcast closure below: the user
  // could decrypt a *different* message in the reading pane while this
  // handoff is still retrying (up to REPLY_HANDOFF_TIMEOUT_MS later), which
  // would otherwise reassign _decryptedText/_decryptedIsHtml out from under
  // an in-flight handoff and broadcast the wrong message's plaintext into
  // this reply.
  const decryptedText = _decryptedText;
  const decryptedIsHtml = _decryptedIsHtml;

  try {
    // formData is a required parameter for both APIs (not optional despite
    // being commonly called with no visible effect) -- calling these with
    // zero arguments throws synchronously on every invocation. Passing the
    // bare HANDOFF_PENDING_MARKER string here was confirmed live to insert
    // NO visible text on classic Outlook Desktop, despite Microsoft's own
    // docs example showing exactly that usage -- no documented or
    // community-confirmed explanation found (see #22). Using the
    // ReplyFormData OBJECT form instead (`{ htmlBody }`), matching the
    // field name openReplyComposeForm()/displayNewMessageForm already uses
    // successfully elsewhere in this file, is the only structurally
    // different invocation the docs describe, and is what actually works.
    // Outlook still prepends this above its own native quote (so "let
    // Outlook build the native reply" stays untouched) -- see
    // HANDOFF_PENDING_MARKER's docblock for why: it's how
    // Functions/ReplyHandoffRuntime.classic.js knows this specific reply is
    // one this add-in opened for a handoff, and applyReplyHandoff() removes
    // it as part of the same splice that replaces the PGP armor.
    const formData = { htmlBody: HANDOFF_PENDING_MARKER };
    if (replyAll) item.displayReplyAllForm(formData);
    else item.displayReplyForm(formData);
  } catch (e) {
    console.error('Native reply: displayReplyForm/displayReplyAllForm failed', e);
    fallBack(HandoffFallbackReason.DISPLAY_REPLY_FAILED);
    return;
  }

  // The native reply window is now open, and stays open for up to
  // REPLY_HANDOFF_TIMEOUT_MS while this handoff retries -- see #17. Guard
  // this window against a second concurrent handoff from the same pane.
  _nativeReplyHandoffInFlight = true;
  setReplyButtonsDisabled(true);

  const channelName = getReplyHandoffChannelName(scopingId);
  let channel;
  try {
    channel = new BroadcastChannel(channelName);
  } catch (e) {
    console.error('Native reply: BroadcastChannel construction failed', e);
    _nativeReplyHandoffInFlight = false;
    setReplyButtonsDisabled(false);
    fallBack(HandoffFallbackReason.CHANNEL_FAILED);
    return;
  }
  console.log('Native reply: broadcasting on', channelName);

  const token = generateChannelToken();
  let settled = false;
  let broadcastTimer;
  let giveUpTimer;

  const finish = (useFallback) => {
    if (settled) return;
    settled = true;
    clearInterval(broadcastTimer);
    clearTimeout(giveUpTimer);
    channel.close();
    _nativeReplyHandoffInFlight = false;
    setReplyButtonsDisabled(false);
    if (useFallback) fallBack(HandoffFallbackReason.TIMEOUT);
    else showStatusReplyOpened();
  };

  channel.onmessage = (event) => {
    if (event.data?.type === 'pgp-reply-handoff-ack' && event.data.token === token) {
      finish(false);
    }
  };

  const broadcast = () => {
    channel.postMessage({ type: 'pgp-reply-handoff', token, text: decryptedText, isHtml: decryptedIsHtml });
  };
  broadcast();
  broadcastTimer = setInterval(broadcast, REPLY_HANDOFF_BROADCAST_INTERVAL_MS);
  giveUpTimer = setTimeout(() => finish(true), REPLY_HANDOFF_TIMEOUT_MS);
}

// ── Mobile inline compose ─────────────────────────────────────────────────────

/**
 * Show the inline compose section, pre-populated with a plain-text quote of
 * the already-decrypted body (if available) so the user sees context.
 */
function openMobileCompose() {
  const textarea = el('mobile-compose-body');
  const statusEl = el('mobile-compose-status');

  // Reset to write mode in case a previous encryption result is still showing.
  textarea.readOnly = false;
  textarea.style.fontFamily = '';
  textarea.style.fontSize = '';
  el('mobile-compose-title').textContent = 'Compose Encrypted Reply';
  el('mobile-copy-instructions').classList.add('pgp-hidden');
  el('btn-mobile-encrypt-send').classList.remove('pgp-hidden');
  el('btn-mobile-copy-armor').classList.add('pgp-hidden');
  el('btn-mobile-copy-armor').textContent = 'Copy';
  statusEl.classList.add('pgp-hidden');

  if (_decryptedText && !_decryptedIsHtml) {
    const item = Office.context.mailbox.item;
    const senderName = item.from?.displayName || item.from?.emailAddress || '';
    const header = senderName
      ? `\n\n--- Original message from ${senderName} ---\n`
      : '\n\n--- Original message ---\n';
    textarea.value = header + _decryptedText;
    // Position cursor at the very top so the user types above the quote.
    textarea.setSelectionRange(0, 0);
    textarea.scrollTop = 0;
  } else {
    textarea.value = '';
  }

  // Show whether signing will be applied.
  const keyStatusEl = el('mobile-compose-key-status');
  if (getSessionKey()) {
    if (getSignDefault()) {
      keyStatusEl.textContent =
        `Will sign with cached key · ${getSessionEmail() || ''}`;
    } else {
      keyStatusEl.textContent = 'Message will be encrypted (signing is off by default).';
    }
  } else {
    keyStatusEl.textContent =
      'Message will be encrypted without a signature ' +
      '(decrypt the incoming message first to cache your key for signing).';
  }
  keyStatusEl.classList.remove('pgp-hidden');

  showSection('section-mobile-compose');
  textarea.focus();
}

/**
 * Encrypt the text typed in the mobile compose textarea.
 *
 * NOTE: displayReplyForm / displayReplyAllForm are explicitly listed as
 * unsupported on Outlook mobile in the Office.js docs.  There is no API that
 * opens a pre-filled compose window from a read-mode task pane on mobile.
 *
 * Instead we encrypt the text here in the task pane, replace the textarea with
 * the PGP armor (read-only), and expose a Copy button.  The user taps Copy,
 * starts a reply manually in Outlook, and pastes the armor as the body.
 *
 * Recipient keys: sender's public key (discovered via keyring / WKD / VKS) +
 * the user's own public key (encrypt-to-self so sent mail is readable).
 * Signing: applied only when the user's key is already unlocked in the session
 * cache AND signing is their stored default — no extra passphrase prompt needed.
 */
async function handleMobileEncryptReply() {
  const textarea  = el('mobile-compose-body');
  const btn       = el('btn-mobile-encrypt-send');
  const spinner   = el('mobile-encrypt-spinner');
  const statusEl  = el('mobile-compose-status');

  const text = textarea.value.trim();
  if (!text) {
    statusEl.textContent = 'Please type a reply before encrypting.';
    statusEl.className = 'pgp-alert pgp-alert--warning';
    statusEl.classList.remove('pgp-hidden');
    return;
  }

  btn.disabled = true;
  spinner.classList.remove('pgp-hidden');
  statusEl.classList.add('pgp-hidden');

  try {
    const item        = Office.context.mailbox.item;
    const senderEmail = item.from?.emailAddress;

    // ── Discover the sender's public key ───────────────────────────────────
    if (!senderEmail) {
      // item.from is unavailable on Mailbox < 1.7 (Outlook 2019). Show a
      // graceful message rather than throwing an unhandled error.
      statusEl.textContent = !_has17
        ? 'Encrypted reply requires Outlook 2021 or Microsoft 365 — sender information is unavailable on this Outlook version.'
        : 'Cannot determine the sender\'s email address.';
      statusEl.className = 'pgp-alert pgp-alert--error';
      statusEl.classList.remove('pgp-hidden');
      return;
    }
    const { key: senderKey } = await discoverKey(senderEmail);
    if (!senderKey) {
      statusEl.innerHTML =
        `No public key found for <strong>${escHtml(senderEmail)}</strong>. ` +
        `Ask them to share their public key, or have them publish it via ` +
        `WKD / keys.openpgp.org, then try again.`;
      statusEl.className = 'pgp-alert pgp-alert--error';
      statusEl.classList.remove('pgp-hidden');
      return;
    }

    // ── Build recipient list (sender + self) ───────────────────────────────
    const recipientKeys = [senderKey];
    const ownArmoredPub = getPublicKey();
    if (ownArmoredPub) {
      try { recipientKeys.push(await readPublicKey(ownArmoredPub)); } catch { /* skip */ }
    }

    // ── Optional signing (session key must already be cached) ──────────────
    const signingKey = (getSessionKey() && getSignDefault()) ? getSessionKey() : null;

    // ── Encrypt ────────────────────────────────────────────────────────────
    const armor = await encryptMessage(text, recipientKeys, signingKey);

    // ── Show result + Copy button ───────────────────────────────────────────
    // displayReplyForm/displayReplyAllForm are not supported on Outlook mobile.
    // Show the armor in the (now read-only) textarea and let the user copy it.
    textarea.value = armor;
    textarea.readOnly = true;
    textarea.style.fontFamily = 'monospace';
    textarea.style.fontSize = '11px';

    el('mobile-compose-title').textContent = 'Encrypted Reply Ready';
    el('mobile-compose-key-status').classList.add('pgp-hidden');
    el('mobile-copy-instructions').classList.remove('pgp-hidden');
    btn.classList.add('pgp-hidden');
    el('btn-mobile-copy-armor').classList.remove('pgp-hidden');

    // Attempt auto-copy so the user just needs to paste.
    try {
      await navigator.clipboard.writeText(armor);
      statusEl.textContent = 'Copied! Start a reply in Outlook and paste as the message body.';
      statusEl.className = 'pgp-alert pgp-alert--info';
      statusEl.classList.remove('pgp-hidden');
    } catch {
      // Clipboard API unavailable — user will tap the Copy button manually.
    }

  } catch (e) {
    statusEl.textContent = `Encryption failed: ${e.message}`;
    statusEl.className = 'pgp-alert pgp-alert--error';
    statusEl.classList.remove('pgp-hidden');
  } finally {
    btn.disabled = false;
    spinner.classList.add('pgp-hidden');
  }
}

// ── Office.js wrappers ────────────────────────────────────────────────────────

function getBodyAsync(coercionType) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.getAsync(coercionType, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
      else reject(new Error(result.error.message));
    });
  });
}

// ── Bootstrap ─────────────────────────────────────────────────────────────────

Office.onReady(async () => {
  // Detect mobile early so the reply section description is correct.
  const platform = Office.context.diagnostics?.platform;
  _isMobile = platform === 'Android' || platform === 'iOS';

  // Capability flags — evaluated once after Office.js has initialized.
  _has17 = Office.context.requirements.isSetSupported('Mailbox', '1.7');
  _has18 = Office.context.requirements.isSetSupported('Mailbox', '1.8');
  _has14 = Office.context.requirements.isSetSupported('Mailbox', '1.4');

  // Load org config (e.g. companyDecryptedExtensionPrefix) before attachments
  // are rendered/decrypted, so getDecryptedExtensionPrefix() reads populated data.
  const userEmail = Office.context.mailbox.userProfile?.emailAddress || '';
  await loadOrgConfig(userEmail);

  if (_isMobile) {
    el('reply-desktop-hint').classList.add('pgp-hidden');
    el('reply-mobile-hint').classList.remove('pgp-hidden');
  }

  // On Mailbox < 1.7 item.from is undefined, so recipient auto-fill is
  // unavailable. Surface a note below the reply buttons so users aren't
  // surprised when the compose window opens with an empty To field.
  if (!_has17) {
    el('reply-sender-note').classList.remove('pgp-hidden');
  }

  // section-reply is visible on all platforms (mobile hint/desktop hint swap above).

  // Wire reply buttons regardless of key state — the user may want to reply
  // encrypted even if they have no local key pair yet.
  el('btn-reply-encrypted').addEventListener('click', () => handleReplyEncrypted(false));
  el('btn-reply-all-encrypted').addEventListener('click', () => handleReplyEncrypted(true));

  // Mobile inline compose buttons.
  el('btn-mobile-encrypt-send').addEventListener('click', handleMobileEncryptReply);
  el('btn-mobile-copy-armor').addEventListener('click', async () => {
    const textarea = el('mobile-compose-body');
    const armor = textarea.value;
    let copied = false;

    // Modern Clipboard API (Android, desktop).
    if (navigator.clipboard?.writeText) {
      try {
        await navigator.clipboard.writeText(armor);
        copied = true;
      } catch { /* fall through */ }
    }

    // iOS fallback: select the textarea content and use execCommand.
    // Must be synchronous and triggered directly by the user gesture.
    if (!copied) {
      try {
        textarea.setSelectionRange(0, armor.length);
        textarea.focus();
        copied = document.execCommand('copy');
      } catch { /* execCommand also unavailable */ }
    }

    if (copied) {
      el('btn-mobile-copy-armor').textContent = 'Copied!';
      setTimeout(() => { el('btn-mobile-copy-armor').textContent = 'Copy'; }, 2000);
    }
  });
  el('btn-mobile-compose-cancel').addEventListener('click', () => {
    // Reset compose section back to write mode for next use.
    const textarea = el('mobile-compose-body');
    textarea.value = '';
    textarea.readOnly = false;
    textarea.style.fontFamily = '';
    textarea.style.fontSize = '';
    el('mobile-compose-title').textContent = 'Compose Encrypted Reply';
    el('mobile-copy-instructions').classList.add('pgp-hidden');
    el('btn-mobile-encrypt-send').classList.remove('pgp-hidden');
    el('btn-mobile-copy-armor').classList.add('pgp-hidden');
    el('btn-mobile-copy-armor').textContent = 'Copy';
    el('mobile-compose-status').classList.add('pgp-hidden');
    hideSection('section-mobile-compose');
  });

  if (!hasKeyPair()) {
    el('panel-no-key').classList.remove('pgp-hidden');
    el('detection-loading').classList.add('pgp-hidden');
    el('detection-result').innerHTML = `<div class="pgp-alert pgp-alert--warning">Generate a key pair first to use decryption.</div>`;
    el('detection-result').classList.remove('pgp-hidden');
    return;
  }

  // Reflect current session cache state and keep it in sync
  updateSessionStatus();
  onSessionCleared(updateSessionStatus);

  el('btn-lock-session').addEventListener('click', () => {
    clearSessionKey(); // triggers onSessionCleared → updateSessionStatus
    // Locking is explicit user intent to stop showing decrypted content —
    // close any open pop-out dialog rather than leaving it displaying
    // plaintext after the session key it depended on is gone.
    closePopoutDialogQuietly();
  });

  await detectAndRenderBody();

  // window.open() is not available in Outlook mobile WebViews.
  if (_isMobile) {
    el('btn-popout-decrypted').style.display = 'none';
  }
});
