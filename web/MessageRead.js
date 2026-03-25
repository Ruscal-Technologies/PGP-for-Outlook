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

import * as openpgp from './js/openpgp.min.mjs';
import {
  unlockPrivateKey,
  decryptMessage, decryptAttachment,
  detectPgpContent,
  encryptMessage, readPublicKey,
} from './js/pgp/pgp-core.js';
import { hasKeyPair, getPrivateKey, getPublicKey, getSignDefault } from './js/pgp/key-storage.js';
import { getContactKeyObject } from './js/pgp/keyring.js';
import { discoverKey, KeyStatus } from './js/pgp/key-discovery.js';
import {
  cacheSessionKey, getSessionKey, clearSessionKey,
  getSessionEmail, getSessionShortId, onSessionCleared,
} from './js/pgp/session-cache.js';

// ── Module state ──────────────────────────────────────────────────────────────

/** Decrypted payload, stored so reply handlers can quote it. */
let _decryptedText = null;
let _decryptedIsHtml = false;

/** True when running inside Outlook on iOS or Android. */
let _isMobile = false;


// ── Helpers ───────────────────────────────────────────────────────────────────

function el(id) { return document.getElementById(id); }

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
  bar.innerHTML = message;
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

  // Normalise line endings to LF, then trim trailing whitespace per line.
  text = text
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .split('\n')
    .map(line => line.trimEnd())
    .join('\n');

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

  function walk(node) {
    if (node.nodeType === Node.TEXT_NODE) return node.textContent;
    if (node.nodeType !== Node.ELEMENT_NODE) return '';
    const tag = node.tagName.toLowerCase();
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
    el('btn-decrypt').addEventListener('click', () => handleDecryptBody(body));
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
      cacheSessionKey(privateKey, userEmail, '');
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
      showStatus(`Decryption failed: ${escHtml(e.message)}`, 'error');
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
    openDecryptedPopup(text, isHtml, subject);
  });
}

// ── Pop-out window ─────────────────────────────────────────────────────────────

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
  }
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
      statusEl.innerHTML = `<div class="pgp-alert pgp-alert--warning">
        Cannot verify signature — no public key found for <strong>${escHtml(senderEmail || 'sender')}</strong>.
        Import their key via Manage Keys to verify future messages.
      </div>`;
      return;
    }

    // Use openpgp.js directly for clearsigned message verification
    const cleartextMessage = await openpgp.readCleartextMessage({ cleartextMessage: signedBody });
    const verifyResult = await openpgp.verify({ message: cleartextMessage, verificationKeys });
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

function renderPgpAttachments() {
  const item = Office.context.mailbox.item;
  const attachments = item.attachments || [];
  const pgpAttachments = attachments.filter(a => !a.isInline && a.name.endsWith('.pgp'));

  if (pgpAttachments.length === 0) return;

  showSection('section-attachments');
  const list = el('attachment-list');
  list.innerHTML = '';

  pgpAttachments.forEach(att => {
    const li = document.createElement('li');
    li.className = 'pgp-attachment-item';
    li.innerHTML = `
      <span class="pgp-attachment-item__name" title="${escHtml(att.name)}">${escHtml(att.name)}</span>
      <button class="pgp-btn pgp-btn--secondary pgp-btn--sm btn-decrypt-att" data-id="${escHtml(att.id)}" data-name="${escHtml(att.name)}">
        Decrypt &amp; Download
      </button>`;
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
      let privateKey = getSessionKey();
      if (!privateKey) {
        const passphrase = await promptPassphrase(`Enter your passphrase to decrypt ${attachmentName}.`);
        privateKey = await unlockPrivateKey(getPrivateKey(), passphrase);
        const userEmail = Office.context.mailbox.userProfile?.emailAddress || '';
        cacheSessionKey(privateKey, userEmail, '');
        updateSessionStatus();
      }

      const contentResult = await getAttachmentContentAsync(item, attachmentId);
      const armoredMessage = atob(contentResult.content);

      const { data: decryptedBytes, filename } = await decryptAttachment(armoredMessage, privateKey);

      // Trigger browser download
      downloadBytes(decryptedBytes, filename || attachmentName.replace(/\.pgp$/i, ''));
      showStatus(`"${filename || attachmentName}" decrypted and downloaded.`, 'success');

    } catch (e) {
      if (e.message !== 'Cancelled.') {
        showStatus(`Could not decrypt ${attachmentName}: ${escHtml(e.message)}`, 'error');
      }
    } finally {
      btn.disabled = false;
      btn.textContent = 'Decrypt & Download';
    }
  });
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

  // ── Desktop flow ──────────────────────────────────────────────────────────
  const item = Office.context.mailbox.item;
  let quotedBody = '';

  if (_decryptedText) {
    const senderName  = item.from?.displayName  || item.from?.emailAddress || '';
    const quoteHeader = senderName
      ? `<br>--- Original message from ${escHtml(senderName)} ---<br>`
      : '<br>--- Original message ---<br>';

    if (_decryptedIsHtml) {
      quotedBody =
        `<br><div style="border-left:2px solid #888;padding-left:8px;margin-left:4px;">` +
        quoteHeader + _decryptedText + `</div>`;
    } else {
      const safe = _decryptedText
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/\n/g, '<br>');
      quotedBody =
        `<br><blockquote style="border-left:2px solid #888;padding-left:8px;margin-left:4px;">` +
        quoteHeader + safe + `</blockquote>`;
    }
  }

  try {
    if (replyAll) {
      item.displayReplyAllForm(quotedBody);
    } else {
      item.displayReplyForm(quotedBody);
    }
    showStatus(
      'Reply opened — click <strong>Encrypt</strong> in the ribbon to encrypt before sending.',
      'info'
    );
  } catch (e) {
    showStatus(`Could not open reply: ${escHtml(e.message)}`, 'error');
  }
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
      throw new Error('Cannot determine the sender\'s email address.');
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

  if (_isMobile) {
    el('reply-desktop-hint').classList.add('pgp-hidden');
    el('reply-mobile-hint').classList.remove('pgp-hidden');
  } else {
    // On desktop and OWA the ribbon's "Reply Encrypted" button handles this;
    // the in-pane reply section is only needed on mobile.
    hideSection('section-reply');
  }

  // Wire reply buttons regardless of key state — the user may want to reply
  // encrypted even if they have no local key pair yet.
  el('btn-reply-encrypted').addEventListener('click', () => handleReplyEncrypted(false));
  el('btn-reply-all-encrypted').addEventListener('click', () => handleReplyEncrypted(true));

  // Mobile inline compose buttons.
  el('btn-mobile-encrypt-send').addEventListener('click', handleMobileEncryptReply);
  el('btn-mobile-copy-armor').addEventListener('click', async () => {
    const armor = el('mobile-compose-body').value;
    try {
      await navigator.clipboard.writeText(armor);
      el('btn-mobile-copy-armor').textContent = 'Copied!';
      setTimeout(() => { el('btn-mobile-copy-armor').textContent = 'Copy'; }, 2000);
    } catch {
      // Clipboard API blocked — user must long-press the textarea to copy manually.
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
  });

  await detectAndRenderBody();

  // window.open() is not available in Outlook mobile WebViews.
  if (_isMobile) {
    el('btn-popout-decrypted').style.display = 'none';
  }
});
