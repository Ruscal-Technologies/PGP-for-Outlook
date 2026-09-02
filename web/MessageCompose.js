'use strict';
/**
 * MessageCompose.js
 * Task pane for encrypting outgoing messages.
 *
 * Flow:
 *  1. On load, all To/CC recipients are resolved to their PGP public keys using
 *     the key-discovery chain: local keyring → WKD → VKS (keys.openpgp.org).
 *  2. For each recipient without a key the user can trigger a fresh search, or
 *     paste an armored key directly.  Keys discovered from WKD/VKS can be saved
 *     to the local keyring with one click.
 *  3. The company key (if org config enables it) is fetched from WKD/VKS and
 *     added to every encryption unconditionally (or optionally, per org policy).
 *  4. The "Sign this message" toggle is initialized from the user's stored
 *     pgp_sign_default preference (default: false / off).  The user can flip
 *     it for any individual message regardless of the stored default.
 *  5. Clicking "Encrypt Message":
 *       a. If signing is enabled, checks the session cache for an already-
 *          unlocked key.  If none, prompts for the passphrase, unlocks, and
 *          caches it for 15 minutes of inactivity.
 *       b. Assembles the full recipient list: all To/CC keys + own public key
 *          (encrypt-to-self so the sender can read sent mail) + company key(s).
 *       c. Gets the message body as HTML (preserving all formatting), encrypts
 *          the HTML string, then replaces the body with the plain-text PGP armor.
 *          When the recipient decrypts, they recover the original HTML exactly.
 *       d. For each non-inline attachment: reads, encrypts to a .pgp file,
 *          removes the original, and adds the encrypted version.
 *  6. After encryption the Encrypt button is disabled so the message cannot be
 *     double-encrypted.  The user then sends the message normally.
 *
 * Requires: Mailbox 1.5 minimum (ribbon buttons, body encrypt/decrypt).
 * Attachment encryption requires Mailbox 1.8 and is gated at runtime via _has18.
 */

import {
  unlockPrivateKey, readPublicKey, getKeyInfo,
  encryptMessage, encryptAttachment,
  hasWeakEncryptionKey,
  base64ToUint8Array,
  detectPgpContent,
} from './js/pgp/pgp-core.js';
import { hasKeyPair, getPrivateKey, getPublicKey, getSignDefault } from './js/pgp/key-storage.js';
import {
  cacheSessionKey, getSessionKey, clearSessionKey,
  getSessionEmail, getSessionShortId, onSessionCleared,
} from './js/pgp/session-cache.js';
import { resolveRecipients, KeyStatus } from './js/pgp/key-discovery.js';
import {
  loadOrgConfig, isCompanyKeyEnabled, isCompanyKeyRequired,
  getCompanyKeyEmails, fetchCompanyKeys,
} from './js/pgp/org-config.js';
import { formatDecryptedContentAsHtml } from './js/pgp/quoted-content.js';
import { getReplyHandoffChannelName } from './js/pgp/reply-handoff-channel.js';

// ── Session status ────────────────────────────────────────────────────────────

/**
 * Refresh the session status bar that shows whether an unlocked private key is
 * currently cached.  Called on load, after caching a new key, and whenever the
 * cache is cleared (via the onSessionCleared callback registered in onReady).
 */
function updateSessionStatus() {
  const bar   = el('session-status');
  const label = el('session-status-text');

  const email   = getSessionEmail();
  const shortId = getSessionShortId();

  if (email) {
    label.textContent = `Key unlocked: ${email}${shortId ? ' ·  …' + shortId : ''}`;
    bar.classList.remove('pgp-hidden');
  } else {
    bar.classList.add('pgp-hidden');
  }
}

// ── State ─────────────────────────────────────────────────────────────────────

/** @type {Array<{email:string, key:openpgp.Key|null, status:string, source:string|null, armoredKey:string|null}>} */
let _recipientResults = [];

/** @type {Array<{id:string, name:string, contentType:string, size:number}>} */
let _attachments = [];

/** @type {Array<{id:string, name:string, contentType:string, size:number}>} */
let _inlineAttachments = [];

/** @type {Array<{email:string, key:openpgp.Key}>} */
let _companyKeys = [];

/**
 * True when the add-in is running in Outlook on the web (OWA).
 * Set once in Office.onReady — the platform never changes during a session.
 * Used to show/hide the inline-attachment Convert option (not available on
 * desktop because the Office API does not expose clipboard-pasted images).
 * @type {boolean}
 */
let _isWebOutlook = false;

/**
 * True when the host meets Mailbox 1.8 (Outlook 2021+ / Microsoft 365 / OWA).
 * Required for getAttachmentContentAsync and addFileAttachmentFromBase64Async.
 * Set once in Office.onReady via Office.context.requirements.isSetSupported().
 * @type {boolean}
 */
let _has18 = false;

/**
 * True when the host meets Mailbox 1.10. Required for getComposeTypeAsync,
 * used to confirm this compose window is actually a reply before setting up
 * the reply-handoff BroadcastChannel listener (see setupReplyHandoffListener).
 * Set once in Office.onReady via Office.context.requirements.isSetSupported().
 * @type {boolean}
 */
let _has110 = false;

/**
 * True when the host meets Mailbox 1.14. Required for item.inReplyTo, used
 * as a fallback scoping ID for the reply-handoff BroadcastChannel listener
 * when item.conversationId is unavailable (see setupReplyHandoffListener).
 * Set once in Office.onReady via Office.context.requirements.isSetSupported().
 * @type {boolean}
 */
let _has114 = false;

// ── Helpers ───────────────────────────────────────────────────────────────────

function el(id) { return document.getElementById(id); }

function escHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function showStatus(message, type = 'info') {
  const bar = el('status-bar');
  bar.className = `pgp-alert pgp-alert--${type}`;
  bar.textContent = message;
  bar.classList.remove('pgp-hidden');
}

function clearStatus() {
  el('status-bar').classList.add('pgp-hidden');
}

function statusBadge(status, source) {
  switch (status) {
    case KeyStatus.FOUND_LOCAL:
      return `<span class="pgp-badge pgp-badge--success">✓ Key found <span style="font-weight:400">(local keyring)</span></span>`;
    case KeyStatus.FOUND_WKD:
      return `<span class="pgp-badge pgp-badge--success">✓ Key found <span style="font-weight:400">(WKD)</span></span>`;
    case KeyStatus.FOUND_VKS:
      return `<span class="pgp-badge pgp-badge--success">✓ Key found <span style="font-weight:400">(${escHtml(source)})</span></span>`;
    case KeyStatus.NOT_FOUND:
      return `<span class="pgp-badge pgp-badge--error">✗ No key found</span>`;
    default:
      return `<span class="pgp-badge pgp-badge--warning">? Unknown</span>`;
  }
}

// ── Recipient resolution ──────────────────────────────────────────────────────

async function loadRecipients() {
  el('recipients-loading').classList.remove('pgp-hidden');
  el('recipient-list').classList.add('pgp-hidden');
  el('recipients-empty').classList.add('pgp-hidden');

  const item = Office.context.mailbox.item;

  // Gather all recipient fields
  const [toRaw, ccRaw] = await Promise.all([
    getRecipientsAsync(item.to),
    getRecipientsAsync(item.cc),
  ]);

  const all = [...(toRaw || []), ...(ccRaw || [])];

  if (all.length === 0) {
    el('recipients-loading').classList.add('pgp-hidden');
    el('recipients-empty').classList.remove('pgp-hidden');
    updateEncryptButton();
    return;
  }

  const emails = all.map(r => r.emailAddress);
  _recipientResults = await resolveRecipients(emails);

  renderRecipientList();
  updateEncryptButton();
}

// Outlook's getAsync() only returns recipients it has *finished* resolving —
// a recipient that's still being "checked" (e.g. right after "Reply
// encrypted" pre-populates To/Cc, or moments after the user finishes typing)
// is simply omitted, not returned in some half-resolved state. There's no
// public API to force that resolution to happen; polling until two
// consecutive reads agree is the documented workaround.
function getRecipientsAsync(recipients, maxAttempts = 5, delayMs = 300) {
  return new Promise((resolve) => {
    let previous = null;
    let attempt = 0;

    const poll = () => {
      recipients.getAsync((result) => {
        const value = result.status === Office.AsyncResultStatus.Succeeded ? result.value : [];
        attempt++;
        if ((previous !== null && value.length === previous.length) || attempt >= maxAttempts) {
          resolve(value);
        } else {
          previous = value;
          setTimeout(poll, delayMs);
        }
      });
    };

    poll();
  });
}

function renderRecipientList() {
  const list = el('recipient-list');
  list.innerHTML = '';
  el('recipients-loading').classList.add('pgp-hidden');
  list.classList.remove('pgp-hidden');

  _recipientResults.forEach((r, idx) => {
    const li = document.createElement('li');
    li.className = 'pgp-recipient-item';
    li.dataset.idx = idx;

    const hasKey = !!r.key;
    let actionsHtml = '';
    if (!hasKey) {
      actionsHtml = `
        <button class="pgp-btn pgp-btn--secondary pgp-btn--sm btn-retry-key" data-idx="${idx}">Search</button>
        <button class="pgp-btn pgp-btn--secondary pgp-btn--sm btn-paste-key" data-idx="${idx}">Paste Key</button>`;
    }

    li.innerHTML = `
      <div class="pgp-recipient-item__header">
        <span class="pgp-recipient-item__email">${escHtml(r.email)}</span>
        <div class="pgp-recipient-item__actions">
          ${actionsHtml}
        </div>
      </div>
      <div>${statusBadge(r.status, r.source)}</div>
      <div id="recipient-paste-form-${idx}" class="pgp-hidden" style="margin-top:8px;">
        <textarea id="recipient-paste-key-${idx}" rows="4" style="width:100%;box-sizing:border-box;font-family:monospace;font-size:11px;padding:4px;border:1px solid #8a8886;border-radius:2px;" placeholder="-----BEGIN PGP PUBLIC KEY BLOCK-----"></textarea>
        <div class="pgp-row pgp-mt-sm">
          <button class="pgp-btn pgp-btn--primary pgp-btn--sm btn-paste-key-confirm" data-idx="${idx}">Use This Key</button>
          <button class="pgp-btn pgp-btn--secondary pgp-btn--sm btn-paste-key-cancel" data-idx="${idx}">Cancel</button>
        </div>
      </div>`;

    list.appendChild(li);
  });
}

// ── Company key panel ─────────────────────────────────────────────────────────

async function loadCompanyKeys() {
  if (!isCompanyKeyEnabled()) {
    el('company-key-disabled').classList.remove('pgp-hidden');
    el('company-key-panel').classList.add('pgp-hidden');
    return;
  }

  // When the company key is required by IT policy the user has no choices to
  // make here, so hide the entire section rather than showing a locked toggle.
  if (isCompanyKeyRequired()) {
    el('section-company-key').classList.add('pgp-hidden');
    _companyKeys = await fetchCompanyKeys();
    return;
  }

  el('company-key-disabled').classList.add('pgp-hidden');
  el('company-key-panel').classList.remove('pgp-hidden');

  _companyKeys = await fetchCompanyKeys();

  const list = el('company-key-list');
  list.innerHTML = '';

  if (_companyKeys.length === 0) {
    list.innerHTML = `<li class="pgp-empty">⚠ Could not load company key(s). Encrypt anyway?</li>`;
    return;
  }

  for (const ck of _companyKeys) {
    const info = await getKeyInfo(ck.key.armor());
    const li = document.createElement('li');
    li.className = 'pgp-key-item';
    li.innerHTML = `
      <div class="pgp-key-item__email">${escHtml(ck.email)}</div>
      <span class="pgp-fingerprint">${escHtml(info.fingerprintFormatted)}</span>`;
    list.appendChild(li);
  }
}

// ── Attachments ───────────────────────────────────────────────────────────────

function loadAttachments() {
  const item    = Office.context.mailbox.item;
  const list    = el('attachment-list');
  const empty   = el('attachments-empty');
  const loading = el('attachments-loading');
  const note    = el('attach-capability-note');

  // On Mailbox < 1.8 clients, swap the description to explain the limitation.
  if (!_has18 && note) {
    note.textContent =
      '⚠ Your version of Outlook does not support attachment encryption. ' +
      'Attachments must be removed before this message can be encrypted. ' +
      'Upgrade to Outlook 2021 or Microsoft 365 for full attachment support.';
    note.style.color = '#797775';
  }

  loading.classList.remove('pgp-hidden');
  empty.classList.add('pgp-hidden');

  // item.attachments is only updated for attachments added programmatically
  // in the current task-pane session.  getAttachmentsAsync() returns the full
  // list including any files the user attached before opening the pane.
  return new Promise((resolve) => {
    item.getAttachmentsAsync({}, (result) => {
      loading.classList.add('pgp-hidden');

      const raw = result.status === Office.AsyncResultStatus.Succeeded
        ? result.value
        : (item.attachments || []);   // graceful fallback for older hosts

      _attachments = raw.filter(a => !a.isInline);
      _inlineAttachments = raw.filter(a => a.isInline);

      if (_attachments.length === 0) {
        empty.classList.remove('pgp-hidden');
        resolve();
        return;
      }

      empty.classList.add('pgp-hidden');

      // Remove only the dynamically-added attachment items, leaving the
      // static #attachments-empty <li> in the DOM so subsequent calls
      // to loadAttachments() can find it via el('attachments-empty').
      Array.from(list.children).forEach(c => {
        if (c.id !== 'attachments-empty') c.remove();
      });

      _attachments.forEach(att => {
        const li = document.createElement('li');
        li.className = 'pgp-attachment-item';
        li.innerHTML = `
          <span class="pgp-attachment-item__name" title="${escHtml(att.name)}">${escHtml(att.name)}</span>
          <span class="pgp-badge pgp-badge--info pgp-badge--sm">→ ${escHtml(att.name)}.pgp</span>`;
        list.appendChild(li);
      });

      resolve();
    });
  });
}

// ── Encrypt button state ──────────────────────────────────────────────────────

function updateEncryptButton() {
  const allHaveKeys = _recipientResults.length > 0 &&
    _recipientResults.every(r => !!r.key);
  const ready = allHaveKeys && hasKeyPair();
  const wasDisabled = el('btn-encrypt').disabled;
  el('btn-encrypt').disabled = !ready;
  if (ready && wasDisabled) el('btn-encrypt').focus(); // only on disabled→enabled transition
}

// ── Passphrase modal ──────────────────────────────────────────────────────────

function promptPassphrase() {
  return new Promise((resolve, reject) => {
    const modal = el('passphrase-modal');
    const input = el('passphrase-input');
    const errEl = el('passphrase-error');

    input.value = '';
    errEl.classList.add('pgp-hidden');
    modal.style.display = 'flex';
    modal.classList.remove('pgp-hidden');
    input.focus();

    function cleanup() {
      modal.style.display = '';
      modal.classList.add('pgp-hidden');
      input.removeEventListener('keydown', onKeydown);
      el('btn-passphrase-ok').removeEventListener('click', onOk);
      el('btn-passphrase-cancel').removeEventListener('click', onCancel);
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

    function onCancel() {
      cleanup();
      reject(new Error('Cancelled by user.'));
    }

    function onKeydown(e) {
      if (e.key === 'Enter') onOk();
      if (e.key === 'Escape') onCancel();
    }

    el('btn-passphrase-ok').addEventListener('click', onOk);
    el('btn-passphrase-cancel').addEventListener('click', onCancel);
    input.addEventListener('keydown', onKeydown);
  });
}

// ── Core encrypt flow ─────────────────────────────────────────────────────────

async function handleEncrypt() {
  clearStatus();
  const btn = el('btn-encrypt');
  const spinner = el('encrypt-spinner');
  btn.disabled = true;
  spinner.classList.remove('pgp-hidden');

  try {
    // 0a. Re-check recipient resolution before doing anything else. If the
    //     compose window was just opened (e.g. via "Reply encrypted") or the
    //     user finished typing a recipient moments ago, Outlook may still be
    //     resolving the To/Cc fields — loadRecipients() re-polls and
    //     re-resolves keys so we don't encrypt against a stale/incomplete
    //     recipient list.
    showStatus('Checking recipients…', 'info');
    await loadRecipients();
    if (!_recipientResults.every(r => !!r.key)) {
      throw new Error('Not all recipients have a resolved key yet — review the recipient list and try again.');
    }

    // 0b. Refresh the attachment list in case attachments were added after the
    //    pane was first opened.  Must be awaited so _attachments is current
    //    before the encryption loop runs.
    await loadAttachments();

    // 1. Unlock the private key — only needed when signing is enabled.
    //    Encrypting to our own public key (step 2) does NOT require the
    //    private key; the public key alone is sufficient for encryption.
    const shouldSign = el('sign-toggle').checked;
    let signingKey = null;

    if (shouldSign) {
      // Check the session cache before prompting — the user may have already
      // entered their passphrase during this task pane session.
      signingKey = getSessionKey();

      if (!signingKey) {
        const passphrase = await promptPassphrase();
        signingKey = await unlockPrivateKey(getPrivateKey(), passphrase);

        // Cache the unlocked key for the remainder of the session (15-minute
        // inactivity timeout; cleared automatically when the pane is closed).
        const userEmail = Office.context.mailbox.userProfile?.emailAddress || '';
        const keyInfo   = await getKeyInfo(getPublicKey());
        cacheSessionKey(signingKey, userEmail, keyInfo.shortId);
        updateSessionStatus();
      }
    }

    // 2. Collect all encryption keys
    //    — own public key (encrypt to self so you can read sent mail)
    //    — all recipient keys
    //    — company keys if enabled
    const ownPublicKey = await readPublicKey(getPublicKey());
    const recipientKeys = _recipientResults.map(r => r.key).filter(Boolean);

    const includeCompanyKey = isCompanyKeyEnabled() && el('company-key-toggle').checked;
    const companyKeyObjects = includeCompanyKey ? _companyKeys.map(ck => ck.key) : [];

    const allEncryptionKeys = [ownPublicKey, ...recipientKeys, ...companyKeyObjects];

    // 2b. Warn (but do not block) if any recipient uses a legacy key algorithm
    //     such as ElGamal. Encryption will still succeed via an automatic retry
    //     with a permissive config inside encryptMessage / encryptAttachment.
    if (await hasWeakEncryptionKey(recipientKeys)) {
      showStatus(
        '⚠ One or more recipients use a legacy key algorithm (e.g. ElGamal/DSA). ' +
        'Encryption will proceed, but their key offers reduced security compared to modern ECC or RSA-2048+ keys.',
        'warning'
      );
      // Brief pause so the user can read the warning before it is replaced by
      // the progress message below.
      await new Promise(r => setTimeout(r, 2000));
    }

    // 3. Get the message body as HTML so that all formatting, inline images, and
    //    rich-text markup are preserved exactly.  We encrypt the raw HTML string
    //    as the PGP payload.  The recipient's decrypt pane will detect that the
    //    decrypted content is HTML and render it in a sandboxed <iframe>.
    showStatus('Encrypting message body…', 'info');
    let bodyHtml = await getBodyAsync(Office.CoercionType.Html);

    // Refuse to double-encrypt.  When the body is HTML the PGP armor block will
    // appear as literal text inside the <body> element if already encrypted.
    if (detectPgpContent(bodyHtml) === 'encrypted') {
      showStatus('Message appears to already be PGP-encrypted.', 'warning');
      btn.disabled = false;
      spinner.classList.add('pgp-hidden');
      return;
    }

    // Warn if the message body contains inline images (e.g. embedded images).
    // These are incompatible with PGP encryption — the cid: URIs cannot be
    // resolved after the body is replaced with armor text.
    // reconcileInlineAttachments() supplements the API's isInline flag with a
    // direct body-HTML scan, because some Outlook environments (e.g. OWA) set
    // isInline=false for user-pasted images, or omit them from the API list.
    reconcileInlineAttachments(bodyHtml);
    if (_inlineAttachments.length > 0) {
      const choice = await confirmInlineAttachments();
      if (!choice) throw new Error('Cancelled by user.');
      if (choice === 'convert') {
        showStatus('Converting inline attachments to regular attachments…', 'info');
        const { cleaned, converted } = await convertInlineAttachments(bodyHtml);
        bodyHtml = cleaned;
        if (converted === 0) {
          // Outlook does not expose clipboard-pasted images through the
          // attachment API, so we can strip the cid: reference but cannot
          // read and re-attach the data.  The body has been cleaned; warn the
          // user so they can re-attach the image manually.
          showStatus(
            '⚠ The inline image(s) could not be accessed via the Outlook API ' +
            'and have been removed from the message body. To include the ' +
            'image(s), save each to disk and re-attach as a regular file.',
            'warning'
          );
          await new Promise(r => setTimeout(r, 4000));
        }
      }
    }

    const encryptedBody = await encryptMessage(bodyHtml, allEncryptionKeys, signingKey);

    // The outer body is plain-text PGP armor — recipients without the add-in
    // will see the raw armor; those with it will decrypt and render the HTML.
    await setBodyAsync(encryptedBody);

    // 4. Encrypt attachments
    if (_attachments.length > 0) {
      if (!_has18) {
        // Mailbox < 1.8: attachment APIs unavailable. Require explicit user
        // consent before removing attachments — never silently destroy data.
        const confirmed = await confirmAttachmentRemoval(_attachments.length);
        if (!confirmed) throw new Error('Cancelled by user.');
        await removeAllAttachments();
      } else {
        showStatus('Encrypting attachments…', 'info');
        await encryptAttachments(allEncryptionKeys, signingKey);
      }
    }

    showStatus('✓ Message encrypted. Click Send when ready.', 'success');

  } catch (e) {
    if (e.message === 'Cancelled by user.') {
      showStatus('Encryption cancelled.', 'info');
    } else {
      showStatus(`Encryption failed: ${e.message}`, 'error');
      console.error(e);
    }
  } finally {
    // Re-enable if there was a non-passphrase error
    const encrypted = el('status-bar').classList.contains('pgp-alert--success');
    btn.disabled = encrypted; // keep disabled after success so user can't re-encrypt
    spinner.classList.add('pgp-hidden');
  }
}

/**
 * Warn the user that the message contains inline attachments which are
 * incompatible with PGP encryption, then let them choose what to do.
 * Resolves to:
 *   'convert'  – move inline attachments to regular attachments, then encrypt
 *   'continue' – encrypt as-is (inline images will break for the recipient)
 *   false      – abort
 */
function confirmInlineAttachments() {
  return new Promise((resolve) => {
    // Show the Convert option only on Outlook on the web, where
    // getAttachmentsAsync() exposes pasted inline images via the API.
    // On desktop builds those images are inaccessible, so we hide the button
    // and show a simpler "fix manually" hint instead.
    el('btn-cid-convert').classList.toggle('pgp-hidden', !_isWebOutlook);
    el('cid-hint-web').classList.toggle('pgp-hidden', !_isWebOutlook);
    el('cid-hint-desktop').classList.toggle('pgp-hidden', _isWebOutlook);

    const modal = el('cid-warning-modal');
    modal.style.display = 'flex';
    modal.classList.remove('pgp-hidden');

    function cleanup() {
      modal.style.display = '';
      modal.classList.add('pgp-hidden');
      el('btn-cid-convert').removeEventListener('click', onConvert);
      el('btn-cid-continue').removeEventListener('click', onContinue);
      el('btn-cid-cancel').removeEventListener('click', onCancel);
    }
    function onConvert()  { cleanup(); resolve('convert'); }
    function onContinue() { cleanup(); resolve('continue'); }
    function onCancel()   { cleanup(); resolve(false); }

    el('btn-cid-convert').addEventListener('click', onConvert);
    el('btn-cid-continue').addEventListener('click', onContinue);
    el('btn-cid-cancel').addEventListener('click', onCancel);
  });
}

/**
 * Supplement the API's isInline flag with a direct scan of the body HTML.
 *
 * Some Outlook environments (notably OWA) report user-pasted images with
 * isInline=false, or don't include them in getAttachmentsAsync() at all, even
 * though the body HTML references them via <img src="cid:…">.
 *
 * This function:
 *   1. Extracts every CID value from <img src="cid:…"> tags in the body.
 *   2. Moves any _attachments entry whose id matches a found CID into
 *      _inlineAttachments (reclassification of false-negative API results).
 *   3. For CIDs that have no matching attachment at all (orphaned), pushes a
 *      sentinel object into _inlineAttachments so the warning still fires and
 *      the img tag is stripped from the body during conversion.
 *
 * Must be called after both loadAttachments() and getBodyAsync() have settled.
 *
 * @param {string} bodyHtml  Current HTML body of the message.
 */
function reconcileInlineAttachments(bodyHtml) {
  const cidRefs = new Set(
    [...bodyHtml.matchAll(/<img\b[^>]*\bsrc=["']cid:([^"']+)["']/gi)].map(m => m[1])
  );
  if (cidRefs.size === 0) return;

  // 1. Reclassify regular attachments whose id matches a body CID.
  //    Single partition pass — avoids scanning _attachments twice.
  const keep = [], reclassified = [];
  for (const a of _attachments) {
    (cidRefs.has(a.id) ? reclassified : keep).push(a);
  }
  if (reclassified.length > 0) {
    _attachments       = keep;
    _inlineAttachments = [..._inlineAttachments, ...reclassified];
  }

  // 2. For CIDs that have no matching attachment by exact id, try matching by
  //    name prefix.  Outlook CIDs are typically "filename.ext@domain-part", so
  //    the part before the first "@" often matches the attachment's name field.
  //    This covers the common OWA case where the API returns isInline=false and
  //    assigns a short numeric id (e.g. "1") that doesn't match the body CID.
  const knownIds = new Set(_inlineAttachments.map(a => a.id));
  for (const cid of cidRefs) {
    if (knownIds.has(cid)) continue;

    const namePrefix  = cid.split('@')[0];
    const byName      = _attachments.find(a => a.name === namePrefix);

    if (byName) {
      // Real attachment found via name — move it to the inline list.
      _attachments       = _attachments.filter(a => a.id !== byName.id);
      _inlineAttachments = [..._inlineAttachments, byName];
      knownIds.add(byName.id);
    } else {
      // True orphan: CID appears in the body but no corresponding attachment
      // is accessible via the API (common for clipboard-pasted images in some
      // Outlook builds).  A sentinel ensures the warning fires and the <img>
      // tag is stripped from the body; read/remove/re-add is skipped for it.
      _inlineAttachments.push({ id: cid, name: namePrefix || cid, contentType: '', size: 0, isInline: true });
      knownIds.add(cid);
    }
  }
}

/**
 * Convert all inline attachments (isInline) to regular file attachments.
 *
 * For each inline attachment:
 *   1. Read its content via the Office API.
 *   2. Remove the inline attachment from the message.
 *   3. Re-add it as a regular (non-inline) file attachment.
 *
 * Also strips <img src="cid:…"> tags from the supplied body HTML (since the
 * cid: URIs will no longer resolve once the inline attachments are gone) and
 * persists the cleaned HTML back to the message body.
 *
 * Returns the cleaned body HTML so the caller can use it for encryption
 * without a redundant round-trip to the Office API.
 *
 * @param {string} bodyHtml  Current HTML body of the message.
 * @returns {Promise<string>} Cleaned HTML body (cid: image tags removed).
 */
async function convertInlineAttachments(bodyHtml) {
  const item = Office.context.mailbox.item;

  let converted = 0;

  for (const att of _inlineAttachments) {
    let contentResult;
    try {
      contentResult = await getAttachmentContentAsync(item, att.id);
    } catch (_) {
      // The attachment is not accessible via the Office API (e.g. a pasted
      // clipboard image that Outlook doesn't expose through getAttachmentsAsync).
      // Skip read/remove/re-add; the cid: img tag is still stripped below.
      continue;
    }
    await removeAttachmentAsync(item, att.id);
    await addAttachmentFromBase64Async(item, contentResult.content, att.name);
    converted++;
  }

  // Strip every <img> whose src is a cid: URI — those images are gone from
  // the body now that they have been promoted to regular file attachments.
  const cleaned = bodyHtml.replace(/<img\b[^>]*\bsrc=["']cid:[^"']*["'][^>]*\/?>/gi, '');

  await setBodyHtmlAsync(cleaned);

  _inlineAttachments = [];
  await loadAttachments();

  return { cleaned, converted };
}

/** Set the message body as HTML without any PGP-armor wrapping. */
function setBodyHtmlAsync(html) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.setAsync(
      html,
      { coercionType: Office.CoercionType.Html },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
        else reject(new Error(result.error.message));
      }
    );
  });
}

function getBodyAsync(coercionType = Office.CoercionType.Text) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.getAsync(
      coercionType,
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
        else reject(new Error(result.error.message));
      }
    );
  });
}

function setBodyAsync(armoredText) {
  // Wrap the PGP armor in a <pre> block so Outlook preserves line breaks.
  // Setting CoercionType.Text in an HTML-mode compose window causes Outlook
  // to wrap lines in <p> tags (collapsing newlines), which corrupts the armor
  // structure and makes it undetectable when the recipient opens the message.
  const safe = armoredText
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
  const html = `<html><body><pre style="font-family:monospace;white-space:pre-wrap;">${safe}</pre></body></html>`;
  return setBodyHtmlAsync(html);
}

// ── Reply handoff (native-reply armor splice) ─────────────────────────────────
//
// For large decrypted messages, MessageRead.js's Reply/Reply All opens
// Outlook's NATIVE reply (displayReplyForm/displayReplyAllForm, no custom
// body) instead of building one from scratch — that quotes the ORIGINAL
// message as Outlook has it: still PGP-armored, since Outlook has no notion
// of decryption. MessageRead.js then hands the decrypted plaintext to this
// window over a BroadcastChannel (there's no way for it to pass a
// per-instance token into this window's URL, unlike the pop-out dialog),
// because Body.setAsync (1 MB limit) — called from inside this compose
// window's own script context — is the only body-write path not bound by
// the 32 KB htmlBody cap shared by displayNewMessageForm and
// displayReplyForm/displayReplyAllForm.
//
// BroadcastChannel is same-origin-wide, not scoped to just these two
// windows — any script running anywhere in the add-in's origin could listen
// in on decrypted plaintext if this were unconditional and indefinite. Two
// mitigations keep that blast radius narrow:
//   - The listener is only set up when this compose window is confirmed to
//     be a reply (getComposeTypeAsync, Mailbox 1.10+) — an ordinary new
//     message or forward never listens at all. Falls back to listening
//     unconditionally on older hosts that can't be asked (see setupReplyHandoffListener).
//   - Once set up, it only listens for REPLY_HANDOFF_LISTEN_TIMEOUT_MS — a
//     reply window that was never the target of an actual handoff (e.g. the
//     user used Outlook's own Reply button on an encrypted-but-undecrypted
//     message) stops listening after a short grace period rather than for
//     the rest of the window's lifetime.
// The channel name itself is also derived from the conversation ID
// (getReplyHandoffChannelName) rather than fixed, so a listener needs to
// already know which conversation to target — see that module's docblock
// for why this isn't a secrecy boundary on its own.

// Matches (with margin) MessageRead.js's REPLY_HANDOFF_TIMEOUT_MS, so a
// reply window that's genuinely waiting on a handoff isn't cut off before
// the sender itself gives up and falls back.
const REPLY_HANDOFF_LISTEN_TIMEOUT_MS = 12000;

let _replyHandoffConsumed = false;

/**
 * Called from Office.onReady with no arguments in production (`has110`/
 * `has114` default to the real feature-detected _has110/_has114). Exported,
 * and both made explicit parameters rather than only reading module state,
 * so tests can exercise every branch deterministically without needing
 * Office.onReady itself to run.
 *
 * @param {boolean} [has110]
 * @param {boolean} [has114]
 */
export async function setupReplyHandoffListener(has110 = _has110, has114 = _has114) {
  if (typeof BroadcastChannel !== 'function') return;

  if (has110) {
    // Office.MailboxEnums.ComposeType has exactly three values -- Reply,
    // NewMail, Forward -- Reply All is NOT a distinct value; getComposeTypeAsync
    // reports 'reply' for both.
    const composeType = await getComposeTypeAsync().catch((e) => {
      console.error('Reply handoff: getComposeTypeAsync failed', e);
      return null; // unknown -- fall through and listen anyway, see below
    });
    if (composeType !== null && composeType !== Office.MailboxEnums.ComposeType.Reply) return;
  }
  // _has110 false: can't confirm compose type, so listen anyway (broader
  // exposure on older hosts only, still bounded by the timeout below).

  // Prefer conversationId, but fall back to inReplyTo (Mailbox 1.14 -- the
  // internet message ID of the message being replied to) when it's missing.
  // MessageRead.js derives the matching scoping ID the same way, from its
  // own item.conversationId / item.internetMessageId. See
  // openNativeReplyWithHandoff.
  const conversationId = Office.context.mailbox.item.conversationId;
  const inReplyTo = has114 ? Office.context.mailbox.item.inReplyTo : undefined;
  const scopingId = conversationId || inReplyTo;
  if (!scopingId) {
    // No way to scope the channel to this specific conversation/message at
    // all -- falling back to the shared base channel name would let any
    // same-origin page listen for this (and every other) large reply's
    // decrypted plaintext. Treat this the same as "handoff unavailable" and
    // don't listen at all; MessageRead.js makes the matching decision on its
    // side (see openNativeReplyWithHandoff).
    return;
  }

  const channelName = getReplyHandoffChannelName(scopingId);
  let channel;
  try {
    channel = new BroadcastChannel(channelName);
  } catch (e) {
    console.error('Reply handoff: BroadcastChannel construction failed', e);
    return;
  }

  const idleTimer = setTimeout(() => {
    if (!_replyHandoffConsumed) channel.close();
  }, REPLY_HANDOFF_LISTEN_TIMEOUT_MS);
  let handoffInFlight = false;

  channel.onmessage = async (event) => {
    const data = event.data;
    if (!data || data.type !== 'pgp-reply-handoff' || _replyHandoffConsumed || handoffInFlight) return;
    handoffInFlight = true;

    const success = await applyReplyHandoff(data.text, data.isHtml);
    // Only ack on confirmed success -- an ack that arrives despite a failed
    // splice would make MessageRead.js treat this as done and never trigger
    // its own fallback, leaving the user with a still-armored body and no
    // backup window. Staying silent here lets that timeout-based fallback
    // fire instead.
    if (success) {
      _replyHandoffConsumed = true;
      clearTimeout(idleTimer);
      channel.postMessage({ type: 'pgp-reply-handoff-ack', token: data.token });
      channel.close();
      return;
    }
    handoffInFlight = false;
  };
}

function getComposeTypeAsync() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.getComposeTypeAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value.composeType);
      else reject(new Error(result.error.message));
    });
  });
}

/**
 * Reads the current (Outlook-native-quoted) body, strips out the PGP armor
 * block Outlook quoted from the original encrypted message, and splices the
 * decrypted plaintext in at that location.
 *
 * Leaves the body untouched (with a warning) if the armor block can't be
 * found, or if anything else fails — never guesses at a partial edit.
 *
 * @param {string} text - Decrypted payload from MessageRead.js
 * @param {boolean} isHtml - True when the payload is HTML
 * @returns {Promise<boolean>} true only if the splice actually succeeded
 */
async function applyReplyHandoff(text, isHtml) {
  try {
    const bodyHtml = await getBodyAsync(Office.CoercionType.Html);
    const { found, before, after } = stripPgpArmorBlock(bodyHtml);
    if (!found) {
      showStatus('Could not find the encrypted message in this reply to replace — please verify the body before sending.', 'warning');
      return false;
    }
    const formattedContent = formatDecryptedContentAsHtml(text, isHtml);
    await setBodyHtmlAsync(before + formattedContent + after);
    showStatus('Decrypted message inserted into this reply.', 'success');
    return true;
  } catch (e) {
    showStatus(`Could not automatically insert the decrypted message into this reply: ${e.message} — please verify the body before sending.`, 'warning');
    return false;
  }
}

// Internal splitting delimiter for stripPgpArmorBlock() — a Private Use Area
// character sequence that can never appear in real email text, and survives
// HTML (de)serialization unescaped (no &, <, >, or ").
const ARMOR_SPLICE_MARKER_BASE = '__PGP_ARMOR_SPLICE__';

/**
 * Returns a marker guaranteed not to already appear in `html` — the input is
 * attacker-influenceable PGP message content, so the base marker alone can't
 * be assumed unique, however unlikely a literal collision is in practice.
 */
function pickSpliceMarker(html) {
  let marker = ARMOR_SPLICE_MARKER_BASE;
  for (let i = 0; html.includes(marker); i++) {
    marker = `${ARMOR_SPLICE_MARKER_BASE}${i}`;
  }
  return marker;
}

/**
 * Locates the PGP armor block (`-----BEGIN PGP MESSAGE-----` through
 * `-----END PGP MESSAGE-----`, inclusive) inside an HTML body string and
 * splits the HTML around it, so the caller can splice something else in at
 * that location.
 *
 * The armor block may be split across multiple sibling text nodes/lines the
 * way Outlook's native reply-quoting renders a quoted message (including a
 * `<pre>`-wrapped block — this add-in's own setBodyAsync() sends the armor
 * that way, so a reply to one of its own messages commonly quotes it back
 * inside a `<pre>`). This walks a detached DOM the same way MessageRead.js's
 * extractArmorFromHtml() does (same BLOCK-element/`<br>`/`<pre>` handling),
 * but additionally tracks which text node each character came from, so the
 * located range can be mapped back to specific nodes and removed — rather
 * than only extracted, like extractArmorFromHtml() does.
 *
 * Removal itself works by replacing the range with a unique marker inside
 * the DOM, serializing once, then splitting the resulting HTML string on
 * that marker — this guarantees valid, unmangled surrounding HTML regardless
 * of how many nodes the range spans, without needing to reconstruct partial
 * DOM structure by hand.
 *
 * @param {string} html
 * @returns {{found: boolean, before?: string, after?: string}}
 */
export function stripPgpArmorBlock(html) {
  const div = document.createElement('div');
  div.innerHTML = html;

  const BLOCK = new Set([
    'div', 'p', 'li', 'blockquote', 'tr',
    'h1', 'h2', 'h3', 'h4', 'h5', 'h6',
    'article', 'section', 'header', 'footer', 'html', 'body',
  ]);
  const SKIP = new Set(['style', 'script', 'head', 'title', 'noscript']);

  // { node, start, end } — [start, end) is this node's exact character range
  // within `flat`, always equal to flat.slice(start, end) === node.textContent.
  const segments = [];
  let flat = '';

  function walk(node) {
    if (node.nodeType === Node.TEXT_NODE) {
      const start = flat.length;
      flat += node.textContent;
      segments.push({ node, start, end: flat.length });
      return;
    }
    if (node.nodeType !== Node.ELEMENT_NODE) return;
    const tag = node.tagName.toLowerCase();
    if (SKIP.has(tag)) return;
    if (tag === 'br') { flat += '\n'; return; }
    if (tag === 'pre') {
      flat += '\n';
      const start = flat.length;
      flat += node.textContent;
      segments.push({ node, start, end: flat.length });
      flat += '\n';
      return;
    }
    for (const child of Array.from(node.childNodes)) walk(child);
    if (BLOCK.has(tag)) flat += '\n';
  }
  walk(div);

  const beginIdx = flat.indexOf('-----BEGIN PGP MESSAGE-----');
  if (beginIdx === -1) return { found: false };
  const endMarkerIdx = flat.indexOf('-----END PGP MESSAGE-----', beginIdx);
  if (endMarkerIdx === -1) return { found: false };
  const endIdx = endMarkerIdx + '-----END PGP MESSAGE-----'.length;

  const marker = pickSpliceMarker(html);
  let markerPlaced = false;
  for (const seg of segments) {
    if (seg.end <= beginIdx || seg.start >= endIdx) continue; // no overlap

    const localBegin = Math.max(0, beginIdx - seg.start);
    const localEnd = Math.min(seg.end - seg.start, endIdx - seg.start);
    const text = seg.node.textContent;
    seg.node.textContent = text.slice(0, localBegin) + (markerPlaced ? '' : marker) + text.slice(localEnd);
    markerPlaced = true;
  }

  const spliced = div.innerHTML;
  const parts = spliced.split(marker);
  // Anything other than exactly 2 parts means the marker either never made
  // it into the serialized output, or (despite pickSpliceMarker's check
  // against the raw input) something produced more copies of it than
  // expected -- either way, splicing on an assumption that doesn't hold
  // would corrupt the reply body, so bail out the same safe way as "not
  // found" rather than guess.
  if (parts.length !== 2) return { found: false };
  const [before, after] = parts;
  return { found: true, before, after };
}

/**
 * Show a blocking modal asking the user to confirm that their attachments will
 * be permanently removed before body-only encryption proceeds.
 * Returns a Promise<boolean> — true if the user clicked confirm, false if cancelled.
 */
function confirmAttachmentRemoval(count) {
  return new Promise((resolve) => {
    const countEl = el('attach-remove-count');
    countEl.textContent = `${count} attachment${count === 1 ? '' : 's'}`;

    const modal = el('attach-remove-modal');
    modal.style.display = 'flex';
    modal.classList.remove('pgp-hidden');

    function cleanup() {
      modal.style.display = '';
      modal.classList.add('pgp-hidden');
      el('btn-attach-remove-confirm').removeEventListener('click', onConfirm);
      el('btn-attach-remove-cancel').removeEventListener('click', onCancel);
    }
    function onConfirm() { cleanup(); resolve(true); }
    function onCancel()  { cleanup(); resolve(false); }

    el('btn-attach-remove-confirm').addEventListener('click', onConfirm);
    el('btn-attach-remove-cancel').addEventListener('click', onCancel);
  });
}

/**
 * Remove every non-inline attachment from the item and reset local state.
 * Called only on Mailbox < 1.8 after the user has explicitly confirmed removal.
 * Reuses the existing removeAttachmentAsync wrapper — no new Office API surface.
 */
async function removeAllAttachments() {
  const item = Office.context.mailbox.item;
  for (const att of _attachments) {
    await removeAttachmentAsync(item, att.id);
  }
  _attachments = [];
  _inlineAttachments = [];
  await loadAttachments();
}

async function encryptAttachments(encryptionKeys, signingKey) {
  const item = Office.context.mailbox.item;

  for (const att of _attachments) {
    // Read attachment content (requires Mailbox 1.8)
    const contentResult = await getAttachmentContentAsync(item, att.id);

    // getAttachmentContentAsync returns different formats depending on the
    // attachment type.  Regular files are Base64; dragged email items are Eml
    // (raw MIME text); calendar items are ICalendar (raw iCal text).
    // Cloud/URL attachments cannot be read as bytes through this API.
    let rawBytes;
    let encryptedName;
    const fmt = contentResult.format;
    if (fmt === Office.MailboxEnums.AttachmentContentFormat.Base64) {
      rawBytes = base64ToUint8Array(contentResult.content.replace(/[^\x00-\x7F]/g, ''));
      encryptedName = att.name + '.pgp';
    } else if (fmt === Office.MailboxEnums.AttachmentContentFormat.Eml) {
      rawBytes = new TextEncoder().encode(contentResult.content);
      // Ensure the decrypted file opens as .eml (email items have no extension in Outlook)
      const baseName = att.name.toLowerCase().endsWith('.eml') ? att.name : att.name + '.eml';
      encryptedName = baseName + '.pgp';
    } else if (fmt === Office.MailboxEnums.AttachmentContentFormat.ICalendar) {
      rawBytes = new TextEncoder().encode(contentResult.content);
      const baseName = att.name.toLowerCase().endsWith('.ics') ? att.name : att.name + '.ics';
      encryptedName = baseName + '.pgp';
    } else {
      throw new Error(
        `Cannot encrypt "${att.name}": it is a cloud/linked attachment. ` +
        `Download it to your device and re-attach the file before encrypting.`
      );
    }

    // Encrypt — use the corrected base name (e.g. "subject.eml") as the
    // filename stored in the PGP literal data packet so the decryptor
    // recovers the right extension, not a browser-guessed one.
    const plainName = encryptedName.replace(/\.pgp$/i, '');
    const armoredEncrypted = await encryptAttachment(
      rawBytes,
      plainName,
      encryptionKeys,
      signingKey
    );

    // Remove the original
    await removeAttachmentAsync(item, att.id);

    // Add the encrypted version
    const encryptedBase64 = btoa(armoredEncrypted);
    await addAttachmentFromBase64Async(item, encryptedBase64, encryptedName);
  }

  // Refresh attachment list display
  _attachments = [];
  _inlineAttachments = [];
  await loadAttachments();
}

function getAttachmentContentAsync(item, attachmentId) {
  return new Promise((resolve, reject) => {
    item.getAttachmentContentAsync(attachmentId, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
      else reject(new Error(result.error.message));
    });
  });
}

function removeAttachmentAsync(item, attachmentId) {
  return new Promise((resolve, reject) => {
    item.removeAttachmentAsync(attachmentId, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(new Error(result.error.message));
    });
  });
}

function addAttachmentFromBase64Async(item, base64, name) {
  return new Promise((resolve, reject) => {
    item.addFileAttachmentFromBase64Async(base64, name, { asyncContext: null }, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
      else reject(new Error(result.error.message));
    });
  });
}

// ── Delegate recipient list interactions ──────────────────────────────────────

/**
 * Wire a single delegated click handler on the recipient list container.
 * Must be called exactly once after the container is in the DOM.
 *
 * Using event delegation means we can call renderRecipientList() to replace
 * the inner HTML without needing to re-attach listeners — the handler always
 * lives on the stable container element, not on individual buttons.
 */
function wireRecipientListEvents() {
  el('recipient-list').addEventListener('click', async (e) => {
    const idx = parseInt(e.target.closest('[data-idx]')?.dataset.idx ?? '-1');
    if (idx < 0) return;

    // Retry key lookup — re-runs the full WKD/VKS discovery chain
    if (e.target.classList.contains('btn-retry-key')) {
      e.target.disabled = true;
      e.target.textContent = '…';
      const result = await resolveRecipients([_recipientResults[idx].email]);
      _recipientResults[idx] = result[0];
      renderRecipientList(); // replaces innerHTML; delegation keeps the handler alive
      updateEncryptButton();
    }

    // Toggle the inline paste form for a recipient
    if (e.target.classList.contains('btn-paste-key')) {
      el(`recipient-paste-form-${idx}`).classList.toggle('pgp-hidden');
    }

    // Cancel paste — hide the form
    if (e.target.classList.contains('btn-paste-key-cancel')) {
      el(`recipient-paste-form-${idx}`).classList.add('pgp-hidden');
    }

    // Validate and accept a manually pasted armored public key
    if (e.target.classList.contains('btn-paste-key-confirm')) {
      const armoredKey = el(`recipient-paste-key-${idx}`).value.trim();
      if (!armoredKey) return;
      try {
        const key = await readPublicKey(armoredKey);
        _recipientResults[idx].key = key;
        _recipientResults[idx].armoredKey = armoredKey;
        _recipientResults[idx].status = 'found_local';
        _recipientResults[idx].source = 'Pasted';
        renderRecipientList();
        updateEncryptButton();
      } catch (err) {
        alert(`Invalid PGP key: ${err.message}`);
      }
    }

  });
}

// ── Bootstrap ─────────────────────────────────────────────────────────────────

Office.onReady(async () => {
  const userEmail = Office.context.mailbox.userProfile?.emailAddress || '';
  _isWebOutlook = Office.context.platform === Office.PlatformType.OfficeOnline;
  _has18 = Office.context.requirements.isSetSupported('Mailbox', '1.8');
  _has110 = Office.context.requirements.isSetSupported('Mailbox', '1.10');
  _has114 = Office.context.requirements.isSetSupported('Mailbox', '1.14');

  // Fire-and-forget: inert for every ordinary compose window unless a
  // matching reply-handoff broadcast actually arrives (see its own docblock).
  setupReplyHandoffListener();

  // Load org config
  await loadOrgConfig(userEmail);

  // Check for own key pair
  if (!hasKeyPair()) {
    el('panel-no-key').classList.remove('pgp-hidden');
    el('btn-encrypt').disabled = true;
  }

  // Load data in parallel
  await Promise.all([
    loadRecipients(),
    loadCompanyKeys(),
  ]);
  loadAttachments();

  // Apply the user's stored sign-by-default preference.
  // The user can flip the toggle for any individual message.
  el('sign-toggle').checked = getSignDefault();

  // Reflect initial session cache state (user may have just come from KeyManagement)
  updateSessionStatus();

  // Keep the session status bar in sync whenever the cache is cleared (timeout
  // or the user clicking Lock).  onSessionCleared fires for both.
  onSessionCleared(updateSessionStatus);

  // Wire events
  el('btn-refresh-recipients').addEventListener('click', async () => {
    _recipientResults = [];
    await loadRecipients();
    // No need to call wireRecipientListEvents() again — it uses event
    // delegation on the container, which survives innerHTML replacement.
  });

  el('btn-encrypt').addEventListener('click', handleEncrypt);

  el('btn-lock-session').addEventListener('click', () => {
    clearSessionKey(); // triggers onSessionCleared → updateSessionStatus
  });

  // Wire the delegated recipient list handler exactly once.
  wireRecipientListEvents();
});
