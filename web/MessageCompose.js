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
 * Requires: Mailbox 1.8 (for getAttachmentContentAsync / addFileAttachmentFromBase64Async)
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

function getRecipientsAsync(recipients) {
  return new Promise((resolve) => {
    recipients.getAsync((result) => {
      resolve(result.status === Office.AsyncResultStatus.Succeeded ? result.value : []);
    });
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
  el('btn-encrypt').disabled = !allHaveKeys || !hasKeyPair();
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
    // 0. Refresh the attachment list in case attachments were added after the
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
      showStatus('Encrypting attachments…', 'info');
      await encryptAttachments(allEncryptionKeys, signingKey);
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

async function encryptAttachments(encryptionKeys, signingKey) {
  const item = Office.context.mailbox.item;

  for (const att of _attachments) {
    // Read attachment content (requires Mailbox 1.8)
    const contentResult = await getAttachmentContentAsync(item, att.id);
    const rawBytes = base64ToUint8Array(contentResult.content);

    // Encrypt
    const armoredEncrypted = await encryptAttachment(
      rawBytes,
      att.name,
      encryptionKeys,
      signingKey
    );

    // Remove the original
    await removeAttachmentAsync(item, att.id);

    // Add the encrypted version
    const encryptedBase64 = btoa(armoredEncrypted);
    await addAttachmentFromBase64Async(item, encryptedBase64, att.name + '.pgp');
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
