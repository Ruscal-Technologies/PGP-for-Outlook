'use strict';
/**
 * KeyManagement.js
 * Task pane for managing the user's PGP identity and keyring.
 *
 * Sections:
 *
 *  MY KEY PAIR
 *    Generate a new key pair protected by a passphrase (ECC or RSA-4096), or
 *    import an existing key from any OpenPGP-compatible client (GnuPG, Kleopatra,
 *    Thunderbird, etc.).  The passphrase is verified against the imported key
 *    before saving; only the already-encrypted private key blob is persisted.
 *    Keys are stored in Office roaming settings — they follow the user across
 *    devices via their Microsoft 365 account, but never leave the Office
 *    ecosystem unencrypted.  The user can copy or email their public key to
 *    contacts, download a backup, and delete or replace the key pair at any time.
 *
 *  CONTACTS' KEYRING
 *    A local store of trusted contacts' public keys, also in roaming settings.
 *    Keys can be added by searching (WKD → VKS auto-discovery) or by pasting
 *    armored text directly.  A storage-usage warning appears when the 32 KB
 *    roaming settings ceiling is within 20% of being reached.
 *
 *  ORGANIZATION SETTINGS
 *    The add-in tries to load org-level config from (in order):
 *      Primary:  https://<user-email-domain>/.well-known/pgp-for-outlook-addin/company-config.json
 *      Fallback: https://openpgpkey.<user-email-domain>/.well-known/pgp-for-outlook-addin/company-config.json
 *    IT admins can enable/configure the company key feature by publishing that
 *    file.  A manual override (stored in roaming settings) takes precedence
 *    and is intended for orgs that cannot host a well-known file.
 *
 *  PERSONAL PREFERENCES
 *    User-level compose defaults stored in roaming settings.  Currently:
 *      - pgp_sign_default: whether the sign toggle starts checked in the
 *        compose pane.  Default is false; the user can override per-message.
 */

import { generateKeyPair, getKeyInfo, extractPublicKey, unlockPrivateKey, addModernSubkeys, hasModernSubkeys } from './js/pgp/pgp-core.js';
import {
  hasKeyPair, getPrivateKey, getPublicKey, getKeyMetadata,
  saveKeyPair, clearKeyPair,
  getOrgOverride, saveOrgOverride, clearOrgOverride,
  getSignDefault, saveSignDefault,
} from './js/pgp/key-storage.js';
import {
  addContactKey, removeContactKey, listContactKeys, getKeyringStorageInfo,
} from './js/pgp/keyring.js';
import { discoverKey, KeyStatus } from './js/pgp/key-discovery.js';
import {
  loadOrgConfig, getOrgConfig, isCompanyKeyEnabled, isCompanyKeyRequired,
  getCompanyKeyEmails, fetchCompanyKeys, isSupportButtonHidden,
} from './js/pgp/org-config.js';

// ── Helpers ───────────────────────────────────────────────────────────────────

function el(id) { return document.getElementById(id); }

function showStatus(containerId, message, type = 'info') {
  const container = el(containerId);
  container.className = `pgp-alert pgp-alert--${type}`;
  container.textContent = message;
  container.classList.remove('pgp-hidden');
}

function hideStatus(containerId) {
  el(containerId).classList.add('pgp-hidden');
}

function formatDate(date) {
  if (!date) return 'Never';
  return date.toLocaleDateString(undefined, { year: 'numeric', month: 'short', day: 'numeric' });
}

// ── My key pair panel ─────────────────────────────────────────────────────────

async function refreshMyKeyPanel() {
  if (!hasKeyPair()) {
    el('panel-no-key').classList.remove('pgp-hidden');
    el('panel-has-key').classList.add('pgp-hidden');
    return;
  }

  el('panel-no-key').classList.add('pgp-hidden');
  el('panel-has-key').classList.remove('pgp-hidden');

  const meta = getKeyMetadata();
  if (meta) {
    el('key-uid').textContent = `${meta.name} <${meta.email}>`;
    el('key-created').textContent = `Created: ${formatDate(meta.created ? new Date(meta.created) : null)}`;
    el('key-expires').textContent = meta.expires
      ? `Expires: ${formatDate(new Date(meta.expires))}`
      : 'No expiration';
    el('key-fingerprint').textContent = meta.fingerprintFormatted || meta.fingerprint;

    // Warn if expired
    if (meta.expires && new Date(meta.expires) < new Date()) {
      el('key-status-badge').className = 'pgp-badge pgp-badge--error';
      el('key-status-badge').textContent = 'Expired';
    }

    // Show algorithm badge and "Add Modern Subkeys" button for non-ECC primary key types.
    // Covers DSA/ElGamal (weak) and RSA (any key size — ECC is universally preferred).
    // OpenPGP.js reports RSA as 'rsaEncryptSign' / 'rsaEncrypt' / 'rsaSign' (camelCase),
    // so we match by prefix after lowercasing rather than enumerating variants.
    const primaryAlgo = (meta.algorithm || '').toLowerCase();
    const isNonEccAlgo = primaryAlgo.startsWith('rsa') ||
      primaryAlgo === 'dsa' || primaryAlgo === 'elgamal';

    if (isNonEccAlgo) {
      const badgeLabel = (primaryAlgo === 'dsa' || primaryAlgo === 'elgamal')
        ? `⚠ ${meta.algorithm?.toUpperCase() ?? 'Legacy'} key — weak algorithm`
        : `${meta.algorithm?.toUpperCase() ?? 'RSA'} key — ECC subkeys recommended`;
      el('key-algorithm-badge').textContent = badgeLabel;
      el('key-legacy-row').classList.remove('pgp-hidden');

      // Check whether modern subkeys have already been added, then show/hide
      // the Add Modern Subkeys trigger button accordingly.
      const armoredPublic = getPublicKey();
      if (armoredPublic) {
        hasModernSubkeys(armoredPublic).then(alreadyModern => {
          if (alreadyModern) {
            // Modern ECC subkeys already present — no action needed; suppress the warning.
            el('key-legacy-row').classList.add('pgp-hidden');
            el('panel-add-subkeys-trigger').classList.add('pgp-hidden');
          } else {
            el('panel-add-subkeys-trigger').classList.remove('pgp-hidden');
          }
        }).catch(() => {
          // On error, show the button anyway — addModernSubkeys will handle it
          el('panel-add-subkeys-trigger').classList.remove('pgp-hidden');
        });
      }
    } else {
      el('key-legacy-row').classList.add('pgp-hidden');
      el('panel-add-subkeys-trigger').classList.add('pgp-hidden');
      el('panel-add-subkeys').classList.add('pgp-hidden');
    }
  }
}

function showGenerateForm() {
  el('panel-generate-form').classList.remove('pgp-hidden');
  hideStatus('gen-status');
  el('gen-name').focus();
}

function hideGenerateForm() {
  el('panel-delete-confirm').classList.add('pgp-hidden');
  el('panel-generate-form').classList.add('pgp-hidden');
  // Reset key type to recommended default
  el('gen-key-type-ecc').checked = true;
  el('gen-name').value = '';
  el('gen-email').value = '';
  el('gen-passphrase').value = '';
  el('gen-passphrase-confirm').value = '';
  hideStatus('gen-status');
}

async function handleGenerate() {
  const name       = el('gen-name').value.trim();
  const email      = el('gen-email').value.trim();
  const passphrase = el('gen-passphrase').value;
  const confirm    = el('gen-passphrase-confirm').value;

  // Read selected key type from radio buttons (defaults to 'ecc')
  const keyTypeEl = document.querySelector('input[name="gen-key-type"]:checked');
  const keyType   = keyTypeEl?.value || 'ecc';

  if (!name)       return showStatus('gen-status', 'Full name is required.', 'error');
  if (!email)      return showStatus('gen-status', 'Email address is required.', 'error');
  if (!passphrase) return showStatus('gen-status', 'A passphrase is required.', 'error');
  if (passphrase !== confirm) return showStatus('gen-status', 'Passphrases do not match.', 'error');
  if (passphrase.length < 8) return showStatus('gen-status', 'Passphrase must be at least 8 characters.', 'warning');

  const btn = el('btn-generate-confirm');
  const spinner = el('gen-spinner');
  btn.disabled = true;
  spinner.classList.remove('pgp-hidden');

  // RSA-4096 generation can take several seconds — let the user know
  if (keyType === 'rsa4096') {
    showStatus('gen-status', 'Generating RSA-4096 key — this may take 5–15 seconds…', 'info');
  } else {
    hideStatus('gen-status');
  }

  try {
    const { privateKey: armoredPrivate, publicKey: armoredPublic } = await generateKeyPair(name, email, passphrase, keyType);
    const info = await getKeyInfo(armoredPublic);

    await saveKeyPair(armoredPrivate, armoredPublic, {
      name:                 info.name,
      email:                info.email,
      fingerprint:          info.fingerprint,
      fingerprintFormatted: info.fingerprintFormatted,
      keyId:                info.keyId,
      created:              info.created?.toISOString(),
      expires:              info.expires?.toISOString() ?? null,
      algorithm:            info.algorithm,
    });

    hideGenerateForm();
    await refreshMyKeyPanel();
    showStatus('status-bar', 'Key pair generated and saved successfully.', 'success');
  } catch (e) {
    showStatus('gen-status', `Key generation failed: ${e.message}`, 'error');
  } finally {
    btn.disabled = false;
    spinner.classList.add('pgp-hidden');
  }
}

async function handleCopyPublicKey() {
  const armoredKey = getPublicKey();
  if (!armoredKey) return;
  try {
    await navigator.clipboard.writeText(armoredKey);
    showStatus('status-bar', 'Public key copied to clipboard.', 'success');
  } catch {
    // Fallback — show it in a prompt
    window.prompt('Copy the public key below:', armoredKey);
  }
}

function handleSendPublicKey() {
  const armoredKey = getPublicKey();
  const meta = getKeyMetadata();
  if (!armoredKey || !meta) return;

  // displayNewMessageForm opens a new compose window pre-filled with the
  // public key.  The recipient is left blank so the user can address it.
  // The armored key is wrapped in <pre> to preserve its line structure in
  // the HTML body; the recipient can copy it out and import it in any
  // OpenPGP-compatible client.
  Office.context.mailbox.displayNewMessageForm({
    toRecipients: [],
    subject: `PGP Public Key for ${meta.name}`,
    htmlBody:
      `<p>Hi,</p>` +
      `<p>Please find my PGP public key below. ` +
      `You can use this to send me encrypted messages.</p>` +
      `<p><strong>Fingerprint:</strong> ${meta.fingerprintFormatted || meta.fingerprint}</p>` +
      `<pre>${armoredKey}</pre>`,
  });
}

/**
 * Download the user's passphrase-encrypted private key as a .asc file.
 *
 * The exported file is the same armored text stored in roaming settings —
 * it is already encrypted with the user's passphrase (AES-256 via OpenPGP
 * S2K), so it is safe to store on disk, in cloud storage, etc.  It cannot
 * be decrypted or used without the passphrase.
 *
 * The suggested filename includes the key's short fingerprint ID so the user
 * can tell multiple backups apart, e.g. "pgp-private-key-ABCD1234.asc".
 */
function handleExportPrivateKey() {
  const armoredPrivate = getPrivateKey();
  if (!armoredPrivate) return;

  const meta     = getKeyMetadata();
  const shortId  = meta?.fingerprint?.slice(-8) || 'key';
  const filename = `pgp-private-key-${shortId}.asc`;

  const blob = new Blob([armoredPrivate], { type: 'text/plain;charset=utf-8' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href     = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 5000);

  showStatus('status-bar',
    `Private key backup downloaded as "${filename}". ` +
    'Keep this file secure — it is protected by your passphrase.',
    'warning'
  );
}

/**
 * First click on "Delete Key" — show the inline confirmation panel.
 * window.confirm() is blocked in sandboxed Office task-pane iframes (OWA),
 * so we use a visible confirm/cancel panel in the DOM instead.
 */
function handleDeleteKey() {
  el('panel-delete-confirm').classList.remove('pgp-hidden');
}

async function handleDeleteConfirm() {
  el('panel-delete-confirm').classList.add('pgp-hidden');
  await clearKeyPair();
  await refreshMyKeyPanel();
  showStatus('status-bar', 'Key pair deleted.', 'warning');
}

function handleDeleteCancel() {
  el('panel-delete-confirm').classList.add('pgp-hidden');
}

// ── Import existing private key ───────────────────────────────────────────────

/**
 * Holds parsed key data between the initial import attempt and the user's
 * confirmation of the legacy-key warning.  Cleared whenever the import form
 * is closed or the legacy warning is dismissed.
 *
 * @type {{ armoredPrivate: string, armoredPublic: string, info: object } | null}
 */
let _pendingLegacyImport = null;

function showImportKeyForm() {
  el('panel-import-key-form').classList.remove('pgp-hidden');
  el('panel-generate-form').classList.add('pgp-hidden');
  hideStatus('import-key-status');
  el('import-privkey-text').focus();
}

function hideImportKeyForm() {
  _pendingLegacyImport = null;
  el('panel-import-key-form').classList.add('pgp-hidden');
  el('import-privkey-text').value = '';
  el('import-privkey-passphrase').value = '';
  hideStatus('import-key-status');
  el('import-legacy-warning').classList.add('pgp-hidden');
  el('import-key-buttons').classList.remove('pgp-hidden');
}

/**
 * Save a key pair (both armored strings + metadata object) to roaming settings.
 * Extracted as a helper so it is shared between the normal and legacy-confirm paths.
 */
async function _doSaveImport(armoredPrivate, armoredPublic, info) {
  await saveKeyPair(armoredPrivate, armoredPublic, {
    name:                 info.name,
    email:                info.email,
    fingerprint:          info.fingerprint,
    fingerprintFormatted: info.fingerprintFormatted,
    keyId:                info.keyId,
    created:              info.created?.toISOString(),
    expires:              info.expires?.toISOString() ?? null,
    algorithm:            info.algorithm,
  });
  hideImportKeyForm();
  await refreshMyKeyPanel();
}

/**
 * Validate, verify, and save an existing armored private key.
 *
 * Steps:
 *  1. Check the pasted text looks like a private key block (fast fail).
 *  2. Parse the key with OpenPGP.js to catch malformed armor.
 *  3. Unlock with the provided passphrase to verify it is correct
 *     (wrong passphrase → clear error, key not saved).
 *  4. Extract the armored public key from the private key object.
 *  5. Read key metadata (fingerprint, UID, algorithm, etc.).
 *  6. If the key is a legacy DSA/ElGamal type, show the security warning panel
 *     and pause — the user must explicitly confirm before the key is saved.
 *  7. Otherwise, save both keys to roaming settings immediately.
 */
async function handleImportPrivateKey() {
  const armoredPrivate = el('import-privkey-text').value.trim();
  const passphrase     = el('import-privkey-passphrase').value;

  if (!armoredPrivate) {
    return showStatus('import-key-status', 'Paste your armored private key first.', 'error');
  }
  if (!armoredPrivate.includes('-----BEGIN PGP PRIVATE KEY BLOCK-----')) {
    return showStatus('import-key-status',
      'This does not look like a PGP private key. Make sure you paste the full block including the header and footer lines.',
      'error'
    );
  }
  if (!passphrase) {
    return showStatus('import-key-status', 'Passphrase is required to verify the key.', 'error');
  }

  const btn     = el('btn-import-key-confirm');
  const spinner = el('import-key-spinner');
  btn.disabled  = true;
  spinner.classList.remove('pgp-hidden');
  showStatus('import-key-status', 'Verifying key…', 'info');

  try {
    // Steps 1+2+3: Parse and unlock.  unlockPrivateKey automatically falls
    // back to a permissive config for legacy DSA/ElGamal keys that have
    // SHA-1 self-signatures; the fallback is transparent here.
    await unlockPrivateKey(armoredPrivate, passphrase);

    // Step 4: Extract the public key (also fallback-aware for legacy keys).
    const armoredPublic = await extractPublicKey(armoredPrivate);

    // Step 5: Read metadata.
    const info = await getKeyInfo(armoredPublic);

    // Step 6: Detect legacy DSA/ElGamal keys and require explicit confirmation.
    // DSA is the primary key algorithm for the classic DSA+ElGamal key type;
    // OpenPGP.js reports it as algorithm 'dsa' from the key packet.
    const isLegacy = info.algorithm === 'dsa' || info.algorithm === 'elgamal';

    if (isLegacy) {
      // Store the parsed data and let the user read the warning before saving.
      _pendingLegacyImport = { armoredPrivate, armoredPublic, info };
      el('import-key-buttons').classList.add('pgp-hidden');
      el('import-legacy-warning').classList.remove('pgp-hidden');
      hideStatus('import-key-status');
      return; // save deferred to handleConfirmLegacyImport
    }

    // Step 7: Save immediately for non-legacy keys.
    await _doSaveImport(armoredPrivate, armoredPublic, info);
    showStatus('status-bar',
      `Key imported: ${info.name} <${info.email}> — ${info.fingerprintFormatted}`,
      'success'
    );
  } catch (e) {
    // Provide a clearer message for the most common failure (wrong passphrase).
    const msg = e.message?.toLowerCase() ?? '';
    if (msg.includes('passphrase') || msg.includes('decrypt') || msg.includes('session key')) {
      showStatus('import-key-status', 'Incorrect passphrase — please try again.', 'error');
    } else {
      showStatus('import-key-status', `Import failed: ${e.message}`, 'error');
    }
  } finally {
    btn.disabled = false;
    spinner.classList.add('pgp-hidden');
  }
}

/**
 * Called when the user clicks "I understand — import anyway" on the legacy
 * key warning panel.  Saves the key pair that was validated in
 * handleImportPrivateKey and stored in _pendingLegacyImport.
 */
async function handleConfirmLegacyImport() {
  if (!_pendingLegacyImport) return;
  const { armoredPrivate, armoredPublic, info } = _pendingLegacyImport;

  const btn     = el('btn-import-legacy-confirm');
  btn.disabled  = true;

  try {
    await _doSaveImport(armoredPrivate, armoredPublic, info);
    showStatus('status-bar',
      `Legacy key imported: ${info.name} <${info.email}> — ${info.fingerprintFormatted}. ` +
      `Consider generating a new ECC key for stronger security.`,
      'warning'
    );
  } catch (e) {
    el('import-legacy-warning').classList.add('pgp-hidden');
    el('import-key-buttons').classList.remove('pgp-hidden');
    showStatus('import-key-status', `Import failed: ${e.message}`, 'error');
  } finally {
    btn.disabled = false;
  }
}

// ── Add Modern Subkeys ────────────────────────────────────────────────────────

function showAddSubkeysForm() {
  el('panel-add-subkeys-trigger').classList.add('pgp-hidden');
  el('panel-add-subkeys').classList.remove('pgp-hidden');
  hideStatus('add-subkeys-status');
  el('add-subkeys-passphrase').focus();
}

function hideAddSubkeysForm() {
  el('panel-add-subkeys').classList.add('pgp-hidden');
  el('add-subkeys-passphrase').value = '';
  hideStatus('add-subkeys-status');
  // Re-evaluate whether to show the trigger button (same isNonEccAlgo logic as
  // refreshMyKeyPanel — must cover RSA as well as DSA/ElGamal).
  const meta = getKeyMetadata();
  const primaryAlgo = (meta?.algorithm || '').toLowerCase();
  const isNonEccAlgo = primaryAlgo.startsWith('rsa') ||
    primaryAlgo === 'dsa' || primaryAlgo === 'elgamal';
  if (isNonEccAlgo) {
    const armoredPublic = getPublicKey();
    if (armoredPublic) {
      hasModernSubkeys(armoredPublic).then(already => {
        el('panel-add-subkeys-trigger').classList.toggle('pgp-hidden', already);
      }).catch(() => {
        el('panel-add-subkeys-trigger').classList.remove('pgp-hidden');
      });
    }
  }
}

/**
 * Append Ed25519 and X25519 subkeys to the stored legacy private key.
 * The passphrase is used to authorize the operation (unlock the DSA primary key
 * for signing) and to re-encrypt the new subkeys before saving.
 */
async function handleAddSubkeys() {
  const passphrase = el('add-subkeys-passphrase').value;
  if (!passphrase) {
    return showStatus('add-subkeys-status', 'Passphrase is required.', 'error');
  }

  const armoredPrivate = getPrivateKey();
  if (!armoredPrivate) return;

  const btn     = el('btn-add-subkeys-confirm');
  const spinner = el('add-subkeys-spinner');
  btn.disabled  = true;
  spinner.classList.remove('pgp-hidden');
  showStatus('add-subkeys-status', 'Adding modern subkeys — this may take a moment…', 'info');

  try {
    const { armoredPrivate: newPrivate, armoredPublic: newPublic } =
      await addModernSubkeys(armoredPrivate, passphrase);

    // Re-read info from the updated key (algorithm is still 'dsa' for the primary,
    // but subkeys now include Ed25519 and X25519).
    const info = await getKeyInfo(newPublic);
    const meta = getKeyMetadata();

    await saveKeyPair(newPrivate, newPublic, {
      // Preserve original metadata fields; only update algorithm if desired
      name:                 meta?.name  ?? info.name,
      email:                meta?.email ?? info.email,
      fingerprint:          meta?.fingerprint          ?? info.fingerprint,
      fingerprintFormatted: meta?.fingerprintFormatted ?? info.fingerprintFormatted,
      keyId:                meta?.keyId ?? info.keyId,
      created:              meta?.created ?? info.created?.toISOString(),
      expires:              meta?.expires ?? (info.expires?.toISOString() ?? null),
      algorithm:            meta?.algorithm ?? info.algorithm,
    });

    hideAddSubkeysForm();
    // Hide the trigger button — modern subkeys are now present
    el('panel-add-subkeys-trigger').classList.add('pgp-hidden');
    await refreshMyKeyPanel();
    showStatus('status-bar',
      'Modern subkeys added: Ed25519 (sign) + X25519 (encrypt). ' +
      'Share your updated public key with contacts.',
      'success'
    );
  } catch (e) {
    const msg = e.message?.toLowerCase() ?? '';
    if (msg.includes('passphrase') || msg.includes('decrypt')) {
      showStatus('add-subkeys-status', 'Incorrect passphrase — please try again.', 'error');
    } else {
      showStatus('add-subkeys-status', `Failed: ${e.message}`, 'error');
    }
  } finally {
    btn.disabled = false;
    spinner.classList.add('pgp-hidden');
  }
}

// ── Keyring panel ─────────────────────────────────────────────────────────────

async function refreshKeyringPanel() {
  const list = el('keyring-list');
  const empty = el('keyring-empty');
  const countBadge = el('keyring-count');
  const storageWarning = el('storage-warning');

  const contacts = await listContactKeys();
  countBadge.textContent = `${contacts.length} key${contacts.length !== 1 ? 's' : ''}`;

  // Remove all existing items except the empty placeholder
  Array.from(list.querySelectorAll('.pgp-key-item-wrapper')).forEach(el => el.remove());

  if (contacts.length === 0) {
    empty.classList.remove('pgp-hidden');
  } else {
    empty.classList.add('pgp-hidden');
    contacts.forEach(contact => {
      const li = document.createElement('li');
      li.className = 'pgp-key-item pgp-key-item-wrapper';
      li.dataset.email = contact.email;

      if (contact.error) {
        li.innerHTML = `
          <div class="pgp-key-item__header">
            <div class="pgp-key-item__identity">
              <div class="pgp-key-item__email">${escHtml(contact.email)}</div>
              <div class="pgp-key-item__meta" style="color:#a80000;">${escHtml(contact.error)}</div>
            </div>
            <button class="pgp-btn pgp-btn--danger pgp-btn--sm btn-remove-key" data-email="${escHtml(contact.email)}">Remove</button>
          </div>`;
      } else {
        const info = contact.info;
        li.innerHTML = `
          <div class="pgp-key-item__header">
            <div class="pgp-key-item__identity">
              <div class="pgp-key-item__email">${escHtml(contact.email)}</div>
              ${info.name ? `<div class="pgp-key-item__name">${escHtml(info.name)}</div>` : ''}
              <div class="pgp-key-item__meta">
                Algorithm: ${escHtml(info.algorithm)} &nbsp;·&nbsp;
                ${info.expires ? `Expires: ${formatDate(info.expires)}` : 'No expiration'}
              </div>
            </div>
            <button class="pgp-btn pgp-btn--danger pgp-btn--sm btn-remove-key" data-email="${escHtml(contact.email)}">Remove</button>
          </div>
          <span class="pgp-fingerprint">${escHtml(info.fingerprintFormatted)}</span>`;
      }
      list.appendChild(li);
    });
  }

  // Storage warning
  const storageInfo = getKeyringStorageInfo();
  storageWarning.classList.toggle('pgp-hidden', !storageInfo.nearLimit);
}

async function handleFindKey() {
  const email = el('keyring-search').value.trim();
  if (!email) return;

  const container = el('find-result');
  container.innerHTML = `<div class="pgp-alert pgp-alert--info"><span class="pgp-spinner"></span> Looking up key for ${escHtml(email)}…</div>`;
  container.classList.remove('pgp-hidden');

  try {
    const result = await discoverKey(email);

    if (result.status === KeyStatus.NOT_FOUND) {
      container.innerHTML = `<div class="pgp-alert pgp-alert--warning">No key found for <strong>${escHtml(email)}</strong> via WKD or keyserver. You can import one manually.</div>`;
      return;
    }

    const info = await getKeyInfo(result.key.armor());
    const sourceLabel = result.source;

    let html = `
      <div class="pgp-alert pgp-alert--success">
        <div>
          <strong>Key found</strong> via ${escHtml(sourceLabel)}<br/>
          ${info.name ? `${escHtml(info.name)}<br/>` : ''}
          <span class="pgp-fingerprint" style="margin-top:4px;">${escHtml(info.fingerprintFormatted)}</span>
        </div>
      </div>`;

    if (result.status !== KeyStatus.FOUND_LOCAL && result.armoredKey) {
      html += `<button class="pgp-btn pgp-btn--primary pgp-btn--sm pgp-mt-sm" id="btn-save-found-key" data-email="${escHtml(email)}" data-key="${escHtml(result.armoredKey)}">Save to Keyring</button>`;
    } else {
      html += `<div class="pgp-badge pgp-badge--success pgp-mt-sm">Already in keyring</div>`;
    }

    container.innerHTML = html;

    el('btn-save-found-key')?.addEventListener('click', async (e) => {
      const btn = e.currentTarget;
      try {
        await addContactKey(btn.dataset.email, btn.dataset.key);
        await refreshKeyringPanel();
        container.innerHTML = `<div class="pgp-alert pgp-alert--success">Key for ${escHtml(email)} saved to keyring.</div>`;
      } catch (err) {
        container.innerHTML = `<div class="pgp-alert pgp-alert--error">Could not save key: ${escHtml(err.message)}</div>`;
      }
    });

  } catch (e) {
    container.innerHTML = `<div class="pgp-alert pgp-alert--error">Lookup failed: ${escHtml(e.message)}</div>`;
  }
}

async function handleImportKey() {
  const email      = el('import-email').value.trim();
  const armoredKey = el('import-key-text').value.trim();

  if (!email)      return showStatus('import-status', 'Email address is required.', 'error');
  if (!armoredKey) return showStatus('import-status', 'Paste an armored public key.', 'error');

  try {
    const { info, storageWarning } = await addContactKey(email, armoredKey);
    el('import-email').value = '';
    el('import-key-text').value = '';
    el('panel-import-form').classList.add('pgp-hidden');
    await refreshKeyringPanel();
    showStatus('status-bar',
      `Key for ${email} saved. Fingerprint: ${info.fingerprintFormatted}`,
      storageWarning ? 'warning' : 'success'
    );
  } catch (e) {
    showStatus('import-status', `Could not save key: ${e.message}`, 'error');
  }
}

// ── Org settings panel ────────────────────────────────────────────────────────

function refreshOrgPanel() {
  const override = getOrgOverride();

  // The override form must be locked when IT policy (from the domain's
  // company-config.json) marks the company key as required.  An active
  // manual override means IT set this up explicitly, so those controls
  // remain editable.  Only the domain-policy case is locked.
  const lockedByPolicy = !override && isCompanyKeyRequired();

  if (lockedByPolicy) {
    showStatus('org-status',
      'Company key settings are enforced by your organization\'s IT policy and cannot be overridden.',
      'warning'
    );
  } else if (override) {
    showStatus('org-status',
      'Using manual override. Auto-discovery from domain is skipped.',
      'warning'
    );
  } else if (isCompanyKeyEnabled()) {
    showStatus('org-status',
      `Organization config loaded from domain. Company key: ${getCompanyKeyEmails().join(', ')}`,
      'info'
    );
  } else {
    showStatus('org-status',
      'No organization config found. Company key feature is disabled.',
      'neutral'
    );
  }

  el('org-key-enabled').checked  = isCompanyKeyEnabled();
  el('org-key-required').checked = isCompanyKeyRequired();
  el('org-key-emails').value     = getCompanyKeyEmails().join(', ');

  // Disable all override controls when IT policy is enforced
  el('org-key-enabled').disabled  = lockedByPolicy;
  el('org-key-required').disabled = lockedByPolicy;
  el('org-key-emails').disabled   = lockedByPolicy;
  el('btn-save-org').disabled     = lockedByPolicy;
  el('btn-clear-org').disabled    = lockedByPolicy;
}

async function handleSaveOrgOverride() {
  // Defensive guard — the UI should already prevent this via disabled controls,
  // but block programmatic or unexpected calls when IT policy is enforced.
  if (!getOrgOverride() && isCompanyKeyRequired()) {
    showStatus('org-save-status',
      'Cannot save override: company key settings are enforced by IT policy.',
      'error'
    );
    return;
  }

  const enabled  = el('org-key-enabled').checked;
  const required = el('org-key-required').checked;
  const emailsRaw = el('org-key-emails').value;
  const emails = emailsRaw.split(',').map(e => e.trim()).filter(Boolean);

  await saveOrgOverride({ companyKeyEnabled: enabled, companyKeyRequired: required, companyKeyEmails: emails });
  showStatus('org-save-status', 'Override saved.', 'success');
  refreshOrgPanel();
}

async function handleClearOrgOverride() {
  await clearOrgOverride();
  showStatus('org-save-status', 'Override cleared. Auto-discovery will be used.', 'info');
  refreshOrgPanel();
}

// ── Personal preferences panel ────────────────────────────────────────────────

/**
 * Populate the Personal Preferences section from stored settings.
 */
function refreshPrefsPanel() {
  el('pref-sign-default').checked = getSignDefault();
}

/**
 * Persist the user's personal compose preferences to roaming settings.
 */
async function handleSavePrefs() {
  const signDefault = el('pref-sign-default').checked;
  await saveSignDefault(signDefault);
  showStatus('prefs-save-status', 'Preferences saved.', 'success');
}

// ── XSS-safe HTML escaping ────────────────────────────────────────────────────

function escHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// ── Bootstrap ─────────────────────────────────────────────────────────────────

Office.onReady(async () => {
  const userEmail = Office.context.mailbox.userProfile?.emailAddress || '';

  // Load org config from domain or override
  await loadOrgConfig(userEmail);

  // Wire generate-form buttons (there are two "show-generate" buttons)
  document.querySelectorAll('#btn-show-generate').forEach(btn => {
    btn.addEventListener('click', () => {
      hideImportKeyForm();
      showGenerateForm();
    });
  });
  el('btn-generate-cancel').addEventListener('click', hideGenerateForm);
  el('btn-generate-confirm').addEventListener('click', handleGenerate);

  // Wire import-key-form buttons (there are two "show-import-key" buttons)
  document.querySelectorAll('#btn-show-import-key').forEach(btn => {
    btn.addEventListener('click', () => {
      hideGenerateForm();
      showImportKeyForm();
    });
  });
  el('btn-import-key-cancel').addEventListener('click', hideImportKeyForm);
  el('btn-import-key-confirm').addEventListener('click', handleImportPrivateKey);
  el('btn-import-legacy-confirm').addEventListener('click', handleConfirmLegacyImport);
  el('btn-import-legacy-cancel').addEventListener('click', () => {
    _pendingLegacyImport = null;
    el('import-legacy-warning').classList.add('pgp-hidden');
    el('import-key-buttons').classList.remove('pgp-hidden');
    hideStatus('import-key-status');
  });

  // Add Modern Subkeys panel
  el('btn-add-subkeys').addEventListener('click', showAddSubkeysForm);
  el('btn-add-subkeys-cancel').addEventListener('click', hideAddSubkeysForm);
  el('btn-add-subkeys-confirm').addEventListener('click', handleAddSubkeys);

  el('btn-copy-pubkey')?.addEventListener('click', handleCopyPublicKey);
  el('btn-send-pubkey')?.addEventListener('click', handleSendPublicKey);
  el('btn-export-key')?.addEventListener('click', handleExportPrivateKey);
  el('btn-delete-key').addEventListener('click', handleDeleteKey);
  el('btn-delete-confirm').addEventListener('click', handleDeleteConfirm);
  el('btn-delete-cancel').addEventListener('click', handleDeleteCancel);

  // Keyring
  el('btn-find-key').addEventListener('click', handleFindKey);
  el('keyring-search').addEventListener('keydown', e => { if (e.key === 'Enter') handleFindKey(); });

  el('btn-show-import').addEventListener('click', () => {
    el('panel-import-form').classList.toggle('pgp-hidden');
  });
  el('btn-import-cancel').addEventListener('click', () => {
    el('panel-import-form').classList.add('pgp-hidden');
  });
  el('btn-import-confirm').addEventListener('click', handleImportKey);

  // Keyring list — delegate remove button clicks.
  // Two-click confirmation: first click primes the button; second click within
  // 3 seconds executes.  Avoids window.confirm(), which is blocked in sandboxed
  // Office task-pane iframes (OWA).
  el('keyring-list').addEventListener('click', async (e) => {
    const btn = e.target.closest('.btn-remove-key');
    if (!btn) return;

    if (!btn.dataset.confirming) {
      btn.dataset.confirming = '1';
      btn.textContent = 'Confirm?';
      setTimeout(() => {
        if (btn.dataset.confirming) {
          btn.dataset.confirming = '';
          btn.textContent = 'Remove';
        }
      }, 3000);
      return;
    }

    btn.dataset.confirming = '';
    const email = btn.dataset.email;
    await removeContactKey(email);
    await refreshKeyringPanel();
    showStatus('status-bar', `Key for ${email} removed.`, 'info');
  });

  // Org settings
  el('btn-save-org').addEventListener('click', handleSaveOrgOverride);
  el('btn-clear-org').addEventListener('click', handleClearOrgOverride);

  // Personal preferences
  el('btn-save-prefs').addEventListener('click', handleSavePrefs);

  // Initial render
  await refreshMyKeyPanel();
  await refreshKeyringPanel();
  refreshOrgPanel();
  refreshPrefsPanel();

  // Ko-fi support button — loaded dynamically so the external script is never
  // fetched when an org hides the button via hideSupportButton in org config.
  if (!isSupportButtonHidden()) {
    const container = el('section-kofi');
    const script = document.createElement('script');
    script.src = 'https://storage.ko-fi.com/cdn/widget/Widget_2.js';
    script.onload = () => {
      kofiwidget2.init('Support me on Ko-fi', '#72a4f2', 'R6R61WMZMW');
      kofiwidget2.draw();
    };
    container.appendChild(script);
  }
});
