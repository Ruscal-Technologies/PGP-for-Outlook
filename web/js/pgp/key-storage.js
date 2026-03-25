/**
 * key-storage.js
 * Wrapper around Office.context.roamingSettings for persisting PGP keys.
 *
 * WHY roaming settings?
 *   Office roaming settings are tied to the user's Microsoft 365 account and
 *   sync automatically across all devices where the user is signed in.  This
 *   means keys generated on one machine are immediately available in Outlook
 *   Web, desktop, and mobile without any manual export/import step.
 *
 *   The trade-off is a tight storage cap (~32 KB for all settings combined).
 *   We store the private key encrypted at rest (AES-256, protected by the
 *   user's passphrase) and never store the passphrase itself.
 *
 * Storage key layout:
 *   pgp_private_key  — Armored, passphrase-encrypted private key string
 *   pgp_public_key   — Armored public key string
 *   pgp_key_meta     — Object: { name, email, fingerprint, keyId, created, expires, algorithm }
 *   pgp_keyring      — Object: { "email@example.com": "-----BEGIN PGP PUBLIC KEY BLOCK-----..." }
 *   pgp_org_override — Object: manual org config override (see org-config.js)
 *   pgp_sign_default — Boolean: user's personal default for the "sign messages" toggle
 *                      (defaults to false; overridable per-message in the compose pane)
 *
 * Storage budget (approximate):
 *   Private key (ECC/curve25519 + passphrase encryption): ~3–6 KB
 *   Private key (RSA-4096 + passphrase encryption):       ~6–10 KB
 *   Public key (ECC): ~0.5–1 KB  |  Public key (RSA-4096): ~1.5–2 KB
 *   Metadata: <0.2 KB
 *   Per contact public key: ~1–3 KB
 *   ⇒ Comfortable room for ~8–10 ECC contact keys before approaching the limit.
 *   Use estimateStorageUsage() and STORAGE_LIMIT_BYTES to warn the user early.
 */

const KEYS = {
  PRIVATE:      'pgp_private_key',
  PUBLIC:       'pgp_public_key',
  META:         'pgp_key_meta',
  KEYRING:      'pgp_keyring',
  ORG_OVERRIDE: 'pgp_org_override',
  SIGN_DEFAULT: 'pgp_sign_default',
};

function settings() {
  return Office.context.roamingSettings;
}

/**
 * Persist all pending in-memory setting changes to the server.
 *
 * IMPORTANT: roamingSettings.set() only updates the in-memory copy.  Changes
 * are lost if the task pane closes before saveAsync() completes.  Every public
 * write function in this module awaits saveAsync() before resolving, so callers
 * never need to call it themselves.
 */
function saveAsync() {
  return new Promise((resolve, reject) => {
    settings().saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        reject(new Error(result.error.message));
      } else {
        resolve();
      }
    });
  });
}

// ── Own key pair ──────────────────────────────────────────────────────────────

export function hasKeyPair() {
  return !!settings().get(KEYS.PRIVATE);
}

export function getPrivateKey() {
  return settings().get(KEYS.PRIVATE) || null;
}

export function getPublicKey() {
  return settings().get(KEYS.PUBLIC) || null;
}

export function getKeyMetadata() {
  return settings().get(KEYS.META) || null;
}

export async function saveKeyPair(armoredPrivateKey, armoredPublicKey, metadata) {
  settings().set(KEYS.PRIVATE, armoredPrivateKey);
  settings().set(KEYS.PUBLIC, armoredPublicKey);
  settings().set(KEYS.META, metadata);
  await saveAsync();
}

export async function clearKeyPair() {
  settings().remove(KEYS.PRIVATE);
  settings().remove(KEYS.PUBLIC);
  settings().remove(KEYS.META);
  await saveAsync();
}

// ── Keyring (contacts' public keys) ──────────────────────────────────────────

export function getKeyring() {
  return settings().get(KEYS.KEYRING) || {};
}

export async function saveKeyring(keyring) {
  settings().set(KEYS.KEYRING, keyring);
  await saveAsync();
}

// ── Org config override ───────────────────────────────────────────────────────

export function getOrgOverride() {
  return settings().get(KEYS.ORG_OVERRIDE) || null;
}

export async function saveOrgOverride(config) {
  settings().set(KEYS.ORG_OVERRIDE, config);
  await saveAsync();
}

export async function clearOrgOverride() {
  settings().remove(KEYS.ORG_OVERRIDE);
  await saveAsync();
}

// ── Personal compose preferences ─────────────────────────────────────────────

/**
 * Return the user's stored sign-by-default preference.
 * When not set, returns false (signing off by default).
 * Users can change this in Manage Keys → Personal Preferences and
 * override it per-message using the sign toggle in the compose pane.
 *
 * @returns {boolean}
 */
export function getSignDefault() {
  const stored = settings().get(KEYS.SIGN_DEFAULT);
  // Treat any non-boolean stored value as false to be safe.
  return stored === true;
}

/**
 * Persist the user's sign-by-default preference.
 * @param {boolean} value
 */
export async function saveSignDefault(value) {
  settings().set(KEYS.SIGN_DEFAULT, !!value);
  await saveAsync();
}

// ── Storage diagnostics ───────────────────────────────────────────────────────

/**
 * Estimate current roaming settings storage usage in bytes.
 * Roaming settings are serialized as JSON; this gives an approximation.
 */
export function estimateStorageUsage() {
  const data = {
    [KEYS.PRIVATE]:      settings().get(KEYS.PRIVATE) || '',
    [KEYS.PUBLIC]:       settings().get(KEYS.PUBLIC) || '',
    [KEYS.META]:         settings().get(KEYS.META) || {},
    [KEYS.KEYRING]:      settings().get(KEYS.KEYRING) || {},
    [KEYS.ORG_OVERRIDE]: settings().get(KEYS.ORG_OVERRIDE) || {},
    [KEYS.SIGN_DEFAULT]: settings().get(KEYS.SIGN_DEFAULT) || false,
  };
  return JSON.stringify(data).length;
}

export const STORAGE_LIMIT_BYTES = 32768; // 32KB Office roaming settings ceiling
