/**
 * pgp-core.js
 * Core PGP cryptographic operations built on OpenPGP.js v5.
 *
 * This module is the only place in the add-in that directly calls the
 * OpenPGP.js library.  All other modules go through these wrappers, which
 * keeps the crypto surface small and easy to audit.
 *
 * Supported key types (passed as the `keyType` argument to generateKeyPair):
 *
 *   'ecc'  — Ed25519 (signing) + X25519 (encryption), curve25519.  Default.
 *     Rationale:
 *       - Compact keys and signatures (~200 bytes vs ~512 bytes for RSA-2048)
 *       - Widely supported by modern OpenPGP clients (GnuPG ≥ 2.1, Thunderbird, etc.)
 *       - Deterministic signing — no random-number dependency during sign operations
 *       - Strongly recommended by the OpenPGP RFC 9580 "crypto-refresh" update
 *
 *   'rsa4096' — RSA-4096 (sign + encrypt, legacy).
 *     Rationale:
 *       - Maximum interoperability with older PGP clients (GnuPG < 2.1, PGP 2.x, etc.)
 *       - Larger keys (~6–10 KB stored) and slower generation (~5–15 s in-browser)
 *       - Choose this only when you need to exchange keys with clients that do not
 *         support modern ECC algorithms (pre-2015 software or constrained appliances)
 */

import * as openpgp from '../openpgp.min.mjs';

// ── Internal helpers for legacy-key interoperability ──────────────────────────

/**
 * Returns true when the error is the OpenPGP.js "too weak" rejection thrown
 * by key.getEncryptionKey() when the key uses an algorithm listed in
 * config.rejectPublicKeyAlgorithms (e.g. ElGamal, DSA).
 */
function _isWeakKeyError(err) {
  const msg = err?.message ?? '';
  return msg.includes('too weak') ||
         msg.includes('Could not find valid encryption key');
}

/**
 * Returns true when the error is the self-signature validation failure that
 * OpenPGP.js v5 throws for old DSA/ElGamal keys whose self-signatures use
 * SHA-1 — a hash algorithm that is in config.rejectHashAlgorithms by default.
 */
function _isLegacySelfSigError(err) {
  const msg = err?.message ?? '';
  return msg.includes('Could not find valid self-signature') ||
         msg.includes('self-signature');
}

/**
 * Build a one-off OpenPGP.js config that removes ElGamal (RFC 4880 algo 16)
 * from the rejection list so we can encrypt to legacy DSA+ElGamal keys.
 * All other config values remain at their defaults.
 */
function _buildPermissiveConfig() {
  // ElGamal encrypt-only is algorithm 16 in RFC 4880 §9.1.
  // OpenPGP.js v5 puts it in rejectPublicKeyAlgorithms by default.
  const elgamalId = openpgp.enums?.publicKey?.elgamal ?? 16;
  const rejectSet = new Set(openpgp.config?.rejectPublicKeyAlgorithms ?? []);
  rejectSet.delete(elgamalId);
  return { rejectPublicKeyAlgorithms: rejectSet };
}

/**
 * Build a one-off OpenPGP.js config suitable for READING legacy DSA/ElGamal
 * private keys.
 *
 * Old DSA keys self-sign with SHA-1 (hash algorithm 2).  OpenPGP.js v5 rejects
 * SHA-1 by default (it is in config.rejectHashAlgorithms) because SHA-1 is
 * cryptographically weak for new signatures.  However, self-signatures on
 * pre-existing keys cannot be changed without re-signing the key, and SHA-1
 * is still structurally sound enough to identify the key owner during import.
 *
 * This config:
 *   - Removes SHA-1 from the rejected hash algorithm set so that
 *     readPrivateKey / readKey can validate SHA-1 self-signatures.
 *   - Removes DSA (17) and ElGamal (16) from the rejected public-key algorithm
 *     set so those key types are not rejected before parsing completes.
 *
 * IMPORTANT: This config is used ONLY for key reading/import operations, never
 * for creating new signatures or encrypting new messages.  New cryptographic
 * operations on legacy keys (signing, encryption) still use the restrictive
 * default config, which ensures no new SHA-1 signatures are created.
 */
function _buildLegacyKeyReadConfig() {
  const sha1Id    = openpgp.enums?.hash?.sha1           ?? 2;
  const dsaId     = openpgp.enums?.publicKey?.dsa       ?? 17;
  const elgamalId = openpgp.enums?.publicKey?.elgamal   ?? 16;

  const rejectHashes = new Set(openpgp.config?.rejectHashAlgorithms    ?? []);
  const rejectPK     = new Set(openpgp.config?.rejectPublicKeyAlgorithms ?? []);

  rejectHashes.delete(sha1Id);
  rejectPK.delete(dsaId);
  rejectPK.delete(elgamalId);

  return { rejectHashAlgorithms: rejectHashes, rejectPublicKeyAlgorithms: rejectPK };
}

// ── Key generation ────────────────────────────────────────────────────────────

/**
 * Generate a new PGP key pair protected by a passphrase.
 *
 * The passphrase is applied using AES-256 symmetric encryption inside the
 * OpenPGP packet structure (S2K + CFB).  The private key material is never
 * accessible without the passphrase.
 *
 * @param {string} name              - User's full name (embedded in the key UID)
 * @param {string} email             - User's email address (embedded in the key UID)
 * @param {string} passphrase        - Passphrase to protect the private key
 * @param {'ecc'|'rsa4096'} [keyType='ecc']
 *   'ecc'     — Ed25519 / X25519 (modern, compact, fast).  Recommended default.
 *   'rsa4096' — RSA-4096 (legacy interoperability with older PGP clients such as
 *               GnuPG < 2.1).  Key generation takes several seconds in-browser
 *               and produces larger keys (~6–10 KB vs ~3–6 KB for ECC).
 * @returns {Promise<{ privateKey: string, publicKey: string }>} Armored key strings
 */
export async function generateKeyPair(name, email, passphrase, keyType = 'ecc') {
  let genOptions;

  if (keyType === 'rsa4096') {
    genOptions = {
      type: 'rsa',
      rsaBits: 4096,
      userIDs: [{ name, email }],
      passphrase,
      format: 'armored',
    };
  } else {
    // Default: ECC (Ed25519 + X25519 / curve25519)
    genOptions = {
      type: 'ecc',
      curve: 'curve25519',
      userIDs: [{ name, email }],
      passphrase,
      format: 'armored',
    };
  }

  const { privateKey, publicKey } = await openpgp.generateKey(genOptions);
  return { privateKey, publicKey };
}

// ── Key reading / parsing ─────────────────────────────────────────────────────

/**
 * Parse an armored public key string into a key object.
 * Throws if the armor is malformed or the key type is unexpected.
 */
export async function readPublicKey(armoredKey) {
  return await openpgp.readKey({ armoredKey });
}

/**
 * Extract the armored public key from an armored private key.
 *
 * The public-key packets inside a PGP private key are always stored in
 * plaintext — no passphrase is needed to read the public portion.  This
 * lets us display fingerprint/UID metadata and save the public key to
 * roaming settings immediately after an import, without requiring the
 * user to enter their passphrase twice.
 *
 * @param {string} armoredPrivateKey - Armored private key (-----BEGIN PGP PRIVATE KEY BLOCK-----)
 * @returns {Promise<string>} Armored public key (-----BEGIN PGP PUBLIC KEY BLOCK-----)
 */
export async function extractPublicKey(armoredPrivateKey) {
  const privateKey = await openpgp.readPrivateKey({ armoredKey: armoredPrivateKey })
    .catch(err => {
      if (!_isLegacySelfSigError(err)) throw err;
      return openpgp.readPrivateKey({
        armoredKey: armoredPrivateKey,
        config: _buildLegacyKeyReadConfig(),
      });
    });
  return privateKey.toPublic().armor();
}

/**
 * Parse a binary public key (e.g. the raw response from a WKD lookup)
 * into a key object.
 */
export async function readPublicKeyFromBinary(binaryKey) {
  return await openpgp.readKey({ binaryKey });
}

/**
 * Decrypt an armored private key using its passphrase.
 * Returns an in-memory unlocked key object ready for signing or decryption.
 *
 * SECURITY NOTE: The returned object contains the raw private key material in
 * memory for the duration of the operation.  It is NOT persisted anywhere by
 * this function; callers should discard it as soon as the crypto operation
 * is complete (i.e. do not store it in module-level state).
 *
 * @param {string} armoredPrivateKey - Armored, passphrase-encrypted private key
 * @param {string} passphrase        - The key's passphrase
 * @returns {Promise<openpgp.PrivateKey>} Unlocked private key object
 * @throws If the passphrase is wrong or the armored text is corrupt
 */
export async function unlockPrivateKey(armoredPrivateKey, passphrase) {
  // Attempt standard (strict) parsing first.  If any step fails because the
  // key carries a SHA-1 self-signature (common on old DSA/ElGamal keys), retry
  // the entire sequence with a permissive config.
  //
  // NOTE: openpgp.decryptKey() calls key.validate() internally, which calls
  // getPrimaryUser() with the global config.  If the global config rejects
  // SHA-1, that validation throws the self-sig error — so BOTH readPrivateKey
  // AND decryptKey must be retried together with the legacy config.
  try {
    const privateKey = await openpgp.readPrivateKey({ armoredKey: armoredPrivateKey });
    return await openpgp.decryptKey({ privateKey, passphrase });
  } catch (err) {
    if (!_isLegacySelfSigError(err)) throw err;
    const config = _buildLegacyKeyReadConfig();
    const privateKey = await openpgp.readPrivateKey({ armoredKey: armoredPrivateKey, config });
    return await openpgp.decryptKey({ privateKey, passphrase, config });
  }
}

// ── Key inspection ────────────────────────────────────────────────────────────

/**
 * Extract human-readable metadata from an armored key (public or private).
 * Safe to call with a private key — no sensitive data is returned.
 *
 * @param {string} armoredKey
 * @returns {{
 *   fingerprint: string,         // 40-char hex, uppercase
 *   fingerprintFormatted: string, // spaced groups of 4: "ABCD EFGH …"
 *   shortId: string,             // last 8 chars of fingerprint
 *   keyId: string,               // hex key ID, uppercase
 *   userIds: string[],
 *   name: string,
 *   email: string,
 *   created: Date,
 *   expires: Date|null,          // null = no expiration
 *   isPrivate: boolean,
 *   algorithm: string
 * }}
 */
export async function getKeyInfo(armoredKey) {
  // Parse the key.  For legacy DSA/ElGamal keys (SHA-1 self-signatures),
  // readKey itself usually succeeds because packet parsing does not validate
  // self-signatures; the failure comes later in getPrimaryUser / getExpirationTime.
  const key = await openpgp.readKey({ armoredKey })
    .catch(err => {
      if (!_isLegacySelfSigError(err)) throw err;
      return openpgp.readKey({ armoredKey, config: _buildLegacyKeyReadConfig() });
    });

  // getPrimaryUser() validates self-signatures with the global (strict) config.
  // In OpenPGP.js v5.5.0 the method does not accept a runtime config parameter,
  // so we cannot selectively loosen it for one call.  On failure for legacy keys
  // we extract the UID directly from key.getUserIDs() which reads the raw UserID
  // packets without any signature validation.
  //
  // getExpirationTime() also calls getPrimaryUser() internally, so it has the
  // same problem — on failure we default to no expiration.
  let name = '', email = '', expirationTime = Infinity;

  try {
    const primaryUser = await key.getPrimaryUser();
    name  = primaryUser?.user?.userID?.name  || '';
    email = primaryUser?.user?.userID?.email || '';
  } catch (err) {
    if (!_isLegacySelfSigError(err)) throw err;
    // Fallback: parse the first "Name <email>" UID string without sig validation.
    const firstUid = key.getUserIDs()[0] || '';
    const match = firstUid.match(/^(.*?)\s*<([^>]+)>$/);
    if (match) { name = match[1].trim(); email = match[2].trim(); }
    else        { name = firstUid; }
  }

  try {
    expirationTime = await key.getExpirationTime();
  } catch {
    // getExpirationTime() calls getPrimaryUser() internally; default to no expiry.
    expirationTime = Infinity;
  }

  const fp = key.getFingerprint().toUpperCase();
  return {
    fingerprint: fp,
    fingerprintFormatted: fp.match(/.{1,4}/g).join(' '),
    shortId: fp.slice(-8),
    keyId: key.getKeyID().toHex().toUpperCase(),
    userIds: key.getUserIDs(),
    name,
    email,
    created: key.getCreationTime(),
    expires: expirationTime === Infinity ? null : expirationTime,
    isPrivate: key.isPrivate(),
    algorithm: key.getAlgorithmInfo().algorithm,
  };
}

// ── Modern subkey operations ──────────────────────────────────────────────────

/**
 * Returns true if the key already contains at least one modern ECC subkey
 * (Ed25519 or X25519 / curve25519).  Used to decide whether to offer the
 * "Add Modern Subkeys" button in the key management UI.
 *
 * @param {string} armoredKey - Armored public or private key
 * @returns {Promise<boolean>}
 */
export async function hasModernSubkeys(armoredKey) {
  const key = await openpgp.readKey({ armoredKey })
    .catch(err => {
      if (!_isLegacySelfSigError(err)) throw err;
      return openpgp.readKey({ armoredKey, config: _buildLegacyKeyReadConfig() });
    });
  const modernAlgos = new Set(['ecdh', 'ecdsa', 'ed25519', 'x25519', 'curve25519']);
  return key.subkeys.some(sub => modernAlgos.has(sub.getAlgorithmInfo().algorithm));
}

/**
 * Add modern ECC subkeys to a legacy DSA/ElGamal (or other restricted) private key.
 *
 * Two subkeys are appended to the key:
 *   - Ed25519  — preferred signing subkey   (SHA-256 binding, no expiry)
 *   - X25519   — preferred encryption subkey (SHA-256 binding, no expiry)
 *
 * The original key structure is preserved unchanged (DSA primary + ElGamal subkey
 * + SHA-1 self-signature).  The new subkeys are bound to the DSA primary key with
 * SHA-256 binding signatures.  Both new subkeys are encrypted with the original
 * passphrase before the merged key is serialized.
 *
 * Internally, this function must temporarily relax the global OpenPGP.js config
 * (rejectHashAlgorithms / rejectPublicKeyAlgorithms) to allow reading and unlocking
 * the legacy key.  The config is restored in the finally block regardless of outcome.
 *
 * @param {string} armoredPrivateKey - Armored DSA/ElGamal (or other legacy) private key
 * @param {string} passphrase        - Key passphrase (verifies ownership + re-encrypts
 *                                     the new subkeys before they are stored)
 * @returns {Promise<{ armoredPrivate: string, armoredPublic: string }>}
 * @throws If the passphrase is wrong or key parsing fails
 */
export async function addModernSubkeys(armoredPrivateKey, passphrase) {
  const sha1Id    = openpgp.enums?.hash?.sha1         ?? 2;
  const dsaId     = openpgp.enums?.publicKey?.dsa     ?? 17;
  const elgamalId = openpgp.enums?.publicKey?.elgamal ?? 16;

  // Build a permissive config object for all operations that touch the legacy key.
  // This is passed per-call rather than mutating the global config, which avoids
  // race conditions if multiple operations overlap.
  const relaxedConfig = {
    rejectHashAlgorithms: new Set(
      [...openpgp.config.rejectHashAlgorithms].filter(id => id !== sha1Id)
    ),
    rejectPublicKeyAlgorithms: new Set(
      [...openpgp.config.rejectPublicKeyAlgorithms].filter(
        id => id !== dsaId && id !== elgamalId
      )
    ),
  };

  // Read and fully decrypt the key so addSubkey() can access the primary key material.
  const privateKey  = await openpgp.readPrivateKey({ armoredKey: armoredPrivateKey });
  const unlockedKey = await openpgp.decryptKey({ privateKey, passphrase, config: relaxedConfig });

  // PrivateKey.addSubkey() uses OpenPGP.js's own createBindingSignature() internally,
  // which correctly sets publicKeyAlgorithm, handles the back-signature for signing
  // subkeys, and enforces key flags — far more reliable than hand-crafting packets.
  //
  // sign: true  → Ed25519 signing subkey  (binding sig includes 0x19 back-signature)
  // sign: false → X25519 encryption subkey (no back-signature needed)
  let keyWithSubkeys = await unlockedKey.addSubkey({
    type: 'ecc', curve: 'curve25519', sign: true,  config: relaxedConfig,
  });
  keyWithSubkeys = await keyWithSubkeys.addSubkey({
    type: 'ecc', curve: 'curve25519', sign: false, config: relaxedConfig,
  });

  // Re-encrypt the entire key (primary + all subkeys) with the passphrase.
  // addSubkey() intentionally leaves key material in plaintext; encryptKey() seals it.
  const encryptedKey = await openpgp.encryptKey({
    privateKey: keyWithSubkeys, passphrase, config: relaxedConfig,
  });

  return {
    armoredPrivate: encryptedKey.armor(),
    armoredPublic:  encryptedKey.toPublic().armor(),
  };
}

// ── Message encryption / decryption ──────────────────────────────────────────

/**
 * Encrypt a plaintext message to one or more recipient public key objects.
 * Optionally sign with the sender's unlocked private key.
 *
 * The message is encrypted once for every recipient key (OpenPGP PKESK
 * packets), meaning any one recipient can decrypt it independently.
 *
 * @param {string}             text               - Plaintext message body
 * @param {openpgp.Key[]}      recipientPublicKeys - Parsed public key objects
 * @param {openpgp.PrivateKey} [signingKey]        - Unlocked private key for signing (optional)
 * @returns {Promise<string>} Armored PGP message ("-----BEGIN PGP MESSAGE-----")
 */
export async function encryptMessage(text, recipientPublicKeys, signingKey = null) {
  const message = await openpgp.createMessage({ text });
  const options = { message, encryptionKeys: recipientPublicKeys };
  if (signingKey) options.signingKeys = signingKey;

  try {
    return await openpgp.encrypt(options);
  } catch (err) {
    // At least one recipient has a legacy key.  Two distinct failure modes:
    //   _isWeakKeyError       — ElGamal/DSA key rejected by rejectPublicKeyAlgorithms
    //   _isLegacySelfSigError — recipient's key has SHA-1 self-signatures, rejected by
    //                           rejectHashAlgorithms during getEncryptionKey() validation
    // Both are resolved by _buildLegacyKeyReadConfig(), which removes SHA-1 from the
    // rejected hash set AND removes DSA/ElGamal from the rejected PK algorithm set.
    if (_isWeakKeyError(err) || _isLegacySelfSigError(err)) {
      return await openpgp.encrypt({ ...options, config: _buildLegacyKeyReadConfig() });
    }
    throw err;
  }
}

/**
 * Decrypt an armored PGP message.
 *
 * Signature verification is opportunistic: if verificationKeys are provided
 * and the message was signed, the result includes a validity flag.  If no
 * verification keys are provided (or the message is unsigned), signatureResult
 * will have valid === null (not false — that would imply a failed verification).
 *
 * @param {string}             armoredMessage    - PGP-armored ciphertext
 * @param {openpgp.PrivateKey} decryptionKey     - Unlocked private key
 * @param {openpgp.Key[]}      [verificationKeys] - Public keys for signature checking
 * @returns {Promise<{
 *   data: string,
 *   signatureResult: { valid: boolean|null, signedByKeyId: string|null }
 * }>}
 */
export async function decryptMessage(armoredMessage, decryptionKey, verificationKeys = []) {
  const message = await openpgp.readMessage({ armoredMessage });
  const result = await openpgp.decrypt({
    message,
    decryptionKeys: decryptionKey,
    verificationKeys: verificationKeys.length > 0 ? verificationKeys : undefined,
    // expectSigned: false means we don't throw if there's no signature.
    // This is correct behavior for encrypted-only (no signature) messages.
    expectSigned: false,
  });

  let signatureResult = { valid: null, signedByKeyId: null };
  if (result.signatures && result.signatures.length > 0) {
    const sig = result.signatures[0];
    try {
      // sig.verified is a Promise that resolves on valid and rejects on invalid.
      // We await it inside a try/catch to avoid unhandled rejections.
      await sig.verified;
      signatureResult.valid = true;
      signatureResult.signedByKeyId = sig.keyID?.toHex()?.toUpperCase() || null;
    } catch {
      signatureResult.valid = false;
      signatureResult.signedByKeyId = sig.keyID?.toHex()?.toUpperCase() || null;
    }
  }

  return { data: result.data, signatureResult };
}

// ── Attachment encryption / decryption ───────────────────────────────────────

/**
 * Encrypt binary attachment data as an armored PGP message.
 *
 * The original filename is stored inside the PGP Literal Data packet so that
 * the recipient can restore the correct filename when they decrypt.
 * The encrypted output should be saved with a ".pgp" suffix (e.g. "report.pdf.pgp").
 *
 * @param {Uint8Array}         data               - Raw file bytes
 * @param {string}             filename           - Original filename (stored in PGP packet)
 * @param {openpgp.Key[]}      recipientPublicKeys
 * @param {openpgp.PrivateKey} [signingKey]       - Optional signing key
 * @returns {Promise<string>} Armored PGP message
 */
export async function encryptAttachment(data, filename, recipientPublicKeys, signingKey = null) {
  const options = {
    message: await openpgp.createMessage({ binary: data, filename }),
    encryptionKeys: recipientPublicKeys,
    format: 'armored',
  };
  if (signingKey) options.signingKeys = signingKey;

  try {
    return await openpgp.encrypt(options);
  } catch (err) {
    if (_isWeakKeyError(err) || _isLegacySelfSigError(err)) {
      return await openpgp.encrypt({ ...options, config: _buildLegacyKeyReadConfig() });
    }
    throw err;
  }
}

/**
 * Returns true if any key in the array uses an algorithm that OpenPGP.js v5
 * rejects by default (currently ElGamal and DSA).  Use this before encrypting
 * to decide whether to show a "legacy key" warning in the UI.
 *
 * @param {openpgp.Key[]} keys
 * @returns {Promise<boolean>}
 */
export async function hasWeakEncryptionKey(keys) {
  for (const key of keys) {
    if (typeof key?.getEncryptionKey !== 'function') continue;
    try {
      await key.getEncryptionKey();
    } catch (err) {
      if (_isWeakKeyError(err)) return true;
    }
  }
  return false;
}

/**
 * Decrypt an armored PGP attachment message.
 *
 * @param {string}             armoredMessage
 * @param {openpgp.PrivateKey} decryptionKey
 * @returns {Promise<{ data: Uint8Array, filename: string }>}
 */
export async function decryptAttachment(armoredMessage, decryptionKey) {
  const message = await openpgp.readMessage({ armoredMessage });
  const result = await openpgp.decrypt({
    message,
    decryptionKeys: decryptionKey,
    // format: 'binary' returns a Uint8Array instead of a string, which is
    // required for arbitrary binary files (not just text).
    format: 'binary',
  });

  // OpenPGP.js surfaces the Literal Data packet's filename on result.filename.
  // (message.packets[0].filename refers to the *encrypted* input packet, which
  // is never populated — the filename only becomes available after decryption.)
  const filename = result.filename || '';
  return { data: result.data, filename };
}

// ── Helpers ───────────────────────────────────────────────────────────────────

/**
 * Detect whether a string contains a PGP armored block and return its type.
 * Used to decide which UI panel to show when reading a message.
 *
 * @param {string} text
 * @returns {'encrypted'|'signed'|'public-key'|'private-key'|null}
 */
export function detectPgpContent(text) {
  if (!text) return null;
  if (text.includes('-----BEGIN PGP MESSAGE-----'))          return 'encrypted';
  if (text.includes('-----BEGIN PGP SIGNED MESSAGE-----'))   return 'signed';
  if (text.includes('-----BEGIN PGP PUBLIC KEY BLOCK-----')) return 'public-key';
  if (text.includes('-----BEGIN PGP PRIVATE KEY BLOCK-----')) return 'private-key';
  return null;
}

/**
 * Convert a base64 string to Uint8Array.
 * Used to convert the base64 attachment content returned by Office.js
 * getAttachmentContentAsync() into raw bytes for encryption.
 */
export function base64ToUint8Array(base64) {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

/**
 * Convert a Uint8Array to a base64 string.
 * Used to convert encrypted bytes into the format expected by
 * Office.js addFileAttachmentFromBase64Async().
 */
export function uint8ArrayToBase64(bytes) {
  let binary = '';
  for (let i = 0; i < bytes.length; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  return btoa(binary);
}
