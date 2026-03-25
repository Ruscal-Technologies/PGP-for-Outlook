/**
 * session-cache.js
 * In-memory session cache for the unlocked private key.
 *
 * WHY this exists:
 *   Unlocking a private key requires the passphrase.  Without a cache the user
 *   must re-enter it for every encrypt/decrypt/attachment operation, which is
 *   particularly painful when encrypting a message with several attachments.
 *
 * Security model:
 *   - The unlocked key object is held in a module-level variable and is NEVER
 *     written to localStorage, sessionStorage, IndexedDB, or any other
 *     persistent/serialisable storage.
 *   - It exists only in the JavaScript heap of the task pane's WebView.
 *     Closing the task pane, reloading the page, or ending the browser session
 *     destroys it automatically — no cleanup needed on the caller's side.
 *   - An inactivity timer clears the cache after CACHE_TIMEOUT_MS of no use.
 *     Every call to getSessionKey() resets the timer (counts as "use").
 *   - A manual clearSessionKey() call is exposed so the user can lock
 *     immediately via a Lock button in the UI.
 *
 * Scope:
 *   Each task pane (MessageCompose, MessageRead, KeyManagement) is a separate
 *   WebView/iframe with its own JS module scope.  The session cache is
 *   therefore per-pane.  This is intentional: a compromised compose pane
 *   cannot access an unlocked key loaded by the read pane.
 *
 * Timeout:
 *   Default is 15 minutes of inactivity (adjustable via CACHE_TIMEOUT_MS).
 *   "Inactivity" means no call to getSessionKey() — not no user interaction.
 */

export const CACHE_TIMEOUT_MS = 15 * 60 * 1000; // 15 minutes

/** @type {import('../openpgp.min.mjs').PrivateKey|null} */
let _cachedKey   = null;
let _cachedEmail = null;   // owner email, for display only
let _cachedFp    = null;   // fingerprint short ID, for display only
let _timer       = null;
const _onClearCbs = [];

// ── Write ─────────────────────────────────────────────────────────────────────

/**
 * Store an unlocked private key object in the session cache.
 *
 * @param {import('../openpgp.min.mjs').PrivateKey} unlockedKey - Result of openpgp.decryptKey()
 * @param {string} userEmail   - The key owner's email (for UI display only)
 * @param {string} [shortId]   - The key's 8-char short fingerprint ID (for UI display only)
 */
export function cacheSessionKey(unlockedKey, userEmail, shortId = '') {
  _cachedKey   = unlockedKey;
  _cachedEmail = userEmail;
  _cachedFp    = shortId;
  _resetTimer();
}

// ── Read ──────────────────────────────────────────────────────────────────────

/**
 * Retrieve the cached unlocked key, resetting the inactivity timer.
 * Returns null if no key is cached or if the cache has expired.
 *
 * @returns {import('../openpgp.min.mjs').PrivateKey|null}
 */
export function getSessionKey() {
  if (!_cachedKey) return null;
  _resetTimer(); // activity detected — push back the expiry
  return _cachedKey;
}

/** Returns true if a key is currently cached. */
export function hasSessionKey() {
  return _cachedKey !== null;
}

/** Returns the cached key owner's email address, or null. */
export function getSessionEmail() {
  return _cachedEmail;
}

/** Returns the cached key's short fingerprint ID, or null. */
export function getSessionShortId() {
  return _cachedFp;
}

// ── Clear ─────────────────────────────────────────────────────────────────────

/**
 * Immediately clear the session cache and notify all registered callbacks.
 * Call this when the user clicks "Lock", when a decryption operation
 * completes on a security-sensitive pane, etc.
 */
export function clearSessionKey() {
  _cachedKey   = null;
  _cachedEmail = null;
  _cachedFp    = null;
  if (_timer) {
    clearTimeout(_timer);
    _timer = null;
  }
  // Notify all listeners (typically a single UI update function per pane)
  for (const cb of _onClearCbs) {
    try { cb(); } catch { /* never let a UI callback break the cache clear */ }
  }
}

// ── Lifecycle callbacks ───────────────────────────────────────────────────────

/**
 * Register a function to be called whenever the session cache is cleared.
 * Typically used to update the lock/unlock status indicator in the UI.
 *
 * @param {() => void} callback
 */
export function onSessionCleared(callback) {
  _onClearCbs.push(callback);
}

// ── Internal ──────────────────────────────────────────────────────────────────

function _resetTimer() {
  if (_timer) clearTimeout(_timer);
  // When the timer fires we do a full clearSessionKey() to trigger UI callbacks
  _timer = setTimeout(clearSessionKey, CACHE_TIMEOUT_MS);
}
