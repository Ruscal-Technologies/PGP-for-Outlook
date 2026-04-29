/**
 * org-config.js
 * Loads organization-level PGP configuration.
 *
 * Configuration is fetched at runtime from a well-known URL on the user's
 * email domain.  IT admins publish a static JSON file at one of:
 *
 *   Primary:  https://<org-domain>/.well-known/pgp-for-outlook-addin/company-config.json
 *   Fallback: https://openpgpkey.<org-domain>/.well-known/pgp-for-outlook-addin/company-config.json
 *
 * The fallback lets organizations that already operate a WKD server
 * (https://openpgpkey.<domain>) co-locate the add-in config alongside
 * their key-discovery infrastructure without needing to touch the apex domain.
 *
 * Expected schema:
 * {
 *   "companyKeyEnabled":  true,           // whether the feature is active
 *   "companyKeyRequired": false,          // if true, users cannot opt out
 *   "companyKeyEmails":   ["legal@example.com"]  // addresses to look up via WKD/VKS
 * }
 *
 * A manual override stored in roaming settings takes precedence over the
 * well-known URL.  This lets an IT admin configure the add-in for users
 * who cannot host the well-known file, by having them (or a PowerShell
 * script) set the override once via the Key Management settings panel.
 */

import { getOrgOverride } from './key-storage.js';
import { fetchFromWKD, fetchFromVKS } from './key-discovery.js';

const DEFAULT_CONFIG = Object.freeze({
  companyKeyEnabled:  false,
  companyKeyRequired: false,
  companyKeyEmails:   [],
  hideSupportButton:  false,
});

// Config and company keys are cached for the lifetime of the task pane session.
// Call clearOrgConfigCache() after changing the override to force a reload.
let _cachedConfig     = null;
let _cachedCompanyKeys = null; // [{ email, key: openpgp.Key }]

// ── Config loading ────────────────────────────────────────────────────────────

/**
 * Load and cache the org config for the current user.
 * Call once at add-in startup (after Office.onReady).
 *
 * @param {string} userEmail - The signed-in user's email address
 * @returns {object} The effective config object
 */
export async function loadOrgConfig(userEmail) {
  // 1. Roaming settings override (highest priority — set by IT admin or user)
  const override = getOrgOverride();
  if (override && typeof override === 'object') {
    _cachedConfig = { ...DEFAULT_CONFIG, ...override };
    return _cachedConfig;
  }

  // 2. Well-known URL — primary on the apex domain, fallback on openpgpkey subdomain
  const domain = userEmail?.split('@')[1];
  if (domain) {
    const candidates = [
      `https://${domain}/.well-known/pgp-for-outlook-addin/company-config.json`,
      `https://openpgpkey.${domain}/.well-known/pgp-for-outlook-addin/company-config.json`,
    ];
    for (const url of candidates) {
      try {
        const response = await fetch(url, {
          // Fail fast — if the file doesn't exist we try the next candidate
          // rather than leaving the user waiting for a network timeout.
          signal: AbortSignal.timeout(5000),
          // No credentials — this file is intentionally public.  Never add
          // 'include' here; that would send the user's cookies to the domain.
        });
        if (response.ok) {
          const json = await response.json();
          // Validate field types before merging — the server (or a network
          // attacker on a downgraded connection) could supply unexpected types
          // that would crash consumers (e.g. a string where an array is expected).
          _cachedConfig = {
            ...DEFAULT_CONFIG,
            ...(typeof json.companyKeyEnabled  === 'boolean' && { companyKeyEnabled:  json.companyKeyEnabled  }),
            ...(typeof json.companyKeyRequired === 'boolean' && { companyKeyRequired: json.companyKeyRequired }),
            ...(Array.isArray(json.companyKeyEmails)          && { companyKeyEmails:   json.companyKeyEmails   }),
            ...(typeof json.hideSupportButton  === 'boolean' && { hideSupportButton:  json.hideSupportButton  }),
          };
          return _cachedConfig;
        }
      } catch (e) {
        // This candidate isn't available — try the next one.
        console.info(`company-config.json not found at ${url}:`, e.message);
      }
    }
  }

  // 3. Defaults (company key feature disabled)
  _cachedConfig = { ...DEFAULT_CONFIG };
  return _cachedConfig;
}

// ── Accessors ─────────────────────────────────────────────────────────────────

export function getOrgConfig() {
  return _cachedConfig ?? { ...DEFAULT_CONFIG };
}

export function isCompanyKeyEnabled() {
  return getOrgConfig().companyKeyEnabled === true;
}

export function isCompanyKeyRequired() {
  return getOrgConfig().companyKeyRequired === true;
}

export function getCompanyKeyEmails() {
  return getOrgConfig().companyKeyEmails ?? [];
}

export function isSupportButtonHidden() {
  return getOrgConfig().hideSupportButton === true;
}

// ── Company key fetching ──────────────────────────────────────────────────────

/**
 * Fetch and cache the company public key objects.
 * Tries WKD first, then falls back to VKS for each configured email.
 *
 * @returns {Array<{ email: string, key: openpgp.Key }>}
 */
export async function fetchCompanyKeys() {
  if (_cachedCompanyKeys !== null) return _cachedCompanyKeys;

  const emails = getCompanyKeyEmails();
  if (emails.length === 0) {
    _cachedCompanyKeys = [];
    return _cachedCompanyKeys;
  }

  const results = await Promise.allSettled(
    emails.map(async (email) => {
      let result = null;
      try { result = await fetchFromWKD(email); } catch { /* ignore */ }
      if (!result) {
        try { result = await fetchFromVKS(email); } catch { /* ignore */ }
      }
      if (result) return { email, key: result.key };
      throw new Error(`No key found for company address: ${email}`);
    })
  );

  _cachedCompanyKeys = results
    .filter(r => r.status === 'fulfilled')
    .map(r => r.value);

  const failed = results
    .filter(r => r.status === 'rejected')
    .map(r => r.reason?.message);
  if (failed.length > 0) {
    console.warn('Could not fetch some company keys:', failed);
  }

  return _cachedCompanyKeys;
}

/**
 * Return the names of any company key emails that could not be resolved.
 */
export async function getMissingCompanyKeyEmails() {
  const emails = getCompanyKeyEmails();
  const keys = await fetchCompanyKeys();
  const resolved = new Set(keys.map(k => k.email.toLowerCase()));
  return emails.filter(e => !resolved.has(e.toLowerCase()));
}

/**
 * Clear caches — call after config changes.
 */
export function clearOrgConfigCache() {
  _cachedConfig     = null;
  _cachedCompanyKeys = null;
}
