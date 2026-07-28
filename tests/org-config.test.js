import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';

// org-config.js imports getOrgOverride from key-storage.js (Office-backed) and
// fetchFromWKD/fetchFromVKS from key-discovery.js — mock both so tests never
// touch Office.js or the network directly. vi.mock is hoisted, so it applies
// even to the dynamic re-imports below (each with a fresh mock instance,
// since resetModules() reruns the factory too).
vi.mock('../web/js/pgp/key-storage.js', () => ({
  getOrgOverride: vi.fn(() => null),
}));

vi.mock('../web/js/pgp/key-discovery.js', () => ({
  fetchFromWKD: vi.fn(),
  fetchFromVKS: vi.fn(),
}));

let orgConfig;
let keyStorageMock;
let keyDiscoveryMock;

beforeEach(async () => {
  vi.resetModules();
  vi.clearAllMocks();
  orgConfig = await import('../web/js/pgp/org-config.js');
  keyStorageMock = await import('../web/js/pgp/key-storage.js');
  keyDiscoveryMock = await import('../web/js/pgp/key-discovery.js');
  keyStorageMock.getOrgOverride.mockReturnValue(null);
  global.fetch = vi.fn();
});

afterEach(() => {
  delete global.fetch;
});

describe('loadOrgConfig — manual override', () => {
  it('uses the override without going to the network when present', async () => {
    keyStorageMock.getOrgOverride.mockReturnValue({
      companyKeyEnabled: true,
      companyKeyEmails: ['legal@acme.com'],
    });

    const config = await orgConfig.loadOrgConfig('alice@acme.com');

    expect(config.companyKeyEnabled).toBe(true);
    expect(config.companyKeyEmails).toEqual(['legal@acme.com']);
    expect(global.fetch).not.toHaveBeenCalled();
  });
});

describe('loadOrgConfig — well-known URL fetch', () => {
  it('fetches the primary URL and merges validated fields, including companyDecryptedExtensionPrefix', async () => {
    global.fetch.mockResolvedValueOnce({
      ok: true,
      json: async () => ({
        companyKeyEnabled: true,
        companyKeyRequired: true,
        companyKeyEmails: ['legal@acme.com'],
        hideSupportButton: true,
        companyDecryptedExtensionPrefix: 'pgpDecrypted',
      }),
    });

    const config = await orgConfig.loadOrgConfig('alice@acme.com');

    expect(config).toEqual({
      companyKeyEnabled: true,
      companyKeyRequired: true,
      companyKeyEmails: ['legal@acme.com'],
      hideSupportButton: true,
      companyDecryptedExtensionPrefix: 'pgpDecrypted',
    });
    expect(global.fetch).toHaveBeenCalledTimes(1);
    expect(global.fetch.mock.calls[0][0]).toBe(
      'https://acme.com/.well-known/pgp-for-outlook-addin/company-config.json'
    );
  });

  it('ignores fields of the wrong type and keeps their defaults', async () => {
    global.fetch.mockResolvedValueOnce({
      ok: true,
      json: async () => ({ companyDecryptedExtensionPrefix: 123, companyKeyEnabled: 'yes' }),
    });

    const config = await orgConfig.loadOrgConfig('alice@acme.com');

    expect(config.companyDecryptedExtensionPrefix).toBe('');
    expect(config.companyKeyEnabled).toBe(false);
  });

  it('falls back to the openpgpkey subdomain when the primary URL fails', async () => {
    global.fetch
      .mockRejectedValueOnce(new Error('network error'))
      .mockResolvedValueOnce({ ok: true, json: async () => ({ companyKeyEnabled: true }) });

    const config = await orgConfig.loadOrgConfig('alice@acme.com');

    expect(config.companyKeyEnabled).toBe(true);
    expect(global.fetch).toHaveBeenCalledTimes(2);
    expect(global.fetch.mock.calls[1][0]).toBe(
      'https://openpgpkey.acme.com/.well-known/pgp-for-outlook-addin/company-config.json'
    );
  });

  it('returns defaults when both candidate URLs fail', async () => {
    global.fetch.mockRejectedValue(new Error('network error'));

    const config = await orgConfig.loadOrgConfig('alice@acme.com');

    expect(config.companyKeyEnabled).toBe(false);
    expect(config.companyDecryptedExtensionPrefix).toBe('');
  });

  it('returns defaults and skips fetching entirely when the email has no domain', async () => {
    const config = await orgConfig.loadOrgConfig('');

    expect(config.companyKeyEnabled).toBe(false);
    expect(global.fetch).not.toHaveBeenCalled();
  });
});

describe('accessors', () => {
  it('getOrgConfig() returns defaults before loadOrgConfig() has ever run', () => {
    expect(orgConfig.getOrgConfig()).toMatchObject({
      companyKeyEnabled: false,
      companyDecryptedExtensionPrefix: '',
    });
  });

  it('read the cached config after loadOrgConfig() resolves', async () => {
    global.fetch.mockResolvedValueOnce({
      ok: true,
      json: async () => ({
        companyKeyEnabled: true,
        companyKeyRequired: true,
        companyKeyEmails: ['legal@acme.com'],
        hideSupportButton: true,
        companyDecryptedExtensionPrefix: 'pgpDecrypted',
      }),
    });
    await orgConfig.loadOrgConfig('alice@acme.com');

    expect(orgConfig.isCompanyKeyEnabled()).toBe(true);
    expect(orgConfig.isCompanyKeyRequired()).toBe(true);
    expect(orgConfig.getCompanyKeyEmails()).toEqual(['legal@acme.com']);
    expect(orgConfig.isSupportButtonHidden()).toBe(true);
    expect(orgConfig.getDecryptedExtensionPrefix()).toBe('pgpDecrypted');
  });

  it('clearOrgConfigCache() resets accessors back to defaults', async () => {
    global.fetch.mockResolvedValueOnce({ ok: true, json: async () => ({ companyKeyEnabled: true }) });
    await orgConfig.loadOrgConfig('alice@acme.com');
    expect(orgConfig.isCompanyKeyEnabled()).toBe(true);

    orgConfig.clearOrgConfigCache();

    expect(orgConfig.isCompanyKeyEnabled()).toBe(false);
  });
});

describe('fetchCompanyKeys / getMissingCompanyKeyEmails', () => {
  beforeEach(() => {
    keyStorageMock.getOrgOverride.mockReturnValue({
      companyKeyEnabled: true,
      companyKeyEmails: ['legal@acme.com', 'missing@acme.com'],
    });
  });

  it('resolves via WKD, falls back to VKS on a miss, and reports unresolved emails', async () => {
    await orgConfig.loadOrgConfig('alice@acme.com');

    keyDiscoveryMock.fetchFromWKD.mockImplementation(async (email) => {
      if (email === 'legal@acme.com') return { key: { fake: 'legal-key' } };
      throw new Error('not found via WKD');
    });
    keyDiscoveryMock.fetchFromVKS.mockImplementation(async () => {
      throw new Error('not found via VKS either');
    });

    const keys = await orgConfig.fetchCompanyKeys();
    expect(keys).toEqual([{ email: 'legal@acme.com', key: { fake: 'legal-key' } }]);

    const missing = await orgConfig.getMissingCompanyKeyEmails();
    expect(missing).toEqual(['missing@acme.com']);
  });

  it('caches the resolved keys and does not re-fetch on a second call', async () => {
    await orgConfig.loadOrgConfig('alice@acme.com');
    keyDiscoveryMock.fetchFromWKD.mockResolvedValue({ key: { fake: 'key' } });

    await orgConfig.fetchCompanyKeys();
    const callsAfterFirst = keyDiscoveryMock.fetchFromWKD.mock.calls.length;
    await orgConfig.fetchCompanyKeys();

    expect(keyDiscoveryMock.fetchFromWKD.mock.calls.length).toBe(callsAfterFirst);
  });

  it('returns an empty array when no company emails are configured', async () => {
    keyStorageMock.getOrgOverride.mockReturnValue({ companyKeyEnabled: false, companyKeyEmails: [] });
    await orgConfig.loadOrgConfig('alice@acme.com');

    const keys = await orgConfig.fetchCompanyKeys();
    expect(keys).toEqual([]);
  });
});
