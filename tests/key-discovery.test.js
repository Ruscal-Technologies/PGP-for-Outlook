import { describe, it, expect, beforeEach, afterEach, beforeAll, vi } from 'vitest';
import { generateKeyPair } from '../web/js/pgp/pgp-core.js';
import * as openpgp from '../web/js/openpgp.min.mjs';

// keyring.js is a separate module — safe to fully mock via vi.mock().
vi.mock('../web/js/pgp/keyring.js', () => ({
  getContactKeyObject: vi.fn(),
}));

// wkd.js's WKD class does its own fetch()/crypto.subtle work that we don't
// want to exercise here; mock the class so fetchFromWKD's *own* logic (empty
// buffer -> null, binary -> parsed key) is what's under test, not the WKD
// protocol implementation itself.
vi.mock('../web/js/wkd.js', () => {
  const lookup = vi.fn();
  return {
    default: vi.fn(() => ({ lookup })),
    __lookup: lookup, // exposed so tests can configure/inspect the shared mock
  };
});

import * as keyDiscovery from '../web/js/pgp/key-discovery.js';
import * as keyringMock from '../web/js/pgp/keyring.js';
import * as wkdMock from '../web/js/wkd.js';

let alice;
let aliceBinaryKey;

beforeAll(async () => {
  alice = await generateKeyPair('Alice Example', 'alice@example.com', 'correct horse battery staple');
  const aliceKeyObj = await openpgp.readKey({ armoredKey: alice.publicKey });
  aliceBinaryKey = aliceKeyObj.write();
}, 30000);

beforeEach(() => {
  vi.clearAllMocks();
});

afterEach(() => {
  delete global.fetch;
});

describe('discoverKey — source precedence', () => {
  it('returns the local keyring hit first, without trying WKD/VKS', async () => {
    keyringMock.getContactKeyObject.mockResolvedValue({ fake: 'local-key' });

    const result = await keyDiscovery.discoverKey('Alice@Example.com');

    expect(result).toEqual({
      key: { fake: 'local-key' },
      status: keyDiscovery.KeyStatus.FOUND_LOCAL,
      source: 'Local keyring',
      armoredKey: null,
    });
    expect(keyringMock.getContactKeyObject).toHaveBeenCalledWith('alice@example.com');
    expect(wkdMock.__lookup).not.toHaveBeenCalled();
  });

  it('falls through to WKD when there is no local key', async () => {
    keyringMock.getContactKeyObject.mockResolvedValue(null);
    wkdMock.__lookup.mockResolvedValue(aliceBinaryKey);

    const result = await keyDiscovery.discoverKey('bob@example.com');

    expect(result.status).toBe(keyDiscovery.KeyStatus.FOUND_WKD);
    expect(result.source).toBe('WKD');
    expect(result.armoredKey).toContain('-----BEGIN PGP PUBLIC KEY BLOCK-----');
  });

  it('falls through to VKS when WKD returns an empty buffer (no key found)', async () => {
    keyringMock.getContactKeyObject.mockResolvedValue(null);
    wkdMock.__lookup.mockResolvedValue(new Uint8Array());
    global.fetch = vi.fn().mockResolvedValue({ status: 200, text: async () => alice.publicKey });

    const result = await keyDiscovery.discoverKey('carol@example.com');

    expect(result.status).toBe(keyDiscovery.KeyStatus.FOUND_VKS);
    expect(result.source).toBe('keys.openpgp.org');
  });

  it('returns NOT_FOUND when all three sources miss', async () => {
    keyringMock.getContactKeyObject.mockResolvedValue(null);
    wkdMock.__lookup.mockRejectedValue(new Error('WKD lookup failed'));
    global.fetch = vi.fn().mockResolvedValue({ status: 404, text: async () => '' });

    const result = await keyDiscovery.discoverKey('dave@example.com');

    expect(result).toEqual({ key: null, status: keyDiscovery.KeyStatus.NOT_FOUND, source: null, armoredKey: null });
  });

  it('continues to VKS when WKD throws instead of propagating the error', async () => {
    keyringMock.getContactKeyObject.mockResolvedValue(null);
    wkdMock.__lookup.mockRejectedValue(new Error('WKD network error'));
    global.fetch = vi.fn().mockResolvedValue({ status: 200, text: async () => alice.publicKey });

    const result = await keyDiscovery.discoverKey('erin@example.com');

    expect(result.status).toBe(keyDiscovery.KeyStatus.FOUND_VKS);
  });
});

describe('resolveRecipients', () => {
  it('aggregates results for multiple emails (mix of hit/miss)', async () => {
    keyringMock.getContactKeyObject.mockImplementation(async (email) => (
      email === 'alice@example.com' ? { fake: 'alice-key' } : null
    ));
    wkdMock.__lookup.mockResolvedValue(new Uint8Array());
    global.fetch = vi.fn().mockResolvedValue({ status: 404, text: async () => '' });

    const results = await keyDiscovery.resolveRecipients(['alice@example.com', 'bob@example.com']);

    expect(results).toEqual([
      {
        email: 'alice@example.com', key: { fake: 'alice-key' },
        status: keyDiscovery.KeyStatus.FOUND_LOCAL, source: 'Local keyring', armoredKey: null,
      },
      {
        email: 'bob@example.com', key: null,
        status: keyDiscovery.KeyStatus.NOT_FOUND, source: null, armoredKey: null,
      },
    ]);
  });
});

describe('fetchFromWKD (real implementation, mocked WKD class)', () => {
  it('returns a parsed key + armored text on a successful lookup', async () => {
    wkdMock.__lookup.mockResolvedValue(aliceBinaryKey);

    const result = await keyDiscovery.fetchFromWKD('alice@example.com');

    expect(result).not.toBeNull();
    expect(result.armoredKey).toContain('-----BEGIN PGP PUBLIC KEY BLOCK-----');
    expect(result.key).toBeTruthy();
  });

  it('returns null when lookup resolves with an empty buffer', async () => {
    wkdMock.__lookup.mockResolvedValue(new Uint8Array());
    expect(await keyDiscovery.fetchFromWKD('alice@example.com')).toBeNull();
  });
});

describe('fetchFromVKS (real implementation, mocked fetch)', () => {
  it('builds the correct keys.openpgp.org URL and parses a found key', async () => {
    global.fetch = vi.fn().mockResolvedValue({ status: 200, text: async () => alice.publicKey });

    const result = await keyDiscovery.fetchFromVKS('alice@example.com');

    expect(global.fetch).toHaveBeenCalledWith('https://keys.openpgp.org/vks/v1/by-email/alice%40example.com');
    expect(result.armoredKey).toBe(alice.publicKey);
    expect(result.key).toBeTruthy();
  });

  it('returns null when the server responds with a non-200 status', async () => {
    global.fetch = vi.fn().mockResolvedValue({ status: 404, text: async () => '' });
    expect(await keyDiscovery.fetchFromVKS('nobody@example.com')).toBeNull();
  });

  it('returns null when the response body is not a PGP key', async () => {
    global.fetch = vi.fn().mockResolvedValue({ status: 200, text: async () => 'not a key' });
    expect(await keyDiscovery.fetchFromVKS('nobody@example.com')).toBeNull();
  });
});
