import { describe, it, expect, beforeEach, beforeAll, vi } from 'vitest';
import { generateKeyPair, encryptMessage, decryptMessage, unlockPrivateKey } from '../web/js/pgp/pgp-core.js';

// keyring.js persists through key-storage.js (Office-backed) — mock it with an
// in-memory object so tests exercise real key parsing (pgp-core.js) without
// touching Office.js.
vi.mock('../web/js/pgp/key-storage.js', () => ({
  getKeyring: vi.fn(() => ({})),
  saveKeyring: vi.fn(async () => {}),
  estimateStorageUsage: vi.fn(() => 100),
  STORAGE_LIMIT_BYTES: 32768,
}));

// NOTE: keyring.js and key-storage.js are imported statically (not re-imported
// via resetModules()+dynamic import per test) because keyring.js transitively
// imports pgp-core.js — resetting modules would load a *second*, separate
// instance of the vendored openpgp.min.mjs library, and openpgp.Key objects
// from one instance aren't usable with encrypt/decrypt calls on another
// (surfaces as a confusing "Unknown curve" error). keyring.js itself has no
// persistent module-level state to reset, so this is safe — only the
// key-storage.js *mock's* return values need resetting between tests.
import * as keyring from '../web/js/pgp/keyring.js';
import * as keyStorageMock from '../web/js/pgp/key-storage.js';

let alice;

beforeAll(async () => {
  alice = await generateKeyPair('Alice Example', 'alice@example.com', 'correct horse battery staple');
}, 30000);

beforeEach(() => {
  vi.clearAllMocks();
  keyStorageMock.getKeyring.mockReturnValue({});
  keyStorageMock.estimateStorageUsage.mockReturnValue(100);
});

describe('addContactKey / getContactKey / hasContactKey / removeContactKey', () => {
  it('accepts a valid armored public key and stores it lowercase-keyed', async () => {
    const result = await keyring.addContactKey('Alice@Example.com', alice.publicKey);

    expect(result.info.email).toBe('alice@example.com');
    expect(result.storageWarning).toBe(false);
    expect(keyStorageMock.saveKeyring).toHaveBeenCalledWith({ 'alice@example.com': alice.publicKey });
  });

  it('rejects garbage input instead of storing it', async () => {
    await expect(keyring.addContactKey('bob@example.com', 'not a real key')).rejects.toThrow();
    expect(keyStorageMock.saveKeyring).not.toHaveBeenCalled();
  });

  it('refuses to store a private key in the shared keyring', async () => {
    await expect(keyring.addContactKey('alice@example.com', alice.privateKey)).rejects.toThrow(
      /Refusing to store a private key/
    );
  });

  it('flags a storage warning when usage is above 80% of the limit', async () => {
    keyStorageMock.estimateStorageUsage.mockReturnValue(Math.round(32768 * 0.85));
    const result = await keyring.addContactKey('alice@example.com', alice.publicKey);
    expect(result.storageWarning).toBe(true);
  });

  it('getContactKey / hasContactKey / removeContactKey round-trip via the in-memory keyring', async () => {
    let store = {};
    keyStorageMock.getKeyring.mockImplementation(() => store);
    keyStorageMock.saveKeyring.mockImplementation(async (next) => { store = next; });

    expect(keyring.hasContactKey('alice@example.com')).toBe(false);
    expect(keyring.getContactKey('alice@example.com')).toBeNull();

    await keyring.addContactKey('alice@example.com', alice.publicKey);
    expect(keyring.hasContactKey('ALICE@EXAMPLE.COM')).toBe(true);
    expect(keyring.getContactKey('alice@example.com')).toBe(alice.publicKey);

    await keyring.removeContactKey('alice@example.com');
    expect(keyring.hasContactKey('alice@example.com')).toBe(false);
  });
});

describe('listContactKeys', () => {
  it('returns parsed metadata per contact, sorted by email', async () => {
    keyStorageMock.getKeyring.mockReturnValue({
      'zed@example.com': alice.publicKey,
      'alice@example.com': alice.publicKey,
    });

    const results = await keyring.listContactKeys();

    expect(results.map(r => r.email)).toEqual(['alice@example.com', 'zed@example.com']);
    expect(results[0].info.email).toBe('alice@example.com');
  });

  it('includes an error entry for a key that fails to parse instead of throwing', async () => {
    keyStorageMock.getKeyring.mockReturnValue({ 'broken@example.com': 'not a real key' });

    const results = await keyring.listContactKeys();

    expect(results).toHaveLength(1);
    expect(results[0].error).toMatch(/Could not parse key/);
  });
});

describe('getContactKeyObject', () => {
  it('returns a usable openpgp.Key that can encrypt to the contact', async () => {
    keyStorageMock.getKeyring.mockReturnValue({ 'alice@example.com': alice.publicKey });

    const keyObject = await keyring.getContactKeyObject('alice@example.com');
    expect(keyObject).not.toBeNull();

    const armored = await encryptMessage('hello contact', [keyObject]);
    const unlockedAlice = await unlockPrivateKey(alice.privateKey, 'correct horse battery staple');
    const { data } = await decryptMessage(armored, unlockedAlice);
    expect(data).toBe('hello contact');
  });

  it('returns null when the contact is not in the keyring', async () => {
    keyStorageMock.getKeyring.mockReturnValue({});
    expect(await keyring.getContactKeyObject('nobody@example.com')).toBeNull();
  });

  it('returns null when the stored key is unparseable', async () => {
    keyStorageMock.getKeyring.mockReturnValue({ 'broken@example.com': 'not a real key' });
    expect(await keyring.getContactKeyObject('broken@example.com')).toBeNull();
  });
});

describe('getKeyringStorageInfo', () => {
  it('reports count, byte usage, and near-limit status', () => {
    keyStorageMock.getKeyring.mockReturnValue({ 'alice@example.com': alice.publicKey });
    keyStorageMock.estimateStorageUsage.mockReturnValue(Math.round(32768 * 0.9));

    const info = keyring.getKeyringStorageInfo();

    expect(info.count).toBe(1);
    expect(info.totalBytes).toBe(Math.round(32768 * 0.9));
    expect(info.remainingBytes).toBe(32768 - Math.round(32768 * 0.9));
    expect(info.nearLimit).toBe(true);
  });
});
