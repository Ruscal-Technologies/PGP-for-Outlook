import { describe, it, expect, beforeEach, vi } from 'vitest';

// key-storage.js only ever touches Office.context.roamingSettings inside
// function bodies (never at module load time), so a plain in-memory fake
// assigned to global.Office before each test is sufficient — no real Office.js
// runtime is required.
function installFakeOffice() {
  const store = {};
  global.Office = {
    AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
    context: {
      roamingSettings: {
        get: (key) => store[key],
        set: (key, value) => { store[key] = value; },
        remove: (key) => { delete store[key]; },
        saveAsync: (callback) => {
          callback({ status: 'succeeded' });
        },
      },
    },
  };
  return store;
}

let keyStorage;

beforeEach(async () => {
  installFakeOffice();
  // key-storage.js reads `Office` lazily inside each function body — it has no
  // module-level state of its own, so re-importing after resetModules() just
  // keeps every test file's import graph isolated/consistent; a fresh fake
  // Office object per test is what actually isolates state here.
  vi.resetModules();
  keyStorage = await import('../web/js/pgp/key-storage.js');
});

describe('own key pair', () => {
  it('has no key pair before one is saved', () => {
    expect(keyStorage.hasKeyPair()).toBe(false);
    expect(keyStorage.getPrivateKey()).toBeNull();
    expect(keyStorage.getPublicKey()).toBeNull();
    expect(keyStorage.getKeyMetadata()).toBeNull();
  });

  it('saveKeyPair persists private key, public key, and metadata', async () => {
    const meta = { name: 'Alice', email: 'alice@example.com', fingerprint: 'ABCD', keyId: '1234' };
    await keyStorage.saveKeyPair('PRIVATE-ARMOR', 'PUBLIC-ARMOR', meta);

    expect(keyStorage.hasKeyPair()).toBe(true);
    expect(keyStorage.getPrivateKey()).toBe('PRIVATE-ARMOR');
    expect(keyStorage.getPublicKey()).toBe('PUBLIC-ARMOR');
    expect(keyStorage.getKeyMetadata()).toEqual(meta);
  });

  it('clearKeyPair removes private key, public key, and metadata', async () => {
    await keyStorage.saveKeyPair('PRIVATE-ARMOR', 'PUBLIC-ARMOR', { email: 'alice@example.com' });
    await keyStorage.clearKeyPair();

    expect(keyStorage.hasKeyPair()).toBe(false);
    expect(keyStorage.getPrivateKey()).toBeNull();
    expect(keyStorage.getPublicKey()).toBeNull();
    expect(keyStorage.getKeyMetadata()).toBeNull();
  });

  it('saveKeyPair rejects when it would exceed the storage limit', async () => {
    const hugeKey = 'x'.repeat(keyStorage.STORAGE_LIMIT_BYTES);
    await expect(
      keyStorage.saveKeyPair(hugeKey, 'PUBLIC-ARMOR', { email: 'alice@example.com' })
    ).rejects.toThrow(/Storage limit would be exceeded/);
    // Rejected save must not have partially written state.
    expect(keyStorage.hasKeyPair()).toBe(false);
  });

  it('saveKeyPair allows replacing an existing key pair with one of similar size', async () => {
    const meta = { email: 'alice@example.com' };
    await keyStorage.saveKeyPair('PRIVATE-1', 'PUBLIC-1', meta);
    await keyStorage.saveKeyPair('PRIVATE-2', 'PUBLIC-2', meta);
    expect(keyStorage.getPrivateKey()).toBe('PRIVATE-2');
  });
});

describe('keyring', () => {
  it('defaults to an empty object', () => {
    expect(keyStorage.getKeyring()).toEqual({});
  });

  it('round-trips a keyring object', async () => {
    const keyring = { 'bob@example.com': 'BOB-PUBLIC-ARMOR' };
    await keyStorage.saveKeyring(keyring);
    expect(keyStorage.getKeyring()).toEqual(keyring);
  });
});

describe('org override', () => {
  it('defaults to null', () => {
    expect(keyStorage.getOrgOverride()).toBeNull();
  });

  it('round-trips and clears an override', async () => {
    const override = { companyKeyEnabled: true };
    await keyStorage.saveOrgOverride(override);
    expect(keyStorage.getOrgOverride()).toEqual(override);

    await keyStorage.clearOrgOverride();
    expect(keyStorage.getOrgOverride()).toBeNull();
  });
});

describe('sign default preference', () => {
  it('defaults to false', () => {
    expect(keyStorage.getSignDefault()).toBe(false);
  });

  it('round-trips true/false', async () => {
    await keyStorage.saveSignDefault(true);
    expect(keyStorage.getSignDefault()).toBe(true);

    await keyStorage.saveSignDefault(false);
    expect(keyStorage.getSignDefault()).toBe(false);
  });
});

describe('estimateStorageUsage', () => {
  it('grows when keys are stored', async () => {
    const before = keyStorage.estimateStorageUsage();
    await keyStorage.saveKeyPair('PRIVATE-ARMOR', 'PUBLIC-ARMOR', { email: 'alice@example.com' });
    const after = keyStorage.estimateStorageUsage();
    expect(after).toBeGreaterThan(before);
  });
});
