import { describe, it, expect, beforeAll } from 'vitest';
import {
  generateKeyPair,
  readPublicKey,
  unlockPrivateKey,
  getKeyInfo,
  encryptMessage,
  decryptMessage,
  encryptAttachment,
  decryptAttachment,
  stripPgpExtension,
  applyDecryptedExtensionPrefix,
  detectPgpContent,
  base64ToUint8Array,
  uint8ArrayToBase64,
} from '../web/js/pgp/pgp-core.js';

// Real ECC key generation is fast (~tens of ms), so we generate two key pairs
// once and reuse them across tests rather than per-test.
let alice; // { privateKey, publicKey } armored
let bob;

beforeAll(async () => {
  alice = await generateKeyPair('Alice Example', 'alice@example.com', 'correct horse battery staple');
  bob = await generateKeyPair('Bob Example', 'bob@example.com', 'hunter2 hunter2 hunter2');
}, 30000);

describe('generateKeyPair + getKeyInfo', () => {
  it('produces armored keys that getKeyInfo can parse back', async () => {
    const info = await getKeyInfo(alice.publicKey);
    expect(info.email).toBe('alice@example.com');
    expect(info.name).toBe('Alice Example');
    expect(info.isPrivate).toBe(false);
    expect(info.fingerprint).toMatch(/^[0-9A-F]{40}$/);
    expect(info.fingerprintFormatted).toMatch(/^[0-9A-F]{4}( [0-9A-F]{4})+$/);
    expect(info.shortId).toBe(info.fingerprint.slice(-8));
    expect(info.expires).toBeNull();
  });

  it('reports isPrivate: true for an armored private key', async () => {
    const info = await getKeyInfo(alice.privateKey);
    expect(info.isPrivate).toBe(true);
    expect(info.email).toBe('alice@example.com');
  });
});

describe('unlockPrivateKey', () => {
  it('succeeds with the correct passphrase', async () => {
    const unlocked = await unlockPrivateKey(alice.privateKey, 'correct horse battery staple');
    expect(unlocked.isPrivate()).toBe(true);
  });

  it('rejects with the wrong passphrase', async () => {
    await expect(unlockPrivateKey(alice.privateKey, 'wrong passphrase')).rejects.toThrow();
  });
});

describe('encryptMessage + decryptMessage round trip', () => {
  it('preserves plaintext through encrypt/decrypt with no signing', async () => {
    const recipientKey = await readPublicKey(alice.publicKey);
    const armored = await encryptMessage('hello, world', [recipientKey]);
    expect(armored).toContain('-----BEGIN PGP MESSAGE-----');

    const unlockedAlice = await unlockPrivateKey(alice.privateKey, 'correct horse battery staple');
    const { data, signatureResult } = await decryptMessage(armored, unlockedAlice);
    expect(data).toBe('hello, world');
    expect(signatureResult.valid).toBeNull();
  });

  it('reports signatureResult.valid === true when signed and verified against the correct key', async () => {
    const recipientKey = await readPublicKey(bob.publicKey);
    const unlockedAlice = await unlockPrivateKey(alice.privateKey, 'correct horse battery staple');
    const armored = await encryptMessage('signed message', [recipientKey], unlockedAlice);

    const unlockedBob = await unlockPrivateKey(bob.privateKey, 'hunter2 hunter2 hunter2');
    const aliceVerifyKey = await readPublicKey(alice.publicKey);
    const { data, signatureResult } = await decryptMessage(armored, unlockedBob, [aliceVerifyKey]);

    expect(data).toBe('signed message');
    expect(signatureResult.valid).toBe(true);
    expect(signatureResult.signedByKeyId).toBeTruthy();
  });

  it('reports signatureResult.valid === false when verified against the wrong key', async () => {
    const recipientKey = await readPublicKey(bob.publicKey);
    const unlockedAlice = await unlockPrivateKey(alice.privateKey, 'correct horse battery staple');
    const armored = await encryptMessage('signed message', [recipientKey], unlockedAlice);

    const unlockedBob = await unlockPrivateKey(bob.privateKey, 'hunter2 hunter2 hunter2');
    // Verify against Bob's own key instead of Alice's — the actual signer.
    const wrongVerifyKey = await readPublicKey(bob.publicKey);
    const { signatureResult } = await decryptMessage(armored, unlockedBob, [wrongVerifyKey]);

    expect(signatureResult.valid).toBe(false);
  });
});

describe('encryptAttachment + decryptAttachment round trip', () => {
  it('preserves binary data and recovers the filename', async () => {
    const recipientKey = await readPublicKey(alice.publicKey);
    const originalBytes = new Uint8Array([0, 1, 2, 253, 254, 255, 42]);

    const armored = await encryptAttachment(originalBytes, 'report.pdf', [recipientKey]);
    expect(armored).toContain('-----BEGIN PGP MESSAGE-----');

    const unlockedAlice = await unlockPrivateKey(alice.privateKey, 'correct horse battery staple');
    const { data, filename } = await decryptAttachment(armored, unlockedAlice);

    expect(filename).toBe('report.pdf');
    expect(Array.from(data)).toEqual(Array.from(originalBytes));
  });

  it('treats the "_console" and "email" filename markers as empty', async () => {
    const recipientKey = await readPublicKey(alice.publicKey);
    const unlockedAlice = await unlockPrivateKey(alice.privateKey, 'correct horse battery staple');

    const armoredConsole = await encryptAttachment(new Uint8Array([1]), '_console', [recipientKey]);
    const consoleResult = await decryptAttachment(armoredConsole, unlockedAlice);
    expect(consoleResult.filename).toBe('');

    const armoredEmail = await encryptAttachment(new Uint8Array([1]), 'email', [recipientKey]);
    const emailResult = await decryptAttachment(armoredEmail, unlockedAlice);
    expect(emailResult.filename).toBe('');
  });
});

describe('stripPgpExtension', () => {
  it('strips known PGP extensions case-insensitively', () => {
    expect(stripPgpExtension('report.pdf.pgp')).toBe('report.pdf');
    expect(stripPgpExtension('report.pdf.PGP')).toBe('report.pdf');
    expect(stripPgpExtension('archive.tar.gz.gpg')).toBe('archive.tar.gz');
    expect(stripPgpExtension('signed.txt.asc')).toBe('signed.txt');
  });

  it('discards a leading directory/drive path', () => {
    expect(stripPgpExtension('C:\\Users\\me\\Downloads\\report.pdf.pgp')).toBe('report.pdf');
    expect(stripPgpExtension('/home/me/report.pdf.pgp')).toBe('report.pdf');
  });

  it('no-ops on a name with no recognized PGP extension', () => {
    expect(stripPgpExtension('report.pdf')).toBe('report.pdf');
  });
});

describe('applyDecryptedExtensionPrefix', () => {
  it('inserts the prefix before the extension', () => {
    expect(applyDecryptedExtensionPrefix('report.xlsx', 'pgpDecrypted')).toBe('report.pgpDecrypted.xlsx');
  });

  it('appends the prefix as a trailing segment when there is no extension', () => {
    expect(applyDecryptedExtensionPrefix('README', 'pgpDecrypted')).toBe('README.pgpDecrypted');
  });

  it('treats a dotfile (leading dot only) as having no extension', () => {
    expect(applyDecryptedExtensionPrefix('.gitignore', 'pgpDecrypted')).toBe('.gitignore.pgpDecrypted');
  });

  it('is a no-op for a falsy prefix', () => {
    expect(applyDecryptedExtensionPrefix('report.xlsx', '')).toBe('report.xlsx');
    expect(applyDecryptedExtensionPrefix('report.xlsx', null)).toBe('report.xlsx');
    expect(applyDecryptedExtensionPrefix('report.xlsx', undefined)).toBe('report.xlsx');
  });
});

describe('detectPgpContent', () => {
  it('detects each armor type', () => {
    expect(detectPgpContent('-----BEGIN PGP MESSAGE-----\n...')).toBe('encrypted');
    expect(detectPgpContent('-----BEGIN PGP SIGNED MESSAGE-----\n...')).toBe('signed');
    expect(detectPgpContent('-----BEGIN PGP PUBLIC KEY BLOCK-----\n...')).toBe('public-key');
    expect(detectPgpContent('-----BEGIN PGP PRIVATE KEY BLOCK-----\n...')).toBe('private-key');
  });

  it('returns null for plain text or empty input', () => {
    expect(detectPgpContent('just some regular text')).toBeNull();
    expect(detectPgpContent('')).toBeNull();
    expect(detectPgpContent(null)).toBeNull();
  });
});

describe('base64ToUint8Array / uint8ArrayToBase64', () => {
  it('round-trips arbitrary bytes', () => {
    const original = new Uint8Array([0, 1, 2, 127, 128, 200, 255]);
    const base64 = uint8ArrayToBase64(original);
    const roundTripped = base64ToUint8Array(base64);
    expect(Array.from(roundTripped)).toEqual(Array.from(original));
  });
});
