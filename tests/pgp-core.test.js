import { describe, it, expect, beforeAll } from 'vitest';
import {
  generateKeyPair,
  readPublicKey,
  unlockPrivateKey,
  getKeyInfo,
  extractPublicKey,
  hasModernSubkeys,
  addModernSubkeys,
  hasWeakEncryptionKey,
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
import { LEGACY_PUBLIC_KEY, LEGACY_PRIVATE_KEY, LEGACY_PASSPHRASE } from './fixtures/legacy-dsa-elgamal-key.js';

// Flips one base64 character in the armor body to simulate bit-level
// corruption (e.g. a mangled transfer) without touching the header/footer/
// checksum lines, so the armor still parses far enough to reach decryption.
function corruptArmorBody(armored) {
  const lines = armored.split('\n');
  const bodyIndex = lines.findIndex((l, i) => i > 2 && l.length > 20 && !l.startsWith('='));
  if (bodyIndex === -1) throw new Error('corruptArmorBody: no suitable body line found');
  const line = lines[bodyIndex];
  const flippedChar = line[5] === 'A' ? 'B' : 'A';
  lines[bodyIndex] = line.slice(0, 5) + flippedChar + line.slice(6);
  return lines.join('\n');
}

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

describe('extractPublicKey', () => {
  it('extracts a public key matching the original from an armored private key', async () => {
    const extracted = await extractPublicKey(alice.privateKey);
    expect(extracted).toContain('-----BEGIN PGP PUBLIC KEY BLOCK-----');

    const extractedInfo = await getKeyInfo(extracted);
    const originalInfo = await getKeyInfo(alice.publicKey);
    expect(extractedInfo.fingerprint).toBe(originalInfo.fingerprint);
    expect(extractedInfo.isPrivate).toBe(false);
  });
});

describe('generateKeyPair (rsa4096)', () => {
  it('generates a working RSA-4096 key pair', async () => {
    const rsaPair = await generateKeyPair('RSA Legacy Alice', 'rsa-alice@example.com', 'rsa passphrase', 'rsa4096');
    const info = await getKeyInfo(rsaPair.publicKey);
    expect(info.algorithm.toLowerCase()).toContain('rsa');

    const recipientKey = await readPublicKey(rsaPair.publicKey);
    const armored = await encryptMessage('hello rsa', [recipientKey]);
    const unlocked = await unlockPrivateKey(rsaPair.privateKey, 'rsa passphrase');
    const { data } = await decryptMessage(armored, unlocked);
    expect(data).toBe('hello rsa');
  }, 30000);
});

describe('encryptMessage — multiple recipients and text edge cases', () => {
  it('encrypts once to multiple recipients; either can decrypt independently', async () => {
    const aliceKey = await readPublicKey(alice.publicKey);
    const bobKey = await readPublicKey(bob.publicKey);
    const armored = await encryptMessage('shared secret', [aliceKey, bobKey]);

    const unlockedAlice = await unlockPrivateKey(alice.privateKey, 'correct horse battery staple');
    const unlockedBob = await unlockPrivateKey(bob.privateKey, 'hunter2 hunter2 hunter2');

    expect((await decryptMessage(armored, unlockedAlice)).data).toBe('shared secret');
    expect((await decryptMessage(armored, unlockedBob)).data).toBe('shared secret');
  });

  it('round-trips an empty string', async () => {
    const recipientKey = await readPublicKey(alice.publicKey);
    const armored = await encryptMessage('', [recipientKey]);
    const unlockedAlice = await unlockPrivateKey(alice.privateKey, 'correct horse battery staple');
    const { data } = await decryptMessage(armored, unlockedAlice);
    expect(data).toBe('');
  });

  it('round-trips multi-byte unicode text (emoji, CJK)', async () => {
    const recipientKey = await readPublicKey(alice.publicKey);
    const text = 'emoji test 🎉🔒 and CJK text 中文测试';
    const armored = await encryptMessage(text, [recipientKey]);
    const unlockedAlice = await unlockPrivateKey(alice.privateKey, 'correct horse battery staple');
    const { data } = await decryptMessage(armored, unlockedAlice);
    expect(data).toBe(text);
  });
});

describe('tampering and wrong-key failure modes', () => {
  it('decryptMessage throws on a corrupted ciphertext rather than returning garbage', async () => {
    const recipientKey = await readPublicKey(alice.publicKey);
    const armored = await encryptMessage('do not tamper with me', [recipientKey]);
    const corrupted = corruptArmorBody(armored);

    const unlockedAlice = await unlockPrivateKey(alice.privateKey, 'correct horse battery staple');
    await expect(decryptMessage(corrupted, unlockedAlice)).rejects.toThrow();
  });

  it('decryptMessage throws when decrypting with a key that is not a recipient', async () => {
    const recipientKey = await readPublicKey(alice.publicKey);
    const armored = await encryptMessage('for alice only', [recipientKey]);

    const unlockedBob = await unlockPrivateKey(bob.privateKey, 'hunter2 hunter2 hunter2');
    await expect(decryptMessage(armored, unlockedBob)).rejects.toThrow();
  });

  it('decryptAttachment throws on a corrupted ciphertext', async () => {
    const recipientKey = await readPublicKey(alice.publicKey);
    const armored = await encryptAttachment(new Uint8Array([1, 2, 3, 4, 5]), 'data.bin', [recipientKey]);
    const corrupted = corruptArmorBody(armored);

    const unlockedAlice = await unlockPrivateKey(alice.privateKey, 'correct horse battery staple');
    await expect(decryptAttachment(corrupted, unlockedAlice)).rejects.toThrow();
  });

  it('readPublicKey throws on malformed armor instead of silently returning something usable', async () => {
    await expect(readPublicKey('-----BEGIN PGP PUBLIC KEY BLOCK-----\nnot real armor\n-----END PGP PUBLIC KEY BLOCK-----')).rejects.toThrow();
  });

  it('unlockPrivateKey throws on malformed armor', async () => {
    await expect(unlockPrivateKey('not an armored key at all', 'whatever')).rejects.toThrow();
  });
});

describe('legacy DSA + ElGamal key support', () => {
  // LEGACY_PUBLIC_KEY/LEGACY_PRIVATE_KEY are a real GnuPG-generated DSA-1024
  // primary + ElGamal-1024 encryption subkey pair (see tests/fixtures for
  // provenance). This exercises code paths that had zero coverage before:
  // hasWeakEncryptionKey's true branch, hasModernSubkeys' false branch, and
  // addModernSubkeys' full augmentation flow.
  it('getKeyInfo reports the dsa algorithm', async () => {
    const info = await getKeyInfo(LEGACY_PUBLIC_KEY);
    expect(info.algorithm.toLowerCase()).toContain('dsa');
    expect(info.email).toBe('legacy@example.com');
  });

  it('hasWeakEncryptionKey detects the ElGamal encryption subkey as weak', async () => {
    const key = await readPublicKey(LEGACY_PUBLIC_KEY);
    expect(await hasWeakEncryptionKey([key])).toBe(true);
  });

  it('hasModernSubkeys is false before augmentation', async () => {
    expect(await hasModernSubkeys(LEGACY_PUBLIC_KEY)).toBe(false);
  });

  it('unlockPrivateKey works on the legacy key with its passphrase', async () => {
    const unlocked = await unlockPrivateKey(LEGACY_PRIVATE_KEY, LEGACY_PASSPHRASE);
    expect(unlocked.isPrivate()).toBe(true);
  });

  it('encrypts/decrypts a message and an attachment to the legacy ElGamal key', async () => {
    const legacyPublicKey = await readPublicKey(LEGACY_PUBLIC_KEY);
    const unlockedLegacy = await unlockPrivateKey(LEGACY_PRIVATE_KEY, LEGACY_PASSPHRASE);

    const armoredMessage = await encryptMessage('hello legacy key', [legacyPublicKey]);
    expect((await decryptMessage(armoredMessage, unlockedLegacy)).data).toBe('hello legacy key');

    const armoredAttachment = await encryptAttachment(new Uint8Array([9, 8, 7]), 'legacy.bin', [legacyPublicKey]);
    const { data, filename } = await decryptAttachment(armoredAttachment, unlockedLegacy);
    expect(filename).toBe('legacy.bin');
    expect(Array.from(data)).toEqual([9, 8, 7]);
  });

  it('addModernSubkeys augments the key so hasModernSubkeys becomes true, and the new subkey works', async () => {
    const { armoredPrivate, armoredPublic } = await addModernSubkeys(LEGACY_PRIVATE_KEY, LEGACY_PASSPHRASE);

    expect(await hasModernSubkeys(armoredPublic)).toBe(true);

    const augmentedPublicKey = await readPublicKey(armoredPublic);
    const armored = await encryptMessage('hello modern subkey', [augmentedPublicKey]);
    const unlockedAugmented = await unlockPrivateKey(armoredPrivate, LEGACY_PASSPHRASE);
    const { data } = await decryptMessage(armored, unlockedAugmented);
    expect(data).toBe('hello modern subkey');
  }, 15000);
});
