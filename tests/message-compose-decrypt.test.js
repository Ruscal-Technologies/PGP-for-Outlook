import { describe, it, expect, beforeEach, vi } from 'vitest';
import { cacheSessionKey, clearSessionKey } from '../web/js/pgp/session-cache.js';

// key-storage.js and session-cache.js talk to Office.context.roamingSettings /
// are otherwise irrelevant to refreshComposeButtons()/handleDecrypt()'s own
// logic — mock key-storage so tests control the private/public key values
// directly without stubbing roaming settings.
vi.mock('../web/js/pgp/key-storage.js', () => ({
  getPrivateKey: vi.fn(() => 'armored-priv-key'),
  getPublicKey: vi.fn(() => 'armored-pub-key'),
  hasKeyPair: vi.fn(() => true),
  getSignDefault: vi.fn(() => false),
}));

// Keep the real detectPgpContent/stripPgpExtension/uint8ArrayToBase64/
// base64ToUint8Array (pure, no Office dependency) but replace the two
// functions that would otherwise require real OpenPGP key material.
vi.mock('../web/js/pgp/pgp-core.js', async (importOriginal) => {
  const actual = await importOriginal();
  return {
    ...actual,
    decryptMessage: vi.fn(),
    decryptAttachment: vi.fn(),
    unlockPrivateKey: vi.fn(),
    // getKeyInfo normally parses real OpenPGP armor via openpgp.readKey();
    // getPublicKey() is mocked above to a fake string, so this must be
    // mocked too or the passphrase-prompt path (Task 4's third test) would
    // throw trying to parse it.
    getKeyInfo: vi.fn(async () => ({ shortId: 'ABCD1234' })),
  };
});

function installStubs({ bodyText = '', attachments = [] } = {}) {
  const encryptBtn = { classList: { add: vi.fn(), remove: vi.fn(), contains: vi.fn(() => false) }, disabled: false, focus: vi.fn() };
  const decryptBtn = { classList: { add: vi.fn(), remove: vi.fn(), contains: vi.fn(() => false) }, disabled: false, focus: vi.fn() };
  const statusEl = { className: '', textContent: '', classList: { add: vi.fn(), remove: vi.fn(), contains: vi.fn(() => false) } };
  const spinnerEls = {
    'encrypt-spinner': { classList: { add: vi.fn(), remove: vi.fn() } },
    'decrypt-spinner': { classList: { add: vi.fn(), remove: vi.fn() } },
  };
  // loadAttachments() (already defined in MessageCompose.js, called by
  // handleDecrypt() from Task 5 onward) reads/writes these four elements
  // every time it runs — stub them even in tests that don't care about
  // attachments, since handleDecrypt() always calls loadAttachments().
  const attachmentListEl = { children: [], appendChild: vi.fn() };
  const attachmentsEmptyEl = { classList: { add: vi.fn(), remove: vi.fn() } };
  const attachmentsLoadingEl = { classList: { add: vi.fn(), remove: vi.fn() } };

  // Passphrase modal elements, needed whenever a test exercises the
  // no-cached-session-key path (getSessionKey() returns null, so
  // promptPassphrase() is invoked). okBtn/cancelBtn capture their click
  // callback so a test can simulate the user clicking OK/Cancel — same
  // pattern as Task 3's dedicated promptPassphrase test.
  const passphraseInput = { value: '', focus: vi.fn(), addEventListener: vi.fn(), removeEventListener: vi.fn() };
  const passphraseModal = { style: {}, classList: { add: vi.fn(), remove: vi.fn() } };
  const passphraseError = { classList: { add: vi.fn(), remove: vi.fn() } };
  const passphraseMsg = { textContent: '' };
  const okBtn = { addEventListener: (_e, cb) => { okBtn._cb = cb; }, removeEventListener: vi.fn() };
  const cancelBtn = { addEventListener: (_e, cb) => { cancelBtn._cb = cb; }, removeEventListener: vi.fn() };

  // updateSessionStatus() (invoked by handleDecrypt() after caching a newly
  // unlocked key, same as handleEncrypt()) reads/writes these two elements.
  const sessionStatusBar = { classList: { add: vi.fn(), remove: vi.fn(), contains: vi.fn(() => false) } };
  const sessionStatusText = { textContent: '' };

  const elements = {
    'btn-encrypt': encryptBtn,
    'btn-decrypt': decryptBtn,
    'status-bar': statusEl,
    'session-status': sessionStatusBar,
    'session-status-text': sessionStatusText,
    'attachment-list': attachmentListEl,
    'attachments-empty': attachmentsEmptyEl,
    'attachments-loading': attachmentsLoadingEl,
    'passphrase-input': passphraseInput,
    'passphrase-modal': passphraseModal,
    'passphrase-error': passphraseError,
    'passphrase-modal-msg': passphraseMsg,
    'btn-passphrase-ok': okBtn,
    'btn-passphrase-cancel': cancelBtn,
    ...spinnerEls,
  };
  global.document = {
    getElementById: (id) => elements[id] || null,
    // Only used by loadAttachments() to render <li> rows for non-empty
    // attachment lists — a bare object is enough since nothing reads it back.
    createElement: () => ({ className: '', innerHTML: '' }),
  };

  const getAsync = vi.fn((_coercionType, cb) => cb({ status: 'succeeded', value: bodyText }));
  const setAsync = vi.fn((_html, _opts, cb) => cb({ status: 'succeeded' }));
  const getAttachmentsAsync = vi.fn((_opts, cb) => cb({ status: 'succeeded', value: attachments }));

  global.Office = {
    onReady: () => {},
    CoercionType: { Text: 'text', Html: 'html' },
    AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
    context: {
      mailbox: {
        userProfile: { emailAddress: 'me@example.com' },
        item: { body: { getAsync, setAsync }, getAttachmentsAsync },
      },
      requirements: { isSetSupported: () => true },
    },
  };

  return {
    encryptBtn, decryptBtn, statusEl, getAsync, setAsync, getAttachmentsAsync,
    passphraseInput, passphraseMsg, okBtn, cancelBtn,
  };
}

let refreshComposeButtons;

beforeEach(async () => {
  vi.clearAllMocks();
});

describe('refreshComposeButtons', () => {
  it('shows Decrypt and hides Encrypt when the body is PGP-encrypted', async () => {
    const { encryptBtn, decryptBtn } = installStubs({
      bodyText: '-----BEGIN PGP MESSAGE-----\nabc\n-----END PGP MESSAGE-----',
    });
    ({ refreshComposeButtons } = await import('../web/MessageCompose.js'));

    await refreshComposeButtons();

    expect(decryptBtn.classList.remove).toHaveBeenCalledWith('pgp-hidden');
    expect(encryptBtn.classList.add).toHaveBeenCalledWith('pgp-hidden');
  });

  it('shows Encrypt and hides Decrypt when the body is not encrypted', async () => {
    const { encryptBtn, decryptBtn } = installStubs({ bodyText: 'just a normal draft' });
    ({ refreshComposeButtons } = await import('../web/MessageCompose.js'));

    await refreshComposeButtons();

    expect(encryptBtn.classList.remove).toHaveBeenCalledWith('pgp-hidden');
    expect(decryptBtn.classList.add).toHaveBeenCalledWith('pgp-hidden');
  });
});

describe('promptPassphrase', () => {
  it('sets the modal message text from its argument', async () => {
    const msgEl = { textContent: '' };
    const input = { value: '', focus: vi.fn(), addEventListener: vi.fn(), removeEventListener: vi.fn() };
    const errEl = { classList: { add: vi.fn(), remove: vi.fn() } };
    const modal = { style: {}, classList: { add: vi.fn(), remove: vi.fn() } };
    const okBtn = { addEventListener: (_e, cb) => { okBtn._cb = cb; }, removeEventListener: vi.fn() };
    const cancelBtn = { addEventListener: vi.fn(), removeEventListener: vi.fn() };
    const elements = {
      'passphrase-modal': modal,
      'passphrase-input': input,
      'passphrase-error': errEl,
      'passphrase-modal-msg': msgEl,
      'btn-passphrase-ok': okBtn,
      'btn-passphrase-cancel': cancelBtn,
    };
    global.document = { getElementById: (id) => elements[id] || null };
    global.Office = { onReady: () => {} };

    const { promptPassphraseForTest } = await import('../web/MessageCompose.js');
    const resultPromise = promptPassphraseForTest('Enter your passphrase to decrypt this message.');
    input.value = 'hunter2';
    okBtn._cb();
    await resultPromise;

    expect(msgEl.textContent).toBe('Enter your passphrase to decrypt this message.');
  });
});

describe('handleDecrypt — body restore', () => {
  beforeEach(() => {
    clearSessionKey();
  });

  it('restores the original HTML body and switches buttons back to Encrypt', async () => {
    const { decryptBtn, setAsync } = installStubs({
      bodyText: '-----BEGIN PGP MESSAGE-----\narmored\n-----END PGP MESSAGE-----',
    });
    const pgpCore = await import('../web/js/pgp/pgp-core.js');
    pgpCore.unlockPrivateKey.mockResolvedValue({ id: 'unlocked-key' });
    pgpCore.decryptMessage.mockResolvedValue({ data: '<p>original body</p>', signatureResult: { valid: null } });

    // Simulate an already-cached session key so no passphrase prompt is needed.
    cacheSessionKey({ id: 'unlocked-key' }, 'me@example.com', 'ABCD1234');

    const { handleDecrypt } = await import('../web/MessageCompose.js');
    await handleDecrypt();

    expect(pgpCore.decryptMessage).toHaveBeenCalledWith(
      expect.stringContaining('-----BEGIN PGP MESSAGE-----'),
      { id: 'unlocked-key' },
    );
    expect(setAsync).toHaveBeenCalledWith(
      '<p>original body</p>',
      { coercionType: 'html' },
      expect.any(Function),
    );
  });

  it('shows an error status and does not touch the body when the body is not encrypted', async () => {
    const { statusEl, setAsync } = installStubs({ bodyText: 'plain draft, nothing encrypted' });

    const { handleDecrypt } = await import('../web/MessageCompose.js');
    await handleDecrypt();

    expect(setAsync).not.toHaveBeenCalled();
    expect(statusEl.textContent).toMatch(/not.*encrypted|no.*encrypted/i);
  });

  it('prompts for the passphrase, unlocks, and caches the key when no session key is cached', async () => {
    const { setAsync, passphraseInput, passphraseMsg, okBtn } = installStubs({
      bodyText: '-----BEGIN PGP MESSAGE-----\narmored\n-----END PGP MESSAGE-----',
    });
    const pgpCore = await import('../web/js/pgp/pgp-core.js');
    pgpCore.unlockPrivateKey.mockResolvedValue({ id: 'unlocked-key' });
    pgpCore.decryptMessage.mockResolvedValue({ data: '<p>original body</p>', signatureResult: { valid: null } });

    // No cacheSessionKey() call this time — getSessionKey() must return null,
    // forcing handleDecrypt() through the promptPassphrase() branch.
    const { handleDecrypt } = await import('../web/MessageCompose.js');
    const decryptPromise = handleDecrypt();

    // Let the microtask queue advance until promptPassphrase() has populated
    // the modal — the exact number of intervening awaits inside
    // handleDecrypt() before that point is an implementation detail, so poll
    // rather than hardcoding a tick count.
    for (let i = 0; i < 20 && !passphraseMsg.textContent; i++) {
      await Promise.resolve();
    }
    expect(passphraseMsg.textContent).toBe('Enter your passphrase to decrypt this message.');
    passphraseInput.value = 'hunter2';
    okBtn._cb();

    await decryptPromise;

    expect(pgpCore.unlockPrivateKey).toHaveBeenCalledWith('armored-priv-key', 'hunter2');
    expect(setAsync).toHaveBeenCalledWith(
      '<p>original body</p>',
      { coercionType: 'html' },
      expect.any(Function),
    );
  });
});

describe('handleDecrypt — attachment reversal', () => {
  beforeEach(() => {
    clearSessionKey();
  });

  it('reverts every .pgp attachment, leaves non-.pgp attachments alone, and reports success', async () => {
    const attachments = [
      { id: 'a1', name: 'report.pdf.pgp', isInline: false },
      { id: 'a2', name: 'notes.txt', isInline: false },
    ];
    const { statusEl } = installStubs({
      bodyText: '-----BEGIN PGP MESSAGE-----\narmored\n-----END PGP MESSAGE-----',
      attachments,
    });
    const pgpCore = await import('../web/js/pgp/pgp-core.js');
    pgpCore.decryptMessage.mockResolvedValue({ data: '<p>body</p>', signatureResult: { valid: null } });
    pgpCore.decryptAttachment.mockResolvedValue({ data: new Uint8Array([1, 2, 3]), filename: 'report.pdf' });
    cacheSessionKey({ id: 'k' }, 'me@example.com', 'ABCD1234');

    const item = global.Office.context.mailbox.item;
    item.getAttachmentContentAsync = vi.fn((id, cb) => cb({ status: 'succeeded', value: { format: 'base64', content: btoa('armored-attachment-text') } }));
    item.removeAttachmentAsync = vi.fn((id, cb) => cb({ status: 'succeeded' }));
    item.addFileAttachmentFromBase64Async = vi.fn((base64, name, opts, cb) => cb({ status: 'succeeded', value: name }));

    const { handleDecrypt } = await import('../web/MessageCompose.js');
    await handleDecrypt();

    expect(pgpCore.decryptAttachment).toHaveBeenCalledTimes(1);
    expect(item.removeAttachmentAsync).toHaveBeenCalledWith('a1', expect.any(Function));
    expect(item.addFileAttachmentFromBase64Async).toHaveBeenCalledWith(
      expect.any(String), 'report.pdf', { asyncContext: null }, expect.any(Function),
    );
    expect(statusEl.textContent).toContain('✓ Message decrypted.');
  });

  it('falls back to stripPgpExtension(name) when decryptAttachment returns no filename', async () => {
    const attachments = [{ id: 'a1', name: 'archive.zip.pgp', isInline: false }];
    installStubs({
      bodyText: '-----BEGIN PGP MESSAGE-----\narmored\n-----END PGP MESSAGE-----',
      attachments,
    });
    const pgpCore = await import('../web/js/pgp/pgp-core.js');
    pgpCore.decryptMessage.mockResolvedValue({ data: '<p>body</p>', signatureResult: { valid: null } });
    pgpCore.decryptAttachment.mockResolvedValue({ data: new Uint8Array([1]), filename: '' });
    cacheSessionKey({ id: 'k' }, 'me@example.com', 'ABCD1234');

    const item = global.Office.context.mailbox.item;
    item.getAttachmentContentAsync = vi.fn((id, cb) => cb({ status: 'succeeded', value: { format: 'base64', content: btoa('armored-attachment-text') } }));
    item.removeAttachmentAsync = vi.fn((id, cb) => cb({ status: 'succeeded' }));
    item.addFileAttachmentFromBase64Async = vi.fn((base64, name, opts, cb) => cb({ status: 'succeeded', value: name }));

    const { handleDecrypt } = await import('../web/MessageCompose.js');
    await handleDecrypt();

    expect(item.addFileAttachmentFromBase64Async).toHaveBeenCalledWith(
      expect.any(String), 'archive.zip', { asyncContext: null }, expect.any(Function),
    );
  });

  it('leaves a failed attachment untouched and reports a warning naming it, without blocking the others', async () => {
    const attachments = [
      { id: 'a1', name: 'good.txt.pgp', isInline: false },
      { id: 'a2', name: 'bad.txt.pgp', isInline: false },
    ];
    const { statusEl } = installStubs({
      bodyText: '-----BEGIN PGP MESSAGE-----\narmored\n-----END PGP MESSAGE-----',
      attachments,
    });
    const pgpCore = await import('../web/js/pgp/pgp-core.js');
    pgpCore.decryptMessage.mockResolvedValue({ data: '<p>body</p>', signatureResult: { valid: null } });
    pgpCore.decryptAttachment.mockImplementation(async (armored) => {
      if (armored.includes('bad')) throw new Error('corrupted armor');
      return { data: new Uint8Array([9]), filename: 'good.txt' };
    });
    cacheSessionKey({ id: 'k' }, 'me@example.com', 'ABCD1234');

    const item = global.Office.context.mailbox.item;
    item.getAttachmentContentAsync = vi.fn((id, cb) => {
      const text = id === 'a1' ? 'good-armored' : 'bad-armored';
      cb({ status: 'succeeded', value: { format: 'base64', content: btoa(text) } });
    });
    item.removeAttachmentAsync = vi.fn((id, cb) => cb({ status: 'succeeded' }));
    item.addFileAttachmentFromBase64Async = vi.fn((base64, name, opts, cb) => cb({ status: 'succeeded', value: name }));

    const { handleDecrypt } = await import('../web/MessageCompose.js');
    await handleDecrypt();

    // Only the good attachment was removed/re-added.
    expect(item.removeAttachmentAsync).toHaveBeenCalledTimes(1);
    expect(item.removeAttachmentAsync).toHaveBeenCalledWith('a1', expect.any(Function));
    expect(statusEl.textContent).toContain('bad.txt.pgp');
    expect(statusEl.textContent).toMatch(/could not/i);
  });
});
