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

  const elements = {
    'btn-encrypt': encryptBtn,
    'btn-decrypt': decryptBtn,
    'status-bar': statusEl,
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
