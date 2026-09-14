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
  getAutoEncryptDefault: vi.fn(() => false),
  getAutoSendDefault: vi.fn(() => false),
  hasAcknowledgedWarning: vi.fn(() => false),
  saveAcknowledgedWarning: vi.fn(async () => {}),
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
    // readPublicKey/encryptMessage normally parse/produce real OpenPGP
    // material — mocked so the handleEncrypt() button-visibility test below
    // doesn't need real key objects.
    readPublicKey: vi.fn(async () => ({ fake: 'own-public-key' })),
    encryptMessage: vi.fn(),
  };
});

// handleEncrypt() calls resolveRecipients() (key-discovery.js) to resolve
// To/Cc emails to keys — mocked so the button-visibility test can supply a
// recipient with an already-resolved key without hitting WKD/VKS.
vi.mock('../web/js/pgp/key-discovery.js', () => ({
  resolveRecipients: vi.fn(async (emails) => emails.map((email) => (
    { email, key: { fake: 'recipient-key' }, status: 'found', source: 'keyring', armoredKey: null }
  ))),
  KeyStatus: { FOUND: 'found', NOT_FOUND: 'not-found', ERROR: 'error' },
}));

// org-config.js is only consulted for the company-key panel — mocked to the
// "disabled" state so handleEncrypt()'s company-key branch is a no-op.
vi.mock('../web/js/pgp/org-config.js', () => ({
  loadOrgConfig: vi.fn(async () => ({})),
  isCompanyKeyEnabled: vi.fn(() => false),
  isCompanyKeyRequired: vi.fn(() => false),
  getCompanyKeyEmails: vi.fn(() => []),
  fetchCompanyKeys: vi.fn(async () => []),
}));

function installStubs({ bodyText = '', attachments = [], recipients = [] } = {}) {
  const encryptBtn = { classList: { add: vi.fn(), remove: vi.fn(), contains: vi.fn(() => false) }, disabled: false, focus: vi.fn() };
  const decryptBtn = { classList: { add: vi.fn(), remove: vi.fn(), contains: vi.fn(() => false) }, disabled: false, focus: vi.fn() };
  // classList.contains reflects the real className string (rather than a
  // hardcoded false) since handleAutoEncrypt() checks it after handleEncrypt()
  // calls showStatus(), which sets className directly, not via classList.add.
  const statusEl = { className: '', textContent: '', classList: { add: vi.fn(), remove: vi.fn(), contains: vi.fn(function (cls) { return statusEl.className.split(/\s+/).includes(cls); }) } };
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

  // Elements handleEncrypt()'s recipient-loading/company-key/sign-toggle path
  // touches — stubbed even in decrypt-only tests since installStubs() is
  // shared, following the same "stub everything handleX always calls"
  // pattern already used above for attachments.
  const recipientListEl = { innerHTML: '', classList: { add: vi.fn(), remove: vi.fn() }, appendChild: vi.fn() };
  const recipientsLoadingEl = { classList: { add: vi.fn(), remove: vi.fn() } };
  const recipientsEmptyEl = { classList: { add: vi.fn(), remove: vi.fn() } };
  const signToggle = { checked: false };
  const companyKeyToggle = { checked: false };
  const companyKeyDisabledEl = { classList: { add: vi.fn(), remove: vi.fn() } };
  const companyKeyPanelEl = { classList: { add: vi.fn(), remove: vi.fn() } };
  const encryptSendConfirmPanel = { classList: { add: vi.fn(), remove: vi.fn() } };
  const encryptSendConfirmBtn = { addEventListener: (_e, cb) => { encryptSendConfirmBtn._cb = cb; }, removeEventListener: vi.fn() };
  const encryptSendCancelBtn = { addEventListener: (_e, cb) => { encryptSendCancelBtn._cb = cb; }, removeEventListener: vi.fn() };

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
    'recipient-list': recipientListEl,
    'recipients-loading': recipientsLoadingEl,
    'recipients-empty': recipientsEmptyEl,
    'sign-toggle': signToggle,
    'company-key-toggle': companyKeyToggle,
    'company-key-disabled': companyKeyDisabledEl,
    'company-key-panel': companyKeyPanelEl,
    'panel-encrypt-send-confirm': encryptSendConfirmPanel,
    'btn-encrypt-send-confirm': encryptSendConfirmBtn,
    'btn-encrypt-send-cancel': encryptSendCancelBtn,
    ...spinnerEls,
  };
  global.document = {
    getElementById: (id) => elements[id] || null,
    // Only used by loadAttachments() to render <li> rows for non-empty
    // attachment lists — a bare object is enough since nothing reads it back.
    createElement: () => ({ className: '', innerHTML: '', dataset: {} }),
  };

  // Stateful body: setAsync updates what a later getAsync returns, so
  // refreshComposeButtons()'s post-action re-read reflects an Encrypt/Decrypt
  // that just ran, the same as it would against a real Outlook body.
  let currentBody = bodyText;
  const getAsync = vi.fn((_coercionType, cb) => cb({ status: 'succeeded', value: currentBody }));
  const setAsync = vi.fn((html, _opts, cb) => { currentBody = html; cb({ status: 'succeeded' }); });
  const getAttachmentsAsync = vi.fn((_opts, cb) => cb({ status: 'succeeded', value: attachments }));
  // loadRecipients() polls item.to/item.cc via getRecipientsAsync() until two
  // consecutive reads agree on the count — returning the same list every call
  // satisfies that on the first poll.
  const to = { getAsync: vi.fn((cb) => cb({ status: 'succeeded', value: recipients })) };
  const cc = { getAsync: vi.fn((cb) => cb({ status: 'succeeded', value: [] })) };

  global.Office = {
    onReady: () => {},
    CoercionType: { Text: 'text', Html: 'html' },
    AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
    context: {
      mailbox: {
        userProfile: { emailAddress: 'me@example.com' },
        item: { body: { getAsync, setAsync }, getAttachmentsAsync, to, cc },
      },
      requirements: { isSetSupported: () => true },
    },
  };

  return {
    encryptBtn, decryptBtn, statusEl, getAsync, setAsync, getAttachmentsAsync,
    passphraseInput, passphraseMsg, okBtn, cancelBtn, signToggle,
    encryptSendConfirmPanel, encryptSendConfirmBtn, encryptSendCancelBtn,
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

  it('never removes the .pgp original when re-adding the decrypted file fails, so the attachment is not lost', async () => {
    const attachments = [{ id: 'a1', name: 'report.pdf.pgp', isInline: false }];
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
    // Simulate the re-add itself failing (quota, size limit, transient Office error).
    item.addFileAttachmentFromBase64Async = vi.fn((base64, name, opts, cb) => cb({ status: 'failed', error: { message: 'attachment too large' } }));

    const { handleDecrypt } = await import('../web/MessageCompose.js');
    await handleDecrypt();

    // The decrypted file is added BEFORE the .pgp original is removed, so a
    // failed add must leave the original attachment in place -- never
    // deleted with nothing to replace it.
    expect(item.addFileAttachmentFromBase64Async).toHaveBeenCalledTimes(1);
    expect(item.removeAttachmentAsync).not.toHaveBeenCalled();
    expect(statusEl.textContent).toContain('report.pdf.pgp');
    expect(statusEl.textContent).toMatch(/could not/i);
  });
});

describe('handleEncrypt — button visibility', () => {
  beforeEach(() => {
    clearSessionKey();
  });

  it('hides Encrypt and shows Decrypt after a successful encrypt', async () => {
    const { encryptBtn, decryptBtn } = installStubs({
      bodyText: '<p>hello</p>',
      recipients: [{ emailAddress: 'friend@example.com' }],
    });
    const pgpCore = await import('../web/js/pgp/pgp-core.js');
    pgpCore.encryptMessage.mockResolvedValue(
      '-----BEGIN PGP MESSAGE-----\nencrypted\n-----END PGP MESSAGE-----',
    );

    const { handleEncrypt } = await import('../web/MessageCompose.js');
    await handleEncrypt();

    expect(decryptBtn.classList.remove).toHaveBeenCalledWith('pgp-hidden');
    expect(encryptBtn.classList.add).toHaveBeenCalledWith('pgp-hidden');
  });

  it('returns true only for the run that actually encrypted, even when a concurrent run leaves a stale success status visible', async () => {
    const { statusEl } = installStubs({
      bodyText: '<p>hello</p>',
      recipients: [{ emailAddress: 'friend@example.com' }],
    });
    const pgpCore = await import('../web/js/pgp/pgp-core.js');
    pgpCore.encryptMessage.mockResolvedValue(
      '-----BEGIN PGP MESSAGE-----\nencrypted\n-----END PGP MESSAGE-----',
    );

    const { handleEncrypt } = await import('../web/MessageCompose.js');

    // handleEncrypt() checks its in-flight guard synchronously, before its
    // own first await -- so calling it a second time in the same tick,
    // before the first call has had a chance to progress, deterministically
    // finds the guard already held (no fake timers or manual deferred
    // promises needed for this ordering).
    const firstRun = handleEncrypt();
    const secondRun = handleEncrypt();
    const [firstResult, secondResult] = await Promise.all([firstRun, secondRun]);

    // The first (real) run succeeded and left the status bar showing
    // success -- exactly the ambient state a caller must NOT infer its own
    // success from.
    expect(firstResult).toBe(true);
    expect(statusEl.className).toContain('pgp-alert--success');
    // The second run did nothing (skipped by the guard) and must report
    // that honestly, regardless of what the status bar -- written by the
    // OTHER run -- currently shows.
    expect(secondResult).toBe(false);
  });

  it('refuses to encrypt when there are no recipients at all, even though an empty array vacuously passes .every()', async () => {
    const { statusEl } = installStubs({
      bodyText: '<p>hello</p>',
      recipients: [], // no To/Cc recipients whatsoever
    });
    const pgpCore = await import('../web/js/pgp/pgp-core.js');

    const { handleEncrypt } = await import('../web/MessageCompose.js');
    const result = await handleEncrypt();

    expect(result).toBe(false);
    expect(pgpCore.encryptMessage).not.toHaveBeenCalled();
    expect(statusEl.textContent).toContain('Not all recipients have a resolved key yet');
  });
});

describe('mergePreservingManuallyResolvedKeys', () => {
  it('preserves a prior key when the fresh pass has none for the same email', async () => {
    const { mergePreservingManuallyResolvedKeys } = await import('../web/MessageCompose.js');
    const priorResults = [
      { email: 'a@example.com', key: { fake: 'pasted-key' }, status: 'found_local', source: 'Pasted', armoredKey: 'ARMOR' },
    ];
    const freshResults = [
      { email: 'a@example.com', key: null, status: 'not_found', source: null, armoredKey: null },
    ];

    expect(mergePreservingManuallyResolvedKeys(freshResults, priorResults)).toEqual(priorResults);
  });

  it('uses the fresh result when it already has its own key', async () => {
    const { mergePreservingManuallyResolvedKeys } = await import('../web/MessageCompose.js');
    const priorResults = [{ email: 'a@example.com', key: { fake: 'old-key' }, status: 'found_local', source: 'Pasted', armoredKey: 'OLD' }];
    const freshResults = [{ email: 'a@example.com', key: { fake: 'new-key' }, status: 'found', source: 'WKD', armoredKey: 'NEW' }];

    expect(mergePreservingManuallyResolvedKeys(freshResults, priorResults)).toEqual(freshResults);
  });

  it('uses the fresh result (still keyless) when there is no matching prior entry to preserve', async () => {
    const { mergePreservingManuallyResolvedKeys } = await import('../web/MessageCompose.js');
    const priorResults = [];
    const freshResults = [{ email: 'new-recipient@example.com', key: null, status: 'not_found', source: null, armoredKey: null }];

    expect(mergePreservingManuallyResolvedKeys(freshResults, priorResults)).toEqual(freshResults);
  });
});

describe('auto-encrypt on pane load', () => {
  beforeEach(async () => {
    // MessageCompose.js tracks _autoEncryptFired as module-level state so it
    // can only fire once per real pane session — vi.resetModules() gives
    // each test here its own fresh module instance so that guard doesn't
    // leak across tests, matching the same pattern message-compose.test.js
    // already uses for this file's other module-level state.
    vi.resetModules();
    clearSessionKey();

    // vi.resetModules() only reloads real (non-mocked) modules — the
    // key-discovery.js mock instance registered by vi.mock() at the top of
    // this file is a singleton that survives across it, so a test that
    // overrides resolveRecipients with a persistent .mockResolvedValue(...)
    // (rather than a self-consuming .mockResolvedValueOnce(...)) would
    // otherwise leak that override into the next test. Restore the default
    // "every recipient resolves immediately" behavior here so each test
    // starts from the same baseline.
    const keyDiscovery = await import('../web/js/pgp/key-discovery.js');
    keyDiscovery.resolveRecipients.mockReset();
    keyDiscovery.resolveRecipients.mockImplementation(async (emails) => emails.map((email) => (
      { email, key: { fake: 'recipient-key' }, status: 'found', source: 'keyring', armoredKey: null }
    )));
  });

  it('fires handleEncrypt automatically when auto-encrypt is on and all recipients already have keys at pane load', async () => {
    const keyStorage = await import('../web/js/pgp/key-storage.js');
    keyStorage.getAutoEncryptDefault.mockReturnValue(true);
    keyStorage.getAutoSendDefault.mockReturnValue(false);

    const { encryptBtn, decryptBtn } = installStubs({
      bodyText: '<p>hello</p>',
      recipients: [{ emailAddress: 'friend@example.com' }],
    });
    const pgpCore = await import('../web/js/pgp/pgp-core.js');
    pgpCore.encryptMessage.mockResolvedValue(
      '-----BEGIN PGP MESSAGE-----\nencrypted\n-----END PGP MESSAGE-----',
    );

    const { maybeAutoEncryptForTest } = await import('../web/MessageCompose.js');
    await maybeAutoEncryptForTest();

    expect(decryptBtn.classList.remove).toHaveBeenCalledWith('pgp-hidden');
    expect(encryptBtn.classList.add).toHaveBeenCalledWith('pgp-hidden');
  });

  it('does not fire when auto-encrypt is off', async () => {
    const keyStorage = await import('../web/js/pgp/key-storage.js');
    keyStorage.getAutoEncryptDefault.mockReturnValue(false);

    const { encryptBtn } = installStubs({
      bodyText: '<p>hello</p>',
      recipients: [{ emailAddress: 'friend@example.com' }],
    });
    const pgpCore = await import('../web/js/pgp/pgp-core.js');

    const { maybeAutoEncryptForTest } = await import('../web/MessageCompose.js');
    await maybeAutoEncryptForTest();

    expect(pgpCore.encryptMessage).not.toHaveBeenCalled();
    expect(encryptBtn.classList.add).not.toHaveBeenCalledWith('pgp-hidden');
  });

  it('does not fire a second time on the same pane session even if called again', async () => {
    const keyStorage = await import('../web/js/pgp/key-storage.js');
    keyStorage.getAutoEncryptDefault.mockReturnValue(true);
    keyStorage.getAutoSendDefault.mockReturnValue(false);

    installStubs({
      bodyText: '<p>hello</p>',
      recipients: [{ emailAddress: 'friend@example.com' }],
    });
    const pgpCore = await import('../web/js/pgp/pgp-core.js');
    pgpCore.encryptMessage.mockResolvedValue(
      '-----BEGIN PGP MESSAGE-----\nencrypted\n-----END PGP MESSAGE-----',
    );

    const { maybeAutoEncryptForTest } = await import('../web/MessageCompose.js');
    await maybeAutoEncryptForTest();
    expect(pgpCore.encryptMessage).toHaveBeenCalledTimes(1);

    await maybeAutoEncryptForTest();
    expect(pgpCore.encryptMessage).toHaveBeenCalledTimes(1); // still 1, not 2
  });

  it('waits and retries when a recipient has no key yet, then fires once one appears', async () => {
    vi.useFakeTimers();
    const keyStorage = await import('../web/js/pgp/key-storage.js');
    keyStorage.getAutoEncryptDefault.mockReturnValue(true);
    keyStorage.getAutoSendDefault.mockReturnValue(false);

    const { statusEl } = installStubs({
      bodyText: '<p>hello</p>',
      recipients: [{ emailAddress: 'slow@example.com' }],
    });
    const keyDiscovery = await import('../web/js/pgp/key-discovery.js');
    keyDiscovery.resolveRecipients
      .mockResolvedValueOnce([{ email: 'slow@example.com', key: null, status: 'not-found', source: null, armoredKey: null }])
      .mockResolvedValueOnce([{ email: 'slow@example.com', key: { fake: 'recipient-key' }, status: 'found', source: 'keyring', armoredKey: null }]);
    const pgpCore = await import('../web/js/pgp/pgp-core.js');
    pgpCore.encryptMessage.mockResolvedValue(
      '-----BEGIN PGP MESSAGE-----\nencrypted\n-----END PGP MESSAGE-----',
    );

    const { maybeAutoEncryptForTest } = await import('../web/MessageCompose.js');
    const runPromise = maybeAutoEncryptForTest();

    // maybeAutoEncrypt()'s own first action is a fresh loadRecipients() call,
    // which internally polls Outlook's recipients collection twice (via
    // getRecipientsAsync's own ~300ms internal retry) before resolving — so
    // even this FIRST pass needs a timer advance to complete, not just the
    // 2s wait between auto-encrypt's own poll passes below. 500ms safely
    // covers that internal ~300ms delay.
    await vi.advanceTimersByTimeAsync(500);
    expect(statusEl.textContent).toContain('Waiting for all recipients to resolve');

    // Advance past the 2s poll interval (which itself contains another
    // internal ~300ms recipient-poll delay, safely covered within this
    // window) to trigger the retry pass, which now finds the key.
    await vi.advanceTimersByTimeAsync(2500);
    await runPromise;

    expect(pgpCore.encryptMessage).toHaveBeenCalledTimes(1);
    vi.useRealTimers();
  });

  it('gives up and shows a notice when no progress is made between two consecutive passes', async () => {
    vi.useFakeTimers();
    const keyStorage = await import('../web/js/pgp/key-storage.js');
    keyStorage.getAutoEncryptDefault.mockReturnValue(true);

    const { statusEl } = installStubs({
      bodyText: '<p>hello</p>',
      recipients: [{ emailAddress: 'nokey@example.com' }],
    });
    const keyDiscovery = await import('../web/js/pgp/key-discovery.js');
    keyDiscovery.resolveRecipients.mockResolvedValue(
      [{ email: 'nokey@example.com', key: null, status: 'not-found', source: null, armoredKey: null }],
    );
    const pgpCore = await import('../web/js/pgp/pgp-core.js');

    const { maybeAutoEncryptForTest } = await import('../web/MessageCompose.js');
    const runPromise = maybeAutoEncryptForTest();

    await vi.advanceTimersByTimeAsync(500); // flush the initial loadRecipients() call
    await vi.advanceTimersByTimeAsync(2500); // second pass: identical result -> give up
    await runPromise;

    expect(pgpCore.encryptMessage).not.toHaveBeenCalled();
    expect(statusEl.textContent).toContain("didn't complete");
    vi.useRealTimers();
  });

  it('aborts without firing when inline attachments are present', async () => {
    const keyStorage = await import('../web/js/pgp/key-storage.js');
    keyStorage.getAutoEncryptDefault.mockReturnValue(true);

    const { statusEl } = installStubs({
      bodyText: '<p>hello <img src="cid:abc123"></p>',
      recipients: [{ emailAddress: 'friend@example.com' }],
    });
    const pgpCore = await import('../web/js/pgp/pgp-core.js');

    const { maybeAutoEncryptForTest } = await import('../web/MessageCompose.js');
    await maybeAutoEncryptForTest();

    expect(pgpCore.encryptMessage).not.toHaveBeenCalled();
    expect(statusEl.textContent).toContain('inline images');
  });

  it('aborts without firing when regular attachments are present on a host below Mailbox 1.8', async () => {
    const keyStorage = await import('../web/js/pgp/key-storage.js');
    keyStorage.getAutoEncryptDefault.mockReturnValue(true);

    // _has18 is only ever set inside Office.onReady, which this test file's
    // stubbed Office.onReady never invokes -- it stays at its module-default
    // `false` throughout every test here, so a non-empty, non-inline
    // attachment list is exactly what's needed to exercise this branch;
    // isSetSupported doesn't need overriding for this one.
    const { statusEl } = installStubs({
      bodyText: '<p>hello</p>',
      recipients: [{ emailAddress: 'friend@example.com' }],
      attachments: [{ id: 'a1', name: 'report.pdf', isInline: false }],
    });
    const pgpCore = await import('../web/js/pgp/pgp-core.js');

    const { maybeAutoEncryptForTest } = await import('../web/MessageCompose.js');
    await maybeAutoEncryptForTest();

    expect(pgpCore.encryptMessage).not.toHaveBeenCalled();
    expect(statusEl.textContent).toContain("can't encrypt attachments");
  });

  it('does not report ready with zero recipients if the recipient list empties out mid-wait', async () => {
    vi.useFakeTimers();
    const keyStorage = await import('../web/js/pgp/key-storage.js');
    keyStorage.getAutoEncryptDefault.mockReturnValue(true);
    keyStorage.getAutoSendDefault.mockReturnValue(false);

    const { statusEl } = installStubs({
      bodyText: '<p>hello</p>',
      recipients: [{ emailAddress: 'nokey@example.com' }],
    });

    // Simulate the recipient being removed from To/Cc entirely partway
    // through the wait loop: the first loadRecipients() poll (getAsync needs
    // two consecutive equal-length reads to settle, per getRecipientsAsync())
    // still sees the recipient; every loadRecipients() poll after that sees
    // an empty To field -- loadRecipients()'s own "reset _recipientResults
    // to []" branch (it never calls resolveRecipients() in that branch).
    let toCallCount = 0;
    global.Office.context.mailbox.item.to.getAsync = vi.fn((cb) => {
      toCallCount++;
      const value = toCallCount <= 2 ? [{ emailAddress: 'nokey@example.com' }] : [];
      cb({ status: 'succeeded', value });
    });

    const keyDiscovery = await import('../web/js/pgp/key-discovery.js');
    keyDiscovery.resolveRecipients.mockResolvedValueOnce(
      [{ email: 'nokey@example.com', key: null, status: 'not-found', source: null, armoredKey: null }],
    );
    const pgpCore = await import('../web/js/pgp/pgp-core.js');

    const { maybeAutoEncryptForTest } = await import('../web/MessageCompose.js');
    const runPromise = maybeAutoEncryptForTest();

    // Generously advance past: the initial loadRecipients() poll, then two
    // more 2s wait-loop passes (one that observes the list going empty, one
    // more so the give-up snapshot comparison -- empty vs empty -- settles).
    await vi.advanceTimersByTimeAsync(500);
    await vi.advanceTimersByTimeAsync(3000);
    await vi.advanceTimersByTimeAsync(3000);
    await vi.advanceTimersByTimeAsync(3000);
    await runPromise;

    expect(pgpCore.encryptMessage).not.toHaveBeenCalled();
    // With the length > 0 guard dropped, waitForAllRecipientKeys() would see
    // _recipientResults reset to [] by loadRecipients() and [].every(...)
    // vacuously report "ready" -- maybeAutoEncrypt() would then proceed into
    // handleAutoEncrypt()/handleEncrypt(), which re-checks recipients itself
    // and fails with "Encryption failed: Not all recipients...". With the
    // guard restored, waitForAllRecipientKeys() never reports ready here, so
    // handleEncrypt() is never even entered and the status bar never reaches
    // that failure text.
    expect(statusEl.textContent).not.toContain('Encryption failed');
    vi.useRealTimers();
  });
});

describe('auto-send after auto-encrypt', () => {
  beforeEach(async () => {
    // Same reasoning as the 'auto-encrypt on pane load' describe above:
    // maybeAutoEncrypt() only ever fires once per module instance
    // (_autoEncryptFired), so each test here needs its own fresh module.
    vi.resetModules();
    clearSessionKey();

    const keyDiscovery = await import('../web/js/pgp/key-discovery.js');
    keyDiscovery.resolveRecipients.mockReset();
    keyDiscovery.resolveRecipients.mockImplementation(async (emails) => emails.map((email) => (
      { email, key: { fake: 'recipient-key' }, status: 'found', source: 'keyring', armoredKey: null }
    )));
  });

  it('calls item.sendAsync when auto-send is on, the host supports Mailbox 1.15, and auto-encrypt just succeeded', async () => {
    const keyStorage = await import('../web/js/pgp/key-storage.js');
    keyStorage.getAutoEncryptDefault.mockReturnValue(true);
    keyStorage.getAutoSendDefault.mockReturnValue(true);

    installStubs({
      bodyText: '<p>hello</p>',
      recipients: [{ emailAddress: 'friend@example.com' }],
    });
    global.Office.context.requirements.isSetSupported = () => true;
    const sendAsync = vi.fn((cb) => cb({ status: 'succeeded' }));
    global.Office.context.mailbox.item.sendAsync = sendAsync;

    const pgpCore = await import('../web/js/pgp/pgp-core.js');
    pgpCore.encryptMessage.mockResolvedValue(
      '-----BEGIN PGP MESSAGE-----\nencrypted\n-----END PGP MESSAGE-----',
    );

    const { maybeAutoEncryptForTest } = await import('../web/MessageCompose.js');
    await maybeAutoEncryptForTest();

    expect(sendAsync).toHaveBeenCalledTimes(1);
  });

  it('does not call sendAsync when the host does not support Mailbox 1.15, even if auto-send is on', async () => {
    const keyStorage = await import('../web/js/pgp/key-storage.js');
    keyStorage.getAutoEncryptDefault.mockReturnValue(true);
    keyStorage.getAutoSendDefault.mockReturnValue(true);

    installStubs({
      bodyText: '<p>hello</p>',
      recipients: [{ emailAddress: 'friend@example.com' }],
    });
    // isSetSupported returning false for 1.15 specifically simulates an
    // older host that supports enough for auto-encrypt but not sendAsync.
    global.Office.context.requirements.isSetSupported = (family, version) => version !== '1.15';
    const sendAsync = vi.fn((cb) => cb({ status: 'succeeded' }));
    global.Office.context.mailbox.item.sendAsync = sendAsync;

    const pgpCore = await import('../web/js/pgp/pgp-core.js');
    pgpCore.encryptMessage.mockResolvedValue(
      '-----BEGIN PGP MESSAGE-----\nencrypted\n-----END PGP MESSAGE-----',
    );

    const { maybeAutoEncryptForTest } = await import('../web/MessageCompose.js');
    await maybeAutoEncryptForTest();

    expect(sendAsync).not.toHaveBeenCalled();
  });

  it('does not call sendAsync when auto-send is off', async () => {
    const keyStorage = await import('../web/js/pgp/key-storage.js');
    keyStorage.getAutoEncryptDefault.mockReturnValue(true);
    keyStorage.getAutoSendDefault.mockReturnValue(false);

    installStubs({
      bodyText: '<p>hello</p>',
      recipients: [{ emailAddress: 'friend@example.com' }],
    });
    global.Office.context.requirements.isSetSupported = () => true;
    const sendAsync = vi.fn((cb) => cb({ status: 'succeeded' }));
    global.Office.context.mailbox.item.sendAsync = sendAsync;

    const pgpCore = await import('../web/js/pgp/pgp-core.js');
    pgpCore.encryptMessage.mockResolvedValue(
      '-----BEGIN PGP MESSAGE-----\nencrypted\n-----END PGP MESSAGE-----',
    );

    const { maybeAutoEncryptForTest } = await import('../web/MessageCompose.js');
    await maybeAutoEncryptForTest();

    expect(sendAsync).not.toHaveBeenCalled();
  });

  it('shows an error status and does not throw when sendAsync itself fails', async () => {
    const keyStorage = await import('../web/js/pgp/key-storage.js');
    keyStorage.getAutoEncryptDefault.mockReturnValue(true);
    keyStorage.getAutoSendDefault.mockReturnValue(true);

    const { statusEl } = installStubs({
      bodyText: '<p>hello</p>',
      recipients: [{ emailAddress: 'friend@example.com' }],
    });
    global.Office.context.requirements.isSetSupported = () => true;
    const sendAsync = vi.fn((cb) => cb({ status: 'failed', error: { message: 'blocked by another add-in' } }));
    global.Office.context.mailbox.item.sendAsync = sendAsync;

    const pgpCore = await import('../web/js/pgp/pgp-core.js');
    pgpCore.encryptMessage.mockResolvedValue(
      '-----BEGIN PGP MESSAGE-----\nencrypted\n-----END PGP MESSAGE-----',
    );

    const { maybeAutoEncryptForTest } = await import('../web/MessageCompose.js');
    await maybeAutoEncryptForTest();

    expect(statusEl.textContent).toContain('Automatic send failed');
    expect(statusEl.textContent).toContain('blocked by another add-in');
  });

  it('does not call sendAsync when auto-encrypt itself fails', async () => {
    const keyStorage = await import('../web/js/pgp/key-storage.js');
    keyStorage.getAutoEncryptDefault.mockReturnValue(true);
    keyStorage.getAutoSendDefault.mockReturnValue(true);

    installStubs({
      bodyText: '<p>hello</p>',
      recipients: [{ emailAddress: 'friend@example.com' }],
    });
    global.Office.context.requirements.isSetSupported = () => true;
    const sendAsync = vi.fn((cb) => cb({ status: 'succeeded' }));
    global.Office.context.mailbox.item.sendAsync = sendAsync;

    const pgpCore = await import('../web/js/pgp/pgp-core.js');
    pgpCore.encryptMessage.mockRejectedValue(new Error('boom'));

    const { maybeAutoEncryptForTest } = await import('../web/MessageCompose.js');
    await maybeAutoEncryptForTest();

    expect(sendAsync).not.toHaveBeenCalled();
  });
});

describe('runForceEncryptAndSend', () => {
  beforeEach(async () => {
    vi.resetModules();
    clearSessionKey();

    const keyDiscovery = await import('../web/js/pgp/key-discovery.js');
    keyDiscovery.resolveRecipients.mockReset();
    keyDiscovery.resolveRecipients.mockImplementation(async (emails) => emails.map((email) => (
      { email, key: { fake: 'recipient-key' }, status: 'found', source: 'keyring', armoredKey: null }
    )));

    const keyStorage = await import('../web/js/pgp/key-storage.js');
    keyStorage.hasAcknowledgedWarning = vi.fn(() => false);
    keyStorage.saveAcknowledgedWarning = vi.fn(async () => {});
    // Reset explicitly each test -- vi.resetModules() doesn't recreate the
    // vi.mock()'d module object itself, so a per-test override (e.g. the
    // "no key pair" test below) would otherwise leak into later tests.
    keyStorage.hasKeyPair = vi.fn(() => true);
  });

  it('shows the confirmation panel and does nothing until confirmed', async () => {
    const { encryptSendConfirmPanel, encryptSendCancelBtn } = installStubs({
      bodyText: '<p>hello</p>',
      recipients: [{ emailAddress: 'friend@example.com' }],
    });
    const pgpCore = await import('../web/js/pgp/pgp-core.js');

    const { runForceEncryptAndSend } = await import('../web/MessageCompose.js');
    const runPromise = runForceEncryptAndSend();

    // Let the confirmation panel show before cancelling.
    for (let i = 0; i < 20 && !encryptSendConfirmPanel.classList.remove.mock.calls.length; i++) {
      await Promise.resolve();
    }
    expect(encryptSendConfirmPanel.classList.remove).toHaveBeenCalledWith('pgp-hidden');

    encryptSendCancelBtn._cb();
    await runPromise;

    expect(pgpCore.encryptMessage).not.toHaveBeenCalled();
    const keyStorage = await import('../web/js/pgp/key-storage.js');
    expect(keyStorage.saveAcknowledgedWarning).not.toHaveBeenCalled();
  });

  it('shows an error and never shows the confirmation panel when the user has no key pair', async () => {
    const keyStorage = await import('../web/js/pgp/key-storage.js');
    keyStorage.hasKeyPair = vi.fn(() => false);

    const { encryptSendConfirmPanel, statusEl } = installStubs({
      bodyText: '<p>hello</p>',
      recipients: [{ emailAddress: 'friend@example.com' }],
    });

    const { runForceEncryptAndSend } = await import('../web/MessageCompose.js');
    await runForceEncryptAndSend();

    expect(encryptSendConfirmPanel.classList.remove).not.toHaveBeenCalledWith('pgp-hidden');
    expect(keyStorage.saveAcknowledgedWarning).not.toHaveBeenCalled();
    expect(statusEl.textContent).toMatch(/don't have a PGP key pair/);
  });

  it('encrypts and sends once confirmed, on a host that supports Mailbox 1.15', async () => {
    const { encryptSendConfirmPanel, encryptSendConfirmBtn } = installStubs({
      bodyText: '<p>hello</p>',
      recipients: [{ emailAddress: 'friend@example.com' }],
    });
    global.Office.context.requirements.isSetSupported = () => true;
    const sendAsync = vi.fn((cb) => cb({ status: 'succeeded' }));
    global.Office.context.mailbox.item.sendAsync = sendAsync;

    const pgpCore = await import('../web/js/pgp/pgp-core.js');
    pgpCore.encryptMessage.mockResolvedValue(
      '-----BEGIN PGP MESSAGE-----\nencrypted\n-----END PGP MESSAGE-----',
    );

    const { runForceEncryptAndSend } = await import('../web/MessageCompose.js');
    const runPromise = runForceEncryptAndSend();

    for (let i = 0; i < 20 && !encryptSendConfirmPanel.classList.remove.mock.calls.length; i++) {
      await Promise.resolve();
    }
    encryptSendConfirmBtn._cb();
    await runPromise;

    const keyStorage = await import('../web/js/pgp/key-storage.js');
    expect(keyStorage.saveAcknowledgedWarning).toHaveBeenCalledWith('encryptSendConfirm');
    expect(pgpCore.encryptMessage).toHaveBeenCalled();
    expect(sendAsync).toHaveBeenCalledTimes(1);
  });

  it('skips the confirmation panel entirely when already acknowledged', async () => {
    const keyStorage = await import('../web/js/pgp/key-storage.js');
    keyStorage.hasAcknowledgedWarning = vi.fn(() => true);

    const { encryptSendConfirmPanel } = installStubs({
      bodyText: '<p>hello</p>',
      recipients: [{ emailAddress: 'friend@example.com' }],
    });
    global.Office.context.requirements.isSetSupported = () => true;
    const sendAsync = vi.fn((cb) => cb({ status: 'succeeded' }));
    global.Office.context.mailbox.item.sendAsync = sendAsync;

    const pgpCore = await import('../web/js/pgp/pgp-core.js');
    pgpCore.encryptMessage.mockResolvedValue(
      '-----BEGIN PGP MESSAGE-----\nencrypted\n-----END PGP MESSAGE-----',
    );

    const { runForceEncryptAndSend } = await import('../web/MessageCompose.js');
    await runForceEncryptAndSend();

    expect(encryptSendConfirmPanel.classList.remove).not.toHaveBeenCalledWith('pgp-hidden');
    expect(sendAsync).toHaveBeenCalledTimes(1);
  });

  it('encrypts but does not send on a host below Mailbox 1.15, and shows an explanatory status', async () => {
    const keyStorage = await import('../web/js/pgp/key-storage.js');
    keyStorage.hasAcknowledgedWarning = vi.fn(() => true);

    const { statusEl } = installStubs({
      bodyText: '<p>hello</p>',
      recipients: [{ emailAddress: 'friend@example.com' }],
    });
    global.Office.context.requirements.isSetSupported = () => false;
    const sendAsync = vi.fn((cb) => cb({ status: 'succeeded' }));
    global.Office.context.mailbox.item.sendAsync = sendAsync;

    const pgpCore = await import('../web/js/pgp/pgp-core.js');
    pgpCore.encryptMessage.mockResolvedValue(
      '-----BEGIN PGP MESSAGE-----\nencrypted\n-----END PGP MESSAGE-----',
    );

    const { runForceEncryptAndSend } = await import('../web/MessageCompose.js');
    await runForceEncryptAndSend();

    expect(pgpCore.encryptMessage).toHaveBeenCalled();
    expect(sendAsync).not.toHaveBeenCalled();
    expect(statusEl.textContent).toMatch(/doesn't support automatic sending/);
  });

  it('does not send when encryption itself fails', async () => {
    const keyStorage = await import('../web/js/pgp/key-storage.js');
    keyStorage.hasAcknowledgedWarning = vi.fn(() => true);

    installStubs({
      bodyText: '<p>hello</p>',
      recipients: [{ emailAddress: 'friend@example.com' }],
    });
    global.Office.context.requirements.isSetSupported = () => true;
    const sendAsync = vi.fn((cb) => cb({ status: 'succeeded' }));
    global.Office.context.mailbox.item.sendAsync = sendAsync;

    const pgpCore = await import('../web/js/pgp/pgp-core.js');
    pgpCore.encryptMessage.mockRejectedValue(new Error('boom'));

    const { runForceEncryptAndSend } = await import('../web/MessageCompose.js');
    await runForceEncryptAndSend();

    expect(sendAsync).not.toHaveBeenCalled();
  });

  it('force-disables the ordinary Encrypt button for the duration of the recipient-key wait, and restores it afterward', async () => {
    // Regression: a manual click on btn-encrypt while runForceEncryptAndSend()
    // is still inside waitForAllRecipientKeys() could race ahead of the
    // forced flow's own handleEncrypt() call, causing that call to hit the
    // "already encrypted" bailout and silently stop without ever sending.
    vi.useFakeTimers();
    const keyStorage = await import('../web/js/pgp/key-storage.js');
    keyStorage.hasAcknowledgedWarning = vi.fn(() => true);

    const { encryptBtn } = installStubs({
      bodyText: '<p>hello</p>',
      recipients: [{ emailAddress: 'slow@example.com' }],
    });
    global.Office.context.requirements.isSetSupported = () => true;
    const sendAsync = vi.fn((cb) => cb({ status: 'succeeded' }));
    global.Office.context.mailbox.item.sendAsync = sendAsync;

    const keyDiscovery = await import('../web/js/pgp/key-discovery.js');
    keyDiscovery.resolveRecipients
      .mockResolvedValueOnce([{ email: 'slow@example.com', key: null, status: 'not-found', source: null, armoredKey: null }])
      .mockResolvedValueOnce([{ email: 'slow@example.com', key: { fake: 'recipient-key' }, status: 'found', source: 'keyring', armoredKey: null }]);

    const pgpCore = await import('../web/js/pgp/pgp-core.js');
    pgpCore.encryptMessage.mockResolvedValue(
      '-----BEGIN PGP MESSAGE-----\nencrypted\n-----END PGP MESSAGE-----',
    );

    const { runForceEncryptAndSend } = await import('../web/MessageCompose.js');
    const runPromise = runForceEncryptAndSend();

    // Flush the initial loadRecipients() poll (no key yet) -- btn-encrypt
    // must already be forced disabled at this point, before any recipient
    // has resolved, since the flag is set as the very first statement.
    await vi.advanceTimersByTimeAsync(500);
    expect(encryptBtn.disabled).toBe(true);

    // Advance past the 2s wait-loop retry, which now finds the key and lets
    // the flow proceed into its own handleEncrypt()/sendAsync() call.
    await vi.advanceTimersByTimeAsync(2500);
    await runPromise;

    expect(pgpCore.encryptMessage).toHaveBeenCalledTimes(1);
    expect(sendAsync).toHaveBeenCalledTimes(1);
    // The flag is cleared in the finally block and updateEncryptButton() is
    // called once more -- with a resolved recipient and a key pair, the
    // button should now reflect a normal ready state (re-enabled), not the
    // forced-disabled state from during the wait.
    expect(encryptBtn.disabled).toBe(false);

    vi.useRealTimers();
  });

  it('disables the ordinary Encrypt button synchronously, before the confirmation panel or any await runs', async () => {
    // Regression: setting the _forceEncryptSendActive flag alone doesn't
    // repaint anything -- without an immediate updateEncryptButton() call,
    // btn-encrypt stays in whatever state the initial pane-load
    // loadRecipients() left it (typically enabled) through the confirmation
    // panel's await and the hasKeyPair()/hasAcknowledgedWarning() checks,
    // since nothing else calls updateEncryptButton() until
    // waitForAllRecipientKeys()'s own loadRecipients() polls run.
    const { encryptBtn } = installStubs({
      bodyText: '<p>hello</p>',
      recipients: [{ emailAddress: 'friend@example.com' }],
    });

    const { runForceEncryptAndSend } = await import('../web/MessageCompose.js');

    // Deliberately do NOT await -- a JS function runs synchronously up to
    // its first await, so checking encryptBtn.disabled right here proves
    // whether the disable happens before or only after some later await
    // (e.g. inside the confirmation panel or the recipient-wait loop).
    runForceEncryptAndSend();

    expect(encryptBtn.disabled).toBe(true);
  });
});
