# Compose-window Decrypt Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a "Decrypt" button to the compose pane (`web/MessageCompose.js` / `.html`) that reverses the add-in's own Encrypt action — restoring the original body and any encrypted attachments — so a user can edit recipients/body/attachments after encrypting, per issue #25.

**Architecture:** A new `refreshComposeButtons()` function drives Encrypt/Decrypt visibility from `detectPgpContent()` on the current body. `handleDecrypt()` mirrors `handleEncrypt()`'s structure: unlock the private key (session cache or passphrase prompt, reusing the existing modal), call `decryptMessage()` to restore the body, then best-effort reverse each `.pgp` attachment via `decryptAttachment()`. Encrypting to your own public key is always done unconditionally by `handleEncrypt()`, so the sender's own private key is always sufficient here regardless of recipients.

**Tech Stack:** Plain ES modules, Office.js Mailbox APIs, OpenPGP.js (via `web/js/pgp/pgp-core.js`), Vitest for tests.

## Global Constraints

- No changes to `MessageRead.js`, `KeyManagement.js`, or the manifest.
- No change to how Encrypt itself works — Decrypt is purely additive.
- Do not attempt to restore native Eml/ICalendar attachment types — decrypted attachments are always re-added as plain files (matches `MessageRead.js`'s existing decrypt-and-download behavior).
- Do not apply `companyDecryptedExtensionPrefix` when recovering attachment filenames in this feature — that setting is for the recipient's decrypt-and-download naming, not for restoring the sender's own original file.
- Attachment reversal is best-effort per attachment: one failure must not stop the others or roll back the body or already-reverted attachments.
- Follow the existing plain-object Office/document stub test pattern used in `tests/message-read-native-reply-handoff.test.js` — no real DOM/jsdom needed for this feature's tests.

---

## File Structure

- Modify: `web/MessageCompose.html` — add `#btn-decrypt` button + spinner next to `#btn-encrypt`; add `id="passphrase-modal-msg"` to the existing passphrase modal's message `<p>`.
- Modify: `web/MessageCompose.js` — add `decryptMessage`, `decryptAttachment`, `stripPgpExtension`, `uint8ArrayToBase64` to the `pgp-core.js` import; parameterize `promptPassphrase(message)`; add `refreshComposeButtons()`; add `handleDecrypt()`; wire the new button; export both new functions for testing.
- Create: `tests/message-compose-decrypt.test.js` — covers `refreshComposeButtons()` and `handleDecrypt()` using mocked `pgp-core.js`/`key-storage.js` and a plain-object Office/document stub (no jsdom).

---

## Task 1: HTML — add the Decrypt button and passphrase modal message id

**Files:**
- Modify: `web/MessageCompose.html:108-116` (Encrypt button block), `web/MessageCompose.html:170-185` (passphrase modal)

**Interfaces:**
- Produces: DOM element ids `btn-decrypt`, `decrypt-spinner`, `passphrase-modal-msg` — consumed by Task 2/3/4's JS changes.

- [ ] **Step 1: Add the Decrypt button markup**

Replace the Encrypt button block (`web/MessageCompose.html:108-116`):

```html
  <!-- ═══════════════════════════════════════════════════════
       ENCRYPT / DECRYPT BUTTON
  ══════════════════════════════════════════════════════════ -->
  <div>
    <button class="pgp-btn pgp-btn--primary pgp-btn--full" id="btn-encrypt" disabled>
      <span class="pgp-spinner pgp-hidden" id="encrypt-spinner"></span>
      Encrypt Message
    </button>
    <button class="pgp-btn pgp-btn--secondary pgp-btn--full pgp-hidden" id="btn-decrypt">
      <span class="pgp-spinner pgp-hidden" id="decrypt-spinner"></span>
      Decrypt Message
    </button>
    <p style="font-size:11px;color:#605e5c;margin:6px 0 0;text-align:center;">
      After encrypting, click <strong>Send</strong> in Outlook as normal.
    </p>
  </div>
```

- [ ] **Step 2: Give the passphrase modal's message paragraph an id**

In `web/MessageCompose.html:170-185`, replace:

```html
      <p style="font-size:13px;color:#605e5c;margin:0 0 10px;">
        Your private key passphrase is required to sign and encrypt this message.
      </p>
```

with:

```html
      <p style="font-size:13px;color:#605e5c;margin:0 0 10px;" id="passphrase-modal-msg">
        Your private key passphrase is required to sign and encrypt this message.
      </p>
```

- [ ] **Step 3: Verify the HTML is well-formed**

Run: `node -e "require('fs').readFileSync('web/MessageCompose.html','utf8')" && echo OK`
Expected: `OK` (this just proves the file is still readable/no corruption from the edit; there is no HTML linter in this repo).

- [ ] **Step 4: Commit**

```bash
git add web/MessageCompose.html
git commit -m "feat: add Decrypt button markup to compose pane (#25)"
```

---

## Task 2: `refreshComposeButtons()` — Encrypt/Decrypt visibility driven by body state

**Files:**
- Modify: `web/MessageCompose.js` (import block near line 33-39; add function near `updateEncryptButton()` at line ~370; export list at line ~40)
- Test: `tests/message-compose-decrypt.test.js` (new file)

**Interfaces:**
- Consumes: `detectPgpContent(text)` from `pgp-core.js` (already imported) — returns `'encrypted' | 'signed' | 'public-key' | 'private-key' | null`.
- Produces: `export async function refreshComposeButtons()` — reads the body via `Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, cb)`, toggles `pgp-hidden` on `#btn-encrypt` / `#btn-decrypt`. Consumed by Task 4 (called at the end of `handleDecrypt()`) and Task 6 (called in `Office.onReady` and at the end of `handleEncrypt()`'s `finally` block).

- [ ] **Step 1: Write the failing test**

Create `tests/message-compose-decrypt.test.js`:

```javascript
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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run tests/message-compose-decrypt.test.js`
Expected: FAIL — `refreshComposeButtons is not a function` (it isn't exported yet).

- [ ] **Step 3: Add the import and `refreshComposeButtons()` implementation**

In `web/MessageCompose.js`, replace the `pgp-core.js` import block (lines 33-39):

```javascript
import {
  unlockPrivateKey, readPublicKey, getKeyInfo,
  encryptMessage, encryptAttachment,
  decryptMessage, decryptAttachment,
  hasWeakEncryptionKey,
  base64ToUint8Array, uint8ArrayToBase64, stripPgpExtension,
  detectPgpContent,
} from './js/pgp/pgp-core.js';
```

Update the export line (was `export { stripPgpArmorBlock, pickSpliceMarker };`):

```javascript
export { stripPgpArmorBlock, pickSpliceMarker, refreshComposeButtons, handleDecrypt };
```

Add this function immediately after `updateEncryptButton()` (after line 375, i.e. right after its closing `}`):

```javascript
/**
 * Show Decrypt / hide Encrypt when the body is currently PGP-armored, and
 * vice versa. Called on load and after every Encrypt/Decrypt action so the
 * two buttons never both suggest an available action at once.
 */
async function refreshComposeButtons() {
  const bodyText = await getBodyAsync(Office.CoercionType.Text);
  const isEncrypted = detectPgpContent(bodyText) === 'encrypted';
  if (isEncrypted) {
    el('btn-decrypt').classList.remove('pgp-hidden');
    el('btn-encrypt').classList.add('pgp-hidden');
  } else {
    el('btn-encrypt').classList.remove('pgp-hidden');
    el('btn-decrypt').classList.add('pgp-hidden');
  }
}
```

Note: `getBodyAsync` is defined later in the file (line 757) but is a hoisted `function` declaration, so calling it here before its textual definition is valid.

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run tests/message-compose-decrypt.test.js`
Expected: PASS (2 tests)

- [ ] **Step 5: Commit**

```bash
git add web/MessageCompose.js tests/message-compose-decrypt.test.js
git commit -m "feat: add refreshComposeButtons() to toggle Encrypt/Decrypt visibility (#25)"
```

---

## Task 3: Parameterize `promptPassphrase()` with a message

**Files:**
- Modify: `web/MessageCompose.js:381` (`promptPassphrase` definition) and its call site at line ~469 inside `handleEncrypt()`
- Test: `tests/message-compose-decrypt.test.js`

**Interfaces:**
- Produces: `promptPassphrase(message = 'Your private key passphrase is required to sign and encrypt this message.')` — sets `#passphrase-modal-msg`'s `textContent` before showing the modal. Consumed by Task 4's `handleDecrypt()`.

- [ ] **Step 1: Write the failing test**

Add to `tests/message-compose-decrypt.test.js` (new `describe` block, after the `refreshComposeButtons` block):

```javascript
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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run tests/message-compose-decrypt.test.js`
Expected: FAIL — `promptPassphraseForTest is not a function` (not exported yet).

- [ ] **Step 3: Parameterize `promptPassphrase` and export a test alias**

In `web/MessageCompose.js`, change the function signature (line 381) from:

```javascript
function promptPassphrase() {
  return new Promise((resolve, reject) => {
    const modal = el('passphrase-modal');
    const input = el('passphrase-input');
    const errEl = el('passphrase-error');
```

to:

```javascript
function promptPassphrase(message = 'Your private key passphrase is required to sign and encrypt this message.') {
  return new Promise((resolve, reject) => {
    const modal = el('passphrase-modal');
    const input = el('passphrase-input');
    const errEl = el('passphrase-error');
    el('passphrase-modal-msg').textContent = message;
```

Update the export line from Task 2 to also export a test-only alias:

```javascript
export { stripPgpArmorBlock, pickSpliceMarker, refreshComposeButtons, handleDecrypt, promptPassphrase as promptPassphraseForTest };
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run tests/message-compose-decrypt.test.js`
Expected: PASS (3 tests)

- [ ] **Step 5: Commit**

```bash
git add web/MessageCompose.js tests/message-compose-decrypt.test.js
git commit -m "feat: parameterize compose passphrase modal message (#25)"
```

---

## Task 4: `handleDecrypt()` — body restore

**Files:**
- Modify: `web/MessageCompose.js` (new function after `handleEncrypt()`, ends at line ~588)
- Test: `tests/message-compose-decrypt.test.js`

**Interfaces:**
- Consumes: `getBodyAsync`, `setBodyHtmlAsync` (both already defined in the file), `getSessionKey`/`cacheSessionKey`/`updateSessionStatus` (already imported), `promptPassphrase(message)` from Task 3, `unlockPrivateKey`, `decryptMessage` from `pgp-core.js`, `refreshComposeButtons()` from Task 2.
- Produces: `async function handleDecrypt()` (already in the Task 2 export line) — on success restores the body and calls `refreshComposeButtons()`; attachment reversal is added in Task 5.

- [ ] **Step 1: Write the failing test**

Add to `tests/message-compose-decrypt.test.js` (below the existing `describe` blocks — `cacheSessionKey`/`clearSessionKey` are already imported at the top of the file from Task 2):

```javascript
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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run tests/message-compose-decrypt.test.js`
Expected: FAIL — `handleDecrypt is not a function` (or, once exported as an empty stub, assertion failures — implement fully in Step 3 rather than stubbing).

- [ ] **Step 3: Implement `handleDecrypt()` (body-only for now)**

Add this function in `web/MessageCompose.js` immediately after the closing `}` of `handleEncrypt()` (after line 588):

```javascript
async function handleDecrypt() {
  clearStatus();
  const btn = el('btn-decrypt');
  const spinner = el('decrypt-spinner');
  btn.disabled = true;
  spinner.classList.remove('pgp-hidden');

  try {
    const bodyText = await getBodyAsync(Office.CoercionType.Text);
    if (detectPgpContent(bodyText) !== 'encrypted') {
      showStatus('This message does not appear to be PGP-encrypted.', 'error');
      return;
    }

    let privateKey = getSessionKey();
    if (!privateKey) {
      const passphrase = await promptPassphrase('Enter your passphrase to decrypt this message.');
      privateKey = await unlockPrivateKey(getPrivateKey(), passphrase);
      const userEmail = Office.context.mailbox.userProfile?.emailAddress || '';
      const keyInfo = await getKeyInfo(getPublicKey());
      cacheSessionKey(privateKey, userEmail, keyInfo.shortId);
      updateSessionStatus();
    }

    showStatus('Decrypting message body…', 'info');
    const { data: originalHtml } = await decryptMessage(bodyText, privateKey);
    await setBodyHtmlAsync(originalHtml);

    showStatus('✓ Message decrypted.', 'success');
  } catch (e) {
    if (e.message === 'Cancelled by user.') {
      showStatus('Decryption cancelled.', 'info');
    } else {
      showStatus(`Decryption failed: ${e.message}`, 'error');
      console.error(e);
    }
  } finally {
    spinner.classList.add('pgp-hidden');
    await refreshComposeButtons();
    el('btn-decrypt').disabled = false;
  }
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run tests/message-compose-decrypt.test.js`
Expected: PASS (6 tests)

- [ ] **Step 5: Commit**

```bash
git add web/MessageCompose.js tests/message-compose-decrypt.test.js
git commit -m "feat: implement handleDecrypt() body restore (#25)"
```

---

## Task 5: `handleDecrypt()` — best-effort attachment reversal

**Files:**
- Modify: `web/MessageCompose.js` (extend `handleDecrypt()` added in Task 4)
- Test: `tests/message-compose-decrypt.test.js`

**Interfaces:**
- Consumes: `loadAttachments()`, `_attachments` (module-level state, already defined), `getAttachmentContentAsync`, `removeAttachmentAsync`, `addAttachmentFromBase64Async` (all already defined in the file), `decryptAttachment`, `stripPgpExtension`, `uint8ArrayToBase64` from `pgp-core.js`.
- Produces: `handleDecrypt()` now also reverses `.pgp` attachments and reports a warning status naming any it couldn't revert.

- [ ] **Step 1: Write the failing test**

Add to `tests/message-compose-decrypt.test.js`:

```javascript
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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run tests/message-compose-decrypt.test.js`
Expected: FAIL — attachments untouched, `decryptAttachment` never called (attachment reversal doesn't exist yet).

- [ ] **Step 3: Extend `handleDecrypt()` with attachment reversal**

In `web/MessageCompose.js`, replace the body of `handleDecrypt()`'s `try` block added in Task 4 — insert the attachment loop between the `setBodyHtmlAsync(originalHtml);` line and the final `showStatus('✓ Message decrypted.', 'success');` line:

```javascript
    showStatus('Decrypting message body…', 'info');
    const { data: originalHtml } = await decryptMessage(bodyText, privateKey);
    await setBodyHtmlAsync(originalHtml);

    await loadAttachments();
    const pgpAttachments = _attachments.filter(a => /\.pgp$/i.test(a.name));
    const failedAttachments = [];

    if (pgpAttachments.length > 0) {
      showStatus('Decrypting attachments…', 'info');
      const item = Office.context.mailbox.item;

      for (const att of pgpAttachments) {
        try {
          const contentResult = await getAttachmentContentAsync(item, att.id);
          const armoredMessage = atob(contentResult.content.replace(/[^\x00-\x7F]/g, ''));
          const { data: decryptedBytes, filename } = await decryptAttachment(armoredMessage, privateKey);
          const recoveredName = filename || stripPgpExtension(att.name);

          await removeAttachmentAsync(item, att.id);
          await addAttachmentFromBase64Async(item, uint8ArrayToBase64(decryptedBytes), recoveredName);
        } catch (e) {
          console.error(`Decrypt: failed to revert attachment "${att.name}"`, e);
          failedAttachments.push(att.name);
        }
      }

      await loadAttachments();
    }

    if (failedAttachments.length > 0) {
      showStatus(`✓ Body decrypted. Could not revert: ${failedAttachments.join(', ')}.`, 'warning');
    } else {
      showStatus('✓ Message decrypted.', 'success');
    }
```

This replaces the single `showStatus('✓ Message decrypted.', 'success');` line that Task 4 added at the end of the `try` block.

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run tests/message-compose-decrypt.test.js`
Expected: PASS (9 tests)

- [ ] **Step 5: Commit**

```bash
git add web/MessageCompose.js tests/message-compose-decrypt.test.js
git commit -m "feat: best-effort attachment reversal in handleDecrypt() (#25)"
```

---

## Task 6: Wire the button, integrate `refreshComposeButtons()` into load/encrypt, full suite check

**Files:**
- Modify: `web/MessageCompose.js` (`handleEncrypt()`'s `finally` block at lines ~582-587; `Office.onReady` wiring block at lines ~1015-1069)

**Interfaces:**
- Consumes: `refreshComposeButtons()` (Task 2), `handleDecrypt` (Task 4/5).
- Produces: fully wired feature — no further tasks depend on this one.

- [ ] **Step 1: Replace `handleEncrypt()`'s `finally` block to use `refreshComposeButtons()`**

Replace (current lines ~582-587):

```javascript
  } finally {
    // Re-enable if there was a non-passphrase error
    const encrypted = el('status-bar').classList.contains('pgp-alert--success');
    btn.disabled = encrypted; // keep disabled after success so user can't re-encrypt
    spinner.classList.add('pgp-hidden');
  }
```

with:

```javascript
  } finally {
    spinner.classList.add('pgp-hidden');
    await refreshComposeButtons();
    btn.disabled = false;
  }
```

`refreshComposeButtons()` now owns hiding the Encrypt button after a successful encrypt (by hiding it entirely, not just disabling it), so the old `status-bar`-sniffing logic is no longer needed.

- [ ] **Step 2: Wire the Decrypt button and call `refreshComposeButtons()` on load**

In the `Office.onReady` callback, after the existing line `el('btn-encrypt').addEventListener('click', handleEncrypt);` (line ~1061), add:

```javascript
  el('btn-decrypt').addEventListener('click', handleDecrypt);
```

And after the existing `loadAttachments();` call (line ~1040), add:

```javascript
  await refreshComposeButtons();
```

- [ ] **Step 3: Run the full test suite**

Run: `npm test`
Expected: All tests pass, including every test in `tests/message-compose-decrypt.test.js` and the pre-existing `tests/message-compose.test.js`.

- [ ] **Step 4: Commit**

```bash
git add web/MessageCompose.js
git commit -m "feat: wire Decrypt button and refresh button state on load/encrypt (#25)"
```

---

## Task 7: Update documentation

**Files:**
- Modify: `CLAUDE.md` (entry-points table description for `MessageCompose.html/.js`, and the "Encryption scope" section)
- Modify: `README.md` (if it documents the compose-window feature set — check for an existing "Encrypt" feature bullet list to extend)

**Interfaces:** None — documentation only.

- [ ] **Step 1: Update `CLAUDE.md`**

In the entry-points table, change the `MessageCompose.html/.js` row's Purpose column from:

```
Encrypt outgoing messages, manage recipient keys
```

to:

```
Encrypt/decrypt outgoing messages, manage recipient keys
```

Add a new paragraph at the end of the "Encryption scope" section:

```markdown
### Decrypting from Compose

Encrypting to your own public key is always done unconditionally alongside recipient/company keys (so the sender can read their own sent mail) — this also means the sender's own private key is always sufficient to reverse an encrypt performed from this add-in's own compose pane, regardless of which recipients were chosen. A "Decrypt" button (`refreshComposeButtons()`/`handleDecrypt()` in `MessageCompose.js`) appears whenever the compose body is currently PGP-armored, letting the user restore the original body and revert any `.pgp` attachments back to their originals so recipients, body, or attachments can be edited before re-encrypting (#25). Attachment reversal is best-effort per file: a `.pgp` attachment that fails to decrypt (e.g. hand-corrupted armor) is left alone and named in a warning status, without blocking the body restore or any other attachment.
```

- [ ] **Step 2: Check `README.md` for a feature list to extend**

Run: `grep -n -i "encrypt" README.md | head -20`

If a bullet list of compose-window features exists (e.g. "Encrypt message body", "Encrypt attachments"), add a bullet: `- Decrypt a message you've already encrypted, to edit recipients/body/attachments before re-sending`. If no such list exists in README.md, skip this step — CLAUDE.md is the authoritative doc per this repo's existing convention (see `feedback_docs_with_code.md` — CLAUDE.md/README are both kept current, but README's structure varies).

- [ ] **Step 3: Commit**

```bash
git add CLAUDE.md README.md
git commit -m "docs: document compose-window Decrypt feature (#25)"
```

---

## Self-Review Notes

- **Spec coverage:** UI visibility rule (Task 2), passphrase modal reuse/parameterization (Task 3), body decrypt (Task 4), best-effort attachment reversal + filename fallback (Task 5), wiring/integration (Task 6), docs (Task 7). All spec sections have a corresponding task.
- **Type/name consistency checked:** `refreshComposeButtons`, `handleDecrypt`, `promptPassphrase(message)` used identically across all tasks; `decryptMessage`/`decryptAttachment`/`stripPgpExtension`/`uint8ArrayToBase64` names match their actual exports in `web/js/pgp/pgp-core.js` (verified against source during design).
- **No placeholders:** every step has literal code, not descriptions.
