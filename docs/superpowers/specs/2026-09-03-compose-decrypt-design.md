# Compose-window Decrypt (Issue #25)

## Problem

Once a user hits **Encrypt** in the compose window, the body is replaced with PGP armor and (if attachments were present) each attachment is replaced with a `.pgp` file. There is currently no way back: if the user then realizes they need to change a recipient, tweak the body, or swap an attachment, they have no path except discarding the draft.

## Goal

Add a **Decrypt** button to the compose pane that reverses both the body encryption and any attachment encryption performed by this add-in's own Encrypt action, so the user can edit and re-encrypt.

## Why this works: self-encryption

`handleEncrypt()` (`web/MessageCompose.js`) always includes the sender's own public key in `allEncryptionKeys`, regardless of recipients or company-key settings (encrypting to yourself so you can read your own sent mail). This means the sender's own private key is *always* sufficient to decrypt a message this add-in encrypted from compose — decrypt-in-compose never depends on recipient selection.

## UI

- A new `btn-decrypt` button is added to `web/MessageCompose.html`, positioned alongside the existing `btn-encrypt` button.
- Visibility is state-driven, not a static toggle: a new `refreshComposeButtons()` function reads the current body via `getBodyAsync(Office.CoercionType.Text)` and runs the existing `detectPgpContent()` (`web/js/pgp/pgp-core.js`).
  - `detectPgpContent(...) === 'encrypted'` → show Decrypt, hide/disable Encrypt.
  - otherwise → show Encrypt (existing behavior), hide Decrypt.
- `refreshComposeButtons()` is called:
  - once during `Office.onReady` (covers reopening a draft that was already encrypted in a previous session), and
  - at the end of `handleEncrypt()` and `handleDecrypt()` (covers the immediate before/after of each action).
- The existing passphrase modal (`#passphrase-modal`) is reused for the decrypt passphrase prompt. `promptPassphrase()` gains an optional message parameter so the modal body text can read "...to decrypt this message" instead of "...to sign and encrypt this message" depending on the caller.

## `handleDecrypt()` flow

1. Clear status, disable the Decrypt button, show its spinner (mirrors `handleEncrypt()`'s button/spinner handling).
2. Read the body via `getBodyAsync(Office.CoercionType.Text)`. If `detectPgpContent(...) !== 'encrypted'`, show an error status and abort — this should be unreachable given the button's visibility rule, but guards against a stale UI state.
3. Obtain the unlocked private key:
   - Check `getSessionKey()` first (may already be cached from signing during Encrypt, or from KeyManagement).
   - If not cached, prompt via `promptPassphrase()`, then `unlockPrivateKey(getPrivateKey(), passphrase)`, then `cacheSessionKey(...)` and `updateSessionStatus()` — identical pattern to the signing branch already in `handleEncrypt()`.
4. `const { data } = await decryptMessage(armorText, unlockedKey);` — `data` is the original HTML body string (exactly what was passed into `encryptMessage()` at encrypt time). Signature verification is not surfaced to the user here: this is an internal round-trip of the user's own just-encrypted draft, not a received message, so there's no recipient-facing signature UX to show. A verification failure/absence is not treated as an error.
5. `await setBodyHtmlAsync(data);` — restores the original rich body, replacing the armored `<pre>` block.
6. Attachment reversal:
   - `await loadAttachments();` to get the current attachment list.
   - Filter `_attachments` to those whose `name` ends with `.pgp` (case-insensitive) — these are candidates for reversal.
   - For each candidate, independently (best-effort — a failure on one does not stop the others or roll back the body or any already-reverted attachment):
     a. `getAttachmentContentAsync(item, att.id)` → expect `Office.MailboxEnums.AttachmentContentFormat.Base64`; anything else is treated as a failure for this attachment.
     b. Base64-decode to get the armored ASCII text (`atob(contentResult.content)`).
     c. `const { data, filename } = await decryptAttachment(armoredText, unlockedKey);`
     d. Recovered filename: `filename || stripPgpExtension(att.name)` (reuses the existing exported helper from `pgp-core.js` — same fallback logic `MessageRead.js` already uses on the receive side). Do **not** apply `companyDecryptedExtensionPrefix` here — that setting governs the recipient's decrypt-and-download naming, not restoring the sender's own original file.
     e. `removeAttachmentAsync(item, att.id)`, then `addAttachmentFromBase64Async(item, base64FromBytes(data), recoveredName)`.
     f. On any failure in a–e, leave the `.pgp` attachment untouched and record `att.name` in a `failedAttachments` list; continue to the next candidate.
   - `await loadAttachments();` again to refresh `_attachments` / the UI list.
7. Status message:
   - All succeeded (or there were no `.pgp` attachments): `"✓ Message decrypted."` (success).
   - One or more attachments failed: `"✓ Body decrypted. Could not revert: <names>."` (warning) — body and successfully-reverted attachments are still kept; only the failed ones remain encrypted.
8. `refreshComposeButtons()`; re-enable button, hide spinner (`finally` block, mirroring `handleEncrypt()`).

## Known limitation (inherited from existing behavior, not new)

Attachments originally read via `AttachmentContentFormat.Eml` or `.ICalendar` (forwarded email items / calendar items) were encrypted with a corrected filename (e.g. `subject.eml.pgp`) but can only be **re-added as a plain file attachment** on decrypt — Office's `addFileAttachmentFromBase64Async` cannot recreate a native forwarded-item or calendar-item attachment. This is identical to the tradeoff `MessageRead.js`'s existing "Decrypt & Download" already makes for recipients; Decrypt-in-compose does not make this any worse, so it is not addressed here.

## Error handling

- Wrong/cancelled passphrase: same handling as `handleEncrypt()` — `Error('Cancelled by user.')` shows an info status and re-enables the button; any other unlock error shows an error status.
- Corrupted/non-decryptable body armor (shouldn't happen for a body this add-in itself produced, but could if the user hand-edited the `<pre>` block): `decryptMessage` throws, caught by the outer `try/catch`, shown as an error status, body left untouched, Decrypt button re-enabled.
- Attachment failures never throw out of `handleDecrypt()` — they're caught per-attachment as described in step 6f.

## Testing

New/extended tests in `tests/message-compose.test.js` following the file's existing plain-object-stub pattern (no real DOM needed):
- `refreshComposeButtons()`: encrypted body → Decrypt shown/Encrypt hidden, and vice versa.
- `handleDecrypt()` body-only path: mocked `decryptMessage` returning HTML, asserts `setBodyHtmlAsync` called with that HTML and button state flips back to Encrypt.
- Attachment reversal: mocked `_attachments` list with a mix of `.pgp` and non-`.pgp` names; asserts only `.pgp` ones are processed, decrypted content is re-added under the recovered name, and a filename-recovery fallback (`stripPgpExtension`) is exercised when `decryptAttachment` returns an empty `filename`.
- Best-effort partial failure: one `.pgp` attachment's `decryptAttachment` rejects — asserts the other attachment(s) still succeed, the failed one is left untouched, and the final status is the warning variant naming it.
- Passphrase-required path: session cache empty → `promptPassphrase()` invoked with decrypt-specific message, `unlockPrivateKey`/`cacheSessionKey` called, matching the existing signing-branch test pattern.

## Out of scope

- No changes to `MessageRead.js`, `KeyManagement.js`, or the manifest.
- No change to how Encrypt itself works — Decrypt is purely additive.
- No attempt to restore native Eml/ICalendar attachment types (see Known limitation above).
