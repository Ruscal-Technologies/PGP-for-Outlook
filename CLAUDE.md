# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project overview

A Microsoft Outlook add-in (MailApp) that provides end-to-end PGP encryption in Outlook web, desktop, and mobile. It is a **zero-build-step** static web app — plain HTML/CSS/ES modules served over HTTPS. There is no npm, no bundler, no transpilation step.

## Local development

```bash
# One-time: install a dev HTTPS server and trust its self-signed cert
npm install -g office-addin-dev-certs http-server
office-addin-dev-certs install

# Serve the add-in — http-server's --ssl flag defaults to looking for
# cert.pem/key.pem in the current directory, NOT the certs office-addin-dev-certs
# just installed, so those must be pointed to explicitly or the server fails
# with "Could not find certificate cert.pem":
http-server web --ssl --cert "$HOME/.office-addin-dev-certs/localhost.crt" --key "$HOME/.office-addin-dev-certs/localhost.key" --port 3000
```

PowerShell equivalent for the last command:
```powershell
http-server web --ssl --cert "$env:USERPROFILE\.office-addin-dev-certs\localhost.crt" --key "$env:USERPROFILE\.office-addin-dev-certs\localhost.key" --port 3000
```

Update `manifest/manifest.xml` to point at `https://localhost:3000/`, then sideload it in Outlook following Microsoft's [sideloading guide](https://docs.microsoft.com/en-us/office/dev/add-ins/testing/test-debug-office-add-ins). Sideloading (not installing from Microsoft Marketplace/the Office Store) is Microsoft's own documented method for testing event-based-activation features (see `web/ReplyHandoffRuntime.js`) locally — it works on classic Outlook Windows, new Outlook Windows, web, and Mac.

### Tests

A Vitest suite in `tests/` covers the shared business-logic modules in `web/js/pgp/*.js` plus `web/js/wkd.js` — crypto round trips including multi-recipient, empty-string, unicode, and RSA-4096 (`pgp-core.js`), tampering/wrong-key failure modes (corrupted ciphertext, decrypting with a non-recipient key, malformed armor), legacy DSA+ElGamal key support (`hasWeakEncryptionKey`, `hasModernSubkeys`, `addModernSubkeys`, `extractPublicKey`) against a real fixture key in `tests/fixtures/`, storage (`key-storage.js`, mocked `Office.context.roamingSettings`, including its `saveAsync` failure path), the contact keyring (`keyring.js`), key discovery precedence and fallbacks (`key-discovery.js`, mocked WKD/fetch), org config parsing and merging (`org-config.js`, mocked fetch), the in-memory session cache/timeout (`session-cache.js`, fake timers), and `wkd.js`'s `lookup()` (advanced/direct URL fallback, hashing, error handling — mocked fetch, real WebCrypto). Run it with:

```bash
npm ci
npm test              # vitest run
npm run test:watch
npm run test:coverage # adds a v8 coverage report (text + html + lcov in coverage/)
```

Coverage is scoped to `web/js/pgp/**` + `web/js/wkd.js` (see `vitest.config.js`) — it's currently ~97% statements/~97% functions for those files, reported for visibility only, with no enforced threshold (see gaps below for why a hard gate would be misleading right now).

**Known gaps** (not covered — see the "Scope decision" reasoning in the PR that added the suite): the four Office.js UI entry points (`MessageCompose.js`, `MessageRead.js`, `KeyManagement.js`, `Functions/FunctionFile.js`) plus the `DecryptedPopup.js` dialog page — these execute `Office.onReady()` against real DOM element IDs and need a DOM environment plus a fuller Office.js mock to test properly. Narrow exceptions, all using plain object stubs for `Office`/`document`/`window` rather than a real DOM, since these are pure decision/rendering logic that don't touch the rest of either file's `Office.onReady`-gated setup: `tests/message-read-popout.test.js` and `tests/decrypted-popup.test.js` unit-test `handleDialogOpenFailure` (exported from `MessageRead.js`) and `renderPayload` (exported from `DecryptedPopup.js`) directly; `tests/message-read-reply-body.test.js` unit-tests `buildQuotedReplyHtml` (exported from `MessageRead.js`) with a hand-rolled `DOMParser` stub; `tests/message-read-native-reply-handoff.test.js` unit-tests `openNativeReplyWithHandoff` (also exported) with plain-object `Office`/`document` stubs. `tests/message-compose.test.js` and `tests/reply-handoff-runtime-core.test.js` are the exceptions to "no real DOM": between them they unit-test `stripPgpArmorBlock`/`pickSpliceMarker`/`armReplyHandoffListener` (all in `web/js/pgp/reply-handoff-runtime-core.js`, re-exported from `MessageCompose.js` for backward compatibility — see "Reply / Reply All quoting" below), which parse/mutate/re-serialize HTML via `document.createElement('div').innerHTML`, and the reply-handoff `BroadcastChannel` flow end-to-end — both run under a real DOM (`// @vitest-environment jsdom`, `jsdom` devDependency, per-file override of this repo's default `environment: 'node'`) rather than a hand-rolled stub, since a hand-rolled fake DOM risks subtly diverging from real HTML parsing/serialization semantics for exactly the function whose correctness matters most here. `tests/reply-handoff-channel.test.js` covers `getReplyHandoffChannelName`/`HANDOFF_PENDING_MARKER` directly (plain Node, no DOM needed). Also not covered: the specific SHA-1-self-signature legacy-key retry path in `pgp-core.js` (`_isLegacySelfSigError`/`_buildLegacyKeyReadConfig`'s hash-rejection branch — the fixture key uses a modern SHA-256 self-signature since reproducing a genuine SHA-1 one needs forging packets or an ancient GnuPG version; the ElGamal/DSA "weak key" rejection path is covered); and the three event-based-activation entry points added for #22 — `web/ReplyHandoffRuntime.html`/`.js`, `web/Functions/ReplyHandoffRuntime.classic.js`, and `web/ReplyHandoffPane.js` — none of which have any automated test coverage at all. All three are thin Office.js glue around the already-tested `armReplyHandoffListener()` core (or, for the classic file, plain string/callback logic with no shared module to test against, per its own no-ES-modules constraint); their actual behavior (does the event fire, does the notification post and gate correctly, does the pane open/splice/close) can only be verified by live sideload testing, which is how #22 was actually validated — see that PR's description for what was manually confirmed on classic Outlook Desktop specifically. CI (`.github/workflows/deploy-pages.yml`) runs `npm run test:coverage` in the `build` job before packaging/deploying, uploading the coverage report as a workflow artifact — a failing test blocks the GitHub Pages deploy; coverage itself is not a gate.

## Regenerating icons

```bash
pip install Pillow
python generate_icons.py
```

Outputs PNG files to `web/images/` at all required sizes. There are four icon families: `Icon*` (group button), `IconEncrypt*`, `IconDecrypt*`, `IconKeys*`. Sizes generated: 16, 32, 64, 80, 128, 192 px. The 128 px variant (`Icon128.png`) is required by AppSource for `HighResolutionIconUrl` in mail add-ins.

## Architecture

The add-in has four entry points in `web/`:

| File | Purpose | Office API requirement |
|------|---------|----------------------|
| `MessageCompose.html/.js` | Encrypt/decrypt outgoing messages, manage recipient keys | Mailbox 1.5 min; attachment encryption gates on `_has18` (1.8) at runtime |
| `MessageRead.html/.js` | Decrypt incoming messages, verify signatures | Mailbox 1.5 min; attachment decrypt gates on `_has18` (1.8); sender info gates on `_has17` (1.7) |
| `KeyManagement.html/.js` | Key generation, import, export, contacts keyring, org settings, Ko-fi support button (dynamically injected; suppressed by `hideSupportButton` org config) | Mailbox 1.1 |
| `Functions/FunctionFile.html/.js` | UI-less ribbon action host | Mailbox 1.1 |

`web/DecryptedPopup.html/.js` is a fifth page, but not a ribbon/task-pane entry point — it's a dialog spawned by `MessageRead.js`'s "Pop Out" button via `Office.context.ui.displayDialogAsync` (Mailbox 1.4+), receiving the decrypted payload over a same-origin `BroadcastChannel` rather than Office's `Dialog.messageChild` (which needs Mailbox 1.9). Below 1.4, if `BroadcastChannel` is unavailable, or if the dialog opens but its handshake never completes / it reports its own failure / it closes unexpectedly, `MessageRead.js` falls back to the legacy `window.open()`-based `openDecryptedPopup()` — see `triggerPopoutFallback()`, the single chokepoint all of those failure signals funnel through. On mobile the "Pop Out" button is hidden entirely (mobile WebViews block `window.open()`), so neither path is reachable there.

Three more non-ribbon pages exist for the reply-handoff feature (see "Reply / Reply All quoting" → "Event-based activation" below for the full picture):
- `web/ReplyHandoffRuntime.html`/`.js` — the browser-runtime `OnNewMessageCompose` event handler (web/new Outlook Windows/new Mac UI). Currently a no-op stub; real logic deferred.
- `web/Functions/ReplyHandoffRuntime.classic.js` — the classic-Windows `OnNewMessageCompose` event handler, declared via the manifest's `<Runtimes>`/`<Override>` elements rather than a ribbon command.
- `web/ReplyHandoffPane.html`/`.js` — a dedicated minimal task pane, opened by that classic handler's notification action (also has its own `msgComposeReplyHandoffButton` ribbon control as a manual fallback).

All four main entry points import from the shared modules in `web/js/pgp/`. The strict dependency order (no reverse imports):

```
pgp-core.js                   ← sole importer of openpgp.min.mjs
key-storage.js                ← sole caller of Office.context.roamingSettings
keyring.js                    ← calls key-storage + pgp-core
key-discovery.js              ← calls keyring + pgp-core + wkd.js
org-config.js                 ← calls key-storage + key-discovery
session-cache.js              ← standalone (in-memory only, no imports from pgp/)
quoted-content.js             ← standalone (no imports) — shared by MessageRead.js and MessageCompose.js, which otherwise have no dependency on each other
reply-handoff-channel.js      ← standalone (no imports) — BroadcastChannel name derivation + HANDOFF_PENDING_MARKER
reply-handoff-runtime-core.js ← calls quoted-content.js + reply-handoff-channel.js — shared reply-handoff listener/splice logic (MessageCompose.js's own listener + web/ReplyHandoffPane.js)
```

`pgp-core.js` is the only file that touches the OpenPGP.js library. All crypto goes through it.

Message and attachment encryption also enable DEFLATE compression (`config.preferredCompressionAlgorithm`) to shrink the armored payload. This is a soft opt-in: OpenPGP.js only compresses when every recipient key's self-signature advertises support for it, silently falling back to uncompressed otherwise — so it's safe for legacy/stripped keys. Decompression is fully automatic on the decrypt side (the packet carries its own algorithm ID); no decrypt-side code or config is needed.

## Storage model

Everything persists in **Office roaming settings** (32 KB total cap, syncs across devices):

| Key | Content |
|-----|---------|
| `pgp_private_key` | Armored, passphrase-encrypted private key |
| `pgp_public_key` | Armored public key |
| `pgp_key_meta` | `{ name, email, fingerprint, keyId, created, expires, algorithm }` |
| `pgp_keyring` | `{ "email": "armored public key", … }` — contacts' keys |
| `pgp_org_override` | Manual org config override |
| `pgp_sign_default` | Boolean — user's default for the sign-messages toggle |

Storage budget is tight: ~8–10 ECC contact keys fit comfortably. Call `estimateStorageUsage()` to warn users before hitting the limit. RSA-4096 keys are ~2–3× larger than ECC keys.

## Recipient resolution (MessageCompose.js)

`item.to.getAsync()` / `item.cc.getAsync()` only return recipients Outlook has *finished* resolving — a recipient that's still being resolved (e.g. right after "Encrypted Reply" pre-populates To/Cc, or moments after the user finishes typing) is silently omitted rather than returned in some partial state. There is no public Office.js API to force that resolution (no `checkNames()`/`resolveRecipients()` equivalent), so `getRecipientsAsync()` polls up to 5 times at 300ms intervals until two consecutive reads agree on the recipient count. `handleEncrypt()` calls `loadRecipients()` (which uses this polling read and re-runs key discovery for any newly-resolved recipient) as its first step, before unlocking the signing key or touching attachments, and aborts with a status message if any recipient still lacks a resolved key afterward. This prevents encrypting against a stale/incomplete recipient list when the user clicks Encrypt quickly.

## Session cache (`session-cache.js`)

The unlocked private key is held **only in the JavaScript heap** — never written to any persistent storage. Key facts:
- Each task pane (Compose / Read / KeyManagement) is a separate WebView with its own module scope; the cache is per-pane by design.
- Default timeout: 15 minutes of inactivity (no `getSessionKey()` calls). Every call resets the timer.
- The passphrase itself is never retained — only the derived in-memory key object is cached.
- `clearSessionKey()` is the programmatic lock.

## Manifest

`manifest/manifest.xml` is an XML-format Office add-in manifest (VersionOverrides 1.0). It targets `MailApp` type with a 2-tier Mailbox requirement: a legacy `<Requirements>` block pinned to 1.1 for the add-in to load at all, and a `DefaultMinVersion="1.5"` inside `VersionOverrides` for the ribbon/task-pane surface. Attachment encryption (1.8) and sender-info APIs (1.7) are feature-detected at runtime via `_has18`/`_has17` rather than gated in the manifest — see the entry-points table above. It also declares Office.js event-based activation (`<Runtimes>` + `<ExtensionPoint xsi:type="LaunchEvent">`, `OnNewMessageCompose`) and an extra `msgComposeReplyHandoffButton` ribbon control for the reply-handoff feature — see "Reply / Reply All quoting" → "Event-based activation" below for the full rationale. `DefaultMinVersion` deliberately stays at 1.5 rather than being bumped to 1.10 for this.

The manifest in the repo points to `https://pgp-outlook.ruscaltech.com`. When forking or self-hosting, replace every URL in the file and regenerate the `<Id>` GUID. The `<AppDomains>` section controls task-pane navigation only, **not** `fetch()`/XHR (which is governed by CORS on the target server).

The root `<Version>` element must strictly increase on any change that will be re-imported through admin-managed deployment (Microsoft 365 admin center → Integrated Apps) — it rejects a re-upload whose version isn't higher than what's currently installed, even for manifest changes that don't touch functionality. `manifest/manifest.json` (the unified/JSON manifest used by the separate tag-triggered `v*` release, see below) is kept in sync with the same version number by convention, though its own schema/content is otherwise independent.

### Rolling "latest" release

`.github/workflows/deploy-pages.yml`'s `build` job publishes/updates a GitHub Release tagged `latest` containing only `manifest/manifest.xml`, whenever a push to `main`/`master` both changes that file (checked via `dorny/paths-filter`, comparing against the prior commit) and passes the test suite. This gives IT admins a stable import-by-URL target for Exchange Online / M365 admin center's Integrated Apps (`https://github.com/<org>/<repo>/releases/latest/download/manifest.xml`) that updates itself — distinct from the separate, tag-triggered (`v*`) release which bundles `manifest.json` + icons instead.

## Key discovery chain

`key-discovery.js` resolves a recipient email to a public key in this order:
1. Local keyring (`key-storage.js`)
2. WKD (Web Key Directory) — authoritative for the recipient's domain
3. VKS (keys.openpgp.org) — email-verified keys

Automatically discovered keys are always shown with their source before the user can save them. The company key (org config) is fetched via the same WKD→VKS chain.

## Organization config

IT admins publish a JSON file at:
```
https://<email-domain>/.well-known/pgp-for-outlook-addin/company-config.json
```
(fallback: `https://openpgpkey.<email-domain>/...`). See `docs/company-config.example.json` for the schema. The add-in fetches it anonymously and derives the URL from the signed-in user's email domain. A manual override stored in roaming settings takes precedence.

Key config fields: `companyKeyEnabled`, `companyKeyRequired`, `companyKeyEmails[]`, `hideSupportButton` (hides the Ko-fi donation button in the Key Management pane; when `true` the external CDN script is never loaded), and `companyDecryptedExtensionPrefix` (string, default `""`; when set, inserted before the extension of every decrypted attachment filename in `MessageRead.js`, e.g. `"pgpDecrypted"` turns `report.xlsx` into `report.pgpDecrypted.xlsx`; empty/`null`/absent leaves filenames unchanged).

## Encryption scope

The add-in encrypts the **HTML body** of the message and replaces it with PGP armor. Subject lines, sender/recipient headers, and metadata are not encrypted (fundamental OpenPGP-over-email limitation). On decrypt, the original HTML is recovered and rendered in a sandboxed iframe.

Attachments are encrypted individually to `filename.ext.pgp`. Inline (clipboard-pasted) images in the body cannot be read by the Office API on Outlook desktop; the add-in detects the broken `cid:` reference and warns the user. On Outlook Web it can convert them to regular attachments automatically. On decrypt, `MessageRead.js` offers both per-attachment "Decrypt & Download" buttons and a "Save All" button that decrypts and downloads every PGP attachment on the message sequentially, prompting for the passphrase only once.

### Decrypting from Compose

Encrypting to your own public key is always done unconditionally alongside recipient/company keys (so the sender can read their own sent mail) — this also means the sender's own private key is always sufficient to reverse an encrypt performed from this add-in's own compose pane, regardless of which recipients were chosen. A "Decrypt" button (`refreshComposeButtons()`/`handleDecrypt()` in `MessageCompose.js`) appears whenever the compose body is currently PGP-armored, letting the user restore the original body and revert any `.pgp` attachments back to their originals so recipients, body, or attachments can be edited before re-encrypting (#25). Attachment reversal is best-effort per file: a `.pgp` attachment that fails to decrypt (e.g. hand-corrupted armor) is left alone and named in a warning status, without blocking the body restore or any other attachment.

### Reply / Reply All quoting

**Normal-sized messages** (the common case): `MessageRead.js`'s reply buttons call `Office.context.mailbox.displayNewMessageFormAsync`/`displayNewMessageForm` with a quoted copy of the decrypted body as `formData.htmlBody`. Office.js caps `htmlBody` at 32 KB (32,768 characters) — confirmed to be the exact same cap on `Office.ReplyFormData.htmlBody` (used by `displayReplyForm`/`displayReplyAllForm`), so switching APIs alone would not raise it — and Outlook Classic enforces this synchronously, throwing `Sys.ArgumentOutOfRangeException` if exceeded. `buildQuotedReplyHtml()` (exported from `MessageRead.js`) keeps the quote under that limit: it returns `{ html, truncated }` — `html` unchanged and `truncated: false` when it fits; if the formatted HTML would exceed the limit it falls back to a plain-text quote (never truncates raw HTML, to avoid emitting unbalanced tags) and `truncated` is already `true` at that point — falling back to plain text is itself a real degradation (all formatting lost), even if the resulting plain text turns out short enough that no further cutting is needed; if even that is too large, it additionally truncates the text and appends a visible "[Original message truncated...]" notice. Callers branch on the `truncated` flag itself, not by searching `html` for the notice text — the decrypted message can legitimately *contain* that exact string (e.g. quoting an earlier reply that itself got truncated), which would make a substring check false-positive on an otherwise normal-sized message. Content formatting itself (HTML-vs-plain-text rendering, `<head><style>` preservation — see below) is delegated to the shared `web/js/pgp/quoted-content.js` module (`formatDecryptedContentAsHtml`/`formatDecryptedContentAsPlainTextHtml`).

Office also rejects a nested `<html>` tag inside `htmlBody`, so the HTML quote uses only `doc.body.innerHTML` from the decrypted document — but it explicitly carries forward any `<style>` block(s) from `<head>` too. Outlook Desktop's Word-based HTML export commonly relies on such a rule (e.g. `p.MsoNormal { margin:0 }`) to render single-spaced lines correctly; the decrypt preview iframe and pop-out window render the full document (`<head>` intact) so they're unaffected, but a bare `doc.body.innerHTML` would silently drop it, letting default browser paragraph margins reappear as extra blank lines in the reply compose window only.

**Large messages** (when `buildQuotedReplyHtml()` returns `truncated: true`): `handleReplyEncrypted()` instead calls `Office.context.mailbox.item.displayReplyForm({ htmlBody: HANDOFF_PENDING_MARKER })`/`displayReplyAllForm({ htmlBody: HANDOFF_PENDING_MARKER })` — `formData` is a required parameter of both APIs despite commonly being called with none, so *some* value is required to avoid a synchronous throw. The `ReplyFormData` **object** form is required here, not a bare string: confirmed live that passing the marker as a plain string inserts no visible text into the reply body on classic Outlook Desktop, despite Microsoft's own docs example showing exactly that usage (no documented or community-confirmed explanation found). `HANDOFF_PENDING_MARKER` (`web/js/pgp/reply-handoff-channel.js`) is used instead of an empty string specifically so the classic-Windows event handler (see below) can tell, from the body alone, that this particular reply was opened for a handoff — Outlook prepends it above its own native quote, letting Outlook build its native reply — proper `In-Reply-To`/`References` threading, recipients (self excluded from Reply All automatically), and "Re:" subject, all handled by Outlook itself. Outlook's native quote still shows the original message *as Outlook has it: PGP-armored*, since Outlook has no notion of decryption. `MessageRead.js` then hands the decrypted plaintext (captured into locals before starting, so a later reading-pane navigation/decrypt can't leak a different message's plaintext into an in-flight handoff) to the new compose window over a `BroadcastChannel`, repeating the broadcast every ~400 ms (since a message posted before the compose window's listener is ready is silently dropped, with no queueing for late subscribers) for up to `REPLY_HANDOFF_TIMEOUT_MS` (20 s — doubled from the original 10 s for extra breathing room on a resource-constrained system/connection, #22).

The channel name is derived from the message's `conversationId` (`getReplyHandoffChannelName()`, `web/js/pgp/reply-handoff-channel.js`, standalone/no imports so both sides can never derive it differently) rather than fixed — there is no way to control the new window's URL the way the pop-out dialog does, so there's nowhere to embed a per-instance token for it to read on load, and a single hardcoded name would let *any* same-origin script (BroadcastChannel is same-origin-wide, not scoped to just these two windows) passively collect every large-message reply's plaintext across the whole product. It uses the conversation ID directly (URI-encoded), not a hash of it — a fixed-width hash has a real collision probability at scale, which would make two unrelated conversations share a channel name and cross-talk (one conversation's plaintext spliced into, or leaked to, another's reply); the conversation ID itself isn't secret (anyone with message access already has it), so encoding it directly costs nothing and can't collide. If `conversationId` is missing, both sides fall back to a second scoping ID before giving up: `MessageRead.js` tries `item.internetMessageId` (available since Mailbox 1.1 — broader coverage than `conversationId` itself), and `MessageCompose.js` tries `item.inReplyTo` (Mailbox 1.14, gated behind a `_has114` check — "the internet message ID of the original message being replied to", i.e. the same value the read side computed). **Only if *neither* ID is available do both sides independently refuse the handoff entirely** rather than falling back to the shared base channel name (which would reintroduce the broad-exposure problem this whole scheme exists to avoid): `openNativeReplyWithHandoff()` (exported for testing) skips straight to the existing `displayNewMessageForm` path without even opening a native reply, and `setupReplyHandoffListener()` (also exported, with both `has110` and `has114` explicit parameters for testing) never subscribes at all. None of this is a secrecy boundary on its own, just a floor against the broadest, most trivial form of that exposure — see `MessageCompose.js`'s other mitigations below.

The listener/splice logic itself (`armReplyHandoffListener()`, `applyReplyHandoff()`, `stripPgpArmorBlock()`, `pickSpliceMarker()`) lives in `web/js/pgp/reply-handoff-runtime-core.js`, shared between `MessageCompose.js`'s own listener and `web/ReplyHandoffPane.js` (see "Event-based activation" below) so both provably run the same code. `armReplyHandoffListener()` only listens when the compose window is confirmed to be a reply (`getComposeTypeAsync()`, Mailbox 1.10+ — Office's `ComposeType` enum has only `Reply`/`NewMail`/`Forward`, no separate value for Reply All, so a plain reply and a reply-all report the same `composeType`; an ordinary new message or forward never listens at all; falls back to listening unconditionally on older hosts that can't be asked), and only for a bounded time (~24 s, kept proportionally ahead of `REPLY_HANDOFF_TIMEOUT_MS`) — a reply window that was never the target of an actual handoff (e.g. the user used Outlook's own Reply button on an encrypted-but-undecrypted message) stops listening after a short grace period rather than for the rest of its lifetime. `MessageCompose.js`'s exported `setupReplyHandoffListener(has110, has114)` is now a thin wrapper (`armReplyHandoffListener({ has110, has114, onStatus: showStatus })`) kept for its own pane's use and for existing tests — both params stay explicit (default: the real feature-detected flags) so every branch is directly testable. If a splice attempt fails, the listener stays open (rather than closing outright) so a later re-broadcast from `MessageRead.js`'s retry loop gets another chance to succeed before the timeout — only a *successful* splice closes the channel and stops listening. `applyReplyHandoff()` itself shows no status — it returns `{ success, message }` and the caller decides when to surface `message` (via `onStatus`), showing a failure warning only once rather than on every retry (a splice that fails once typically fails identically on every subsequent ~400ms re-broadcast for the full ~20s window, so re-showing it each time would flicker the status bar). An optional `onSettled(result)` callback fires once, on every settle path (ack, timeout, or an immediate skip like "not a reply"/"no scoping ID") — `MessageCompose.js` doesn't pass one (preserving its exact pre-extraction resolve-immediately-after-arming timing, since existing tests depend on it), but `web/ReplyHandoffPane.js` does, to decide when to close itself.

`stripPgpArmorBlock()` locates and removes the armor block by replacing it with an internal splitting marker, serializing the DOM once, then splitting the resulting HTML string on that marker (see its docblock for why, over reconstructing partial DOM structure by hand). Since the input is attacker-influenceable PGP message content, the marker isn't assumed unique on its own: `pickSpliceMarker()` checks the raw input first and falls back to a suffixed variant if the base marker already appears there, and the split result is validated to be exactly 2 parts before use — anything else (marker never landed in the output, or more copies of it turned up than expected) is treated the same as "armor not found" rather than risking a corrupted splice.

On receipt of a handoff, it reads its own current body via `item.body.getAsync` (already-open compose items write via `Body.setAsync`, which is capped at 1 MB — not the 32 KB `htmlBody` cap — since this is the one Office.js body-write path not bound by the constraint the whole large-message path exists to avoid), locates and removes the PGP armor block via the exported `stripPgpArmorBlock()` (walks a detached DOM the same way `MessageRead.js`'s `extractArmorFromHtml()` does — same `<pre>`/`<br>`/block-element handling — but tracks text-node offsets so the located range can be removed rather than only extracted; robust to the armor being split across sibling nodes, or wrapped in a `<pre>` the way this add-in's own `setBodyAsync()` sends it), and splices the formatted decrypted content in at that location. **Only once that splice is confirmed to have written successfully** does it ack — never before attempting it — because an ack sent early would make `MessageRead.js` treat the handoff as done and never trigger its own fallback below, leaving the user with a still-armored body and no backup window. If the armor block can't be found, or `getAsync`/`setAsync` throws, the body is left untouched, a warning status is shown, and (critically) no ack is sent.

If no ack arrives within the timeout — whether because the compose side never got the broadcast, isn't a reply, or the splice itself failed — `MessageRead.js` falls back to the **normal-sized-message path** above (`displayNewMessageForm` + `buildQuotedReplyHtml`, truncation and all). `openNativeReplyWithHandoff()` falls back to this path for four distinct reasons (`HandoffFallbackReason`, all handled by `openReplyComposeForm()`'s `handoffFallbackReason` parameter), and only two of them actually leave a second window open: a missing scoping ID and a `displayReplyForm`/`displayReplyAllForm` throw both happen *before* any native reply is opened, so the fallback becomes the only window and its warning says so; a `BroadcastChannel` construction failure and a handoff timeout both happen *after* the native reply already succeeded, so the fallback opens a genuine *second* window and its warning tells the user to close the other (still-blank/still-armored) one.

While a native-reply handoff from a given reading pane is in flight (between `displayReplyForm`/`displayReplyAllForm` succeeding and the handoff settling), that pane's Reply/Reply All buttons are disabled and a second click is refused with a status message (`_nativeReplyHandoffInFlight` in `MessageRead.js`) — this narrows, but does not eliminate, the case where two concurrent large-message replies on the same conversation cross-wire the shared conversation-scoped `BroadcastChannel`: it only guards against two attempts from the *same* pane, not two separate Outlook windows racing each other on the same thread.

#### Event-based activation (#22)

Everything above still depends on `MessageCompose.js`'s listener actually running — which, historically, only happened if the user manually opened this add-in's task pane inside the new reply window (Outlook never loads a task pane's JS just because a compose window exists). `manifest/manifest.xml` now also declares Office.js **event-based activation**: a `<Runtimes>` element plus an `<ExtensionPoint xsi:type="LaunchEvent">` registering `OnNewMessageCompose` (Mailbox 1.10), which fires automatically the instant *any* new compose window opens (New Mail, Reply, Reply All, or Forward) — no task pane required. `DefaultMinVersion` stays at 1.5 in that `VersionOverridesV1_1` node (bumping it would gate the whole node, including `MobileFormFactor`, behind 1.10) — hosts below 1.10 simply never fire the event, same graceful-degradation pattern as `_has110`/`_has114` elsewhere.

Two different runtime files back this, declared via the same `<Runtimes>` element:
- `web/ReplyHandoffRuntime.html`/`.js` — used by Outlook on the web, new Outlook on Windows, and new Mac UI (full browser runtime: modern JS, ES modules, `BroadcastChannel` all fine). **Not yet implemented beyond a no-op** — confirmed live that the event does not fire there via sideload even with the manifest genuinely active (ruled out a stale-sideload false negative); root cause unconfirmed, deferred.
- `web/Functions/ReplyHandoffRuntime.classic.js` — used by **classic** Outlook on Windows. Runs in Outlook's stripped-down "JavaScript-only runtime", confirmed live to have **neither `document` nor `BroadcastChannel`** — so unlike the browser-runtime file, it cannot receive or splice the decrypted payload itself. Must also avoid `async`/`await` and the ternary operator (both break the add-in on classic builds before ~April 2024 / Build 17425.20000) and `import`/ES modules (single flat file, nothing bundled) — see its own comments.

Every alternative for getting the payload into that restricted runtime was considered and ruled out: `Office.context.mailbox.item.sessionData` doesn't bridge across items (the original message and the new reply are two different items with separate `sessionData` stores, and it's compose-mode-only besides) and is capped at 50,000 characters on hosts ≤ Mailbox 1.15 anyway; `Office.addin.showAsTaskpane()`/`hide()` require `SharedRuntime`, which Outlook does not support at all (confirmed: "Shared runtimes aren't supported in Outlook" — each of task pane / function command / event-based task / dialog gets its own runtime); raw temp-file I/O isn't exposed to JS in either runtime (no `fs`-equivalent; the browser File System Access API needs a user gesture per call, so nothing can write or read one silently in the background).

The actual design: `ReplyHandoffRuntime.classic.js` posts an `Office.MailboxEnums.ItemNotificationMessageType.InsightMessage` notification (Mailbox 1.10) whose one supported action (`actionType: 'showTaskPane'` — passed as a plain string, not `Office.MailboxEnums.ActionType.ShowTaskPane`, since that enum may not be populated in this restricted runtime and an undefined-property access inside an async callback throws silently, uncatchable by the surrounding try/catch which has already returned by the time the callback runs) opens `web/ReplyHandoffPane.html`/`.js` — a **dedicated minimal pane** (its own `msgComposeReplyHandoffButton` ribbon control, also manually clickable as a fallback), deliberately separate from the full Encrypt/Decrypt pane so the user isn't shown unrelated recipient/sign UI for what should be a one-click action. Because opening a task pane this way is a real user gesture (clicking the notification's action), not a programmatic pane-open, it's unaffected by the `SharedRuntime` restriction above. That pane has a full browser runtime (a real `BroadcastChannel`), arms `armReplyHandoffListener()` exactly like `MessageCompose.js`'s own pane, and — once the handoff settles successfully — closes itself via `Office.context.ui.closeContainer()` (plain Mailbox 1.5, not `SharedRuntime`-gated; confirmed to be the same mechanism third-party add-ins like KnowBe4's Phish Alert Button use for an equivalent "submit, then auto-close" flow). It stays open on failure so the user can read why, rather than auto-closing and hiding the one signal that something needs attention.

The classic-Windows handler also gates its notification so it only appears on a reply this add-in actually opened for a handoff — not on every reply to every encrypted message. `HANDOFF_PENDING_MARKER` (`web/js/pgp/reply-handoff-channel.js`) is passed as the `htmlBody` of the `ReplyFormData` object given to `openNativeReplyWithHandoff()`'s `displayReplyForm`/`displayReplyAllForm` call instead of an empty string; the classic handler checks the reply body for that literal text (via `body.getAsync`) before posting anything, and `applyReplyHandoff()` strips it as part of the same splice that replaces the PGP armor. The marker deliberately contains no quotes or HTML-reserved characters, so a plain string search/removal is robust regardless of how either runtime (de)serializes it. `ReplyHandoffRuntime.classic.js` can't import `reply-handoff-channel.js` (no ES modules in its runtime), so it duplicates the literal string — `tests/reply-handoff-channel.test.js` asserts the two stay in sync.

This currently only fully closes the loop on **classic Outlook Desktop** — the browser-runtime file's real listener logic (for web/new Outlook/Mac) is deferred pending the `OnNewMessageCompose`-not-firing-there investigation above. `MessageCompose.js`'s own listener (armed only when the user manually opens the pane) remains the fallback for every platform: hosts below Mailbox 1.10, tenants that haven't yet re-approved the updated manifest (event-based add-ins require fresh admin consent on every manifest change, per Microsoft's docs — the whole add-in is blocked until re-approved, not just the new feature), or simply a user who gets there first on their own.

## Skill routing

When the user's request matches an available skill, invoke it via the Skill tool. The
skill has multi-step workflows, checklists, and quality gates that produce better
results than an ad-hoc answer. When in doubt, invoke the skill. A false positive is
cheaper than a false negative.

Key routing rules:
- Product ideas, "is this worth building", brainstorming → invoke /office-hours
- Strategy, scope, "think bigger", "what should we build" → invoke /plan-ceo-review
- Architecture, "does this design make sense" → invoke /plan-eng-review
- Design system, brand, "how should this look" → invoke /design-consultation
- Design review of a plan → invoke /plan-design-review
- Developer experience of a plan → invoke /plan-devex-review
- "Review everything", full review pipeline → invoke /autoplan
- Bugs, errors, "why is this broken", "wtf", "this doesn't work" → invoke /investigate
- Test the site, find bugs, "does this work" → invoke /qa (or /qa-only for report only)
- Code review, check the diff, "look at my changes" → invoke /review
- Visual polish, design audit, "this looks off" → invoke /design-review
- Developer experience audit, try onboarding → invoke /devex-review
- Ship, deploy, create a PR, "send it" → invoke /ship
- Merge + deploy + verify → invoke /land-and-deploy
- Configure deployment → invoke /setup-deploy
- Post-deploy monitoring → invoke /canary
- Update docs after shipping → invoke /document-release
- Weekly retro, "how'd we do" → invoke /retro
- Second opinion, codex review → invoke /codex
- Safety mode, careful mode, lock it down → invoke /careful or /guard
- Restrict edits to a directory → invoke /freeze or /unfreeze
- Upgrade gstack → invoke /gstack-upgrade
- Save progress, "save my work" → invoke /context-save
- Resume, restore, "where was I" → invoke /context-restore
- Security audit, OWASP, "is this secure" → invoke /cso
- Make a PDF, document, publication → invoke /make-pdf
- Launch real browser for QA → invoke /open-gstack-browser
- Import cookies for authenticated testing → invoke /setup-browser-cookies
- Performance regression, page speed, benchmarks → invoke /benchmark
- Review what gstack has learned → invoke /learn
- Tune question sensitivity → invoke /plan-tune
- Code quality dashboard → invoke /health
