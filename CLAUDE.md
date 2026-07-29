# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project overview

A Microsoft Outlook add-in (MailApp) that provides end-to-end PGP encryption in Outlook web, desktop, and mobile. It is a **zero-build-step** static web app — plain HTML/CSS/ES modules served over HTTPS. There is no npm, no bundler, no transpilation step.

## Local development

```bash
# One-time: install a dev HTTPS server and trust its self-signed cert
npm install -g office-addin-dev-certs http-server
office-addin-dev-certs install

# Serve the add-in
http-server web --ssl --port 3000
```

Update `manifest/manifest.xml` to point at `https://localhost:3000/`, then sideload it in Outlook following Microsoft's [sideloading guide](https://docs.microsoft.com/en-us/office/dev/add-ins/testing/test-debug-office-add-ins).

### Tests

A Vitest suite in `tests/` covers the shared business-logic modules in `web/js/pgp/*.js` plus `web/js/wkd.js` — crypto round trips including multi-recipient, empty-string, unicode, and RSA-4096 (`pgp-core.js`), tampering/wrong-key failure modes (corrupted ciphertext, decrypting with a non-recipient key, malformed armor), legacy DSA+ElGamal key support (`hasWeakEncryptionKey`, `hasModernSubkeys`, `addModernSubkeys`, `extractPublicKey`) against a real fixture key in `tests/fixtures/`, storage (`key-storage.js`, mocked `Office.context.roamingSettings`, including its `saveAsync` failure path), the contact keyring (`keyring.js`), key discovery precedence and fallbacks (`key-discovery.js`, mocked WKD/fetch), org config parsing and merging (`org-config.js`, mocked fetch), the in-memory session cache/timeout (`session-cache.js`, fake timers), and `wkd.js`'s `lookup()` (advanced/direct URL fallback, hashing, error handling — mocked fetch, real WebCrypto). Run it with:

```bash
npm ci
npm test              # vitest run
npm run test:watch
npm run test:coverage # adds a v8 coverage report (text + html + lcov in coverage/)
```

Coverage is scoped to `web/js/pgp/**` + `web/js/wkd.js` (see `vitest.config.js`) — it's currently ~97% statements/~97% functions for those files, reported for visibility only, with no enforced threshold (see gaps below for why a hard gate would be misleading right now).

**Known gaps** (not covered — see the "Scope decision" reasoning in the PR that added the suite): the four Office.js UI entry points (`MessageCompose.js`, `MessageRead.js`, `KeyManagement.js`, `Functions/FunctionFile.js`) — these execute `Office.onReady()` against real DOM element IDs and need a DOM environment plus a fuller Office.js mock to test properly; and the specific SHA-1-self-signature legacy-key retry path in `pgp-core.js` (`_isLegacySelfSigError`/`_buildLegacyKeyReadConfig`'s hash-rejection branch — the fixture key uses a modern SHA-256 self-signature since reproducing a genuine SHA-1 one needs forging packets or an ancient GnuPG version; the ElGamal/DSA "weak key" rejection path is covered). CI (`.github/workflows/deploy-pages.yml`) runs `npm run test:coverage` in the `build` job before packaging/deploying, uploading the coverage report as a workflow artifact — a failing test blocks the GitHub Pages deploy; coverage itself is not a gate.

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
| `MessageCompose.html/.js` | Encrypt outgoing messages, manage recipient keys | Mailbox 1.5 min; attachment encryption gates on `_has18` (1.8) at runtime |
| `MessageRead.html/.js` | Decrypt incoming messages, verify signatures | Mailbox 1.5 min; attachment decrypt gates on `_has18` (1.8); sender info gates on `_has17` (1.7) |
| `KeyManagement.html/.js` | Key generation, import, export, contacts keyring, org settings, Ko-fi support button (dynamically injected; suppressed by `hideSupportButton` org config) | Mailbox 1.1 |
| `Functions/FunctionFile.html/.js` | UI-less ribbon action host | Mailbox 1.1 |

All four import from the shared modules in `web/js/pgp/`. The strict dependency order (no reverse imports):

```
pgp-core.js        ← sole importer of openpgp.min.mjs
key-storage.js     ← sole caller of Office.context.roamingSettings
keyring.js         ← calls key-storage + pgp-core
key-discovery.js   ← calls keyring + pgp-core + wkd.js
org-config.js      ← calls key-storage + key-discovery
session-cache.js   ← standalone (in-memory only, no imports from pgp/)
```

`pgp-core.js` is the only file that touches the OpenPGP.js library. All crypto goes through it.

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

`manifest/manifest.xml` is an XML-format Office add-in manifest (VersionOverrides 1.0). It targets `MailApp` type with a 2-tier Mailbox requirement: a legacy `<Requirements>` block pinned to 1.1 for the add-in to load at all, and a `DefaultMinVersion="1.5"` inside `VersionOverrides` for the ribbon/task-pane surface. Attachment encryption (1.8) and sender-info APIs (1.7) are feature-detected at runtime via `_has18`/`_has17` rather than gated in the manifest — see the entry-points table above.

The manifest in the repo points to `https://pgp-outlook.ruscaltech.com`. When forking or self-hosting, replace every URL in the file and regenerate the `<Id>` GUID. The `<AppDomains>` section controls task-pane navigation only, **not** `fetch()`/XHR (which is governed by CORS on the target server).

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
