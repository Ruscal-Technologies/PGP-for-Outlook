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

There is no test suite yet (tracked as a roadmap item — Vitest/Jest for `pgp-core.js`, Office mock libraries for integration tests).

## Regenerating icons

```bash
pip install Pillow
python generate_icons.py
```

Outputs PNG files to `web/images/` at all required sizes. There are three icon families: `Icon*` (group button), `IconEncrypt*`, `IconDecrypt*`, `IconKeys*`.

## Architecture

The add-in has four entry points in `web/`:

| File | Purpose | Office API requirement |
|------|---------|----------------------|
| `MessageCompose.html/.js` | Encrypt outgoing messages, manage recipient keys | Mailbox 1.8 (attachment APIs) |
| `MessageRead.html/.js` | Decrypt incoming messages, verify signatures | Mailbox 1.3 |
| `KeyManagement.html/.js` | Key generation, import, export, contacts keyring, org settings | Mailbox 1.1 |
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

## Session cache (`session-cache.js`)

The unlocked private key is held **only in the JavaScript heap** — never written to any persistent storage. Key facts:
- Each task pane (Compose / Read / KeyManagement) is a separate WebView with its own module scope; the cache is per-pane by design.
- Default timeout: 15 minutes of inactivity (no `getSessionKey()` calls). Every call resets the timer.
- The passphrase itself is never retained — only the derived in-memory key object is cached.
- `clearSessionKey()` is the programmatic lock.

## Manifest

`manifest/manifest.xml` is an XML-format Office add-in manifest (VersionOverrides 1.0). It targets `MailApp` type with Mailbox 1.8 requirement.

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

## Encryption scope

The add-in encrypts the **HTML body** of the message and replaces it with PGP armor. Subject lines, sender/recipient headers, and metadata are not encrypted (fundamental OpenPGP-over-email limitation). On decrypt, the original HTML is recovered and rendered in a sandboxed iframe.

Attachments are encrypted individually to `filename.ext.pgp`. Inline (clipboard-pasted) images in the body cannot be read by the Office API on Outlook desktop; the add-in detects the broken `cid:` reference and warns the user. On Outlook Web it can convert them to regular attachments automatically.
