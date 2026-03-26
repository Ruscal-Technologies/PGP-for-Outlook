# PGP for Outlook

A modern Microsoft Outlook add-in that brings end-to-end PGP encryption to the
Outlook ribbon — on web, desktop, and mobile — without requiring any desktop
software, plugins, or IT-managed infrastructure.

---

## Features

| Feature | Details |
|---------|---------|
| **Key pair generation** | ECC (Ed25519/X25519) or RSA-4096 keys generated in-browser; private key stored passphrase-encrypted in Office roaming settings |
| **Key import** | Import an existing PGP private key from GnuPG, Kleopatra, Thunderbird, or any OpenPGP-compatible client; passphrase is verified before saving |
| **Key export / sharing** | Copy public key to clipboard, send via email, or download a passphrase-protected private key backup |
| **Contacts' keyring** | Store, search, and remove trusted contacts' public keys |
| **Key discovery** | Automatic lookup via WKD → VKS (keys.openpgp.org) → manual paste |
| **Message encryption** | Encrypts the message body and replaces it with PGP armor before sending |
| **Message signing** | Optionally signs outgoing messages so recipients can verify authorship; off by default, configurable per-user and per-message |
| **Encrypt to self** | Your own public key is always included so you can read your Sent items |
| **Company / legal key** | Org-level key added to every encrypted message (configurable, optional or required) |
| **Attachment encryption** | Each attachment encrypted individually to `filename.ext.pgp` |
| **Inline image handling** | Detects inline (embedded) images before encryption and warns the user. On Outlook on the web, offers automatic conversion to regular file attachments. On Outlook desktop the Office API does not expose clipboard-pasted images, so the image is removed from the message body with guidance to re-attach manually. |
| **Message decryption** | Detects and decrypts PGP-encrypted message bodies; works on desktop, OWA, and Outlook mobile |
| **Attachment decryption** | One-click decrypt and download for `.pgp` attachments |
| **Signature verification** | Verifies inline signatures against the local keyring or WKD |
| **Signed-only messages** | Displays and verifies PGP cleartext-signed messages |
| **Mobile encrypted reply** | On iOS/Android (where compose add-ins are unavailable), encrypts the reply in-pane and copies the armor to the clipboard for pasting into a normal Outlook reply |
| **Session key cache** | Unlocked private key cached in memory for 15 minutes; passphrase is never stored |

---

## Requirements

| Requirement | Minimum version |
|-------------|----------------|
| Microsoft 365 / Outlook | Any current subscription (web, Windows, Mac) |
| Office JavaScript API | Mailbox **1.8** (required for compose-side attachment access) |
| Browser / WebView2 | Edge WebView2 or any modern browser (Chrome 90+, Firefox 90+, Safari 15+) |

> **Outlook 2019 and earlier (perpetual license):** The add-in will load in
> read mode, but compose-side attachment encryption requires Mailbox 1.8, which
> is only available in Microsoft 365.  Message body encryption/decryption will
> still work.

---

## Project structure

```
manifest/
├── manifest.xml                      ← Office add-in XML manifest (classic)
└── manifest.json                     ← Unified manifest (Teams / new Outlook)

web/                                  ← Web app (host on any HTTPS server)
├── MessageRead.html / .js            ← Decrypt & verify incoming messages
├── MessageCompose.html / .js         ← Encrypt outgoing messages
├── KeyManagement.html / .js          ← Key generation, keyring, org settings
├── Functions/
│   └── FunctionFile.html / .js       ← UI-less ribbon action host
├── js/
│   ├── openpgp.min.mjs               ← OpenPGP.js v5 (ES module)
│   ├── wkd.js                        ← WKD client
│   └── pgp/                          ← Shared PGP modules
│       ├── pgp-core.js               ← Cryptographic operations
│       ├── key-storage.js            ← Office roaming settings wrapper
│       ├── keyring.js                ← Contacts' key management
│       ├── key-discovery.js          ← WKD / VKS / keyring lookup
│       ├── org-config.js            ← Organization-level configuration
│       └── session-cache.js          ← In-memory unlocked-key cache
├── css/
│   └── pgp-addon.css                 ← Shared Fluent UI styles
└── images/                           ← Add-in icons

docs/
└── company-config.example.json      ← Example org config file (see below)
```

---

## Architecture

```
┌──────────────────────────────────────────────────────────────┐
│                      Outlook ribbon                          │
│  [ Encrypt ]  [ Decrypt ]  [ Manage Keys ]                   │
└────────┬───────────────┬──────────────┬──────────────────────┘
         │               │              │
    Compose          Read pane      Key Mgmt
    task pane        task pane      task pane
         │                │              │
         └───────┬────────┘              │
                 │                       │
         ┌───────▼───────────────────────▼─────────┐
         │           Shared JS modules             │
         │  pgp-core ─ key-storage ─ keyring       │
         │  key-discovery ─ org-config             │
         └───────────────────────────────────────┬─┘
                                                 │
                          ┌──────────────────────┼──────────────────────┐
                          │                      │                      │
               Office Roaming           WKD / VKS             Well-known URL
               Settings (32 KB)         key lookup             org config
               (private key,            (network)              (network)
               public key,
               keyring,
               org override)
```

The five modules in `js/pgp/` have a strict dependency direction:

```
pgp-core.js        ← Only file that imports openpgp.min.mjs directly
key-storage.js     ← Only file that calls Office.context.roamingSettings
keyring.js         ← calls key-storage + pgp-core
key-discovery.js   ← calls keyring + pgp-core + wkd.js
org-config.js      ← calls key-storage + key-discovery
```

Page scripts (`MessageRead.js` etc.) import from the modules above and from
Office.js.  No page script imports Office.js internals from another page script.

---

## Security model

### Key storage
The private key is **never stored in plaintext**.  It is stored in armored
format, encrypted with AES-256 using the user's passphrase (standard OpenPGP
S2K + symmetric cipher).  Only the encrypted blob is written to Office roaming
settings; the plaintext private key material exists in browser memory only
during the brief window of a single encrypt/decrypt/sign operation.

### Passphrase handling
The passphrase unlocks the private key once per session.  The **unlocked key
object** is held in browser memory for up to 15 minutes of inactivity, then
discarded automatically.  The passphrase itself is **never retained** after the
unlock step — only the derived in-memory key object is cached.

The session can be locked at any time with the **Lock** button in the read pane.
Navigating away from a message or closing the task pane also clears the cache
because the in-memory state is not persisted to sessionStorage or any other
durable store.

### Key discovery trust
Keys discovered automatically from WKD or VKS are shown with their source
before the user can save them.  VKS keys (keys.openpgp.org) have had their
email addresses verified by the owner.  WKD keys are authoritative for the
domain that published them.

**Important:** Even verified keys should have their fingerprints confirmed
out-of-band (phone call, Signal, in-person) before sending truly sensitive data
to a new contact.

### Scope of encryption
The add-in encrypts the **HTML content** of the message body (preserving
formatting) and replaces it with PGP armor.  When the recipient decrypts, the
original HTML is recovered and rendered in a sandboxed iframe.  Subject lines,
sender/recipient headers, and message metadata are **not** encrypted (this is a
fundamental limitation of OpenPGP applied to email; metadata encryption requires
a different transport layer entirely).

---

## Deployment

### 1. Host the web app

The `web/` folder is a static web app — no server-side code, no database.
Host it on any HTTPS-enabled server.

---

#### Option A: GitHub Pages (recommended — no infrastructure needed)

This repository includes a ready-to-use GitHub Actions workflow
(`.github/workflows/deploy-pages.yml`) that publishes the add-in automatically
on every push to `main`/`master`.

**One-time setup:**

1. Push the repository to GitHub (public or private — Pages works for both).
2. Go to your repository's **Settings → Pages**.
3. Under *Build and deployment*, select **GitHub Actions** as the source.
4. Push any commit to `main`; the workflow runs and your Pages URL appears in
   the Settings → Pages panel, e.g.:
   ```
   https://<your-org-or-username>.github.io/<repo-name>/
   ```
5. Update the manifest (see step 2 below) with that URL.

> **HTTPS is automatic** — GitHub Pages always serves over HTTPS.

---

#### Option B: Azure Static Web Apps

Drag and drop the `web/` folder into the Azure portal,
or connect it to the repository for automatic deployment.  Free tier available.

#### Option C: SharePoint

Host the folder as a SharePoint App Page.  Useful for organizations that want
the add-in files inside their Microsoft 365 tenant.

#### Option D: IIS / Apache / Nginx

Copy the folder to a virtual directory configured for HTTPS.

> HTTPS is **required** in all cases.  Office add-ins will not load over plain HTTP.

---

### 2. Update the manifest

The manifest in this repository is pre-configured to point at the GitHub Pages
deployment for this project (`https://pgp-outlook.ruscaltech.com`).
If you fork the project or host the `web/` folder elsewhere, open
`manifest/manifest.xml` and update every URL accordingly.  Do **not** include a
trailing slash:

```xml
<!-- Example: update all resource URLs for a fork -->
<bt:Url id="messageReadTaskPaneUrl"
        DefaultValue="https://your-org.github.io/your-repo/MessageRead.html"/>
```

A quick way to do all substitutions at once (Linux / macOS):

```bash
sed -i 's|https://pgp-outlook.ruscaltech.com|https://your-org.github.io/your-repo|g' \
    manifest/manifest.xml
```

Also replace the `<Id>` GUID with a freshly generated one so the manifest has
a unique identity in your tenant:

```xml
<Id>YOUR-NEW-GUID-HERE</Id>
```

Generate a GUID with `uuidgen` (Linux/macOS) or PowerShell's `[guid]::NewGuid()`.

### 3. Deploy the manifest

**Personal / testing:**
Outlook → Get Add-ins → My Add-ins → Custom add-ins → **+ Add from file** →
upload the XML.

**Enterprise (centralised deployment):**
Microsoft 365 Admin Center → Settings → Integrated Apps → Upload custom app.
Or use the `New-OrganizationAddIn` PowerShell cmdlet.

---

## Organization configuration (company / legal key)

IT administrators can enable the company key feature — which automatically adds
a designated legal or compliance key to every encrypted message — by publishing
a small JSON file on their domain.

### Step 1: Create the config file

```json
{
  "companyKeyEnabled": true,
  "companyKeyRequired": false,
  "companyKeyEmails": ["legal@your-company.com"]
}
```

See `docs/company-config.example.json` for the full documented template.

| Field | Type | Default | Meaning |
|-------|------|---------|---------|
| `companyKeyEnabled` | boolean | `false` | Whether the company key feature is active |
| `companyKeyRequired` | boolean | `false` | If `true`, users cannot opt out per-message |
| `companyKeyEmails` | string[] | `[]` | Email addresses whose keys are added to every encryption |
| `hideSupportButton` | boolean | `false` | If `true`, hides the Ko-fi support button from the Key Management pane (the external CDN script is never loaded) |

### Step 2: Publish it

Upload the file to your web server at one of the following paths (the add-in
tries them in order, using the first that returns a successful response):

```
Primary:  https://<your-email-domain>/.well-known/pgp-for-outlook-addin/company-config.json
Fallback: https://openpgpkey.<your-email-domain>/.well-known/pgp-for-outlook-addin/company-config.json
```

The add-in derives the URL from the signed-in user's email domain automatically.
For `alice@acme.com` it first tries
`https://acme.com/.well-known/pgp-for-outlook-addin/company-config.json`, then
`https://openpgpkey.acme.com/.well-known/pgp-for-outlook-addin/company-config.json`.

The fallback path lets organisations that already run a WKD server on
`openpgpkey.<domain>` co-locate the add-in config there without needing to
publish anything at the apex domain.

The file must be served without authentication.  If your add-in is hosted on a
different origin you may need to add a CORS header:

```
Access-Control-Allow-Origin: https://your-addin-host.example.com
```

### Step 3: Publish the company public key

The email address(es) in `companyKeyEmails` must have their PGP public keys
discoverable via **WKD** (preferred) or **VKS** (keys.openpgp.org).

### Fallback: manual override

If your org cannot host a well-known file, an admin (or the user themselves)
can set the org config manually via **Manage Keys → Organization Settings →
Manual Override → Save Override**.  This stores the config in the user's own
roaming settings and takes precedence over any well-known URL.

---

## First-run guide (for users)

### Initial setup

1. Open any email in Outlook and click **Manage Keys** in the ribbon.
2. Set up your key pair — choose one of:
   - **Generate New Key Pair** — choose ECC (recommended) or RSA-4096 for
     legacy compatibility, fill in your name, email, and a strong passphrase.
   - **Import Existing Key** — paste your armored private key block (from
     GnuPG, Kleopatra, Thunderbird, etc.) and enter its passphrase to verify.
     Any OpenPGP key type is accepted (RSA, ECC, DSA/ElGamal).
3. Click **Copy Public Key** and share it with contacts who need to send you
   encrypted mail (email it, upload to [keys.openpgp.org](https://keys.openpgp.org),
   or publish via WKD).
4. Optionally, go to **Personal Preferences** and enable **Sign encrypted
   messages by default** if you want signing on for every message without
   toggling it each time.

### Adding a contact's key

1. In **Manage Keys → Contacts' Keyring**, type the contact's email and click
   **Find** — the add-in checks WKD and keys.openpgp.org automatically.
2. If their key is found, verify the fingerprint with them out-of-band, then
   click **Save to Keyring**.
3. If no key is found, ask them to send you their public key and click
   **Import** to paste it in.

### Encrypting a message

1. Compose a new message and add recipients normally.
2. Click **Encrypt** in the ribbon.
3. The pane shows key status for each recipient.  Resolve any missing keys.
4. Optionally enable **Sign this message** (requires your passphrase at send
   time).  The toggle starts in the state you set in Personal Preferences and
   can be flipped for any individual message.
5. Click **Encrypt Message** and enter your passphrase if signing is enabled.
   - If the message contains **inline images** (e.g. pasted from the clipboard),
     a warning appears. On **Outlook on the web** you can click **Convert to
     Regular Attachments** to move them automatically. On **Outlook desktop**
     the Office API does not expose pasted images, so you must save the image
     to disk, remove it from the body, and re-attach it as a file before
     encrypting.
6. Click Outlook's normal **Send** button.

### Decrypting a message

1. Open the encrypted message and click **Decrypt** in the ribbon.
2. Enter your passphrase — the decrypted content appears in the task pane.
3. For `.pgp` attachments, click **Decrypt & Download** next to each file.

---

## Roadmap

Recommended enhancements not yet implemented, roughly in priority order:

### High priority

**Key expiration**
`generateKeyPair()` currently creates keys with no expiration.  Best practice
is a 2-year default.  Add a date picker to the generate form and pass
`keyExpirationTime` to `openpgp.generateKey()`.

**Revocation certificate**
When generating a key pair, immediately generate a revocation certificate and
prompt the user to save it.  If the private key is ever compromised, the cert
can be uploaded to keyservers to invalidate the key.

### Medium priority

**One-click public key import from message body**
When the read pane detects a `public-key` type in the message body, show an
**Import to Keyring** button alongside the existing instructions.

**Key refresh**
Periodically re-fetch contact keys from WKD/VKS to pick up revocations and
replacements.  A staleness threshold of ~7 days, checked silently on pane open,
would cover most use cases.

**Multiple UIDs per key**
OpenPGP keys can carry multiple email addresses.  The current keyring uses a
single email as the lookup key.  A fuller implementation would index by
fingerprint and maintain a multi-email reverse index.

### Lower priority

**Inline image conversion on Outlook desktop**
The Office.js `getAttachmentsAsync()` API does not expose clipboard-pasted
inline images in Outlook desktop (Win32 / Mac).  The add-in detects the broken
`cid:` reference in the body HTML and warns the user, but cannot read the image
data to re-attach it programmatically.  A future solution could use the
Microsoft Graph API (with delegated mail permissions) to retrieve the inline
attachment bytes and re-upload them — but this requires an OAuth token exchange
outside the scope of the current task-pane-only architecture.

**Test suite**
Unit tests for `pgp-core.js` (no DOM required) using Vitest or Jest.
Integration tests for the compose/read flows using Office mock libraries.

**Audit log**
For compliance environments, log encryption/decryption events (which keys,
which user, timestamp) to a company endpoint configurable via the org config.

---

## Development

### Prerequisites

- Any HTTPS-capable local development server
- A Microsoft 365 developer account
  ([free via M365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program))

### Run locally

The web app has no build step — it is plain HTML/CSS/ES modules.

```bash
# Install a dev server with self-signed HTTPS certs
npm install -g office-addin-dev-certs http-server
office-addin-dev-certs install

# Serve the web app
http-server web --ssl --port 3000
```

Update the manifest to point at `https://localhost:3000/`, then sideload it in
Outlook following
[Microsoft's sideloading guide](https://docs.microsoft.com/en-us/office/dev/add-ins/testing/test-debug-office-add-ins).

---

## Dependencies

| Library | Version | License | Purpose |
|---------|---------|---------|---------|
| [OpenPGP.js](https://openpgpjs.org/) | 5.5.0 | LGPL-3.0 | All cryptographic operations |
| [wkd-client](https://github.com/wiktor-k/openpgp-wkd) | bundled | LGPL-3.0 | WKD key lookup |
| [Office.js](https://docs.microsoft.com/en-us/javascript/api/overview/outlook) | CDN | MIT | Outlook add-in API |
| [Fluent UI Core](https://developer.microsoft.com/en-us/fluentui) | 9.6.0 | MIT | CSS design system |

---

## License

Copyright (C) 2025 Ruscal Technologies.
This project is licensed under the [GNU Affero General Public License v3.0](LICENSE) (AGPL-3.0).

In brief: you may use, modify, and distribute this software, but any modified
version — including one run as a network service — must also be released under
the AGPL-3.0 with source code made available.  See the [LICENSE](LICENSE) file
for the full terms.

Third-party dependencies are under their own licenses (see the Dependencies
table above).

## AI Assistance

Documentation and code review for this project may have been assisted by
[Claude Code](https://claude.ai/code) (Anthropic).  All content is the
property of Ruscal Technologies.

---

[![ko-fi](https://ko-fi.com/img/githubbutton_sm.svg)](https://ko-fi.com/R6R61WMZMW)
