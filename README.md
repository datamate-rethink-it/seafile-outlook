# Seafile for Outlook

An Outlook add-in that integrates [Seafile](https://www.seafile.com) into Microsoft Outlook. Email attachments can be uploaded to your Seafile server and replaced with download links. Existing files on Seafile can be browsed and inserted as share links into emails. Received attachments can be saved directly to Seafile.

Built with the modern [Office Web Add-in](https://learn.microsoft.com/en-us/office/dev/add-ins/) platform (Office.js). This is the successor to the legacy VSTO-based Seafile Outlook Add-in.

Built by [datamate](https://datamate.org), the Seafile partner for Europe.

## Features

### Upload attachments (outgoing)

- **Upload & replace** — upload email attachments to Seafile and replace them with download links
- **Direct file upload** — upload local files via file picker or drag & drop (no size limitation from the Outlook attachment API)
- **Share link creation** — each uploaded file gets a Seafile download link inserted into the email
- **Password protection** — optionally protect share links with a password
- **Link expiration** — set an automatic expiry (in days) for share links
- **Existing link handling** — if a share link already exists, it is replaced automatically
- **Password modes** — no password, random password (configurable length), or custom password
- **Rich email template** — inserted links include file name, size, link URL, password info, and Seafile logo

### Insert Seafile links (compose)

- **Browse & insert** — browse your Seafile libraries and folders directly in Outlook's task pane
- **File selection** — click a file to select it, then configure link options before inserting
- **File type icons** — color-coded SVG icons for common file types (PDF, images, spreadsheets, archives, audio, video, code, etc.)
- **Password & expiration** — set password and expiration per link, or use defaults from settings
- **Password generator** — generate secure random passwords with one click (cryptographically secure, configurable length)
- **Show password in email** — choose to display the password in the email or show a "sent separately" hint
- **File filter** — search/filter files by name when folders contain many entries
- **Existing link detection** — reuse existing share links or create a new one
- **Cursor position insert** — links are inserted at the cursor position without modifying existing email content

### Save attachments (incoming)

- **Save attachments to Seafile** — click the Seafile button in the ribbon to save received attachments
- **Library & folder selection** — choose target library and navigate folders with a collapsible folder picker
- **Rename before saving** — click any filename to rename it before uploading (name part auto-selected, reset button to restore)
- **Batch saving** — select multiple attachments at once, with synced "Select all" checkbox
- **Per-file status** — visual SVG feedback for each file during upload
- **Duplicate handling** — configurable: rename automatically (default) or overwrite existing files

### Authentication

- **Username & password** — standard Seafile login
- **Two-factor authentication (2FA)** — optional TOTP code field for accounts with 2FA enabled
- **Single Sign-On (SSO)** — login via browser using SAML, OAuth, Keycloak, or any SSO method configured on the server
- **Display name** — shows user display name and contact email from Seafile account info
- **Connection status** — clearly shows server, username, and authentication method (SSO or password)
- **Disconnect** — one-click disconnect with automatic cleanup
- **HTTPS validation** — warns when connecting over HTTP to non-localhost servers

### Settings & UI

- **Multi-account support** — configure multiple Seafile accounts, switch between them in task panes (last used is remembered)
- **Tabbed settings** — Connection, Upload, Share Links, Save Attachments
- **Auto-save** — all configuration changes are saved immediately with visual feedback
- **Collapsible folder picker** — browse and select folders visually
- **Encrypted library filtering** — encrypted libraries are excluded automatically
- **Localization** — English, German, French, Spanish, Portuguese (BR), Russian, Chinese (Simplified)
- **Dark mode** — automatically follows the system/Outlook theme preference
- **Cross-platform** — works in Outlook on Windows, Mac, and Outlook Web (OWA)

## Requirements

- **Outlook 2019** or later, **Microsoft 365**, or **Outlook on the web**
- **Seafile Server** 10.0 or later (Community Edition or Professional Edition)

## Installation

### Using the hosted version (recommended)

The add-in is hosted centrally by [datamate](https://datamate.org) at `https://outlook.de.seafile.com`. No self-hosting required.

1. **Configure CORS** on your Seafile server (see [CORS configuration](#cors-configuration) below)
2. **Download the manifest** from `https://outlook.de.seafile.com/manifest.xml`
3. **Deploy the manifest** to your users:
   - **Organization-wide**: Upload via **Microsoft 365 Admin Center** → **Integrated Apps** and assign to users/groups
   - **Individual**: Go to `https://aka.ms/olksideload`, click **Add a custom add-in** → **Add from File**, and upload `manifest.xml`

The add-in will appear automatically in users' Outlook ribbon.

### Self-hosting

You can also host the add-in files on your own HTTPS web server (e.g. the same server as Seafile):

1. Clone this repository and copy the files to your web server
2. Update all URLs in `manifest.xml` to point to your server
3. If hosted on the same domain as Seafile, no CORS configuration is needed
4. Deploy the manifest as described above

> **Note:** The add-in consists entirely of static files (HTML, CSS, JavaScript). There is no server-side code, no database, and no secrets in the source. The code is open source and always visible to anyone who can access the hosting URL — this is by design, as with any web application.

### For development (sideloading)

1. Clone this repository
2. Generate dev certificates and start the local server:
   ```bash
   ./dev/gen-certs.sh
   node dev/server.js
   ```
3. Open `https://localhost:3000` in your browser and accept the self-signed certificate
4. Sideload the manifest:
   - **Outlook Web**: Go to `https://aka.ms/olksideload` and upload `manifest.xml`
   - **Outlook Desktop**: Settings → Manage Add-ins → Upload custom add-in

The dev server includes a CORS proxy so you can test against any Seafile server without configuring CORS headers.

## Configuration

After installation, click **Seafile Settings** in the Outlook ribbon.

1. Enter your **Seafile server URL** (e.g. `https://cloud.seafile.com`). HTTPS is strongly recommended.
2. Log in using one of two methods:
   - **Username/password**: Enter credentials and optionally a **2FA code**, then click **Connect**
   - **SSO**: Click **Login via SSO** — a browser window opens for authentication
3. **Upload tab**: Select target library and upload folder, configure password protection (none/random/custom) and link expiration
4. **Share Links tab**: Set defaults for share links when inserting Seafile links (password, expiration, show password in email)
5. **Save Attachments tab**: Select default library and folder for saving received attachments

All settings are saved automatically.

### SSO setup

To use SSO login, the Seafile server admin must enable client SSO in `seahub_settings.py`:

```python
CLIENT_SSO_VIA_LOCAL_BROWSER = True
```

This works with any SSO method configured on the server (SAML, OAuth, Keycloak, Shibboleth, etc.).

### CORS configuration

When using the hosted version (or any setup where the add-in and Seafile are on different domains), the Seafile server must allow cross-origin requests from the add-in domain.

Add the following to your Seafile **nginx configuration** (typically in the `location /` or `location /api` block):

```nginx
# CORS for Seafile Outlook Add-in
add_header Access-Control-Allow-Origin "https://outlook.de.seafile.com" always;
add_header Access-Control-Allow-Methods "GET, POST, PUT, DELETE, OPTIONS" always;
add_header Access-Control-Allow-Headers "Authorization, Content-Type, X-SEAFILE-OTP" always;

# Handle preflight requests
if ($request_method = OPTIONS) {
    return 204;
}
```

If you self-host the add-in on the same domain as Seafile, no CORS configuration is needed.

## Usage

### Uploading attachments

When composing an email, click **Upload to Seafile** in the ribbon. The task pane shows two options:
- **Email attachments** — select existing attachments to upload and replace with links
- **File picker** — click or drag & drop files to upload directly from your computer

### Inserting Seafile links

Click **Insert Seafile Link** in the compose ribbon. Browse your Seafile libraries, select a file, optionally set password and expiration, and click **Insert link into email**. The link is inserted at the cursor position.

### Saving attachments

When reading an email, click **Save to Seafile** in the ribbon. Select attachments, choose a target library and folder, and click **Save to Seafile**.

## Project Structure

```
├── manifest.xml               # Office Add-in XML manifest
├── api/
│   └── seafile.js             # Seafile API client
├── shared/
│   ├── file-icons.js          # SVG file type + status icons
│   ├── password.js            # Cryptographic password generator
│   ├── utils.js               # Shared utilities (escapeHtml, formatSize, link templates)
│   ├── storage.js             # Account storage manager (localStorage)
│   ├── i18n.js                # Localization helper
│   └── shared.css             # Common styles
├── settings/
│   ├── settings.html          # Account configuration page (tabbed)
│   └── settings.js            # Configuration logic
├── compose/
│   ├── compose.html           # Compose task pane (upload + insert link)
│   └── compose.js             # Upload and insert logic
├── read/
│   ├── read.html              # Read task pane (save attachments)
│   └── read.js                # Save attachments logic
├── assets/                    # Add-in icons
├── _locales/                  # Translations (en, de, fr, es, pt_BR, ru, zh_CN)
└── PRIVACY.md                 # Privacy policy
```

## Technology

- **Office.js** — Microsoft's web add-in platform for Outlook
- **Vanilla JavaScript** — no frameworks, no build step, no dependencies
- **Requirement Set 1.8** — for `getAttachmentContentAsync` (Outlook 2019+)
- **Seafile Web API** v2 / v2.1

### Why no React/TypeScript/bundler?

This add-in follows the same architecture as the [Seafile Thunderbird extension](https://github.com/datamate-rethink-it/seafile-thunderbird): vanilla JavaScript, no build step, no external dependencies. This means:
- The Seafile API client (`api/seafile.js`) shares the same API logic (endpoints, parameters, response parsing) with the Thunderbird version, adapted for Outlook's web context (CORS proxy for dev, no browser.i18n in the API layer)
- No Node.js/npm required for development or deployment
- Simple debugging (source = deployed code)
- Lower maintenance burden

### Comparison with the legacy VSTO add-in

| | Legacy VSTO Add-in | This Add-in |
|---|---|---|
| **Platform** | Windows only | Windows, Mac, Web |
| **Technology** | C# / .NET Framework / ClickOnce | JavaScript / Office.js |
| **Certificate** | Code-signing required | Not required (HTTPS only) |
| **Deployment** | Per-client installer | Centralized via Admin Center |
| **Features** | Upload only | Upload + Insert Link + Save Attachments |
| **Status** | Unmaintained since 2022 | Actively developed |

## Privacy

This add-in does not collect or share any data with third parties. All data is stored locally in the browser's localStorage (within the Outlook add-in sandbox) and communicated exclusively with your configured Seafile server. Your Seafile password is never stored — only the API token is persisted.

See [PRIVACY.md](PRIVACY.md) for the full privacy policy.

## Contributing

Contributions are welcome! Please open an issue to discuss your idea before submitting a pull request.

- Bug reports: [GitHub Issues](https://github.com/datamate-rethink-it/seafile-outlook/issues)
- Translation improvements: Edit the files in `_locales/` and submit a PR

## License

[Apache License 2.0](LICENSE)
