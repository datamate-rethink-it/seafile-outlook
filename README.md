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
- **Folder filter** — search/filter folders by name when there are many subfolders
- **Reset to defaults** — quickly return to your configured default library and folder
- **Rename before saving** — click any filename to rename it before uploading (name part auto-selected, reset button to restore)
- **Batch saving** — select multiple attachments at once, with synced "Select all" checkbox
- **Per-file status** — visual SVG feedback for each file during upload (with error tooltips)
- **Duplicate handling** — configurable: rename automatically (default) or overwrite existing files

### Authentication

- **Username & password** — standard Seafile login
- **Two-factor authentication (2FA)** — optional TOTP code field for accounts with 2FA enabled
- **Single Sign-On (SSO)** — login via browser using SAML, OAuth, Keycloak, or any SSO method configured on the server
- **Display name** — shows user display name and contact email from Seafile account info
- **Connection status** — clearly shows server, username, and authentication method (SSO or password)
- **Disconnect** — one-click disconnect with server-side token revocation
- **SSO reconnect** — dedicated reconnect UI when SSO session expires (instead of showing the password form)
- **HTTPS validation** — warns when connecting over HTTP to non-localhost servers

### Settings & UI

- **Multi-account support** — configure multiple Seafile accounts, switch between them in task panes (last used is remembered)
- **Tabbed settings** — Connection, Upload, Share Links, Save Attachments
- **Auto-save** — all configuration changes are saved immediately with visual feedback
- **Collapsible folder picker** — browse and select folders visually
- **Encrypted library filtering** — encrypted libraries are excluded automatically
- **Error handling** — clear messages for quota exceeded, file too large, too many files, and network errors
- **Localization** — English, German, French, Spanish, Portuguese (BR), Russian, Chinese (Simplified)
- **Dark mode** — automatically follows the system/Outlook theme preference
- **Cross-platform** — works in Outlook on Windows, Mac, and Outlook Web (OWA)

## Requirements

- **Outlook 2019** or later, **Microsoft 365**, or **Outlook on the web**
- **Seafile Server** 10.0 or later (Community Edition or Professional Edition)

## Installation

### Using the hosted version (recommended)

The add-in is hosted centrally by [datamate](https://datamate.org) at `https://outlook.datamate.org`. No self-hosting required.

1. **Configure CORS** on your Seafile server (see [CORS configuration](#cors-configuration) below)
2. **Deploy the add-in** via the **Microsoft 365 Admin Center** (requires M365 admin rights):
   1. Go to [admin.cloud.microsoft.com](https://admin.cloud.microsoft.com) → **Settings** → **Integrated Apps**
   2. Click **Upload custom apps**
   3. Set **App type** to **Office Add-in**
   4. Select **Provide link to manifest file** and enter: `https://outlook.datamate.org/manifest.xml`
   5. Click **Validate**, then **Next**
   6. Assign to users (individual, group, or entire organization) and complete the deployment

The add-in will appear in users' Outlook ribbon within a few hours (may take up to 24 hours due to caching).

> **Note:** The `https://aka.ms/olksideload` sideloading method no longer supports custom add-in uploads in the new Outlook. Use the Admin Center deployment described above instead.

### Self-hosting

You can also host the add-in files on your own HTTPS web server (e.g. the same server as Seafile):

1. Clone this repository and copy the files to your web server
2. Update all URLs in `manifest.xml` to point to your server
3. If hosted on the same domain as Seafile, no CORS configuration is needed
4. Deploy the manifest as described above

> **Note:** The add-in consists entirely of static files (HTML, CSS, JavaScript). There is no server-side code, no database, and no secrets in the source. The code is open source and always visible to anyone who can access the hosting URL — this is by design, as with any web application.

### For development

1. Clone this repository
2. Generate dev certificates and start the local server:
   ```bash
   ./dev/gen-certs.sh
   node dev/server.js
   ```
3. Open `https://localhost:3000` in your browser and accept the self-signed certificate
4. The task panes are accessible directly in the browser for UI testing:
   - `https://localhost:3000/settings/settings.html` — Settings (login, SSO, all config)
   - `https://localhost:3000/compose/compose.html` — Compose (upload, insert link)
   - `https://localhost:3000/read/read.html` — Read (save attachments)
5. To test inside Outlook, deploy via the Admin Center as described above (localhost URLs won't work — use a public HTTPS URL)

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

When using the hosted version (or any setup where the add-in and Seafile are on different domains), the Seafile server must allow cross-origin requests from the add-in domain. Without CORS, the add-in can authenticate but all subsequent API calls will fail silently.

**Important:** CORS headers must be set on the **outermost proxy** — the one that terminates the TLS connection from the browser. If you have Caddy in front of nginx, configure CORS in Caddy, not in nginx. If nginx is your only/outermost proxy, configure CORS there.

You can verify your CORS setup with:

```bash
curl -I -X OPTIONS \
  -H "Origin: https://outlook.datamate.org" \
  -H "Access-Control-Request-Method: GET" \
  -H "Access-Control-Request-Headers: Authorization" \
  https://YOUR-SEAFILE-SERVER/api2/account/info/
```

The response must include `Access-Control-Allow-Origin: https://outlook.datamate.org` and return status `204`.

#### nginx

Add the following to your nginx `server` block (in the `location /` block). The `proxy_hide_header` line removes the default `Access-Control-Allow-Origin: *` that Seahub sets, preventing duplicate values:

```nginx
# CORS for Seafile Outlook Add-in
proxy_hide_header Access-Control-Allow-Origin;
add_header Access-Control-Allow-Origin "https://outlook.datamate.org" always;
add_header Access-Control-Allow-Methods "GET, POST, PUT, DELETE, OPTIONS" always;
add_header Access-Control-Allow-Headers "Authorization, Content-Type, X-SEAFILE-OTP" always;

if ($request_method = OPTIONS) {
    return 204;
}
```

#### Caddy (Caddyfile)

```
seafile.example.com {
    header Access-Control-Allow-Origin "https://outlook.datamate.org"
    header Access-Control-Allow-Methods "GET, POST, PUT, DELETE, OPTIONS"
    header Access-Control-Allow-Headers "Authorization, Content-Type, X-SEAFILE-OTP"

    @options method OPTIONS
    respond @options 204

    reverse_proxy localhost:80 {
        header_down -Access-Control-Allow-Origin
    }
}
```

#### Caddy Docker Proxy (lucaslorentz/caddy-docker-proxy)

Add these labels to your Seafile service in `docker-compose.yml`:

```yaml
labels:
  caddy: seafile.example.com
  caddy.header.Access-Control-Allow-Origin: '"https://outlook.datamate.org"'
  caddy.header.Access-Control-Allow-Methods: '"GET, POST, PUT, DELETE, OPTIONS"'
  caddy.header.Access-Control-Allow-Headers: '"Authorization, Content-Type, X-SEAFILE-OTP"'
  caddy.@options.method: OPTIONS
  caddy.respond: "@options 204"
  caddy.reverse_proxy: "{{upstreams 80}}"
  caddy.reverse_proxy.header_down_1: -Access-Control-Allow-Origin
```

The `header_down_1` line removes the `Access-Control-Allow-Origin` header that Seahub sets by default (`*`), so only the correct value from Caddy's `header` directive remains. Without this, the browser receives two conflicting values and rejects the response.

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
├── assets/                    # Add-in icons (PNG required by Microsoft, SVG sources included)
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
