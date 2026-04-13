# Seafile for Outlook

Office Web Add-in that integrates Seafile into Outlook (Web, Desktop, Mobile).
Built by datamate GmbH. Apache 2.0 license.

## Quick Reference

- **Outlook** via Office.js (Mailbox Requirement Set 1.8+), **Seafile Server** >= 10.0
- Vanilla JavaScript (ES2020+), no build tools, no dependencies, no TypeScript
- Add-in ID: `a1b2c3d4-e5f6-7890-abcd-ef1234567890` (manifest.xml)

## Project Structure

```
manifest.xml               # Office Add-in XML manifest (defines ribbon buttons, permissions, URLs)
api/seafile.js             # Seafile REST API client (SeafileAPI class)
shared/                    # Shared modules (loaded via <script> tags in all HTML files)
  shared.css               # Shared styles + CSS custom properties (light/dark theme)
  file-icons.js            # SVG file type icons + status icons
  password.js              # Cryptographic password generator
  utils.js                 # escapeHtml, formatSize, getHostLabel, showStatus, buildShareLinkHtml/Text
  storage.js               # localStorage-based account config (replaces Thunderbird's browser.storage.local)
  i18n.js                  # Lightweight i18n (JSON locale files, data-i18n attributes)
settings/                  # Account settings page (tabbed: Connection, Upload, Share Links, Save Attachments)
compose/                   # Compose task pane: upload attachments + insert Seafile links
read/                      # Read task pane: save received attachments to Seafile
assets/                    # SVG icons (16px, 32px, 80px) for Office ribbon
_locales/{en,de,fr,es,pt_BR,ru,zh_CN}.json  # i18n translations (flat JSON format)
dev/server.js              # Local HTTPS dev server with CORS proxy
dev/gen-certs.sh           # Generate self-signed certs for dev server
```

## Architecture

- **No background script** — unlike Thunderbird, Office.js add-ins have no persistent background. Each task pane (compose, read, settings) is a standalone HTML page that loads all shared modules directly.
- **api/seafile.js** contains the `SeafileAPI` class — all Seafile REST API interaction goes through this. In dev mode, API calls are routed through a local CORS proxy (`/seafile-proxy/{url}`).
- **State**: persistent config in `localStorage` via `shared/storage.js`. Each account stored under `seafile_{accountId}`. Account list in `seafile_accounts`.
- **Multi-account**: account selector in settings and task panes. Each account has independent config (server, token, library, folder, password mode, etc.).
- **Office Dialog API**: SSO login opens a browser dialog via `Office.context.ui.displayDialogAsync()` with fallback to `window.open()`.

## Build & Development

```bash
# Start dev server (HTTPS on localhost:3000 with CORS proxy)
cd dev && bash gen-certs.sh   # one-time: generate self-signed certs
node server.js                # serves all files + proxies Seafile API calls

# Sideload in Outlook
# Upload manifest.xml via Microsoft 365 admin center or Outlook's "Get Add-ins" dialog
# Dev manifest points to https://localhost:3000/
```

There is no linter, formatter, test suite, or CI pipeline.

## Code Conventions

- **Indentation**: 2 spaces
- **Semicolons**: always
- **Quotes**: double quotes preferred
- **Naming**: camelCase for functions/variables, PascalCase for classes, kebab-case for HTML element IDs
- **JSDoc**: all functions have JSDoc comments with @param/@returns
- **Security**: all user input escaped via `escapeHtml()` before DOM insertion — never use innerHTML with unescaped data
- **No modules**: scripts are loaded via HTML `<script>` tags in dependency order, not ES modules
- **Shared utilities**: common functions live in `shared/` — do not duplicate them in page-specific JS files

## Design & CSS

**No CSS framework** — no Tailwind, no Bootstrap. All styles are pure custom CSS.

**Shared styles** live in `shared/shared.css`, imported by all HTML files via `<link rel="stylesheet">`. Each HTML file keeps only page-specific styles in an inline `<style>` tag.

**CSS Custom Properties** are defined in `:root` in `shared.css`. Always use variables for colors, font sizes, etc. — never hardcode values that are already defined as variables. Key variable groups:
- `--color-primary`, `--color-danger` — action colors
- `--color-text`, `--color-text-secondary`, `--color-text-hint`, `--color-text-disabled`, `--color-text-icon` — text hierarchy
- `--color-border`, `--color-border-light`, `--color-border-lighter`, `--color-border-subtle` — border hierarchy
- `--color-bg-body`, `--color-bg-surface`, `--color-bg-input`, `--color-bg-button`, `--color-bg-highlight` — backgrounds
- `--color-success-*`, `--color-error-*`, `--color-info-*` — status colors
- `--color-hint-bg`, `--color-hint-border` — hint/defaults boxes
- `--font-size-base` (13px), `--font-size-small` (12px), `--font-size-hint` (11px) — typography
- `--radius`, `--focus-shadow`, `--select-arrow`, `--search-icon` — misc

**Dark mode**: fully supported via `@media (prefers-color-scheme: dark)` in `shared.css`. All color variables are overridden for dark mode. Native form controls use `color-scheme: light dark`. When adding new colors, always define both light and dark variants.

**SVG icons**: use `stroke="currentColor"` for theme-aware strokes. Fills use `fill-opacity` instead of solid fills to work on both light and dark backgrounds (e.g., folder icons use `fill="#e8a735" fill-opacity="0.15"`).

**Layout**: Flexbox only (no CSS Grid). `box-sizing: border-box` on all elements.

**Typography**: `"Segoe UI"` first (Fluent UI), then system font stack. Base `13px`, labels `font-weight: 600`.

**Email templates** (in `shared/utils.js`): `buildShareLinkHtml()` generates inline-styled HTML for email insertion. These colors are intentionally hardcoded because they render in the recipient's email client, not in the add-in.

## i18n

- Translation files in `_locales/` as flat JSON (`{ "key": "value" }`)
- HTML uses `data-i18n`, `data-i18n-placeholder`, `data-i18n-empty`, `data-i18n-title` attributes
- `I18N.init()` detects locale from `Office.context.displayLanguage`, falls back to `navigator.language`, then English
- 7 languages: English (primary), German, French, Spanish, Portuguese (BR), Russian, Chinese (Simplified)
- When adding user-visible strings, add keys to all locale files

## SeafileAPI Class (api/seafile.js)

All Seafile server communication goes through this class:

- **Auth**: `getToken()`, `getServerInfo()`, `createSSOLink()`, `checkSSOStatus()`, `getAccountInfo()`, `logout()`
- **Libraries**: `listRepos()`, `listDir()`, `dirExists()`, `createDir()`
- **Files**: `getUploadLink()`, `uploadFile()` (with AbortSignal support), `deleteFile()`
- **Share links**: `getShareLinks()`, `createShareLink()`, `deleteShareLink()`

The `proxyUrl()` helper routes calls through the dev proxy on localhost.

## Key Patterns

- **Token revocation**: on disconnect, calls `POST /api2/logout-device/` to invalidate the token server-side (best-effort, wrapped in try/catch)
- **SSO reconnect**: settings page shows a dedicated reconnect UI when an SSO session expires, preserving the auth method context instead of showing a generic login form
- **Password generation**: uses `crypto.getRandomValues()` (WebCrypto API), never Math.random()
- **Folder picker**: reusable collapsible tree component loaded via `loadFolderPicker()`
- **Auto-save**: settings changes are debounced (300ms) and saved via `autoSave()` with visual checkmark feedback
- **Existing link detection**: before creating a new share link, checks for existing links via `getShareLinks()`

## External Behavior: Never Guess, Always Verify

When reasoning about how **Seafile Server** or **Office.js** behaves (API responses, error codes, dialog behavior, etc.), **never guess or assume**. Always verify by reading the actual source code, testing against a running instance, or checking official documentation. If verification is not possible, explicitly state that the claim is unverified.

## Sister Project

This add-in is ported from [seafile-thunderbird](https://github.com/datamate-rethink-it/seafile-thunderbird), a Thunderbird WebExtension with the same Seafile integration. The API client, password generator, and file icons are shared across both projects.
