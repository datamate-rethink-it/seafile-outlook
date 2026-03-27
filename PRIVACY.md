# Privacy Policy — Seafile for Outlook

## Data Collection

This add-in does **not** collect, transmit, or store any data on external servers operated by the add-in developer.

## Data Flow

All communication occurs exclusively between:
1. **Your Outlook client** (where the add-in runs)
2. **Your configured Seafile server** (which you specify in the settings)

No data is sent to any third party.

## Stored Data

The add-in stores the following data locally in your browser's localStorage (within the Outlook add-in sandbox):

- **Seafile server URL** — the address of your Seafile server
- **API token** — an authentication token obtained from your Seafile server (your password is never stored)
- **Display name and email** — fetched from your Seafile account for display purposes
- **Preferences** — your configured library, folder, password, and expiration settings

This data remains on your device and is not transmitted anywhere except to your Seafile server during API calls.

## Permissions

The add-in requests the following Outlook permissions:

- **ReadWriteMailbox** — required to read email attachments (for uploading to Seafile and saving to Seafile) and to insert share links into email bodies

## Open Source

This add-in is open source. You can review the complete source code to verify these claims.
