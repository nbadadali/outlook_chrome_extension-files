# Sa.AI for Outlook

An AI-powered Chrome extension that adds a smart sidebar to Outlook Web, letting you chat with an AI assistant about your inbox, summarize email threads, and manage your email workflow — all without leaving Outlook.

This is the Outlook counterpart to the [Sa.AI Gmail Assistant](https://github.com/nbadadali/gmail_chrome_extension-files), built with the same architecture adapted for Microsoft's platform.

## Features

- **Outlook Sidebar**: Sidebar injected directly into Outlook Web across all domains (office.com, live.com, outlook.com)
- **Microsoft OAuth with PKCE**: Secure sign-in using Proof Key for Code Exchange — a more secure OAuth flow that protects against authorization code interception
- **AI Chat Interface**: Ask questions about your inbox, emails, and tasks
- **Thread Summarization**: Open any email thread and ask the AI to summarize it — the extension reads the thread ID from the URL automatically
- **Automatic Token Refresh**: Silently refreshes your session on expiry without requiring you to sign in again
- **Extension Reload Recovery**: Detects when the extension has been reloaded and prompts a clean page refresh rather than silently failing
- **Session Persistence**: Stays connected across browser sessions
- **SPA Compatibility**: Works with Outlook's single-page navigation

## Tech Stack

- **Platform**: Chrome Extension (Manifest V3)
- **Language**: JavaScript (Vanilla)
- **Auth**: Microsoft OAuth 2.0 with PKCE via Azure App Registration and `chrome.identity`
- **Automation**: n8n (workflow automation for AI and email processing)
- **API**: Microsoft Graph API (`User.Read`, `Mail.Read`)
- **Styling**: CSS

## Installation

1. Clone or download this repository
2. Open Chrome and go to `chrome://extensions/`
3. Enable **Developer mode** (toggle in the top right)
4. Click **Load unpacked** and select the `outlook_chrome_extension_files` folder
5. The Sa.AI icon will appear in your Chrome toolbar

## Usage

### Connecting Your Outlook Account
1. Click the Sa.AI icon in your Chrome toolbar
2. Click **Connect to Microsoft & Sa.AI**
3. Complete the Microsoft sign-in and permissions screen
4. Once connected, click **Open Chat Assistant**

### Using the Assistant
1. Open Outlook Web at [outlook.office.com](https://outlook.office.com) or [outlook.live.com](https://outlook.live.com)
2. The Sa.AI sidebar will appear
3. Type a message and press **Enter** or click **Send**
4. Try asking:
   - *"Summarize my inbox"*
   - *"What emails need my attention?"*
   - *"Summarize this thread"* (when an email thread is open)

### Sidebar Controls
- **× button**: Closes the sidebar
- **Connect button**: Appears when not signed in — starts the Microsoft OAuth flow

## How It Works

The extension is built around three components that communicate via Chrome's messaging API:

1. **Popup** (`popup.js`): Handles Microsoft OAuth sign-in and shows your connection status
2. **Background script** (`background.js`): Manages the PKCE OAuth flow, stores and refreshes JWTs, and forwards chat messages to n8n
3. **Content script** (`content-outlook.js`): Injects the sidebar into Outlook Web, manages the chat UI, and extracts thread IDs from the URL for thread summarization

### Authentication Flow

The extension uses PKCE (Proof Key for Code Exchange) for a more secure sign-in:

1. The extension requests PKCE parameters from the n8n start webhook
2. A Microsoft login pop-up opens via Chrome's identity API
3. After sign-in, Microsoft sends an authorization code to the n8n callback
4. n8n exchanges the code for tokens, generates a JWT, and redirects back to the extension
5. The extension stores the JWT securely and uses it for all future requests

### n8n Webhooks

| Purpose | Endpoint |
|---------|----------|
| OAuth Start | `https://connector.saai.dev/webhook/outlook/oauth/start` |
| OAuth Callback | `https://connector.saai.dev/webhook/outlook/oauth/callback` |
| AI Chat & Thread Summarization | `https://connector.saai.dev/webhook/outlook/cahtbot` |
| Session Renewal | `https://connector.saai.dev/webhook/outlook/session/renew-outlook` |
| Credit Tracking | `https://connector.saai.dev/webhook/Credit-Tracking` |

## Project Structure

```
├── manifest.json          # Extension config — permissions, scripts, OAuth scopes
├── background.js          # Service worker — OAuth (PKCE), JWT management, n8n requests
├── background-outlook.js  # Additional background utilities
├── content-outlook.js     # Outlook sidebar — chat UI, thread ID extraction
├── popup.js               # Toolbar popup — connect/disconnect, status
├── popup.html             # Popup HTML
├── styles.css             # Sidebar and popup styles
└── icons/                 # Extension icons (16px, 48px, 128px)
```

## Azure Configuration

The extension authenticates via a Microsoft Azure app. To set it up:

1. Register an app in [Azure Active Directory](https://portal.azure.com)
2. Under **Authentication**, add two redirect URIs:
   - `https://connector.saai.dev/webhook/outlook/oauth/callback` — for n8n to receive the authorization code
   - `https://<your-extension-id>.chromiumapp.org/` — for Chrome to capture the final redirect

To find your extension's redirect URI, open the popup and check the browser console — the URI is logged on load.

See `OAUTH_SETUP.md` for the full step-by-step configuration guide.

## Security

- JWT tokens are stored in `chrome.storage.local`, sandboxed to the extension
- PKCE prevents authorization code interception attacks
- An OAuth concurrency lock prevents duplicate sign-in flows from running at the same time
- Thread IDs are validated for format and length before being sent to the backend
- All communication with n8n and Microsoft uses HTTPS
- All console logging is silenced in production builds

## Version History

### v2.2.3
- Thread summarization: reads the thread ID directly from the Outlook URL and validates it before sending to the backend
- Extension reload recovery: prompts the user to refresh the page instead of silently failing
- Automatic JWT refresh on session expiry, with up to two retry attempts
- Silent sign-out when a Microsoft token is permanently revoked
- Fallback responses when the n8n backend is unreachable
- OAuth concurrency lock to prevent duplicate sign-in flows

### v2.0
- Initial Outlook implementation based on the Gmail extension architecture
- Microsoft PKCE OAuth flow via n8n
- Sidebar chat UI adapted for Outlook Web
- JWT-based authentication for all n8n requests
