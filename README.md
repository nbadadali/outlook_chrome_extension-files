# Sa.AI for Outlook

Most inboxes show you emails. Yours shows you problems — the missed task, the forgotten thread, the reply you never got to.

Sa.AI is a Chrome extension that lives inside Outlook Web as a sidebar. It summarizes your inbox, extracts tasks, drafts replies, and answers questions about your emails — so you always know what matters and what needs your attention.

This is the Outlook counterpart to the [Sa.AI Gmail Assistant](https://github.com/nbadadali/gmail_chrome_extension-files), built with the same purpose and architecture adapted for Microsoft's platform.

## Features

- **Inbox Summary**: Opens with a clear summary of what actually matters in your inbox — no more scanning through everything manually
- **Task Extraction**: Pulls tasks buried inside email threads and turns them into a prioritized to-do list, kept up to date automatically
- **AI Chat**: Ask anything about your inbox, the web, or your day and get instant, context-aware answers
- **Voice Mode**: Switch to voice instead of typing — talk to your assistant naturally
- **Thread Summarization**: Get the key points of any long email thread in seconds, without scrolling through it
- **Draft Replies**: Ask the assistant to draft a reply and get a first version ready to send
- **Auto-Labeling**: Quietly labels and organizes incoming emails in the background — your inbox stays tidy without you touching a thing
- **Microsoft OAuth with PKCE**: Secure sign-in using Proof Key for Code Exchange — protects against authorization code interception
- **Automatic Token Refresh**: Silently refreshes your session on expiry without requiring you to sign in again
- **Session Persistence**: Stays connected across browser sessions — sign in once and you're set
- **SPA Compatibility**: Works with Outlook's single-page navigation across all domains (office.com, live.com, outlook.com)

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
2. The Sa.AI sidebar appears on the right side of the page
3. Your inbox summary loads automatically — tasks, priorities, and key threads
4. Type a message or switch to **voice mode** to ask questions
5. Try asking:
   - *"Summarize my inbox"*
   - *"What tasks do I have today?"*
   - *"Draft a reply to [sender]"*
   - *"Summarize this thread"* (when an email thread is open)

### Sidebar Controls
- **× button**: Closes the sidebar
- **Chat input**: Auto-focuses when the sidebar opens, ready to type immediately
- **Voice mode**: Switch from typing to speaking with your assistant
- **Connect button**: Appears when not signed in — starts the Microsoft OAuth flow

## How It Works

The extension is built around three components that communicate via Chrome's messaging API:

1. **Popup** (`popup.js`): Handles Microsoft OAuth sign-in and shows your connection status
2. **Background script** (`background.js`): Manages the PKCE OAuth flow, stores and refreshes JWTs, and forwards requests to n8n
3. **Content script** (`content-outlook.js`): Injects the sidebar into Outlook Web, manages the full chat and task UI, and extracts thread IDs from the URL for thread summarization

When you open Outlook, the content script injects the sidebar and loads your inbox summary via n8n. When you send a message or request a task list, the background script makes an authenticated request to the n8n webhook, which processes it using AI and returns the response to your sidebar.

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
- No email content or credentials are stored locally
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
