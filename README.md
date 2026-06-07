# Sa.AI for Outlook

An AI-powered Chrome extension that integrates with Outlook Web to provide intelligent inbox summarization, thread analysis, and chat functionality through n8n workflows and Microsoft Graph API.

## Features

- **Seamless Outlook Integration**: Sidebar injected directly into Outlook Web (office.com, live.com, outlook.com)
- **Microsoft OAuth (PKCE)**: Secure authentication flow using Proof Key for Code Exchange via n8n
- **AI Chat Interface**: Interactive chat with AI assistant for email queries and inbox management
- **Thread Summarization**: Summarize specific email threads by opening them and asking the AI
- **Credit Tracking**: Automatic usage tracking sent to the backend after each interaction
- **JWT Token Management**: Automatic token refresh on 401/403 with retry logic
- **Extension Reload Detection**: Detects orphaned content scripts and prompts a page refresh
- **State Persistence**: Connection status, JWT, and user ID persisted across browser sessions
- **SPA Compatibility**: Handles Outlook's single-page navigation with on-demand script injection
- **Fallback Responses**: Graceful degradation with helpful messages when n8n is unavailable

## Installation

1. Clone or download this repository
2. Open Chrome and navigate to `chrome://extensions/`
3. Enable "Developer mode" in the top right
4. Click "Load unpacked" and select the `outlook_chrome_extension_files` folder
5. The Sa.AI extension icon should appear in your Chrome toolbar

## Configuration

### Azure App Registration

The extension uses a Microsoft Azure OAuth app for authentication:

- **Client ID**: `029d6ab9-7fbf-4476-9bf5-deb353bb79a6` (configured in `background.js`)
- **Scopes**: `openid`, `profile`, `email`, `offline_access`, `User.Read`, `Mail.Read`

You must register **two** redirect URIs in Azure → App Registrations → Authentication:
1. `https://connector.saai.dev/webhook/outlook/oauth/callback` (n8n callback)
2. `https://<your-extension-id>.chromiumapp.org/` (Chrome identity redirect — see OAUTH_SETUP.md)

To get the extension's redirect URI, open the popup and check the browser console for:
```
🔑 EXTENSION REDIRECT URI (add this to Azure): https://<extension-id>.chromiumapp.org/
```

### n8n Integration

All AI and authentication logic runs through n8n webhooks on `connector.saai.dev`:

| Purpose | Endpoint |
|---------|----------|
| OAuth start (PKCE) | `https://connector.saai.dev/webhook/outlook/oauth/start` |
| OAuth callback | `https://connector.saai.dev/webhook/outlook/oauth/callback` |
| Chat & thread summarization | `https://connector.saai.dev/webhook/outlook/cahtbot` |
| Session renewal | `https://connector.saai.dev/webhook/outlook/session/renew-outlook` |
| Credit tracking | `https://connector.saai.dev/webhook/Credit-Tracking` |
| Task management | `https://connector.saai.dev/webhook/task-management-outlook` |
| Feedback / error reporting | `https://connector.saai.dev/webhook/Feedback_error` |
| Heartbeat (disabled by default) | `https://connector.saai.dev/webhook/oauth/heartbeat` |

### Feature Flags

Located in `background.js` inside the `FEATURES` object:

```js
const FEATURES = {
    heartbeatEnabled: false,          // Enable to activate periodic session heartbeats
    outlookTokenRefreshEnabled: true  // Enable automatic JWT refresh on 401/403
};
```

Set `outlookTokenRefreshEnabled: false` if the `outlook/session/renew-outlook` webhook is not yet deployed.

## Usage

### Initial Setup
1. Click the Sa.AI extension icon in your Chrome toolbar
2. Click **"Connect to Microsoft & Sa.AI"**
3. Complete the Microsoft login and consent screen
4. Once connected, click **"Open Chat Assistant"**

### Using the Assistant
1. Navigate to Outlook Web (`outlook.office.com`, `outlook.live.com`, or `outlook.com`)
2. The Sa.AI sidebar will appear on the page
3. Type your questions in the chat input
4. Press **Enter** or click **Send** to submit
5. To summarize a specific thread — open the email thread first, then ask "Summarize this thread"

### Sidebar Controls
- **Close Button (×)**: Closes the sidebar
- **Chat Input**: Type messages and press Enter or click Send
- **Connect Button**: Appears when not authenticated — starts the OAuth flow

## OAuth Flow (PKCE)

The extension uses PKCE (Proof Key for Code Exchange) for secure Microsoft OAuth:

1. Extension calls `connector.saai.dev/webhook/outlook/oauth/start` to get PKCE parameters (`state`, `code_challenge`)
2. Extension launches Microsoft login via `chrome.identity.launchWebAuthFlow()`
3. Microsoft redirects to the n8n callback URI with an authorization code
4. n8n exchanges the code for tokens, generates a JWT, and **redirects** to the extension's chromiumapp.org URI with `?jwt=...&userId=...`
5. Extension captures the redirect URL and stores the JWT and userId in `chrome.storage.local`

> See `OAUTH_SETUP.md` for the full n8n redirect configuration required for this flow to work.

## Chat Flow

When a message is sent from the sidebar:

1. **Content script** (`content-outlook.js`) → `handleSendMessage()` builds the payload: `{ query, userId, context: 'OutlookChat', ... }`
2. **Content** → Background via `chrome.runtime.sendMessage({ action: 'sendToN8N', data: { endpoint: 'chat', payload } })`
3. **Background** (`background.js`) → `handleN8NRequest()` fetches the JWT via `ensureValidJWTToken()`
4. **Background** → POSTs to `outlook/cahtbot` with `Authorization: Bearer <jwt>`
5. **Content** receives and renders the AI response; on error, displays `Error: <message>`

### Thread Summarization
If the payload contains `{ action: 'summarize_thread', threadId: '...' }`, the same `outlook/cahtbot` webhook is used with the thread ID and subject line attached to the payload. The thread ID is extracted from the Outlook URL path (e.g. `/mail/0/inbox/id/<threadId>`) and validated for length (20–600 chars) and base64-like charset before being sent.

## Token Management

- **On startup**: The extension checks for an existing JWT in storage without proactively refreshing it.
- **On 401/403**: If `outlookTokenRefreshEnabled` is `true`, the background script automatically calls `outlook/session/renew-outlook` with the refresh token and retries the original request up to twice.
- **On `invalid_grant`**: The user is silently disconnected and prompted to reconnect.
- **On extension reload**: Orphaned content scripts detect the stale connection and show: *"The extension was reloaded or updated. Please refresh this page (F5) to connect again."*

## Testing Instructions

### 1. Basic Functionality Test
```
1. Load the extension in Chrome
2. Open Outlook Web (outlook.office.com or outlook.live.com)
3. Click extension icon → "Connect to Microsoft & Sa.AI"
4. Complete Microsoft OAuth flow
5. Click "Open Chat Assistant"
6. Verify sidebar appears
7. Test typing in chat input and sending a message
8. Test closing sidebar with × button
```

### 2. OAuth Flow Test
```
1. Clear extension storage: chrome.storage.local.clear()
2. Reload extension
3. Click "Connect to Microsoft & Sa.AI" in popup
4. Confirm Microsoft login screen appears
5. After login, check storage in DevTools console:
   chrome.storage.local.get(['userId', 'jwtToken', 'isConnected', 'provider'], console.log)
6. Expected: { userId: "...", jwtToken: "...", isConnected: true, provider: "outlook" }
```

### 3. Chat Functionality Test
```
1. Ensure connected (sidebar shows chat UI, not Connect prompt)
2. Type "hello" and press Enter
3. Verify message appears in the chat
4. Check for AI response or meaningful error message
5. Open DevTools → check background service worker console for [SaAI-BG] logs
```

### 4. Thread Summarization Test
```
1. Open a specific email thread in Outlook Web
2. Open the Sa.AI sidebar
3. Type "Summarize this thread"
4. Verify the sidebar sends the correct threadId (check [SaAI-Content] logs)
5. Check for AI-generated summary in response
```

### 5. Token Refresh Test
```
1. Manually expire the JWT (or wait for natural expiry)
2. Send a chat message
3. Verify the extension automatically refreshes the token (check background logs)
4. Verify the original message is retried and returns a response
```

### 6. Extension Reload Test
```
1. Open Outlook Web with the sidebar active
2. Go to chrome://extensions and click Reload on Sa.AI
3. Return to Outlook and try sending a message
4. Verify the extension shows the "reload detected" message
5. Refresh the Outlook tab (F5) and verify the sidebar works again
```

## Troubleshooting

### Sidebar Not Opening
- Verify the content script is injected: open DevTools and check for `[SaAI-Content]` logs
- Refresh the Outlook tab — extension reloads require a tab refresh to re-inject the script
- Check `chrome://extensions` → Sa.AI for Outlook → Errors

### OAuth Issues
- Confirm both redirect URIs are registered in Azure (n8n callback + chromiumapp.org URI)
- Verify n8n `outlook/oauth/callback` workflow **redirects** to the extension URI instead of returning JSON
- Check the background service worker console for `[SaAI-BG]` error logs
- Clear storage and retry: `chrome.storage.local.clear()`

### Chat Not Responding
- Verify the JWT is present: `chrome.storage.local.get(['jwtToken'], console.log)`
- Check the background service worker for "No JWT token found" or timeout errors
- Confirm `https://connector.saai.dev/webhook/outlook/cahtbot` is deployed and accepting `context: 'OutlookChat'`
- Inspect n8n execution logs for 4xx/5xx errors

### "Receiving End Does Not Exist"
- This means the extension was reloaded but the Outlook tab still runs the old content script
- Fix: Refresh the Outlook tab (F5) — no need to reconnect, storage is preserved

### "Session Expired / Please Reconnect"
- JWT expired and token refresh failed (or `outlookTokenRefreshEnabled` is `false`)
- Click Connect in the sidebar to start a fresh OAuth flow
- Long-term fix: ensure `outlook/session/renew-outlook` is deployed and returns a `jwt` field

### Redirect URI Mismatch
- The exact chromiumapp.org URI must be added to Azure (including the trailing slash)
- Get the correct URI from the popup console: `🔑 EXTENSION REDIRECT URI: https://...chromiumapp.org/`

## Debug Mode

To enable verbose logging, set `PRODUCTION_MODE = false` in both:
- `content-outlook.js` (line ~6)
- `background.js` (line ~5, already set to `false` during development)

### Console Log Prefixes
- `[SaAI-Content]` — Content script logs (sidebar, chat, thread extraction)
- `[SaAI-BG]` — Background script logs (OAuth, JWT, n8n requests)
- `[SaAI-Popup]` — Popup script logs (connect flow, status)

### Inspect Storage
Open DevTools → Console on the Outlook page:
```js
chrome.storage.local.get(['userId', 'jwtToken', 'isConnected', 'provider', 'refreshToken'], console.log)
```

### Storage Keys Reference

| Key | Description |
|-----|-------------|
| `userId` | User identifier from OAuth |
| `isConnected` | Connection status (`true`/`false`) |
| `jwtToken` | JWT for authenticated n8n requests |
| `refreshToken` | Used to renew the JWT without re-login |
| `provider` | Set to `'outlook'` to distinguish from Gmail |
| `oauthData` | Full OAuth response metadata |
| `oauthInProgress` | Lock flag to prevent concurrent OAuth flows |
| `tokenRefreshCount` | Number of times the JWT has been refreshed |
| `lastSuccessfulRefresh` | Timestamp of most recent successful refresh |
| `sessionActive` | Heartbeat-tracked session activity (when enabled) |

## Development

### File Structure
```
├── manifest.json              # Extension configuration (Manifest V3)
├── background.js              # Service worker (OAuth, JWT, n8n forwarding)
├── background-outlook.js      # Additional background utilities
├── content-outlook.js         # Content script (sidebar, chat, thread extraction)
├── popup.js                   # Popup script (connect flow, status UI)
├── popup.html                 # Popup HTML
├── styles.css                 # CSS styles for sidebar and popup
└── icons/                     # Extension icons (16px, 48px, 128px)
```

### Key Functions

| Function | File | Purpose |
|----------|------|---------|
| `handleOAuthFlow()` | `background.js` | Launches Microsoft PKCE OAuth via chrome.identity |
| `ensureValidJWTToken()` | `background.js` | Retrieves JWT from storage; throws if missing |
| `refreshJWTToken()` | `background.js` | Calls `outlook/session/renew-outlook` to get a new JWT |
| `handleN8NRequest()` | `background.js` | Routes requests to the correct n8n webhook endpoint |
| `handleFallbackResponse()` | `background.js` | Returns graceful offline responses when n8n is down |
| `silentlyDisconnectUser()` | `background.js` | Clears all auth storage on token revocation |
| `extractThreadId()` | `content-outlook.js` | Parses thread ID from the Outlook URL pathname |
| `normalizeThreadId()` | `content-outlook.js` | Validates thread ID length and charset before sending |
| `handleSendMessage()` | `content-outlook.js` | Builds chat payload and dispatches to background |
| `isOutlookConnected()` | `content-outlook.js` | Returns `true` if `userId` exists and `provider === 'outlook'` |

### Message Flow
1. **Popup → Background**: OAuth connection request
2. **Background → n8n** (`/oauth/start`): Get PKCE `state` + `code_challenge`
3. **Background → Microsoft**: OAuth consent screen via `chrome.identity.launchWebAuthFlow()`
4. **Microsoft → n8n** (`/oauth/callback`): Authorization code exchange
5. **n8n → Background**: Redirect with `?jwt=...&userId=...` to `chromiumapp.org` URI
6. **Background → Storage**: Store `jwtToken`, `userId`, `isConnected`, `provider`
7. **Content → Background**: Chat message via `chrome.runtime.sendMessage`
8. **Background → n8n** (`/outlook/cahtbot`): Authenticated POST with JWT
9. **Background → Content**: AI response rendered in sidebar

## Security Considerations

- JWT tokens are stored in `chrome.storage.local` (sandboxed to the extension)
- PKCE prevents authorization code interception attacks
- The OAuth lock (`oauthInProgress`) prevents concurrent auth flows
- Thread IDs are validated before being sent to the backend to prevent malformed requests
- All webhook calls use HTTPS
- Feedback and error endpoints do not require authentication
- `PRODUCTION_MODE = true` silences all console output for Chrome Web Store builds

## Support

For issues or questions:
1. Enable debug mode (`PRODUCTION_MODE = false`) and reproduce the issue
2. Check the background service worker console in `chrome://extensions`
3. Verify n8n webhook status and execution logs at `connector.saai.dev`
4. See `OAUTH_SETUP.md` for Azure configuration details
5. See `OUTLOOK_STATUS_AND_NEXT_STEPS.md` for known issues and the current development checklist

## Version History

### v2.2.3
- Thread ID extraction uses URL pathname to avoid query-string contamination
- Thread ID normalization with length (20–600 chars) and charset validation
- Extension reload detection with user-facing refresh prompt
- Silent disconnect on `invalid_grant` / revoked tokens
- Fallback responses with descriptive messages when n8n is unavailable
- OAuth concurrency lock (in-memory + storage) to prevent duplicate flows
- Token refresh auto-retry on 401/403 with two-attempt logic
- `provider: 'outlook'` flag in storage to distinguish from Gmail extension

### v2.0
- Initial Outlook implementation based on Gmail extension architecture
- Microsoft PKCE OAuth flow via n8n
- Sidebar chat UI adapted for Outlook Web layout
- JWT-based authentication for all n8n requests
