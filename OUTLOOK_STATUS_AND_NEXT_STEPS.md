# Outlook Extension – Status & Where to Pick Up

## Current State (as of last session)

| Item | Status |
|------|--------|
| Outlook sidebar opens | ✅ Working |
| Conversations (send message → AI reply) | ❌ Not working |
| OAuth flow | Depends on n8n redirect setup |
| Chat webhook | `outlook/cahtbot` (Outlook-specific) |

---

## What’s Already Done

1. **File separation** – Gmail (`gmail/`) and Outlook (`outlook/`) are separate folders.
2. **Content script** – `content-outlook.js` injects sidebar, chat UI, handles `handleSendMessage()`.
3. **OAuth** – Microsoft PKCE via n8n; n8n must **redirect** to extension URI (`https://<extension-id>.chromiumapp.org/?jwt=...&userId=...`) instead of returning JSON.
4. **Background** – `background.js` handles OAuth, token refresh, and forwards chat to `https://connector.saai.dev/webhook/outlook/cahtbot` with JWT.
5. **Connection check** – `isOutlookConnected()` returns true if `userId` (or `isConnected`) and `provider === 'outlook'` (or unset).
6. **Thread ID extraction** – `extractThreadId()` reads the ID from the URL path (e.g. `/mail/0/inbox/id/XXX`). It uses **pathname** so query params (e.g. `?path=%2Fmail&wa=wsignin1.0`) never get mixed into the ID. Different Outlook account types use different ID formats (e.g. **AAkALg...** vs **AQQk...**); the extension accepts any base64-like segment. The **backend** (n8n / Microsoft Graph) must support both formats when summarizing or fetching a thread.
7. **Thread ID safeguards** – Before sending to the backend, `normalizeThreadId()` validates length (20–600 chars) and allowed charset (base64-like). Invalid IDs are not sent; summarization is blocked with a clear message (“We couldn’t read the message ID…”) when the URL looks like a thread but the ID fails validation. This avoids backend errors from malformed or future-format IDs.

---

## Chat Flow (when you send a message)

1. **Content** (`content-outlook.js`) → `handleSendMessage()`  
   - Checks `userId`, credits, then builds payload: `{ query, userId, context: 'OutlookChat', ... }`
2. **Content** → `chrome.runtime.sendMessage({ action: 'sendToN8N', data: { endpoint: 'chat', payload } })`
3. **Background** (`background.js`) → `handleN8NRequest()`  
   - Gets JWT with `ensureValidJWTToken()` → if none: **"No JWT token found. Please authenticate first."**
   - POSTs to `outlook/cahtbot` with `Authorization: Bearer <jwt>`
4. **Content** receives response  
   - If `response.success` → parses and shows reply  
   - If `response.success === false` → shows `Error: <response.error>`

---

## Why “no conversations” could happen

| Possible cause | How to check |
|----------------|--------------|
| **Not connected** | Sidebar shows “Connect Outlook Account” instead of chat UI. If you see chat UI, you’re connected. |
| **No JWT** | Background throws “No JWT token found. Please authenticate first.” You should see this in the chat as an error message. |
| **OAuth failed / not completed** | No `userId` or `jwtToken` in storage. Use the “Debug” button in the Connect prompt or run the debug check below. |
| **Outlook chatbot webhook rejects** | Backend may not accept `context: 'OutlookChat'` or JWT. Check n8n logs. |
| **Receiving end doesn't exist** | Extension was reloaded; refresh the Outlook page (F5) to re-inject the script. |
| **"Service temporarily unavailable" / "Session expired"** | JWT expired and **token refresh failed**. Reconnect Outlook (click Connect) for a fresh session. Fix long-term: ensure n8n `outlook/session/renew` webhook works and returns a new JWT. |
| **Request timeout** | 90s timeout. If it times out, you see “Error: Request timed out”. |

---

## Logic: Extension reload and “receiving end doesn’t exist”

When you **reload the extension** (e.g. after changing code or via chrome://extensions → Reload), the **background** restarts but **already-open tabs** still run the **old** content script (orphaned). That script’s `chrome.runtime.sendMessage()` then fails with **“Receiving end does not exist”**.

**Outlook now:** When this is detected (on send or Connect), the UI shows: *“The extension was reloaded or updated. Please refresh this page (F5 or reload) to connect again.”*

**What to do:** **Refresh the Outlook tab (F5).** The new background injects the content script and messaging works. Login stays in storage, so you usually don’t need to log in again.

**Gmail vs Outlook:** Same idea—after extension reload, refresh the mail tab to re-inject the script. Gmail doesn’t clear storage on reload; Outlook now follows the same pattern (refresh tab → script re-injects; re-login only if storage was cleared).

---

## Token refresh (renew) paused

Outlook renew URL: **outlook/session/renew-outlook**. Automatic token refresh is **disabled** until you set `outlookTokenRefreshEnabled: true` in `outlook/background.js`. When the backend returns 401/403, the extension will show “Session expired. Please reconnect your Outlook account” and will **not** call the renew endpoint.

To turn refresh back on after renew is ready: in **`outlook/background.js`**, set `outlookTokenRefreshEnabled: true` in the `FEATURES` object (near the top).

---

## Next Steps – Debug Checklist

### 1. Turn on debug mode (temporary)

In **`outlook/content-outlook.js`** line ~6:

```js
const PRODUCTION_MODE = false;  // was: true
```

In **`outlook/background.js`** it’s already `false` (line 5).

### 2. Check connection status

1. Open Outlook web.
2. Open the Sa.AI sidebar.
3. If you see “Connect Outlook Account”:
   - Click Connect and finish OAuth.
   - Confirm n8n redirects to `https://<extension-id>.chromiumapp.org/?jwt=...&userId=...`.
4. If you see the chat UI, you’re connected.

### 3. Inspect storage

1. Open DevTools (F12) → Console on the Outlook page.
2. Run:

```js
chrome.storage.local.get(['userId', 'jwtToken', 'isConnected', 'provider'], console.log)
```

Expected when connected: `{ userId: "...", jwtToken: "...", isConnected: true, provider: "outlook" }`.

### 4. Test a message and watch logs

1. Keep DevTools open (Console tab).
2. Type a simple message (e.g. “hello”) and send.
3. Look for:
   - Content: `[SaAI-Content]` logs (if `PRODUCTION_MODE = false`).
   - Background: open `chrome://extensions` → Sa.AI for Outlook → “service worker” → Inspect.
   - Errors like “No JWT token found” or “Request timed out”.

### 5. Check Outlook chatbot webhook

- Confirm `https://connector.saai.dev/webhook/outlook/cahtbot` handles `context: 'OutlookChat'` (or falls back like Gmail).
- Inspect n8n execution logs for any 4xx/5xx or failed runs.

---

## Files to Touch for Debugging

| File | Purpose |
|------|---------|
| `outlook/content-outlook.js` | Set `PRODUCTION_MODE = false` to see content logs |
| `outlook/background.js` | Debug mode already on; check service worker console |
| `outlook/popup.js` | Connect flow, status display |

---

## After Fixing

1. Set `PRODUCTION_MODE = true` in `content-outlook.js` again.
2. Verify chat works end-to-end (send → reply).
3. Optional: fix the “Connecting to Google…” text in `showOAuthLoader()` to “Connecting to Microsoft…” (search for `showOAuthLoader`).
