# Outlook OAuth Setup Guide

## Problem
The n8n callback is working and returning JSON, but `chrome.identity.launchWebAuthFlow()` only captures redirects to the extension's redirect URI, not external URLs like `connector.saai.dev`.

## Solution
n8n must redirect to the extension's redirect URI after processing the callback.

## Step 1: Get Extension Redirect URI

1. Open the extension popup
2. Open browser console (F12 → Console tab)
3. You'll see a log message with the extension redirect URI:
   ```
   🔑 EXTENSION REDIRECT URI (add this to Azure): https://<extension-id>.chromiumapp.org/
   ```

**OR** check the background script logs:
- Go to `chrome://extensions/`
- Find "Sa.AI for Outlook"
- Click "service worker" (or "background page")
- Look for: `Extension redirect URI: https://<extension-id>.chromiumapp.org/`

## Step 2: Add to Azure App Registration

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** → **App registrations**
3. Find your app: **SaAI Outlook Integration** (or the app with client ID: `029d6ab9-7fbf-4476-9bf5-deb353bb79a6`)
4. Click **Authentication** in the left menu
5. Under **Platform configurations**, click **Add a platform** → **Web**
6. Add **TWO** redirect URIs:
   - `https://connector.saai.dev/webhook/outlook/oauth/callback` (already there)
   - `https://<extension-id>.chromiumapp.org/` (NEW - use the URI from Step 1)
7. Click **Configure**

## Step 3: Update n8n Callback Workflow

Your n8n callback workflow (`/webhook/outlook/oauth/callback`) currently:
1. Receives `code` and `state` from Microsoft
2. Exchanges code for tokens
3. Generates JWT
4. Returns JSON: `[{userId, jwt}]`

**You need to change step 4** to redirect instead of returning JSON:

### Current (Returns JSON):
```json
[{userId: "...", jwt: "..."}]
```

### New (Redirect to Extension):
After generating the JWT, redirect to the extension's redirect URI with the JWT and userId as URL parameters.

**Example redirect URL format:**
```
https://abcdefghijklmnopqrstuvwxyz123456.chromiumapp.org/?jwt=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9...&userId=AAAAAAAAAAAAAAAAAAAAAAK1gV0Fap0WucJHR7ISxvw
```

### n8n Workflow Update:

**Option 1: Use the extension_redirect_uri from start request (Recommended)**
1. In `/webhook/outlook/oauth/start` workflow:
   - Store the `extension_redirect_uri` query parameter (passed by extension)
   - Save it in your database/workflow state mapped to the `state` value
   
2. In `/webhook/outlook/oauth/callback` workflow:
   - After generating JWT, retrieve the stored `extension_redirect_uri` using the `state` parameter
   - Use a **Redirect** node (or HTTP Response with 302 status)
   - Set redirect URL to: `{{$json.extension_redirect_uri}}?jwt={{$json.jwt}}&userId={{$json.userId}}`
   - Replace `{{$json.jwt}}` and `{{$json.userId}}` with your actual JWT and userId variables

**Option 2: Hardcode extension redirect URI (Not recommended, but simpler)**
1. Get your extension redirect URI from the console (see Step 1)
2. In `/webhook/outlook/oauth/callback` workflow:
   - After generating JWT, add a **Redirect** node
   - Set redirect URL to: `https://YOUR-EXTENSION-ID.chromiumapp.org/?jwt={{$json.jwt}}&userId={{$json.userId}}`
   - Replace `YOUR-EXTENSION-ID` with your actual extension ID
   - Replace `{{$json.jwt}}` with your JWT variable (e.g., `{{$json.jwt}}` or `{{$node["Generate JWT"].json.jwt}}`)
   - Replace `{{$json.userId}}` with your userId variable (e.g., `{{$json.userId}}` or `{{$node["Get User"].json.userId}}`)

**Important:** 
- `<jwt>`, `<userId>`, and `<extension-id>` are **placeholders** - replace them with actual values!
- Use n8n expressions like `{{$json.jwt}}` to get the actual JWT value from your workflow
- Use n8n expressions like `{{$json.userId}}` to get the actual userId value from your workflow

## Step 4: Test

1. Reload the extension
2. Click "Connect to Microsoft & Sa.AI"
3. Complete Microsoft login
4. n8n should redirect to the extension's redirect URI
5. Extension should capture the redirect and store the JWT

## Troubleshooting

### "Redirect URI mismatch" error
- Ensure the extension redirect URI is **exactly** added to Azure (including trailing slash)
- Check that the URI matches what's logged in console

### Callback still not being hit
- Check n8n workflow logs
- Verify redirect is happening (not just returning JSON)
- Check browser network tab for redirect to `chromiumapp.org`

### Extension not capturing redirect
- Verify redirect URI is correct
- Check background script logs for "OAuth redirect URL received"
- Ensure redirect includes `jwt` and `userId` query parameters
