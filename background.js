// Background script for Sa.AI Inbox Assistant

// Production-safe logging - ALL logging disabled for Chrome Web Store compliance
// For local debugging only: Set PRODUCTION_MODE = false in all 3 files
const PRODUCTION_MODE = false; // TEMPORARILY ENABLED FOR DEBUGGING

// No-op functions for production - completely silent
function debugLog() {}
function debugError() {}
function debugWarn() {}

// Robust network request wrapper with timeout and error handling
async function safeRequest(url, options = {}, timeout = 90000) {
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), timeout);
  
  try {
    const response = await fetch(url, {
      ...options,
      signal: controller.signal
    });
    
    clearTimeout(timeoutId);
    
    if (!response.ok) {
      // Allow caller to handle non-2xx responses when explicitly requested
      if (options && options._allowErrorResponse === true) {
        return response;
      }
      // Try to get error details from response body
      let errorDetails = response.statusText;
      try {
        const errorBody = await response.text();
        if (errorBody) errorDetails = errorBody;
      } catch (e) {
        // Ignore if can't read body
      }
      throw new Error(`HTTP ${response.status}: ${errorDetails}`);
    }
    
    return response;
  } catch (error) {
    clearTimeout(timeoutId);
    
    if (error.name === 'AbortError') {
      throw new Error('Request timed out');
    }
    
    debugError('Network request failed:', error);
    throw error;
  }
}

// Feature flags
const FEATURES = {
    heartbeatEnabled: false,
    // Set to true when n8n webhook outlook/session/renew-outlook is ready; until then, 401/403 will ask user to reconnect
    outlookTokenRefreshEnabled: true
};

chrome.runtime.onInstalled.addListener(async (details) => {
    debugLog('Sa.AI Inbox Assistant installed/updated', details.reason);
    
    // Clear any stale OAuth lock
    const stored = await chrome.storage.local.get(['oauthInProgress']);
    if (stored.oauthInProgress) {
        await chrome.storage.local.remove('oauthInProgress');
        debugLog('Cleared stale OAuth lock');
    }
    
    // Only clear storage on fresh install, not on updates
    if (details.reason === 'install') {
    chrome.storage.local.clear(() => {
            debugLog('Storage cleared on fresh installation');
        });
    } else if (details.reason === 'update') {
        debugLog('Extension updated - preserving existing user data');
        // Do not clear storage on updates
    }
});

// Check existing session on startup
chrome.runtime.onStartup.addListener(async () => {
    debugLog('Extension startup - checking existing session');
    
    // Clear any stale OAuth lock from previous session
    const stored = await chrome.storage.local.get(['oauthInProgress']);
    if (stored.oauthInProgress) {
        await chrome.storage.local.remove('oauthInProgress');
        debugLog('Cleared stale OAuth lock from previous session');
    }
    
    chrome.storage.local.get(['isConnected', 'jwtToken', 'userId'], (result) => {
        if (result.isConnected && result.jwtToken && result.userId) {
            debugLog('Existing session found');
        } else {
            debugLog('No existing session found');
        }
    });
});

// Periodic token refresh disabled - only refresh on 401/403 errors from n8n
function setupPeriodicTokenRefresh() {
    debugLog('Periodic token refresh disabled - will refresh only on auth errors');
    
    // No automatic refresh alarms
    // Token will be refreshed only when n8n returns 401/403
}

// heartbeat function to maintain active connection (disabled by default)
async function performHeartbeat() {
    if (!FEATURES.heartbeatEnabled) {
        debugLog('Heartbeat skipped (disabled)');
        return;
    }

    try {
        debugLog('Performing connection heartbeat');
        
        const token = await getJWTToken();
        const userId = await getStoredUserId();
        
        if (!token || !userId) {
            debugLog('Heartbeat: No token or userId, skipping');
            return;
        }
        
        // Call n8n heartbeat endpoint to maintain session
        const heartbeatResponse = await safeRequest('https://connector.saai.dev/webhook/oauth/heartbeat', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${token}`
            },
            body: JSON.stringify({
                userId: userId,
                action: 'heartbeat',
                timestamp: Date.now(),
                source: 'chrome_extension'
            })
        });
        
        if (heartbeatResponse.ok) {
            const heartbeatData = await heartbeatResponse.json();
            debugLog('Heartbeat successful:', heartbeatData.success);
            
            // Update last activity
            await chrome.storage.local.set({
                lastHeartbeat: Date.now(),
                sessionActive: true
            });
            
            // If heartbeat returns a new refresh token, update it
            if (heartbeatData.refreshToken) {
                debugLog('Heartbeat provided new refresh token');
                await chrome.storage.local.set({
                    refreshToken: heartbeatData.refreshToken
                });
            }
        }
        
    } catch (error) {
        debugLog('Heartbeat failed:', error.message);
        
        // Mark session as potentially inactive
        await chrome.storage.local.set({
            sessionActive: false,
            lastHeartbeatFailed: Date.now()
        });
    }
}

// Alarm handler disabled - no periodic refresh
chrome.alarms.onAlarm.addListener(async (alarm) => {
    debugLog('Alarm triggered but periodic refresh is disabled:', alarm.name);
    // No automatic token refresh - only refresh on 401/403 errors
});

// Handle message forwarding for OAuth and chat
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
    // PRODUCTION: console.log('[SaAI-BG] Message received - action:', request.action);
    debugLog('Message received - action:', request.action);
    
    if (request.action === 'sendToN8N') {
        // PRODUCTION: console.log('[SaAI-BG] Processing sendToN8N request - endpoint:', request.data?.endpoint);
        debugLog('Processing sendToN8N request - endpoint:', request.data?.endpoint);
        
        if (request.data.endpoint === 'oauth') {
            // PRODUCTION: console.log('[SaAI-BG] OAuth endpoint detected - calling handleOAuthFlow()');
            debugLog('OAuth endpoint detected - calling handleOAuthFlow()');
        }
        
        handleN8NRequest(request.data)
            .then(response => {
                // PRODUCTION: console.log('[SaAI-BG] sendToN8N success response:', response);
                debugLog('sendToN8N success response:', response);
                sendResponse({success: true, data: response});
            })
            .catch(error => {
                // PRODUCTION: console.error('[SaAI-BG] sendToN8N error:', error);
                debugError('sendToN8N error:', error);
                sendResponse({success: false, error: error.message});
            });
        return true;
    } else if (request.action === 'trackCredits') {
        // PRODUCTION: console.log('[SaAI-BG] Processing trackCredits request:', request.data);
        debugLog('Processing trackCredits request:', request.data);
        
        handleCreditTracking(request.data)
            .then(response => {
                // PRODUCTION: console.log('[SaAI-BG] trackCredits success response:', response);
                debugLog('trackCredits success response:', response);
                sendResponse({success: true, data: response});
            })
            .catch(error => {
                // PRODUCTION: console.error('[SaAI-BG] trackCredits error:', error);
                debugError('trackCredits error:', error);
                sendResponse({success: false, error: error.message});
            });
        return true;
    } else if (request.action === 'refreshToken') {
        // PRODUCTION: console.log('[SaAI-BG] Processing refreshToken request');
        debugLog('Processing refreshToken request');
        refreshJWTToken()
            .then(token => {
                // PRODUCTION: console.log('[SaAI-BG] refreshToken success:', token ? 'Token received' : 'No token');
                debugLog('refreshToken success:', token ? 'Token received' : 'No token');
                sendResponse({success: true, token: token});
            })
            .catch(error => {
                // PRODUCTION: console.error('[SaAI-BG] refreshToken error:', error);
                debugError('refreshToken error:', error);
                sendResponse({success: false, error: error.message});
            });
        return true;
    }
    
    // PRODUCTION: console.log('[SaAI-BG] Unknown message action:', request.action);
    debugLog('Unknown message action:', request.action);
    sendResponse({success: false, error: 'Unknown action'});
});

// Handle tab updates to inject content script only if needed
chrome.tabs.onUpdated.addListener((tabId, changeInfo, tab) => {
    if (changeInfo.status === 'complete' && tab.url && (
        tab.url.includes('outlook.office.com') || 
        tab.url.includes('outlook.live.com') || 
        tab.url.includes('outlook.com')
    )) {
        // Ping first - only inject if content script is missing
        chrome.tabs.sendMessage(tabId, {action: 'checkInitialization'}).then((response) => {
            // Content script responded - already loaded, no injection needed
            debugLog('Content script already loaded on tab', tabId);
        }).catch((error) => {
            // Content script not responding - safe to inject
            debugLog('Content script missing, injecting on tab', tabId);
            chrome.scripting.executeScript({
                target: {tabId: tabId},
                files: ['content-outlook.js']
            }).catch((injectError) => {
                debugError('Failed to inject content script:', injectError);
            });
        });
    }
});

// Get JWT token from storage
async function getJWTToken() {
    const { jwtToken } = await chrome.storage.local.get(['jwtToken']);
    return jwtToken;
}

// Get stored user ID from storage
async function getStoredUserId() {
    const { userId } = await chrome.storage.local.get(['userId']);
    return userId;
}

// Get refresh token from storage
async function getRefreshToken() {
    const { refreshToken } = await chrome.storage.local.get(['refreshToken']);
    return refreshToken;
}

// Get token refresh count for tracking
async function getTokenRefreshCount() {
    const { tokenRefreshCount } = await chrome.storage.local.get(['tokenRefreshCount']);
    return tokenRefreshCount || 0;
}

// Check if JWT token is expired (parse JWT payload for exp claim)
async function isJWTTokenExpired(jwtToken) {
    try {
        if (!jwtToken) return true;
        
        // Decode JWT payload (base64url decode)
        const parts = jwtToken.split('.');
        if (parts.length !== 3) return true;
        
        const payload = JSON.parse(atob(parts[1].replace(/-/g, '+').replace(/_/g, '/')));
        const exp = payload.exp;
        
        if (!exp) return false; // No expiration claim, assume valid
        
        // Check if token expires within next 5 minutes (buffer for refresh)
        const currentTime = Math.floor(Date.now() / 1000);
        const bufferTime = 5 * 60; // 5 minutes
        
        return (exp - currentTime) < bufferTime;
    } catch (error) {
        debugError('Error checking JWT expiration:', error);
        return true; // If we can't parse it, assume expired
    }
}

// Get JWT token without proactive refresh - only check if it exists
async function ensureValidJWTToken() {
    const jwtToken = await getJWTToken();
    
    if (!jwtToken) {
        throw new Error('No JWT token found. Please authenticate first.');
    }
    
    // Don't check expiry or refresh proactively
    // Let n8n return 401 if token is expired, then we refresh
    return jwtToken;
}

// Refresh JWT token using n8n refresh endpoint
async function refreshJWTToken() {
    try {
        debugLog('JWT token expired, attempting refresh via n8n');
        
        const currentUserId = await getStoredUserId();
        if (!currentUserId) {
            throw new Error('No userId found. Please re-authenticate.');
        }
        
        const refreshResponse = await safeRequest('https://connector.saai.dev/webhook/outlook/session/renew-outlook', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${await getJWTToken()}`
            },
            body: JSON.stringify({
                userId: currentUserId,
                refreshToken: await getRefreshToken(),
                grantType: 'refresh_token',
                extendLifetime: true
            })
        });
        
        if (!refreshResponse.ok) {
            const errorText = await refreshResponse.text();
            debugError('Token refresh failed:', errorText);
            
            // Handle Outlook-specific error codes
            if (refreshResponse.status === 401) {
                debugError('Token invalid - forcing reconnection');
                await silentlyDisconnectUser();
                throw new Error('Token invalid. Please reconnect your Outlook account.');
            } else if (refreshResponse.status === 404) {
                debugError('User not found');
                throw new Error('User not found. Please reconnect your Outlook account.');
            }
            
            throw new Error(`Token refresh failed: ${refreshResponse.status} ${refreshResponse.statusText}`);
        }
        
        let refreshData = await refreshResponse.json();
        
        // Handle array response from n8n (extract first element)
        if (Array.isArray(refreshData) && refreshData.length > 0) {
            debugLog('Refresh response is array, extracting first element');
            refreshData = refreshData[0];
        }
        
        // Check for JWT token in various field names
        const newJwtToken = refreshData.jwt || refreshData.jwtToken || refreshData.token;
        
        if (!newJwtToken) {
            debugError('No JWT in refresh response:', refreshData);
            throw new Error('Token refresh response missing JWT token');
        }
        
        debugLog('JWT token refreshed successfully');
        
        // Store the new token (preserve provider flag)
        const currentStorage = await chrome.storage.local.get(['provider']);
        await chrome.storage.local.set({
            jwtToken: newJwtToken,
            refreshToken: refreshData.refreshToken || await getRefreshToken(),
            userId: refreshData.userId || currentUserId,
            provider: currentStorage.provider || 'outlook', // Preserve provider flag
            tokenIssuedAt: Date.now(),
            tokenRefreshCount: (await getTokenRefreshCount()) + 1,
            lastSuccessfulRefresh: Date.now()
        });
        
        return newJwtToken;
        
    } catch (error) {
        debugError('Token refresh failed:', error);
        // PRODUCTION: console.error('[SaAI-BG] Token refresh error - check if webhook/session/renew is active and returning jwt field');
        throw new Error('Unable to refresh authentication. Please re-authenticate.');
    }
}

// Handle credit tracking webhook requests
async function handleCreditTracking(data) {
    const { userId, prompt, creditsUsed } = data;
    
    // PRODUCTION: console.log('[SaAI-BG] Credit tracking - creditsUsed:', creditsUsed);
    debugLog('Sending credit tracking to n8n - creditsUsed:', creditsUsed);
    
    // Validate data
    if (!userId || !prompt || !creditsUsed) {
        // PRODUCTION: console.error('[SaAI-BG] Invalid credit tracking data:', data);
        throw new Error('Missing required credit tracking parameters');
    }
    
    // Get and validate JWT token
    const jwtToken = await ensureValidJWTToken();
    if (!jwtToken) {
        throw new Error('No JWT token found. Please authenticate first.');
    }
    
    const url = 'https://connector.saai.dev/webhook/Credit-Tracking';
    
    const payload = {
        userId: userId,
        prompt: prompt,
        creditsUsed: creditsUsed
    };
    
    // PRODUCTION: console.log('[SaAI-BG] Sending credit tracking - credits:', creditsUsed);
    // PRODUCTION: console.log('[SaAI-BG] JWT token:', jwtToken ? 'Present' : 'Missing');
    
    const response = await safeRequest(url, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            'Accept': 'application/json',
            'Authorization': `Bearer ${jwtToken}`
        },
        body: JSON.stringify(payload)
    });
    
    // PRODUCTION: console.log('[SaAI-BG] Credit tracking response status:', response.status);
    
    if (!response.ok) {
        const errorText = await response.text();
        // PRODUCTION: console.error('[SaAI-BG] Credit tracking error response:', errorText);
        throw new Error(`Credit tracking failed: ${response.status} ${response.statusText}`);
    }
    
    const result = await response.json();
    // PRODUCTION: console.log('[SaAI-BG] Credit tracking response data:', result);
    debugLog('Credit tracking response:', result);
    
    return result;
}

// Handle N8N webhook requests
async function handleN8NRequest(data) {
    const { endpoint, payload } = data;
    
    if (endpoint === 'oauth') {
        // For OAuth, we need to launch the Microsoft Outlook OAuth flow first
        return await handleOAuthFlow();
    }
    
    // Determine URL first
    let url;
    switch (endpoint) {
        case 'chat':
            // Check if this is a thread summarization request
            if (payload.action === 'summarize_thread' && payload.threadId) {
                // PRODUCTION: console.log('[Background] Thread summarization request detected:', {
                //     threadId: payload.threadId,
                //     subjectLine: payload.subjectLine || 'Not provided',
                //     query: payload.query
                // });
                // Use the same webhook but with thread context
                url = 'https://connector.saai.dev/webhook/outlook/cahtbot';
            } else {
                url = 'https://connector.saai.dev/webhook/outlook/cahtbot';
            }
            break;
        case 'task':
            url = 'https://connector.saai.dev/webhook/task-management-outlook';
            break;
        case 'feedback':
            url = 'https://connector.saai.dev/webhook/Feedback_error';
            break;
        case 'dataDeletion':
            url = 'https://connector.saai.dev/webhook/Feedback_error';
            break;
        default:
            throw new Error('Invalid endpoint');
    }
    
    // PRODUCTION: console.log('[Background] Sending request to n8n - endpoint:', endpoint, 'url:', url);
    
    // For feedback and dataDeletion, no authentication required - send directly
    if (endpoint === 'feedback' || endpoint === 'dataDeletion') {
        try {
            const response = await safeRequest(url, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'Accept': 'application/json'
                },
                body: JSON.stringify(payload),
                _allowErrorResponse: true
            });
            
            // PRODUCTION: console.log(`[Background] ${endpoint} response status:`, response.status);
            // PRODUCTION: console.log(`[Background] ${endpoint} response URL:`, response.url);
            
            if (response.ok) {
                const result = await response.json();
                // PRODUCTION: console.log(`[Background] ${endpoint} sent successfully:`, result);
                return result;
            } else {
                const errorText = await response.text();
                // PRODUCTION: console.error(`[Background] ${endpoint} error response:`, {
                //     status: response.status,
                //     statusText: response.statusText,
                //     url: response.url,
                //     body: errorText
                // });
                throw new Error(`${endpoint} submission failed: ${response.status} - ${errorText || response.statusText}`);
            }
        } catch (error) {
            // PRODUCTION: console.error(`[Background] ${endpoint} request failed:`, error);
            throw error;
        }
    }
    
    // For other endpoints, get and validate JWT token for authenticated requests (auto-refresh if needed)
    const jwtToken = await ensureValidJWTToken();
    if (!jwtToken) {
        throw new Error('No JWT token found. Please authenticate first.');
    }
    
    // PRODUCTION: console.log('[Background] Sending authenticated request to n8n - url:', url);
    
    try {
    const response = await safeRequest(url, {
        method: 'POST',
        headers: {
                'Content-Type': 'application/json',
                'Accept': 'application/json',
                'Authorization': `Bearer ${jwtToken}`
        },
        body: JSON.stringify(payload),
        _allowErrorResponse: true
    });
        
        // PRODUCTION: console.log('[Background] n8n response status:', response.status);
        // PRODUCTION: console.log('[Background] n8n response headers:', response.headers);
    
    if (!response.ok) {
            // Handle specific error cases
            if (response.status === 401 || response.status === 402 || response.status === 403) {
                // Token refresh paused until n8n outlook/session/renew is set up
                if (!FEATURES.outlookTokenRefreshEnabled) {
                    debugLog('401/403 received - token refresh disabled, asking user to reconnect');
                    throw new Error('Session expired. Please reconnect your Outlook account (click Connect in the sidebar) and try again.');
                }
                
                try {
                    // Automatically refresh the token
                    const newJwtToken = await refreshJWTToken();
                    debugLog('Token refreshed successfully, retrying original request');
                    
                    // Retry the original request with new token
                    const retryResponse = await safeRequest(url, {
                        method: 'POST',
                        headers: {
                            'Content-Type': 'application/json',
                            'Accept': 'application/json',
                            'Authorization': `Bearer ${newJwtToken}`
                        },
                        body: JSON.stringify(payload)
                    });
                    
                    if (retryResponse.ok) {
                        const retryResult = await retryResponse.json();
                        debugLog('Request retry successful after token refresh');
                        return retryResult;
                    } else {
                        // If retry still fails, try one more time with fresh token
                        debugLog('First retry failed, attempting second retry');
                        const finalJwtToken = await ensureValidJWTToken();
                        const finalRetryResponse = await safeRequest(url, {
                            method: 'POST',
                            headers: {
                                'Content-Type': 'application/json',
                                'Accept': 'application/json',
                                'Authorization': `Bearer ${finalJwtToken}`
                            },
                            body: JSON.stringify(payload)
                        });
                        
                        if (finalRetryResponse.ok) {
                            const finalResult = await finalRetryResponse.json();
                            debugLog('Second retry successful');
                            return finalResult;
                        } else {
                            throw new Error(`Request failed after token refresh attempts: ${finalRetryResponse.status}`);
                        }
                    }
                } catch (refreshError) {
                    debugError('Automatic token refresh failed:', refreshError);
                    // Tell user to reconnect so they get a fresh session
                    throw new Error('Session expired. Please reconnect your Outlook account (click Connect in the sidebar) and try again.');
                }
            } else if (response.status === 400) {
                // Check for token expiration/revocation (invalid_grant)
                const errorText = await response.text().catch(() => 'Unknown error');
                let errorData;
                try {
                    errorData = JSON.parse(errorText);
                } catch (e) {
                    errorData = { error: errorText };
                }
                
                // Check if this is a token expiration error
                if (errorData.error === 'invalid_grant' || 
                    (errorData.error_description && errorData.error_description.includes('Token has been expired or revoked'))) {
                    debugLog('Token expired or revoked - silently disconnecting user');
                    await silentlyDisconnectUser();
                    // Return a special response indicating disconnection (not an error)
                    return { 
                        disconnected: true,
                        message: 'Session expired. Please reconnect.'
                    };
                }
                
                // Other 400 errors - throw normally
                throw new Error(`n8n webhook error (400): ${errorText}`);
            } else if (response.status === 404) {
                // PRODUCTION: console.error('[Background] n8n webhook not found (404)');
                // Return a fallback response instead of throwing error
                return await handleFallbackResponse(endpoint, payload);
            } else if (response.status === 500) {
                // PRODUCTION: console.error('[Background] n8n 500 error for endpoint:', endpoint, 'URL:', url);
                throw new Error(`n8n server error (500) - webhook ${endpoint} crashed. Check n8n execution logs.`);
            } else {
                const errorText = await response.text().catch(() => 'Unknown error');
                throw new Error(`n8n webhook error (${response.status}): ${errorText}`);
            }
        }
        
        // Try to parse JSON response
        const responseText = await response.text();
        // PRODUCTION: console.log('[Background] n8n raw response:', responseText);
        
        let result;
        try {
            result = responseText ? JSON.parse(responseText) : {};
        } catch (parseError) {
            // PRODUCTION: console.warn('[Background] Failed to parse JSON response:', parseError);
            // If response is not JSON, treat it as text
            result = { message: responseText || 'Response received but not in expected format' };
        }
        
        // Check if the response body contains token expiration error
        if (result.error === 'invalid_grant' || 
            (result.error_description && result.error_description.includes('Token has been expired or revoked'))) {
            debugLog('Token expired or revoked in response body - silently disconnecting user');
            await silentlyDisconnectUser();
            // Return a special response indicating disconnection (not an error)
            return { 
                disconnected: true,
                message: 'Session expired. Please reconnect.'
            };
        }
        
        // PRODUCTION: console.log('[Background] n8n parsed response:', result);
        return result;
        
    } catch (error) {
        // PRODUCTION: console.error('[Background] Network error:', error);
        
        // If it's a network error (not a 404), try fallback
        if (error.name === 'TypeError' || error.message.includes('fetch')) {
            // PRODUCTION: console.log('[Background] Network error detected, using fallback response');
            return await handleFallbackResponse(endpoint, payload);
        }
        
        throw error;
    }
}

// Handle fallback responses when n8n is unavailable
async function handleFallbackResponse(endpoint, payload) {
    // PRODUCTION: console.log('[Background] Using fallback response for:', endpoint);
    
    if (endpoint === 'chat') {
        const userMessage = payload.query || payload.message || 'Hello';
        
        // Check if this is a thread summarization request
        if (payload.action === 'summarize_thread' && payload.threadId) {
            // PRODUCTION: console.log('[Background] Thread summarization fallback for threadId:', payload.threadId);
            const subjectInfo = payload.subjectLine ? ` (Subject: "${payload.subjectLine}")` : '';
            return {
                message: `I can see you want me to summarize the email thread (ID: ${payload.threadId})${subjectInfo}. However, the n8n webhook is currently unavailable, so I cannot access the thread content to provide a summary. Please check your webhook configuration and try again.`,
                fallback: true,
                webhookStatus: 'unavailable',
                suggestion: 'Please verify your n8n webhook is deployed and accessible',
                threadId: payload.threadId,
                subjectLine: payload.subjectLine,
                action: 'summarize_thread'
            };
        }
        
        // Generate a mock response based on the user's message
        const mockResponses = {
            'hello': 'Hi there! I\'m your Inbox assistant. How can I help you today?',
            'help': 'I can help you with:\n• Summarizing your inbox\n• Finding important emails\n• Managing your tasks\n• Answering questions about your emails\n• Summarizing specific email threads (open the thread first)',
            'summarize': 'I\'d be happy to summarize your inbox! However, the n8n webhook is currently unavailable. Please check your webhook configuration.',
            'inbox': 'I can help you with your inbox! The n8n integration needs to be configured to access your Outlook data.',
            'email': 'I\'m here to help with your emails! Please ensure the n8n webhook is properly set up to process your requests.',
            'thread': 'To summarize a specific email thread, please open the thread first, then ask me to summarize it.'
        };
        
        // Find the best matching response
        const lowerMessage = userMessage.toLowerCase();
        let response = 'I understand you\'re asking about "' + userMessage + '". The n8n webhook is currently unavailable. Please check your webhook configuration at: https://connector.saai.dev/webhook/outlook/cahtbot';
        
        for (const [key, value] of Object.entries(mockResponses)) {
            if (lowerMessage.includes(key)) {
                response = value;
                break;
            }
        }
        
        return {
            message: response,
            fallback: true,
            webhookStatus: 'unavailable',
            suggestion: 'Please verify your n8n webhook is deployed and accessible'
        };
    }
    
    return {
        message: 'Service temporarily unavailable. Please check your n8n webhook configuration.',
        fallback: true,
        webhookStatus: 'unavailable'
    };
}

// OAuth flow lock to prevent concurrent flows
let oauthInProgress = false;

// Silently disconnect user when token is expired/revoked
async function silentlyDisconnectUser() {
    debugLog('Silently disconnecting user - clearing all connection data');
    try {
        await chrome.storage.local.remove([
            'isConnected',
            'userId',
            'jwtToken',
            'refreshToken',
            'oauthData'
        ]);
        debugLog('User disconnected - all connection data cleared');
    } catch (error) {
        debugError('Error during silent disconnection:', error);
    }
}

// PKCE OAuth flow with n8n webhook
async function handleOAuthFlow() {
    // PRODUCTION: console.log('[SaAI-BG] ===== handleOAuthFlow() called =====');
    
    // Check both in-memory and persisted lock state
    const stored = await chrome.storage.local.get(['oauthInProgress']);
    if (oauthInProgress || stored.oauthInProgress) {
        debugLog('OAuth flow already in progress, rejecting duplicate request');
        throw new Error('OAuth flow already in progress. Please wait for the current authentication to complete.');
    }
    
    oauthInProgress = true;
    // Persist lock state to survive service worker restarts
    await chrome.storage.local.set({ oauthInProgress: true });
    
    try {
        // PRODUCTION: console.log('[SaAI-BG] Starting PKCE OAuth flow with n8n');
        debugLog('Starting PKCE OAuth flow with n8n');
        
        // Step 1: Get extension's redirect URI first (needed for n8n to redirect back)
        const extensionRedirectUri = chrome.identity.getRedirectURL();
        debugLog('Extension redirect URI:', extensionRedirectUri);
        
        // Step 1: Call n8n /outlook/oauth/start endpoint to get PKCE parameters
        // Pass extension redirect URI so n8n knows where to redirect after processing callback
        debugLog('Calling n8n /outlook/oauth/start endpoint...');
        
        let startData;
        try {
            // Pass extension redirect URI to n8n so it can redirect back to extension
            const startUrl = `https://connector.saai.dev/webhook/outlook/oauth/start?extension_redirect_uri=${encodeURIComponent(extensionRedirectUri)}`;
            const startResponse = await safeRequest(startUrl, {
                method: 'GET'
            });
            
            debugLog('PKCE start response status:', startResponse.status);
            debugLog('PKCE start response headers:', Object.fromEntries(startResponse.headers.entries()));
            
            if (!startResponse.ok) {
                throw new Error(`n8n /oauth/start failed with status ${startResponse.status}: ${startResponse.statusText}`);
            }
            
            startData = await startResponse.json();
            debugLog('PKCE start response data:', startData);
            
            if (!startData.state || !startData.code_challenge) {
                debugError('Invalid PKCE parameters from n8n:', startData);
                throw new Error(`Invalid PKCE parameters from n8n. Response: ${JSON.stringify(startData)}`);
            }
            
            debugLog('PKCE parameters received successfully:', {
                state: startData.state,
                code_challenge: startData.code_challenge,
                code_challenge_method: startData.code_challenge_method
            });
            
        } catch (startError) {
            debugError('Failed to get PKCE parameters from /oauth/start:', startError);
            // PRODUCTION: console.error('[SaAI-BG] OAuth start error:', startError.message);
            throw new Error(`Failed to get OAuth parameters: ${startError.message}`);
        }
        
        // Step 2: Build Microsoft OAuth URL with n8n's PKCE parameters
        // Azure App Registration Client ID
        const clientId = '029d6ab9-7fbf-4476-9bf5-deb353bb79a6';
        
        // CRITICAL: chrome.identity.launchWebAuthFlow() only captures redirects to extension's redirect URI
        // We need to use the extension's redirect URI, not the n8n callback URL directly
        // n8n callback URL must be registered in Azure, but Microsoft should redirect to extension URI
        // OR: n8n processes callback and redirects to extension URI
        
        // For Microsoft OAuth, we use n8n callback URL (must be registered in Azure)
        // n8n will then redirect to extension's redirect URI with the data
        const n8nCallbackUri = 'https://connector.saai.dev/webhook/outlook/oauth/callback';
        
        const scopes = [
            'openid',
            'profile',
            'email',
            'offline_access',
            'User.Read',
            'Mail.Read'
        ];
        
        // Build Microsoft OAuth URL with n8n callback as redirect_uri
        // n8n must be configured to redirect to extensionRedirectUri after processing
        const authUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?` +
            `client_id=${clientId}` +
            `&response_type=code` +
            `&redirect_uri=${encodeURIComponent(n8nCallbackUri)}` +
            `&response_mode=query` +
            `&scope=${encodeURIComponent(scopes.join(' '))}` +
            `&state=${encodeURIComponent(startData.state)}` +
            `&code_challenge=${encodeURIComponent(startData.code_challenge)}` +
            `&code_challenge_method=${startData.code_challenge_method || 'S256'}` +
            `&prompt=consent`;
        
        debugLog('Launching Outlook OAuth flow with PKCE');
        debugLog('n8n callback URI (registered in Azure):', n8nCallbackUri);
        debugLog('Extension redirect URI (where n8n should redirect):', extensionRedirectUri);
        debugLog('Using n8n state:', startData.state);
        debugLog('Microsoft OAuth URL:', authUrl);
        
        // Step 3: Launch OAuth flow - await the Promise so finally block executes after completion
        // Add timeout to prevent hanging
        const oauthTimeout = setTimeout(() => {
            debugError('OAuth flow timed out after 5 minutes');
            reject(new Error('OAuth flow timed out. Please try again.'));
        }, 5 * 60 * 1000); // 5 minute timeout
        
        const result = await new Promise((resolve, reject) => {
            debugLog('Launching Chrome identity web auth flow...');
            debugLog('Waiting for redirect to extension URI:', extensionRedirectUri);
        chrome.identity.launchWebAuthFlow({
            url: authUrl,
            interactive: true
            }, async function(redirectUrl) {
                clearTimeout(oauthTimeout); // Clear timeout on response
            if (chrome.runtime.lastError) {
                    debugError('OAuth failed:', chrome.runtime.lastError);
                    
                    // Handle user cancellation specifically
                    const errorMessage = chrome.runtime.lastError.message || '';
                    if (errorMessage.includes('user did not approve') || errorMessage.includes('cancelled')) {
                        reject(new Error('OAUTH_CANCELLED: User cancelled the authentication process'));
                    } else {
                        reject(new Error(`OAuth failed: ${errorMessage}`));
                    }
                return;
            }
            
            if (!redirectUrl) {
                debugError('No redirect URL received from OAuth flow');
                reject(new Error('OAUTH_CANCELLED: User cancelled the authentication process or redirect failed'));
                return;
            }
            
            debugLog('OAuth redirect URL received:', redirectUrl);
            debugLog('Full redirect URL for debugging:', redirectUrl);
            
            // Check if Microsoft returned an error in the redirect URL
            if (redirectUrl.includes('error=') || redirectUrl.includes('error_description=')) {
                const urlParams = new URLSearchParams(redirectUrl.split('?')[1] || '');
                const error = urlParams.get('error');
                const errorDescription = urlParams.get('error_description');
                debugError('Microsoft OAuth error in redirect:', { error, errorDescription });
                
                if (error === 'invalid_client') {
                    reject(new Error('Invalid client ID. Please check Azure App Registration configuration.'));
                } else if (error === 'redirect_uri_mismatch') {
                    reject(new Error('Redirect URI mismatch. The redirect URI must be exactly registered in Azure App Registration: https://connector.saai.dev/webhook/outlook/oauth/callback'));
                } else if (error === 'invalid_request') {
                    reject(new Error(`Invalid OAuth request: ${errorDescription || error}`));
                } else {
                    reject(new Error(`Microsoft OAuth error: ${error} - ${errorDescription || 'Unknown error'}`));
                }
                return;
            }
            
            // Step 4: Handle n8n callback response
            // Flow: Microsoft → n8n callback → n8n redirects to extension redirect URI
            // Chrome captures the redirect to extension's redirect URI (chromiumapp.org)
            // n8n should have redirected with JWT data in URL params
            
            try {
                debugLog('Processing redirect URL from n8n:', redirectUrl);
                
                const urlParams = new URLSearchParams(redirectUrl.split('?')[1] || '');
                let jwtToken = urlParams.get('jwt_token') || urlParams.get('jwt');
                let userId = urlParams.get('user_id') || urlParams.get('userId');
                const refreshToken = urlParams.get('refresh_token');
                const returnedState = urlParams.get('state');
                
                debugLog('Extracted from URL params:', {
                    hasJwtToken: !!jwtToken,
                    hasUserId: !!userId,
                    hasRefreshToken: !!refreshToken,
                    hasState: !!returnedState,
                    redirectUrl: redirectUrl
                });
                
                // If JWT not in URL params, n8n might have returned JSON
                // This shouldn't happen if n8n redirects properly, but handle it as fallback
                if (!jwtToken || !userId) {
                    debugLog('JWT not in URL params, checking if n8n callback returned JSON');
                    
                    // If redirectUrl is still the n8n callback (shouldn't happen, but handle it)
                    if (redirectUrl.includes('/webhook/outlook/oauth/callback')) {
                        debugLog('Redirect URL is still n8n callback, attempting to fetch JSON');
                        
                        try {
                            // Fetch the JSON response from callback
                            const callbackResponse = await safeRequest(redirectUrl, {
                                method: 'GET',
                                headers: {
                                    'Accept': 'application/json'
                                },
                                _allowErrorResponse: true
                            });
                            
                            debugLog('Callback fetch response status:', callbackResponse.status);
                            
                            if (callbackResponse.ok) {
                                const responseText = await callbackResponse.text();
                                debugLog('Callback response text (first 200 chars):', responseText.substring(0, 200));
                                
                                let responseJson;
                                try {
                                    responseJson = JSON.parse(responseText);
                                    debugLog('Callback JSON parsed successfully');
                                } catch (parseError) {
                                    debugError('Failed to parse callback JSON:', parseError);
                                    throw new Error('Callback returned invalid JSON');
                                }
                                
                                // Handle array response: [{userId, jwt}]
                                if (Array.isArray(responseJson) && responseJson.length > 0) {
                                    const responseData = responseJson[0];
                                    userId = responseData.userId || responseData.sub;
                                    jwtToken = responseData.jwt || responseData.jwtToken || responseData.token;
                                    debugLog('Extracted from JSON array:', { hasUserId: !!userId, hasJwtToken: !!jwtToken });
                                } else if (responseJson && typeof responseJson === 'object') {
                                    // Handle object response: {userId, jwt}
                                    userId = responseJson.userId || responseJson.sub;
                                    jwtToken = responseJson.jwt || responseJson.jwtToken || responseJson.token;
                                    debugLog('Extracted from JSON object:', { hasUserId: !!userId, hasJwtToken: !!jwtToken });
                                }
                            } else {
                                const errorText = await callbackResponse.text().catch(() => 'Unknown error');
                                debugError('Callback returned error status:', callbackResponse.status, errorText);
                                
                                // Handle specific error status codes
                                if (callbackResponse.status === 400) {
                                    throw new Error('Invalid State: Callback returned 400');
                                } else if (callbackResponse.status === 408) {
                                    throw new Error('NOT_WHITELISTED: User not in allow-list');
                                } else {
                                    throw new Error(`Callback error (${callbackResponse.status}): ${errorText}`);
                                }
                            }
                        } catch (fetchError) {
                            debugError('Failed to fetch JSON from callback:', fetchError);
                            // Don't throw here - check URL params for errors below
                        }
                    }
                }
                
                // Check if this is an error response
                const error = urlParams.get('error');
                const errorCode = urlParams.get('error_code');
                if (error) {
                    const errorDescription = urlParams.get('error_description') || 'Unknown error';
                    
                    // Handle invalid state (HTTP 400)
                    if (errorCode === '400' || error === 'invalid_state') {
                        debugError('Invalid state:', errorDescription);
                        throw new Error(`Invalid State: ${errorDescription}`);
                    }
                    
                    // Handle user not in allow-list (HTTP 408)
                    if (errorCode === '408' || error === 'user_not_whitelisted' || error === 'user_not_in_allowlist') {
                        debugError('User not in allow-list:', errorDescription);
                        throw new Error(`NOT_WHITELISTED: ${errorDescription}`);
                    }
                    
                    // Handle access denied specifically
                    if (error === 'access_denied') {
                        debugError('Access denied:', errorDescription);
                        throw new Error(`Access Denied: ${errorDescription}`);
                    }
                    
                    throw new Error(`OAuth error: ${error} - ${errorDescription}`);
                }
                
                if (!jwtToken || !userId) {
                    debugError('No JWT token or user ID found in redirect URL or callback response');
                    debugError('Redirect URL:', redirectUrl);
                    throw new Error('No JWT token or user ID in OAuth response. Please try again.');
                }
                
                if (returnedState && returnedState !== startData.state) {
                    throw new Error('State mismatch in OAuth response');
                }
                
                debugLog('JWT token extracted: [REDACTED]');
                debugLog('User ID extracted: [REDACTED]');
                debugLog('Refresh token extracted:', refreshToken ? 'Present' : 'Not provided');
                if (returnedState) {
                    debugLog('State verified: [REDACTED]');
                }
                
                // Step 5: Store JWT token, refresh token, and user ID with Outlook provider flag
                const storageData = {
                    isConnected: true, 
                    userId: userId,
                    jwtToken: jwtToken,
                    provider: 'outlook', // Mark as Outlook provider
                    oauthData: {
                        userId: userId,
                        jwtToken: jwtToken,
                        redirectUrl: redirectUrl,
                        provider: 'outlook'
                    }
                };
                
                // Store refresh token if available
                if (refreshToken) {
                    storageData.refreshToken = refreshToken;
                    storageData.oauthData.refreshToken = refreshToken;
                }
                
                await chrome.storage.local.set(storageData);
                
                debugLog('PKCE OAuth flow completed successfully');
                resolve({ 
                    success: true, 
                    userId: userId,
                    jwtToken: jwtToken
                });
                
            } catch (error) {
                debugError('Error in PKCE OAuth flow:', error);
                reject(error);
            }
        });
    });
        
        // Return the result after Promise resolves
        return result;
        
    } catch (error) {
        debugError('PKCE OAuth flow failed:', error);
        debugError('PKCE error details:', {
            message: error.message,
            stack: error.stack
        });
        
        // No fallback - let the OAuth flow fail completely
        throw error;
    } finally {
        // Always release the lock, even if flow fails
        // This executes AFTER the Promise resolves/rejects because we awaited it
        oauthInProgress = false;
        chrome.storage.local.remove('oauthInProgress').then(() => {
            debugLog('OAuth flow lock released from memory and storage');
        }).catch(() => {
            debugLog('OAuth flow lock released from memory (storage cleanup failed)');
        });
    }
}

// Handle extension icon click - opens popup for Outlook connection
chrome.action.onClicked.addListener((tab) => {
    // The popup will handle the connection flow
    // This ensures users connect their Outlook first
    // PRODUCTION: console.log('Extension icon clicked - popup will handle connection');
});

// Handle storage changes
chrome.storage.onChanged.addListener((changes, namespace) => {
    if (namespace === 'local') {
        if (changes.isConnected) {
            // PRODUCTION: console.log('Connection status updated:', changes.isConnected.newValue);
        }
        if (changes.userId) {
            // PRODUCTION: console.log('User ID updated:', changes.userId.newValue);
        }
        if (changes.jwtToken) {
            // PRODUCTION: console.log('JWT token updated:', changes.jwtToken.newValue ? 'Token present' : 'Token removed');
        }
    }
});
