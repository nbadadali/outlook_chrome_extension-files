// Production-safe logging - ALL logging disabled for Chrome Web Store compliance
// For local debugging only: Set PRODUCTION_MODE = false in all 3 files
const PRODUCTION_MODE = false; // TEMPORARILY ENABLED FOR DEBUGGING

// No-op functions for production - completely silent
function debugLog() {}
function debugError() {}
function debugWarn() {}

document.addEventListener('DOMContentLoaded', function() {
    const authorizeBtn = document.getElementById('authorize');
    const testBtn = document.getElementById('test-connection');
    const statusDiv = document.getElementById('status');
    
    // Log extension redirect URI for Azure configuration
    const extensionRedirectUri = chrome.identity.getRedirectURL();
    console.log('🔑 EXTENSION REDIRECT URI (add this to Azure):', extensionRedirectUri);
    console.log('📋 Copy this URI and add it to Azure App Registration → Authentication → Redirect URIs');

    chrome.storage.local.get(['userId', 'isConnected', 'provider'], function(result) {
        debugLog('Popup loaded, storage result:', result);
        // Check if connected and provider is Outlook
        if (result.userId && (result.provider === 'outlook' || !result.provider)) {
            // If userId is present and provider is Outlook (or not set), ensure isConnected is set
            chrome.storage.local.set({ isConnected: true, provider: 'outlook' });
            debugLog('Found userId, setting isConnected and showing test button');
            showStatus('Connected to Sa.AI Inbox Assistant', 'success');
            authorizeBtn.style.display = 'none';
            testBtn.style.display = 'block';
        } else if (result.isConnected && (result.provider === 'outlook' || !result.provider)) {
            // Fallback for legacy state
            debugLog('Found legacy isConnected, showing test button');
            showStatus('Connected to Sa.AI Inbox Assistant', 'success');
            authorizeBtn.style.display = 'none';
            testBtn.style.display = 'block';
        } else {
            debugLog('No connection found, showing authorize button');
        }
    });

    // Debounce flag to prevent rapid clicks
    let isConnecting = false;
    
    authorizeBtn.addEventListener('click', function() {
        // Prevent multiple rapid clicks
        if (isConnecting) {
            debugLog('OAuth connection already in progress, ignoring click');
            return;
        }
        
        // Check if already connected
        chrome.storage.local.get(['userId', 'isConnected', 'provider'], function(result) {
            if ((result.userId || result.isConnected) && (result.provider === 'outlook' || !result.provider)) {
                showStatus('Already connected to Sa.AI!', 'info');
                authorizeBtn.style.display = 'none';
                testBtn.style.display = 'block';
                return;
            }
            
            // Set connecting flag
            isConnecting = true;
            showStatus('Connecting to Microsoft...', 'info');
            
            // Send message to background script to handle OAuth
            debugLog('Sending OAuth request to background script...');
            chrome.runtime.sendMessage({
                action: 'sendToN8N',
                data: {
                    endpoint: 'oauth',
                    payload: { context: 'OutlookConnectClicked' }
                }
            }, (response) => {
                // Reset connecting flag
                isConnecting = false;
                
                // Check for Chrome runtime errors
                if (chrome.runtime.lastError) {
                    debugError('Chrome runtime error:', chrome.runtime.lastError);
                    showStatus('❌ Connection failed: ' + chrome.runtime.lastError.message, 'error');
                    return;
                }
                
                // Check if response exists
                if (!response) {
                    debugError('No response from background script');
                    showStatus('❌ No response from extension. Please reload and try again.', 'error');
                    return;
                }
                
                debugLog('Response received from background:', response);
                
                if (response?.success) {
                    debugLog('OAuth success response from background:', response.data);
                    showStatus('✅ Connected to Sa.AI!', 'success');
                    authorizeBtn.style.display = 'none';
                    testBtn.style.display = 'block';
                    debugLog('UI updated - authorizeBtn hidden, testBtn shown');
                } else {
                    debugError('OAuth via background failed:', response?.error);
                    // Handle specific error messages
                    if (response?.error && response.error.includes('already in progress')) {
                        showStatus('⏳ Authentication in progress. Please wait...', 'info');
                    } else if (response?.error && response.error.includes('NOT_WHITELISTED')) {
                        showStatus('❌ User not in allow-list. Please contact support.', 'error');
                    } else if (response?.error && response.error.includes('Invalid State')) {
                        showStatus('❌ Authentication error. Please try again.', 'error');
                    } else {
                        showStatus('❌ Connection failed: ' + (response?.error || 'Unknown error'), 'error');
                    }
                }
            });
        });
    });

    testBtn.addEventListener('click', function() {
        debugLog('Test button clicked');
        chrome.tabs.query({active: true, currentWindow: true}, function(tabs) {
            debugLog('Current tab:', tabs[0]);
            const isOutlook = tabs[0].url && (
                tabs[0].url.includes('outlook.office.com') || 
                tabs[0].url.includes('outlook.live.com') || 
                tabs[0].url.includes('outlook.com')
            );
            
            if (isOutlook) {
                debugLog('Attempting to communicate with tab:', tabs[0].id);
                
                // Ping first to check if content script is already loaded
                chrome.tabs.sendMessage(tabs[0].id, {action: 'ping'}, function(pingResponse) {
                    if (chrome.runtime.lastError) {
                        // Content script not loaded - inject it first
                        debugLog('Content script not found, injecting...');
                        chrome.scripting.executeScript({
                            target: { tabId: tabs[0].id },
                            files: ['content-outlook.js']
                        }, function() {
                            if (chrome.runtime.lastError) {
                                debugError('Script injection error:', chrome.runtime.lastError);
                                showStatus('Error injecting script: ' + chrome.runtime.lastError.message, 'error');
                                return;
                            }
                            
                            // Wait for script to initialize, then send message
                            setTimeout(() => {
                                sendOpenSaaiMessage(tabs[0].id);
                            }, 500);
                        });
                    } else {
                        // Content script already loaded - send message directly
                        debugLog('Content script ready, sending message directly');
                        sendOpenSaaiMessage(tabs[0].id);
                    }
                });
            } else {
                debugLog('Not on Outlook, current URL:', tabs[0].url);
                showStatus('Please open Outlook first to use the assistant', 'error');
            }
        });
    });

    // Helper function to show status messages
    function showStatus(message, type) {
        statusDiv.innerHTML = `<div class="status ${type}">${message}</div>`;
        if (type !== 'info') {
            setTimeout(() => {
                statusDiv.innerHTML = '';
            }, 3000);
        }
    }
});

// Helper function to send the open_saai message
function sendOpenSaaiMessage(tabId) {
    const statusDiv = document.getElementById('status');
    
    // Helper function for status in this scope
    function showStatusMessage(message, type) {
        if (statusDiv) {
            statusDiv.innerHTML = `<div class="status ${type}">${message}</div>`;
            if (type !== 'info') {
                setTimeout(() => {
                    statusDiv.innerHTML = '';
                }, 3000);
            }
        }
    }
    
    chrome.tabs.sendMessage(tabId, {action: 'open_saai'}, function(response) {
        debugLog('Message response:', response);
        if (chrome.runtime.lastError) {
            debugError('Message error:', chrome.runtime.lastError);
            showStatusMessage('Error: ' + chrome.runtime.lastError.message, 'error');
        } else {
            debugLog('Sa.AI Assistant opened successfully');
            window.close();
        }
    });
}
