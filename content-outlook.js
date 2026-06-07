// Sa.AI Inbox Assistant Content Script
// Complete rewrite for seamless Outlook integration
// Wrapped in IIFE so re-injection (manifest + programmatic) does not redeclare globals
(function() {
'use strict';

// Production-safe logging - ALL logging disabled for Chrome Web Store compliance
// For local debugging only: Set PRODUCTION_MODE = false in all 3 files
var PRODUCTION_MODE = false;  // ENABLED FOR DEBUGGING – set true when done

// No-op functions for production - completely silent
function debugLog() {}
function debugError() {}
function debugWarn() {}

// Message shown when extension was reloaded and content script is orphaned (no receiving end)
var SAAI_REFRESH_PAGE_MSG = 'The extension was reloaded or updated. Please refresh this page (F5 or reload) to connect again.';

function isExtensionContextInvalidated(err) {
  if (!err || !err.message) return false;
  var msg = (err.message || '').toLowerCase();
  return (
    msg.indexOf('receiving end does not exist') !== -1 ||
    msg.indexOf('extension context invalidated') !== -1 ||
    msg.indexOf('could not establish connection') !== -1
  );
}

function showRefreshPageMessage(chatArea) {
  var text = SAAI_REFRESH_PAGE_MSG;
  if (chatArea) {
    appendMessage('bot', text, chatArea);
  } else {
    var sidebar = document.getElementById('saai-outlook-sidebar');
    if (sidebar) {
      var area = sidebar.querySelector('#saai-chat-area');
      if (area) appendMessage('bot', text, area);
    }
  }
}

// Robust network request wrapper with timeout and error handling
async function safeRequest(url, options = {}, timeout = 50000) {
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), timeout);
  
  try {
    const response = await fetch(url, {
      ...options,
      signal: controller.signal
    });
    
    clearTimeout(timeoutId);
    
    if (!response.ok) {
      throw new Error(`HTTP ${response.status}: ${response.statusText}`);
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

// Safe storage operations with error handling
async function safeStorageGet(keys) {
  try {
    return await chrome.storage.local.get(keys);
  } catch (error) {
    debugError('Storage get failed:', error);
    return {};
  }
}

async function safeStorageSet(data) {
  try {
    await chrome.storage.local.set(data);
    return true;
  } catch (error) {
    debugError('Storage set failed:', error);
    return false;
  }
}

// Performance optimization - DOM element cache
const domCache = new Map();
function getCachedElement(id) {
  if (!domCache.has(id)) {
    domCache.set(id, document.getElementById(id));
  }
  return domCache.get(id);
}

// User experience - loading states and network monitoring
function showLoadingState(element, message = 'Loading...') {
  if (element) {
    element.innerHTML = `<div style="display: flex; align-items: center; gap: 8px;"><div style="width: 16px; height: 16px; border: 2px solid #ccc; border-top: 2px solid #0f172a; border-radius: 50%; animation: spin 1s linear infinite;"></div>${message}</div>`;
  }
}

// Enhanced thinking indicator with Sa.AI branding
function showThinkingIndicator(chatArea) {
  const thinkingDiv = document.createElement('div');
  thinkingDiv.className = 'message bot-message temporary';
  thinkingDiv.innerHTML = `
    <div class="message-content saai-thinking">
      <div class="saai-thinking-dots">
        <div class="saai-thinking-dot"></div>
        <div class="saai-thinking-dot"></div>
        <div class="saai-thinking-dot"></div>
      </div>
      <span class="saai-thinking-text">Sa.AI is thinking...</span>
    </div>
  `;
  
  chatArea.appendChild(thinkingDiv);
  chatArea.scrollTop = chatArea.scrollHeight;
  
  return thinkingDiv;
}

// Utility to safely escape HTML entities in dynamic text
function escapeHtml(str) {
  if (typeof str !== 'string') {
    return str;
  }
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function formatCardTimestamp(date = new Date()) {
  try {
    return date.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
  } catch (error) {
    debugError('Failed to format timestamp for card', error);
    return '';
  }
}

function extractPlainTextContent(value) {
  if (!value) {
    return '';
  }
  const temp = document.createElement('div');
  temp.innerHTML = value;
  return (temp.textContent || temp.innerText || '').trim();
}

function parseModernSummaryCard(text) {
  if (!text || typeof text !== 'string') {
    return null;
  }

  const normalized = text.replace(/\r/g, '\n').trim();
  if (!normalized) {
    return null;
  }

  // Check for structured email summary patterns
  const hasSubjectPattern = /(?:^|\n)\s*\d+\.\s*Subject:\s*/i.test(normalized);
  const hasSenderPattern = /(?:^|\n)\s*(?:—|[-*])\s*Sender:\s*/i.test(normalized);
  const hasDatePattern = /(?:^|\n)\s*(?:—|[-*])\s*Date:\s*/i.test(normalized);
  
  // If it looks like an email summary with subjects, parse it specially
  if (hasSubjectPattern || (hasSenderPattern && hasDatePattern)) {
    return parseEmailSummaryContent(normalized);
  }

  // Original parsing logic for other structured content
  const lines = normalized.split(/\n+/).map(line => line.trim()).filter(Boolean);
  if (lines.length < 2) {
    return null;
  }

  const potentialTitle = lines[0];
  if (potentialTitle.length < 6 || potentialTitle.length > 140) {
    return null;
  }

  const items = [];
  for (let i = 1; i < lines.length; i++) {
    const line = lines[i];

    const bulletMatch = line.match(/^(?:[-•*]|[✅✔☑✓])\s*(.+)$/);
    if (bulletMatch && bulletMatch[1].trim().length > 0) {
      items.push(bulletMatch[1].trim());
      continue;
    }

    const numberedMatch = line.match(/^(\d+[\).]?\s*)(.+)$/);
    if (numberedMatch && numberedMatch[2].trim().length > 0) {
      items.push(`${numberedMatch[1].trim()} ${numberedMatch[2].trim()}`);
      continue;
    }

    const simpleMatch = line.match(/^([0-9]+\s+.+)$/);
    if (simpleMatch) {
      items.push(simpleMatch[1].trim());
      continue;
    }

    if (/^note[:\-]/i.test(line)) {
      items.push(line.trim());
      continue;
    }

    if (line.length <= 120 && items.length > 0) {
      const previous = items[items.length - 1];
      items[items.length - 1] = `${previous} ${line}`.trim();
    }
  }

  const filteredItems = items.filter(item => item && item.length > 0).slice(0, 6);

  if (filteredItems.length < 2) {
    return null;
  }

  const title = potentialTitle.replace(/\s+$/, '');
  return { title, items: filteredItems };
}

function parseEmailSummaryContent(text) {
  // Split by sentences and common email patterns
  const items = [];
  
  // Split by numbered subjects first
  const subjectSections = text.split(/(?=\d+\.\s*Subject:)/i);
  
  for (const section of subjectSections) {
    if (!section.trim()) continue;
    
    const lines = section.split(/\n+/).map(line => line.trim()).filter(Boolean);
    
    for (const line of lines) {
      // Skip very short lines
      if (line.length < 10) continue;
      
      // Add subjects, senders, dates, reasons, etc.
      if (/^\d+\.\s*Subject:/i.test(line) || 
          /^(?:—|[-*])\s*(?:Sender|Date|Reason):/i.test(line) ||
          /^(?:Sender|Date|Reason):/i.test(line)) {
        items.push(line);
      } else if (line.length > 20 && line.length < 300) {
        // Add substantial content lines
        items.push(line);
      }
    }
  }
  
  // If no structured items found, treat as simple content
  if (items.length === 0) {
    const sentences = text.split(/[.!?]+/).map(s => s.trim()).filter(s => s.length > 20);
    if (sentences.length > 0) {
      return { title: sentences[0], items: sentences.slice(1, 4) };
    }
    return null;
  }
  
  const title = items[0] || 'Email Summary';
  return { title, items: items.slice(1) };
}

function parseSummaryWithSources(text) {
  if (!text || typeof text !== 'string') {
    return null;
  }
  
  // PRODUCTION: console.log('🔍 Parsing text for sources:', text); // Debug log
  
  // Check if text contains any source indicators
  const hasSourceIndicators = text.includes('Sources:') || 
                              text.includes('Source:') ||
                              /\[\d+\]\(https?:\/\//.test(text) ||
                              text.includes('https://') ||
                              text.includes('http://');
  
  // If no source indicators, not a summary with sources
  if (!hasSourceIndicators && !text.startsWith('Summary:') && !text.startsWith('Insight:')) {
    return null;
  }
  
  let summaryText = '';
  let sourcesText = '';
  let sources = [];
  
  // Try multiple parsing approaches
  
  // Approach 1: Split by "Sources:" (most common)
  if (text.includes('Sources:')) {
    const parts = text.split(/\s*Sources?:\s*/i);
    if (parts.length >= 2) {
      summaryText = parts[0].trim();
      sourcesText = parts.slice(1).join(' Sources: ').trim();
    }
  }
  // Approach 2: Split by "Source:" (singular)  
  else if (text.includes('Source:')) {
    const parts = text.split(/\s*Source:\s*/i);
    if (parts.length >= 2) {
      summaryText = parts[0].trim();
      sourcesText = parts.slice(1).join(' Source: ').trim();
      // PRODUCTION: console.log('📝 Split by Source:', { summaryText, sourcesText });
    }
  }
  // Approach 3: Extract inline sources from entire text
  else {
    summaryText = text;
    sourcesText = text; // Check entire text for sources
    // PRODUCTION: console.log('📝 Using entire text for source extraction');
  }
  
  // Clean up summary text prefixes
  summaryText = summaryText.replace(/^(Summary|Insight):\s*/i, '').trim();
  
  // Extract all possible source formats
  if (sourcesText) {
    // Format 1: [1](url), [2](url) - Markdown links
    const markdownLinks = sourcesText.match(/\[(\d+)\]\((https?:\/\/[^)]+)\)/g);
    if (markdownLinks) {
      // PRODUCTION: console.log('🔗 Found markdown links:', markdownLinks);
      markdownLinks.forEach(match => {
        const linkMatch = match.match(/\[(\d+)\]\((https?:\/\/[^)]+)\)/);
        if (linkMatch) {
          sources.push({
            number: linkMatch[1],
            url: linkMatch[2]
          });
        }
      });
    }
    
    // Format 2: 1. https://url, 2. https://url - Numbered URLs
    const numberedUrls = sourcesText.match(/(\d+)\.\s*(https?:\/\/[^\s,]+)/g);
    if (numberedUrls) {
      // PRODUCTION: console.log('🔗 Found numbered URLs:', numberedUrls);
      numberedUrls.forEach(match => {
        const urlMatch = match.match(/(\d+)\.\s*(https?:\/\/[^\s,]+)/);
        if (urlMatch) {
          sources.push({
            number: urlMatch[1],
            url: urlMatch[2]
          });
        }
      });
    }
    
    // Format 3: Just URLs with numbers nearby - 1 https://url, 2 https://url
    const nearbyNumberUrls = sourcesText.match(/(\d+)\s+(https?:\/\/[^\s,]+)/g);
    if (nearbyNumberUrls) {
      // PRODUCTION: console.log('🔗 Found nearby number URLs:', nearbyNumberUrls);
      nearbyNumberUrls.forEach(match => {
        const urlMatch = match.match(/(\d+)\s+(https?:\/\/[^\s,]+)/);
        if (urlMatch) {
          const exists = sources.some(s => s.number === urlMatch[1]);
          if (!exists) {
            sources.push({
              number: urlMatch[1],
              url: urlMatch[2]
            });
          }
        }
      });
    }
    
    // Format 4: Plain URLs without numbers
    if (sources.length === 0) {
      const plainUrls = sourcesText.match(/(https?:\/\/[^\s,)]+)/g);
      if (plainUrls) {
        // PRODUCTION: console.log('🔗 Found plain URLs:', plainUrls);
        plainUrls.forEach((url, index) => {
          sources.push({
            number: (index + 1).toString(),
            url: url
          });
        });
      }
    }
  }
  
  // Remove source references from summary text for clean display
  summaryText = summaryText.replace(/\[\d+\]\(https?:\/\/[^)]+\)/g, '').trim();
  summaryText = summaryText.replace(/\[\d+\]/g, '').trim();
  
  // PRODUCTION: console.log('✅ Final parsing result:', { 
  //   summaryLength: summaryText.length, 
  //   sourcesCount: sources.length
  // });
  
  if (!summaryText || summaryText.length < 10) {
    return null;
  }
  
  return {
    summary: summaryText,
    sources: sources
  };
}

function buildSummaryWithSourcesHtml(data) {
  if (!data || !data.summary) {
    return null;
  }
  
  const summaryHtml = `
    <div class="saai-summary-section">
      <div class="saai-summary-header"><strong>Summary:</strong></div>
      <div class="saai-summary-text">${escapeHtml(data.summary)}</div>
    </div>
  `;
  
  let sourcesHtml = '';
  if (data.sources && data.sources.length > 0) {
    const sourcesList = data.sources
      .map(source => `
        <div class="saai-source-item">
          <span class="saai-modern-card-dash" aria-hidden="true">—</span>
          <span class="saai-source-text">
            <a href="${escapeHtml(source.url)}" target="_blank" rel="noopener noreferrer" class="saai-source-link">
              [${escapeHtml(source.number)}] ${escapeHtml(source.url)}
            </a>
          </span>
        </div>
      `)
      .join('');
    
    sourcesHtml = `
      <div class="saai-sources-section">
        <div class="saai-sources-header"><strong>Sources:</strong></div>
        <div class="saai-sources-list">
          ${sourcesList}
        </div>
      </div>
    `;
  }
  
  return `
    <div class="saai-modern-card-wrapper">
      <div class="saai-modern-card" role="group" aria-label="Sa.AI summary with sources">
        <div class="saai-modern-card-body">
          ${summaryHtml}
          ${sourcesHtml}
        </div>
        <div class="saai-modern-card-footer">
          <span class="saai-modern-card-time">${escapeHtml(formatCardTimestamp())}</span>
          <div class="saai-modern-card-actions" aria-hidden="true">
            <span class="saai-modern-card-action" title="Thumbs up">
              <svg width="16" height="16" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
                <path d="M6.6665 14.6667H13.3332C13.7013 14.6667 14.0548 14.5202 14.3123 14.2626C14.5698 14.0051 14.7163 13.6516 14.7163 13.2834V7.78338C14.7163 7.41522 14.5698 7.06174 14.3123 6.80421C14.0548 6.54669 13.7013 6.40022 13.3332 6.40022H10.6665L11.3132 3.46689C11.346 3.31244 11.3409 3.15178 11.2985 3.00033C11.2562 2.84888 11.1778 2.71204 11.0705 2.60295C10.9632 2.49386 10.8309 2.41674 10.6866 2.37939C10.5422 2.34203 10.3909 2.3457 10.2498 2.38989L5.99984 3.73322C5.79423 3.79798 5.61245 3.92036 5.47629 4.08507C5.34013 4.24977 5.2557 4.44994 5.23317 4.66155L4.6665 10.0002V13.3336" stroke="#475569" stroke-width="1.2" stroke-linecap="round" stroke-linejoin="round"/>
                <path d="M3.3335 6.66699H1.3335V13.3337H3.3335V6.66699Z" stroke="#475569" stroke-width="1.2" stroke-linecap="round" stroke-linejoin="round"/>
              </svg>
            </span>
            <span class="saai-modern-card-action" title="Thumbs down">
              <svg width="16" height="16" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
                <path d="M9.33346 1.33301H2.66679C2.29863 1.33301 1.94515 1.47948 1.68762 1.73701C1.4301 1.99453 1.28363 2.34801 1.28363 2.71618L1.28346 8.21601C1.28346 8.58417 1.42993 8.93765 1.68746 9.19518C1.94499 9.45271 2.29846 9.59918 2.66663 9.59918H5.33329L4.68663 12.5325C4.65393 12.687 4.65908 12.8477 4.70145 12.9991C4.74382 13.1506 4.82219 13.2874 4.92947 13.3965C5.03675 13.5056 5.16902 13.5827 5.31336 13.6201C5.45771 13.6574 5.60905 13.6538 5.75013 13.6096L10.0001 12.2663C10.2058 12.2015 10.3876 12.0791 10.5237 11.9144C10.6599 11.7497 10.7443 11.5495 10.7668 11.3379L11.3335 5.99918V2.66584" stroke="#475569" stroke-width="1.2" stroke-linecap="round" stroke-linejoin="round"/>
                <path d="M12.6665 9.33268H14.6665V2.66602H12.6665V9.33268Z" stroke="#475569" stroke-width="1.2" stroke-linecap="round" stroke-linejoin="round"/>
              </svg>
            </span>
          </div>
        </div>
      </div>
    </div>
  `;
}

function buildModernSummaryCardHtml(cardData) {
  if (!cardData) {
    return null;
  }

  const subjectPattern = /^(?:\d+\.\s*)?Subject:\s*/i;
  const titleText = cardData.title ? cardData.title.trim() : '';
  const titleLooksLikeSubject = titleText && subjectPattern.test(titleText);

  const itemsToRender = titleLooksLikeSubject
    ? [titleText, ...cardData.items]
    : cardData.items;

  const groups = [];
  let currentSubjectGroup = null;

  itemsToRender.forEach(rawItem => {
    const originalText = typeof rawItem === 'string' ? rawItem.trim() : '';
    if (!originalText) {
      return;
    }

    const escapedText = escapeHtml(originalText);

    if (subjectPattern.test(originalText)) {
      if (currentSubjectGroup) {
        groups.push(currentSubjectGroup);
      }

      currentSubjectGroup = {
        subject: `<strong>${escapedText}</strong>`,
        rows: []
      };
    } else if (currentSubjectGroup) {
      currentSubjectGroup.rows.push(escapedText);
    } else {
      groups.push({ subject: null, rows: [escapedText] });
    }
  });

  if (currentSubjectGroup) {
    groups.push(currentSubjectGroup);
  }

  const renderRow = (text) => `
        <div class="saai-modern-card-item">
          <span class="saai-modern-card-dash" aria-hidden="true">—</span>
          <span class="saai-modern-card-text">${text}</span>
        </div>
      `;

  const sectionsHtml = groups
    .map(group => {
      if (group.subject) {
        const rowsHtml = group.rows.length > 0
          ? group.rows.map(renderRow).join('')
          : '';

        return `
        <div class="saai-modern-card-section">
          <div class="saai-modern-card-subject">${group.subject}</div>
          ${rowsHtml}
        </div>
      `;
      }

      return group.rows.map(renderRow).join('');
    })
    .join('');

  const headerHtml = !titleLooksLikeSubject && titleText
    ? `<div class="saai-modern-card-header"><div class="saai-modern-card-title">${escapeHtml(titleText)}</div></div>`
    : '';

  return `
    <div class="saai-modern-card-wrapper">
      <div class="saai-modern-card" role="group" aria-label="Sa.AI summary">
        ${headerHtml}
        <div class="saai-modern-card-body">
          ${sectionsHtml}
        </div>
        <div class="saai-modern-card-footer">
          <span class="saai-modern-card-time">${escapeHtml(formatCardTimestamp())}</span>
          <div class="saai-modern-card-actions" aria-hidden="true">
            <span class="saai-modern-card-action" title="Thumbs up">
              <svg width="16" height="16" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
                <path d="M6.6665 14.6667H13.3332C13.7013 14.6667 14.0548 14.5202 14.3123 14.2626C14.5698 14.0051 14.7163 13.6516 14.7163 13.2834V7.78338C14.7163 7.41522 14.5698 7.06174 14.3123 6.80421C14.0548 6.54669 13.7013 6.40022 13.3332 6.40022H10.6665L11.3132 3.46689C11.346 3.31244 11.3409 3.15178 11.2985 3.00033C11.2562 2.84888 11.1778 2.71204 11.0705 2.60295C10.9632 2.49386 10.8309 2.41674 10.6866 2.37939C10.5422 2.34203 10.3909 2.3457 10.2498 2.38989L5.99984 3.73322C5.79423 3.79798 5.61245 3.92036 5.47629 4.08507C5.34013 4.24977 5.2557 4.44994 5.23317 4.66155L4.6665 10.0002V13.3336" stroke="#475569" stroke-width="1.2" stroke-linecap="round" stroke-linejoin="round"/>
                <path d="M3.3335 6.66699H1.3335V13.3337H3.3335V6.66699Z" stroke="#475569" stroke-width="1.2" stroke-linecap="round" stroke-linejoin="round"/>
              </svg>
            </span>
            <span class="saai-modern-card-action" title="Thumbs down">
              <svg width="16" height="16" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
                <path d="M9.33346 1.33301H2.66679C2.29863 1.33301 1.94515 1.47948 1.68762 1.73701C1.4301 1.99453 1.28363 2.34801 1.28363 2.71618L1.28346 8.21601C1.28346 8.58417 1.42993 8.93765 1.68746 9.19518C1.94499 9.45271 2.29846 9.59918 2.66663 9.59918H5.33329L4.68663 12.5325C4.65393 12.687 4.65908 12.8477 4.70145 12.9991C4.74382 13.1506 4.82219 13.2874 4.92947 13.3965C5.03675 13.5056 5.16902 13.5827 5.31336 13.6201C5.45771 13.6574 5.60905 13.6538 5.75013 13.6096L10.0001 12.2663C10.2058 12.2015 10.3876 12.0791 10.5237 11.9144C10.6599 11.7497 10.7443 11.5495 10.7668 11.3379L11.3335 5.99918V2.66584" stroke="#475569" stroke-width="1.2" stroke-linecap="round" stroke-linejoin="round"/>
                <path d="M12.6665 9.33268H14.6665V2.66602H12.6665V9.33268Z" stroke="#475569" stroke-width="1.2" stroke-linecap="round" stroke-linejoin="round"/>
              </svg>
            </span>
          </div>
        </div>
      </div>
    </div>
  `;
}

function buildSimpleModernCardHtml(cardData) {
  if (!cardData) {
    return null;
  }

  return `
    <div class="saai-modern-card-wrapper">
      <div class="saai-modern-card saai-simple-card" role="group" aria-label="Sa.AI response">
        <div class="saai-modern-card-body">
          <div class="saai-simple-card-text">${escapeHtml(cardData.title)}</div>
        </div>
        <div class="saai-modern-card-footer">
          <span class="saai-modern-card-time">${escapeHtml(formatCardTimestamp())}</span>
          <div class="saai-modern-card-actions" aria-hidden="true">
            <span class="saai-modern-card-action" title="Thumbs up">
              <svg width="16" height="16" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
                <path d="M6.6665 14.6667H13.3332C13.7013 14.6667 14.0548 14.5202 14.3123 14.2626C14.5698 14.0051 14.7163 13.6516 14.7163 13.2834V7.78338C14.7163 7.41522 14.5698 7.06174 14.3123 6.80421C14.0548 6.54669 13.7013 6.40022 13.3332 6.40022H10.6665L11.3132 3.46689C11.346 3.31244 11.3409 3.15178 11.2985 3.00033C11.2562 2.84888 11.1778 2.71204 11.0705 2.60295C10.9632 2.49386 10.8309 2.41674 10.6866 2.37939C10.5422 2.34203 10.3909 2.3457 10.2498 2.38989L5.99984 3.73322C5.79423 3.79798 5.61245 3.92036 5.47629 4.08507C5.34013 4.24977 5.2557 4.44994 5.23317 4.66155L4.6665 10.0002V13.3336" stroke="#475569" stroke-width="1.2" stroke-linecap="round" stroke-linejoin="round"/>
                <path d="M3.3335 6.66699H1.3335V13.3337H3.3335V6.66699Z" stroke="#475569" stroke-width="1.2" stroke-linecap="round" stroke-linejoin="round"/>
              </svg>
            </span>
            <span class="saai-modern-card-action" title="Thumbs down">
              <svg width="16" height="16" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
                <path d="M9.33346 1.33301H2.66679C2.29863 1.33301 1.94515 1.47948 1.68762 1.73701C1.4301 1.99453 1.28363 2.34801 1.28363 2.71618L1.28346 8.21601C1.28346 8.58417 1.42993 8.93765 1.68746 9.19518C1.94499 9.45271 2.29846 9.59918 2.66663 9.59918H5.33329L4.68663 12.5325C4.65393 12.687 4.65908 12.8477 4.70145 12.9991C4.74382 13.1506 4.82219 13.2874 4.92947 13.3965C5.03675 13.5056 5.16902 13.5827 5.31336 13.6201C5.45771 13.6574 5.60905 13.6538 5.75013 13.6096L10.0001 12.2663C10.2058 12.2015 10.3876 12.0791 10.5237 11.9144C10.6599 11.7497 10.7443 11.5495 10.7668 11.3379L11.3335 5.99918V2.66584" stroke="#475569" stroke-width="1.2" stroke-linecap="round" stroke-linejoin="round"/>
                <path d="M12.6665 9.33268H14.6665V2.66602H12.6665V9.33268Z" stroke="#475569" stroke-width="1.2" stroke-linecap="round" stroke-linejoin="round"/>
              </svg>
            </span>
          </div>
        </div>
      </div>
    </div>
  `;
}

function tryFormatModernCard(content) {
  const plainText = extractPlainTextContent(content);
  
  // Check for Summary/Insight with Sources format first
  const summaryWithSources = parseSummaryWithSources(plainText);
  if (summaryWithSources) {
    return buildSummaryWithSourcesHtml(summaryWithSources);
  }
  
  // Try structured card format for email summaries with subjects
  const cardData = parseModernSummaryCard(plainText);
  if (cardData && cardData.items && cardData.items.length > 0) {
    // Check if this looks like an email summary with subjects
    const hasSubjects = cardData.items.some(item => 
      /^\d+\.\s*Subject:/i.test(item) || /Subject:/i.test(item)
    );
    
    if (hasSubjects) {
      return buildModernSummaryCardHtml(cardData);
    }
  }
  
  // Check for other structured patterns
  if (plainText) {
    // Pattern: "Here are the subjects of emails from X: ..."
    if (/^Here are the subjects of emails from/i.test(plainText)) {
      return buildEmailSubjectsCard(plainText);
    }
    
    // Pattern: Any text starting with "Insight:" 
    if (/^Insight:/i.test(plainText)) {
      const simpleCardData = {
        title: plainText,
        items: []
      };
      return buildSimpleModernCardHtml(simpleCardData);
    }
  }
  
  // For all other content, use simple card format for consistency
  if (plainText && plainText.length > 10) {
    const simpleCardData = {
      title: plainText,
      items: []
    };
    return buildSimpleModernCardHtml(simpleCardData);
  }
  
  return null;
}

function buildEmailSubjectsCard(text) {
  // Parse email subjects list format
  const lines = text.split(/\n+/).map(line => line.trim()).filter(Boolean);
  if (lines.length < 2) {
    return buildSimpleModernCardHtml({ title: text, items: [] });
  }
  
  const title = lines[0];
  const subjects = [];
  
  for (let i = 1; i < lines.length; i++) {
    const line = lines[i];
    // Match numbered subjects or bullet points
    if (/^\d+\.\s*/.test(line) || /^[-•*]\s*/.test(line)) {
      subjects.push(line);
    } else if (line.length > 10) {
      subjects.push(line);
    }
  }
  
  if (subjects.length === 0) {
    return buildSimpleModernCardHtml({ title: text, items: [] });
  }
  
  const itemsHtml = subjects
    .map(subject => `
      <div class="saai-modern-card-item">
        <span class="saai-modern-card-dash" aria-hidden="true">—</span>
        <span class="saai-modern-card-text">${escapeHtml(subject)}</span>
      </div>
    `)
    .join('');
  
  return `
    <div class="saai-modern-card-wrapper">
      <div class="saai-modern-card" role="group" aria-label="Email subjects">
        <div class="saai-modern-card-header">
          <div class="saai-modern-card-title">${escapeHtml(title)}</div>
        </div>
        <div class="saai-modern-card-body">
          ${itemsHtml}
        </div>
        <div class="saai-modern-card-footer">
          <span class="saai-modern-card-time">${escapeHtml(formatCardTimestamp())}</span>
          <div class="saai-modern-card-actions" aria-hidden="true">
            <span class="saai-modern-card-action" title="Thumbs up">
              <svg width="16" height="16" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
                <path d="M6.6665 14.6667H13.3332C13.7013 14.6667 14.0548 14.5202 14.3123 14.2626C14.5698 14.0051 14.7163 13.6516 14.7163 13.2834V7.78338C14.7163 7.41522 14.5698 7.06174 14.3123 6.80421C14.0548 6.54669 13.7013 6.40022 13.3332 6.40022H10.6665L11.3132 3.46689C11.346 3.31244 11.3409 3.15178 11.2985 3.00033C11.2562 2.84888 11.1778 2.71204 11.0705 2.60295C10.9632 2.49386 10.8309 2.41674 10.6866 2.37939C10.5422 2.34203 10.3909 2.3457 10.2498 2.38989L5.99984 3.73322C5.79423 3.79798 5.61245 3.92036 5.47629 4.08507C5.34013 4.24977 5.2557 4.44994 5.23317 4.66155L4.6665 10.0002V13.3336" stroke="#475569" stroke-width="1.2" stroke-linecap="round" stroke-linejoin="round"/>
                <path d="M3.3335 6.66699H1.3335V13.3337H3.3335V6.66699Z" stroke="#475569" stroke-width="1.2" stroke-linecap="round" stroke-linejoin="round"/>
              </svg>
            </span>
            <span class="saai-modern-card-action" title="Thumbs down">
              <svg width="16" height="16" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
                <path d="M9.33346 1.33301H2.66679C2.29863 1.33301 1.94515 1.47948 1.68762 1.73701C1.4301 1.99453 1.28363 2.34801 1.28363 2.71618L1.28346 8.21601C1.28346 8.58417 1.42993 8.93765 1.68746 9.19518C1.94499 9.45271 2.29846 9.59918 2.66663 9.59918H5.33329L4.68663 12.5325C4.65393 12.687 4.65908 12.8477 4.70145 12.9991C4.74382 13.1506 4.82219 13.2874 4.92947 13.3965C5.03675 13.5056 5.16902 13.5827 5.31336 13.6201C5.45771 13.6574 5.60905 13.6538 5.75013 13.6096L10.0001 12.2663C10.2058 12.2015 10.3876 12.0791 10.5237 11.9144C10.6599 11.7497 10.7443 11.5495 10.7668 11.3379L11.3335 5.99918V2.66584" stroke="#475569" stroke-width="1.2" stroke-linecap="round" stroke-linejoin="round"/>
                <path d="M12.6665 9.33268H14.6665V2.66602H12.6665V9.33268Z" stroke="#475569" stroke-width="1.2" stroke-linecap="round" stroke-linejoin="round"/>
              </svg>
            </span>
          </div>
        </div>
      </div>
    </div>
  `;
}

// Enhanced message formatting function
// Sanitize HTML to prevent XSS from backend responses
function sanitizeHtml(html) {
  // Remove script tags and event handlers
  let sanitized = html
    .replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, '')
    .replace(/on\w+\s*=\s*["'][^"']*["']/gi, '')
    .replace(/on\w+\s*=\s*[^\s>]*/gi, '')
    .replace(/javascript:/gi, '')
    .replace(/<iframe/gi, '&lt;iframe')
    .replace(/<object/gi, '&lt;object')
    .replace(/<embed/gi, '&lt;embed');
  
  return sanitized;
}

function formatMessageContent(content) {
  // Convert markdown-style lists to HTML with enhanced styling
  let formatted = content;

  const hasHtmlTags = /<\s*\w+[^>]*>/.test(formatted);
  if (!hasHtmlTags) {
    const cardHtml = tryFormatModernCard(formatted);
    if (cardHtml) {
      debugLog('Modern summary card formatting applied');
      return cardHtml;
    }
  }

  // Check if content already contains HTML tags - sanitize it before returning
  if (hasHtmlTags) {
    debugLog('Content contains HTML, sanitizing before use');
    return sanitizeHtml(formatted);
  }

  // Convert numbered lists with em dash pattern to bullet points
  // Pattern: "1. Title — Date Description"
  formatted = formatted.replace(/^(\d+)\.\s+(.+?)\s*—\s*(.+)$/gm, (match, number, title, description) => {
    return `<div class="list-item"><span class="bullet-point">•</span><div class="item-content"><strong>${title.trim()}</strong> — ${description.trim()}</div></div>`;
  });

  // Handle remaining numbered lists (convert to bullet points)
  formatted = formatted.replace(/^(\d+)\.\s+(.+)$/gm, '<div class="list-item"><span class="bullet-point">•</span><div class="item-content">$2</div></div>');

  // Handle bullet points  
  formatted = formatted.replace(/^[•\-\*]\s+(.+)$/gm, '<div class="list-item">$1</div>');

  // Handle headers
  formatted = formatted.replace(/^###\s+(.+)$/gm, '<h3 class="response-header">$1</h3>');
  formatted = formatted.replace(/^##\s+(.+)$/gm, '<h2 class="response-header">$1</h2>');
  formatted = formatted.replace(/^#\s+(.+)$/gm, '<h1 class="response-header">$1</h1>');

  // Handle bold text
  formatted = formatted.replace(/\*\*(.+?)\*\*/g, '<strong class="highlight">$1</strong>');

  // Handle code blocks
  formatted = formatted.replace(/`([^`]+)`/g, '<code>$1</code>');

  // Convert line breaks to paragraphs for plain text
  if (!formatted.includes('<div') && !formatted.includes('<h') && !formatted.includes('<ul') && !formatted.includes('<ol')) {
    const paragraphs = formatted.split(/\n\s*\n/);
    if (paragraphs.length > 1) {
      formatted = paragraphs.map(p => p.trim()).filter(p => p).map(p => `<p>${p.replace(/\n/g, '<br>')}</p>`).join('');
    } else {
      formatted = `<p>${formatted.replace(/\n/g, '<br>')}</p>`;
    }
  }

  // Wrap in formatted response container
  return `<div class="formatted-response">${formatted}</div>`;
}

let isOnline = navigator.onLine;
window.addEventListener('online', () => {
  isOnline = true;
  debugLog('Connection restored');
});

window.addEventListener('offline', () => {
  isOnline = false;
  debugWarn('Connection lost');
});

// Robust idempotent initialization - prevent duplicate script loading
if (window.__saaiOutlookInjected || document.getElementById('saai-outlook-sidebar')) {
  debugLog('Script already initialized, skipping...');
} else {
  window.__saaiOutlookInjected = true;
  
  let SIDEBAR_WIDTH = 320; // Will be updated from storage
  const SIDEBAR_ID = 'saai-outlook-sidebar';
  const CHAT_AREA_ID = 'saai-chat-area';

  let isInitialized = false;
  let isSidebarOpen = false;
  let sidebarElement = null;

// === CORE INITIALIZATION ===

async function initialize() {
  if (isInitialized || document.getElementById(SIDEBAR_ID)) {
    debugLog('Already initialized, skipping...');
    return;
  }
  
  debugLog('Initializing Sa.AI Inbox Assistant');
  
  try {
    // Load saved sidebar width from storage
    try {
      const { sidebarWidth } = await chrome.storage.local.get(['sidebarWidth']);
      if (sidebarWidth && sidebarWidth >= 320 && sidebarWidth <= 500) {
        SIDEBAR_WIDTH = sidebarWidth;
        debugLog('Loaded saved sidebar width:', SIDEBAR_WIDTH);
      }
    } catch (error) {
      debugLog('Using default sidebar width:', SIDEBAR_WIDTH);
    }
    
    // Wait for Outlook to be fully loaded
    await waitForOutlookReady();
    
    // Set up message listeners
    setupMessageListeners();
    
    // Set up storage change listeners
    setupStorageListeners();
    
    // Set up page visibility change listener
    setupVisibilityListener();
    
    // Check if sidebar should be open from previous session
    const { sidebarOpen } = await chrome.storage.local.get(['sidebarOpen']);
    debugLog('Checking sidebar state from storage:', sidebarOpen);
    if (sidebarOpen) {
      debugLog('Restoring sidebar from previous session');
      await openSidebar();
    } else {
      debugLog('No previous sidebar state found, starting closed');
    }
    
    isInitialized = true;
    debugLog('Initialization complete');
    
    // Initialize credit system
    await initializeCreditSystem();
    
    // Double-check sidebar state after initialization
    await checkAndRestoreSidebarState();
    
    // Additional delayed check to ensure sidebar is properly restored
    setTimeout(async () => {
      debugLog('Delayed sidebar state check...');
      await checkAndRestoreSidebarState();
    }, 1000);
    
  } catch (error) {
    debugError('Initialization failed:', error);
    // Silently handle initialization errors - don't break Outlook
    // Errors from Outlook's own CSP/network issues are harmless
  }
}

// === VOICE MODE ===
let voiceSession = { 
  active: false, 
  segments: [], 
  recognizer: null, 
  mediaStream: null, 
  speaking: false,
  processing: false,  // NEW: Prevents recording restart during n8n processing
  awaitingConfirmation: false,
  pendingCommand: null,
  errorCount: 0
};
let voiceSpeed = 1.0; // Default speed
let voiceHeartbeat = null; // Keep-alive mechanism

function setupVoiceModeControls(sidebar) {
  const voiceBtn = sidebar.querySelector('#voice-btn');
  const exitBtn = sidebar.querySelector('#voice-exit-btn');
  const indicator = sidebar.querySelector('#voice-indicator');
  const speedBtn125 = sidebar.querySelector('#voice-speed-125');
  const speedBtn150 = sidebar.querySelector('#voice-speed-150');
  const speedBtn100 = sidebar.querySelector('#voice-speed-100');
  
  if (!voiceBtn || !exitBtn || !indicator) return;

  voiceBtn.addEventListener('click', async () => {
    if (!voiceSession.active) {
      await startVoiceMode(indicator, exitBtn, voiceBtn);
    } else {
      await stopListening(indicator, voiceBtn);
    }
  });

  exitBtn.addEventListener('click', async () => {
    await exitVoiceMode(indicator, exitBtn, voiceBtn);
  });

  // Speed control buttons
  if (speedBtn100) {
    speedBtn100.addEventListener('click', () => setVoiceSpeed(1.0, speedBtn100, speedBtn125, speedBtn150));
  }
  if (speedBtn125) {
    speedBtn125.addEventListener('click', () => setVoiceSpeed(1.25, speedBtn100, speedBtn125, speedBtn150));
  }
  if (speedBtn150) {
    speedBtn150.addEventListener('click', () => setVoiceSpeed(1.5, speedBtn100, speedBtn125, speedBtn150));
  }
}

async function startVoiceMode(indicator, exitBtn, voiceBtn) {
  voiceSession.active = true;
  voiceSession.segments = [];
  voiceSession.errorCount = 0; // Reset error count
  voiceSession.processing = false; // Reset processing state
  voiceSession.awaitingConfirmation = false; // Reset confirmation state
  voiceSession.pendingCommand = null; // Clear any pending commands
  
  // Show voice indicator and exit button
  indicator.style.display = 'flex';
  exitBtn.style.display = 'flex';
  
  // Hide voice and send buttons
  voiceBtn.style.display = 'none';
  const sendBtn = document.querySelector('#send-btn');
  if (sendBtn) sendBtn.style.display = 'none';
  
  // Load saved speed preference
  await loadVoiceSpeedPreference();
  
  // Start heartbeat to keep session alive
  startVoiceHeartbeat();
  
  await startListening(indicator);
}

function startVoiceHeartbeat() {
  // Clear any existing heartbeat
  if (voiceHeartbeat) {
    clearInterval(voiceHeartbeat);
  }
  
  // Conservative heartbeat - only restart if truly lost
  voiceHeartbeat = setInterval(() => {
    if (voiceSession.active && !voiceSession.speaking && !voiceSession.processing && !voiceSession.recognizer && !voiceSession.awaitingConfirmation) {
      const indicator = document.getElementById('voice-indicator');
      if (indicator) {
        debugLog('[Voice] Heartbeat: Recognition seems lost, restarting');
        startListening(indicator);
      }
    } else if (!voiceSession.active) {
      // Clean up heartbeat when voice mode is inactive
      clearInterval(voiceHeartbeat);
      voiceHeartbeat = null;
    }
  }, 10000); // Increased to 10 seconds, much less aggressive
}

async function loadVoiceSpeedPreference() {
  try {
    const { voiceSpeed: savedSpeed } = await chrome.storage.local.get(['voiceSpeed']);
    if (savedSpeed && [1.0, 1.25, 1.5].includes(savedSpeed)) {
      voiceSpeed = savedSpeed;
      
      // Update button states
      const btn100 = document.querySelector('#voice-speed-100');
      const btn125 = document.querySelector('#voice-speed-125');
      const btn150 = document.querySelector('#voice-speed-150');
      
      [btn100, btn125, btn150].forEach(btn => {
        if (btn) btn.classList.remove('active');
      });
      
      if (savedSpeed === 1.0 && btn100) btn100.classList.add('active');
      if (savedSpeed === 1.25 && btn125) btn125.classList.add('active');
      if (savedSpeed === 1.5 && btn150) btn150.classList.add('active');
    }
  } catch (e) {
    debugWarn('[Voice] Failed to load speed preference:', e);
  }
}

async function exitVoiceMode(indicator, exitBtn, voiceBtn) {
  voiceSession.active = false;
  voiceSession.processing = false; // Clear processing state when exiting
  
  // Stop heartbeat
  if (voiceHeartbeat) {
    clearInterval(voiceHeartbeat);
    voiceHeartbeat = null;
  }
  
  await stopListening(indicator, voiceBtn);
  speechSynthesis.cancel();
  
  // Render transcript with proper formatting (same as regular chat)
  const chatArea = document.getElementById(CHAT_AREA_ID);
  if (chatArea && voiceSession.segments.length) {
    voiceSession.segments.forEach(seg => {
      if (seg.role === 'user') {
        // User messages - plain text
        appendMessage('user', seg.text, chatArea);
      } else {
        // Bot messages - apply formatting like regular chat
        const formattedText = formatMessageContent(seg.text);
        appendMessage('bot', formattedText, chatArea, false, true);
      }
    });
    chatArea.scrollTop = chatArea.scrollHeight;
  }
  
  // Hide exit button and restore voice/send buttons
  if (exitBtn) exitBtn.style.display = 'none';
  if (voiceBtn) voiceBtn.style.display = 'flex';
  const sendBtn = document.querySelector('#send-btn');
  if (sendBtn) sendBtn.style.display = 'flex';
  if (indicator) indicator.style.display = 'none';
}

async function startListening(indicator) {
  // Clean up any existing recognizer first
  if (voiceSession.recognizer) {
    try {
      voiceSession.recognizer.stop();
    } catch (e) {}
    voiceSession.recognizer = null;
  }

  const Rec = window.SpeechRecognition || window.webkitSpeechRecognition;
  if (!Rec) {
    updateVoiceIndicator('error', 'Speech recognition not supported');
    speak('Speech recognition is not supported in this browser');
    return;
  }

  // Test microphone access first
  try {
    await testMicrophoneAccess();
  } catch (e) {
    debugError('[Voice] Microphone test failed:', e);
    updateVoiceIndicator('error', 'Microphone access failed');
    speak('Unable to access microphone. Please check permissions.');
    return;
  }

    const rec = new Rec();
    voiceSession.recognizer = rec;
  
  // Enhanced configuration for better accuracy
  rec.continuous = true;  // Keep listening for better flow
  rec.interimResults = true;  // Get partial results for feedback
  rec.maxAlternatives = 3;  // Get multiple interpretations
    rec.lang = 'en-US';
  
    rec.onresult = async (e) => {
    try {
      if (e.results.length > 0) {
        const result = e.results[0];
        if (result.isFinal) {
          // Get all alternatives and find the best one
          const alternatives = [];
          for (let i = 0; i < Math.min(result.length, 3); i++) {
            alternatives.push({
              transcript: result[i].transcript.trim(),
              confidence: result[i].confidence || 0
            });
          }
          
          debugLog('[Voice] Alternatives:', alternatives);
          
          // Apply smart confidence filtering
          const bestResult = getBestTranscript(alternatives);
          
          if (bestResult) {
            debugLog('[Voice] Selected:', bestResult.transcript, 'confidence:', bestResult.confidence);
            
            // Reset error count on successful recognition
            voiceSession.errorCount = 0;
            
            // Apply common corrections
            const correctedTranscript = applySpeechCorrections(bestResult.transcript);
            
            // Stop current recognition before processing
            if (voiceSession.recognizer) {
              try {
                voiceSession.recognizer.stop();
              } catch (e) {
                debugWarn('[Voice] Error stopping recognizer:', e);
              }
            }
            voiceSession.recognizer = null;
            await handleVoiceTurn(correctedTranscript, indicator, bestResult.confidence);
          } else {
            debugLog('[Voice] All alternatives below confidence threshold');
            speak('I didn\'t catch that clearly. Could you repeat?');
            // Don't stop recognition, keep listening
          }
        } else {
          // Show interim results for user feedback
          const interim = result[0].transcript;
          if (interim.length > 3) {
            updateVoiceIndicator('listening', `Hearing: "${interim}..."`);
          }
        }
      }
    } catch (error) {
      debugError('[Voice] Error processing speech result:', error);
      speak('Sorry, there was an error processing your speech');
    }
  };
  
  rec.onerror = (e) => {
    debugError('[Voice] Recognition error:', e.error);
    voiceSession.recognizer = null;
    
    // Handle different error types with appropriate user feedback
    switch (e.error) {
      case 'not-allowed':
        updateVoiceIndicator('error', 'Microphone access denied');
        speak('Microphone access was denied. Please allow microphone access and try again.');
        return;
        
      case 'no-speech':
        debugLog('[Voice] No speech detected, continuing...');
        // This is normal, just restart
        break;
        
      case 'audio-capture':
        updateVoiceIndicator('error', 'Audio capture failed');
        speak('Audio capture failed. Please check your microphone connection.');
        break;
        
      case 'network':
        updateVoiceIndicator('error', 'Network error');
        speak('Network error occurred. Retrying in a moment.');
        break;
        
      case 'service-not-allowed':
        updateVoiceIndicator('error', 'Speech service not allowed');
        speak('Speech recognition service is not available. Please try refreshing the page.');
        return;
        
      case 'bad-grammar':
        debugLog('[Voice] Grammar error, continuing...');
        break;
        
      default:
        debugWarn('[Voice] Unknown error:', e.error);
        updateVoiceIndicator('error', `Error: ${e.error}`);
        // Don't speak for unknown errors to prevent loops
        debugLog('[Voice] Suppressing speech for unknown error to prevent loops');
    }
    
    // For recoverable errors, restart after a delay (but limit retries)
    if (voiceSession.active && !voiceSession.speaking && !voiceSession.processing) {
      // Track error count to prevent infinite loops
      voiceSession.errorCount = (voiceSession.errorCount || 0) + 1;
      
      if (voiceSession.errorCount > 3) {
        debugError('[Voice] Too many errors, stopping voice mode');
        updateVoiceIndicator('error', 'Too many errors - voice mode stopped');
        speak('Too many recognition errors occurred. Please try restarting voice mode.');
        voiceSession.active = false;
        return;
      }
      
      const delay = e.error === 'network' ? 3000 : 1500; // Longer delays
      setTimeout(() => {
        if (voiceSession.active && !voiceSession.speaking && !voiceSession.processing) {
          debugLog('[Voice] Restarting after error:', e.error, 'attempt:', voiceSession.errorCount);
          startListening(indicator);
        }
      }, delay);
    }
  };
  
  rec.onend = () => {
    debugLog('[Voice] Recognition ended, active:', voiceSession.active, 'speaking:', voiceSession.speaking);
    voiceSession.recognizer = null;
    
    // DON'T auto-restart here - let the speech completion handle restart
    // This prevents the continuous listening loop
    debugLog('[Voice] Not auto-restarting from onend - waiting for speech completion');
  };
  
  rec.onstart = () => {
    debugLog('[Voice] Recognition started successfully');
    updateVoiceIndicator('listening', 'Listening… (say something)');
  };
  
  try {
    rec.start();
    debugLog('[Voice] Starting new recognition session');
  } catch (err) {
    debugError('[Voice] Failed to start recognition:', err);
    voiceSession.recognizer = null;
    updateVoiceIndicator('error', 'Failed to start listening');
  }
}

async function stopListening(indicator, voiceBtn) {
  debugLog('[Voice] Stopping listening...');
  
  if (voiceSession.recognizer) {
    try { 
      voiceSession.recognizer.onend = null; 
      voiceSession.recognizer.onerror = null;
      voiceSession.recognizer.onresult = null;
      voiceSession.recognizer.onstart = null;
      voiceSession.recognizer.stop(); 
    } catch (err) {
      debugWarn('[Voice] Error stopping recognizer:', err);
    }
    voiceSession.recognizer = null;
  }
  
  if (voiceSession.mediaStream) {
    voiceSession.mediaStream.getTracks().forEach(t => t.stop());
    voiceSession.mediaStream = null;
  }
  
  if (indicator) indicator.style.display = 'none';
}

async function handleVoiceTurn(userText, indicator, confidence = 1) {
  debugLog('[Voice] Processing:', userText, 'confidence:', confidence);
  
  // Check for voice commands first
  if (handleVoiceCommands(userText)) {
    return;
  }
  
  // Check if we need confirmation for this command
  if (shouldConfirmCommand(userText, confidence)) {
    await handleCommandConfirmation(userText, indicator, confidence);
    return;
  }
  
  voiceSession.segments.push({ role: 'user', text: userText });
  
  // Get thread context
  const { isOnThreadPage, threadId, subjectLine } = getThreadContext();
  const userId = await getOrGenerateUserId();
  
  // Check for draft/reply requests
  const isDraftRequest = isDraftOrReplyRequest(userText);
  const isSummarizeThreadRequest = /summari[sz]e\s+(this\s+)?thread/i.test(userText);
  
  // Validation 1: Draft/Reply requests require an open thread
  if (isDraftRequest && !isOnThreadPage) {
    voiceSession.segments.push({ 
      role: 'assistant', 
      text: 'Open a thread for this process and then prompt again.' 
    });
    await speak('Open a thread for this process and then prompt again.');
    return;
  }
  
  // Validation 2: Thread summarization requests require an open thread
  if (isSummarizeThreadRequest && !isOnThreadPage) {
    voiceSession.segments.push({ 
      role: 'assistant', 
      text: 'Open a thread for this process and then prompt again.' 
    });
    await speak('Open a thread for this process and then prompt again.');
    return;
  }
  // Validation 3: Summarize requires a valid threadId (safeguard: ID may have failed validation)
  if (isSummarizeThreadRequest && isOnThreadPage && !threadId) {
    const msg = 'We couldn\'t read the message ID for this thread. Try opening the email in a new tab and try again.';
    voiceSession.segments.push({ role: 'assistant', text: msg });
    await speak(msg);
    return;
  }
  
  voiceSession.processing = true;  // NEW: Set processing state
  updateVoiceIndicator('processing', 'Processing…');
  
  // Simplified payload for faster processing
  const payload = {
    query: userText,
    userId: userId,
    context: 'OutlookChat',
    voiceMode: true
  };
  
  // When a thread is open: send threadId and subjectLine for both draft and summarize
  if (isOnThreadPage) {
    if (threadId) payload.threadId = threadId;
    payload.subjectLine = subjectLine || (document.title && document.title.split(/\s*[-|]\s*/)[0].trim()) || '';
    if (payload.subjectLine || payload.threadId) debugLog('[Voice] Including thread context:', { threadId: threadId || 'none', subjectLine: payload.subjectLine || 'none' });
  }
  if (isSummarizeThreadRequest && isOnThreadPage) {
    payload.action = 'summarize_thread';
  }
  
  let replyText = 'I understand.';
  
  try {
    debugLog('[Voice] Sending to n8n...');
    const startTime = Date.now();
    
    const resp = await chrome.runtime.sendMessage({ 
      action: 'sendToN8N', 
      data: { endpoint: 'chat', payload } 
    });
    
    const endTime = Date.now();
    debugLog('[Voice] Response time:', endTime - startTime, 'ms');
    debugLog('[Voice] Full response:', JSON.stringify(resp, null, 2));
    
    if (resp && resp.success) {
      const data = resp.data;
      debugLog('[Voice] Response data:', JSON.stringify(data, null, 2));
      
      const extractedText = extractVoiceReplyText(data);
      debugLog('[Voice] Extracted text:', extractedText);
      
              if (extractedText && extractedText.trim()) {
          // Apply formatting to voice responses as well
          replyText = extractedText; // keep it raw for voice
        } else {
        debugWarn('[Voice] No valid text extracted, using fallback');
        replyText = 'I received your request but couldn\'t extract a proper response.';
      }
    } else {
      debugError('[Voice] Request failed:', resp);
      replyText = 'Sorry, I had trouble with that request.';
    }
  } catch (e) {
    debugError('[Voice] Error:', e);
    replyText = 'Network issue. Please try again.';
    voiceSession.processing = false;  // NEW: Clear processing state on error
  }
  
  debugLog('[Voice] Speaking:', replyText);
  voiceSession.segments.push({ role: 'assistant', text: replyText });
  voiceSession.processing = false;  // NEW: Clear processing state before speaking
  await speak(replyText);
}

function extractVoiceReplyText(data) {
  debugLog('[Voice] extractVoiceReplyText input:', typeof data, data);
  
  try {
    if (!data) {
      debugLog('[Voice] No data provided');
      return '';
    }
    
    // If server wrapped in array, use first item
    if (Array.isArray(data) && data.length > 0) {
      debugLog('[Voice] Data is array, using first item');
      data = data[0];
    }
    
    // If data is a string, try to parse as JSON or return as-is
    if (typeof data === 'string') {
      debugLog('[Voice] Data is string:', data);
      // Sometimes n8n returns a JSON string
      if (data.trim().startsWith('{') || data.trim().startsWith('[')) {
        try { 
          data = JSON.parse(data);
          debugLog('[Voice] Parsed JSON string:', data);
        } catch { 
          debugLog('[Voice] Failed to parse JSON, returning string as-is');
          return data.trim(); 
        }
      } else {
        return data.trim();
      }
    }
    
    // Check common response fields
    const fields = ['message', 'replyText', 'reply', 'response', 'text', 'content', 'summary'];
    
    for (const field of fields) {
      if (data[field]) {
        debugLog('[Voice] Found field:', field, typeof data[field]);
        
        if (typeof data[field] === 'string' && data[field].trim()) {
          debugLog('[Voice] Returning string field:', field);
          return data[field].trim();
        }
        
        if (typeof data[field] === 'object') {
          debugLog('[Voice] Processing object field:', field);
          const extracted = extractVoiceReplyText(data[field]);
          if (extracted) return extracted;
        }
      }
    }
    
    // Handle reply field specially (legacy support)
    if (data.reply) {
      let r = data.reply;
      debugLog('[Voice] Processing reply field:', typeof r);
      
      if (typeof r === 'string') {
        // If reply contains JSON, parse it
        if (r.trim().startsWith('{') || r.trim().startsWith('[')) {
          try { 
            r = JSON.parse(r); 
            debugLog('[Voice] Parsed reply JSON:', r);
          } catch {
            debugLog('[Voice] Failed to parse reply JSON, using string');
            return r.trim();
          }
        } else {
          return r.trim();
        }
      }
      
      // If reply is an array of result objects
      if (Array.isArray(r)) {
        debugLog('[Voice] Reply is array with', r.length, 'items');
        const parts = [];
        for (const item of r) {
          if (!item) continue;
          if (typeof item.summary === 'string' && item.summary.trim()) parts.push(item.summary.trim());
          else if (typeof item.message === 'string' && item.message.trim()) parts.push(item.message.trim());
          else if (typeof item.text === 'string' && item.text.trim()) parts.push(item.text.trim());
        }
        if (parts.length) {
          debugLog('[Voice] Extracted from reply array:', parts.join(' '));
          return parts.join(' ');
      }
      }
      
      // If reply is an object
      if (r && typeof r === 'object') {
        debugLog('[Voice] Reply is object, recursing');
        const extracted = extractVoiceReplyText(r);
        if (extracted) return extracted;
      }
    }
    
    // Check if it's a structured email summary
    const keys = ['high_priority_emails','medium_priority','low_priority','already_replied_closed_threads','missed_or_ignored_emails'];
    let found = false; const parts = [];
    for (const k of keys) {
      if (Array.isArray(data[k])) { 
        found = true; 
        parts.push(`${data[k].length} in ${k.replace(/_/g,' ')}`); 
      }
    }
    if (found) {
      const summary = `Summary prepared: ${parts.join(', ')}. Say show chat to view details.`;
      debugLog('[Voice] Generated email summary - length:', summary.length);
      return summary;
    }
    
    debugLog('[Voice] No extractable text found in data');
    return '';
    
  } catch (e) {
    debugError('[Voice] extractVoiceReplyText failed:', e);
  return '';
  }
}

async function speak(text) {
  voiceSession.speaking = true;
  updateVoiceIndicator('speaking', 'Speaking…');
  
  try {
    // Cancel any ongoing speech
    speechSynthesis.cancel();
    
    // Wait for cancellation to complete
    await new Promise(resolve => setTimeout(resolve, 100));
    
    const utter = new SpeechSynthesisUtterance(text);
    
    // Enhanced voice settings
    const voices = speechSynthesis.getVoices();
    const preferredVoice = voices.find(v => 
      v.lang.startsWith('en') && v.name.includes('Google')
    ) || voices.find(v => v.lang.startsWith('en'));
    
    if (preferredVoice) {
      utter.voice = preferredVoice;
    }
    
    utter.rate = voiceSpeed;
    utter.pitch = 1.0;
    utter.volume = 0.8;
    
    // Add pauses for better comprehension
    const processedText = text
      .replace(/\./g, '. ')
      .replace(/,/g, ', ')
      .replace(/:/g, ': ');
    utter.text = processedText;
    
    utter.onend = () => { 
      debugLog('[Voice] Speech finished, voiceSession.active:', voiceSession.active);
      voiceSession.speaking = false;
      
      // Only restart if voice mode is active, not processing, and we don't already have a recognizer
      if (voiceSession.active && !voiceSession.processing && !voiceSession.recognizer) {
        debugLog('[Voice] Scheduling recognition restart after speech...');
        setTimeout(() => {
          debugLog('[Voice] Speech timeout - active:', voiceSession.active, 'speaking:', voiceSession.speaking, 'processing:', voiceSession.processing, 'recognizer:', !!voiceSession.recognizer);
          if (voiceSession.active && !voiceSession.speaking && !voiceSession.processing && !voiceSession.recognizer) {
            const indicator = document.getElementById('voice-indicator');
            if (indicator) {
              debugLog('[Voice] ✅ Restarting recognition after speech');
              startListening(indicator);
            } else {
              debugLog('[Voice] ❌ No indicator found for restart');
            }
          } else {
            debugLog('[Voice] ❌ Not restarting - conditions not met');
          }
        }, 1500); // Longer delay to ensure speech is fully finished
      } else {
        debugLog('[Voice] ❌ Not scheduling restart - active:', voiceSession.active, 'recognizer exists:', !!voiceSession.recognizer);
      }
    };
    
    utter.onerror = (e) => {
      debugError('[Voice] Speech error:', e);
      voiceSession.speaking = false;
      updateVoiceIndicator('error', 'Speech error');
    };
    
    speechSynthesis.speak(utter);
  } catch (e) {
    debugError('[Voice] Speak error:', e);
    voiceSession.speaking = false;
    updateVoiceIndicator('error', 'Speech error');
  }
}

async function setVoiceSpeed(speed, btn100, btn125, btn150) {
  voiceSpeed = speed;
  
  // Update button states
  [btn100, btn125, btn150].forEach(btn => {
    if (btn) btn.classList.remove('active');
  });
  
  if (speed === 1.0 && btn100) btn100.classList.add('active');
  if (speed === 1.25 && btn125) btn125.classList.add('active');
  if (speed === 1.5 && btn150) btn150.classList.add('active');
  
  // Save preference
  try {
    await chrome.storage.local.set({ voiceSpeed: speed });
  } catch (e) {
    debugWarn('[Voice] Failed to save speed preference:', e);
  }
  
  // Test the new speed if voice mode is active
  if (voiceSession.active) {
    speak('Speed updated');
  }
}

// Enhanced UI feedback function
function updateVoiceIndicator(state, message) {
  const indicator = document.getElementById('voice-indicator');
  if (!indicator) return;
  
  const label = indicator.querySelector('.label');
  const bars = indicator.querySelectorAll('.bars i');
  
  switch (state) {
    case 'listening':
      label.textContent = message || 'Listening…';
      bars.forEach(bar => bar.style.animation = 'saai-bars 1s infinite ease-in-out');
      bars.forEach(bar => bar.style.background = '#0f172a');
      break;
    case 'processing':
      label.textContent = message || 'Processing…';
      bars.forEach(bar => bar.style.animation = 'none');
      bars.forEach(bar => bar.style.background = '#f59e0b');
      break;
    case 'speaking':
      label.textContent = message || 'Speaking…';
      bars.forEach(bar => bar.style.animation = 'none');
      bars.forEach(bar => bar.style.background = '#3b82f6');
      break;
    case 'error':
      label.textContent = message || 'Error occurred';
      bars.forEach(bar => bar.style.animation = 'none');
      bars.forEach(bar => bar.style.background = '#ef4444');
      break;
  }
}

// === VOICE QUALITY IMPROVEMENTS ===

async function testMicrophoneAccess() {
  return new Promise(async (resolve, reject) => {
    try {
      // Simple microphone access test - just check if we can get the stream
      const stream = await navigator.mediaDevices.getUserMedia({
        audio: {
          echoCancellation: true,
          noiseSuppression: true,
          autoGainControl: true
        }
      });
      
      debugLog('[Voice] Microphone access granted');
      
      // Clean up immediately - don't test levels as it's causing issues
      stream.getTracks().forEach(track => track.stop());
      
      resolve();
      
    } catch (error) {
      debugError('[Voice] Microphone access failed:', error);
      reject(error);
    }
  });
}

function getBestTranscript(alternatives) {
  if (!alternatives || alternatives.length === 0) return null;
  
  // Filter by confidence and length
  const validAlternatives = alternatives.filter(alt => {
    const transcript = alt.transcript.trim();
    if (!transcript) return false;
    
    // Dynamic confidence threshold based on length
    const minConfidence = transcript.length > 10 ? 0.6 : 0.75;
    return alt.confidence >= minConfidence;
  });
  
  if (validAlternatives.length === 0) return null;
  
  // Sort by confidence and return best
  return validAlternatives.sort((a, b) => b.confidence - a.confidence)[0];
}

function applySpeechCorrections(transcript) {
  if (!transcript) return transcript;
  
  // Common misheard words in email/assistant context
  const corrections = {
    // Email related
    'e-mail': 'email',
    'gmail': 'email', 
    'male': 'email',
    'mail': 'email',
    
    // Actions
    'some arise': 'summarize',
    'sum arise': 'summarize',
    'summary': 'summarize',
    'some rice': 'summarize',
    
    // Thread related
    'red': 'thread',
    'bread': 'thread', 
    'threat': 'thread',
    'fred': 'thread',
    
    // Common commands
    'show me': 'show',
    'tell me': 'tell',
    'open up': 'open',
    'close down': 'close',
    
    // Technical terms
    'in box': 'inbox',
    'chat bot': 'chatbot',
    'ai assistant': 'AI assistant'
  };
  
  let corrected = transcript.toLowerCase();
  
  // Apply corrections
  for (const [wrong, right] of Object.entries(corrections)) {
    const regex = new RegExp(`\\b${wrong}\\b`, 'gi');
    corrected = corrected.replace(regex, right);
  }
  
  // Clean up extra spaces
  corrected = corrected.replace(/\s+/g, ' ').trim();
  
  if (corrected !== transcript.toLowerCase()) {
    debugLog('[Voice] Applied correction:', transcript, '->', corrected);
  }
  
  return corrected;
}

function shouldConfirmCommand(transcript, confidence) {
  // Confirm if confidence is low or command seems critical
  if (confidence < 0.7) return true;
  
  // Critical commands that should be confirmed
  const criticalCommands = [
    'delete', 'remove', 'clear', 'exit', 'stop', 'close',
    'send email', 'reply', 'forward', 'archive'
  ];
  
  const lowerTranscript = transcript.toLowerCase();
  return criticalCommands.some(cmd => lowerTranscript.includes(cmd));
}

async function handleCommandConfirmation(userText, indicator, confidence) {
  debugLog('[Voice] Requesting confirmation for:', userText);
  
  // Store the pending command
  voiceSession.pendingCommand = { text: userText, confidence };
  
  if (confidence < 0.7) {
    speak(`I heard "${userText}" but I'm not completely sure. Say "yes" to confirm or repeat your command.`);
  } else {
    speak(`Did you say "${userText}"? Say "yes" to confirm or "no" to cancel.`);
  }
  
  // Set flag to handle next input as confirmation
  voiceSession.awaitingConfirmation = true;
}

// Enhanced voice command handler that includes confirmation responses
function handleVoiceCommands(text) {
  const lowerText = text.toLowerCase().trim();
  
  // Handle confirmation responses
  if (voiceSession.awaitingConfirmation) {
    voiceSession.awaitingConfirmation = false;
    
    if (lowerText.includes('yes') || lowerText.includes('confirm') || lowerText.includes('correct')) {
      const pendingCommand = voiceSession.pendingCommand;
      if (pendingCommand) {
        speak('Confirmed. Processing your request.');
        voiceSession.pendingCommand = null;
        // Process the original command
        setTimeout(() => {
          const indicator = document.getElementById('voice-indicator');
          handleVoiceTurn(pendingCommand.text, indicator, 1.0); // Set confidence to 1.0 since confirmed
        }, 500);
        return true;
      }
    } else if (lowerText.includes('no') || lowerText.includes('cancel') || lowerText.includes('wrong')) {
      speak('Cancelled. What would you like me to do?');
      voiceSession.pendingCommand = null;
      return true;
    } else {
      // Treat as new command if not yes/no
      voiceSession.pendingCommand = null;
      speak('Let me process that as a new command.');
      return false; // Let it be processed as normal command
    }
  }
  
  // Original voice commands...
  if (lowerText.includes('stop') || lowerText.includes('exit voice')) {
    const indicator = document.getElementById('voice-indicator');
    const exitBtn = document.querySelector('#voice-exit-btn');
    const voiceBtn = document.querySelector('#voice-btn');
    exitVoiceMode(indicator, exitBtn, voiceBtn);
    return true;
  }
  
  if (lowerText.includes('repeat') || lowerText.includes('say again')) {
    const lastResponse = voiceSession.segments
      .filter(s => s.role === 'assistant')
      .pop();
    if (lastResponse) {
      speak(lastResponse.text);
      return true;
    }
  }
  
  if (lowerText.includes('clear chat') || lowerText.includes('start over')) {
    voiceSession.segments = [];
    speak('Chat cleared. How can I help you?');
    return true;
  }
  
  if (lowerText.includes('slower') || lowerText.includes('slow down')) {
    if (voiceSpeed > 1.0) {
      const newSpeed = voiceSpeed === 1.5 ? 1.25 : 1.0;
      const btn100 = document.querySelector('#voice-speed-100');
      const btn125 = document.querySelector('#voice-speed-125');
      const btn150 = document.querySelector('#voice-speed-150');
      setVoiceSpeed(newSpeed, btn100, btn125, btn150);
      speak('Speed reduced');
      return true;
    }
  }
  
  if (lowerText.includes('faster') || lowerText.includes('speed up')) {
    if (voiceSpeed < 1.5) {
      const newSpeed = voiceSpeed === 1.0 ? 1.25 : 1.5;
      const btn100 = document.querySelector('#voice-speed-100');
      const btn125 = document.querySelector('#voice-speed-125');
      const btn150 = document.querySelector('#voice-speed-150');
      setVoiceSpeed(newSpeed, btn100, btn125, btn150);
      speak('Speed increased');
      return true;
    }
  }
  
  return false;
}
async function checkAndRestoreSidebarState() {
  try {
    const { sidebarOpen } = await chrome.storage.local.get(['sidebarOpen']);
    debugLog('Post-initialization sidebar state check:', sidebarOpen);
    
    if (sidebarOpen && !isSidebarOpen) {
      debugLog('Sidebar should be open but isn\'t, restoring...');
      await openSidebar();
    } else if (!sidebarOpen && isSidebarOpen) {
      debugLog('Sidebar should be closed but isn\'t, closing...');
      await closeSidebar();
    } else {
      debugLog('Sidebar state is consistent');
    }
    
    // Additional check: if sidebar should be open but element doesn't exist
    if (sidebarOpen && !document.getElementById(SIDEBAR_ID)) {
      debugLog('Sidebar should be open but element missing, recreating...');
      await openSidebar();
    }
  } catch (error) {
    debugError('Error checking sidebar state:', error);
  }
}

async function waitForOutlookReady() {
  return new Promise((resolve) => {
    const checkOutlook = () => {
      // Check if Outlook's main containers are present
      // Outlook web uses various selectors - we'll try multiple common patterns
      const outlookContainer = document.querySelector(
        '[role="main"], ' +
        '.ms-Fabric, ' +
        '[data-app-section="Mail"], ' +
        '#app, ' +
        '.owaContent, ' +
        '[id*="MailModule"], ' +
        '[class*="mailContainer"]'
      );
      if (outlookContainer) {
        resolve();
      } else {
        setTimeout(checkOutlook, 100);
      }
    };
    checkOutlook();
  });
}

function setupMessageListeners() {
  chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
    debugLog('Message received:', request);
    
    switch (request.action) {
      case 'ping':
        sendResponse({ status: 'ready' });
        break;
        
      case 'open_saai':
        toggleSidebar().then(() => {
          sendResponse({ success: true });
        }).catch(error => {
          sendResponse({ success: false, error: error.message });
        });
        return true; // Keep message channel open for async response
        
      case 'checkInitialization':
        sendResponse({ initialized: isInitialized });
        break;
    }
  });
}

function setupStorageListeners() {
  chrome.storage.onChanged.addListener((changes, namespace) => {
    if (namespace === 'local') {
      if (changes.isConnected) {
        debugLog('Connection status changed:', changes.isConnected.newValue);
        if (isSidebarOpen) {
          updateSidebarContent();
        }
      }
    }
  });
}

function setupVisibilityListener() {
  document.addEventListener('visibilitychange', async () => {
    if (!document.hidden) {
      debugLog('Page became visible, checking sidebar state...');
      await checkAndRestoreSidebarState();
    }
  });
}

// === CHAT HISTORY MANAGEMENT ===

function saveChatHistory() {
  try {
    const chatArea = document.getElementById(CHAT_AREA_ID);
    if (!chatArea) return;
    
    const messages = Array.from(chatArea.querySelectorAll('.message')).map(messageDiv => {
      const isUser = messageDiv.classList.contains('user-message');
      const content = messageDiv.querySelector('.message-content');
      return {
        type: isUser ? 'user' : 'bot',
        content: content ? content.innerHTML : '',
        timestamp: Date.now()
      };
    });
    
    chrome.storage.local.set({ chatHistory: messages });
    debugLog('Chat history saved:', messages.length, 'messages');
  } catch (error) {
    debugError('Error saving chat history:', error);
  }
}

function loadChatHistory() {
  try {
    const chatArea = document.getElementById(CHAT_AREA_ID);
    if (!chatArea) return;
    
    chrome.storage.local.get(['chatHistory'], (result) => {
      if (result.chatHistory && result.chatHistory.length > 0) {
        debugLog('Loading chat history:', result.chatHistory.length, 'messages');
        
        // Clear existing content
        chatArea.innerHTML = '';
        
        // Load each message
        result.chatHistory.forEach(messageData => {
          const messageDiv = document.createElement('div');
          messageDiv.className = `message ${messageData.type}-message`;
          messageDiv.innerHTML = `
            <div class="message-content">
              ${messageData.content}
            </div>
          `;
          chatArea.appendChild(messageDiv);
        });
        
        // Scroll to bottom
        chatArea.scrollTop = chatArea.scrollHeight;
        
        debugLog('Chat history loaded successfully');
      }
    });
  } catch (error) {
    debugError('Error loading chat history:', error);
  }
}

function clearChatHistory() {
  try {
    chrome.storage.local.remove(['chatHistory']);
    debugLog('Chat history cleared');
  } catch (error) {
    debugError('Error clearing chat history:', error);
  }
}

// === SIDEBAR MANAGEMENT ===

async function toggleSidebar() {
  debugLog('Toggle sidebar called, current state:', isSidebarOpen);
  
  if (isSidebarOpen) {
    await closeSidebar();
  } else {
    await openSidebar();
  }
}

async function openSidebar() {
  if (isSidebarOpen) return;
  
  debugLog('Opening sidebar');
  
  try {
    // Create and inject sidebar
    await createSidebar();
    
    // Adjust Outlook layout
    adjustOutlookLayout(true);
    
    // Update state
    isSidebarOpen = true;
    await chrome.storage.local.set({ sidebarOpen: true });
    
    debugLog('Sidebar opened successfully, state saved');
    
  } catch (error) {
    debugError('Failed to open sidebar:', error);
    throw error;
  }
}

async function closeSidebar() {
  if (!isSidebarOpen) return;
  
  debugLog('Closing sidebar');
  
  try {
    // Remove sidebar element
    if (sidebarElement) {
      sidebarElement.remove();
      sidebarElement = null;
    }
    
    // Restore Outlook layout
    adjustOutlookLayout(false);
    
    // Clean up flex container if no sidebar is open
    const flexContainer = document.querySelector('.saai-flex-container');
    if (flexContainer) {
      // Find Outlook main wrapper to restore children
      const outlookMainWrapper = document.querySelector(
        '[role="main"], ' +
        '.ms-Fabric, ' +
        '[data-app-section="Mail"], ' +
        '#app, ' +
        '.owaContent, ' +
        '[id*="MailModule"], ' +
        '[class*="mailContainer"]'
      ) || document.body;
      
      if (outlookMainWrapper) {
        while (flexContainer.firstChild) {
          outlookMainWrapper.appendChild(flexContainer.firstChild);
        }
        flexContainer.remove();
      }
    }
    
    // Update state
    isSidebarOpen = false;
    await chrome.storage.local.set({ sidebarOpen: false });
    
    debugLog('Sidebar closed successfully');
    
  } catch (error) {
    debugError('Failed to close sidebar:', error);
    throw error;
  }
}

async function createSidebar() {
  // Check connection status
  const isConnected = await isOutlookConnected();
  
  // Create settings sidebar
  createSettingsSidebar();
  
  // Find Outlook's main wrapper container
  // Try multiple common Outlook selectors
  const outlookMainWrapper = document.querySelector(
    '[role="main"], ' +
    '.ms-Fabric, ' +
    '[data-app-section="Mail"], ' +
    '#app, ' +
    '.owaContent, ' +
    '[id*="MailModule"], ' +
    '[class*="mailContainer"], ' +
    'body > div:first-child'
  );
  
  if (!outlookMainWrapper) {
    debugError('Could not find Outlook main wrapper container');
    // Fallback: use body as container
    const body = document.body;
    if (!body) {
      debugError('Body element not found');
      return;
    }
    
    // Create flex container and inject into body
    let flexContainer = body.querySelector('.saai-flex-container');
    if (!flexContainer) {
      flexContainer = document.createElement('div');
      flexContainer.className = 'saai-flex-container';
      flexContainer.style.cssText = `
        display: flex !important;
        width: 100% !important;
        height: 100vh !important;
        position: relative !important;
        box-sizing: border-box !important;
        margin: 0 !important;
        padding: 0 !important;
        gap: 0 !important;
      `;
      
      // Move body's direct children into flex container
      while (body.firstChild && body.firstChild !== flexContainer) {
        flexContainer.appendChild(body.firstChild);
      }
      body.appendChild(flexContainer);
    }
    
    // Add sidebar to flex container
    sidebarElement = document.createElement('div');
    sidebarElement.id = SIDEBAR_ID;
    sidebarElement.className = 'saai-sidebar';
    sidebarElement.style.width = `${SIDEBAR_WIDTH}px`;
    sidebarElement.style.setProperty('--saai-sidebar-width', `${SIDEBAR_WIDTH}px`);
    
    if (isConnected) {
      sidebarElement.innerHTML = createWelcomePageHTML();
    } else {
      sidebarElement.innerHTML = createConnectPromptHTML();
    }
    
    const resizeHandle = document.createElement('div');
    resizeHandle.className = 'saai-resize-handle';
    resizeHandle.title = 'Drag to resize sidebar';
    sidebarElement.appendChild(resizeHandle);
    
    flexContainer.appendChild(sidebarElement);
    addSidebarEventListeners(sidebarElement, isConnected);
    addResizeFunctionality(sidebarElement, resizeHandle);
    adjustOutlookLayout(true);
    return;
  }
  
  // Check if we already have a flex container
  let flexContainer = outlookMainWrapper.querySelector('.saai-flex-container');
  
  if (!flexContainer) {
    // Create flex container to hold Outlook content and sidebar
    flexContainer = document.createElement('div');
    flexContainer.className = 'saai-flex-container';
    flexContainer.style.cssText = `
      display: flex !important;
      width: 100% !important;
      height: 100vh !important;
      position: relative !important;
      box-sizing: border-box !important;
      margin: 0 !important;
      padding: 0 !important;
      gap: 0 !important;
    `;
    
    // Move all existing children of Outlook main wrapper into the flex container
    while (outlookMainWrapper.firstChild) {
      flexContainer.appendChild(outlookMainWrapper.firstChild);
    }
    
    // Add flex container back to Outlook main wrapper
    outlookMainWrapper.appendChild(flexContainer);
  }
  
  // Create sidebar container
  sidebarElement = document.createElement('div');
  sidebarElement.id = SIDEBAR_ID;
  sidebarElement.className = 'saai-sidebar';
  
  // Apply saved width to sidebar
  sidebarElement.style.width = `${SIDEBAR_WIDTH}px`;
  sidebarElement.style.setProperty('--saai-sidebar-width', `${SIDEBAR_WIDTH}px`);
  
  // Set content based on connection status
  if (isConnected) {
    sidebarElement.innerHTML = createWelcomePageHTML();
  } else {
    sidebarElement.innerHTML = createConnectPromptHTML();
  }
  
  // Add resize handle (always add, regardless of connection status)
  const resizeHandle = document.createElement('div');
  resizeHandle.className = 'saai-resize-handle';
  resizeHandle.title = 'Drag to resize sidebar';
  sidebarElement.appendChild(resizeHandle);
  
  debugLog('Resize handle created and added to sidebar');
  
  // Add sidebar to flex container
  flexContainer.appendChild(sidebarElement);
  
  // Add event listeners
  addSidebarEventListeners(sidebarElement, isConnected);
  
  // Add resize functionality
  addResizeFunctionality(sidebarElement, resizeHandle);
  
  // Add CSS class to body
  document.body.classList.add('saai-sidebar-open');
  
  // Add data attribute to sidebar for easy detection
  sidebarElement.setAttribute('data-saai-provider', 'outlook');
  
  // Store reference to flex container for cleanup
  sidebarElement._saaiFlexContainer = flexContainer;
}

function addResizeFunctionality(sidebar, resizeHandle) {
  let isResizing = false;
  let startX = 0;
  let startWidth = 0;
  const useCapture = true;

  const resize = (e) => {
    if (!isResizing) return;
    const deltaX = startX - e.clientX;
    let newWidth = startWidth + deltaX;
    newWidth = Math.max(320, Math.min(500, newWidth));
    SIDEBAR_WIDTH = newWidth;
    sidebar.style.width = `${newWidth}px`;
    sidebar.style.setProperty('--saai-sidebar-width', `${newWidth}px`);
    adjustOutlookLayout(true);
  };

  const stopResize = () => {
    if (!isResizing) return;
    isResizing = false;
    chrome.storage.local.set({ sidebarWidth: SIDEBAR_WIDTH });
    debugLog('Saved sidebar width:', SIDEBAR_WIDTH);
    document.body.style.userSelect = '';
    document.body.style.cursor = '';
    document.removeEventListener('mousemove', resize, useCapture);
    document.removeEventListener('mouseup', stopResize, useCapture);
    window.removeEventListener('mousemove', resize, useCapture);
    window.removeEventListener('mouseup', stopResize, useCapture);
    window.removeEventListener('blur', stopResize);
  };

  const startResize = (e) => {
    e.preventDefault();
    isResizing = true;
    startX = e.clientX;
    startWidth = SIDEBAR_WIDTH;
    document.body.style.userSelect = 'none';
    document.body.style.cursor = 'col-resize';
    document.addEventListener('mousemove', resize, useCapture);
    document.addEventListener('mouseup', stopResize, useCapture);
    window.addEventListener('mousemove', resize, useCapture);
    window.addEventListener('mouseup', stopResize, useCapture);
    window.addEventListener('blur', stopResize);
  };

  resizeHandle.addEventListener('mousedown', startResize);
  resizeHandle.addEventListener('dragstart', (e) => e.preventDefault());
}

function adjustOutlookLayout(sidebarOpen) {
  debugLog('Adjusting Outlook layout, sidebar open:', sidebarOpen);
  
  // Find the flex container
  const flexContainer = document.querySelector('.saai-flex-container');
  if (!flexContainer) {
    debugLog('No flex container found, layout adjustment not needed');
    return;
  }
  
  // Get all direct children of the flex container (Outlook content + sidebar)
  const flexChildren = Array.from(flexContainer.children);
  
  flexChildren.forEach(child => {
    if (child.id === SIDEBAR_ID) {
      // This is our sidebar - keep it as is
      return;
    }
    
    // This is Outlook content - adjust its flex properties
    if (sidebarOpen) {
      // When sidebar is open, make Outlook content take remaining space
      const currentWidth = SIDEBAR_WIDTH;
      child.style.setProperty('flex', `1 1 calc(100vw - ${currentWidth}px)`, 'important');
      child.style.setProperty('width', `calc(100vw - ${currentWidth}px)`, 'important');
      child.style.setProperty('max-width', `calc(100vw - ${currentWidth}px)`, 'important');
      child.style.setProperty('min-width', `calc(100vw - ${currentWidth}px)`, 'important');
      child.style.setProperty('box-sizing', 'border-box', 'important');
      child.style.setProperty('overflow', 'hidden', 'important');
    } else {
      // When sidebar is closed, restore full width
      child.style.setProperty('flex', '1 1 100vw', 'important');
      child.style.setProperty('width', '100vw', 'important');
      child.style.setProperty('max-width', '100vw', 'important');
      child.style.setProperty('min-width', '100vw', 'important');
      child.style.setProperty('box-sizing', 'border-box', 'important');
      child.style.setProperty('overflow', 'auto', 'important');
    }
  });
  
  // Prevent horizontal scrolling when sidebar is open
  if (sidebarOpen) {
    document.body.style.setProperty('overflow-x', 'hidden', 'important');
    document.documentElement.style.setProperty('overflow-x', 'hidden', 'important');
  } else {
    document.body.style.setProperty('overflow-x', 'auto', 'important');
    document.documentElement.style.setProperty('overflow-x', 'auto', 'important');
  }
}

// === UI COMPONENTS ===

function createChatInterfaceHTML() {
  return `
    <div class="saai-header">
      <span class="saai-title">Sa.AI Assistant</span>
      <div class="saai-header-actions">
        <button id="task-list-btn" class="saai-task-btn">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <polyline points="9,11 12,14 22,4"/>
            <path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11"/>
          </svg>
          Tasks
        </button>
        <button id="clear-chat-btn" class="saai-task-btn" title="Clear chat history">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <path d="M3 6h18"/>
            <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6"/>
            <path d="M8 6V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/>
          </svg>
          Clear
        </button>
        <button id="close-sidebar" class="saai-close-btn" title="Close">×</button>
      </div>
    </div>
    <div id="${CHAT_AREA_ID}" class="chat-area">
      <div class="chat-welcome">
        <div class="chat-welcome-icon">
            <img src="${chrome.runtime.getURL('icons/icon-128.png')}" alt="Sa.AI Logo" style="height: 48px; width: auto;" class="saai-logo-img-small"/>
        </div>
        <h3 class="chat-welcome-title">Hi! I'm your Inbox assistant</h3>
        <p class="chat-welcome-subtitle">How can I help you today?</p>
      </div>
    </div>
    <div class="chat-input-container">
      <div id="voice-indicator" class="saai-voice-indicator" style="display:none">
        <span class="mic">🎙️</span>
        <span class="bars"><i></i><i></i><i></i></span>
        <span class="label">Listening…</span>
        <div class="voice-speed-controls">
          <button id="voice-speed-100" class="voice-speed-btn active" title="Normal speed">1.0x</button>
          <button id="voice-speed-125" class="voice-speed-btn" title="1.25x speed">1.25x</button>
          <button id="voice-speed-150" class="voice-speed-btn" title="1.5x speed">1.5x</button>
        </div>
      </div>
      <div class="chat-input-row">
        <div class="chat-input-wrapper">
          <textarea id="chat-input" placeholder="Ask me anything about your emails..." rows="1"></textarea>
        </div>
        <div class="chat-input-controls">
          <button id="voice-btn" class="saai-action-btn saai-voice-btn" title="Voice mode">
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
              <path d="M12 1a3 3 0 0 0-3 3v8a3 3 0 0 0 6 0V4a3 3 0 0 0-3-3z"/>
              <path d="M19 10v2a7 7 0 0 1-14 0v-2"/>
              <line x1="12" y1="19" x2="12" y2="23"/>
              <line x1="8" y1="23" x2="16" y2="23"/>
            </svg>
          </button>
          <button id="send-btn" class="saai-action-btn saai-send-btn">
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
              <path d="M22 2L11 13"/>
              <path d="M22 2L15 22L11 13L2 9L22 2Z"/>
            </svg>
          </button>
          <button id="voice-exit-btn" class="saai-action-btn saai-voice-exit-btn" title="Show chat transcript" style="display:none">
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
          <path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/>
        </svg>
      </button>
    </div>
      </div>
    </div>
  `;
}

function createConnectPromptHTML() {
  return `
    <div class="saai-header">
      <span class="saai-title">Sa.AI Assistant</span>
      <button id="close-sidebar" class="saai-close-btn" title="Close">×</button>
    </div>
    <div class="saai-connect-content">
      <div class="saai-connect-icon">
          <img src="${chrome.runtime.getURL('icons/icon-128.png')}" alt="Sa.AI Logo" style="height: 32px; width: auto;" class="saai-logo-img-small"/>
      </div>
      
      <h2 class="saai-connect-heading">Welcome to Sa.AI</h2>
      
      <p class="saai-connect-description">
        Your intelligent Inbox assistant is ready to help you manage your inbox more efficiently.
      </p>

      <div class="saai-features">
        <div class="saai-feature-card">
          <div class="saai-feature-icon">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
              <path d="M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z"/>
              <polyline points="22,6 12,13 2,6"/>
            </svg>
          </div>
          <div class="saai-feature-content">
            <h4 class="saai-feature-title">Inbox Summarization</h4>
            <p class="saai-feature-description">Get instant summaries of your important emails</p>
          </div>
        </div>
        
        <div class="saai-feature-card">
          <div class="saai-feature-icon">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
              <path d="M12 3l1.912 5.813a2 2 0 0 0 1.088 1.088L21 12l-5.813 1.912a2 2 0 0 0-1.088 1.088L12 21l-1.912-5.813a2 2 0 0 0-1.088-1.088L3 12l5.813-1.912a2 2 0 0 0 1.088-1.088L12 3z"/>
            </svg>
          </div>
          <div class="saai-feature-content">
            <h4 class="saai-feature-title">Task Extraction</h4>
            <p class="saai-feature-description">Automatically extract tasks and action items</p>
          </div>
        </div>
        
        <div class="saai-feature-card">
          <div class="saai-feature-icon">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
              <polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"/>
            </svg>
          </div>
          <div class="saai-feature-content">
            <h4 class="saai-feature-title">Smart Drafting</h4>
            <p class="saai-feature-description">AI-powered email composition assistance</p>
          </div>
        </div>
      </div>

      <button id="saai-connect-btn" class="saai-connect-button">
        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
          <path d="M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z"/>
          <polyline points="22,6 12,13 2,6"/>
        </svg>
        Connect Outlook Account
      </button>
      
      <p class="saai-disclaimer">
        We'll need access to your Outlook to provide personalized assistance
      </p>
    </div>
  `;
}

function createWelcomePageHTML() {
  return `
    <div class="saai-header">
      <span class="saai-title">Sa.AI Assistant</span>
      <button id="close-sidebar" class="saai-close-btn" title="Close">×</button>
    </div>
    <div id="onboarding-container" class="onboarding-container">
      <div class="progress-dots">
        <span class="dot active" data-card="0"></span>
        <span class="dot" data-card="1"></span>
        <span class="dot" data-card="2"></span>
        <span class="dot" data-card="3"></span>
        <span class="dot" data-card="4"></span>
        <span class="dot" data-card="5"></span>
      </div>
      
      <!-- Skip button -->
      <button id="skip-intro" class="skip-btn">Skip intro</button>

      <!-- Card 1: Welcome / Gratitude -->
      <div class="onboarding-card active" data-card="0">
        <div class="card-content">
          <div class="saai-logo-container">
            <div class="saai-logo-main">
              <img src="${chrome.runtime.getURL('icons/icon-128.png')}" alt="Sa.AI Logo" style="height: 56px; width: auto;" class="saai-logo-img"/>
            </div>
            <div class="confetti-animation"></div>
          </div>
          <h2>Thank you for joining our beta phase</h2>
          <p>We're grateful for your support. You're early, and that means you're helping shape the future of inbox autonomy.</p>
        </div>
      </div>

      <!-- Card 2: Features Overview -->
      <div class="onboarding-card" data-card="1">
        <div class="card-content">
          <h2>Your assistant can...</h2>
          <div class="features-carousel">
            <div class="feature-item" data-feature="1">
              <div class="feature-number">1</div>
              <div class="feature-text">
                <div class="feature-title">Summarize your inbox in seconds</div>
                <small>(click on view summary button for better view)</small>
              </div>
            </div>
            <div class="feature-item" data-feature="2">
              <div class="feature-number">2</div>
              <div class="feature-text">
                <div class="feature-title">Extract tasks</div>
                <div class="feature-subtitle">automatically or add manually</div>
                <small>(click on task button)</small>
              </div>
            </div>
            <div class="feature-item" data-feature="3">
              <div class="feature-number">3</div>
              <div class="feature-text">
                <div class="feature-title">Summarize threads</div>
                <div class="feature-subtitle">just open thread</div>
                <small>and type "summarise this thread"</small>
              </div>
            </div>
            <div class="feature-item" data-feature="4">
              <div class="feature-number">4</div>
              <div class="feature-text">
                <div class="feature-title">Clear chat</div>
                <div class="feature-subtitle">when needed</div>
                <small>(click on clear button)</small>
              </div>
            </div>
            <div class="feature-item" data-feature="5">
              <div class="feature-number">5</div>
              <div class="feature-text">
                <div class="feature-title">Talk to your inbox</div>
                <div class="feature-subtitle">voice & text</div>
                <small>Ask Sa.AI any question about your emails</small>
              </div>
            </div>
          </div>
        </div>
        </div>
        
      <!-- Card 3: Credits & Usage -->
      <div class="onboarding-card" data-card="2">
        <div class="card-content">
          <h2>Your beta credits</h2>
          <div class="credits-display">
            <div class="credits-counter">
              <span class="credits-number">1000</span>
              <span class="credits-label">credits</span>
          </div>
            <div class="credits-progress">
              <div class="progress-bar">
                <div class="progress-fill"></div>
              </div>
            </div>
          </div>
          
          <!-- Credit Usage Breakdown -->
          <div class="credit-breakdown">
            <h3 style="font-size: 18px; color: var(--saai-text-primary); font-weight: 600; margin: 20px 0 15px 0;">Credit Usage</h3>
            <div class="credit-items">
              <div class="credit-item">
                <span class="credit-feature">Thread Summarization</span>
                <span class="credit-cost">5 credits</span>
              </div>
              <div class="credit-item">
                <span class="credit-feature">Web Search</span>
                <span class="credit-cost">8 credits</span>
              </div>
              <div class="credit-item">
                <span class="credit-feature">Task Extraction</span>
                <span class="credit-cost">10 credits</span>
              </div>
              <div class="credit-item">
                <span class="credit-feature">Ask Question (Inbox)</span>
                <span class="credit-cost">10 credits</span>
              </div>
              <div class="credit-item">
                <span class="credit-feature">Inbox Summarization</span>
                <span class="credit-cost">15 credits</span>
              </div>
              <div class="credit-item">
                <span class="credit-feature">Voice Mode</span>
                <span class="credit-cost">varies by feature</span>
              </div>
              <div class="credit-item credit-item-danger">
                <span class="credit-feature">Delete All Data</span>
                <span class="credit-cost">250 credits</span>
              </div>
            </div>
          </div>
          
          <p class="credits-note">We'll add more ways to earn/extend credits soon.</p>
        </div>
        </div>
        
      <!-- Card 4: Extensibility -->
      <div class="onboarding-card" data-card="3">
        <div class="card-content">
          <h2>Resize and customize</h2>
          <div class="extensibility-demo">
            <div class="demo-border">
              <div class="resize-arrow">↔</div>
              <p>Drag the border to resize the assistant</p>
          </div>
          </div>
          <p style="font-size: 16px; color: var(--saai-text-primary); font-weight: 500;">Make Sa.AI fit perfectly in your workflow by adjusting the sidebar width.</p>
        </div>
      </div>

      <!-- Card 5: Privacy & Trust -->
      <div class="onboarding-card" data-card="4">
        <div class="card-content">
          <div class="privacy-shield">
            <div class="shield-icon">🛡️</div>
          </div>
          <h2>Your privacy, non-negotiable</h2>
          <div class="privacy-points">
            <p>• All processing happens securely with your OAuth token.</p>
            <p>• Email metadata stored for 30 days for context awareness.</p>
            <p>• Data automatically deleted after retention period.</p>
            <p>• Full GDPR compliance and data protection.</p>
            <p>• You can request data deletion anytime.</p>
            <p><strong>Your data. Your control. Always.</strong></p>
          </div>
        </div>
      </div>

      <!-- Card 6: Get Started -->
      <div class="onboarding-card" data-card="5">
        <div class="card-content">
          <div class="launch-animation">
            <div class="chat-bubble">💬</div>
          </div>
          <h2>Ready to launch</h2>
          <p>Remember, this is your inbox, we're just making it smarter.</p>
          <button class="card-btn launch-btn" id="launch-saai">Start chatting</button>
        </div>
      </div>
    </div>
  `;
}

// OnboardingManager class to handle the card system
class OnboardingManager {
  constructor(sidebar) {
    this.sidebar = sidebar;
    this.currentCard = 0;
    this.totalCards = 6;
    this.isOnboardingComplete = false;
    this.onboardingContainer = sidebar.querySelector('#onboarding-container');
    this.cards = sidebar.querySelectorAll('.onboarding-card');
    this.dots = sidebar.querySelectorAll('.dot');
    this.skipBtn = sidebar.querySelector('#skip-intro');
    this.launchBtn = sidebar.querySelector('#launch-saai');
    
    debugLog('OnboardingManager initialized:', {
      onboardingContainer: !!this.onboardingContainer,
      cards: this.cards.length,
      dots: this.dots.length,
      skipBtn: !!this.skipBtn,
      launchBtn: !!this.launchBtn
    });
    
    this.bindEvents();
  }

  bindEvents() {
    // Progress dots navigation
    this.dots.forEach(dot => {
      dot.addEventListener('click', (e) => {
        e.preventDefault();
        const cardIndex = parseInt(e.target.dataset.card);
        this.showCard(cardIndex);
      });
    });

    // Skip button
    if (this.skipBtn) {
      this.skipBtn.addEventListener('click', (e) => {
        e.preventDefault();
        this.completeOnboarding();
      });
    }

    // Launch button
    if (this.launchBtn) {
      this.launchBtn.addEventListener('click', (e) => {
        e.preventDefault();
        this.completeOnboarding();
      });
    }

    // Keyboard navigation
    document.addEventListener('keydown', (e) => {
      if (!this.isOnboardingComplete && this.onboardingContainer) {
        if (e.key === 'ArrowRight' || e.key === ' ') {
          e.preventDefault();
          this.nextCard();
        } else if (e.key === 'ArrowLeft') {
          e.preventDefault();
          this.prevCard();
        } else if (e.key === 'Escape') {
          this.completeOnboarding();
        }
      }
    });
  }

  showCard(cardIndex) {
    if (cardIndex < 0 || cardIndex >= this.totalCards) {
      return;
    }

    this.currentCard = cardIndex;

    // Update cards
    this.cards.forEach((card, index) => {
      card.classList.remove('active', 'prev');
      if (index === cardIndex) {
        card.classList.add('active');
      } else if (index < cardIndex) {
        card.classList.add('prev');
      }
    });

    // Update dots
    this.dots.forEach((dot, index) => {
      dot.classList.toggle('active', index === cardIndex);
    });

    // Trigger animations for specific cards
    this.triggerCardAnimations(cardIndex);
  }

  triggerCardAnimations(cardIndex) {
    const card = this.cards[cardIndex];
    if (!card) return;

    switch(cardIndex) {
      case 1: // Features card
        this.animateFeatures(card);
        break;
      case 2: // Credits card
        this.animateCredits(card);
        break;
      case 3: // Extensibility card
        this.animateExtensibility(card);
        break;
    }
  }

  animateFeatures(card) {
    const features = card.querySelectorAll('.feature-item');
    features.forEach((feature, index) => {
      feature.style.animationDelay = `${index * 0.1}s`;
    });
  }

  animateCredits(card) {
    const progressFill = card.querySelector('.progress-fill');
    const creditsNumber = card.querySelector('.credits-number');
    if (progressFill && creditsNumber) {
      // Reset animations
      progressFill.style.animation = 'none';
      creditsNumber.style.animation = 'none';
      
      // Force reflow
      progressFill.offsetHeight;
      creditsNumber.offsetHeight;
      
      // Start animations
      progressFill.style.animation = 'fillProgress 2s ease-out 0.5s forwards';
      creditsNumber.style.animation = 'countUp 2s ease-out';
    }
  }

  animateExtensibility(card) {
    const demoArrow = card.querySelector('.resize-arrow');
    if (demoArrow) {
      demoArrow.style.animation = 'none';
      demoArrow.offsetHeight;
      demoArrow.style.animation = 'arrowMove 2s ease-in-out infinite';
    }
  }

  nextCard() {
    if (this.currentCard < this.totalCards - 1) {
      this.showCard(this.currentCard + 1);
    } else {
      this.completeOnboarding();
    }
  }

  prevCard() {
    if (this.currentCard > 0) {
      this.showCard(this.currentCard - 1);
    }
  }

  async completeOnboarding() {
    this.isOnboardingComplete = true;
    await chrome.storage.local.set({ onboardingCompleted: true });
    showChatInterface();
  }
}

// Initialize onboarding or show chat interface
async function initializeOnboarding(sidebar) {
  const result = await chrome.storage.local.get(['onboardingCompleted']);
  
  if (result.onboardingCompleted) {
    // User has already seen onboarding, show chat interface
    showChatInterface();
  } else {
    // First time user, show onboarding
    setTimeout(() => {
      const onboardingManager = new OnboardingManager(sidebar);
      sidebar._onboardingManager = onboardingManager;
      
      // Show first card
      onboardingManager.showCard(0);
    }, 100);
  }
}

function addSidebarEventListeners(sidebar, isConnected) {
  // PRODUCTION: console.log('[SaAI-Content] Adding sidebar event listeners, isConnected:', isConnected);
  // PRODUCTION: console.log('[SaAI-Content] Sidebar element:', sidebar);
  
  // Close button
  const closeBtn = sidebar.querySelector('#close-sidebar');
  if (closeBtn) {
    closeBtn.addEventListener('click', () => {
      debugLog('Close button clicked');
      closeSidebar();
    });
  }
  
  if (isConnected) {
    // Check if this is the onboarding container
    const onboardingContainer = sidebar.querySelector('#onboarding-container');
    if (onboardingContainer) {
      // Initialize onboarding system
      initializeOnboarding(sidebar);
    } else {
      // This is the chat interface - add chat listeners
      // Send button
      const sendBtn = sidebar.querySelector('#send-btn');
      if (sendBtn) {
        sendBtn.addEventListener('click', handleSendMessage);
      }
      
      // Chat input
      const chatInput = sidebar.querySelector('#chat-input');
      if (chatInput) {
        // Auto-expanding textarea functionality
        chatInput.addEventListener('input', function() {
          const baseHeight = 40;
          const maxHeight = 100;

          this.style.height = baseHeight + 'px';
          this.classList.remove('expanded');

          const scrollHeight = this.scrollHeight;

          if (scrollHeight > baseHeight) {
            const newHeight = Math.min(scrollHeight, maxHeight);
            this.style.height = newHeight + 'px';

            if (newHeight >= maxHeight) {
              this.classList.add('expanded');
            }
          }
        });
        
        chatInput.addEventListener('keydown', (e) => {
          if (e.key === 'Enter' && !e.shiftKey) {
            e.preventDefault();
            handleSendMessage();
          }
        });
        chatInput.focus();
      }
      
      // Voice controls
      setupVoiceModeControls(sidebar);
      
        // Task list button
  const taskListBtn = sidebar.querySelector('#task-list-btn');
  // PRODUCTION: console.log('[SaAI-Content] Task button found:', taskListBtn);
  if (taskListBtn) {
    // PRODUCTION: console.log('[SaAI-Content] Adding click listener to task button');
    taskListBtn.addEventListener('click', () => {
      // PRODUCTION: console.log('[SaAI-Content] Task button clicked!');
      showTaskModal();
    });
  } else {
    // PRODUCTION: console.error('[SaAI-Content] Task button not found!');
  }
  
  // Feedback button
  const feedbackBtn = sidebar.querySelector('#feedback-btn');
  if (feedbackBtn) {
    feedbackBtn.addEventListener('click', () => {
      // PRODUCTION: console.log('[SaAI-Content] Feedback button clicked!');
      showFeedbackModal();
    });
  }
  
  // Clear chat button
  const clearChatBtn = sidebar.querySelector('#clear-chat-btn');
  if (clearChatBtn) {
    clearChatBtn.addEventListener('click', () => {
      if (confirm('Are you sure you want to clear the chat history?')) {
        const chatArea = document.getElementById(CHAT_AREA_ID);
        if (chatArea) {
          chatArea.innerHTML = '';
          clearChatHistory();
          injectSuggestions(); // Re-add suggestions
        }
      }
    });
  }
      
      // Inject suggestions
      setTimeout(() => {
        injectSuggestions();
      }, 100);
    }
  } else {
    // Connect button
    const connectBtn = sidebar.querySelector('#saai-connect-btn');
    if (connectBtn) {
      connectBtn.addEventListener('click', () => {
        // PRODUCTION: console.log('[SaAI-Content] ===== Connect Outlook button clicked =====');
        debugLog('Connect Outlook button clicked');
        startOAuthFlow();
      });
    }
    
    // Debug button
    const debugBtn = sidebar.querySelector('#saai-debug-btn');
    if (debugBtn) {
      debugBtn.addEventListener('click', async () => {
        debugLog('Debug button clicked');
        const debugInfo = await debugConnectionStatus();
        
        alert(`Connection Debug Info:
- userId: ${debugInfo.userId || 'null'}
- isConnected: ${debugInfo.isConnected || 'null'}
- hasUserId: ${!!debugInfo.userId}
- hasIsConnected: ${!!debugInfo.isConnected}
- connectionCheck: ${!!(debugInfo.userId || debugInfo.isConnected)}

Check console for more details.`);
      });
    }
  }
}

async function showChatInterface() {
  if (!sidebarElement) return;
  
  sidebarElement.innerHTML = createChatInterfaceHTML();
  
  // Re-add resize handle after content change
  const resizeHandle = document.createElement('div');
  resizeHandle.className = 'saai-resize-handle';
  resizeHandle.title = 'Drag to resize sidebar';
  sidebarElement.appendChild(resizeHandle);
  
  // Re-add resize functionality
  addResizeFunctionality(sidebarElement, resizeHandle);
  
  addSidebarEventListeners(sidebarElement, true);
  
  // Load chat history after interface is created
  setTimeout(() => {
    loadChatHistory();
  }, 100);
}

async function updateSidebarContent() {
  if (!sidebarElement) return;
  
  const isConnected = await isOutlookConnected();
  
  if (isConnected) {
    sidebarElement.innerHTML = createWelcomePageHTML();
  } else {
    sidebarElement.innerHTML = createConnectPromptHTML();
  }
  
  // Re-add resize handle after content change
  const resizeHandle = document.createElement('div');
  resizeHandle.className = 'saai-resize-handle';
  resizeHandle.title = 'Drag to resize sidebar';
  sidebarElement.appendChild(resizeHandle);
  
  // Re-add resize functionality
  addResizeFunctionality(sidebarElement, resizeHandle);
  
  addSidebarEventListeners(sidebarElement, isConnected);
}

// === CHAT FUNCTIONALITY ===

// === THREAD SUMMARIZATION LOGIC ===

/**
 * Check if the current page is an Outlook thread/message page
 * @returns {boolean} True if on a thread/message page
 */
function isThreadPage() {
  const pathname = window.location.pathname || '';
  const hash = window.location.hash || '';
  const search = window.location.search || '';
  const pathAndHash = pathname + hash;
  const threadPatterns = [
    /\/mail\/.*\/id\//i,
    /\/owa\/.*[?&]ItemID=/i,
    /[?&]itemid=/i,
    /[?&]ItemID=/i,
    /\/mail\/.*[?&]id=/i
  ];
  function urlIsThread(urlStr) {
    if (!urlStr) return false;
    const beforeQ = urlStr.split('?')[0];
    return threadPatterns.some(p => p.test(beforeQ) || p.test((urlStr.match(/\?.*/) || [''])[0]));
  }
  if (urlIsThread(pathAndHash) || urlIsThread(search)) return true;
  // Old Outlook / live.com: thread URL may be in a same-origin iframe
  try {
    for (let i = 0; i < window.frames.length; i++) {
      try {
        if (urlIsThread(window.frames[i].location.href)) return true;
      } catch (_) {}
    }
  } catch (_) {}
  return false;
}

/**
 * Get threadId and subjectLine from DOM in one pass (same row/source) for consistency
 * @returns {{ threadId: string|null, subjectLine: string|null }}
 */
function getThreadContextFromDOM() {
  const idInHref = idInHrefRegex;
  let linkWithId = null;
  let threadId = null;

  const selectedRow = document.querySelector('[aria-selected="true"], [data-selected="true"], .is-selected, [class*="selected"][class*="item"]');
  if (selectedRow) {
    linkWithId = selectedRow.querySelector('a[href*="/id/"]');
    if (linkWithId) {
      const m = (linkWithId.getAttribute('href') || '').match(idInHref);
      if (m && m[1]) threadId = decodeId(m[1]);
    }
  }
  if (!threadId) {
    const links = document.querySelectorAll('a[href*="/id/"]');
    for (const a of links) {
      const m = (a.getAttribute('href') || '').match(idInHref);
      if (m && m[1]) {
        threadId = decodeId(m[1]);
        linkWithId = a;
        break;
      }
    }
  }
  const subjectLine = extractSubjectFromRowOrPage(linkWithId) || extractSubjectLine();
  return { threadId, subjectLine };
}

// Gather thread context: URL first, then DOM (same source for threadId + subject) for draft/summarize
function getThreadContext() {
  const urlHasThread = isThreadPage();
  let threadId = extractThreadId();
  let subjectLine = extractSubjectLine();

  const onMailView = (window.location.pathname + (window.location.hash || '')).indexOf('/mail/') !== -1;
  if (!threadId && onMailView) {
    const fromDom = getThreadContextFromDOM();
    if (fromDom.threadId) threadId = fromDom.threadId;
    if (fromDom.subjectLine) subjectLine = fromDom.subjectLine;
  }
  // Safeguard: only use threadId if it passes validation (length, charset); avoid sending invalid IDs to backend
  threadId = normalizeThreadId(threadId);
  if (!subjectLine && (urlHasThread || threadId)) {
    subjectLine = extractSubjectLine();
  }
  const isOnThreadPage = urlHasThread || !!threadId;
  return { isOnThreadPage, threadId, subjectLine };
}

// Retrieve or generate a userId consistent with text flow
async function getOrGenerateUserId() {
  const { userId } = await chrome.storage.local.get(['userId']);
  return userId || 'anonymous-user';
}

/** Regex to capture Outlook message ID after /id/ (base64-like; stops at ? or end) */
var outlookIdInUrlRegex = /\/id\/([A-Za-z0-9%_.=+-]+)/i;

function decodeIdSafe(encoded) {
  try {
    return decodeURIComponent(encoded);
  } catch (_) {
    return encoded;
  }
}

/**
 * Extract conversation/message ID from an arbitrary URL string (e.g. from current window or a frame).
 * Used so we can run the same logic on window.location and on same-origin iframe locations.
 * @param {string} urlString - Full URL (e.g. location.href)
 * @returns {string|null} Decoded ID or null
 */
function extractThreadIdFromUrlString(urlString) {
  if (!urlString || typeof urlString !== 'string') return null;
  const beforeQuery = urlString.split('?')[0];
  const match = beforeQuery.match(outlookIdInUrlRegex);
  if (!match || !match[1]) return null;
  return decodeIdSafe(match[1]);
}

/**
 * Extract thread/message ID from the current Outlook URL.
 * Outlook Live (old Outlook) may show the message in an iframe while the top window has a different URL,
 * so we also check same-origin frames. We support path, hash, and path+hash.
 * @returns {string|null} Thread/Message ID or null if not found
 */
function extractThreadId() {
  const currentHref = window.location.href;
  const urlObj = new URL(currentHref);
  const pathname = urlObj.pathname;
  const hash = window.location.hash || '';
  const pathAndHash = pathname + hash;

  // 1: /id/XXX in path (outlook.live.com: /mail/0/inbox/id/AAkALg...?path=%2Fmail&wa=wsignin1.0)
  const idInPath = pathname.match(outlookIdInUrlRegex);
  if (idInPath && idInPath[1]) return decodeIdSafe(idInPath[1]);

  // 2: /id/XXX in hash (SPA routing)
  const idInHash = hash.match(outlookIdInUrlRegex);
  if (idInHash && idInHash[1]) return decodeIdSafe(idInHash[1]);

  // 3: Path+hash before ? (edge cases on live.com)
  const beforeQuery = pathAndHash.split('?')[0];
  const idInPathOrHash = beforeQuery.match(outlookIdInUrlRegex);
  if (idInPathOrHash && idInPathOrHash[1]) return decodeIdSafe(idInPathOrHash[1]);

  // 4: Query param itemid / ItemID
  const itemId = urlObj.searchParams.get('itemid') || urlObj.searchParams.get('ItemID');
  if (itemId) return itemId;

  // 5: Classic OWA – /owa/?ItemID=...
  const owaMatch = (urlObj.search || '').match(/[?&]ItemID=([A-Za-z0-9%_.=-]+)/i);
  if (owaMatch && owaMatch[1]) return decodeIdSafe(owaMatch[1]);

  // 6: Old Outlook / live.com – conversation ID may live in a same-origin iframe (reading pane)
  try {
    for (let i = 0; i < window.frames.length; i++) {
      try {
        const frameHref = window.frames[i].location.href;
        const id = extractThreadIdFromUrlString(frameHref);
        if (id) return id;
      } catch (_) { /* cross-origin frame, skip */ }
    }
  } catch (_) {}

  return null;
}

var idInHrefRegex = /\/id\/([A-Za-z0-9%_.=+-]+)/i;

function decodeId(idStr) {
  try {
    return decodeURIComponent(idStr);
  } catch (e) {
    return idStr;
  }
}

/** Min/max length for Outlook message/thread IDs (base64-like; avoid sending garbage or truncated values) */
var OUTLOOK_THREAD_ID_MIN_LEN = 20;
var OUTLOOK_THREAD_ID_MAX_LEN = 600;

/**
 * Safeguard: validate thread/message ID before sending to backend.
 * Outlook IDs are long base64-like strings (AAkALg..., AQQk..., etc.). Rejects empty, too short,
 * too long, or obviously non-ID strings so we never send invalid data or trigger backend errors.
 * @param {*} id - Value from URL or DOM
 * @returns {boolean}
 */
function isValidOutlookThreadId(id) {
  if (id == null || typeof id !== 'string') return false;
  const trimmed = id.trim();
  if (trimmed.length < OUTLOOK_THREAD_ID_MIN_LEN || trimmed.length > OUTLOOK_THREAD_ID_MAX_LEN) return false;
  // Allow only chars that appear in Outlook IDs (letters, digits, /, +, =, and URL-safe)
  if (!/^[A-Za-z0-9/+=_.%-]+$/.test(trimmed)) return false;
  return true;
}

/**
 * Normalize and validate thread ID; return valid string or null. Use this before putting threadId in payload.
 * @param {*} id - Raw thread/message ID
 * @returns {string|null}
 */
function normalizeThreadId(id) {
  if (id == null || typeof id !== 'string') return null;
  const trimmed = id.trim();
  return isValidOutlookThreadId(trimmed) ? trimmed : null;
}

/**
 * Try to get thread/message ID from the DOM when URL has no /id/ (e.g. reading pane at /mail/0/)
 * Prefers selected row / reading pane so we get the visible message.
 * @returns {string|null} Message/thread ID or null
 */
function extractThreadIdFromDOM() {
  const idInHref = idInHrefRegex;
  // 1. URL hash (e.g. #/mail/0/inbox/id/XXX)
  const hash = window.location.hash || '';
  const hashMatch = hash.match(idInHref);
  if (hashMatch && hashMatch[1]) return decodeId(hashMatch[1]);

  // 2. Selected list row: [aria-selected="true"] or similar, then link with /id/
  const selectedRow = document.querySelector('[aria-selected="true"], [data-selected="true"], .is-selected, [class*="selected"][class*="item"], [class*="Selected"][class*="Item"]');
  if (selectedRow) {
    const link = selectedRow.querySelector('a[href*="/id/"]');
    if (link) {
      const m = (link.getAttribute('href') || '').match(idInHref);
      if (m && m[1]) return decodeId(m[1]);
    }
  }

  // 3. Any link with /id/ inside selected/focused/main/reading pane (prefer by order)
  const links = Array.from(document.querySelectorAll('a[href*="/id/"]'));
  const priorityContainers = ['[aria-selected="true"]', '[role="main"]', '[class*="readingPane"]', '[data-app-section="Mail"]', '[class*="ReadingPane"]'];
  for (const sel of priorityContainers) {
    const container = document.querySelector(sel);
    if (!container) continue;
    for (const a of links) {
      if (!container.contains(a)) continue;
      const m = (a.getAttribute('href') || '').match(idInHref);
      if (m && m[1]) return decodeId(m[1]);
    }
  }

  // 4. First link with /id/ in document (often the open message in list)
  for (const a of links) {
    const m = (a.getAttribute('href') || '').match(idInHref);
    if (m && m[1]) return decodeId(m[1]);
  }

  // 5. Data attributes (selected or main first)
  const withDataId = document.querySelector('[aria-selected="true"] [data-item-id], [aria-selected="true"] [data-id], [role="main"] [data-item-id], [data-item-id], [data-id], [data-message-id]');
  if (withDataId) {
    const id = withDataId.getAttribute('data-item-id') || withDataId.getAttribute('data-id') || withDataId.getAttribute('data-message-id');
    if (id && id.length > 10) return id.trim();
  }
  return null;
}

/**
 * Get subject from the same row/element that contains the thread link (for consistency with threadId)
 * @param {Element} [anchor] - Optional link element that has href with /id/; subject is taken from its row
 * @returns {string|null}
 */
function extractSubjectFromRowOrPage(anchor) {
  const notSubject = /^(inbox|focused|other|re|fw|fwd|draft|sent|unread|\d{1,2}\/\d{1,2}|\d{1,2}:\d{2})$/i;
  const maxLen = 400;

  if (anchor) {
    const row = anchor.closest('[role="listitem"], [role="row"], [class*="listItem"], [class*="ListItem"], [class*="mailListItem"], tr');
    if (row) {
      const spans = row.querySelectorAll('span, div, a');
      let best = '';
      for (const s of spans) {
        const t = (s.textContent || '').trim();
        if (t.length > best.length && t.length <= maxLen && !notSubject.test(t) && !/^\d{1,2}\s*(am|pm)$/i.test(t))
          best = t;
      }
      if (best) return best;
    }
  }

  const main = document.querySelector('[role="main"], [data-app-section="Mail"], [class*="readingPane"]');
  if (main) {
    const headings = main.querySelectorAll('h1, h2, h3, [role="heading"]');
    for (const h of headings) {
      const t = (h.textContent || '').trim();
      if (t && t.length <= maxLen && !notSubject.test(t)) return t;
    }
  }
  return null;
}

/**
 * Check if the message contains summarization keywords
 * @param {string} message - The user's message
 * @returns {boolean} True if message contains summarization keywords
 */
// Check if message is asking to draft/reply/compose an email
function isDraftOrReplyRequest(message) {
  const lowerMessage = message.toLowerCase();
  const draftKeywords = [
    'draft',
    'compose',
    'write',
    'reply',
    'respond',
    'answer',
    'email back',
    'send',
    'generate'
  ];
  
  const emailKeywords = [
    'email',
    'message',
    'response',
    'reply',
    'thread'
  ];
  
  // Check if message contains any draft keyword
  const hasDraftKeyword = draftKeywords.some(keyword => lowerMessage.includes(keyword));
  // Check if it's email-related
  const hasEmailContext = emailKeywords.some(keyword => lowerMessage.includes(keyword));
  
  return hasDraftKeyword && (hasEmailContext || lowerMessage.length < 50);
}

function isSummarizationRequest(message) {
  const lowerMessage = message.toLowerCase();
  const summarizationKeywords = [
    'summarise',
    'summarize', 
    'summary',
    'summaries'
  ];
  
  const contextKeywords = [
    'email',
    'thread',
    'this',
    'current',
    'whole',
    'entire'
  ];
  
  const inboxKeywords = [
    'inbox',
    'all emails',
    'my emails',
    'email list'
  ];
  
  // Check if message contains any summarization keyword
  const hasSummarizationKeyword = summarizationKeywords.some(keyword => 
    lowerMessage.includes(keyword)
  );
  
  // Check if message contains inbox-specific keywords
  const hasInboxKeyword = inboxKeywords.some(keyword => 
    lowerMessage.includes(keyword)
  );
  
  // Check if message contains context keywords (optional)
  const hasContextKeyword = contextKeywords.some(keyword => 
    lowerMessage.includes(keyword)
  );
  
  // Return true if it has summarization keyword and either context keyword or is short (likely about current thread)
  return hasSummarizationKeyword && (hasContextKeyword || message.length < 50);
}

/**
 * Check if the message is specifically requesting inbox summarization
 * @param {string} message - The user's message
 * @returns {boolean} True if message is requesting inbox summarization
 */
function isInboxSummarizationRequest(message) {
  const lowerMessage = message.toLowerCase();
  const summarizationKeywords = [
    'summarise',
    'summarize', 
    'summary',
    'summaries'
  ];
  
  const inboxKeywords = [
    'inbox',
    'all emails',
    'my emails',
    'email list'
  ];
  
  // Check if message contains both summarization and inbox keywords
  const hasSummarizationKeyword = summarizationKeywords.some(keyword => 
    lowerMessage.includes(keyword)
  );
  
  const hasInboxKeyword = inboxKeywords.some(keyword => 
    lowerMessage.includes(keyword)
  );
  
  return hasSummarizationKeyword && hasInboxKeyword;
}

/**
 * Get current page context for better error messages
 * @returns {string} Description of current page
 */
function getCurrentPageContext() {
  if (isThreadPage()) {
    const threadId = extractThreadId();
    return `thread page (ID: ${threadId})`;
  } else {
    return 'inbox page';
  }
}

/**
 * Extract subject line from the current Outlook thread/message (URL thread view or reading pane)
 * @returns {string|null} Subject line or null if not found
 */
function extractSubjectLine() {
  const onMailView = window.location.pathname.indexOf('/mail/') !== -1;
  if (!isThreadPage() && !onMailView) {
    return null;
  }
  
  const subjectSelectors = [
    // Gmail-style (shared codebase)
    'h2[data-thread-perm-id]',
    '.hP',
    '[data-thread-perm-id] h2',
    '.gD h2',
    '.adn h2',
    // Outlook / Fluent: common patterns
    '[role="main"] h1',
    '[role="main"] h2',
    '[role="main"] h3',
    '[role="heading"]',
    '.readingPane h1', '.readingPane h2', '.readingPane h3',
    '[class*="subject"]', '[class*="Subject"]',
    '[class*="messageHeader"] h2', '[class*="messageHeader"] h1',
    '[aria-label*="Subject"]', '[aria-label*="subject"]',
    // Generic in main content
    '[role="main"] h2',
    'h2', 'h1', '.subject'
  ];
  
  for (let i = 0; i < subjectSelectors.length; i++) {
    const el = document.querySelector(subjectSelectors[i]);
    if (el) {
      const subject = (el.textContent || '').trim();
      if (subject && subject.length > 0 && subject.length < 500) {
        debugLog('Subject found using selector:', subjectSelectors[i], 'Subject:', subject);
        return subject;
      }
    }
  }
  
  // First heading in main (any h1/h2/h3) to catch Outlook reading pane
  const main = document.querySelector('[role="main"], .readingPane, [data-app-section="Mail"]');
  if (main) {
    const heading = main.querySelector('h1, h2, h3');
    if (heading) {
      const t = (heading.textContent || '').trim();
      if (t && t.length > 0 && t.length < 500) return t;
    }
  }
  
  // Page title: "Subject - Outlook" or "Subject | Microsoft Outlook" or "Outlook - Subject"
  const pageTitle = document.title || '';
  if (pageTitle) {
    const parts = pageTitle.split(/\s*[-|]\s*/);
    for (let j = 0; j < parts.length; j++) {
      const candidate = parts[j].trim();
      if (candidate && candidate.length > 0 && candidate.length < 500 &&
          !/^outlook$/i.test(candidate) && !/^microsoft$/i.test(candidate)) {
        debugLog('Subject from page title:', candidate);
        return candidate;
      }
    }
  }
  
  debugLog('No subject line found');
  return null;
}

// === RESPONSE FORMATTING ===

/**
 * Format unstructured text responses into well-formatted, readable content
 * @param {string} text - The raw response text from backend
 * @returns {string} - Formatted HTML string
 */
function formatUnstructuredResponse(text) {
  if (!text || typeof text !== 'string') {
    return text;
  }
  
  debugLog('Formatting unstructured response:', text);
  
  let formattedText = text;
  
  // 0. Pre-process: Protect product names and technical terms from being formatted as numbered lists
  // This prevents "NeuroTrack NX-500" from being split into "500." as a numbered item
  // We'll use a more sophisticated approach in the numbered list detection instead
  
  // 1. Format markdown headers (e.g., "### News", "### Sources:")
  formattedText = formattedText.replace(/^###\s+(.+)$/gm, '<h3 class="response-header">$1</h3>');
  
  // 2. Format numbered lists (e.g., "1. **Device**: ...")
  formattedText = formattedText.replace(/(\d+\.\s+\*\*([^*]+)\*\*:\s*([^2-9]+?)(?=\d+\.\s+\*\*|$))/g, 
    '<div class="response-section"><div class="section-header"><span class="section-number">$1</span></div><div class="section-content">$3</div></div>'
  );
  
  // 3. Format bold headers (e.g., "**Device**:", "**Fault**:")
  formattedText = formattedText.replace(/\*\*([^*]+)\*\*:\s*/g, '<strong class="field-label">$1:</strong> ');
  
  // 4. Format vendor responses section
  formattedText = formattedText.replace(/(\*\*Vendor Responses\*\*:\s*)(.*?)(?=\*\*Next Steps\*\*|$)/s, (match, header, content) => {
    const vendorResponses = content.split('-').filter(item => item.trim());
    if (vendorResponses.length > 0) {
      const formattedResponses = vendorResponses.map(response => {
        const trimmed = response.trim();
        if (trimmed) {
          return `<div class="vendor-response">• ${trimmed}</div>`;
        }
        return '';
      }).join('');
      return `<div class="vendor-section"><strong class="field-label">Vendor Responses:</strong><div class="vendor-list">${formattedResponses}</div></div>`;
    }
    return match;
  });
  
  // 5. Format next steps section
  formattedText = formattedText.replace(/(\*\*Next Steps\*\*:\s*)(.*?)(?=\*\*|$)/s, (match, header, content) => {
    const steps = content.split(',').filter(step => step.trim());
    if (steps.length > 0) {
      const formattedSteps = steps.map(step => {
        const trimmed = step.trim();
        if (trimmed) {
          return `<div class="next-step">• ${trimmed}</div>`;
        }
        return '';
      }).join('');
      return `<div class="next-steps-section"><strong class="field-label">Next Steps:</strong><div class="steps-list">${formattedSteps}</div></div>`;
    }
    return match;
  });
  
  // 6. Convert numbered lists with em dash pattern to bullet points
  // Pattern: "1. Title — Date Description" or "1. Title — Description"
  formattedText = formattedText.replace(/(^|\n)(\d+)\.\s+(.+?)\s*—\s*(.+?)(?=\n\d+\.\s+|$)/g, (match, start, number, title, description) => {
    const trimmedTitle = title.trim();
    const trimmedDesc = description.trim();
    // Convert to bullet point format
    return `${start}<div class="list-item"><span class="bullet-point">•</span><div class="item-content"><strong>${trimmedTitle}</strong> — ${trimmedDesc}</div></div>`;
  });

  // 7. Format any remaining numbered items that weren't caught (but be much smarter about it)
  // Only format if it's clearly a numbered list item, not part of a product name or technical term
  // Convert numbered lists to bullet points for better readability
  formattedText = formattedText.replace(/(^|\n)(\d+\.\s+)([^1-9\n]+?)(?=\n\d+\.\s+|$)/g, (match, start, number, content) => {
    const trimmedContent = content.trim();
    
    // Skip formatting if this looks like it's part of a product name or technical specification
    const isProductName = /(NeuroTrack|device|model|serial|version|firmware|hardware)/i.test(trimmedContent);
    const isTechnicalSpec = /(NX-\d+|v\d+\.\d+|rev\s*\d+)/i.test(trimmedContent);
    const isPartOfSentence = trimmedContent.includes('.') || trimmedContent.includes(',') || trimmedContent.includes(';');
    
    // Only format if it's clearly a standalone list item
    if (!isProductName && !isTechnicalSpec && !isPartOfSentence && trimmedContent.length > 20) {
      // Convert to bullet point instead of numbered item
      return `${start}<div class="list-item"><span class="bullet-point">•</span><div class="item-content">${content}</div></div>`;
    }
    
    return match;
  });
  
  // 8. Add general styling classes
  formattedText = formattedText.replace(/\*\*([^*]+)\*\*/g, '<strong class="highlight">$1</strong>');
  
  // 9. Format any remaining lists or bullet points
  formattedText = formattedText.replace(/([•\-\*]\s+)([^\n]+)/g, '<div class="list-item">$1$2</div>');
  
  // 10. Format citations/references like [1][2] or [4]
  formattedText = formattedText.replace(/\[(\d+)\]/g, '<span class="citation">[$1]</span>');
  
  // 11. Format source links section at the end (multiple formats)
  formattedText = formattedText.replace(/(Sources?:?\s*)(.*?)$/s, (match, header, content) => {
    // Try different source formats
    let formattedSources = [];
    
    // Format 1: [Name](URL) format
    const nameUrlMatches = content.match(/\d+\.\s+\[([^\]]+)\]\(([^)]+)\)/g);
    if (nameUrlMatches) {
      nameUrlMatches.forEach(match => {
        const nameMatch = match.match(/\[([^\]]+)\]\(([^)]+)\)/);
        if (nameMatch) {
          formattedSources.push(`<div class="source-item"><span class="source-name">${nameMatch[1]}</span><a href="${nameMatch[2]}" target="_blank" class="source-link">${nameMatch[2]}</a></div>`);
        }
      });
    }
    
    // Format 2: [1]: URL format (for news summaries)
    if (formattedSources.length === 0) {
      const numberUrlMatches = content.match(/\[(\d+)\]:\s*(https?:\/\/[^\s]+)/g);
      if (numberUrlMatches) {
        numberUrlMatches.forEach(match => {
          const numberUrlMatch = match.match(/\[(\d+)\]:\s*(https?:\/\/[^\s]+)/);
          if (numberUrlMatch) {
            formattedSources.push(`<div class="source-item"><span class="source-name">Source ${numberUrlMatch[1]}</span><a href="${numberUrlMatch[2]}" target="_blank" class="source-link">${numberUrlMatch[2]}</a></div>`);
          }
        });
      }
    }
    
    if (formattedSources.length > 0) {
      return `<div class="sources-section"><strong class="field-label">${header}</strong><div class="sources-list">${formattedSources.join('')}</div></div>`;
    }
    return match;
  });
  
  // 11. Format news headlines (e.g., "News: Google has launched...")
  formattedText = formattedText.replace(/^(News:\s*)([^[]+?)(?=\[|$)/gm, (match, prefix, content) => {
    if (content.trim().length > 10) {
      return `<div class="news-headline"><strong class="news-label">${prefix}</strong><span class="news-content">${content.trim()}</span></div>`;
    }
    return match;
  });
  
  // 13. Format paragraphs for better readability
  formattedText = formattedText.replace(/([^<>]+?)(?=\n\n|\n\s*\n|$)/g, (match, content) => {
    // Skip if this is already formatted HTML
    if (content.includes('<') && content.includes('>')) {
      return content;
    }
    // Only format if it's substantial text (more than 50 characters)
    if (content.trim().length > 50 && !content.includes('Sources?:')) {
      return `<div class="content-paragraph">${content.trim()}</div>`;
    }
    return content;
  });
  
  // 14. Add container wrapper
  formattedText = `<div class="formatted-response">${formattedText}</div>`;
  
  debugLog('Formatted response created');
  return formattedText;
}

// === MODIFIED MESSAGE HANDLING ===

async function handleSendMessage() {
  // PRODUCTION: console.log('[SaAI-Content] ===== handleSendMessage() called =====');
  
  const input = document.getElementById('chat-input');
  if (!input) {
    debugError('Chat input not found');
    return;
  }
  
  const message = input.value.trim();
  // PRODUCTION: console.log('[SaAI-Content] Message received - length:', message.length);
  if (!message) return;
  
  // Check credit limit first
  if (userCredits.remaining === 0) {
    showCreditLimitReached();
    return;
  }
  
  // Classify the prompt to determine credit cost
  const classification = classifyPrompt(message);
  debugLog('Prompt classified:', classification);
  
  // Check if user has enough credits
  if (!hasEnoughCredits(classification.cost)) {
    showCreditWarning('insufficient');
    return;
  }
  
  const chatArea = document.getElementById(CHAT_AREA_ID);
  if (!chatArea) {
    debugError('Chat area not found');
    return;
  }
  
  // Add user message
  appendMessage('user', message, chatArea);
  input.value = '';
  
  // Save chat history
  saveChatHistory();
  
  // Hide suggestions after first user message
  const suggestions = chatArea.querySelector('#saai-suggestions');
  if (suggestions) {
    suggestions.remove();
  }
  
  // Show enhanced thinking indicator
  const typingIndicator = showThinkingIndicator(chatArea);
  
  try {
    // Get userId from storage
    const { userId } = await chrome.storage.local.get(['userId']);
    
    if (!userId) {
      throw new Error('Please connect your Outlook first');
    }
    
    // Check if this is a summarization request
    const isSummarizeRequest = isSummarizationRequest(message);
    const isInboxSummarizeRequest = isInboxSummarizationRequest(message);
    const isDraftRequest = isDraftOrReplyRequest(message);
    const { isOnThreadPage, threadId, subjectLine } = getThreadContext();
    
    debugLog('Message analysis:', {
      message,
      isSummarizeRequest,
      isInboxSummarizeRequest,
      isDraftRequest,
      isOnThreadPage,
      threadId,
      subjectLine,
      currentUrl: window.location.href,
      sidebarWidth: SIDEBAR_WIDTH,
      sidebarElement: !!sidebarElement
    });
    
    // Validation 1: Draft/Reply requests require an open thread
    if (isDraftRequest && !isOnThreadPage) {
      if (typingIndicator) typingIndicator.remove();
      appendMessage('bot', 'Open a thread for this process and then prompt again.', chatArea);
      return;
    }
    
    // Validation 2: Thread summarization requires a detected thread (URL or DOM)
    if (isSummarizeRequest && !isInboxSummarizeRequest && !isOnThreadPage) {
      if (typingIndicator) typingIndicator.remove();
      const onMailView = window.location.pathname.indexOf('/mail/') !== -1;
      const msg = onMailView
        ? 'We couldn\'t detect which message is open. Try opening the email in its own tab (right‑click the email → Open in new tab), then ask again.'
        : 'Open a thread for this process and then prompt again.';
      appendMessage('bot', msg, chatArea);
      return;
    }
    // Validation 3: Summarize requires a valid threadId (safeguard: URL might have /id/ but ID failed validation)
    if (isSummarizeRequest && !isInboxSummarizeRequest && isOnThreadPage && !threadId) {
      if (typingIndicator) typingIndicator.remove();
      appendMessage('bot', 'We couldn\'t read the message ID for this thread. Try opening the email in a new tab and try again.', chatArea);
      return;
    }
    
    // Prepare payload
    const payload = {
      query: message,
      userId: userId,
      context: 'OutlookChat',
      sidebarWidth: SIDEBAR_WIDTH // Add sidebar width for debugging
    };
    
    // When a thread is open: always send threadId and subjectLine for draft and summarize (backend needs both)
    if (isOnThreadPage) {
      if (threadId) payload.threadId = threadId;
      payload.subjectLine = subjectLine || (document.title && document.title.split(/\s*[-|]\s*/)[0].trim()) || '';
      if (payload.subjectLine) debugLog('Including thread context:', { threadId: payload.threadId || 'none', subjectLine: payload.subjectLine });
    }

    // Summarize thread: mark action so backend can resolve conversation
    if (isSummarizeRequest && isOnThreadPage) {
      payload.action = 'summarize_thread';
      debugLog('Adding thread summarization data:', { threadId: payload.threadId || 'none', subjectLine: payload.subjectLine || 'none' });
    }
    
    debugLog('Sending message to n8n');
    
    debugLog('Sending message to background script');
    
    // Send message to background script
    const response = await chrome.runtime.sendMessage({
      action: 'sendToN8N',
      data: {
        endpoint: 'chat',
        payload: payload
      }
    });
    
    debugLog('Response received from background script:', response);
    
    // Remove typing indicator
    if (typingIndicator) {
      typingIndicator.remove();
    }
    
    // Check if user was disconnected due to token expiration
    if (response?.data?.disconnected) {
      debugLog('User disconnected due to token expiration - not showing error');
      // Silently handle disconnection - refresh UI to show connect button
      // Clear connection status and reload the page state
      chrome.storage.local.get(['userId', 'isConnected'], (result) => {
        if (!result.userId && !result.isConnected) {
          // Connection was cleared, refresh the UI
          location.reload();
        }
      });
      return;
    }
    
    if (response?.success) {
      // PRODUCTION: console.log('[SaAI-Content] ===== RESPONSE SUCCESS - Starting credit tracking logic =====');
      debugLog('Response received:', response.data);
      debugLog('Response data type:', typeof response.data);
      debugLog('Response data keys:', response.data ? Object.keys(response.data) : 'null');
      debugLog('Response data length:', Array.isArray(response.data) ? response.data.length : 'not array');
      
      // Deduct credits after successful response
      try {
        // PRODUCTION: console.log('[SaAI-Content] About to check shouldTrackCreditConsumption');
        // Decide first if we should track credits for this response
        const shouldTrackCredits = shouldTrackCreditConsumption(response.data, message);
        // PRODUCTION: console.log('[SaAI-Content] shouldTrackCredits result:', shouldTrackCredits);

        if (shouldTrackCredits) {
          // Deduct credits
          await deductCredits(classification.cost, classification.feature);
          debugLog(`Credits deducted: ${classification.cost} for ${classification.feature}`);

          // Send credit tracking data to webhook
          try {
            // PRODUCTION: console.log('[SaAI-Content] About to send credit tracking');

            const creditData = await sendCreditTrackingToWebhook(userId, message, classification.cost);
            if (creditData) {
              await updateUIWithTotalCredits(creditData);
            }
          } catch (creditError) {
            debugError('Credit tracking failed:', creditError);
          }
        } else {
          debugLog('Skipping credit deduction and tracking - validation message detected');
        }
      } catch (error) {
        debugError('Credit handling failed:', error);
      }
      
      // Handle array-wrapped responses from n8n
      let responseData = response.data;
      debugLog('Raw response data type:', typeof responseData);
      debugLog('Raw response data:', responseData);
      debugLog('Raw response data length:', Array.isArray(responseData) ? responseData.length : 'not array');
      
      // Handle string responses (double-encoded JSON)
      if (typeof responseData === 'string') {
        debugLog('Response is a string, attempting to parse...');
        try {
          responseData = JSON.parse(responseData);
          debugLog('Successfully parsed string to:', typeof responseData);
          debugLog('Parsed response data:', responseData);
        } catch (parseError) {
          debugError('Failed to parse response string:', parseError);
        }
      }
      
      if (Array.isArray(responseData)) {
        debugLog('Response is an array, extracting first item');
        responseData = responseData[0];
        debugLog('Extracted response data:', responseData);
        debugLog('Extracted response data type:', typeof responseData);
        debugLog('Extracted response data keys:', Object.keys(responseData));
        debugLog('Extracted response data JSON:', JSON.stringify(responseData, null, 2));
        debugLog('Has reply property:', responseData.hasOwnProperty('reply'));
        debugLog('Reply property value:', responseData.reply);
        debugLog('Reply property type:', typeof responseData.reply);
      }
      
      // Check if this is a fallback response
      if (responseData?.fallback) {
        debugLog('Using fallback response - n8n webhook unavailable');
        
        // Create a special message for fallback responses (sanitized)
        const fallbackDiv = document.createElement('div');
        fallbackDiv.className = 'message bot-message';
        
        const messageContent = document.createElement('div');
        messageContent.className = 'message-content';
        
        // Sanitize by using textContent instead of innerHTML
        const mainMessage = document.createElement('div');
        mainMessage.style.marginBottom = '8px';
        mainMessage.textContent = responseData.message || 'No response message';
        
        const statusInfo = document.createElement('div');
        statusInfo.style.fontSize = '12px';
        statusInfo.style.opacity = '0.8';
        statusInfo.style.borderTop = '1px solid rgba(0,0,0,0.1)';
        statusInfo.style.paddingTop = '8px';
        statusInfo.style.marginTop = '8px';
        
        const webhookLabel = document.createElement('strong');
        webhookLabel.textContent = 'Webhook Status: ';
        statusInfo.appendChild(webhookLabel);
        statusInfo.appendChild(document.createTextNode(responseData.webhookStatus || 'Unknown'));
        statusInfo.appendChild(document.createElement('br'));
        
        const suggestionLabel = document.createElement('strong');
        suggestionLabel.textContent = 'Suggestion: ';
        statusInfo.appendChild(suggestionLabel);
        statusInfo.appendChild(document.createTextNode(responseData.suggestion || 'Check n8n configuration'));
        
        messageContent.appendChild(mainMessage);
        messageContent.appendChild(statusInfo);
        fallbackDiv.appendChild(messageContent);
        chatArea.appendChild(fallbackDiv);
        chatArea.scrollTop = chatArea.scrollHeight;
        return;
      }
      
      // Handle normal response
      if (responseData && (responseData.high_priority_emails || responseData.medium_priority || responseData.already_replied_closed_threads || responseData.missed_or_ignored_emails)) {
        // Email summary data
        appendTableMessage('bot', responseData, chatArea);
      } else if (responseData && responseData.reply) {
        // Handle nested JSON reply format
        debugLog('Found reply field, processing...');
        debugLog('Reply field exists and has value');
        try {
          debugLog('Processing reply:', responseData.reply);
          debugLog('Reply type:', typeof responseData.reply);
          debugLog('Reply starts with [ or {:', responseData.reply.startsWith('[') || responseData.reply.startsWith('{'));
          
          // First, try to parse the outer reply
          let parsedReply = responseData.reply;
          
          // If it's a string that looks like JSON, parse it
          if (typeof parsedReply === 'string' && (parsedReply.startsWith('[') || parsedReply.startsWith('{'))) {
            debugLog('Parsing JSON string...');
            parsedReply = JSON.parse(parsedReply);
            debugLog('Parsed successfully, result type:', typeof parsedReply);
            debugLog('Parsed result is array:', Array.isArray(parsedReply));
          }
          
          // Handle array format with nested JSON strings
          if (Array.isArray(parsedReply)) {
            debugLog('Parsed reply is an array:', parsedReply);
            debugLog('Array length:', parsedReply.length);
            // Extract summary from the first item if it has a summary field
            const firstItem = parsedReply[0];
            debugLog('First item:', firstItem);
            debugLog('First item type:', typeof firstItem);
            debugLog('First item keys:', firstItem ? Object.keys(firstItem) : 'null');
            
            if (firstItem && firstItem.summary) {
              debugLog('Found summary field:', firstItem.summary);
              appendMessage('bot', firstItem.summary, chatArea);
            } else if (firstItem && typeof firstItem === 'object') {
              // Handle complex structured response
              debugLog('Found complex structured response');
              debugLog('First item is object, checking for structured format...');
              const formattedResponse = formatStructuredResponse(firstItem);
              debugLog('Formatted response created, length:', formattedResponse.length);
              appendMessage('bot', formattedResponse, chatArea);
            } else {
              // If no summary field, try to extract from the item itself
              const summaryText = typeof firstItem === 'string' ? firstItem : JSON.stringify(firstItem);
              debugLog('Using fallback summary text:', summaryText);
              // Apply formatting to unstructured text
              const formattedText = formatMessageContent(summaryText);
              debugLog('Using formatted fallback text, length:', formattedText.length);
              appendMessage('bot', formattedText, chatArea, false, true);
            }
          } else if (parsedReply && typeof parsedReply === 'object') {
            // Handle object format
            if (parsedReply.summary) {
              appendMessage('bot', parsedReply.summary, chatArea);
            } else if (parsedReply.message) {
              appendMessage('bot', parsedReply.message, chatArea);
            } else {
              appendMessage('bot', JSON.stringify(parsedReply), chatArea);
            }
                      } else {
              // Fallback to raw text - apply formatting
              const formattedText = formatMessageContent(parsedReply);
              debugLog('Using formatted fallback text, length:', formattedText.length);
              appendMessage('bot', formattedText, chatArea, false, true);
            }
          
        } catch (parseError) {
          debugError('Error parsing reply:', parseError);
          debugError('Parse error details:', parseError.message);
          debugError('Parse error stack:', parseError.stack);
          debugLog('Raw reply data:', responseData.reply);
          
          // If parsing fails, try to extract plain text
          let plainText = responseData.reply;
          
          // Try to extract text from common patterns
          if (plainText.includes('summary')) {
            // Try to extract summary from JSON-like string
            const summaryMatch = plainText.match(/"summary":"([^"]+)"/);
            if (summaryMatch) {
              plainText = summaryMatch[1];
            }
          }
          
          appendMessage('bot', plainText, chatArea);
        }
      } else if (responseData && typeof responseData === 'string') {
        // Handle direct string response
        debugLog('Found direct string response:', responseData);
        // Apply formatting to unstructured text
        const formattedText = formatMessageContent(responseData);
        debugLog('Using formatted string response, length:', formattedText.length);
        appendMessage('bot', formattedText, chatArea, false, true);
              } else if (responseData && responseData.summary) {
          // Handle direct summary field
          debugLog('Found direct summary field:', responseData.summary);
          // Apply formatting to summary text
          const formattedText = formatMessageContent(responseData.summary);
          debugLog('Using formatted summary, length:', formattedText.length);
          appendMessage('bot', formattedText, chatArea, false, true);
              } else if (responseData && responseData.message) {
          // Direct message response
          // Apply formatting to message text
          const formattedText = formatMessageContent(responseData.message);
          debugLog('Using formatted message, length:', formattedText.length);
          appendMessage('bot', formattedText, chatArea, false, true);
      } else if (responseData && Array.isArray(responseData)) {
        // Handle direct array response
        debugLog('Found direct array response:', responseData);
        const firstItem = responseData[0];
        if (firstItem && firstItem.summary) {
          debugLog('Found summary in direct array:', firstItem.summary);
          appendMessage('bot', firstItem.summary, chatArea);
                  } else {
            debugLog('Using first array item as text:', firstItem);
            const itemText = typeof firstItem === 'string' ? firstItem : JSON.stringify(firstItem);
            // Apply formatting to array item text
            const formattedText = formatMessageContent(itemText);
            debugLog('Using formatted array item, length:', formattedText.length);
            appendMessage('bot', formattedText, chatArea, false, true);
          }
      } else if (responseData && typeof responseData === 'object' && !Array.isArray(responseData)) {
        // Handle direct structured object response
        debugLog('Checking for direct structured object response...');
        const keys = Object.keys(responseData);
        debugLog('Object keys:', keys);
        
        // Check if this is a structured response with sections
        if (keys.some(key => key.includes('Subject') || key.includes('Purpose') || key.includes('Points') || key.includes('Decisions') || key.includes('Questions') || key.includes('Action'))) {
          debugLog('Found direct structured response object');
          const formattedResponse = formatStructuredResponse(responseData);
          debugLog('Formatted direct response, length:', formattedResponse.length);
          appendMessage('bot', formattedResponse, chatArea);
        } else if (responseData.summary) {
          // Check for direct summary field
          debugLog('Found direct summary field');
          appendMessage('bot', responseData.summary, chatArea);
        } else if (responseData.message) {
          // Check for direct message field
          debugLog('Found direct message field');
          debugLog('Raw message before formatting:', responseData.message);
          
          // If message contains HTML tags, use it directly as it should be properly formatted
          if (responseData.message.includes('<div') || responseData.message.includes('<p>') || responseData.message.includes('<span>')) {
            debugLog('Message contains HTML, using directly');
            appendMessage('bot', responseData.message, chatArea, false, true);
          } else {
            // Plain text message, apply formatting
            const formattedMessage = formatMessageContent(responseData.message);
            debugLog('Formatted message:', formattedMessage);
            appendMessage('bot', formattedMessage, chatArea, false, true);
          }
        } else {
          // Fallback for other object types
          debugLog('Using object as JSON string');
          appendMessage('bot', JSON.stringify(responseData, null, 2), chatArea);
        }
      } else {
        // Check for empty response
        if (!responseData || (typeof responseData === 'object' && Object.keys(responseData).length === 0)) {
          debugLog('Empty response detected:', responseData);
          appendMessage('bot', 'I received an empty response from the AI. This might be due to a webhook configuration issue or the AI service being temporarily unavailable. Please try again in a moment.', chatArea);
          return;
        }
        
        // Fallback
        debugLog('No valid response format found, using fallback');
        debugLog('responseData:', responseData);
        debugLog('responseData.reply:', responseData?.reply);
        debugLog('responseData.message:', responseData?.message);
        debugLog('responseData.summary:', responseData?.summary);
        debugLog('responseData type:', typeof responseData);
        debugLog('responseData keys:', responseData ? Object.keys(responseData) : 'null');
        debugLog('responseData.reply exists:', !!responseData?.reply);
        debugLog('responseData.reply truthy:', !!responseData?.reply);
        debugLog('responseData.reply length:', responseData?.reply?.length);
        appendMessage('bot', 'I received your message!', chatArea);
      }
    } else {
      throw new Error(response?.error || 'Failed to get response from AI');
    }
    
  } catch (error) {
    debugError('Chat error:', error);
    
    // Remove typing indicator
    if (typingIndicator) {
      typingIndicator.remove();
    }
    
    // Check if this is a disconnection (not an error to show)
    if (error.message && error.message.includes('disconnected')) {
      debugLog('User disconnected - not showing error message');
      // Silently handle disconnection - user will need to reconnect
      // Refresh the page to show connect button
      chrome.storage.local.get(['userId', 'isConnected'], (result) => {
        if (!result.userId && !result.isConnected) {
          location.reload();
        }
      });
      return;
    }
    
    if (isExtensionContextInvalidated(error)) {
      showRefreshPageMessage(chatArea);
      return;
    }
    appendMessage('bot', `Error: ${error.message}`, chatArea);
  }
}

function appendMessage(sender, text, chatArea, temporary = false, isFormatted = false) {
  const messageDiv = document.createElement('div');
  messageDiv.className = `message ${sender}-message${temporary ? ' temporary' : ''}`;
  
  const messageContent = document.createElement('div');
  messageContent.className = 'message-content';
  
  // Auto-detect simple HTML if caller forgot to pass isFormatted
  const looksLikeHtml = typeof text === 'string' && /<\w+[^>]*>/.test(text);

  if (isFormatted || looksLikeHtml) {
    // For formatted HTML content - sanitize before injecting
    debugLog('Setting innerHTML for formatted content (sanitized):', text);
    try {
      messageContent.innerHTML = sanitizeHtml(text);
      debugLog('innerHTML set successfully');
    } catch (error) {
      debugError('Error setting innerHTML:', error);
      messageContent.textContent = text; // Fallback to text
    }
  } else {
    // For plain text content
    messageContent.textContent = text;
  }
  
  messageDiv.appendChild(messageContent);
  
  // Add timestamp for user messages
  if (sender === 'user' && !temporary) {
    const timeStamp = document.createElement('div');
    timeStamp.className = 'user-message-timestamp';
    timeStamp.textContent = formatCardTimestamp();
    messageDiv.appendChild(timeStamp);
  }
  
  chatArea.appendChild(messageDiv);
  chatArea.scrollTop = chatArea.scrollHeight;

  // Save chat history after each message
  if (!temporary) {
    saveChatHistory();
  }

  return temporary ? messageDiv : null;
}

// === UTILITY FUNCTIONS ===

/**
 * Compute days old from various date inputs, with guardrails.
 * Accepts ISO/RFC strings, epoch seconds/ms, and numeric strings.
 */
function computeDaysOld(input) {
  if (input === null || input === undefined || input === '') return null;
  try {
    let emailDate;
    if (typeof input === 'number') {
      const ms = input < 1e12 ? input * 1000 : input;
      emailDate = new Date(ms);
    } else if (typeof input === 'string' && /^\d+$/.test(input.trim())) {
      const n = parseInt(input.trim(), 10);
      const ms = n < 1e12 ? n * 1000 : n;
      emailDate = new Date(ms);
    } else {
      emailDate = new Date(input);
    }
    if (isNaN(emailDate.getTime())) return null;
    const now = new Date();
    // Guard against absurd dates
    if (emailDate.getTime() > now.getTime() + 24 * 60 * 60 * 1000) return null;
    if (emailDate.getFullYear() < 1995) return null;
    const diffTime = now.getTime() - emailDate.getTime();
    const diffDays = Math.floor(diffTime / (1000 * 60 * 60 * 24));
    return diffDays;
  } catch (error) {
    debugError('Error computing days old:', error);
    return null;
  }
}

// Format a date/time-like value into compact relative time (e.g., 2h ago, 1d ago)
function formatRelativeTime(dateLike) {
  if (!dateLike) return '—';
  try {
    let dateObj;
    if (typeof dateLike === 'number') {
      const ms = dateLike < 1e12 ? dateLike * 1000 : dateLike;
      dateObj = new Date(ms);
    } else if (typeof dateLike === 'string' && /^\d+$/.test(dateLike.trim())) {
      const n = parseInt(dateLike.trim(), 10);
      const ms = n < 1e12 ? n * 1000 : n;
      dateObj = new Date(ms);
    } else {
      dateObj = new Date(dateLike);
    }
    if (isNaN(dateObj.getTime())) return '—';

    const now = new Date();
    const diffMs = now.getTime() - dateObj.getTime();
    if (diffMs < 0) return '—';

    const sec = Math.floor(diffMs / 1000);
    const min = Math.floor(sec / 60);
    const hr = Math.floor(min / 60);
    const day = Math.floor(hr / 24);

    if (day > 0) return day === 1 ? '1 day' : `${day} days`;
    if (hr > 0) return hr === 1 ? '1h ago' : `${hr}h ago`;
    if (min > 0) return min === 1 ? '1m ago' : `${min}m ago`;
    return 'Just now';
  } catch (e) {
    return '—';
  }
}

// Format numeric days or date-like into compact label (Today, 1 day, N days)
function formatDaysLabel(value) {
  if (value === null || value === undefined) return '—';
  if (typeof value === 'number' && Number.isFinite(value)) {
    if (value === 0) return 'Today';
    if (value === 1) return '1 day';
    return `${value} days`;
  }
  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (/^\d+$/.test(trimmed)) {
      const n = parseInt(trimmed, 10);
      return formatDaysLabel(n);
    }
    const computed = computeDaysOld(trimmed);
    return computed === null ? '—' : formatDaysLabel(computed);
  }
  const computed = computeDaysOld(value);
  return computed === null ? '—' : formatDaysLabel(computed);
}

// Determine the best RECEIVED text using multiple fallbacks
function getReceivedText(email) {
  if (!email || typeof email !== 'object') return '—';
  // If backend already gave a relative string, prefer it
  const relativeCandidates = [
    email.received,
    email.received_text,
    email.relativeTime,
    email.relative_time
  ];
  for (const v of relativeCandidates) {
    if (typeof v === 'string' && v.trim()) return v.trim();
  }
  // If we have days_old, convert to label
  if (email.days_old !== undefined) return formatDaysLabel(email.days_old);
  if (email['days old'] !== undefined) return formatDaysLabel(email['days old']);

  // Try timestamp-like fields
  const timeCandidates = [
    email.receivedAt,
    email.received_at,
    email.dateTime,
    email.datetime,
    email.date,
    email.time,
    email.timestamp,
    email.ts,
    email.message_date,
    email.internalDate
  ];
  for (const t of timeCandidates) {
    const rel = formatRelativeTime(t);
    if (rel && rel !== '—') return rel;
  }
  return '—';
}

function appendTableMessage(sender, emailData, chatArea) {
  const messageDiv = document.createElement('div');
  messageDiv.className = `message ${sender}-message`;
  
  const tableContainer = document.createElement('div');
  tableContainer.className = 'email-summary-table inbox-v2';
  
  const title = document.createElement('h3');
  title.textContent = '📧 Inbox Summary';
  tableContainer.appendChild(title);
  
  // Helper function to clean and format email addresses
  function formatEmailAddress(emailString) {
    if (!emailString) return 'Unknown Sender';
    
    // Remove angle brackets and extract email
    let cleanEmail = emailString.replace(/[<>]/g, '').trim();
    
    // Truncate for compact display
    if (cleanEmail.length > 25) {
      return cleanEmail.substring(0, 22) + '...';
    }
    
    // For normal email addresses, just clean them up
    return cleanEmail;
  }
  
  // Helper function to truncate long subjects
  function truncateSubject(subject, maxLength = 50) {
    if (!subject) return 'No Subject';
    if (subject.length <= maxLength) return subject;
    return subject.substring(0, maxLength) + '...';
  }
  
  // Helper function to format next action
  function formatNextAction(action) {
    if (!action) return '—';
    if (typeof action === 'string') {
      // No truncation - let CSS handle text wrapping
      return action;
    }
    return '—';
  }
  
  // Helper function to format days old (robust, handles numeric strings and dates)
  function formatDaysOld(daysOld) {
    if (daysOld === null || daysOld === undefined) return '—';
    if (typeof daysOld === 'string') {
      const trimmed = daysOld.trim();
      // Pure number? treat as days count
      if (/^\d+$/.test(trimmed)) {
        const n = parseInt(trimmed, 10);
        if (n === 0) return 'Today';
        if (n === 1) return '1 day';
        return `${n} days`;
      }
      // Try as date string
      const computed = computeDaysOld(trimmed);
      if (computed !== null) {
        if (computed === 0) return 'Today';
        if (computed === 1) return '1 day';
        return `${computed} days`;
      }
      return trimmed || '—';
    }
    if (typeof daysOld === 'number' && Number.isFinite(daysOld)) {
      if (daysOld === 0) return 'Today';
      if (daysOld === 1) return '1 day';
      return `${daysOld} days`;
    }
    return '—';
  }
  
  const priorities = [
    { key: 'high_priority_emails', label: 'High Priority', color: '#ffebee', icon: '⊗' },
    { key: 'medium_priority', label: 'Medium Priority', color: '#fff3e0', icon: '⚠' },
    { key: 'low_priority', label: 'Low Priority', color: '#f0f9ff', icon: 'ⓘ' },
    { key: 'already_replied_closed_threads', label: 'Closed/Already Replied', color: '#e8f5e8', icon: '✓' },
    { key: 'missed_or_ignored_emails', label: 'Missed or Ignored', color: '#f5f5f5', icon: '⏱' }
  ];
  
  let totalRows = 0;
  priorities.forEach(priority => {
    const emails = emailData[priority.key];
    if (emails && emails.length > 0) totalRows += emails.length;
  });
  
  // Show preview if large
  const PREVIEW_ROWS = 3;
  let previewMode = totalRows > 6;
  
  debugLog('Table preview mode calculation:', {
    totalRows,
    PREVIEW_ROWS,
    previewMode,
    willShowButton: previewMode
  });
  
  priorities.forEach(priority => {
    const emails = emailData[priority.key];
    if (emails && emails.length > 0) {
      const section = document.createElement('div');
      section.className = 'priority-section';
      
      // Add data-priority attribute for enhanced styling
      if (priority.key === 'high_priority_emails') {
        section.setAttribute('data-priority', 'high');
      } else if (priority.key === 'medium_priority') {
        section.setAttribute('data-priority', 'medium');
      } else if (priority.key === 'low_priority') {
        section.setAttribute('data-priority', 'low');
      } else if (priority.key === 'already_replied_closed_threads') {
        section.setAttribute('data-priority', 'replied');
      } else if (priority.key === 'missed_or_ignored_emails') {
        section.setAttribute('data-priority', 'missed');
      }
      
      const header = document.createElement('h4');
      // Safe: priority.icon and priority.label are from our own code, not backend
      header.innerHTML = `${priority.icon} ${priority.label} <span class="email-count">(${emails.length})</span>`;
      section.appendChild(header);
      
      const table = document.createElement('table');
      table.className = 'email-table';
      
      // Define column widths matching EmailTable.tsx reference
      const colgroup = document.createElement('colgroup');
      colgroup.innerHTML = `
        <col style="width: 200px;">
        <col style="width: 280px;">
        <col style="width: 120px;">
        <col style="width: 380px;">
        <col style="width: 140px;">
      `;
      table.appendChild(colgroup);
      
      const thead = document.createElement('thead');
      thead.innerHTML = `
        <tr>
          <th>From</th>
          <th>Subject</th>
          <th>Received</th>
          <th>Next Action</th>
          <th></th>
        </tr>
      `;
      table.appendChild(thead);
      
      const tbody = document.createElement('tbody');
      const emailsToShow = previewMode ? emails.slice(0, PREVIEW_ROWS) : emails;
      
      emailsToShow.forEach((email, idx) => {
        const row = document.createElement('tr');
        row.className = idx % 2 === 0 ? 'even-row' : 'odd-row';
        
        const senderCell = document.createElement('td');
        senderCell.className = 'sender-cell';
        senderCell.textContent = formatEmailAddress(email.sender);
        senderCell.title = email.sender; // Show full email on hover
        
        const subjectCell = document.createElement('td');
        subjectCell.className = 'subject-cell';
        subjectCell.textContent = email.subject || 'No Subject';
        subjectCell.title = email.subject; // Show full subject on hover
        
        const receivedCell = document.createElement('td');
        receivedCell.className = 'received-cell';
        const receivedText = getReceivedText(email);
        receivedCell.textContent = receivedText;
        receivedCell.title = receivedText;
        
        const nextActionCell = document.createElement('td');
        nextActionCell.className = 'next-action-cell';
        const nextActionValue = (email && (email.next_action || email.nextAction || email.next_action_required || email.nextActionRequired || email.action || email.todo)) || '';
        nextActionCell.textContent = formatNextAction(nextActionValue);
        nextActionCell.title = nextActionValue; // Show full action on hover
        
        row.appendChild(senderCell);
        row.appendChild(subjectCell);
        row.appendChild(receivedCell);
        row.appendChild(nextActionCell);

        const taskCell = document.createElement('td');
        taskCell.className = 'task-cell';
        const taskBtn = document.createElement('button');
        taskBtn.innerHTML = '<span style="margin-right: 4px;">+</span>Add to Tasks';
        taskBtn.className = 'saai-task-btn';
        taskBtn.addEventListener('click', async (e) => {
          e.preventDefault();
          e.stopPropagation();
          
          // PRODUCTION: console.log('[SaAI] Add to Tasks button clicked!');
          
          // Determine priority based on the section
          let priorityLevel = 'medium'; // default
          if (priority.key === 'high_priority_emails') priorityLevel = 'high';
          else if (priority.key === 'medium_priority') priorityLevel = 'medium';
          else if (priority.key === 'low_priority') priorityLevel = 'low';
          else if (priority.key === 'missed_or_ignored_emails') priorityLevel = 'low'; // Missed emails = low priority
          else if (priority.key === 'already_replied_closed_threads') priorityLevel = 'low'; // Already replied = low priority
          
          // PRODUCTION: console.log('[SaAI] Priority determined:', priorityLevel);
          // PRODUCTION: console.log('[SaAI] Email data - priority:', priorityLevel, 'has subject:', !!(email.subject));
          
          // Send task to webhook
          try {
            // Disable button and show loading state
            taskBtn.disabled = true;
            taskBtn.innerHTML = '<span style="margin-right: 4px;">⏳</span>Adding...';
            
            await addEmailToTasks({
              from: email.sender,
              context: email.subject || 'No Subject',
              taskName: nextActionValue || 'Review and respond to email',
              priority: priorityLevel
            });
            
            // Success state
            taskBtn.innerHTML = '<span style="margin-right: 4px;">✓</span>Added';
            taskBtn.style.background = '#16a34a';
            taskBtn.style.cursor = 'default';
            
          } catch (error) {
            // PRODUCTION: console.error('[SaAI] Error in button click handler:', error);
            // Reset button on error
            taskBtn.disabled = false;
            taskBtn.innerHTML = '<span style="margin-right: 4px;">+</span>Add to Tasks';
          }
        });
        taskCell.appendChild(taskBtn);
        row.appendChild(taskCell);
        tbody.appendChild(row);
      });
      
      table.appendChild(tbody);
      section.appendChild(table);
      
      // Add "show more" indicator if in preview mode
      if (previewMode && emails.length > PREVIEW_ROWS) {
        const moreIndicator = document.createElement('div');
        moreIndicator.className = 'more-indicator';
        moreIndicator.textContent = `+${emails.length - PREVIEW_ROWS} more emails`;
        section.appendChild(moreIndicator);
      }
      
      tableContainer.appendChild(section);
    }
  });
  
  if (previewMode) {
    debugLog('Creating View Full Summary button');
    const viewFullBtn = document.createElement('button');
    viewFullBtn.textContent = '📋 View Full Summary';
    viewFullBtn.className = 'view-full-btn';
    viewFullBtn.id = 'view-full-summary-btn';
    
    // Single handler to prevent double-opening
    viewFullBtn.addEventListener('click', (e) => {
      debugLog('View Full Summary button clicked (addEventListener)');
      e.preventDefault();
      e.stopPropagation();
      debugLog('Opening full table');
      openFullTable(emailData);
    });
    
    tableContainer.appendChild(viewFullBtn);
    debugLog('View Full Summary button created and added to DOM');
    
    // Verify button was added
    const addedButton = tableContainer.querySelector('#view-full-summary-btn');
    debugLog('Button verification:', {
      buttonFound: !!addedButton,
      buttonText: addedButton?.textContent,
      buttonParent: addedButton?.parentNode?.className
    });
  } else {
    debugLog('Preview mode is false, not creating button');
  }
  
  // If no emails found
  if (!priorities.some(p => emailData[p.key] && emailData[p.key].length > 0)) {
    const noEmailsMsg = document.createElement('div');
    noEmailsMsg.className = 'no-emails';
    noEmailsMsg.innerHTML = `
      <div class="no-emails-icon">📭</div>
      <div class="no-emails-text">No emails found in your inbox.</div>
    `;
    tableContainer.appendChild(noEmailsMsg);
  }
  
  messageDiv.appendChild(tableContainer);
  chatArea.appendChild(messageDiv);
  chatArea.scrollTop = chatArea.scrollHeight;
}

function openFullTable(emailData) {
  try {
    debugLog('Opening full table');
    
    const tableWindow = window.open('', '_blank');
    
    if (!tableWindow) {
      debugError('Failed to open new window - popup blocked?');
      alert('Failed to open full table. Please allow popups for this site.');
      return;
    }
  
  // Normalize incoming data: handle outer array and JSON string in `reply`
  function normalizeEmailData(raw) {
    try {
      // If raw is a string, try JSON.parse first
      if (typeof raw === 'string') {
        const parsed = JSON.parse(raw);
        return normalizeEmailData(parsed);
      }
      // If raw is an array like [{ reply: "{...}" }]
      if (Array.isArray(raw)) {
        const first = raw[0];
        if (first && typeof first.reply === 'string') {
          return JSON.parse(first.reply);
        }
        return first || {};
      }
      // If raw has a `reply` string directly
      if (raw && typeof raw.reply === 'string') {
        return JSON.parse(raw.reply);
      }
      // Otherwise assume it's already the final object shape
      return raw || {};
    } catch (e) {
      debugWarn('Failed to normalize email data, using raw:', e);
      return raw || {};
    }
  }

  const normalizedData = normalizeEmailData(emailData);

  // Helper function to clean and format email addresses (same as in appendTableMessage)
  function formatEmailAddress(emailString) {
    if (!emailString) return 'Unknown Sender';
    
    // Remove angle brackets and extract email
    let cleanEmail = emailString.replace(/[<>]/g, '').trim();
    
    // Truncate for compact display
    if (cleanEmail.length > 25) {
      return cleanEmail.substring(0, 22) + '...';
    }
    
    // For normal email addresses, just clean them up
    return cleanEmail;
  }
  
  // Helper function to truncate long subjects
  function truncateSubject(subject, maxLength = 80) {
    if (!subject) return 'No Subject';
    if (subject.length <= maxLength) return subject;
    return subject.substring(0, maxLength) + '...';
  }
  
  // Extract next action from multiple possible keys
  function extractNextAction(email) {
    if (!email || typeof email !== 'object') return '—';
    const candidates = [
      email.next_action,
      email.nextAction,
      email.next_action_required,
      email.nextActionRequired,
      email.action,
      email.todo
    ];
    for (const val of candidates) {
      if (typeof val === 'string' && val.trim()) return val.trim();
    }
    return '—';
  }

  // Helper function to format next action (delegates to extractor)
  function formatNextAction(action) {
    if (!action) return '—';
    if (typeof action === 'string') return action;
    return '—';
  }
  
  // Helper function to format days old (same as in main scope)
  function formatDaysOld(daysOld) {
    if (daysOld === null || daysOld === undefined) return '—';
    // Numeric strings like "3"
    if (typeof daysOld === 'string') {
      const trimmed = daysOld.trim();
      if (/^\d+$/.test(trimmed)) {
        const n = parseInt(trimmed, 10);
        if (n === 0) return 'Today';
        if (n === 1) return '1 day';
        return `${n} days`;
      }
      // Non-numeric string; try as date string
      const computed = computeDaysOld(trimmed);
      if (computed !== null) {
        if (computed === 0) return 'Today';
        if (computed === 1) return '1 day';
        return `${computed} days`;
      }
      return trimmed || '—';
    }
    if (typeof daysOld === 'number' && Number.isFinite(daysOld)) {
      if (daysOld === 0) return 'Today';
      if (daysOld === 1) return '1 day';
      return `${daysOld} days`;
    }
    return '—';
  }
  
  // Helper function to compute days old (same as in main scope)
  function computeDaysOld(input) {
    if (input === null || input === undefined || input === '') return null;
    try {
      let emailDate;
      if (typeof input === 'number') {
        // Treat as epoch seconds if small, else ms
        const ms = input < 1e12 ? input * 1000 : input;
        emailDate = new Date(ms);
      } else if (typeof input === 'string' && /^\d+$/.test(input.trim())) {
        const n = parseInt(input.trim(), 10);
        const ms = n < 1e12 ? n * 1000 : n;
        emailDate = new Date(ms);
      } else {
        emailDate = new Date(input);
      }
      if (isNaN(emailDate.getTime())) return null;
      // Guard against absurd dates
      const now = new Date();
      if (emailDate.getTime() > now.getTime() + 24 * 60 * 60 * 1000) return null;
      if (emailDate.getFullYear() < 1995) return null;
      const diffTime = now.getTime() - emailDate.getTime();
      const diffDays = Math.floor(diffTime / (1000 * 60 * 60 * 24));
      return diffDays;
    } catch (error) {
      debugError('Error computing days old:', error);
      return null;
    }
  }
  
  tableWindow.document.write(`
    <html>
      <head>
        <title>Inbox Summary</title>
        <style>
          body { 
            background: #ffffff; 
            color: #0f172a; 
            font-family: Inter, Segoe UI, Arial, sans-serif; 
            font-size: 16px; /* bump overall */
            padding: 32px; 
            margin: 0;
          }
          h2 { 
            color: #0f172a; 
            text-align: center;
            margin-bottom: 32px;
            font-size: 24px;
            font-weight: 600;
          }
          .priority-section {
            margin-bottom: 32px;
            background: #f8fafc;
            border-radius: 12px;
            padding: 20px;
            border: 1px solid #e2e8f0;
          }
          .priority-header {
            color: #0f172a;
            font-size: 18px;
            font-weight: 600;
            margin-bottom: 16px;
            display: flex;
            align-items: center;
            gap: 8px;
          }
          .email-count {
            font-size: 14px;
            color: #64748b;
            font-weight: 400;
          }
          table { 
            width: 100%; 
            border-collapse: collapse; 
            margin-bottom: 16px; 
            background: #ffffff;
            border-radius: 8px;
            overflow: hidden;
            border: 1px solid #e2e8f0;
            table-layout: fixed;
          }
          th, td { 
            padding: 12px 16px; 
            text-align: left;
            border-bottom: 1px solid #f1f5f9;
            vertical-align: top;
          }
          th { 
            background: #f1f5f9; 
            color: #0f172a; 
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            font-size: 12px;
          }
          th:nth-child(1) { width: 30%; text-align: left; } /* Subject */
          th:nth-child(2) { width: 20%; text-align: left; } /* From */
          th:nth-child(3) { width: 35%; text-align: left; } /* Next Action */
          th:nth-child(4) { width: 15%; text-align: center; } /* Days Old */
          tr:nth-child(even) { 
            background: #ffffff; 
          }
          tr:nth-child(odd) { 
            background: #f8fafc; 
          }
          tr:hover {
            background: #f1f5f9;
            transition: background 0.2s ease;
          }
          .subject-cell {
            font-weight: 500;
            max-width: 100%;
            word-wrap: break-word;
            line-height: 1.4;
            color: #0f172a;
            padding-right: 8px;
          }
          .sender-cell {
            font-size: 16px;
            color: #0f172a;
            max-width: 100%;
            word-wrap: break-word;
            font-family: Inter, Segoe UI, Arial, sans-serif;
            padding-left: 8px;
          }
          .next-action-cell {
            font-size: 16px;
            color: #0f172a;
            max-width: 100%;
            word-wrap: break-word;
            line-height: 1.4;
            font-weight: 500;
          }
          .days-old-cell {
            font-size: 16px;
            color: #0f172a;
            max-width: 100%;
            word-wrap: break-word;
            font-weight: 500;
            text-align: center;
            font-family: 'Inter', sans-serif;
            vertical-align: middle;
          }
          
          /* Force center alignment for Days Old column in full table */
          .days-old-cell {
            text-align: center !important;
          }
          button { 
            padding: 12px 24px; 
            border-radius: 8px; 
            border: none; 
            background: #0f172a; 
            color: #ffffff; 
            font-weight: 600; 
            font-size: 14px; 
            cursor: pointer; 
            margin-top: 24px;
            transition: all 0.2s ease;
          }
          button:hover {
            background: #1e293b;
            transform: translateY(-1px);
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
          }
          button:active {
            transform: translateY(0);
          }
          /* Desktop-only; mobile rules removed */
        </style>
        <style>
          /* Inbox Summary V2 overrides for full-view window */
          body {
            max-width: 1200px;
            margin: 32px auto;
          }
          .priority-section {
            background: #fff;
            border: 1px solid #e5e7eb;
            border-left-width: 4px;
            border-radius: 12px;
            box-shadow: 0 1px 2px rgba(0,0,0,.04);
            overflow: hidden;
            padding: 0;
            margin-bottom: 24px;
          }
          .priority-section > .priority-header {
            margin: 0;
            padding: 14px 16px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            line-height: 1.2;
            font-weight: 800;
            font-size: 16px;
            color: #111827;
            background: linear-gradient(90deg, #f9fafb, #ffffff);
            border: none;
          }
          .priority-section[data-priority="high"] { border-left-color: #dc2626; }
          .priority-section[data-priority="medium"] { border-left-color: #d97706; }
          .priority-section[data-priority="low"] { border-left-color: #2563eb; }
          .priority-section[data-priority="replied"] { border-left-color: #16a34a; }
          .priority-section[data-priority="missed"] { border-left-color: #4b5563; }

          .priority-section[data-priority="high"] > .priority-header { background: linear-gradient(90deg, #fef2f2, #fee2e2); }
          .priority-section[data-priority="medium"] > .priority-header { background: linear-gradient(90deg, #fffbeb, #fef3c7); }
          .priority-section[data-priority="low"] > .priority-header { background: linear-gradient(90deg, #eff6ff, #dbeafe); }
          .priority-section[data-priority="replied"] > .priority-header { background: linear-gradient(90deg, #f0fdf4, #dcfce7); }
          .priority-section[data-priority="missed"] > .priority-header { background: linear-gradient(90deg, #f9fafb, #f3f4f6); }

          .priority-section .email-count {
            display: inline-block;
            padding: 6px 10px;
            border-radius: 999px;
            background: #fff;
            font-size: 12px;
            font-weight: 800;
            box-shadow: 0 1px 2px rgba(0,0,0,.06);
          }
          .priority-section[data-priority="high"] .email-count { color: #dc2626; }
          .priority-section[data-priority="medium"] .email-count { color: #d97706; }
          .priority-section[data-priority="low"] .email-count { color: #2563eb; }
          .priority-section[data-priority="replied"] .email-count { color: #16a34a; }
          .priority-section[data-priority="missed"] .email-count { color: #4b5563; }

          .email-table { 
            width: 100%;
            border-collapse: collapse;
            table-layout: fixed;
            font-size: 16px; 
            border: none;
            box-shadow: none;
          }
          .email-table {
            table-layout: fixed;
            width: 100%;
          }
          .email-table th { 
            font-size: 16px !important; 
            padding: 16px 24px;
          }
          .email-table td { 
            font-size: 16px !important; 
            padding: 20px 24px;
          }
          .email-table thead tr { background: linear-gradient(90deg, #f9fafb, #ffffff); }
          .email-table th {
            text-align: left;
            padding: 12px 16px;
            font-weight: 800;
            font-size: 11px;
            text-transform: uppercase;
            letter-spacing: .08em;
            color: #6b7280;
            border-bottom: 2px solid #e5e7eb;
          }
          .email-table tbody tr { border-bottom: 1px solid #f3f4f6; }
          .email-table tbody tr:hover { background: linear-gradient(90deg, rgba(249,250,251,.5), transparent); }
          .email-table td {
            padding: 16px 18px;
            vertical-align: top;
            color: #111827;
            font-weight: 500;
            line-height: 1.5;
          }
          .sender-cell {
            font-weight: 700;
            color: #111827;
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif;
            max-width: 200px;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
          }
          .subject-cell {
            font-weight: 700;
            color: #111827;
            max-width: 280px;
            white-space: normal;
            word-wrap: break-word;
            line-height: 1.4;
          }
          .received-cell {
            white-space: normal !important;
            word-wrap: break-word;
            line-height: 1.4;
          }
          .next-action-cell {
            color: #111827;
            font-weight: 500;
            max-width: 380px;
            white-space: normal !important;
            overflow: hidden !important;
            display: -webkit-box !important;
            -webkit-line-clamp: 2;
            -webkit-box-orient: vertical;
            line-height: 1.5;
          }
          .days-old-cell { color: #4b5563; font-weight: 600; text-align: center; }
          .task-cell {
            text-align: right;
            vertical-align: middle;
            padding: 20px 24px;
            width: 140px;
          }
          .saai-task-btn-popup {
            background: #111827;
            color: #ffffff;
            border: none;
            border-radius: 6px;
            padding: 6px 12px;
            font-size: 13px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.15s ease;
            white-space: nowrap;
            display: inline-flex;
            align-items: center;
            gap: 6px;
            box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
          }
          .saai-task-btn-popup:hover {
            background: #1f2937;
            transform: translateY(-1px);
            box-shadow: 0 4px 8px rgba(0, 0, 0, 0.15);
          }
          .saai-task-btn-popup:active {
            transform: translateY(0);
            box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
          }
        </style>
      </head>
      <body>
        <h2>📧 Inbox Summary (Full Table)</h2>
        ${[
          { key: 'high_priority_emails', label: 'High Priority', icon: '⊗' },
          { key: 'medium_priority', label: 'Medium Priority', icon: '⚠' },
          { key: 'low_priority', label: 'Low Priority', icon: 'ⓘ' },
          { key: 'already_replied_closed_threads', label: 'Closed/Already Replied', icon: '✓' },
          { key: 'missed_or_ignored_emails', label: 'Missed or Ignored', icon: '⏱' }
        ].map(priority => {
          const emails = normalizedData[priority.key];
          if (emails && emails.length > 0) {
            return `
              <div class="priority-section" data-priority="${priority.key === 'high_priority_emails' ? 'high' : 
                                                           priority.key === 'medium_priority' ? 'medium' : 
                                                           priority.key === 'low_priority' ? 'low' :
                                                           priority.key === 'already_replied_closed_threads' ? 'replied' : 'missed'}">
                <div class="priority-header">
                  ${priority.icon} ${priority.label} 
                  <span class="email-count">(${emails.length} emails)</span>
                </div>
                <table class="email-table">
                  <colgroup>
                    <col style="width: 200px;">
                    <col style="width: 280px;">
                    <col style="width: 120px;">
                    <col style="width: 380px;">
                    <col style="width: 140px;">
                  </colgroup>
                  <thead>
                    <tr>
                      <th>From</th>
                      <th>Subject</th>
                      <th>Received</th>
                      <th>Next Action</th>
                      <th></th>
                    </tr>
                  </thead>
                  <tbody>
                    ${emails.map(email => {
                      // Handle different priority types for full table view
                      let senderText, nextActionText, receivedText;
                      
                        senderText = formatEmailAddress(email.sender);
                        const na = extractNextAction(email);
                      nextActionText = priority.key === 'missed_or_ignored_emails' ? '—' : formatNextAction(na);
                      receivedText = getReceivedText(email);
                      
                      const emailData = JSON.stringify({
                        sender: email.sender,
                        subject: email.subject || 'No Subject',
                        nextAction: na || 'Review and respond to email',
                        priorityKey: priority.key
                      }).replace(/"/g, '&quot;');
                      
                      const btnId = `task-btn-${priority.key}-${emails.indexOf(email)}`;
                      
                      return `
                        <tr>
                          <td class="sender-cell" title="${email.sender}">${senderText}</td>
                          <td class="subject-cell" title="${email.subject}">${email.subject || 'No Subject'}</td>
                          <td class="received-cell">${receivedText}</td>
                          <td class="next-action-cell" title="${(extractNextAction(email) || '').replace(/"/g, '&quot;')}">${nextActionText}</td>
                          <td class="task-cell">
                            <button id="${btnId}" class="saai-task-btn-popup" data-email="${emailData}">
                              <span style="margin-right: 4px;">+</span>Add to Tasks
                            </button>
                          </td>
                        </tr>
                      `;
                    }).join('')}
                  </tbody>
                </table>
              </div>
            `;
          }
          return '';
        }).join('')}
        <div style="text-align: center;">
          <button id="close-btn">Close</button>
        </div>
      </body>
    </html>
  `);
    tableWindow.document.close();
    
    // Attach event listeners after the document is loaded (CSP-compliant way)
    setTimeout(() => {
      try {
        // PRODUCTION: console.log('[SaAI] Attaching event listeners to popup buttons');
        
        // Track current button being processed
        let currentButton = null;
        
        // Get all task buttons in the popup
        const taskButtons = tableWindow.document.querySelectorAll('.saai-task-btn-popup');
        
        taskButtons.forEach(btn => {
          btn.addEventListener('click', async function() {
            const emailData = this.getAttribute('data-email');
            
            currentButton = this;
            // Disable button and show loading
            this.disabled = true;
            this.innerHTML = '<span style="margin-right: 4px;">⏳</span>Adding...';
            
            // Parse email data and add to tasks directly (we're already in parent context)
            try {
              const parsedData = JSON.parse(emailData);
              
              // Determine priority level
              let priorityLevel = 'medium';
              if (parsedData.priorityKey === 'high_priority_emails') priorityLevel = 'high';
              else if (parsedData.priorityKey === 'medium_priority') priorityLevel = 'medium';
              else if (parsedData.priorityKey === 'low_priority') priorityLevel = 'low';
              else if (parsedData.priorityKey === 'missed_or_ignored_emails') priorityLevel = 'low';
              else if (parsedData.priorityKey === 'already_replied_closed_threads') priorityLevel = 'low';
              
              // Add task directly
              await addEmailToTasks({
                from: parsedData.sender,
                context: parsedData.subject,
                taskName: parsedData.nextAction,
                priority: priorityLevel
              });
              
              // Success
            currentButton.innerHTML = '<span style="margin-right: 4px;">✓</span>Added';
            currentButton.style.background = '#16a34a';
            currentButton.style.cursor = 'default';
            currentButton = null;
            } catch (error) {
              // Error
            currentButton.disabled = false;
            currentButton.innerHTML = '<span style="margin-right: 4px;">+</span>Add to Tasks';
              alert('Failed to add task: ' + error.message);
            currentButton = null;
          }
          });
        });
        
        // Close button
        const closeBtn = tableWindow.document.getElementById('close-btn');
        if (closeBtn) {
          closeBtn.addEventListener('click', function() {
            tableWindow.close();
          });
        }
        
        // PRODUCTION: console.log('[SaAI] Event listeners attached successfully');
      } catch (error) {
        // PRODUCTION: console.error('[SaAI] Error attaching event listeners:', error);
      }
    }, 100);
    
    debugLog('Full table opened successfully');
    
  } catch (error) {
    debugError('Error opening full table:', error);
    alert('Error opening full table: ' + error.message);
  }
}

// === CREDIT TRACKING SYSTEM ===

// Credit costs for each feature
const CREDIT_COSTS = {
  THREAD_SUMMARIZATION: 5,
  WEB_SEARCH: 8,
  TASK_EXTRACTION: 10,
  ASK_QUESTION_INBOX: 10,
  INBOX_SUMMARIZATION: 15,
  VOICE_MODE: 0, // Varies by feature
  DELETE_ALL_DATA: 250
};

// Current user credit state
let userCredits = {
  total: 1000,
  used: 0,
  remaining: 1000
};

// Prompt classification patterns
const PROMPT_PATTERNS = {
  THREAD_SUMMARIZATION: [
    /summarize.*thread/i,
    /thread.*summary/i,
    /summarize.*email.*thread/i,
    /what.*happened.*thread/i,
    /thread.*overview/i
  ],
  WEB_SEARCH: [
    /search.*web/i,
    /google.*search/i,
    /web.*search/i,
    /search.*internet/i,
    /find.*online/i
  ],
  TASK_EXTRACTION: [
    /extract.*task/i,
    /find.*task/i,
    /what.*task/i,
    /task.*extraction/i,
    /show.*task/i
  ],
  ASK_QUESTION_INBOX: [
    /ask.*question/i,
    /question.*about/i,
    /help.*with/i,
    /explain.*email/i,
    /what.*does.*mean/i
  ],
  INBOX_SUMMARIZATION: [
    /summarize.*inbox/i,
    /inbox.*summary/i,
    /email.*summary/i,
    /summarize.*emails/i,
    /overview.*inbox/i
  ],
  DELETE_ALL_DATA: [
    /delete.*all.*data/i,
    /clear.*all.*data/i,
    /remove.*all.*data/i,
    /wipe.*data/i
  ]
};

// Classify user prompt to determine feature type and credit cost
function classifyPrompt(prompt) {
  const lowerPrompt = prompt.toLowerCase();
  
  // Check for specific patterns
  for (const [feature, patterns] of Object.entries(PROMPT_PATTERNS)) {
    for (const pattern of patterns) {
      if (pattern.test(lowerPrompt)) {
        return {
          feature: feature,
          cost: CREDIT_COSTS[feature],
          confidence: 'high'
        };
      }
    }
  }
  
  // Default classification based on context
  if (lowerPrompt.includes('summarize') || lowerPrompt.includes('summary')) {
    return {
      feature: 'INBOX_SUMMARIZATION',
      cost: CREDIT_COSTS.INBOX_SUMMARIZATION,
      confidence: 'medium'
    };
  }
  
  if (lowerPrompt.includes('search') || lowerPrompt.includes('find')) {
    return {
      feature: 'WEB_SEARCH',
      cost: CREDIT_COSTS.WEB_SEARCH,
      confidence: 'medium'
    };
  }
  
  // Default to ask question
  return {
    feature: 'ASK_QUESTION_INBOX',
    cost: CREDIT_COSTS.ASK_QUESTION_INBOX,
    confidence: 'low'
  };
}

// Check if user has enough credits
function hasEnoughCredits(requiredCredits) {
  // If it's a feature name, get the cost from CREDIT_COSTS
  const cost = typeof requiredCredits === 'string' ? CREDIT_COSTS[requiredCredits] : requiredCredits;
  // PRODUCTION: console.log('[SaAI-Content] Checking credits - required:', requiredCredits, 'cost:', cost, 'remaining:', userCredits.remaining);
  return userCredits.remaining >= cost;
}

// Deduct credits and update state
async function deductCredits(amount, feature) {
  // If amount is a feature name, get the cost from CREDIT_COSTS
  const cost = typeof amount === 'string' ? CREDIT_COSTS[amount] : amount;
  
  if (!hasEnoughCredits(cost)) {
    throw new Error('Insufficient credits');
  }
  
  userCredits.used += cost;
  userCredits.remaining -= cost;
  
  // Update storage
  await chrome.storage.local.set({ 
    userCredits: userCredits,
    lastCreditUpdate: Date.now()
  });
  
  // Update UI
  updateCreditDisplay();
  
  // Check for warnings
  checkCreditWarnings();
  
  // Send to backend
  await sendCreditUpdateToBackend(feature, amount);
  
  debugLog(`Credits deducted: ${amount} for ${feature}. Remaining: ${userCredits.remaining}`);
}

// Update credit display in UI
function updateCreditDisplay() {
  const creditsNumber = document.querySelector('.credits-number');
  const usageFill = document.querySelector('.usage-fill');
  const usageText = document.querySelector('.usage-text');
  
  // Settings sidebar elements
  const settingsUsageFill = document.getElementById('settings-usage-fill');
  const settingsUsageText = document.getElementById('settings-usage-text');
  
  // Modal elements
  const modalUsageNumber = document.getElementById('modal-usage-number');
  const modalUsageFill = document.getElementById('modal-usage-fill');
  
  const percentage = (userCredits.used / userCredits.total) * 100;
  
  // Update onboarding card
  if (creditsNumber) {
    creditsNumber.textContent = userCredits.remaining;
  }
  
  if (usageFill) {
    usageFill.style.width = `${percentage}%`;
  }
  
  if (usageText) {
    usageText.textContent = `${userCredits.used} / ${userCredits.total} requests used`;
  }
  
  // Update settings sidebar
  if (settingsUsageFill) {
    settingsUsageFill.style.width = `${percentage}%`;
  }
  
  if (settingsUsageText) {
    settingsUsageText.textContent = `${userCredits.used} / ${userCredits.total} requests used`;
  }
  
  // Update modal
  if (modalUsageNumber) {
    modalUsageNumber.textContent = userCredits.used;
  }
  
  if (modalUsageFill) {
    modalUsageFill.style.width = `${percentage}%`;
  }
}

// Check for credit warnings
function checkCreditWarnings() {
  if (userCredits.remaining <= 100 && userCredits.remaining > 0) {
    showCreditWarning('low');
  } else if (userCredits.remaining === 0) {
    showCreditLimitReached();
  }
}

// Show credit warning
function showCreditWarning(type) {
  let warningMessage;
  
  if (type === 'low') {
    warningMessage = `⚠️ You have ${userCredits.remaining} credits remaining. Consider upgrading or contact devang@saai.dev for more credits.`;
  } else if (type === 'insufficient') {
    warningMessage = `🚨 Insufficient credits! You need more credits to use this feature. Contact devang@saai.dev for more credits.`;
  } else {
    warningMessage = `🚨 You've reached your credit limit! Please upgrade or contact devang@saai.dev to continue using Sa.AI.`;
  }
  
  showStatus(warningMessage, 'warning');
  
  // Show modal for low credits
  if (type === 'low') {
    showCreditWarningModal();
  } else if (type === 'insufficient') {
    showInsufficientCreditsModal();
  }
}

// Show credit limit reached modal
function showCreditLimitReached() {
  const modal = document.createElement('div');
  modal.className = 'task-modal-overlay';
  modal.innerHTML = `
    <div class="task-modal-content">
      <div class="task-modal-header">
        <h3>🚨 Credit Limit Reached</h3>
        <button class="task-modal-close" onclick="this.closest('.task-modal-overlay').remove()">×</button>
      </div>
      <div class="task-modal-body">
        <div style="text-align: center; padding: 20px;">
          <div style="font-size: 48px; margin-bottom: 16px;">💳</div>
          <div style="font-size: 18px; color: var(--saai-text-primary); margin-bottom: 16px; font-weight: 600;">
            You've used all 1000 credits!
          </div>
          <div style="font-size: 14px; color: var(--saai-text-secondary); margin-bottom: 24px; line-height: 1.5;">
            To continue using Sa.AI, you can:
            <br>• Contact devang@saai.dev for more credits
            <br>• Wait for credit reset (if applicable)
            <br>• Upgrade your plan
          </div>
        </div>
        <div style="display: flex; gap: 12px;">
          <button style="flex: 1; padding: 12px; background: var(--saai-accent); border: 1px solid var(--saai-border); border-radius: var(--saai-border-radius); cursor: pointer; font-family: inherit;" onclick="this.closest('.task-modal-overlay').remove()">Close</button>
          <button style="flex: 1; padding: 12px; background: var(--saai-primary); color: white; border: none; border-radius: var(--saai-border-radius); cursor: pointer; font-family: inherit;" onclick="window.open('mailto:devang@saai.dev?subject=Credit Request&body=Hi Devang, I need more credits for Sa.AI. Please let me know how to proceed.', '_blank')">Contact Devang</button>
        </div>
      </div>
    </div>
  `;
  
  document.body.appendChild(modal);
}

// Show credit warning modal
function showCreditWarningModal() {
  const modal = document.createElement('div');
  modal.className = 'task-modal-overlay';
  modal.innerHTML = `
    <div class="task-modal-content">
      <div class="task-modal-header">
        <h3>⚠️ Low Credits Warning</h3>
        <button class="task-modal-close" onclick="this.closest('.task-modal-overlay').remove()">×</button>
      </div>
      <div class="task-modal-body">
        <div style="text-align: center; padding: 20px;">
          <div style="font-size: 48px; margin-bottom: 16px;">⚠️</div>
          <div style="font-size: 18px; color: var(--saai-text-primary); margin-bottom: 16px; font-weight: 600;">
            Only ${userCredits.remaining} credits remaining!
          </div>
          <div style="font-size: 14px; color: var(--saai-text-secondary); margin-bottom: 24px; line-height: 1.5;">
            You're running low on credits. Consider:
            <br>• Contacting devang@saai.dev for more credits
            <br>• Using features more efficiently
            <br>• Planning your usage
          </div>
        </div>
        <div style="display: flex; gap: 12px;">
          <button style="flex: 1; padding: 12px; background: var(--saai-accent); border: 1px solid var(--saai-border); border-radius: var(--saai-border-radius); cursor: pointer; font-family: inherit;" onclick="this.closest('.task-modal-overlay').remove()">Continue</button>
          <button style="flex: 1; padding: 12px; background: var(--saai-primary); color: white; border: none; border-radius: var(--saai-border-radius); cursor: pointer; font-family: inherit;" onclick="window.open('mailto:devang@saai.dev?subject=Credit Request&body=Hi Devang, I need more credits for Sa.AI. Please let me know how to proceed.', '_blank')">Request Credits</button>
        </div>
      </div>
    </div>
  `;
  
  document.body.appendChild(modal);
}

// Show insufficient credits modal
function showInsufficientCreditsModal() {
  const modal = document.createElement('div');
  modal.className = 'task-modal-overlay';
  modal.innerHTML = `
    <div class="task-modal-content">
      <div class="task-modal-header">
        <h3>🚨 Insufficient Credits</h3>
        <button class="task-modal-close" onclick="this.closest('.task-modal-overlay').remove()">×</button>
      </div>
      <div class="task-modal-body">
        <div style="text-align: center; padding: 20px;">
          <div style="font-size: 48px; margin-bottom: 16px;">💳</div>
          <div style="font-size: 18px; color: var(--saai-text-primary); margin-bottom: 16px; font-weight: 600;">
            Not enough credits!
          </div>
          <div style="font-size: 14px; color: var(--saai-text-secondary); margin-bottom: 24px; line-height: 1.5;">
            You have ${userCredits.remaining} credits remaining, but this feature requires more.
            <br><br>
            To continue using Sa.AI, contact devang@saai.dev for more credits.
          </div>
        </div>
        <div style="display: flex; gap: 12px;">
          <button style="flex: 1; padding: 12px; background: var(--saai-accent); border: 1px solid var(--saai-border); border-radius: var(--saai-border-radius); cursor: pointer; font-family: inherit;" onclick="this.closest('.task-modal-overlay').remove()">Close</button>
          <button style="flex: 1; padding: 12px; background: var(--saai-primary); color: white; border: none; border-radius: var(--saai-border-radius); cursor: pointer; font-family: inherit;" onclick="window.open('mailto:devang@saai.dev?subject=Credit Request&body=Hi Devang, I need more credits for Sa.AI. Please let me know how to proceed.', '_blank')">Request Credits</button>
        </div>
      </div>
    </div>
  `;
  
  document.body.appendChild(modal);
}
async function sendCreditUpdateToBackend(feature, amount) {
  // DEPRECATED: This function is no longer used
  // Credit tracking is now handled by sendCreditTrackingToWebhook() via background script
  // to avoid CORS issues
  debugLog('sendCreditUpdateToBackend called (deprecated, skipping)');
  return;
}

// Get JWT token from storage
async function getJWTToken() {
  const storage = await chrome.storage.local.get(['jwtToken']);
  return storage.jwtToken;
}

// Get User ID from storage
async function getUserId() {
  const storage = await chrome.storage.local.get(['userId']);
  return storage.userId;
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

// Auto-refresh token by requesting background script to handle refresh
async function ensureValidJWTToken() {
  const jwtToken = await getJWTToken();
  
  if (!jwtToken || await isJWTTokenExpired(jwtToken)) {
    debugLog('JWT token expired or missing, requesting refresh from background script');
    
    try {
      // Request background script to refresh the token
      const response = await chrome.runtime.sendMessage({
        action: 'refreshToken'
      });
      
      if (response.success) {
        debugLog('Token refreshed successfully via background script');
        return await getJWTToken(); // Get the newly stored token
      } else {
        throw new Error('Token refresh failed');
      }
    } catch (error) {
      debugError('Token refresh failed:', error);
      throw error;
    }
  }
  
  return jwtToken;
}

// Auto-retry request with JWT refresh on 401/402 errors
async function safeRequestWithRetry(url, options) {
  try {
    const response = await safeRequest(url, options);
    
    // If 401/402 error, automatically refresh token and retry
    if (!response.ok && (response.status === 401 || response.status === 402)) {
      debugLog('Request failed with 401/402, automatically refreshing token and retrying');
      
      try {
        // Request background script to refresh token
        const refreshResponse = await chrome.runtime.sendMessage({
          action: 'refreshToken'
        });
        
        if (refreshResponse.success) {
          debugLog('Token refreshed successfully, retrying request');
          
          // Update Authorization header with new token
          const newToken = await ensureValidJWTToken();
          const retryOptions = {
            ...options,
            headers: {
              ...options.headers,
              'Authorization': `Bearer ${newToken}`
            }
          };
          
          // Retry the request
          const retryResponse = await safeRequest(url, retryOptions);
          
          if (retryResponse.ok) {
            debugLog('Request retry successful after token refresh');
            return retryResponse;
          } else {
            debugLog('Retry still failed, attempting second retry');
            
            // Second retry with fresh token validation
            const finalToken = await ensureValidJWTToken();
            const finalRetryOptions = {
              ...options,
              headers: {
                ...options.headers,
                'Authorization': `Bearer ${finalToken}`
              }
            };
            
            const finalRetryResponse = await safeRequest(url, finalRetryOptions);
            
            if (finalRetryResponse.ok) {
              debugLog('Second retry successful');
              return finalRetryResponse;
            } else {
              throw new Error(`Request failed after token refresh attempts: ${finalRetryResponse.status}`);
            }
          }
        } else {
          throw new Error('Token refresh failed');
        }
      } catch (refreshError) {
        debugError('Automatic token refresh failed:', refreshError);
        throw new Error('Service temporarily unavailable. Please try again.');
      }
    }
    
    return response;
  } catch (error) {
    debugError('Request with retry failed:', error);
    throw error;
  }
}

// Load user credits from storage
async function loadUserCredits() {
  const storage = await chrome.storage.local.get(['userCredits']);
  if (storage.userCredits) {
    userCredits = storage.userCredits;
  }
  updateCreditDisplay();
}

// Initialize credit system
async function initializeCreditSystem() {
  await loadUserCredits();
  debugLog('Credit system initialized:', userCredits);
}

// OAuth flow debounce flag
let oauthFlowInProgress = false;

function startOAuthFlow() {
  // PRODUCTION: console.log('[SaAI-Content] ===== startOAuthFlow() called =====');
  debugLog('Starting OAuth flow');
  
  // Prevent multiple concurrent OAuth flows
  if (oauthFlowInProgress) {
    debugLog('OAuth flow already in progress, ignoring request');
    showStatus('Authentication already in progress. Please wait...', 'info');
    return;
  }
  
  // Check if already connected
  chrome.storage.local.get(['userId', 'isConnected'], (result) => {
    if (result.userId || result.isConnected) {
      debugLog('Already connected, skipping OAuth flow');
      showStatus('Already connected to Sa.AI!', 'info');
      return;
    }
    
    // Set flag to prevent concurrent flows
    oauthFlowInProgress = true;
  
  // Show OAuth loader
  showOAuthLoader();
  
  // Send message to background script
    // PRODUCTION: console.log('[SaAI-Content] Sending OAuth request to background script...');
  debugLog('Sending OAuth request to background script...');
  
  const message = {
    action: 'sendToN8N',
    data: {
      endpoint: 'oauth',
      payload: { context: 'OutlookConnectClicked' }
    }
  };
  
    // PRODUCTION: console.log('[SaAI-Content] OAuth message being sent');
  
  chrome.runtime.sendMessage(message, (response) => {
      // Reset flag after response
      oauthFlowInProgress = false;
      
      // Extension was reloaded - content script is orphaned (no receiving end)
      if (chrome.runtime.lastError && isExtensionContextInvalidated(chrome.runtime.lastError)) {
        showStatus(SAAI_REFRESH_PAGE_MSG, 'error');
        showRefreshPageMessage(null); // Show in sidebar chat area if open
        return;
      }
      
      // PRODUCTION: console.log('[SaAI-Content] OAuth response received:', response);
    debugLog('OAuth response received:', response);
    if (response?.success) {
      debugLog('OAuth success response:', response.data);
      showOAuthSuccess();
    } else {
      debugError('OAuth failed:', response?.error);
      
      // Check if user cancelled the OAuth flow
      if (response?.error && response.error.includes('OAUTH_CANCELLED')) {
        showStatus('Authentication cancelled. Click "Connect" to try again.', 'info');
      }
        // Check if OAuth is already in progress
        else if (response?.error && response.error.includes('already in progress')) {
          showStatus('⏳ Authentication in progress. Please wait...', 'info');
        }
      // Check if user is not in whitelist
      else if (response?.error && response.error.includes('NOT_WHITELISTED')) {
        showStatus('Thank you for your interest! Your email is not yet registered for early access. Please contact our team to request access.', 'error');
      }
      // Check if this is an access denied error
      else if (response?.error && response.error.includes('Access Denied')) {
        showStatus('Access Denied: Your email is not authorized. Please contact admin for access.', 'error');
      } else {
      showStatus('OAuth failed. Please try again.', 'error');
      }
    }
    });
  });
}

async function isOutlookConnected() {
  const { userId, isConnected, provider } = await chrome.storage.local.get(['userId', 'isConnected', 'provider']);
  // Check if connected and provider is Outlook (or not set, for backward compatibility)
  return !!(userId || isConnected) && (provider === 'outlook' || !provider);
}

async function debugConnectionStatus() {
  const storage = await chrome.storage.local.get(['userId', 'isConnected', 'oauthData']);
  debugLog('Debug connection status:', storage);
  return storage;
}

// === UI Helpers ===

function showOAuthLoader() {
  const connectContent = document.querySelector('.saai-connect-content');
  if (connectContent) {
    connectContent.innerHTML = `
      <div class="loader">Connecting to Google...</div>
    `;
  }
}

function showOAuthSuccess() {
  const connectContent = document.querySelector('.saai-connect-content');
  if (connectContent) {
    connectContent.innerHTML = `
      <div class="saai-connect-icon">✅</div>
      <h2 class="saai-connect-heading">Connected Successfully!</h2>
      <p class="saai-connect-description">Your Outlook is now connected to Sa.AI Assistant.</p>
    `;
  }
  
  // Update sidebar content after a delay
  setTimeout(() => {
    updateSidebarContent();
  }, 2000);
}

function showStatus(message, type) {
  // Use the new elegant toast notification system
  showToast(message, type);
}

// Elegant Toast Notification System
function showToast(message, type = 'info', duration = 5000) {
  // Remove any existing toast
  const existingToast = document.querySelector('.saai-toast-notification');
  if (existingToast) {
    existingToast.remove();
  }
  
  // Create toast container
  const toast = document.createElement('div');
  toast.className = `saai-toast-notification saai-toast-${type}`;
  
  // Icon based on type
  let icon = '';
  switch(type) {
    case 'success':
      icon = '✓';
      break;
    case 'error':
      icon = '✕';
      break;
    case 'warning':
      icon = '⚠';
      break;
    case 'info':
    default:
      icon = 'ℹ';
      break;
  }
  
  toast.innerHTML = `
    <div class="saai-toast-icon">${icon}</div>
    <div class="saai-toast-content">
      <div class="saai-toast-message">${message}</div>
    </div>
    <button class="saai-toast-close" onclick="this.parentElement.remove()">×</button>
  `;
  
  // Add to body
  document.body.appendChild(toast);
  
  // Trigger animation
    setTimeout(() => {
    toast.classList.add('saai-toast-show');
  }, 10);
  
  // Auto-remove after duration
  setTimeout(() => {
    toast.classList.remove('saai-toast-show');
    setTimeout(() => {
      if (toast.parentNode) {
        toast.remove();
      }
    }, 300);
  }, duration);
}

function injectSuggestions() {
  const chatArea = document.getElementById(CHAT_AREA_ID);
  if (!chatArea) return;
  
  // Remove existing suggestions if any
  const existingSuggestions = chatArea.querySelector('#saai-suggestions');
  if (existingSuggestions) {
    existingSuggestions.remove();
  }
  
  // Check if user has already sent a message (hide suggestions after first use)
  const messages = chatArea.querySelectorAll('.message');
  if (messages.length > 0) {
    return; // Don't show suggestions if user has already interacted
  }
  
  const suggestions = [
    {
      text: 'Summarize my inbox',
      icon: `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
        <path d="M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z"/>
        <polyline points="22,6 12,13 2,6"/>
      </svg>`
    },
    {
      text: 'Summarize this thread',
      icon: `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
        <path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/>
        <path d="M13 8H7"/>
        <path d="M17 12H7"/>
      </svg>`
    },
    {
      text: 'Draft an email',
      icon: `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
        <polyline points="9,11 12,14 22,4"/>
        <path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11"/>
      </svg>`
    },
    {
      text: 'Show follow-ups needed',
      icon: `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
        <circle cx="12" cy="12" r="10"/>
        <polyline points="12,6 12,12 16,14"/>
      </svg>`
    },
  ];
  
  const suggestionsDiv = document.createElement('div');
  suggestionsDiv.id = 'saai-suggestions';
  suggestionsDiv.className = 'saai-suggestions';
  
  suggestions.forEach(suggestion => {
    const button = document.createElement('button');
    button.className = 'saai-suggestion';
    button.innerHTML = `
      <span class="saai-suggestion-icon">${suggestion.icon}</span>
      <span class="saai-suggestion-text">${suggestion.text}</span>
    `;
    button.addEventListener('click', () => {
      const input = document.getElementById('chat-input');
      if (input) {
        input.value = suggestion.text;
        input.focus();
      }
    });
    suggestionsDiv.appendChild(button);
  });
  
  chatArea.appendChild(suggestionsDiv);
}

// Check if credit consumption should be tracked (skip validation messages)
function shouldTrackCreditConsumption(responseData, userMessage) {
  // If no response data, don't track
  if (!responseData) {
    return false;
  }
  
  // Extract response text for analysis
  let responseText = '';
  
  if (typeof responseData === 'string') {
    responseText = responseData;
  } else if (typeof responseData === 'object') {
    // Handle different response formats
    if (responseData.message) {
      responseText = responseData.message;
    } else if (responseData.response) {
      responseText = responseData.response;
    } else if (responseData.text) {
      responseText = responseData.text;
    } else if (Array.isArray(responseData) && responseData.length > 0) {
      responseText = responseData[0].message || responseData[0].response || responseData[0].text || '';
    }
  }
  
  // Convert to lowercase for easier matching
  const lowerResponseText = responseText.toLowerCase();
  
  // List of validation messages that should NOT count as credit consumption
  const validationMessages = [
    'please open a thread to process',
    'please open a thread',
    'open a thread',
    'thread not found',
    'no thread selected',
    'please select a thread',
    'connect your outlook',
    'please connect your outlook',
    'authentication required',
    'login required',
    'please login',
    'invalid request',
    'error occurred',
    'something went wrong',
    'try again',
    'please try again'
  ];
  
  // Check if response contains any validation message
  const isValidationMessage = validationMessages.some(msg => 
    lowerResponseText.includes(msg.toLowerCase())
  );
  
  // PRODUCTION: console.log('[SaAI-Content] Credit tracking check:', {
  //   responseLength: responseText.length,
  //   isValidationMessage,
  //   shouldTrack: !isValidationMessage
  // });
  
  // Track credits only if it's NOT a validation message
  return !isValidationMessage;
}

// === Credit Tracking ===

// Send credit tracking data to webhook via background script (to avoid CORS)
async function sendCreditTrackingToWebhook(userId, prompt, creditsUsed) {
  try {
    // PRODUCTION: console.log('[SaAI-Content] ===== sendCreditTrackingToWebhook called =====');
    // PRODUCTION: console.log('[SaAI-Content] Parameters received:', { 
    //   creditsUsed: creditsUsed,
    //   promptLength: prompt ? prompt.length : 0
    // });
    
    // Validate parameters
    if (!userId) {
      // PRODUCTION: console.error('[SaAI-Content] Error: userId is missing');
      return null;
    }
    
    if (!prompt || prompt.trim() === '') {
      // PRODUCTION: console.error('[SaAI-Content] Error: prompt is empty or missing');
      return null;
    }
    
    if (!creditsUsed || creditsUsed <= 0) {
      // PRODUCTION: console.error('[SaAI-Content] Error: creditsUsed is invalid:', creditsUsed);
      return null;
    }
    
    const payload = {
      userId: userId,
      prompt: prompt.trim(),
      creditsUsed: creditsUsed
    };
    
    // PRODUCTION: console.log('[SaAI-Content] Sending payload to background script for credit tracking - credits:', creditsUsed);
    
    // Send through background script to avoid CORS issues
    const response = await chrome.runtime.sendMessage({
      action: 'trackCredits',
      data: payload
    });
    
    // PRODUCTION: console.log('[SaAI-Content] Credit tracking response from background:', response);
    
    if (!response?.success) {
      throw new Error(response?.error || 'Credit tracking failed');
    }
    
    const responseData = response.data;
    // PRODUCTION: console.log('[SaAI-Content] Credit tracking data received:', responseData);
    
    return responseData;
    
  } catch (error) {
    // PRODUCTION: console.error('[SaAI-Content] Credit tracking error:', error);
    return null;
  }
}

// Update UI with total credits consumed
async function updateUIWithTotalCredits(creditData) {
  if (!creditData || typeof creditData.totalCreditsConsumed !== 'number') {
    // PRODUCTION: console.warn('[SaAI-Content] Invalid credit data received:', creditData);
    return;
  }
  
  // PRODUCTION: console.log('[SaAI-Content] Updating UI with total credits consumed:', creditData.totalCreditsConsumed);
  
  // Update userCredits object
  userCredits.used = creditData.totalCreditsConsumed;
  userCredits.remaining = userCredits.total - userCredits.used;
  
  // Update storage
  await chrome.storage.local.set({ 
    userCredits: userCredits,
    lastCreditUpdate: Date.now()
  });
  
  // Update UI display
  updateCreditDisplay();
  
  // PRODUCTION: console.log('[SaAI-Content] UI updated - Used:', userCredits.used, 'Remaining:', userCredits.remaining);
}

// === Task Management ===

async function showTaskModal() {
  // PRODUCTION: console.log('[SaAI-Content] ===== showTaskModal() called =====');
  
  // Check credits before showing task modal
  // PRODUCTION: console.log('[SaAI-Content] Checking credits, remaining:', userCredits.remaining);
  if (userCredits.remaining === 0) {
    // PRODUCTION: console.log('[SaAI-Content] No credits remaining, showing insufficient credits modal');
    showInsufficientCreditsModal();
    return;
  }
  
  // PRODUCTION: console.log('[SaAI-Content] Checking if has enough credits for TASK_EXTRACTION');
  if (!hasEnoughCredits('TASK_EXTRACTION')) {
    // PRODUCTION: console.log('[SaAI-Content] Not enough credits for TASK_EXTRACTION, showing insufficient credits modal');
    showInsufficientCreditsModal();
    return;
  }
  
  // PRODUCTION: console.log('[SaAI-Content] Deducting credits for TASK_EXTRACTION');
  // Deduct credits for task extraction
  deductCredits('TASK_EXTRACTION');
  
  // PRODUCTION: console.log('[SaAI-Content] Creating task modal HTML');
  
  // Show loading state first
  const modal = document.createElement('div');
  modal.id = 'task-modal';
  modal.className = 'task-modal-overlay';
  
  modal.innerHTML = `
    <div class="task-modal-content">
      <div class="task-modal-header">
        <h3 class="task-modal-title">
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <polyline points="9,11 12,14 22,4"/>
            <path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11"/>
          </svg>
          Task Management
        </h3>
        <button class="task-modal-close-btn">×</button>
      </div>
      <div class="task-modal-body">
        <div class="task-loading">
          <div class="task-loading-spinner"></div>
          <p>Extracting tasks from your emails...</p>
        </div>
      </div>
    </div>
  `;
  
  // PRODUCTION: console.log('[SaAI-Content] Appending modal to document.body');
  document.body.appendChild(modal);
  // PRODUCTION: console.log('[SaAI-Content] Modal appended successfully');
  
  // Add close button functionality
  const closeBtn = modal.querySelector('.task-modal-close-btn');
  // PRODUCTION: console.log('[SaAI-Content] Close button found:', closeBtn);
  if (closeBtn) {
    closeBtn.addEventListener('click', () => closeTaskModal(modal));
  }
  
  // PRODUCTION: console.log('[SaAI-Content] Starting webhook call...');
  
  try {
    // Get user ID from storage
    const { userId } = await chrome.storage.local.get(['userId']);
    
    if (!userId) {
      throw new Error('User ID not found. Please connect your Outlook account first.');
    }
    
    // Call the webhook to extract tasks
    const response = await safeRequest('https://connector.saai.dev/webhook/task-management-outlook', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${await ensureValidJWTToken()}`
      },
      body: JSON.stringify({
        userId: userId,
        context: 'ExtractTasksList'
      })
    });
    
    if (!response.ok) {
      throw new Error(`Webhook request failed: ${response.status} ${response.statusText}`);
    }
    
    // Check if response has content before parsing JSON
    const responseText = await response.text();
    // PRODUCTION: console.log('[SaAI-Content] Webhook response text:', responseText);
    
    if (!responseText || responseText.trim() === '') {
      throw new Error('Webhook returned empty response');
    }
    
    let responseData;
    try {
      responseData = JSON.parse(responseText);
    } catch (parseError) {
      // PRODUCTION: console.error('[SaAI-Content] JSON parse error:', parseError);
      // PRODUCTION: console.error('[SaAI-Content] Response text that failed to parse:', responseText);
      throw new Error(`Invalid JSON response from webhook: ${parseError.message}`);
    }
    
    // Update modal with tasks from webhook
    updateTaskModalWithData(modal, responseData);
    
    // Send credit tracking data to webhook for task extraction
    try {
      const creditData = await sendCreditTrackingToWebhook(userId, 'Task Extraction', 10);
      if (creditData) {
        await updateUIWithTotalCredits(creditData);
      }
    } catch (creditError) {
      debugError('Credit tracking for task extraction failed:', creditError);
    }
    
  } catch (error) {
    debugError('Task extraction failed:', error);
    
    // Show error state
    modal.querySelector('.task-modal-body').innerHTML = `
      <div class="task-error">
        <div class="task-error-icon">
          <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <circle cx="12" cy="12" r="10"/>
            <line x1="15" y1="9" x2="9" y2="15"/>
            <line x1="9" y1="9" x2="15" y2="15"/>
          </svg>
        </div>
        <div class="task-error-text">Failed to load tasks</div>
        <p style="color: var(--saai-text-secondary); font-size: 12px; margin-top: 8px;">
          ${error.message}
        </p>
        <button class="task-retry-btn">Retry</button>
      </div>
    `;
    
    // Add retry functionality
    const retryBtn = modal.querySelector('.task-retry-btn');
    if (retryBtn) {
      retryBtn.addEventListener('click', () => {
        closeTaskModal(modal);
        showTaskModal();
      });
    }
  }
}

function updateTaskModalWithData(modal, data) {
  // Check if data contains tasks array
  const webhookTasks = Array.isArray(data) ? data : (data.tasks || []);
  
  // Debug: Log webhook tasks to see what fields they have
  // PRODUCTION: console.log('[Task Modal] Webhook tasks received - count:', webhookTasks.length);
  if (webhookTasks.length > 0) {
    // PRODUCTION: console.log('[Task Modal] First webhook task fields:', Object.keys(webhookTasks[0]));
  }
  
  // Get manual tasks from localStorage
  const manualTasks = getTasks().filter(task => task.isManual);
  
  // Create a map to avoid duplicates based on task content
  const taskMap = new Map();
  
  // Add webhook tasks first
  webhookTasks.forEach(task => {
    const taskKey = task.task || task.text || task.description || task.title || '';
    if (taskKey && !taskMap.has(taskKey)) {
      taskMap.set(taskKey, { ...task, source: 'webhook' });
    }
  });
  
  // Add manual tasks, avoiding duplicates
  manualTasks.forEach(task => {
    const taskKey = task.task || task.text || task.description || task.title || '';
    if (taskKey && !taskMap.has(taskKey)) {
      taskMap.set(taskKey, { ...task, source: 'manual' });
    }
  });
  
  // Convert map values back to array
  const allTasks = Array.from(taskMap.values());
  
  // Debug: Log all tasks with their sources and from fields
  // PRODUCTION: console.log('[Task Modal] All tasks after merging:', allTasks);
  allTasks.forEach((task, idx) => {
    // PRODUCTION: console.log(`[Task Modal] Task ${idx}:`, {
    //   task: task.task || task.text,
    //   source: task.source,
    //   from: task.from,
    //   hasFrom: !!task.from
    // });
  });
  
  modal.querySelector('.task-modal-body').innerHTML = `
    <div class="task-list">
      ${allTasks.length === 0 ? `
        <div class="no-tasks">
          <div class="no-tasks-icon">
            <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
              <polyline points="9,11 12,14 22,4"/>
              <path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11"/>
            </svg>
          </div>
          <div class="no-tasks-text">No tasks found</div>
          <p style="color: var(--saai-text-secondary); font-size: 12px; margin-top: 8px;">
            Tasks will be automatically extracted from your emails or add them manually below
          </p>
        </div>
      ` : ''}
      ${allTasks.map((task, index) => {
        // Determine the "added by" text based on source
        let addedByText = '';
        // PRODUCTION: console.log(`[Task Modal Render] Task ${index} - source:`, task.source);
        
        if (task.source === 'manual') {
          // For manual tasks, check if 'from' was provided
          const email = task.from || task.email || task.sender || task.from_email || task.fromEmail;
          if (email) {
            const truncatedEmail = email.length > 25 ? email.substring(0, 22) + '...' : email;
            addedByText = `from ${truncatedEmail}`;
          } else {
            addedByText = 'added by you';
          }
        } else {
          // Check for email in multiple possible fields
          const email = task.from || task.email || task.sender || task.from_email || task.fromEmail;
          if (email) {
            const truncatedEmail = email.length > 25 ? email.substring(0, 22) + '...' : email;
            addedByText = `from ${truncatedEmail}`;
          }
        }
        
        return `
        <div class="task-item" data-id="${task.id || `task-${index}`}" data-source="${task.source || 'unknown'}">
          <div class="task-content">
            <input type="checkbox" class="task-checkbox" ${task.completed ? 'checked' : ''}>
            <div class="task-info">
            <span class="task-text">${task.task || task.text || task.description || task.title || 'Untitled Task'}</span>
              ${addedByText ? `<span class="task-added-by" title="${task.from || ''}">${addedByText}</span>` : ''}
            </div>
            ${task.priority ? `<span class="task-priority ${task.priority.toLowerCase()}">${task.priority}</span>` : ''}
            <button class="task-delete-btn">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                <line x1="18" y1="6" x2="6" y2="18"/>
                <line x1="6" y1="6" x2="18" y2="18"/>
              </svg>
            </button>
          </div>
        </div>
      `;
      }).join('')}
    </div>
    
    <!-- Manual Task Addition Section -->
    <div class="add-task-section">
      <div class="task-input-row">
      <input type="text" class="new-task-input" placeholder="Add new task manually...">
        <input type="text" class="task-from-input" placeholder="From (optional)...">
      </div>
      <div class="task-controls-row">
      <select class="task-priority-select">
        <option value="low">Low</option>
        <option value="medium" selected>Medium</option>
        <option value="high">High</option>
      </select>
      <button class="add-task-btn">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
          <line x1="12" y1="5" x2="12" y2="19"/>
          <line x1="5" y1="12" x2="19" y2="12"/>
        </svg>
      </button>
      </div>
    </div>
    
    <div class="task-actions">
      <button class="task-refresh-btn">
        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
          <polyline points="23,4 23,10 17,10"/>
          <polyline points="1,20 1,14 7,14"/>
          <path d="M20.49 9A9 9 0 0 0 5.64 5.64L1 10m22 4l-4.64 4.36A9 9 0 0 1 3.51 15"/>
        </svg>
        Refresh Tasks
      </button>
      <button class="task-modal-ok-btn">Done</button>
    </div>
  `;
  
  // Add event listeners for the updated modal
  addTaskModalEventListeners(modal);
}

function addTaskModalEventListeners(modal) {
  const closeBtn = modal.querySelector('.task-modal-close-btn');
  const okBtn = modal.querySelector('.task-modal-ok-btn');
  const refreshBtn = modal.querySelector('.task-refresh-btn');
  const addBtn = modal.querySelector('.add-task-btn');
  const input = modal.querySelector('.new-task-input');
  const prioritySelect = modal.querySelector('.task-priority-select');
  
  // Close buttons
  [closeBtn, okBtn].forEach(btn => {
    if (btn) {
      btn.addEventListener('click', () => closeTaskModal(modal));
    }
  });
  
  // Refresh tasks button
  if (refreshBtn) {
    refreshBtn.addEventListener('click', () => {
      closeTaskModal(modal);
      showTaskModal();
    });
  }
  
  // Add manual task
  const fromInput = modal.querySelector('.task-from-input');
  if (addBtn && input && prioritySelect) {
    addBtn.addEventListener('click', async () => {
      const text = input.value.trim();
      const fromValue = fromInput ? fromInput.value.trim() : '';
      if (text) {
        // Disable button during webhook call
        addBtn.disabled = true;
        addBtn.innerHTML = `
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <circle cx="12" cy="12" r="10"/>
            <polyline points="9,11 12,14 22,4"/>
          </svg>
        `;
        
        try {
          await addManualTask(text, prioritySelect.value, fromValue);
          input.value = '';
          if (fromInput) fromInput.value = '';
          // Just close the modal - the task is already saved locally and will show on next open
          closeTaskModal(modal);
        } catch (error) {
          debugError('Failed to add task:', error);
          // Re-enable button on error
          addBtn.disabled = false;
          addBtn.innerHTML = `
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
              <line x1="12" y1="5" x2="12" y2="19"/>
              <line x1="5" y1="12" x2="19" y2="12"/>
            </svg>
          `;
        }
      }
    });
    
    input.addEventListener('keypress', (e) => {
      if (e.key === 'Enter') {
        addBtn.click();
      }
    });
  }
  
  // Task interactions (checkboxes and delete buttons)
  const taskItems = modal.querySelectorAll('.task-item');
  taskItems.forEach(item => {
    const checkbox = item.querySelector('.task-checkbox');
    const deleteBtn = item.querySelector('.task-delete-btn');
    
    if (checkbox) {
      checkbox.addEventListener('change', () => {
        const taskId = item.dataset.id;
        
        // Check if it's a manual task (stored locally)
        const taskSource = item.dataset.source;
        if (taskSource === 'manual') {
          toggleTaskComplete(taskId);
          debugLog('Manual task completion toggled:', taskId);
        } else {
          // Webhook task - just log for now
          debugLog('Webhook task completion toggled:', taskId);
          // TODO: Could send completion update to webhook later
        }
      });
    }
    
    if (deleteBtn) {
      deleteBtn.addEventListener('click', async () => {
        const taskId = item.dataset.id;
        const taskText = item.querySelector('.task-text').textContent;
        
        // Disable delete button during webhook call
        deleteBtn.disabled = true;
        deleteBtn.innerHTML = `
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <circle cx="12" cy="12" r="10"/>
            <polyline points="9,11 12,14 22,4"/>
          </svg>
        `;
        
        try {
          // Get user ID from storage
          const { userId } = await chrome.storage.local.get(['userId']);
          
          if (!userId) {
            throw new Error('User ID not found. Please connect your Outlook account first.');
          }
          
          // Send delete task request to webhook
          const response = await safeRequestWithRetry('https://connector.saai.dev/webhook/task-management-outlook', {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json',
              'Authorization': `Bearer ${await ensureValidJWTToken()}`
            },
            body: JSON.stringify({
              userId: userId,
              context: 'DeleteTask',
              taskName: taskText
            })
          });
          
          if (!response.ok) {
            throw new Error(`Webhook request failed: ${response.status} ${response.statusText}`);
          }
          
          debugLog('Task deletion sent to webhook successfully');
          
          // Remove from local storage if it's a manual task
          const taskSource = item.dataset.source;
          if (taskSource === 'manual') {
            removeTask(taskId);
            debugLog('Manual task deleted from local storage:', taskId);
          } else {
            debugLog('Webhook task deletion confirmed:', taskId);
          }
          
          // Remove from UI
          item.remove();
          
        } catch (error) {
          debugError('Failed to delete task via webhook:', error);
          
          // Re-enable delete button on error
          deleteBtn.disabled = false;
          deleteBtn.innerHTML = `
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
              <line x1="18" y1="6" x2="6" y2="18"/>
              <line x1="6" y1="6" x2="18" y2="18"/>
            </svg>
          `;
          
          // Still remove from local storage if it's a manual task
          const taskSource = item.dataset.source;
          if (taskSource === 'manual') {
            removeTask(taskId);
            debugLog('Manual task deleted locally (webhook failed):', taskId);
            item.remove();
          }
          
          // Could show user notification about sync failure
          // alert('Task deleted locally but failed to sync with backend. Please try again later.');
        }
      });
    }
  });
}

function closeTaskModal(modal) {
  if (modal && modal.parentNode) {
    modal.remove();
  }
}

// === Feedback Modal ===

function showFeedbackModal() {
  // Remove existing modal if any
  const existingModal = document.querySelector('.saai-feedback-modal');
  if (existingModal) existingModal.remove();
  
  const modal = document.createElement('div');
  modal.className = 'saai-modal saai-feedback-modal';
  modal.innerHTML = `
    <div class="saai-modal-content saai-feedback-modal-content">
      <div class="saai-modal-header">
        <h2 class="saai-modal-title">
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <path d="M21 11.5a8.38 8.38 0 0 1-.9 3.8 8.5 8.5 0 0 1-7.6 4.7 8.38 8.38 0 0 1-3.8-.9L3 21l1.9-5.7a8.38 8.38 0 0 1-.9-3.8 8.5 8.5 0 0 1 4.7-7.6 8.38 8.38 0 0 1 3.8-.9h.5a8.48 8.48 0 0 1 8 8v.5z"/>
          </svg>
          Send Feedback
        </h2>
        <button class="saai-modal-close-btn">×</button>
      </div>
      <div class="saai-feedback-modal-body">
        <!-- Feedback Form -->
        <form id="feedback-form" class="feedback-form">
          <div class="feedback-form-row">
            <div class="feedback-form-group">
              <label for="feedback-name">Your Name *</label>
              <input type="text" id="feedback-name" class="feedback-input" placeholder="John Doe" required />
            </div>
            
            <div class="feedback-form-group">
              <label for="feedback-designation">Designation *</label>
              <input type="text" id="feedback-designation" class="feedback-input" placeholder="Product Manager" required />
            </div>
          </div>
          
          <div class="feedback-form-group">
            <label for="feedback-email">Email Address *</label>
            <input type="email" id="feedback-email" class="feedback-input" placeholder="john@example.com" required />
          </div>
          
          <div class="feedback-form-group">
            <label for="feedback-experience">How was your experience? *</label>
            <textarea id="feedback-experience" class="feedback-textarea" rows="4" placeholder="Tell us about your experience with Sa.AI..." required></textarea>
          </div>
          
          <div class="feedback-form-group">
            <label for="feedback-improvement">What would make us improve? *</label>
            <textarea id="feedback-improvement" class="feedback-textarea" rows="4" placeholder="What features or improvements would you like to see?" required></textarea>
          </div>
          
          <div class="feedback-form-group">
            <label for="feedback-picture">Your Picture (Optional)</label>
            <p class="feedback-hint" style="margin: 4px 0 8px 0; color: #666;">Please upload your picture</p>
            <input type="file" id="feedback-picture" class="feedback-file-input" accept="image/*" />
            <small class="feedback-hint">JPG, PNG, GIF (Images will be compressed automatically)</small>
          </div>
          
          <div class="feedback-form-group" id="consent-group" style="display: none;">
            <label style="display: flex; align-items: flex-start; gap: 8px; cursor: pointer; user-select: none;">
              <input type="checkbox" id="feedback-consent" style="margin-top: 2px; cursor: pointer;" />
              <span style="font-size: 13px; color: #555; line-height: 1.4;">
                I consent to Sa.AI using my picture on their website and marketing materials *
              </span>
            </label>
          </div>
          
          <div class="feedback-form-actions">
            <button type="submit" class="feedback-submit-btn">
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <line x1="22" y1="2" x2="11" y2="13"/>
                <polygon points="22 2 15 22 11 13 2 9 22 2"/>
              </svg>
              Submit Feedback
            </button>
          </div>
        </form>
      </div>
    </div>
  `;
  
  document.body.appendChild(modal);
  addFeedbackModalListeners(modal, 'feedback');
}

function showSupportModal() {
  // Remove existing modal if any
  const existingModal = document.querySelector('.saai-feedback-modal');
  if (existingModal) existingModal.remove();
  
  const modal = document.createElement('div');
  modal.className = 'saai-modal saai-feedback-modal';
  modal.innerHTML = `
    <div class="saai-modal-content saai-feedback-modal-content">
      <div class="saai-modal-header">
        <h2 class="saai-modal-title">
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <circle cx="12" cy="12" r="10"/>
            <line x1="12" y1="8" x2="12" y2="12"/>
            <line x1="12" y1="16" x2="12.01" y2="16"/>
          </svg>
          Report an Issue
        </h2>
        <button class="saai-modal-close-btn">×</button>
      </div>
      <div class="saai-feedback-modal-body">
        <!-- Issue Report Form -->
        <form id="issue-form" class="feedback-form">
          <div class="feedback-form-row">
            <div class="feedback-form-group">
              <label for="issue-name">Your Name *</label>
              <input type="text" id="issue-name" class="feedback-input" placeholder="John Doe" required />
            </div>
            
            <div class="feedback-form-group">
              <label for="issue-email">Email Address *</label>
              <input type="email" id="issue-email" class="feedback-input" placeholder="john@example.com" required />
            </div>
          </div>
          
          <div class="feedback-form-group">
            <label for="issue-description">Tell us about the issue *</label>
            <textarea id="issue-description" class="feedback-textarea" rows="6" placeholder="Please describe the issue you're experiencing in detail..." required></textarea>
          </div>
          
          <div class="feedback-form-group">
            <label for="issue-screenshot">Supporting Screenshots (Optional)</label>
            <input type="file" id="issue-screenshot" class="feedback-file-input" accept="image/*" multiple />
            <small class="feedback-hint">Upload multiple screenshots (Max 5MB each)</small>
          </div>
          
          <div class="feedback-form-actions">
            <button type="submit" class="feedback-submit-btn">
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <line x1="22" y1="2" x2="11" y2="13"/>
                <polygon points="22 2 15 22 11 13 2 9 22 2"/>
              </svg>
              Submit Issue
            </button>
          </div>
        </form>
      </div>
    </div>
  `;
  
  document.body.appendChild(modal);
  addFeedbackModalListeners(modal, 'issue');
}

function addFeedbackModalListeners(modal, formType) {
  const closeBtn = modal.querySelector('.saai-modal-close-btn');
  const feedbackForm = modal.querySelector('#feedback-form');
  const issueForm = modal.querySelector('#issue-form');
  
  // Close button
  if (closeBtn) {
    closeBtn.addEventListener('click', () => closeFeedbackModal(modal));
  }
  
  // Click outside to close
  modal.addEventListener('click', (e) => {
    if (e.target === modal) {
      closeFeedbackModal(modal);
    }
  });
  
  // Form submission for feedback
  if (feedbackForm && formType === 'feedback') {
    const pictureInput = document.getElementById('feedback-picture');
    const consentGroup = document.getElementById('consent-group');
    const consentCheckbox = document.getElementById('feedback-consent');
    
    // Show/hide consent checkbox based on file selection
    if (pictureInput && consentGroup && consentCheckbox) {
      pictureInput.addEventListener('change', () => {
        if (pictureInput.files && pictureInput.files[0]) {
          // Image selected - show consent and make it required
          consentGroup.style.display = 'block';
          consentCheckbox.required = true;
        } else {
          // No image - hide consent and make it optional
          consentGroup.style.display = 'none';
          consentCheckbox.required = false;
          consentCheckbox.checked = false;
        }
      });
    }
    
    feedbackForm.addEventListener('submit', async (e) => {
      e.preventDefault();
      
      // Additional validation: if image is selected, consent must be checked
      if (pictureInput && pictureInput.files && pictureInput.files[0]) {
        if (!consentCheckbox.checked) {
          alert('Please check the consent checkbox to allow us to use your picture.');
          return;
        }
      }
      
      await submitFeedback(modal, feedbackForm, 'feedback');
    });
  }
  
  // Form submission for issue
  if (issueForm && formType === 'issue') {
    issueForm.addEventListener('submit', async (e) => {
      e.preventDefault();
      await submitFeedback(modal, issueForm, 'issue');
    });
  }
}

// Compress and convert image to Base64
async function compressImageToBase64(file, maxWidth = 800, quality = 0.8) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    
    reader.onload = (e) => {
      const img = new Image();
      
      img.onload = () => {
        // Create canvas
        const canvas = document.createElement('canvas');
        let width = img.width;
        let height = img.height;
        
        // Calculate new dimensions
        if (width > maxWidth) {
          height = (height * maxWidth) / width;
          width = maxWidth;
        }
        
        canvas.width = width;
        canvas.height = height;
        
        // Draw and compress
        const ctx = canvas.getContext('2d');
        ctx.drawImage(img, 0, 0, width, height);
        
        // Convert to Base64 with compression
        const compressedBase64 = canvas.toDataURL('image/jpeg', quality);
        
        // PRODUCTION: console.log('[Image Compression]', {
        //   originalSize: (file.size / 1024).toFixed(2) + ' KB',
        //   compressedSize: (compressedBase64.length * 0.75 / 1024).toFixed(2) + ' KB',
        //   compressionRatio: ((1 - (compressedBase64.length * 0.75 / file.size)) * 100).toFixed(1) + '%'
        // });
        
        resolve(compressedBase64);
      };
      
      img.onerror = () => reject(new Error('Failed to load image'));
      img.src = e.target.result;
    };
    
    reader.onerror = () => reject(new Error('Failed to read file'));
    reader.readAsDataURL(file);
  });
}

async function submitFeedback(modal, form, formType) {
  const submitBtn = form.querySelector('.feedback-submit-btn');
  const originalBtnText = submitBtn.innerHTML;
  
  try {
    // Get user ID
    const userId = await getUserId();
    
    let payload = {
      userId: userId,
      formType: formType,
      timestamp: new Date().toISOString(),
      userAgent: navigator.userAgent
    };
    
    // Get form values based on form type
    if (formType === 'feedback') {
      const experience = document.getElementById('feedback-experience').value;
      const improvement = document.getElementById('feedback-improvement').value;
      const name = document.getElementById('feedback-name').value;
      const designation = document.getElementById('feedback-designation').value;
      const email = document.getElementById('feedback-email').value;
      const pictureInput = document.getElementById('feedback-picture');
      const consentCheckbox = document.getElementById('feedback-consent');
      
      payload = {
        ...payload,
        experience: experience,
        improvement: improvement,
        name: name,
        designation: designation,
        email: email,
        pictureConsent: consentCheckbox.checked
      };
      
      // Handle picture upload with compression
      if (pictureInput.files && pictureInput.files[0]) {
        const file = pictureInput.files[0];
        
        // PRODUCTION: console.log('[Feedback] Processing image upload:', {
        //   fileName: file.name,
        //   fileSize: file.size,
        //   fileType: file.type
        // });
        
        submitBtn.innerHTML = 'Compressing image...';
        
        try {
          // Compress and convert to Base64
          const compressedBase64 = await compressImageToBase64(file, 800, 0.8);
          
          payload.picture = compressedBase64;
          payload.pictureName = file.name;
          payload.pictureOriginalSize = file.size;
          payload.pictureCompressedSize = Math.round(compressedBase64.length * 0.75);
          
          // PRODUCTION: console.log('[Feedback] Image compressed successfully', {
          //   pictureName: payload.pictureName,
          //   originalSize: payload.pictureOriginalSize,
          //   compressedSize: payload.pictureCompressedSize,
          //   base64Length: compressedBase64.length,
          //   base64Preview: compressedBase64.substring(0, 100) + '...'
          // });
        } catch (error) {
          // PRODUCTION: console.error('[Feedback] Image compression failed:', error);
          throw new Error('Failed to process image. Please try a different image.');
        }
      } else {
        // PRODUCTION: console.log('[Feedback] No image uploaded');
      }
      
    } else if (formType === 'issue') {
      const description = document.getElementById('issue-description').value;
      const name = document.getElementById('issue-name').value;
      const email = document.getElementById('issue-email').value;
      const screenshotInput = document.getElementById('issue-screenshot');
      
      payload = {
        ...payload,
        description: description,
        name: name,
        email: email
      };
      
      // Handle screenshot uploads with compression
      if (screenshotInput.files && screenshotInput.files.length > 0) {
        // PRODUCTION: console.log('[Issue] Compressing screenshots:', screenshotInput.files.length);
        submitBtn.innerHTML = `Compressing ${screenshotInput.files.length} image(s)...`;
        
        try {
          const compressedScreenshots = [];
          
          for (let i = 0; i < screenshotInput.files.length; i++) {
            const file = screenshotInput.files[i];
            // PRODUCTION: console.log(`[Issue] Compressing screenshot ${i + 1}/${screenshotInput.files.length}:`, file.name);
            
            const compressedBase64 = await compressImageToBase64(file, 1200, 0.85);
            
            compressedScreenshots.push({
              fileName: file.name,
              data: compressedBase64,
              originalSize: file.size,
              compressedSize: Math.round(compressedBase64.length * 0.75)
            });
          }
          
          payload.screenshots = compressedScreenshots;
          payload.screenshotCount = compressedScreenshots.length;
          
          // PRODUCTION: console.log('[Issue] All screenshots compressed successfully');
        } catch (error) {
          // PRODUCTION: console.error('[Issue] Screenshot compression failed:', error);
          throw new Error('Failed to process screenshots. Please try different images.');
        }
      }
    }
    
    // Disable submit button
    submitBtn.disabled = true;
    submitBtn.innerHTML = `
      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="animation: spin 1s linear infinite;">
        <circle cx="12" cy="12" r="10"/>
      </svg>
      Sending...
    `;
    
    // Send to n8n webhook via background script
    // PRODUCTION: console.log('[Feedback] Sending to background script:', { 
    //   endpoint: 'feedback', 
    //   payload: payload,
    //   hasPicture: !!payload.picture,
    //   pictureSize: payload.picture ? payload.picture.length : 0
    // });
    
    const response = await chrome.runtime.sendMessage({
      action: 'sendToN8N',
      data: {
        endpoint: 'feedback',
        payload: payload
      }
    });
    
    // PRODUCTION: console.log('[Feedback] Response from background:', response);
    
    if (response && response.success) {
      // Show success message
      submitBtn.innerHTML = `
        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
          <polyline points="20 6 9 17 4 12"/>
        </svg>
        Sent!
      `;
      submitBtn.style.background = '#16a34a';
      
      // Close modal after 1.5 seconds
      setTimeout(() => {
        closeFeedbackModal(modal);
      }, 1500);
    } else {
      throw new Error(response?.error || 'Failed to send feedback');
    }
    
  } catch (error) {
    // PRODUCTION: console.error('[Feedback] Error:', error);
    
    // Show error
    submitBtn.innerHTML = `
      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
        <circle cx="12" cy="12" r="10"/>
        <line x1="12" y1="8" x2="12" y2="12"/>
        <line x1="12" y1="16" x2="12.01" y2="16"/>
      </svg>
      Failed - Try Again
    `;
    submitBtn.style.background = '#dc2626';
    
    // Reset after 2 seconds
    setTimeout(() => {
      submitBtn.disabled = false;
      submitBtn.innerHTML = originalBtnText;
      submitBtn.style.background = '';
    }, 2000);
  }
}

function closeFeedbackModal(modal) {
  if (modal && modal.parentNode) {
    modal.remove();
  }
}

function getTasks() {
  const tasks = localStorage.getItem('saai-tasks');
  return tasks ? JSON.parse(tasks) : [];
}

function saveTasks(tasks) {
  localStorage.setItem('saai-tasks', JSON.stringify(tasks));
}

function addTask(taskText, priority = 'medium', dueDate = null) {
  const tasks = getTasks();
  const newTask = {
    id: Date.now().toString(),
    text: taskText,
    priority: priority,
    dueDate: dueDate,
    completed: false,
    createdAt: new Date().toISOString()
  };
  tasks.push(newTask);
  saveTasks(tasks);
}

function removeTask(taskId) {
  const tasks = getTasks();
  const filteredTasks = tasks.filter(task => task.id !== taskId);
  saveTasks(filteredTasks);
}

function toggleTaskComplete(taskId) {
  const tasks = getTasks();
  const task = tasks.find(t => t.id === taskId);
  if (task) {
    task.completed = !task.completed;
    saveTasks(tasks);
  }
}

// Add manual task function
async function addManualTask(taskText, priority = 'medium', fromValue = '') {
  try {
    // Get user ID from storage
    const { userId } = await chrome.storage.local.get(['userId']);
    
    if (!userId) {
      throw new Error('User ID not found. Please connect your Outlook account first.');
    }
    
    // Build webhook payload
    const payload = {
      userId: userId,
      context: 'ManualTaskAddition',
      taskName: taskText,
      priority: priority
    };
    
    // Add 'from' field if provided
    if (fromValue) {
      payload.from = fromValue;
    }
    
    // Send manual task to webhook for Supabase storage
    const response = await safeRequestWithRetry('https://connector.saai.dev/webhook/task-management-outlook', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${await ensureValidJWTToken()}`
      },
      body: JSON.stringify(payload)
    });
    
    if (!response.ok) {
      throw new Error(`Webhook request failed: ${response.status} ${response.statusText}`);
    }
    
    debugLog('Manual task sent to webhook successfully');
    
    // Also save locally for immediate display
    const tasks = getTasks();
    const newTask = {
      id: `manual-${Date.now()}`,
      task: taskText,
      priority: priority,
      completed: false,
      createdAt: new Date().toISOString(),
      isManual: true, // Flag to identify manually added tasks
      from: fromValue || undefined // Store 'from' if provided
    };
    tasks.push(newTask);
    saveTasks(tasks);
    
    debugLog('Manual task added locally:', newTask);
    
  } catch (error) {
    debugError('Failed to add manual task:', error);
    
    // Still save locally even if webhook fails
    const tasks = getTasks();
    const newTask = {
      id: `manual-${Date.now()}`,
      task: taskText,
      priority: priority,
      completed: false,
      createdAt: new Date().toISOString(),
      isManual: true,
      webhookFailed: true // Flag to indicate webhook sync failed
    };
    tasks.push(newTask);
    saveTasks(tasks);
    
    debugLog('Manual task saved locally (webhook failed):', newTask);
    
    // Could show user notification about sync failure
    // alert('Task added locally but failed to sync with backend. Please try again later.');
  }
}

// === Cleanup ===

function cleanup() {
  if (isSidebarOpen) {
    closeSidebar();
  }
  document.body.classList.remove('saai-sidebar-open');
}

// === Page Lifecycle ===

// === SETTINGS SIDEBAR FUNCTIONS ===

// Create settings sidebar
function createSettingsSidebar() {
  // Remove any existing triggers/sidebars first to prevent duplicates
  const existingSidebar = document.getElementById('saai-settings-sidebar');
  const existingTrigger = document.getElementById('saai-settings-trigger');
  
  if (existingSidebar) {
    existingSidebar.remove();
    debugLog('Removed existing settings sidebar');
  }
  
  if (existingTrigger) {
    existingTrigger.remove();
    debugLog('Removed existing settings trigger');
  }

  debugLog('Creating settings sidebar...');
  
  // Create the trigger separately (always visible)
  const trigger = document.createElement('div');
  trigger.id = 'saai-settings-trigger';
  trigger.className = 'saai-settings-trigger';
  
  // Add settings icon only
  const settingsIcon = document.createElement('div');
  settingsIcon.className = 'trigger-icon settings-icon';
  settingsIcon.onclick = function(e) {
    e.stopPropagation();
    toggleSettingsSidebar();
  };
  
  trigger.appendChild(settingsIcon);
  
  // Create the sidebar (hidden by default)
  const sidebar = document.createElement('div');
  sidebar.id = 'saai-settings-sidebar';
  sidebar.className = 'saai-settings-sidebar';
  
  sidebar.innerHTML = `
    <div class="saai-settings-header">
      <h3 class="saai-settings-title">
        <span>⚙</span>
        Settings
      </h3>
      <button class="saai-settings-close" id="saai-close-btn">×</button>
    </div>
    
    <div class="saai-settings-content">
      <button class="saai-settings-btn" id="saai-usage-btn">
        <div class="saai-settings-btn-text">
          Current Usage
          <div class="saai-settings-btn-desc">View your API usage and limits</div>
          <div class="usage-progress">
            <div class="usage-bar">
              <div class="usage-fill" id="settings-usage-fill" style="width: 0%"></div>
            </div>
            <div class="usage-text" id="settings-usage-text">0 / 1000 requests used</div>
          </div>
        </div>
      </button>
      
      <button class="saai-settings-btn" id="saai-feedback-btn">
        <div class="saai-settings-btn-text">
          Send Feedback
          <div class="saai-settings-btn-desc">Help us improve Sa.AI</div>
        </div>
      </button>
      
      <button class="saai-settings-btn" id="saai-support-btn">
        <div class="saai-settings-btn-text">
          Report Issue
          <div class="saai-settings-btn-desc">Get help with technical problems</div>
        </div>
      </button>
      
      <button class="saai-settings-btn danger" id="saai-clear-data-btn">
        <div class="saai-settings-btn-text">
          Clear All Data
          <div class="saai-settings-btn-desc" style="color: rgba(0, 0, 0, 0.8) !important;">Delete all stored information about me</div>
        </div>
      </button>
    </div>
  `;
  
  // Add both elements to body
  document.body.appendChild(trigger);
  document.body.appendChild(sidebar);
  
  // Add close button event listener
  const closeBtn = document.getElementById('saai-close-btn');
  if (closeBtn) {
    closeBtn.addEventListener('click', function(e) {
      e.stopPropagation();
      toggleSettingsSidebar();
      debugLog('Settings sidebar closed via close button');
    });
  }
  
  // Add usage button event listeners (CSP-compliant)
  const usageBtn = document.getElementById('saai-usage-btn');
  if (usageBtn) {
    usageBtn.addEventListener('click', function() {
      showUsageModal();
    });
    
    usageBtn.addEventListener('mouseenter', function() {
      showCreditTooltip(this);
    });
    
    usageBtn.addEventListener('mouseleave', function() {
      hideCreditTooltip();
    });
  }
  
  // Add feedback button event listener
  const feedbackBtn = document.getElementById('saai-feedback-btn');
  if (feedbackBtn) {
    feedbackBtn.addEventListener('click', function() {
      showFeedbackModal();
    });
  }
  
  // Add support button event listener
  const supportBtn = document.getElementById('saai-support-btn');
  if (supportBtn) {
    supportBtn.addEventListener('click', function() {
      showSupportModal();
    });
  }
  
  // Add clear data button event listener
  const clearDataBtn = document.getElementById('saai-clear-data-btn');
  if (clearDataBtn) {
    clearDataBtn.addEventListener('click', function() {
      confirmClearData();
    });
  }
  
  // Add background click to close functionality
  document.addEventListener('click', function(e) {
    const sidebar = document.getElementById('saai-settings-sidebar');
    const trigger = document.getElementById('saai-settings-trigger');
    
    if (sidebar && sidebar.classList.contains('open')) {
      // Check if click is outside both sidebar and trigger
      if (!sidebar.contains(e.target) && !trigger.contains(e.target)) {
        sidebar.classList.remove('open');
        debugLog('Settings sidebar closed by background click');
      }
    }
  });
  
  debugLog('Settings trigger and sidebar created and appended to body');
  
  // Verify they were added
  const addedTrigger = document.getElementById('saai-settings-trigger');
  const addedSidebar = document.getElementById('saai-settings-sidebar');
  if (addedTrigger && addedSidebar) {
    debugLog('Settings sidebar and trigger successfully added to DOM');
  } else {
    debugError('Settings sidebar or trigger failed to add to DOM');
  }
}

// Toggle settings sidebar
window.toggleSettingsSidebar = function() {
  debugLog('Toggle settings sidebar called');
  const sidebar = document.getElementById('saai-settings-sidebar');
  if (sidebar) {
    sidebar.classList.toggle('open');
    const isOpen = sidebar.classList.contains('open');
    debugLog('Settings sidebar toggled, now open:', isOpen);
  } else {
    debugError('Settings sidebar not found when trying to toggle');
  }
};

// Test tooltip function
window.testTooltip = function() {
  // PRODUCTION: console.log('Testing tooltip...');
  const button = document.querySelector('.saai-settings-btn');
  if (button) {
    // PRODUCTION: console.log('Found button:', button);
    showCreditTooltip(button);
  } else {
    // PRODUCTION: console.log('Button not found');
  }
};
window.showCreditTooltip = function(button) {
  // PRODUCTION: console.log('showCreditTooltip called - button:', button);
  debugLog('showCreditTooltip called');
  
  // Remove any existing tooltip
  hideCreditTooltip();
  
  const tooltip = document.createElement('div');
  tooltip.id = 'saai-credit-tooltip';
  tooltip.className = 'saai-credit-tooltip-inline';
  
  // PRODUCTION: console.log('Tooltip element created:', tooltip);
  
  tooltip.innerHTML = `
    <div class="credit-tooltip-header">
      <h4>Credit Usage Breakdown</h4>
    </div>
    <div class="credit-tooltip-content">
      <div class="credit-tooltip-item">
        <span class="credit-tooltip-feature">Thread Summarization</span>
        <span class="credit-tooltip-cost">5 credits</span>
      </div>
      <div class="credit-tooltip-item">
        <span class="credit-tooltip-feature">Web Search</span>
        <span class="credit-tooltip-cost">8 credits</span>
      </div>
      <div class="credit-tooltip-item">
        <span class="credit-tooltip-feature">Task Extraction</span>
        <span class="credit-tooltip-cost">10 credits</span>
      </div>
      <div class="credit-tooltip-item">
        <span class="credit-tooltip-feature">Ask Question (Inbox)</span>
        <span class="credit-tooltip-cost">10 credits</span>
      </div>
      <div class="credit-tooltip-item">
        <span class="credit-tooltip-feature">Inbox Summarization</span>
        <span class="credit-tooltip-cost">15 credits</span>
      </div>
      <div class="credit-tooltip-item">
        <span class="credit-tooltip-feature">Voice Mode</span>
        <span class="credit-tooltip-cost">varies by feature</span>
      </div>
      <div class="credit-tooltip-item credit-tooltip-danger">
        <span class="credit-tooltip-feature">Delete All Data</span>
        <span class="credit-tooltip-cost">250 credits</span>
      </div>
    </div>
  `;
  
  // Find the Settings sidebar container
  const sidebar = document.getElementById('saai-settings-sidebar');
  if (!sidebar) {
    // PRODUCTION: console.log('Settings sidebar not found; cannot show tooltip');
    return;
  }
  
  // Insert inline tooltip at the end of the current usage button block
  // Find the usage button and insert after it
  const usageBtn = button.closest('.saai-settings-btn');
  if (usageBtn && usageBtn.parentNode) {
    usageBtn.parentNode.insertBefore(tooltip, usageBtn.nextSibling);
    // PRODUCTION: console.log('Tooltip inserted after usage button');
  } else {
    sidebar.appendChild(tooltip);
    // PRODUCTION: console.log('Tooltip appended to sidebar');
  }
  
  debugLog('Tooltip created and inserted in sidebar');
  
  // Add fade-in animation
  setTimeout(() => {
    tooltip.classList.add('show');
    // PRODUCTION: console.log('Tooltip show class added, final element:', tooltip);
    debugLog('Tooltip show class added');
  }, 10);
};

// Hide credit breakdown tooltip
window.hideCreditTooltip = function() {
  debugLog('hideCreditTooltip called');
  const tooltip = document.getElementById('saai-credit-tooltip');
  if (tooltip) {
    tooltip.classList.remove('show');
    setTimeout(() => {
      if (tooltip.parentNode) {
        tooltip.parentNode.removeChild(tooltip);
        debugLog('Tooltip removed from DOM');
      }
    }, 200);
  }
};
window.showUsageModal = function() {
  const modal = document.createElement('div');
  modal.className = 'task-modal-overlay';
  modal.innerHTML = `
    <div class="task-modal-content">
      <div class="task-modal-header">
        <h3 class="task-modal-title">
          <span>📊</span>
          Current Usage
        </h3>
        <button class="task-modal-close-btn" onclick="this.closest('.task-modal-overlay').remove()">×</button>
      </div>
      <div class="task-modal-body">
        <div style="text-align: center; padding: 20px;">
          <div style="font-size: 48px; color: var(--saai-primary); font-weight: 700; margin-bottom: 16px;" id="modal-usage-number">0</div>
          <div style="font-size: 16px; color: var(--saai-text-secondary); margin-bottom: 24px;">requests used this month</div>
          
          <div class="usage-progress" style="max-width: 300px; margin: 0 auto 24px;">
            <div class="usage-bar" style="height: 12px;">
              <div class="usage-fill" id="modal-usage-fill" style="width: 0%"></div>
            </div>
            <div style="display: flex; justify-content: space-between; margin-top: 8px; font-size: 12px; color: var(--saai-text-secondary);">
              <span>0</span>
              <span>1000 limit</span>
            </div>
          </div>
          
          <div style="background: var(--saai-background-secondary); padding: 16px; border-radius: var(--saai-border-radius); border: 1px solid var(--saai-border);">
            <div style="font-size: 14px; color: var(--saai-text-primary); margin-bottom: 8px; font-weight: 500;">Usage Details</div>
            <div style="font-size: 12px; color: var(--saai-text-secondary); line-height: 1.5;">
              • Email summaries: 420 requests<br>
              • Chat conversations: 180 requests<br>
              • Task extractions: 50 requests<br>
              • Resets on: ${new Date(Date.now() + 30 * 24 * 60 * 60 * 1000).toLocaleDateString()}
            </div>
          </div>
        </div>
        <button class="task-modal-ok-btn" onclick="this.closest('.task-modal-overlay').remove()">Close</button>
      </div>
    </div>
  `;
  
  document.body.appendChild(modal);
  
  // Close on backdrop click
  modal.addEventListener('click', (e) => {
    if (e.target === modal) {
      modal.remove();
    }
  });
};

// Show feedback modal
window.showFeedbackModal = function() {
  const modal = document.createElement('div');
  modal.className = 'task-modal-overlay';
  modal.innerHTML = `
    <div class="task-modal-content">
      <div class="task-modal-header">
        <h3 class="task-modal-title">
          <span>💬</span>
          Send Feedback
        </h3>
        <button class="task-modal-close-btn" onclick="this.closest('.task-modal-overlay').remove()">×</button>
      </div>
      <div class="task-modal-body">
        <div style="margin-bottom: 20px;">
          <label style="display: block; font-weight: 500; margin-bottom: 8px; color: var(--saai-text-primary);">Feedback Type</label>
          <select id="feedback-type" style="width: 100%; padding: 8px 12px; border: 1px solid var(--saai-border); border-radius: var(--saai-border-radius); font-family: inherit;">
            <option>Feature Request</option>
            <option>Bug Report</option>
            <option>General Feedback</option>
            <option>Improvement Suggestion</option>
          </select>
        </div>
        
        <div style="margin-bottom: 20px;">
          <label style="display: block; font-weight: 500; margin-bottom: 8px; color: var(--saai-text-primary);">Your Feedback</label>
          <textarea id="feedback-text" placeholder="Tell us what you think about Sa.AI..." style="width: 100%; height: 120px; padding: 12px; border: 1px solid var(--saai-border); border-radius: var(--saai-border-radius); font-family: inherit; resize: vertical;"></textarea>
        </div>
        
        <div style="margin-bottom: 20px;">
          <label style="display: block; font-weight: 500; margin-bottom: 8px; color: var(--saai-text-primary);">Email (optional)</label>
          <input type="email" id="feedback-email" placeholder="your.email@example.com" style="width: 100%; padding: 8px 12px; border: 1px solid var(--saai-border); border-radius: var(--saai-border-radius); font-family: inherit;">
        </div>
        
        <button class="task-modal-ok-btn" onclick="submitFeedback()">Send Feedback</button>
      </div>
    </div>
  `;
  
  document.body.appendChild(modal);
  
  // Close on backdrop click
  modal.addEventListener('click', (e) => {
    if (e.target === modal) {
      modal.remove();
    }
  });
};

// Show support modal
window.showSupportModal = function() {
  const modal = document.createElement('div');
  modal.className = 'task-modal-overlay';
  modal.innerHTML = `
    <div class="task-modal-content">
      <div class="task-modal-header">
        <h3 class="task-modal-title">
          <span>🆘</span>
          Report Issue
        </h3>
        <button class="task-modal-close-btn" onclick="this.closest('.task-modal-overlay').remove()">×</button>
      </div>
      <div class="task-modal-body">
        <div style="margin-bottom: 20px;">
          <label style="display: block; font-weight: 500; margin-bottom: 8px; color: var(--saai-text-primary);">Issue Type</label>
          <select id="issue-type" style="width: 100%; padding: 8px 12px; border: 1px solid var(--saai-border); border-radius: var(--saai-border-radius); font-family: inherit;">
            <option>Extension not working</option>
            <option>Outlook integration issues</option>
            <option>Authentication problems</option>
            <option>Performance issues</option>
            <option>Other technical issue</option>
          </select>
        </div>
        
        <div style="margin-bottom: 20px;">
          <label style="display: block; font-weight: 500; margin-bottom: 8px; color: var(--saai-text-primary);">Describe the Issue</label>
          <textarea id="issue-description" placeholder="Please describe what happened and what you expected to happen..." style="width: 100%; height: 120px; padding: 12px; border: 1px solid var(--saai-border); border-radius: var(--saai-border-radius); font-family: inherit; resize: vertical;"></textarea>
        </div>
        
        <div style="margin-bottom: 20px;">
          <label style="display: block; font-weight: 500; margin-bottom: 8px; color: var(--saai-text-primary);">Contact Email</label>
          <input type="email" id="support-email" placeholder="your.email@example.com" style="width: 100%; padding: 8px 12px; border: 1px solid var(--saai-border); border-radius: var(--saai-border-radius); font-family: inherit;">
        </div>
        
        <div style="background: var(--saai-background-secondary); padding: 12px; border-radius: var(--saai-border-radius); margin-bottom: 20px; font-size: 12px; color: var(--saai-text-secondary);">
          <strong>Debug Info:</strong><br>
          Browser: ${navigator.userAgent.split(' ').slice(-2).join(' ')}<br>
          Extension Version: 2.2.3<br>
          Timestamp: ${new Date().toISOString()}
        </div>
        
        <button class="task-modal-ok-btn" onclick="submitSupportRequest()">Submit Report</button>
      </div>
    </div>
  `;
  
  document.body.appendChild(modal);
  
  // Close on backdrop click
  modal.addEventListener('click', (e) => {
    if (e.target === modal) {
      modal.remove();
    }
  });
};

// Confirm clear data
window.confirmClearData = function() {
  const modal = document.createElement('div');
  modal.className = 'saai-modal saai-confirmation-modal';
  modal.innerHTML = `
    <div class="saai-modal-content saai-confirmation-content">
      <div class="saai-modal-header">
        <h2 class="saai-modal-title">
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/>
            <line x1="12" y1="9" x2="12" y2="13"/>
            <line x1="12" y1="17" x2="12.01" y2="17"/>
          </svg>
          Warning!
        </h2>
        <button class="saai-modal-close-btn" id="clear-data-close-btn">×</button>
      </div>
      <div class="saai-modal-body saai-confirmation-body">
        <div class="confirmation-icon">⚠️</div>
        <div class="confirmation-message">
          <p class="confirmation-title">Sa.ai will lose your data</p>
          <p class="confirmation-subtitle">This action cannot be undone. All your data will be permanently deleted.</p>
          </div>
        <div class="confirmation-actions">
          <button class="confirmation-btn confirmation-btn-no" id="clear-data-no-btn">
            No
          </button>
          <button class="confirmation-btn confirmation-btn-yes" id="clear-data-yes-btn">
            Yes, Clear Data
          </button>
        </div>
      </div>
    </div>
  `;
  
  document.body.appendChild(modal);
  
  // Attach event listeners to all buttons
  const closeBtn = modal.querySelector('#clear-data-close-btn');
  const noBtn = modal.querySelector('#clear-data-no-btn');
  const yesBtn = modal.querySelector('#clear-data-yes-btn');
  
  if (closeBtn) {
    closeBtn.addEventListener('click', () => {
      // PRODUCTION: console.log('[SaAI] Close button clicked');
      modal.remove();
    });
  }
  
  if (noBtn) {
    noBtn.addEventListener('click', () => {
      // PRODUCTION: console.log('[SaAI] No button clicked');
      modal.remove();
    });
  }
  
  if (yesBtn) {
    yesBtn.addEventListener('click', async () => {
      // PRODUCTION: console.log('[SaAI] Yes button clicked');
      try {
        await window.clearAllData();
      } catch (error) {
        // PRODUCTION: console.error('[SaAI] Error in clearAllData:', error);
        alert('Failed to clear data: ' + error.message);
      }
    });
  }
  
  // Close on backdrop click
  modal.addEventListener('click', (e) => {
    if (e.target === modal) {
      // PRODUCTION: console.log('[SaAI] Backdrop clicked');
      modal.remove();
    }
  });
};

// Submit feedback
window.submitFeedback = function() {
  const type = document.getElementById('feedback-type').value;
  const text = document.getElementById('feedback-text').value;
  const email = document.getElementById('feedback-email').value;
  
  if (!text.trim()) {
    alert('Please enter your feedback before submitting.');
    return;
  }
  
  // Here you would typically send to your feedback endpoint
  // PRODUCTION: console.log('Feedback submitted - type:', type);
  
  // Show success message
  const modal = document.querySelector('.task-modal-overlay');
  modal.innerHTML = `
    <div class="task-modal-content">
      <div class="task-modal-header">
        <h3 class="task-modal-title">
          <span>✅</span>
          Feedback Sent
        </h3>
      </div>
      <div class="task-modal-body">
        <div style="text-align: center; padding: 20px;">
          <div style="font-size: 48px; margin-bottom: 16px;">🙏</div>
          <div style="font-size: 16px; color: var(--saai-text-primary); margin-bottom: 16px; font-weight: 500;">
            Thank you for your feedback!
          </div>
          <div style="font-size: 14px; color: var(--saai-text-secondary); margin-bottom: 24px;">
            We appreciate you taking the time to help us improve Sa.AI.
          </div>
        </div>
        <button class="task-modal-ok-btn" onclick="this.closest('.task-modal-overlay').remove()">Close</button>
      </div>
    </div>
  `;
};

// Submit support request
window.submitSupportRequest = function() {
  const type = document.getElementById('issue-type').value;
  const description = document.getElementById('issue-description').value;
  const email = document.getElementById('support-email').value;
  
  if (!description.trim() || !email.trim()) {
    alert('Please fill in all required fields.');
    return;
  }
  
  // Here you would typically send to your support endpoint
  // PRODUCTION: console.log('Support request submitted - type:', type);
  
  // Show success message
  const modal = document.querySelector('.task-modal-overlay');
  modal.innerHTML = `
    <div class="task-modal-content">
      <div class="task-modal-header">
        <h3 class="task-modal-title">
          <span>✅</span>
          Report Submitted
        </h3>
      </div>
      <div class="task-modal-body">
        <div style="text-align: center; padding: 20px;">
          <div style="font-size: 48px; margin-bottom: 16px;">📧</div>
          <div style="font-size: 16px; color: var(--saai-text-primary); margin-bottom: 16px; font-weight: 500;">
            Support request submitted successfully!
          </div>
          <div style="font-size: 14px; color: var(--saai-text-secondary); margin-bottom: 24px;">
            We'll get back to you within 24 hours at the provided email address.
          </div>
        </div>
        <button class="task-modal-ok-btn" onclick="this.closest('.task-modal-overlay').remove()">Close</button>
      </div>
    </div>
  `;
};

// Clear all data
window.clearAllData = async function() {
  // PRODUCTION: console.log('[SaAI] clearAllData() function called');
  
  // Close the confirmation modal
  const confirmModal = document.querySelector('.saai-confirmation-modal');
  if (confirmModal) {
    confirmModal.remove();
  }
  
  try {
    // PRODUCTION: console.log('[SaAI] Getting userId from storage...');
    // Get user email before clearing
    const { userId } = await chrome.storage.local.get(['userId']);
    const userEmail = userId || 'unknown_user';
    // PRODUCTION: console.log('[SaAI] User email available:', !!userEmail);
    
    // Send notification to backend about data deletion using webhook pattern
    try {
      // PRODUCTION: console.log('[SaAI] Sending data deletion notification to webhook...');
      debugLog('Sending data deletion notification to webhook');
      
      const response = await chrome.runtime.sendMessage({
        action: 'sendToN8N',
        data: {
          endpoint: 'dataDeletion',
          payload: {
          email: userEmail,
            userId: userEmail,
          timestamp: new Date().toISOString(),
            action: 'data_cleared',
            userAgent: navigator.userAgent
          }
        }
      });
      
      // PRODUCTION: console.log('[SaAI] Data deletion notification response:', response);
      debugLog('Data deletion notification sent:', response);
    } catch (notifyError) {
      // PRODUCTION: console.error('[SaAI] Failed to send deletion notification:', notifyError);
      debugError('Failed to send deletion notification:', notifyError);
      // Continue with deletion even if notification fails
    }
    
  // Clear all storage
  chrome.storage.local.clear(() => {
    chrome.storage.sync.clear(() => {
        // Show final success message
        const modal = document.createElement('div');
        modal.className = 'saai-modal saai-confirmation-modal';
      modal.innerHTML = `
          <div class="saai-modal-content saai-confirmation-content">
            <div class="saai-modal-header">
              <h2 class="saai-modal-title">
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                  <polyline points="20 6 9 17 4 12"/>
                </svg>
                Deletion Request Submitted
              </h2>
          </div>
            <div class="saai-modal-body saai-confirmation-body">
              <div class="confirmation-icon success">✓</div>
              <div class="confirmation-message">
                <p class="confirmation-title">Your request has been received</p>
                <p class="confirmation-subtitle">
                  devang@saai.dev is notified. Once your data is fully erased you will get a confirmation email within 6 hours
                </p>
              </div>
              <div class="confirmation-actions">
                <button class="confirmation-btn confirmation-btn-primary" onclick="window.location.reload()">
                  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                    <polyline points="23 4 23 10 17 10"/>
                    <polyline points="1 20 1 14 7 14"/>
                    <path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/>
                  </svg>
                  Reload Page
                </button>
              </div>
          </div>
        </div>
      `;
        
        document.body.appendChild(modal);
    });
  });
  } catch (error) {
    debugError('Error in clearAllData:', error);
    alert('An error occurred while clearing data. Please try again.');
  }
};

// Function to add email to tasks via webhook
async function addEmailToTasks(emailTaskData) {
  try {
    // PRODUCTION: console.log('[SaAI] addEmailToTasks called - priority:', emailTaskData?.priority);
    debugLog('Adding email to tasks - priority:', emailTaskData?.priority);
    
    // Get userId and JWT token
    const userId = await getUserId();
    const jwtToken = await getJWTToken();
    
    // PRODUCTION: console.log('[SaAI] User credentials:', { userId, hasToken: !!jwtToken });
    
    if (!userId || !jwtToken) {
      throw new Error('User not authenticated');
    }
    
    // Build request body
    const requestBody = {
      userId: userId,
      from: emailTaskData.from,
      context: "ManualTaskAddition",
      taskName: emailTaskData.taskName,
      priority: emailTaskData.priority
    };
    
    // PRODUCTION: console.log('[SaAI] Request body:', requestBody);
    debugLog('Sending task to webhook:', requestBody);
    
    // Send to Task Management webhook
    const response = await chrome.runtime.sendMessage({
      action: 'sendToN8N',
      data: {
        endpoint: 'task',
        payload: requestBody
      }
    });
    
    debugLog('Task webhook response received');
    
    if (response && response.success) {
      // PRODUCTION: console.log('[SaAI] ✅ Task added successfully!');
      debugLog('✅ Task added successfully!');
      // Success - caller will handle UI update
      return true;
    } else {
      throw new Error(response?.error || 'Failed to add task');
    }
    
  } catch (error) {
    // PRODUCTION: console.error('[SaAI] Error adding email to tasks:', error);
    debugError('Error adding email to tasks:', error);
    alert('❌ Failed to add task: ' + error.message);
    throw error; // Re-throw so caller can handle it
  }
}

// Listen for messages (for future cross-window communication if needed)
window.addEventListener('message', (event) => {
  // Security: Validate origin to prevent malicious messages
  if (event.origin !== 'https://mail.google.com') {
    // PRODUCTION: console.warn('[SaAI Main] Message rejected - invalid origin:', event.origin);
    return;
  }
  
  // Security: Validate event source
  if (!event.source) {
    // PRODUCTION: console.warn('[SaAI Main] Message rejected - no source');
    return;
  }
  
  // PRODUCTION: console.log('[SaAI Main] Message received - action:', event.data?.action);
  
  if (event.data && event.data.action === 'openTaskModal') {
    showTaskModal();
  }
  // Note: addEmailToTask is now handled directly in the popup's click handler
  // since the handler runs in the parent's context
});

// Initialize when DOM is ready
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', initialize);
} else {
  initialize();
}

// Cleanup on page unload
window.addEventListener('beforeunload', cleanup);

// Handle Outlook navigation and sidebar persistence
let lastUrl = location.href;
let sidebarCheckInterval = null;

// MutationObserver for URL changes (Outlook navigation)
new MutationObserver(() => {
  const url = location.href;
  if (url !== lastUrl) {
    lastUrl = url;
    debugLog('Outlook navigation detected');
    // Re-initialize if needed
    if (!isInitialized) {
      initialize();
    } else {
      // Check if sidebar was removed during navigation
      checkAndRestoreSidebarState();
    }
  }
}).observe(document, { subtree: true, childList: true });

// Additional safeguard: Periodically check if sidebar still exists
// This handles cases where Outlook re-renders and removes our sidebar
if (!sidebarCheckInterval) {
  sidebarCheckInterval = setInterval(() => {
    if (isSidebarOpen && !document.getElementById(SIDEBAR_ID)) {
      debugLog('Sidebar missing but should be open - restoring...');
      openSidebar().catch(error => {
        debugError('Failed to restore sidebar:', error);
      });
    }
  }, 2000); // Check every 2 seconds
}

} // Close the if block for duplicate script prevention

})(); // End IIFE
