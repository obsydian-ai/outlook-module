// Office.js initialization
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    if (document.readyState === 'loading') {
      document.addEventListener('DOMContentLoaded', initializeApp);
    } else {
      initializeApp();
    }
  }
});

// Configuration
const API_ENDPOINT = 'https://api-obsydian.up.railway.app/api/claims/create-claim';
const API_AUTH_TOKEN = 'eyJhbGciOiJIUzI1NiIsImtpZCI6IjV3MGlwRldVbDhsNC9aNkUiLCJ0eXAiOiJKV1QifQ.eyJpc3MiOiJodHRwczovL21reGVya2RybXBna29yanF0eHR3LnN1cGFiYXNlLmNvL2F1dGgvdjEiLCJzdWIiOiJiMDFjYmFlYi1kYjM0LTQ3Y2UtOTRjNy1hYzIyYjQ5MjZmNGMiLCJhdWQiOiJhdXRoZW50aWNhdGVkIiwiZXhwIjoxNzYxOTI0NzQ5LCJpYXQiOjE3NjE5MjExNDksImVtYWlsIjoicGFibG9Ab2JzeWRpYW5haS5jb20iLCJwaG9uZSI6IiIsImFwcF9tZXRhZGF0YSI6eyJvcmdhbml6YXRpb25faWQiOiJkZW1vX29yZ19pZCIsInByb3ZpZGVyIjoiZW1haWwiLCJwcm92aWRlcnMiOlsiZW1haWwiXX0sInVzZXJfbWV0YWRhdGEiOnsiZW1haWxfdmVyaWZpZWQiOnRydWV9LCJyb2xlIjoiYXV0aGVudGljYXRlZCIsImFhbCI6ImFhbDEiLCJhbXIiOlt7Im1ldGhvZCI6InBhc3N3b3JkIiwidGltZXN0YW1wIjoxNzYxOTEzNDE2fV0sInNlc3Npb25faWQiOiIzMTBkMDM2OS1hY2M2LTQwZTEtYWMxNi02MGNlYmI5ZWYyN2UiLCJpc19hbm9ueW1vdXMiOmZhbHNlfQ.Uu5qiNKsmubdCCfyLmOTXdcmIcdOanvn7zWr9uFn_QI';
const API_ORGANIZATION_ID = 'demo_org_id';
const GEMINI_API_KEY = 'AIzaSyD9ZNemQeajKaH1gS30RYC6eReiJdHnQWg';
const GEMINI_API_URL = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent';

// Global state
let currentView = 'landing';

function initializeApp() {
  if (window.appInitialized) return;
  
  // Wait for DOM elements to be ready
  const checkElements = setInterval(() => {
    const elements = {
      landingView: document.getElementById("landingView"),
      formView: document.getElementById("formView"),
      successView: document.getElementById("successView"),
      loadingView: document.getElementById("loadingView"),
      createClaimBtn: document.getElementById("createClaimBtn"),
      backToLandingBtn: document.getElementById("backToLandingBtn"),
      claimForm: document.getElementById("claimForm"),
      carrierSelect: document.getElementById("carrierSelect"),
      cancelBtn: document.getElementById("cancelBtn"),
      newClaimBtn: document.getElementById("newClaimBtn")
    };
    
    const allReady = Object.values(elements).every(el => el !== null);
    
    if (allReady) {
      clearInterval(checkElements);
      setupEventListeners(elements);
      window.appInitialized = true;
    }
  }, 100);
  
  // Timeout after 5 seconds
  setTimeout(() => clearInterval(checkElements), 5000);
}

function setupEventListeners(elements) {
  const {
    createClaimBtn,
    backToLandingBtn,
    claimForm,
    carrierSelect,
    cancelBtn,
    newClaimBtn
  } = elements;

  // Create Claim button
  createClaimBtn.addEventListener("click", async () => {
    if (!Office?.context) {
      showView('form');
      return;
    }
    
    showView('loading');
    updateLoadingText("AI is analyzing email content...");
    
    try {
      await new Promise(resolve => setTimeout(resolve, 800));
      
      updateLoadingText("Fetching email content...");
      const emailData = await fetchEmailContent();
      window.emailData = emailData;
      
      await new Promise(resolve => setTimeout(resolve, 1200));
      
      updateLoadingText("AI is extracting claim information...");
      const extractedFormData = await summarizeWithGemini(emailData);
      window.extractedFormData = extractedFormData;
      
      await new Promise(resolve => setTimeout(resolve, 1000));
      
      updateLoadingText("Preparing your claim form...");
      await new Promise(resolve => setTimeout(resolve, 600));
      
      showView('form');
      prefillForm(extractedFormData);
    } catch (error) {
      console.error("Error in Create Claim process:", error);
      alert(`Warning: Could not complete the process (${error.message}). You can still create a claim manually.`);
      showView('form');
    }
  });

  // Navigation buttons
  backToLandingBtn.addEventListener("click", () => showView('landing'));
  cancelBtn.addEventListener("click", () => showView('landing'));
  newClaimBtn.addEventListener("click", () => {
    resetForm();
    showView('landing');
  });
  
  // View details button
  const viewDetailsBtn = document.getElementById('viewDetailsBtn');
  if (viewDetailsBtn) {
    viewDetailsBtn.addEventListener('click', () => {
      const claimIdElement = document.getElementById('successClaimId');
      const claimId = window.lastClaimId || claimIdElement?.textContent?.trim();
      
      if (claimId && claimId !== '-') {
        window.open(`https://dashboard.obsydianai.com/claims/${claimId}`, '_blank');
      } else {
        alert('Claim ID not available.');
      }
    });
  }

  // Carrier selection
  carrierSelect.addEventListener("change", (e) => {
    if (e.target.value === 'ups_001') {
      showUpsFields();
    } else {
      hideUpsFields();
    }
  });

  // Incidence type - show supporting evidence only for DAMAGED
  const incidenceSelect = document.getElementById('incidenceType');
  const supportingEvidenceGroup = document.getElementById('supportingEvidenceGroup');
  if (incidenceSelect && supportingEvidenceGroup) {
    function updateSupportingEvidenceVisibility() {
      supportingEvidenceGroup.style.display = incidenceSelect.value === 'DAMAGED' ? '' : 'none';
    }
    incidenceSelect.addEventListener('change', updateSupportingEvidenceVisibility);
    updateSupportingEvidenceVisibility();
  }

  // Shipment items
  const addItemBtn = document.getElementById('addItemBtn');
  const itemsList = document.getElementById('itemsList');
  if (itemsList && addItemBtn) {
    addItemBtn.addEventListener('click', () => addShipmentItemRow());
    if (itemsList.children.length === 0) {
      addShipmentItemRow();
    }
  }

  // File upload handlers
  const proofOfPurchaseInput = document.getElementById('proofOfPurchase');
  const supportingEvidenceInput = document.getElementById('supportingEvidence');
  
  if (proofOfPurchaseInput) {
    proofOfPurchaseInput.addEventListener('change', (e) => {
      handleFileUpload(e, 'proofOfPurchaseName');
    });
  }
  
  if (supportingEvidenceInput) {
    supportingEvidenceInput.addEventListener('change', (e) => {
      handleFileUpload(e, 'supportingEvidenceName');
    });
  }

  // Form submission
  if (claimForm) {
    claimForm.addEventListener("submit", async (e) => {
      e.preventDefault();
      
      if (claimForm._submitting) return;
      claimForm._submitting = true;
      
      try {
        const items = getShipmentItems();
        const totalAmount = items.reduce((sum, it) => sum + (Number(it.amount) || 0) * (Number(it.quantity) || 0), 0);
        
        const formattedItems = items.map((it, index) => ({
          id: `item-${index}`,
          description: it.description || '',
          quantity: Number(it.quantity) || 1,
          amount: Number(it.amount) || 0
        }));
        const contentsDescription = JSON.stringify(formattedItems);

        const formData = new FormData(claimForm);
        const claimData = Object.fromEntries(formData.entries());
        claimData.actualAmount = totalAmount.toFixed(2);
        claimData.contentsDescription = contentsDescription;
        
        if (!validateForm(claimData)) {
          return;
        }
        
        showView('loading');
        updateLoadingText("Submitting your claim...");
        
        const result = await submitClaim(claimData);
        populateSuccessView(claimData, result);
        showView('success');
        
        // Scroll to top
        requestAnimationFrame(() => {
          window.scrollTo({ top: 0, behavior: 'smooth' });
          document.documentElement.scrollTop = 0;
          document.body.scrollTop = 0;
        });
      } catch (error) {
        console.error("Error submitting claim:", error);
        showView('form');
        alert('Error submitting claim: ' + error.message);
      } finally {
        claimForm._submitting = false;
      }
    });
  }
}

function showView(viewName) {
  const views = ['landingView', 'formView', 'successView', 'loadingView'];
  views.forEach(viewId => {
    const view = document.getElementById(viewId);
    if (view) view.classList.add('hidden');
  });
  
  const targetView = document.getElementById(`${viewName}View`);
  if (targetView) {
    targetView.classList.remove('hidden');
    currentView = viewName;
  }
}

function updateLoadingText(text) {
  const loadingTextElement = document.getElementById('loadingText');
  if (loadingTextElement) {
    loadingTextElement.textContent = text;
  }
}

// File handling
function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const base64String = reader.result.split(',')[1];
      resolve(base64String);
    };
    reader.onerror = () => reject(new Error('Failed to read file'));
    reader.readAsDataURL(file);
  });
}

async function extractItemsFromInvoice(file) {
  try {
    const base64Data = await fileToBase64(file);
    
    const response = await fetch('https://api-obsydian.up.railway.app/api/invoices/extract-items', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        file: {
          base64Data: base64Data,
          mimeType: file.type,
          name: file.name
        }
      })
    });
    
    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }
    
    const result = await response.json();
    if (result.success && Array.isArray(result.items)) {
      return result.items;
    }
    throw new Error('Invalid response format from extraction API');
  } catch (error) {
    console.error('Error extracting items from invoice:', error);
    throw error;
  }
}

async function handleFileUpload(event, fileNameElementId) {
  const file = event.target.files[0];
  const fileNameElement = document.getElementById(fileNameElementId);
  const isProofOfPurchase = fileNameElementId === 'proofOfPurchaseName';
  
  if (isProofOfPurchase && event.target._extracting) return;
  
  if (file) {
    // Validate file type
    const allowedTypes = ['application/pdf', 'image/jpeg', 'image/jpg', 'image/png'];
    if (!allowedTypes.includes(file.type)) {
      alert('Please upload a PDF, JPG, or PNG file.');
      event.target.value = '';
      return;
    }
    
    // Validate file size (max 10MB)
    const maxSize = 10 * 1024 * 1024;
    if (file.size > maxSize) {
      alert('File size must be less than 10MB.');
      event.target.value = '';
      return;
    }
    
    if (fileNameElement) {
      fileNameElement.textContent = file.name;
      fileNameElement.classList.add('show');
    }
    
    // Extract items from invoice if it's proof of purchase
    if (isProofOfPurchase) {
      event.target._extracting = true;
      
      try {
        const fileUploadContainer = fileNameElement?.closest('.file-upload-container');
        let aiLoadingElement = null;
        
        if (fileUploadContainer) {
          fileUploadContainer.querySelectorAll('.ai-extracting, .extraction-success, .extraction-error').forEach(msg => msg.remove());
          
          aiLoadingElement = document.createElement('div');
          aiLoadingElement.className = 'ai-extracting';
          aiLoadingElement.innerHTML = `
            <div class="ai-extracting-spinner"></div>
            <div class="ai-extracting-text">AI is extracting items</div>
          `;
          fileUploadContainer.appendChild(aiLoadingElement);
        }
        
        const extractedItems = await extractItemsFromInvoice(file);
        
        if (aiLoadingElement) aiLoadingElement.remove();
        
        if (extractedItems && extractedItems.length > 0) {
          // Remove empty rows
          const itemsList = document.getElementById('itemsList');
          if (itemsList) {
            const rows = Array.from(itemsList.querySelectorAll('.item-row'));
            rows.forEach(row => {
              const descInput = row.querySelector('input[type="text"]');
              const amountInput = row.querySelector('.amount-only');
              const descValue = descInput ? descInput.value.trim() : '';
              const amountValue = amountInput ? amountInput.value.trim() : '';
              
              if (!descValue && (!amountValue || amountValue === '0.00' || amountValue === '0')) {
                row.remove();
              }
            });
          }
          
          extractedItems.forEach((item) => {
            addShipmentItemRow({
              description: item.description || '',
              quantity: item.quantity?.toString() || '1',
              amount: item.amount ? parseFloat(item.amount).toFixed(2) : '',
              currency: item.currency || 'EUR'
            });
          });
          
          recalcItemsTotals();
          
          // Show success message
          if (fileUploadContainer) {
            const successMessage = document.createElement('div');
            successMessage.className = 'extraction-success';
            successMessage.style.cssText = 'margin-top: 8px; padding: 8px 12px; background-color: #d1fae5; border-radius: 4px; font-size: 13px; color: #065f46;';
            successMessage.textContent = `✓ AI extracted ${extractedItems.length} item(s) from ${file.name}`;
            fileUploadContainer.appendChild(successMessage);
          }
        }
      } catch (error) {
        console.error('Failed to extract items from invoice:', error);
        
        const fileUploadContainer = fileNameElement?.closest('.file-upload-container');
        if (fileUploadContainer) {
          fileUploadContainer.querySelector('.ai-extracting')?.remove();
        }
        
        const errorMessage = document.createElement('div');
        errorMessage.className = 'extraction-error';
        errorMessage.style.cssText = 'margin-top: 8px; padding: 8px 12px; background-color: #fee2e2; border-radius: 4px; font-size: 13px; color: #991b1b;';
        errorMessage.textContent = '⚠ Could not extract items automatically. You can add them manually below.';
        fileUploadContainer?.appendChild(errorMessage);
      } finally {
        event.target._extracting = false;
      }
    }
  } else {
    if (fileNameElement) {
      fileNameElement.classList.remove('show');
    }
    
    if (isProofOfPurchase) {
      const fileUploadContainer = fileNameElement?.closest('.file-upload-container');
      fileUploadContainer?.querySelectorAll('.extraction-success, .extraction-error, .ai-extracting').forEach(el => el.remove());
    }
  }
}

function showUpsFields() {
  const upsFields = document.getElementById('upsFields');
  if (upsFields) upsFields.classList.remove('hidden');
}

function hideUpsFields() {
  const upsFields = document.getElementById('upsFields');
  if (upsFields) upsFields.classList.add('hidden');
}

function validateForm(data) {
  const requiredFields = ['trackingNumber', 'incidenceType', 'description', 'customerAddress'];
  const missingFields = requiredFields.filter(field => !data[field] || data[field].trim() === '');
  
  if (missingFields.length > 0) {
    alert(`Please fill in all required fields: ${missingFields.join(', ')}`);
    return false;
  }
  
  // Validate proof of purchase file
  const proofOfPurchaseInput = document.getElementById('proofOfPurchase');
  if (proofOfPurchaseInput && !proofOfPurchaseInput.files[0]) {
    alert('Please upload a proof of purchase file');
    return false;
  }
  
  // Validate amount format
  const items = getShipmentItems();
  for (const item of items) {
    if (item.amount) {
      const amountStr = item.amount.toString();
      if (!amountStr.includes('.')) {
        alert(`Amount must include a decimal point (e.g., ${item.amount}.00). Please check item: ${item.description || 'unnamed item'}`);
        return false;
      }
      const decimalParts = amountStr.split('.');
      if (decimalParts.length === 2 && decimalParts[1].length < 2) {
        alert(`Amount must have at least 2 decimal places (e.g., ${item.amount}0). Please check item: ${item.description || 'unnamed item'}`);
        return false;
      }
      const numValue = parseFloat(amountStr);
      if (isNaN(numValue) || numValue < 0) {
        alert(`Please enter a valid amount for item: ${item.description || 'unnamed item'}`);
        return false;
      }
    }
  }
  
  return true;
}

// Shipment items management
function addShipmentItemRow(prefill = { description: '', quantity: '1', amount: '' }) {
  const itemsList = document.getElementById('itemsList');
  if (!itemsList) return;
  
  const row = document.createElement('div');
  row.className = 'item-row';
  row.innerHTML = `
    <div class="item-line-top">
      <div class="item-field item-desc">
        <label>Item Description *</label>
        <input type="text" placeholder="Enter item description" value="${escapeHtml(prefill.description || '')}" required>
      </div>
      <div class="item-field item-qty">
        <label>Quantity *</label>
        <input type="number" min="1" step="1" value="${prefill.quantity || '1'}" required>
      </div>
    </div>
    <div class="item-line-bottom">
      <div class="item-field item-curr">
        <label>Currency</label>
        <select class="item-currency" name="itemCurrency">
          <option value="EUR" ${prefill.currency === 'EUR' ? 'selected' : ''}>EUR</option>
          <option value="USD" ${prefill.currency === 'USD' ? 'selected' : ''}>USD</option>
          <option value="GBP" ${prefill.currency === 'GBP' ? 'selected' : ''}>GBP</option>
          <option value="CHF" ${prefill.currency === 'CHF' ? 'selected' : ''}>CHF</option>
          <option value="JPY" ${prefill.currency === 'JPY' ? 'selected' : ''}>JPY</option>
          <option value="CAD" ${prefill.currency === 'CAD' ? 'selected' : ''}>CAD</option>
          <option value="AUD" ${prefill.currency === 'AUD' ? 'selected' : ''}>AUD</option>
          <option value="CNY" ${prefill.currency === 'CNY' ? 'selected' : ''}>CNY</option>
        </select>
      </div>
      <div class="item-field item-amt">
        <label>Amount *</label>
        <input class="amount-only" type="number" step="0.01" min="0" placeholder="0.00" value="${prefill.amount || ''}" required>
      </div>
      <button type="button" class="remove-item-btn" aria-label="Remove item">
        <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
          <polyline points="3 6 5 6 21 6"/>
          <path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/>
          <path d="M10 11v6"/>
          <path d="M14 11v6"/>
          <path d="M9 6V4a2 2 0 0 1 2-2h2a2 2 0 0 1 2 2v2"/>
        </svg>
      </button>
    </div>
  `;
  
  itemsList.appendChild(row);
  
  // Event listeners
  row.querySelector('.remove-item-btn').addEventListener('click', () => {
    row.remove();
    recalcItemsTotals();
  });
  
  row.querySelectorAll('input').forEach(inp => {
    inp.addEventListener('input', recalcItemsTotals);
    if (inp.classList.contains('amount-only')) {
      inp.addEventListener('blur', (e) => {
        const value = e.target.value;
        if (value && value.trim() !== '') {
          const numValue = parseFloat(value);
          if (!isNaN(numValue) && numValue >= 0) {
            const formatted = numValue.toFixed(2);
            if (e.target.value !== formatted) {
              e.target.value = formatted;
              recalcItemsTotals();
            }
          }
        }
      });
    }
  });
  
  const currencySelect = row.querySelector('.item-currency');
  if (currencySelect) {
    currencySelect.addEventListener('change', recalcItemsTotals);
  }
  
  recalcItemsTotals();
}

function getShipmentItems() {
  const itemsList = document.getElementById('itemsList');
  if (!itemsList) return [];
  
  const rows = Array.from(itemsList.querySelectorAll('.item-row'));
  return rows.map(row => {
    const descInput = row.querySelector('input[type="text"]');
    const qtyInput = row.querySelector('input[type="number"]');
    const amountInputs = row.querySelectorAll('.item-amt input[type="number"]');
    const amountInput = amountInputs[0] || null;
    const currencySelect = row.querySelector('.item-currency');
    const currency = currencySelect ? currencySelect.value : 'EUR';
    
    return {
      description: descInput ? descInput.value.trim() : '',
      quantity: qtyInput ? qtyInput.value : '1',
      amount: amountInput ? amountInput.value : '',
      currency: currency
    };
  }).filter(it => it.description || it.amount || it.quantity);
}

function recalcItemsTotals() {
  const items = getShipmentItems();
  const total = items.reduce((sum, it) => sum + (Number(it.amount) || 0) * (Number(it.quantity) || 0), 0);
  
  const actualAmountInput = document.getElementById('actualAmount');
  if (actualAmountInput) actualAmountInput.value = total ? total.toFixed(2) : '';
  
  const contentsDescriptionInput = document.getElementById('contentsDescription');
  if (contentsDescriptionInput) {
    const lines = items.map(it => {
      const currency = it.currency || 'EUR';
      return `${Number(it.quantity)||0} x ${it.description} @ ${(Number(it.amount)||0).toFixed(2)} ${currency}`;
    });
    contentsDescriptionInput.value = lines.join('\n');
  }
}

function escapeHtml(str) {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

// Email content fetching
async function fetchEmailContent() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.getAsync(
      Office.CoercionType.Html,
      { asyncContext: "This is passed to the callback" },
      function (htmlResult) {
        Office.context.mailbox.item.body.getAsync(
          Office.CoercionType.Text,
          { asyncContext: "This is passed to the callback" },
          function (textResult) {
            if (textResult.status === Office.AsyncResultStatus.Failed && htmlResult.status === Office.AsyncResultStatus.Failed) {
              reject(new Error(`Failed to get email body: ${(textResult.error || htmlResult.error).message}`));
              return;
            }
            
            const subject = Office.context.mailbox.item.subject;
            const senderEmail = Office.context.mailbox.item.from?.emailAddress || '';
            const htmlBody = htmlResult.status === Office.AsyncResultStatus.Succeeded ? htmlResult.value : '';
            const textBody = textResult.status === Office.AsyncResultStatus.Succeeded ? textResult.value : '';
            
            resolve({
              subject: subject || '',
              body: textBody || htmlBody || '',
              senderEmail: senderEmail,
              htmlBody: htmlBody || textBody || ''
            });
          }
        );
      }
    );
  });
}

// Gemini AI integration
async function summarizeWithGemini(emailData) {
  const emailText = `Subject: ${emailData.subject || ''}\n\nBody:\n${emailData.body || ''}`;
  
  const prompt = `Analyze this Outlook email and extract information to pre-fill a logistics claim form. Return ONLY a valid JSON object with the following structure. If information is not available, use null for that field.

Required JSON structure:
{
  "carrier": "ups_001" or null,
  "trackingNumber": "string" or null,
  "incidenceType": "DELIVERED_LATE" | "DAMAGED" | "LOST" | "NOT_DELIVERED" or null,
  "description": "string" or null,
  "customerAddress": "string" or null,
  "contentsDescription": "string" or null,
  "currency": "EUR" | "USD" | "GBP" or null,
  "actualAmount": "number" or null
}

Guidelines:
- Look for tracking numbers (patterns like TRK, UPS, FedEx, etc.)
- Identify the type of incident (damaged, lost, late delivery, not delivered)
- Extract customer addresses from the messages
- Find descriptions of package contents
- Look for monetary amounts and currency
- Extract incident descriptions
- If carrier is mentioned (UPS, FedEx, DHL, etc.), use "ups_001" for now
- Return ONLY the JSON object, no additional text

Email Content:
${emailText}`;

  try {
    const response = await fetch(`${GEMINI_API_URL}?key=${GEMINI_API_KEY}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        contents: [{
          parts: [{ text: prompt }]
        }],
        generationConfig: {
          temperature: 0.1,
          topK: 40,
          topP: 0.95,
          maxOutputTokens: 1024,
        }
      })
    });
    
    if (!response.ok) {
      throw new Error(`Gemini API error: ${response.status} ${response.statusText}`);
    }
    
    const result = await response.json();
    const responseText = result.candidates?.[0]?.content?.parts?.[0]?.text || '{}';
    
    // Parse JSON - handle markdown code blocks
    let jsonText = responseText.trim();
    if (jsonText.startsWith('```json')) {
      jsonText = jsonText.replace(/^```json\s*/, '').replace(/\s*```$/, '');
    } else if (jsonText.startsWith('```')) {
      jsonText = jsonText.replace(/^```\s*/, '').replace(/\s*```$/, '');
    }
    
    return JSON.parse(jsonText);
  } catch (error) {
    console.error("Gemini API request failed:", error);
    throw new Error(`Failed to extract form data: ${error.message}`);
  }
}

// API submission
async function submitClaim(data) {
  // Get user email from Outlook
  let userName = 'user@example.com';
  try {
    if (Office?.context?.mailbox?.userProfile) {
      userName = Office.context.mailbox.userProfile.emailAddress || userName;
    }
  } catch (error) {
    console.warn('Could not get user email from Outlook:', error);
  }
  
  // Collect documents
  const documents = [];
  
  const proofOfPurchaseInput = document.getElementById('proofOfPurchase');
  if (proofOfPurchaseInput?.files[0]) {
    try {
      const file = proofOfPurchaseInput.files[0];
      const base64Data = await fileToBase64(file);
      documents.push({
        name: file.name,
        mimeType: file.type,
        base64Data: base64Data,
        size: file.size
      });
    } catch (error) {
      console.error('Error converting proof of purchase to base64:', error);
    }
  }
  
  const supportingEvidenceInput = document.getElementById('supportingEvidence');
  if (supportingEvidenceInput?.files[0]) {
    try {
      const file = supportingEvidenceInput.files[0];
      const base64Data = await fileToBase64(file);
      documents.push({
        name: file.name,
        mimeType: file.type,
        base64Data: base64Data,
        size: file.size
      });
    } catch (error) {
      console.error('Error converting supporting evidence to base64:', error);
    }
  }
  
  // Prepare payload (matching exact API structure)
  const payload = {
    source: 'zendesk',
    organizationId: API_ORGANIZATION_ID,
    userName: userName,
    shipment_trackingNumber: data.trackingNumber || '',
    shipment_carrierId: data.carrier || '',
    shipment_descriptionOfContents: data.contentsDescription || '[]',
    shipment_customerAddress: data.customerAddress || '',
    incidence_incidenceType: data.incidenceType || '',
    incidence_description: data.description || '',
    incidence_actualAmount: data.actualAmount || '0.00',
    documents: documents.length > 0 ? documents : []
  };
  
  // Submit to API
  const response = await fetch(API_ENDPOINT, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${API_AUTH_TOKEN}`
    },
    body: JSON.stringify(payload)
  });
  
  if (!response.ok) {
    const errorText = await response.text();
    console.error(`API Error: ${response.status} ${response.statusText}`, errorText);
    throw new Error(`HTTP error! status: ${response.status}, message: ${errorText}`);
  }
  
  return await response.json().catch(() => response.text());
}

function resetForm() {
  const form = document.getElementById('claimForm');
  if (form) form.reset();
  
  hideUpsFields();
  
  const itemsList = document.getElementById('itemsList');
  if (itemsList) {
    itemsList.innerHTML = '';
    addShipmentItemRow();
  }
}

function prefillForm(extractedData) {
  setTimeout(() => {
    if (extractedData.carrier) {
      const carrierSelect = document.getElementById('carrierSelect');
      if (carrierSelect) {
        carrierSelect.value = extractedData.carrier;
        carrierSelect.dispatchEvent(new Event('change'));
      }
    }
    
    if (extractedData.trackingNumber) {
      const trackingInput = document.getElementById('trackingNumber');
      if (trackingInput) trackingInput.value = extractedData.trackingNumber;
    }
    
    if (extractedData.incidenceType) {
      const incidenceSelect = document.getElementById('incidenceType');
      if (incidenceSelect) {
        incidenceSelect.value = extractedData.incidenceType;
        incidenceSelect.dispatchEvent(new Event('change'));
      }
    }
    
    if (extractedData.description) {
      const descriptionTextarea = document.getElementById('description');
      if (descriptionTextarea) descriptionTextarea.value = extractedData.description;
    }
    
    if (extractedData.customerAddress) {
      const addressTextarea = document.getElementById('customerAddress');
      if (addressTextarea) addressTextarea.value = extractedData.customerAddress;
    }
    
    if (extractedData.contentsDescription) {
      const contentsTextarea = document.getElementById('contentsDescription');
      if (contentsTextarea) contentsTextarea.value = extractedData.contentsDescription;
    }
    
    if (extractedData.currency) {
      const currencySelect = document.getElementById('currency');
      if (currencySelect) currencySelect.value = extractedData.currency;
    }
    
    if (extractedData.actualAmount) {
      const amountInput = document.getElementById('actualAmount');
      if (amountInput) amountInput.value = extractedData.actualAmount;
    }
  }, 100);
}

function populateSuccessView(data, apiResult) {
  const now = new Date();
  const dateStr = now.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });

  const setText = (id, value) => {
    const el = document.getElementById(id);
    if (el) el.textContent = value || '-';
  };

  let claimId = '-';
  if (apiResult) {
    try {
      const result = typeof apiResult === 'string' ? JSON.parse(apiResult) : apiResult;
      claimId = result.data?.claim_id || result.data?.id || result.claim_id || result.id || result.claimId || '-';
    } catch (error) {
      console.warn('Could not parse claim ID from API response:', error);
    }
  }

  setText('successClaimId', claimId);
  setText('successTracking', data.trackingNumber || '-');
  setText('successCreatedOn', dateStr);
  setText('successStatus', 'Created');
  
  window.lastClaimId = claimId;
}
