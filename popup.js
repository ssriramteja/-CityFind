let addressDatabase = [];

document.addEventListener('DOMContentLoaded', function() {
  const fileInput = document.getElementById('fileInput');
  const uploadArea = document.getElementById('uploadArea');
  const searchInput = document.getElementById('searchInput');
  const resultsDiv = document.getElementById('results');
  const fileLoaded = document.getElementById('fileLoaded');
  const fileName = document.getElementById('fileName');
  const fileCount = document.getElementById('fileCount');
  const clearBtn = document.getElementById('clearBtn');
  
  // Load saved data from Chrome storage
  chrome.storage.local.get(['addressData', 'fileName'], function(result) {
    if (result.addressData && result.addressData.length > 0) {
      addressDatabase = result.addressData;
      showFileLoaded(result.fileName || 'addresses.xlsx', addressDatabase.length);
      searchInput.disabled = false;
    }
  });
  
  // Click to upload
  uploadArea.addEventListener('click', () => fileInput.click());
  
  // Drag and drop
  uploadArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadArea.classList.add('dragover');
  });
  
  uploadArea.addEventListener('dragleave', () => {
    uploadArea.classList.remove('dragover');
  });
  
  uploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadArea.classList.remove('dragover');
    const file = e.dataTransfer.files[0];
    if (file) {
      handleFile(file);
    }
  });
  
  // File input change
  fileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) {
      handleFile(file);
    }
  });
  
  // Handle file upload
  function handleFile(file) {
    const validTypes = [
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'application/vnd.ms-excel',
      'text/csv'
    ];
    
    if (!validTypes.includes(file.type) && !file.name.match(/\.(xlsx|xls|csv)$/i)) {
      showError('Please upload a valid Excel or CSV file');
      return;
    }
    
    // Check if this is the expected file
    const expectedFileName = 'us_cities_single_address_format1111';
    const isExpectedFile = file.name.toLowerCase().includes(expectedFileName.toLowerCase());
    
    showLoading(isExpectedFile ? 'Loading US cities addresses...' : 'Processing file...');
    
    const reader = new FileReader();
    reader.onload = function(e) {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        
        // Get first sheet
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(firstSheet);
        
        // Process data
        addressDatabase = processExcelData(jsonData);
        
        if (addressDatabase.length === 0) {
          showError('No valid address data found. Please check columns: address, city, state, zipcode');
          return;
        }
        
        // Save to Chrome storage
        chrome.storage.local.set({
          addressData: addressDatabase,
          fileName: file.name
        });
        
        showFileLoaded(file.name, addressDatabase.length);
        searchInput.disabled = false;
        showNoResults();
        
      } catch (error) {
        showError('Error reading file: ' + error.message);
      }
    };
    
    reader.onerror = function() {
      showError('Failed to read file');
    };
    
    reader.readAsArrayBuffer(file);
  }
  
  // Process Excel data with exact column names
  function processExcelData(data) {
    const processed = [];
    
    data.forEach(row => {
      // Match exact column names: address, city, state, zipcode
      const address = row.address || row.Address || row.ADDRESS;
      const city = row.city || row.City || row.CITY;
      const state = row.state || row.State || row.STATE;
      const zipcode = row.zipcode || row.Zipcode || row.ZIPCODE || row['ZIP Code'] || row['Zip Code'];
      
      if (address && city && state) {
        processed.push({
          address: String(address).trim(),
          city: String(city).trim(),
          state: String(state).trim(),
          zip: zipcode ? String(zipcode).trim() : ''
        });
      }
    });
    
    return processed;
  }
  
  // Clear data
  clearBtn.addEventListener('click', function() {
    chrome.storage.local.clear();
    addressDatabase = [];
    fileLoaded.classList.remove('show');
    searchInput.disabled = true;
    searchInput.value = '';
    fileInput.value = '';
    showNoResults('Upload an Excel file to get started', '📤');
  });
  
  // Show file loaded status
  function showFileLoaded(name, count) {
    fileName.textContent = name;
    fileCount.textContent = `${count} addresses loaded`;
    fileLoaded.classList.add('show');
  }
  
  // Search function
  function performSearch(query) {
    if (!query || query.trim().length === 0) {
      showNoResults();
      return;
    }
    
    if (addressDatabase.length === 0) {
      showNoResults('Please upload an Excel file first', '📤');
      return;
    }
    
    query = query.trim().toLowerCase();
    
    // Filter addresses based on query
    const results = addressDatabase.filter(addr => {
      const cityMatch = addr.city.toLowerCase().includes(query);
      const stateMatch = addr.state.toLowerCase() === query || addr.state.toLowerCase().includes(query);
      const zipMatch = addr.zip.includes(query);
      const fullLocation = `${addr.city.toLowerCase()} ${addr.state.toLowerCase()}`;
      const fullLocationMatch = fullLocation.includes(query);
      
      return cityMatch || stateMatch || zipMatch || fullLocationMatch;
    });
    
    displayResults(results);
  }
 
// Display results
function displayResults(results) {
  if (results.length === 0) {
    resultsDiv.innerHTML = `
      <div class="no-results">
        <div class="no-results-icon">😔</div>
        <p>No addresses found</p>
        <p style="font-size: 11px; margin-top: 5px;">Try searching by city, state code, or zip</p>
      </div>
    `;
    return;
  }
  
  // Limit to first 20 results
  const limitedResults = results.slice(0, 20);
  
  // Clear previous results
  resultsDiv.innerHTML = '';
  
  limitedResults.forEach(addr => {
    // Create result card
    const card = document.createElement('div');
    card.className = 'result-card';
    
    // Address line
    const addressDiv = document.createElement('div');
    addressDiv.className = 'address';
    addressDiv.textContent = addr.address;
    
    // Location line
    const locationDiv = document.createElement('div');
    locationDiv.className = 'location';
    locationDiv.innerHTML = `📍 ${escapeHtml(addr.city)}, ${escapeHtml(addr.state)} ${addr.zip ? `<span class="badge">${escapeHtml(addr.zip)}</span>` : ''}`;
    
    // Copy button
    const copyBtn = document.createElement('button');
    copyBtn.className = 'copy-btn';
    copyBtn.textContent = '📋 Copy Address';
    
    // Create full address for copying (NOT using data attribute)
    const fullAddress = `${addr.address}, ${addr.city}, ${addr.state} ${addr.zip}`.trim();
    
    copyBtn.addEventListener('click', function() {
      navigator.clipboard.writeText(fullAddress).then(() => {
        copyBtn.textContent = '✓ Copied!';
        copyBtn.classList.add('copied');
        setTimeout(() => {
          copyBtn.textContent = '📋 Copy Address';
          copyBtn.classList.remove('copied');
        }, 2000);
      });
    });
    
    // Append elements
    card.appendChild(addressDiv);
    card.appendChild(locationDiv);
    card.appendChild(copyBtn);
    resultsDiv.appendChild(card);
  });
  
  if (results.length > 20) {
    const moreDiv = document.createElement('div');
    moreDiv.style.cssText = 'text-align: center; padding: 15px; color: #6b7280; font-size: 12px;';
    moreDiv.textContent = `Showing 20 of ${results.length} results`;
    resultsDiv.appendChild(moreDiv);
  }
}

  // Show no results message
  function showNoResults(message = 'Enter a city, state, or zip code to search', icon = '📍') {
    resultsDiv.innerHTML = `
      <div class="no-results">
        <div class="no-results-icon">${icon}</div>
        <p>${message}</p>
        ${icon === '📤' ? '<p style="font-size: 11px; margin-top: 5px; color: #667eea; font-weight: 600;">Expected file: us_cities_single_address_format1111.xlsx</p><p style="font-size: 11px; margin-top: 3px;">Columns: address, city, state, zipcode</p>' : ''}
      </div>
    `;
  }
  
  // Show loading
  function showLoading(message = 'Loading addresses...') {
    resultsDiv.innerHTML = `
      <div class="loading">
        <div class="spinner"></div>
        <p>${message}</p>
      </div>
    `;
  }
  
  // Show error
  function showError(message) {
    resultsDiv.innerHTML = `
      <div class="error-message">
        <strong>❌ Error:</strong> ${message}
      </div>
      <div class="no-results">
        <div class="no-results-icon">📤</div>
        <p>Please try uploading again</p>
        <p style="font-size: 11px; margin-top: 5px;">Columns needed: address, city, state, zipcode</p>
      </div>
    `;
  }
  
  // Escape HTML
  function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
  }
  
  // Search on input with debounce
  let searchTimeout;
  searchInput.addEventListener('input', function() {
    clearTimeout(searchTimeout);
    searchTimeout = setTimeout(() => {
      performSearch(this.value);
    }, 300);
  });
  
  // Search on Enter key
  searchInput.addEventListener('keypress', function(e) {
    if (e.key === 'Enter') {
      clearTimeout(searchTimeout);
      performSearch(this.value);
    }
  });
});