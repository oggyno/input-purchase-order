// Global variables
let selectedFiles = [];
let config = {
    webAppUrl: '',
    sheetName: 'Dataset Purchase Order'
};

// Initialize PDF.js worker
if (typeof pdfjsLib !== 'undefined') {
    pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
}

// Event listeners
window.addEventListener('DOMContentLoaded', function() {
    const fileInput = document.getElementById('fileInput');
    const sheetNameInput = document.getElementById('sheetName');
    
    if (fileInput) {
        fileInput.addEventListener('change', handleFileSelect);
    }
    
    if (sheetNameInput) {
        sheetNameInput.addEventListener('input', updateSheetNameDisplay);
    }
});

// Copy code function
function copyCode() {
    const code = document.getElementById('appsScriptCode').textContent;
    navigator.clipboard.writeText(code).then(() => {
        const btn = document.querySelector('.copy-btn');
        const originalText = btn.textContent;
        btn.textContent = '✓ Copied!';
        btn.style.background = '#10b981';
        setTimeout(() => {
            btn.textContent = originalText;
            btn.style.background = '#667eea';
        }, 2000);
    }).catch(err => {
        alert('Gagal copy. Silakan copy manual.');
    });
}

// Save configuration
function saveConfig() {
    const webAppUrlInput = document.getElementById('webAppUrl');
    const sheetNameInput = document.getElementById('sheetName');
    
    if (!webAppUrlInput || !sheetNameInput) {
        alert('Error: Form tidak ditemukan. Refresh halaman dan coba lagi.');
        return;
    }
    
    config.webAppUrl = webAppUrlInput.value.trim();
    config.sheetName = sheetNameInput.value.trim();

    if (!config.webAppUrl) {
        alert('❌ Mohon isi Web App URL dari Google Apps Script');
        return;
    }

    if (!config.webAppUrl.includes('script.google.com')) {
        alert('❌ URL tidak valid. Pastikan URL dari Google Apps Script\n\nFormat: https://script.google.com/macros/s/.../exec');
        return;
    }

    if (!config.webAppUrl.endsWith('/exec')) {
        alert('⚠️ URL harus diakhiri dengan /exec\n\nPastikan menggunakan Web App URL, bukan Test deployment URL');
        return;
    }

    document.getElementById('configSection').style.display = 'none';
    document.getElementById('editConfigBtn').style.display = 'block';
    document.getElementById('uploadSection').style.display = 'block';

    alert('✅ Konfigurasi berhasil disimpan!\n\nSekarang Anda bisa upload file PDF/gambar.');
}

// Toggle configuration visibility
function toggleConfig() {
    const configSection = document.getElementById('configSection');
    const editBtn = document.getElementById('editConfigBtn');
    const isVisible = configSection.style.display !== 'none';
    
    configSection.style.display = isVisible ? 'none' : 'block';
    editBtn.textContent = isVisible ? '⚙️ Edit Konfigurasi' : '❌ Tutup Konfigurasi';
}

// Update sheet name display
function updateSheetNameDisplay() {
    const sheetName = document.getElementById('sheetName').value;
    const displayElement = document.getElementById('sheetNameDisplay');
    if (displayElement) {
        displayElement.textContent = sheetName;
    }
    document.querySelectorAll('.sheet-name-ref').forEach(el => {
        el.textContent = sheetName;
    });
}

// Handle file selection
function handleFileSelect(event) {
    selectedFiles = Array.from(event.target.files);
    
    if (selectedFiles.length === 0) {
        document.getElementById('fileList').style.display = 'none';
        document.getElementById('processBtn').disabled = true;
        return;
    }

    displayFileList();
    document.getElementById('fileList').style.display = 'block';
    document.getElementById('processBtn').disabled = false;
}

// Display selected files
function displayFileList() {
    const fileListContent = document.getElementById('fileListContent');
    const fileCount = document.getElementById('fileCount');
    
    fileCount.textContent = selectedFiles.length;
    fileListContent.innerHTML = '';

    selectedFiles.forEach(file => {
        const fileItem = document.createElement('div');
        fileItem.className = 'file-item';
        
        const icon = file.type === 'application/pdf' ? '📄' : '🖼️';
        const sizeKB = (file.size / 1024).toFixed(1);
        
        fileItem.innerHTML = `
            ${icon} ${file.name}
            <span class="file-size">(${sizeKB} KB)</span>
        `;
        
        fileListContent.appendChild(fileItem);
    });
}

// Extract text from PDF
async function extractTextFromPDF(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = async (e) => {
            try {
                const typedArray = new Uint8Array(e.target.result);
                const pdf = await pdfjsLib.getDocument(typedArray).promise;
                let fullText = '';
                
                for (let i = 1; i <= pdf.numPages; i++) {
                    const page = await pdf.getPage(i);
                    const textContent = await page.getTextContent();
                    const pageText = textContent.items.map(item => item.str).join(' ');
                    fullText += pageText + '\n';
                }
                
                resolve(fullText);
            } catch (error) {
                reject(error);
            }
        };
        
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

// Extract text from image using OCR
async function extractTextFromImage(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = async (e) => {
            try {
                const worker = await Tesseract.createWorker('eng');
                const { data: { text } } = await worker.recognize(e.target.result);
                await worker.terminate();
                
                resolve(text);
            } catch (error) {
                reject(error);
            }
        };
        
        reader.onerror = reject;
        reader.readAsDataURL(file);
    });
}

// Parse PO data from text
function parsePoData(text) {
    const data = {
        poNumber: '',
        poDate: '',
        supplier: '',
        items: [],
        description: ''
    };

    // Extract PO Number
    const poNumberMatch = text.match(/PO Number\s*:\s*(\S+)/i);
    if (poNumberMatch) data.poNumber = poNumberMatch[1].trim();

    // Extract PO Date (fix: only get the date, not everything after)
    const poDateMatch = text.match(/PO Date\s*:\s*([^\n]+?)(?:\s{2,}|$)/i);
    if (poDateMatch) data.poDate = poDateMatch[1].trim();

    // Extract Supplier - get line after "To" but before next field
    const supplierMatch = text.match(/:\s*([A-Z][A-Z\s,\.]+(?:TOKO|PT|CV)?)\s+We confirm/i);
    if (supplierMatch) {
        data.supplier = supplierMatch[1].trim();
    } else {
        // Fallback pattern
        const fallbackMatch = text.match(/To\s+(?:Att|CC|Fax|Tel)[\s\S]{0,200}:\s*([A-Z][A-Z\s,\.]+)/i);
        if (fallbackMatch) data.supplier = fallbackMatch[1].trim();
    }

    // Extract Description at the bottom
    const descMatch = text.match(/Description\s*-\s*([^\n]+)/i);
    if (descMatch) data.description = descMatch[1].trim();

    // Find the table section
    const tableStartMatch = text.match(/No\.\s+Item\s+Description\s+Qty/i);
    const tableEndMatch = text.match(/Requested By/i);
    
    let tableText = text;
    if (tableStartMatch && tableEndMatch) {
        const startIdx = text.indexOf(tableStartMatch[0]) + tableStartMatch[0].length;
        const endIdx = text.indexOf(tableEndMatch[0]);
        tableText = text.substring(startIdx, endIdx);
        
        console.log('📋 Table text extracted (first 500 chars):', tableText.substring(0, 500));
        console.log('📏 Table text length:', tableText.length);
    }

    // Strategy 1: Split and parse (better for complex descriptions)
    console.log('🔍 Strategy 1: Split by pattern and parse...');
    
    // First, let's try to find individual item lines
    // Look for pattern: starts with number, followed by item code
    const itemLines = [];
    const lines = tableText.split('\n');
    
    let currentLine = '';
    for (let i = 0; i < lines.length; i++) {
        const line = lines[i].trim();
        
        // Check if line starts with a number (potential item number)
        if (/^\d+\s/.test(line)) {
            // Save previous line if exists
            if (currentLine) {
                itemLines.push(currentLine);
            }
            currentLine = line;
        } else if (currentLine && line) {
            // Continuation of previous line
            currentLine += ' ' + line;
        }
    }
    // Don't forget last line
    if (currentLine) {
        itemLines.push(currentLine);
    }
    
    console.log(`📝 Found ${itemLines.length} potential item lines`);
    
    // Now parse each line
    let matchCount = 0;
    for (let i = 0; i < itemLines.length; i++) {
        const line = itemLines[i];
        
        // Split by 3+ spaces
        const parts = line.split(/\s{3,}/);
        
        if (parts.length < 7) {
            console.log(`   ⚠️ Line ${i}: Not enough parts (${parts.length}): "${line.substring(0, 60)}..."`);
            continue;
        }
        
        // Expected format: [no, itemCode, description, qty, unitPrice, disc, amount]
        const no = parts[0].trim();
        const itemCode = parts[1].trim();
        const desc = parts[2].trim();
        const qty = parts[3].trim();
        const unitPrice = parts[4].trim();
        const disc = parts[5].trim();
        const amount = parts[6].trim();
        
        console.log(`   🔍 Line ${i}: no="${no}" | item="${itemCode}" | desc="${desc.substring(0, 30)}..." | qty="${qty}" | price="${unitPrice}" | disc="${disc}"`);
        
        // Validation
        const isValidNo = /^\d+$/.test(no);
        const isValidItem = itemCode && itemCode.length >= 2;
        const isValidDesc = desc && desc.length >= 3;
        const isValidQty = /^\d+$/.test(qty) && parseInt(qty) > 0;
        const isValidPrice = /^[\d,.]+$/.test(unitPrice);
        const isValidDisc = /^[\d,.]+$/.test(disc);
        
        if (!isValidNo || !isValidItem || !isValidDesc || !isValidQty || !isValidPrice || !isValidDisc) {
            console.log(`      ❌ Validation failed: no=${isValidNo}, item=${isValidItem}, desc=${isValidDesc}, qty=${isValidQty}, price=${isValidPrice}, disc=${isValidDisc}`);
            continue;
        }
        
        console.log(`      ✅ Valid item ${matchCount + 1}: ${no} | ${itemCode} | ${desc.substring(0, 40)} | ${qty}`);
        
        data.items.push({
            no: no,
            item: itemCode,
            namaBarang: desc,
            quantity: qty
        });
        matchCount++;
    }
    
    console.log(`📊 Split method found ${matchCount} items`);
    
    // If split method worked, return early
    if (matchCount > 0) {
        console.log('✅ Using split method results');
        return data;
    }
    
    // Strategy 2: Try direct regex matching (fallback)
    console.log('🔍 Strategy 2: Direct regex matching (fallback)...');
    
    // Pattern explanation:
    // Format: No | Item | Description | Qty | Unit Price | Disc | Amount
    // Example: 1   11MGPT4PP   MATA GERINDA POTONG 4", BRAND : "WD" (IN PCS)   5   3.378,39   0   16.891,95
    
    // We need to capture until we hit: qty + unit_price + disc (usually 0) + amount
    // Pattern: (\d+)\s{3,}(item)\s{3,}(desc)\s{3,}(\d+)\s{3,}[\d,.]+\s{3,}[\d,.]+\s{3,}[\d,.]+
    //          no           itemCode      description    qty      unit_price   disc      amount
    
    const directPattern = /(\d+)\s{3,}([\w-]+)\s{3,}(.+?)\s{3,}(\d+)\s{3,}[\d,.]+\s{3,}[\d,.]+\s{3,}[\d,.]+/g;
    let match;
    let matchCount = 0;
    
    while ((match = directPattern.exec(tableText)) !== null) {
        const [fullMatch, no, itemCode, desc, qty] = match;
        
        console.log(`   🔍 Raw match: "${fullMatch.substring(0, 100)}..."`);
        console.log(`      Parsed: no="${no}" | item="${itemCode}" | desc="${desc.substring(0, 40)}..." | qty="${qty}"`);
        
        // Clean description - remove excessive whitespace
        const cleanDesc = desc.trim().replace(/\s{2,}/g, ' ');
        
        // Validation
        const isValidNo = /^\d+$/.test(no);
        const isValidItem = itemCode && itemCode.length >= 2;
        const isValidDesc = cleanDesc && cleanDesc.length >= 3;
        const isValidQty = /^\d+$/.test(qty) && parseInt(qty) > 0; // Qty must be > 0
        
        if (!isValidNo || !isValidItem || !isValidDesc || !isValidQty) {
            console.log(`      ❌ Validation failed: no=${isValidNo}, item=${isValidItem}, desc=${isValidDesc}, qty=${isValidQty} (value: ${qty})`);
            continue;
        }
        
        console.log(`      ✅ Valid item ${matchCount + 1}: ${no} | ${itemCode} | ${cleanDesc.substring(0, 40)} | ${qty}`);
        
        data.items.push({
            no: no,
            item: itemCode,
            namaBarang: cleanDesc,
            quantity: qty
        });
        matchCount++;
    }
    
    console.log(`📊 Direct regex found ${matchCount} items`);
    
    // If direct matching worked, return early
    if (matchCount > 0) {
        console.log('✅ Using direct regex results');
        return data;
    }

    // Strategy 2: Split by newlines
    console.log('🔍 Strategy 2: Trying newline split...');
    let lines = tableText.split('\n');
    
    // If no newlines, try other separators
    if (lines.length === 1) {
        console.log('⚠️ No \\n found, trying \\r\\n...');
        lines = tableText.split('\r\n');
    }
    if (lines.length === 1) {
        console.log('⚠️ No \\r\\n found, trying manual parsing...');
        // If still one line, try to find patterns directly in the text
        const itemPattern = /(\d+)\s{2,}(\S+)\s{2,}(.+?)\s{2,}(\d+)\s{2,}[\d,.]+\s{2,}[\d,.]+\s{2,}[\d,.]+/g;
        let match;
        while ((match = itemPattern.exec(tableText)) !== null) {
            const [, no, itemCode, desc, qty] = match;
            console.log(`   ✅ Direct match: ${no} | ${itemCode} | ${desc} | ${qty}`);
            data.items.push({
                no: no.trim(),
                item: itemCode.trim(),
                namaBarang: desc.trim(),
                quantity: qty.trim()
            });
        }
        return data; // Return early if we used direct matching
    }
    
    console.log(`📝 Split into ${lines.length} lines`);
    
    for (let i = 0; i < lines.length; i++) {
        const line = lines[i].trim();
        
        // Skip empty lines
        if (!line) {
            console.log(`🔍 Line ${i}: [empty, skipped]`);
            continue;
        }
        
        // Skip header and footer lines
        if (line.match(/^(No\.|Item|Description|Qty|Unit|Price|Disc|Amount|Sub Total|VAT|Total|Requested|Authorized|Say|Catatan)/i)) {
            console.log(`🔍 Line ${i}: [header/footer, skipped] "${line.substring(0, 50)}..."`);
            continue;
        }
        
        console.log(`🔍 Parsing line ${i}: "${line}"`);
        
        // Method 1: Split by 2 or more spaces
        const parts = line.split(/\s{2,}/);
        console.log('   Parts:', parts);
        
        if (parts.length >= 4) {
            const firstPart = parts[0].trim();
            
            // Check if first part is a single digit (item number)
            if (/^\d+$/.test(firstPart)) {
                const no = firstPart;
                const itemCode = parts[1] ? parts[1].trim() : '';
                const desc = parts[2] ? parts[2].trim() : '';
                const qty = parts[3] ? parts[3].trim() : '';
                
                // Validate quantity is a number
                if (no && itemCode && desc && /^\d+$/.test(qty)) {
                    console.log(`   ✅ Found item: ${no} | ${itemCode} | ${desc} | ${qty}`);
                    data.items.push({
                        no: no,
                        item: itemCode,
                        namaBarang: desc,
                        quantity: qty
                    });
                    continue;
                } else {
                    console.log(`   ❌ Validation failed: qty="${qty}" is not a number`);
                }
            } else {
                console.log(`   ⚠️ First part "${firstPart}" is not a number`);
            }
        } else {
            console.log(`   ⚠️ Not enough parts: ${parts.length}`);
        }
        
        // Method 2: Regex pattern - very specific to this format
        // Format: "1   2017001017   KAPUR BESI @PERLUSIN   4   25.000   0   100.000"
        const match = line.match(/^(\d+)\s+(\S+)\s+(.+?)\s+(\d+)\s+[\d,.]+/);
        if (match) {
            const [, no, itemCode, desc, qty] = match;
            console.log(`   ✅ Regex matched: ${no} | ${itemCode} | ${desc} | ${qty}`);
            data.items.push({
                no: no.trim(),
                item: itemCode.trim(),
                namaBarang: desc.trim(),
                quantity: qty.trim()
            });
        } else {
            console.log(`   ❌ Regex didn't match`);
        }
    }

    return data;
}

// Send data to Google Sheets via Apps Script
async function sendToGoogleSheets(poData) {
    if (!config.webAppUrl) {
        throw new Error('Web App URL belum diisi');
    }

    const rows = poData.items.map(item => [
        poData.poNumber,
        poData.poDate,
        poData.supplier,
        item.item,
        item.namaBarang,
        item.quantity,
        poData.description
    ]);

    try {
        const response = await fetch(config.webAppUrl, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                sheetName: config.sheetName,
                rows: rows
            }),
            mode: 'no-cors'
        });

        // Mode no-cors tidak bisa baca response, tapi request tetap terkirim
        return { success: true };
    } catch (error) {
        throw new Error('Gagal mengirim ke Google Sheets: ' + error.message);
    }
}

// Process all files
async function processFiles() {
    if (selectedFiles.length === 0) {
        alert('❌ Pilih file PDF atau gambar terlebih dahulu');
        return;
    }

    // Show progress
    document.getElementById('processBtn').disabled = true;
    document.getElementById('progressSection').style.display = 'block';
    document.getElementById('resultsSection').style.display = 'none';

    const results = [];

    for (let i = 0; i < selectedFiles.length; i++) {
        const file = selectedFiles[i];
        
        try {
            // Update progress
            document.getElementById('progressText').textContent = 
                `⏳ Processing ${i + 1}/${selectedFiles.length}: ${file.name}...`;

            let text = '';
            
            // Check file type
            if (file.type === 'application/pdf') {
                text = await extractTextFromPDF(file);
            } else if (file.type.startsWith('image/')) {
                text = await extractTextFromImage(file);
            } else {
                throw new Error('Format file tidak didukung. Gunakan PDF atau gambar (JPG/PNG)');
            }
            
            const poData = parsePoData(text);
            
            // Debug: Log extracted data
            console.log('📊 Extracted PO Data:', {
                poNumber: poData.poNumber,
                poDate: poData.poDate,
                supplier: poData.supplier,
                itemCount: poData.items.length,
                description: poData.description,
                items: poData.items
            });
            
            if (poData.items.length === 0) {
                console.error('❌ No items found.');
                console.log('📄 Full text length:', text.length);
                console.log('📄 Text preview (first 1000 chars):', text.substring(0, 1000));
                console.log('📄 Text preview (chars 1000-2000):', text.substring(1000, 2000));
                
                // Try to show table section if exists
                const tableMatch = text.match(/No\.?\s+Item[\s\S]{0,500}/i);
                if (tableMatch) {
                    console.log('📋 Table section found:', tableMatch[0]);
                }
                
                throw new Error('Tidak ada data item yang ditemukan. Cek Console (F12) untuk detail.');
            }
            
            await sendToGoogleSheets(poData);
            
            results.push({
                fileName: file.name,
                status: 'success',
                poNumber: poData.poNumber || 'N/A',
                itemCount: poData.items.length,
                fileType: file.type.startsWith('image/') ? 'image' : 'pdf'
            });
        } catch (error) {
            results.push({
                fileName: file.name,
                status: 'error',
                error: error.message
            });
        }
    }

    // Hide progress and show results
    document.getElementById('progressSection').style.display = 'none';
    document.getElementById('processBtn').disabled = false;
    displayResults(results);
}

// Display results
function displayResults(results) {
    const resultsSection = document.getElementById('resultsSection');
    const resultsContent = document.getElementById('resultsContent');
    
    resultsContent.innerHTML = '';
    
    results.forEach(result => {
        const resultItem = document.createElement('div');
        resultItem.className = `result-item ${result.status === 'success' ? 'result-success' : 'result-error'}`;
        
        const icon = result.status === 'success' 
            ? '<svg class="result-icon result-icon-success" viewBox="0 0 24 24" fill="none" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" /></svg>'
            : '<svg class="result-icon result-icon-error" viewBox="0 0 24 24" fill="none" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 8v4m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" /></svg>';
        
        const fileIcon = result.fileType === 'image' ? '🖼️' : '📄';
        
        const message = result.status === 'success'
            ? `✓ Berhasil! PO: ${result.poNumber} (${result.itemCount} items)`
            : `✗ Error: ${result.error}`;
        
        const messageClass = result.status === 'success' 
            ? 'result-message-success' 
            : 'result-message-error';
        
        resultItem.innerHTML = `
            <div class="result-header">
                ${icon}
                <div class="result-content">
                    <p class="result-filename">${result.fileType ? fileIcon + ' ' : ''}${result.fileName}</p>
                    <p class="result-message ${messageClass}">${message}</p>
                </div>
            </div>
        `;
        
        resultsContent.appendChild(resultItem);
    });
    
    resultsSection.style.display = 'block';
    
    // Scroll to results
    resultsSection.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}
