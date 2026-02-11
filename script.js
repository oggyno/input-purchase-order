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

    // Strategy: Find all item patterns directly in the full table text
    // Pattern matches: No + ItemCode + Description + Qty + UnitPrice + Disc + Amount
    
    console.log('🔍 Parsing strategy: Direct pattern matching with proper column detection...');
    
    // Use a robust pattern that captures all 7 columns
    // Format: "17   2017001315   MAGNALUX PENETRANT 1 SET (CLEANER - DEVELOPER -  PENETRANT)  3   195.000   0   585.000"
    
    const itemPattern = /(\d+)\s{3,}([\w-]+)\s{3,}(.+?)\s{3,}(\d+)\s{3,}([\d,.]+)\s{3,}([\d,.]+)\s{3,}([\d,.]+)/g;
    
    let match;
    let itemCount = 0;
    
    while ((match = itemPattern.exec(tableText)) !== null) {
        const [fullMatch, no, itemCode, rawDesc, qty, unitPrice, disc, amount] = match;
        
        // Clean description - remove excessive spaces and newlines
        const desc = rawDesc.replace(/\s+/g, ' ').trim();
        
        console.log(`   🔍 Item ${itemCount + 1}: no="${no}" | item="${itemCode}" | desc="${desc.substring(0, 40)}..." | qty="${qty}"`);
        
        // Validation
        const isValidNo = /^\d+$/.test(no) && parseInt(no) >= 1 && parseInt(no) <= 100;
        const isValidItem = itemCode && itemCode.length >= 2 && /^[\w-]+$/.test(itemCode);
        const isValidDesc = desc && desc.length >= 3 && desc.length < 500;
        const isValidQty = /^\d+$/.test(qty) && parseInt(qty) > 0 && parseInt(qty) < 10000;
        const isValidPrice = /^[\d,.]+$/.test(unitPrice);
        const isValidDisc = /^[\d,.]+$/.test(disc);
        const isValidAmount = /^[\d,.]+$/.test(amount);
        
        if (!isValidNo || !isValidItem || !isValidDesc || !isValidQty || !isValidPrice || !isValidDisc || !isValidAmount) {
            console.log(`      ❌ Skipped: no=${isValidNo}, item=${isValidItem}, desc=${isValidDesc}, qty=${isValidQty}, price=${isValidPrice}`);
            continue;
        }
        
        console.log(`      ✅ Valid: ${no} | ${itemCode} | ${qty}`);
        
        data.items.push({
            no: no,
            item: itemCode,
            namaBarang: desc,
            quantity: qty
        });
        itemCount++;
    }
    
    console.log(`📊 Found ${itemCount} valid items`);
    
    if (itemCount > 0) {
        console.log('✅ Parsing complete');
        return data;
    }

    // Fallback: If no items found
    console.log('⚠️ No items found with direct pattern.');
    
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
            
            if (poData.items.length === 0) {
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
