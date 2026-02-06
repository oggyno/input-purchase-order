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

// Extract text from PDF (DIPERBAIKI: fallback OCR jika text kosong)
async function extractTextFromPDF(file) {
    return new Promise(async (resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = async (e) => {
            try {
                const typedArray = new Uint8Array(e.target.result);
                const pdf = await pdfjsLib.getDocument(typedArray).promise;
                let fullText = '';
                
                // Coba extract text native dulu
                for (let i = 1; i <= pdf.numPages; i++) {
                    const page = await pdf.getPage(i);
                    const textContent = await page.getTextContent();
                    const pageText = textContent.items.map(item => item.str).join(' ');
                    fullText += pageText + '\n';
                }
                
                console.log('✅ PDF native text extracted:', fullText.substring(0, 500) + '...');
                
                // Jika text terlalu pendek, coba OCR
                if (fullText.trim().length < 100) {
                    console.log('⚠️ Native text terlalu pendek, mencoba OCR...');
                    const ocrText = await ocrPdfPages(typedArray);
                    console.log('✅ PDF OCR text:', ocrText.substring(0, 500) + '...');
                    resolve(ocrText);
                } else {
                    resolve(fullText);
                }
            } catch (error) {
                console.error('❌ Native text extraction failed:', error);
                // Fallback: OCR via canvas render
                try {
                    const ocrText = await ocrPdfPages(typedArray);
                    console.log('✅ PDF OCR fallback success:', ocrText.substring(0, 500) + '...');
                    resolve(ocrText);
                } catch (ocrError) {
                    console.error('❌ OCR failed:', ocrError);
                    reject(ocrError);
                }
            }
        };
        
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

// Fungsi baru: OCR PDF pages via canvas + Tesseract
async function ocrPdfPages(arrayBuffer) {
    const pdf = await pdfjsLib.getDocument(arrayBuffer).promise;
    let fullText = '';
    
    for (let i = 1; i <= pdf.numPages; i++) {
        console.log(`OCR page ${i}/${pdf.numPages}...`);
        
        const page = await pdf.getPage(i);
        const scale = 2.0; // Tinggi untuk akurasi OCR
        const viewport = page.getViewport({ scale });
        
        const canvas = document.createElement('canvas');
        const context = canvas.getContext('2d');
        canvas.height = viewport.height;
        canvas.width = viewport.width;
        
        const renderContext = {
            canvasContext: context,
            viewport: viewport
        };
        
        await page.render(renderContext).promise;
        
        const worker = await Tesseract.createWorker('eng+ind'); // Support English + Indonesia
        const { data: { text } } = await worker.recognize(canvas);
        await worker.terminate();
        
        fullText += text + '\n--- PAGE ' + i + ' ---\n';
    }
    
    return fullText;
}

// Extract text from image using OCR
async function extractTextFromImage(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = async (e) => {
            try {
                console.log('OCR image:', file.name);
                const worker = await Tesseract.createWorker('eng+ind');
                const { data: { text } } = await worker.recognize(e.target.result);
                await worker.terminate();
                
                console.log('✅ Image OCR text:', text.substring(0, 500) + '...');
                resolve(text);
            } catch (error) {
                console.error('❌ Image OCR failed:', error);
                reject(error);
            }
        };
        
        reader.onerror = reject;
        reader.readAsDataURL(file);
    });
}

// Parse PO data (DIPERBAIKI: regex lebih fleksibel)
function parsePoData(text) {
    console.log('🔍 Parsing text sample:', text.substring(0, 1000));
    
    const data = {
        poNumber: '',
        poDate: '',
        supplier: '',
        items: [],
        description: ''
    };

    const normalizedText = text.toLowerCase().replace(/\s+/g, ' ').trim();

    // PO Number: pola lebih luas
    let poMatch = normalizedText.match(/(?:po\s*(?:number|no|#|num)\s*[:\-]?\s*([a-z0-9\-\/]{5,}))/i) ||
                  normalizedText.match(/(?:po[0-9]{4,}|order\s*(?:no|#)\s*[:\-]?\s*([a-z0-9\-\/]{5,}))/i) ||
                  normalizedText.match(/([pP][oO]\s*[0-9\-\/]{5,})/);
    if (poMatch) {
        data.poNumber = poMatch[1] ? poMatch[1].trim() : poMatch[0].trim();
        console.log('✅ PO Number found:', data.poNumber);
    }

    // PO Date: berbagai format tanggal
    let dateMatch = normalizedText.match(/(?:date|tanggal)\s*[:\-]?\s*(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})/i) ||
                    normalizedText.match(/(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})/);
    if (dateMatch) {
        data.poDate = dateMatch[1].trim();
        console.log('✅ PO Date found:', data.poDate);
    }

    // Supplier: berbagai pola
    let suppMatch = normalizedText.match(/(?:to|kepada|supplier|vendor)[:\-]?\s*([a-z\s,]+?)(?:\n|$|qty|item|\d)/i);
    if (suppMatch) {
        data.supplier = suppMatch[1].trim();
        console.log('✅ Supplier found:', data.supplier);
    }

    // Description
    let descMatch = normalizedText.match(/(?:description|deskripsi)[:\-]?\s*([^\n\r]+)/i);
    if (descMatch) {
        data.description = descMatch[1].trim();
    }

    // Items: pola tabel lebih fleksibel
    const lines = text.split(/\n|\r\n/);
    for (let i = 0; i < lines.length; i++) {
        let line = lines[i].trim();
        
        // Skip header/footer lines
        if (line.match(/total|subtotal|grand|tax|ppn/i) || line.length < 10) continue;
        
        // Pola 1: 1 ITEM001 Nama Barang 10
        let itemMatch1 = line.match(/^(\d+[.\)]\s?)?\s*([a-z0-9\-]{3,})\s+(.+?)\s+(\d+(?:,\d+)?)$/i);
        if (itemMatch1 && itemMatch1[2] && itemMatch1[4]) {
            data.items.push({
                no: itemMatch1[1] || '',
                item: itemMatch1[2].trim(),
                namaBarang: itemMatch1[3].trim(),
                quantity: itemMatch1[4].trim()
            });
            continue;
        }
        
        // Pola 2: ITEM001 Nama Barang 10 pcs
        let itemMatch2 = line.match(/^([a-z0-9\-]{3,})\s+(.+?)\s+(\d+(?:,\d+)?)/i);
        if (itemMatch2 && itemMatch2[1] && itemMatch2[3]) {
            data.items.push({
                no: '',
                item: itemMatch2[1].trim(),
                namaBarang: itemMatch2[2].trim(),
                quantity: itemMatch2[3].trim()
            });
        }
    }

    console.log('📊 Parsed data:', {
        poNumber: data.poNumber,
        poDate: data.poDate,
        supplier: data.supplier,
        itemCount: data.items.length,
        items: data.items.slice(0, 3) // Show first 3 items
    });
    
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

        console.log('✅ Data sent to Google Sheets:', rows.length, 'rows');
        return { success: true };
    } catch (error) {
        console.error('❌ Google Sheets error:', error);
        throw new Error('Gagal mengirim ke Google Sheets: ' + error.message);
    }
}

// Process all files (DIPERBAIKI: tidak error jika 0 items)
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
                `⏳ Processing ${i + 1}/${selectedFiles.length}: ${file.name}`;

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
            
            // Selalu kirim data, bahkan jika 0 items (kirim metadata saja)
            await sendToGoogleSheets(poData);
            
            results.push({
                fileName: file.name,
                status: poData.items.length > 0 ? 'success' : 'partial',
                poNumber: poData.poNumber || 'N/A',
                itemCount: poData.items.length,
                fileType: file.type.startsWith('image/') ? 'image' : 'pdf'
            });
        } catch (error) {
            console.error('❌ Process failed:', error);
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
        let statusClass, icon, message;
        
        if (result.status === 'success') {
            statusClass = 'result-success';
            icon = '<svg class="result-icon result-icon-success" viewBox="0 0 24 24" fill="none" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" /></svg>';
            message = `✓ Berhasil! PO: ${result.poNumber} (${result.itemCount} items)`;
        } else if (result.status === 'partial') {
            statusClass = 'result-warning';
            icon = '⚠️';
            message = `⚠️ Metadata OK, tapi 0 items ditemukan: ${result.fileName}`;
        } else {
            statusClass = 'result-error';
            icon = '<svg class="result-icon result-icon-error" viewBox="0 0 24 24" fill="none" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 8v4m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" /></svg>';
            message = `✗ Error: ${result.error}`;
        }
        
        const fileIcon = result.fileType === 'image' ? '🖼️' : '📄';
        
        resultItem.className = `result-item ${statusClass}`;
        resultItem.innerHTML = `
            <div class="result-header">
                <span class="result-icon-text">${icon}</span>
                <div class="result-content">
                    <p class="result-filename">${fileIcon} ${result.fileName}</p>
                    <p class="result-message">${message}</p>
                </div>
            </div>
        `;
        
        resultsContent.appendChild(resultItem);
    });
    
    resultsSection.style.display = 'block';
    resultsSection.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}
