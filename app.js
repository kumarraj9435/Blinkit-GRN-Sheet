// ==================== DATA STORE ====================
let masterData = JSON.parse(localStorage.getItem('blinkitGRN_masterData') || '[]');
let uploadHistory = JSON.parse(localStorage.getItem('blinkitGRN_uploadHistory') || '[]');
let selectedFiles = [];

// ==================== NAVIGATION ====================
document.querySelectorAll('.nav-item').forEach(item => {
    item.addEventListener('click', (e) => {
        e.preventDefault();
        const tab = item.dataset.tab;
        
        // Update nav
        document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
        item.classList.add('active');
        
        // Update content
        document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
        document.getElementById(tab).classList.add('active');
        
        // Update title
        const titles = {
            'dashboard': 'Dashboard',
            'upload': 'Upload Files',
            'master-sheet': 'Master Sheet',
            'comparison': 'R.O. vs GRN Comparison'
        };
        document.getElementById('page-title').textContent = titles[tab];
        
        // Refresh content
        if (tab === 'dashboard') updateDashboard();
        if (tab === 'master-sheet') renderMasterSheet();
        if (tab === 'comparison') renderComparison();
    });
});

// ==================== FILE UPLOAD ====================
const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');
const fileList = document.getElementById('fileList');
const processBtn = document.getElementById('processBtn');

// Drag and drop
dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('dragover');
});

dropZone.addEventListener('dragleave', () => {
    dropZone.classList.remove('dragover');
});

dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('dragover');
    handleFiles(e.dataTransfer.files);
});

fileInput.addEventListener('change', (e) => {
    handleFiles(e.target.files);
});

function handleFiles(files) {
    for (let file of files) {
        const ext = file.name.split('.').pop().toLowerCase();
        const validExts = ['pdf', 'jpg', 'jpeg', 'png', 'xlsx', 'xls', 'csv'];
        if (validExts.includes(ext)) {
            selectedFiles.push(file);
        } else {
            alert(`File "${file.name}" is not supported. Please use PDF, JPEG, PNG, or Excel files.`);
        }
    }
    renderFileList();
}

function renderFileList() {
    fileList.innerHTML = selectedFiles.map((file, idx) => {
        const icon = getFileIcon(file.name);
        const size = formatFileSize(file.size);
        return `
            <div class="file-item">
                <i class="fas ${icon}"></i>
                <span class="file-name">${file.name}</span>
                <span class="file-size">${size}</span>
                <i class="fas fa-times file-remove" onclick="removeFile(${idx})"></i>
            </div>
        `;
    }).join('');
    
    processBtn.disabled = selectedFiles.length === 0;
}

function removeFile(idx) {
    selectedFiles.splice(idx, 1);
    renderFileList();
}

function getFileIcon(name) {
    const ext = name.split('.').pop().toLowerCase();
    if (ext === 'pdf') return 'fa-file-pdf';
    if (['jpg', 'jpeg', 'png'].includes(ext)) return 'fa-file-image';
    if (['xlsx', 'xls', 'csv'].includes(ext)) return 'fa-file-excel';
    return 'fa-file';
}

function formatFileSize(bytes) {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1048576) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / 1048576).toFixed(1) + ' MB';
}

// ==================== FILE PROCESSING ====================
processBtn.addEventListener('click', processFiles);

async function processFiles() {
    if (selectedFiles.length === 0) return;
    
    const progressContainer = document.getElementById('progressContainer');
    const progressFill = document.getElementById('progressFill');
    const progressText = document.getElementById('progressText');
    const fileType = document.getElementById('fileType').value;
    
    progressContainer.style.display = 'block';
    processBtn.disabled = true;
    
    let processed = 0;
    const total = selectedFiles.length;
    
    for (let file of selectedFiles) {
        progressText.textContent = `Processing ${file.name}...`;
        progressFill.style.width = ((processed / total) * 100) + '%';
        
        try {
            const ext = file.name.split('.').pop().toLowerCase();
            let extractedData = [];
            
            if (['xlsx', 'xls', 'csv'].includes(ext)) {
                extractedData = await processExcel(file);
            } else if (ext === 'pdf') {
                extractedData = await processPDF(file);
            } else if (['jpg', 'jpeg', 'png'].includes(ext)) {
                extractedData = await processImage(file);
            }
            
            // Assign type and unique ID
            extractedData.forEach(row => {
                if (fileType !== 'auto') {
                    row.type = fileType;
                }
                row.uniqueId = generateUniqueId(row.roNumber, row.productUPC);
                row.uploadDate = new Date().toISOString();
                row.fileName = file.name;
            });
            
            masterData.push(...extractedData);
            
            uploadHistory.push({
                fileName: file.name,
                fileType: ext,
                dataType: fileType === 'auto' ? 'auto' : fileType,
                records: extractedData.length,
                date: new Date().toISOString()
            });
            
        } catch (err) {
            console.error(`Error processing ${file.name}:`, err);
            alert(`Error processing ${file.name}: ${err.message}`);
        }
        
        processed++;
    }
    
    progressFill.style.width = '100%';
    progressText.textContent = `Done! Processed ${total} file(s).`;
    
    // Save to localStorage
    saveData();
    
    // Reset
    setTimeout(() => {
        selectedFiles = [];
        renderFileList();
        progressContainer.style.display = 'none';
        processBtn.disabled = true;
        progressFill.style.width = '0%';
    }, 2000);
}

// ==================== EXCEL PROCESSING ====================
async function processExcel(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                const jsonData = XLSX.utils.sheet_to_json(firstSheet, { defval: '' });
                
                const extracted = jsonData.map(row => {
                    return normalizeRow(row);
                });
                
                resolve(extracted);
            } catch (err) {
                reject(err);
            }
        };
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

// ==================== PDF PROCESSING ====================
async function processPDF(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = async (e) => {
            try {
                const typedArray = new Uint8Array(e.target.result);
                pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
                
                const pdf = await pdfjsLib.getDocument(typedArray).promise;
                let fullText = '';
                
                for (let i = 1; i <= pdf.numPages; i++) {
                    const page = await pdf.getPage(i);
                    const textContent = await page.getTextContent();
                    const pageText = textContent.items.map(item => item.str).join(' ');
                    fullText += pageText + '\n';
                }
                
                const extracted = parseTextToData(fullText);
                resolve(extracted);
            } catch (err) {
                reject(err);
            }
        };
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

// ==================== IMAGE PROCESSING (OCR) ====================
async function processImage(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = async (e) => {
            try {
                const result = await Tesseract.recognize(e.target.result, 'eng', {
                    logger: m => {
                        if (m.status === 'recognizing text') {
                            document.getElementById('progressText').textContent = 
                                `OCR Processing: ${Math.round(m.progress * 100)}%`;
                        }
                    }
                });
                
                const extracted = parseTextToData(result.data.text);
                resolve(extracted);
            } catch (err) {
                reject(err);
            }
        };
        reader.onerror = reject;
        reader.readAsDataURL(file);
    });
}

// ==================== TEXT PARSING ====================
function parseTextToData(text) {
    const lines = text.split('\n').filter(l => l.trim());
    const data = [];
    
    // Try to detect table-like patterns
    let type = 'ro'; // default
    if (text.toLowerCase().includes('grn') || text.toLowerCase().includes('goods receipt')) {
        type = 'grn';
    } else if (text.toLowerCase().includes('release order') || text.toLowerCase().includes('r.o')) {
        type = 'ro';
    }
    
    // Look for patterns: numbers that look like UPC codes, quantities, amounts
    const numberPattern = /\b(\d{8,14})\b/g;
    const roPattern = /(?:RO|R\.O\.?|PO|Order)\s*(?:#|No\.?|Number)?\s*:?\s*(\w+[-/]?\w+)/gi;
    
    let roNumber = '';
    const roMatch = roPattern.exec(text);
    if (roMatch) {
        roNumber = roMatch[1];
    }
    
    // Try to extract tabular data from text lines
    for (let line of lines) {
        const parts = line.split(/\s{2,}|\t|,/).map(p => p.trim()).filter(p => p);
        if (parts.length >= 3) {
            // Try to identify if this line has product data
            const hasNumber = parts.some(p => /^\d+(\.\d+)?$/.test(p));
            const hasUPC = parts.some(p => /^\d{8,14}$/.test(p));
            
            if (hasNumber || hasUPC) {
                const row = {
                    type: type,
                    roNumber: roNumber || 'N/A',
                    productUPC: '',
                    productName: '',
                    quantity: 0,
                    rate: 0,
                    amount: 0,
                    date: new Date().toISOString().split('T')[0]
                };
                
                // Try to assign values intelligently
                for (let part of parts) {
                    if (/^\d{8,14}$/.test(part)) {
                        row.productUPC = part;
                    } else if (/^\d+$/.test(part) && !row.quantity) {
                        row.quantity = parseInt(part);
                    } else if (/^\d+\.?\d*$/.test(part) && row.quantity && !row.rate) {
                        row.rate = parseFloat(part);
                    } else if (/^\d+\.?\d*$/.test(part) && row.rate && !row.amount) {
                        row.amount = parseFloat(part);
                    } else if (!/^\d+\.?\d*$/.test(part) && !row.productName) {
                        row.productName = part;
                    }
                }
                
                if (row.productUPC || row.productName) {
                    if (!row.amount && row.quantity && row.rate) {
                        row.amount = row.quantity * row.rate;
                    }
                    data.push(row);
                }
            }
        }
    }
    
    // If no structured data found, create a single entry with raw text
    if (data.length === 0) {
        data.push({
            type: type,
            roNumber: roNumber || 'UNKNOWN',
            productUPC: 'MANUAL_ENTRY',
            productName: text.substring(0, 100),
            quantity: 0,
            rate: 0,
            amount: 0,
            date: new Date().toISOString().split('T')[0]
        });
    }
    
    return data;
}

// ==================== DATA NORMALIZATION ====================
function normalizeRow(row) {
    // Map various possible column names to standard fields
    const fieldMaps = {
        roNumber: ['ro number', 'r.o. number', 'ro_number', 'ronumber', 'po number', 'order number', 'ro no', 'release order', 'ro', 'po', 'order_no', 'order no'],
        productUPC: ['upc', 'product upc', 'upc code', 'barcode', 'ean', 'product_upc', 'item code', 'sku', 'article', 'article no', 'article_no'],
        productName: ['product name', 'productname', 'product_name', 'item name', 'item', 'description', 'product', 'name', 'item_name', 'article name'],
        quantity: ['quantity', 'qty', 'qty.', 'units', 'pieces', 'pcs', 'order qty', 'ordered', 'grn qty', 'received qty', 'received', 'dispatch qty'],
        rate: ['rate', 'price', 'unit price', 'mrp', 'cost', 'unit_price', 'unit rate'],
        amount: ['amount', 'total', 'value', 'total amount', 'net amount', 'total_amount', 'net_value'],
        date: ['date', 'order date', 'grn date', 'receipt date', 'delivery date', 'invoice date'],
        type: ['type', 'doc type', 'document type', 'category']
    };
    
    const normalized = {
        type: 'ro',
        roNumber: '',
        productUPC: '',
        productName: '',
        quantity: 0,
        rate: 0,
        amount: 0,
        date: new Date().toISOString().split('T')[0]
    };
    
    const keys = Object.keys(row);
    
    for (let [field, aliases] of Object.entries(fieldMaps)) {
        for (let key of keys) {
            if (aliases.includes(key.toLowerCase().trim())) {
                if (field === 'quantity' || field === 'rate' || field === 'amount') {
                    normalized[field] = parseFloat(row[key]) || 0;
                } else if (field === 'type') {
                    const val = String(row[key]).toLowerCase();
                    normalized[field] = val.includes('grn') ? 'grn' : 'ro';
                } else {
                    normalized[field] = String(row[key]);
                }
                break;
            }
        }
    }
    
    // If no amount but have qty and rate, calculate
    if (!normalized.amount && normalized.quantity && normalized.rate) {
        normalized.amount = normalized.quantity * normalized.rate;
    }
    
    // If no specific fields found, map by position
    if (!normalized.roNumber && !normalized.productUPC && !normalized.productName) {
        const values = Object.values(row);
        if (values.length >= 5) {
            normalized.roNumber = String(values[0] || '');
            normalized.productUPC = String(values[1] || '');
            normalized.productName = String(values[2] || '');
            normalized.quantity = parseFloat(values[3]) || 0;
            normalized.rate = parseFloat(values[4]) || 0;
            normalized.amount = values[5] ? parseFloat(values[5]) || 0 : normalized.quantity * normalized.rate;
        }
    }
    
    return normalized;
}

// ==================== UNIQUE ID GENERATION ====================
function generateUniqueId(roNumber, productUPC) {
    return `${roNumber || 'NA'}_${productUPC || 'NA'}`;
}

// ==================== DATA PERSISTENCE ====================
function saveData() {
    localStorage.setItem('blinkitGRN_masterData', JSON.stringify(masterData));
    localStorage.setItem('blinkitGRN_uploadHistory', JSON.stringify(uploadHistory));
}

// ==================== MASTER SHEET ====================
function renderMasterSheet() {
    const tbody = document.getElementById('masterTableBody');
    const noDataMsg = document.getElementById('noDataMsg');
    const searchTerm = document.getElementById('searchInput').value.toLowerCase();
    const filterType = document.getElementById('filterType').value;
    
    let filteredData = masterData.filter(row => {
        const matchesSearch = !searchTerm || 
            Object.values(row).some(v => String(v).toLowerCase().includes(searchTerm));
        const matchesFilter = filterType === 'all' || row.type === filterType;
        return matchesSearch && matchesFilter;
    });
    
    if (filteredData.length === 0) {
        tbody.innerHTML = '';
        noDataMsg.style.display = 'block';
        return;
    }
    
    noDataMsg.style.display = 'none';
    tbody.innerHTML = filteredData.map(row => `
        <tr>
            <td><strong>${row.uniqueId || ''}</strong></td>
            <td><span class="badge-${row.type}">${row.type ? row.type.toUpperCase() : ''}</span></td>
            <td>${row.roNumber || ''}</td>
            <td>${row.productUPC || ''}</td>
            <td>${row.productName || ''}</td>
            <td>${row.quantity || 0}</td>
            <td>${row.rate || 0}</td>
            <td>${row.amount || 0}</td>
            <td>${row.date || ''}</td>
            <td>${row.uploadDate ? new Date(row.uploadDate).toLocaleDateString() : ''}</td>
        </tr>
    `).join('');
}

// Search and filter events
document.getElementById('searchInput').addEventListener('input', renderMasterSheet);
document.getElementById('filterType').addEventListener('change', renderMasterSheet);

// ==================== COMPARISON SHEET ====================
function renderComparison() {
    const tbody = document.getElementById('comparisonTableBody');
    const noMsg = document.getElementById('noComparisonMsg');
    
    // Group by unique ID
    const roData = {};
    const grnData = {};
    
    masterData.forEach(row => {
        if (row.type === 'ro') {
            if (!roData[row.uniqueId]) roData[row.uniqueId] = row;
        } else if (row.type === 'grn') {
            if (!grnData[row.uniqueId]) grnData[row.uniqueId] = row;
        }
    });
    
    const allIds = new Set([...Object.keys(roData), ...Object.keys(grnData)]);
    
    if (allIds.size === 0) {
        tbody.innerHTML = '';
        noMsg.style.display = 'block';
        return;
    }
    
    noMsg.style.display = 'none';
    
    const rows = [];
    allIds.forEach(id => {
        const ro = roData[id];
        const grn = grnData[id];
        
        let status = 'missing';
        let statusClass = 'status-missing';
        let statusText = 'Missing';
        
        if (ro && grn) {
            if (ro.quantity === grn.quantity && ro.amount === grn.amount) {
                status = 'match';
                statusClass = 'status-match';
                statusText = 'Matched';
            } else {
                status = 'mismatch';
                statusClass = 'status-mismatch';
                statusText = 'Mismatch';
            }
        } else if (ro && !grn) {
            statusText = 'GRN Missing';
        } else {
            statusText = 'R.O. Missing';
        }
        
        const qtyDiff = (ro ? ro.quantity : 0) - (grn ? grn.quantity : 0);
        
        rows.push(`
            <tr>
                <td><strong>${id}</strong></td>
                <td>${(ro || grn).roNumber || ''}</td>
                <td>${(ro || grn).productUPC || ''}</td>
                <td>${(ro || grn).productName || ''}</td>
                <td>${ro ? ro.quantity : '-'}</td>
                <td>${grn ? grn.quantity : '-'}</td>
                <td class="${qtyDiff !== 0 ? 'status-mismatch' : ''}">${qtyDiff}</td>
                <td>${ro ? ro.amount : '-'}</td>
                <td>${grn ? grn.amount : '-'}</td>
                <td class="${statusClass}">${statusText}</td>
            </tr>
        `);
    });
    
    tbody.innerHTML = rows.join('');
}

// ==================== DASHBOARD ====================
function updateDashboard() {
    // Stats
    document.getElementById('total-entries').textContent = masterData.length;
    document.getElementById('total-ro').textContent = masterData.filter(r => r.type === 'ro').length;
    document.getElementById('total-grn').textContent = masterData.filter(r => r.type === 'grn').length;
    
    // Calculate mismatches
    const roData = {};
    const grnData = {};
    masterData.forEach(row => {
        if (row.type === 'ro') roData[row.uniqueId] = row;
        else if (row.type === 'grn') grnData[row.uniqueId] = row;
    });
    
    let mismatches = 0;
    Object.keys(roData).forEach(id => {
        if (grnData[id] && (roData[id].quantity !== grnData[id].quantity || roData[id].amount !== grnData[id].amount)) {
            mismatches++;
        }
    });
    document.getElementById('total-mismatch').textContent = mismatches;
    
    // Recent uploads
    const recentList = document.getElementById('recent-uploads-list');
    if (uploadHistory.length === 0) {
        recentList.innerHTML = '<p class="empty-state">No uploads yet. Start by uploading files.</p>';
    } else {
        const recent = uploadHistory.slice(-10).reverse();
        recentList.innerHTML = recent.map(h => `
            <div class="activity-item">
                <i class="fas fa-file-upload"></i>
                <div class="activity-info">
                    <h4>${h.fileName}</h4>
                    <p>${h.records} records extracted | Type: ${h.dataType.toUpperCase()}</p>
                </div>
                <span class="activity-time">${new Date(h.date).toLocaleDateString()}</span>
            </div>
        `).join('');
    }
    
    // Charts
    renderCharts();
}

function renderCharts() {
    // Upload Activity Chart
    const uploadCtx = document.getElementById('uploadChart').getContext('2d');
    const last7Days = [];
    const uploadCounts = [];
    
    for (let i = 6; i >= 0; i--) {
        const date = new Date();
        date.setDate(date.getDate() - i);
        const dateStr = date.toISOString().split('T')[0];
        last7Days.push(date.toLocaleDateString('en', { weekday: 'short' }));
        
        const count = uploadHistory.filter(h => 
            h.date.split('T')[0] === dateStr
        ).length;
        uploadCounts.push(count);
    }
    
    // Destroy existing chart if any
    if (window.uploadChartInstance) window.uploadChartInstance.destroy();
    window.uploadChartInstance = new Chart(uploadCtx, {
        type: 'bar',
        data: {
            labels: last7Days,
            datasets: [{
                label: 'Files Uploaded',
                data: uploadCounts,
                backgroundColor: 'rgba(248, 201, 39, 0.7)',
                borderColor: '#f8c927',
                borderWidth: 2,
                borderRadius: 6
            }]
        },
        options: {
            responsive: true,
            plugins: { legend: { display: false } },
            scales: {
                y: { beginAtZero: true, ticks: { stepSize: 1 } }
            }
        }
    });
    
    // Comparison Chart
    const compCtx = document.getElementById('comparisonChart').getContext('2d');
    const roCount = masterData.filter(r => r.type === 'ro').length;
    const grnCount = masterData.filter(r => r.type === 'grn').length;
    
    if (window.compChartInstance) window.compChartInstance.destroy();
    window.compChartInstance = new Chart(compCtx, {
        type: 'doughnut',
        data: {
            labels: ['R.O. Orders', 'GRN Receipts'],
            datasets: [{
                data: [roCount || 0, grnCount || 0],
                backgroundColor: ['#667eea', '#f5576c'],
                borderWidth: 0
            }]
        },
        options: {
            responsive: true,
            plugins: {
                legend: { position: 'bottom' }
            }
        }
    });
}

// ==================== EXPORT FUNCTIONS ====================
function exportAllData() {
    if (masterData.length === 0) {
        alert('No data to export!');
        return;
    }
    exportToExcel(masterData, 'Blinkit_GRN_AllData');
}

function exportMasterSheet() {
    if (masterData.length === 0) {
        alert('No data to export!');
        return;
    }
    exportToExcel(masterData, 'Blinkit_GRN_MasterSheet');
}

function exportToExcel(data, fileName) {
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Data');
    XLSX.writeFile(wb, `${fileName}_${new Date().toISOString().split('T')[0]}.xlsx`);
}

// ==================== INITIALIZATION ====================
document.addEventListener('DOMContentLoaded', () => {
    updateDashboard();
});
