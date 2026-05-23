// ==================== GOOGLE APPS SCRIPT - Code.gs ====================

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Blinkit GRN Sheet Manager')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Get or create the Master Sheet
function getMasterSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('MasterData');
  if (!sheet) {
    sheet = ss.insertSheet('MasterData');
    sheet.appendRow(['Unique ID', 'Type', 'R.O. Number', 'Product UPC', 'Product Name', 'Quantity', 'Rate', 'Amount', 'Date', 'Upload Date', 'File Name']);
    sheet.getRange(1, 1, 1, 11).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Get or create Upload History sheet
function getHistorySheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('UploadHistory');
  if (!sheet) {
    sheet = ss.insertSheet('UploadHistory');
    sheet.appendRow(['File Name', 'File Type', 'Data Type', 'Records', 'Upload Date']);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Save extracted data to Master Sheet
function saveDataToSheet(dataArray) {
  var sheet = getMasterSheet();
  var rows = [];
  
  for (var i = 0; i < dataArray.length; i++) {
    var row = dataArray[i];
    var uniqueId = (row.roNumber || 'NA') + '_' + (row.productUPC || 'NA');
    rows.push([
      uniqueId,
      row.type || 'ro',
      row.roNumber || '',
      row.productUPC || '',
      row.productName || '',
      row.quantity || 0,
      row.rate || 0,
      row.amount || 0,
      row.date || new Date().toISOString().split('T')[0],
      new Date().toISOString(),
      row.fileName || ''
    ]);
  }
  
  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 11).setValues(rows);
  }
  
  return rows.length;
}

// Save upload history
function saveUploadHistory(fileName, fileType, dataType, records) {
  var sheet = getHistorySheet();
  sheet.appendRow([fileName, fileType, dataType, records, new Date().toISOString()]);
}

// Get all master data
function getAllMasterData() {
  var sheet = getMasterSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  
  var data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
  var result = [];
  
  for (var i = 0; i < data.length; i++) {
    result.push({
      uniqueId: data[i][0],
      type: data[i][1],
      roNumber: data[i][2],
      productUPC: data[i][3],
      productName: data[i][4],
      quantity: data[i][5],
      rate: data[i][6],
      amount: data[i][7],
      date: data[i][8],
      uploadDate: data[i][9],
      fileName: data[i][10]
    });
  }
  
  return result;
}

// Get upload history
function getUploadHistory() {
  var sheet = getHistorySheet();
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  
  var data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  var result = [];
  
  for (var i = 0; i < data.length; i++) {
    result.push({
      fileName: data[i][0],
      fileType: data[i][1],
      dataType: data[i][2],
      records: data[i][3],
      date: data[i][4]
    });
  }
  
  return result;
}

// Process uploaded file (receives base64 data)
function processUploadedFile(fileData, fileName, fileType, docType) {
  try {
    var extractedData = [];
    
    if (fileType === 'xlsx' || fileType === 'xls' || fileType === 'csv') {
      extractedData = processExcelFile(fileData, fileName, docType);
    } else if (fileType === 'pdf') {
      extractedData = processPDFFile(fileData, fileName, docType);
    } else if (fileType === 'jpg' || fileType === 'jpeg' || fileType === 'png') {
      extractedData = processImageFile(fileData, fileName, docType);
    }
    
    if (extractedData.length > 0) {
      var savedCount = saveDataToSheet(extractedData);
      saveUploadHistory(fileName, fileType, docType, savedCount);
      return { success: true, records: savedCount };
    }
    
    return { success: false, message: 'No data extracted from file.' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// Process Excel/CSV files
function processExcelFile(base64Data, fileName, docType) {
  var decoded = Utilities.base64Decode(base64Data);
  var blob = Utilities.newBlob(decoded, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', fileName);
  
  // For CSV
  if (fileName.toLowerCase().endsWith('.csv')) {
    blob = Utilities.newBlob(decoded, 'text/csv', fileName);
    var csvContent = blob.getDataAsString();
    return parseCSVContent(csvContent, fileName, docType);
  }
  
  // For Excel - upload to Drive temporarily, open as Spreadsheet
  var tempFile = DriveApp.createFile(blob);
  try {
    var tempSS = SpreadsheetApp.open(tempFile);
    var tempSheet = tempSS.getSheets()[0];
    var data = tempSheet.getDataRange().getValues();
    var result = parseSheetData(data, fileName, docType);
    DriveApp.getFileById(tempFile.getId()).setTrashed(true);
    return result;
  } catch (e) {
    DriveApp.getFileById(tempFile.getId()).setTrashed(true);
    throw e;
  }
}

// Parse sheet data (from Excel)
function parseSheetData(data, fileName, docType) {
  if (data.length < 2) return [];
  
  var headers = data[0].map(function(h) { return String(h).toLowerCase().trim(); });
  var result = [];
  
  var fieldMap = {
    roNumber: ['ro number', 'r.o. number', 'ro_number', 'ronumber', 'po number', 'order number', 'ro no', 'release order', 'ro', 'po', 'order_no', 'order no'],
    productUPC: ['upc', 'product upc', 'upc code', 'barcode', 'ean', 'product_upc', 'item code', 'sku', 'article', 'article no', 'article_no'],
    productName: ['product name', 'productname', 'product_name', 'item name', 'item', 'description', 'product', 'name', 'item_name', 'article name'],
    quantity: ['quantity', 'qty', 'qty.', 'units', 'pieces', 'pcs', 'order qty', 'ordered', 'grn qty', 'received qty', 'received', 'dispatch qty'],
    rate: ['rate', 'price', 'unit price', 'mrp', 'cost', 'unit_price', 'unit rate'],
    amount: ['amount', 'total', 'value', 'total amount', 'net amount', 'total_amount', 'net_value'],
    date: ['date', 'order date', 'grn date', 'receipt date', 'delivery date', 'invoice date']
  };
  
  // Map header indices
  var indices = {};
  for (var field in fieldMap) {
    for (var h = 0; h < headers.length; h++) {
      if (fieldMap[field].indexOf(headers[h]) !== -1) {
        indices[field] = h;
        break;
      }
    }
  }
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row.every(function(cell) { return !cell && cell !== 0; })) continue; // skip empty rows
    
    var entry = {
      type: docType === 'auto' ? 'ro' : docType,
      roNumber: indices.roNumber !== undefined ? String(row[indices.roNumber]) : '',
      productUPC: indices.productUPC !== undefined ? String(row[indices.productUPC]) : '',
      productName: indices.productName !== undefined ? String(row[indices.productName]) : '',
      quantity: indices.quantity !== undefined ? parseFloat(row[indices.quantity]) || 0 : 0,
      rate: indices.rate !== undefined ? parseFloat(row[indices.rate]) || 0 : 0,
      amount: indices.amount !== undefined ? parseFloat(row[indices.amount]) || 0 : 0,
      date: indices.date !== undefined ? String(row[indices.date]) : new Date().toISOString().split('T')[0],
      fileName: fileName
    };
    
    // Fallback: if no fields mapped, use positional
    if (!entry.roNumber && !entry.productUPC && !entry.productName) {
      if (row.length >= 5) {
        entry.roNumber = String(row[0] || '');
        entry.productUPC = String(row[1] || '');
        entry.productName = String(row[2] || '');
        entry.quantity = parseFloat(row[3]) || 0;
        entry.rate = parseFloat(row[4]) || 0;
        entry.amount = row[5] ? parseFloat(row[5]) || 0 : entry.quantity * entry.rate;
      }
    }
    
    if (!entry.amount && entry.quantity && entry.rate) {
      entry.amount = entry.quantity * entry.rate;
    }
    
    result.push(entry);
  }
  
  return result;
}

// Parse CSV content
function parseCSVContent(csvContent, fileName, docType) {
  var rows = Utilities.parseCsv(csvContent);
  return parseSheetData(rows, fileName, docType);
}

// Process PDF file (extract text using Drive)
function processPDFFile(base64Data, fileName, docType) {
  var decoded = Utilities.base64Decode(base64Data);
  var blob = Utilities.newBlob(decoded, 'application/pdf', fileName);
  
  // Use OCR via Google Drive to extract text
  var resource = {
    title: fileName,
    mimeType: 'application/pdf'
  };
  
  var file = DriveApp.createFile(blob);
  var fileId = file.getId();
  
  try {
    // Convert PDF to Google Doc for text extraction
    var docFile = Drive.Files.copy(
      { title: fileName + '_ocr', mimeType: 'application/vnd.google-apps.document' },
      fileId,
      { ocr: true }
    );
    
    var doc = DocumentApp.openById(docFile.id);
    var text = doc.getBody().getText();
    
    // Cleanup temp files
    DriveApp.getFileById(fileId).setTrashed(true);
    DriveApp.getFileById(docFile.id).setTrashed(true);
    
    return parseTextContent(text, fileName, docType);
  } catch (e) {
    DriveApp.getFileById(fileId).setTrashed(true);
    // Fallback: try simple text extraction
    return [{ type: docType === 'auto' ? 'ro' : docType, roNumber: 'PDF_IMPORT', productUPC: 'MANUAL', productName: fileName, quantity: 0, rate: 0, amount: 0, date: new Date().toISOString().split('T')[0], fileName: fileName }];
  }
}

// Process Image file (OCR via Google Drive)
function processImageFile(base64Data, fileName, docType) {
  var decoded = Utilities.base64Decode(base64Data);
  var mimeType = 'image/jpeg';
  if (fileName.toLowerCase().endsWith('.png')) mimeType = 'image/png';
  
  var blob = Utilities.newBlob(decoded, mimeType, fileName);
  var file = DriveApp.createFile(blob);
  var fileId = file.getId();
  
  try {
    // Use Drive API OCR
    var docFile = Drive.Files.copy(
      { title: fileName + '_ocr', mimeType: 'application/vnd.google-apps.document' },
      fileId,
      { ocr: true }
    );
    
    var doc = DocumentApp.openById(docFile.id);
    var text = doc.getBody().getText();
    
    DriveApp.getFileById(fileId).setTrashed(true);
    DriveApp.getFileById(docFile.id).setTrashed(true);
    
    return parseTextContent(text, fileName, docType);
  } catch (e) {
    DriveApp.getFileById(fileId).setTrashed(true);
    return [{ type: docType === 'auto' ? 'ro' : docType, roNumber: 'IMG_IMPORT', productUPC: 'MANUAL', productName: fileName, quantity: 0, rate: 0, amount: 0, date: new Date().toISOString().split('T')[0], fileName: fileName }];
  }
}

// Parse text content from PDF/Image OCR
function parseTextContent(text, fileName, docType) {
  var lines = text.split('\n').filter(function(l) { return l.trim(); });
  var data = [];
  
  var type = docType;
  if (docType === 'auto') {
    if (text.toLowerCase().indexOf('grn') !== -1 || text.toLowerCase().indexOf('goods receipt') !== -1) {
      type = 'grn';
    } else {
      type = 'ro';
    }
  }
  
  // Try to find RO number
  var roNumber = '';
  var roMatch = text.match(/(?:RO|R\.O\.?|PO|Order)\s*(?:#|No\.?|Number)?\s*:?\s*(\w+[-\/]?\w+)/i);
  if (roMatch) roNumber = roMatch[1];
  
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i];
    var parts = line.split(/\s{2,}|\t|,/).map(function(p) { return p.trim(); }).filter(function(p) { return p; });
    
    if (parts.length >= 3) {
      var hasNumber = parts.some(function(p) { return /^\d+(\.\d+)?$/.test(p); });
      var hasUPC = parts.some(function(p) { return /^\d{8,14}$/.test(p); });
      
      if (hasNumber || hasUPC) {
        var entry = {
          type: type,
          roNumber: roNumber || 'N/A',
          productUPC: '',
          productName: '',
          quantity: 0,
          rate: 0,
          amount: 0,
          date: new Date().toISOString().split('T')[0],
          fileName: fileName
        };
        
        for (var j = 0; j < parts.length; j++) {
          var part = parts[j];
          if (/^\d{8,14}$/.test(part)) {
            entry.productUPC = part;
          } else if (/^\d+$/.test(part) && !entry.quantity) {
            entry.quantity = parseInt(part);
          } else if (/^\d+\.?\d*$/.test(part) && entry.quantity && !entry.rate) {
            entry.rate = parseFloat(part);
          } else if (/^\d+\.?\d*$/.test(part) && entry.rate && !entry.amount) {
            entry.amount = parseFloat(part);
          } else if (!/^\d+\.?\d*$/.test(part) && !entry.productName) {
            entry.productName = part;
          }
        }
        
        if (entry.productUPC || entry.productName) {
          if (!entry.amount && entry.quantity && entry.rate) {
            entry.amount = entry.quantity * entry.rate;
          }
          data.push(entry);
        }
      }
    }
  }
  
  if (data.length === 0) {
    data.push({
      type: type,
      roNumber: roNumber || 'UNKNOWN',
      productUPC: 'MANUAL_ENTRY',
      productName: text.substring(0, 100),
      quantity: 0,
      rate: 0,
      amount: 0,
      date: new Date().toISOString().split('T')[0],
      fileName: fileName
    });
  }
  
  return data;
}

// Get dashboard stats
function getDashboardStats() {
  var masterData = getAllMasterData();
  var history = getUploadHistory();
  
  var totalEntries = masterData.length;
  var totalRO = masterData.filter(function(r) { return r.type === 'ro'; }).length;
  var totalGRN = masterData.filter(function(r) { return r.type === 'grn'; }).length;
  
  // Calculate mismatches
  var roData = {};
  var grnData = {};
  masterData.forEach(function(row) {
    if (row.type === 'ro') roData[row.uniqueId] = row;
    else if (row.type === 'grn') grnData[row.uniqueId] = row;
  });
  
  var mismatches = 0;
  Object.keys(roData).forEach(function(id) {
    if (grnData[id] && (roData[id].quantity !== grnData[id].quantity || roData[id].amount !== grnData[id].amount)) {
      mismatches++;
    }
  });
  
  return {
    totalEntries: totalEntries,
    totalRO: totalRO,
    totalGRN: totalGRN,
    mismatches: mismatches,
    recentUploads: history.slice(-10).reverse()
  };
}
