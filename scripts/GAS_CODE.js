// GOOGLE APPS SCRIPT CODE
// 1. Go to https://script.google.com/
// 2. Create a new project attached to your Google Sheet
// 3. Paste this code
// 4. Deploy > New Deployment > Web App
//    - Execute as: Me
//    - Who has access: Anyone
// 5. Copy the Web App URL and put it in your .env file as VITE_GOOGLE_SCRIPT_URL

const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(30000); // Wait up to 30 seconds

  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;

    let result;
    switch (action) {
      case 'login':
        result = handleLogin(data);
        break;
      case 'getReports':
        result = handleGetReports(data);
        break;
      case 'submitReport':
        result = handleSubmitReport(data);
        break;
      case 'getAccountCodes':
        result = handleGetAccountCodes(data);
        break;
      case 'saveAccountCode':
        result = handleSaveAccountCode(data);
        break;
      case 'deleteAccountCode':
        result = handleDeleteAccountCode(data);
        break;
      case 'getCollections':
        result = handleGetCollections(data);
        break;
      case 'saveCollection':
        result = handleSaveCollection(data);
        break;
      case 'saveCollectionBulk':
        result = handleSaveCollectionBulk(data);
        break;
      case 'updateCollection':
        result = handleUpdateCollection(data);
        break;
      case 'deleteCollection':
        result = handleDeleteCollection(data);
        break;
      case 'getSignatories':
        result = handleGetSignatories(data);
        break;
      case 'saveSignatory':
        result = handleSaveSignatory(data);
        break;
      case 'deleteSignatory':
        result = handleDeleteSignatory(data);
        break;
      case 'getRPTCollections':
        result = handleGetRPTCollections(data);
        break;
      case 'saveRPTCollection':
        result = handleSaveRPTCollection(data);
        break;
      case 'deleteRPTCollection':
        result = handleDeleteRPTCollection(data);
        break;
      case 'getSpreadsheet':
        result = handleGetSpreadsheet(data);
        break;
      case 'updateSpreadsheet':
        result = handleUpdateSpreadsheet(data);
        break;
      default:
        result = errorResponse('Invalid action');
    }

    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);

  } catch (e) {
    return ContentService.createTextOutput(JSON.stringify({ 'result': 'error', 'message': e.toString() })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// --- HANDLERS ---

function handleLogin(data) {
  const sheet = getSheet('Users');
  if (!sheet) return errorResponse('Users sheet not found');
  
  const rows = sheet.getDataRange().getValues();
  // Skip header
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] == data.email && rows[i][1] == data.password) {
      return successResponse({
        email: rows[i][0],
        name: rows[i][2],
        role: rows[i][3]
      });
    }
  }
  return errorResponse('Invalid credentials');
}

function handleGetReports(data) {
  const sheet = getSheet('Reports');
  if (!sheet) return successResponse({ reports: [] });

  const rows = sheet.getDataRange().getValues();
  const reports = [];
  
  // Skip header, read from bottom up (newest first), limit to 50
  for (let i = rows.length - 1; i >= 1; i--) {
    if (reports.length >= 50) break;
    try {
      // Column 6 (index 6) is the JSON data
      const jsonStr = rows[i][6];
      if (jsonStr) {
        reports.push(JSON.parse(jsonStr));
      }
    } catch (e) {
      // Ignore parse errors
    }
  }
  return successResponse({ reports });
}

function handleSubmitReport(data) {
  const sheet = getSheet('Reports');
  const report = data.report;
  sheet.appendRow([
    report.date,
    report.reportNumber,
    report.collectorName,
    report.fundType,
    report.totalCollection,
    report.status,
    JSON.stringify(report)
  ]);
  return successResponse({ message: 'Report saved successfully' });
}

function handleGetAccountCodes(data) {
  const sheet = getSheet('Account Codes');
  if (!sheet) return successResponse({ accountCodes: [] });

  const rows = sheet.getDataRange().getValues();
  const codes = [];
  
  for (let i = 1; i < rows.length; i++) {
    if (rows[i].length < 4) continue;
    codes.push({
      id: parseInt(rows[i][0]),
      mainCategory: rows[i][1],
      subCategory: rows[i][2],
      code: rows[i][3]
    });
  }
  return successResponse({ accountCodes: codes });
}

function handleSaveAccountCode(data) {
  const sheet = getSheet('Account Codes');
  const accountCode = data.accountCode;
  
  const rows = sheet.getDataRange().getValues();
  let rowIndex = -1;
  
  for (let i = 1; i < rows.length; i++) {
    if (parseInt(rows[i][0]) === accountCode.id) {
      rowIndex = i + 1;
      break;
    }
  }

  const rowData = [
    accountCode.id,
    accountCode.mainCategory,
    accountCode.subCategory,
    accountCode.code
  ];

  if (rowIndex > -1) {
    sheet.getRange(rowIndex, 1, 1, 4).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }
  return successResponse({});
}

function handleDeleteAccountCode(data) {
  const sheet = getSheet('Account Codes');
  const id = data.id;
  
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (parseInt(rows[i][0]) === id) {
      sheet.deleteRow(i + 1);
      return successResponse({});
    }
  }
  return successResponse({}); // Or error if not found? Server implementation just returned success.
}

function handleGetCollections(data) {
  const sheet = getSheet('Collections');
  if (!sheet) return successResponse({ entries: [] });

  const rows = sheet.getDataRange().getValues();
  const entries = [];
  
  for (let i = 1; i < rows.length; i++) {
    if (rows[i].length < 1) continue;
    entries.push({
      id: parseInt(rows[i][0]),
      afNo: rows[i][1],
      orNo: rows[i][2],
      payor: rows[i][3],
      subCategory: rows[i][4],
      mainCategory: rows[i][5],
      accountCode: rows[i][6],
      amount: parseFloat(rows[i][7] || 0),
      date: rows[i][8],
      remarks: rows[i][9]
    });
  }
  return successResponse({ entries });
}

function handleSaveCollection(data) {
  const sheet = getSheet('Collections');
  const entry = data.entry;
  
  sheet.appendRow([
    entry.id,
    entry.afNo,
    entry.orNo,
    entry.payor,
    entry.subCategory,
    entry.mainCategory,
    entry.accountCode,
    entry.amount,
    entry.date,
    entry.remarks
  ]);
  return successResponse({});
}

function handleSaveCollectionBulk(data) {
  const sheet = getSheet('Collections');
  const header = data.header;
  const charges = data.charges;
  
  // Calculate start ID
  const rows = sheet.getDataRange().getValues();
  let maxId = 0;
  for (let i = 1; i < rows.length; i++) {
    const val = parseInt(rows[i][0]);
    if (!isNaN(val)) maxId = Math.max(maxId, val);
  }
  const startId = maxId + 1;
  
  const newRows = charges.map((c, i) => [
    startId + i,
    header.afNo,
    header.orNo,
    header.payor,
    c.subCategory,
    c.mainCategory,
    c.accountCode,
    c.amount,
    header.date,
    header.remarks
  ]);
  
  if (newRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 10).setValues(newRows);
  }
  
  return successResponse({ startId, count: charges.length });
}

function handleUpdateCollection(data) {
  const sheet = getSheet('Collections');
  const entry = data.entry;
  
  const rows = sheet.getDataRange().getValues();
  let rowIndex = -1;
  
  for (let i = 1; i < rows.length; i++) {
    if (parseInt(rows[i][0]) === entry.id) {
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex > -1) {
    const rowData = [
      entry.id,
      entry.afNo,
      entry.orNo,
      entry.payor,
      entry.subCategory,
      entry.mainCategory,
      entry.accountCode,
      entry.amount,
      entry.date,
      entry.remarks
    ];
    sheet.getRange(rowIndex, 1, 1, 10).setValues([rowData]);
    return successResponse({});
  } else {
    return errorResponse('Entry not found');
  }
}

function handleDeleteCollection(data) {
  const sheet = getSheet('Collections');
  const id = data.id;
  
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (parseInt(rows[i][0]) === id) {
      sheet.deleteRow(i + 1);
      return successResponse({});
    }
  }
  return errorResponse('Entry not found');
}

function handleGetSignatories(data) {
  const sheet = getSheet('Signatories');
  if (!sheet) return successResponse({ signatories: [] });

  const rows = sheet.getDataRange().getValues();
  const signatories = [];
  
  for (let i = 1; i < rows.length; i++) {
    if (rows[i].length < 4) continue;
    signatories.push({
      id: parseInt(rows[i][0]),
      fullName: rows[i][1],
      position: rows[i][2],
      department: rows[i][3],
      remarks: rows[i][4] || ''
    });
  }
  return successResponse({ signatories });
}

function handleSaveSignatory(data) {
  const sheet = getSheet('Signatories');
  const signatory = data.signatory;
  
  const rows = sheet.getDataRange().getValues();
  let rowIndex = -1;
  
  if (signatory.id) {
    for (let i = 1; i < rows.length; i++) {
      if (parseInt(rows[i][0]) === signatory.id) {
        rowIndex = i + 1;
        break;
      }
    }
  }

  const rowData = [
    signatory.id,
    signatory.fullName,
    signatory.position,
    signatory.department,
    signatory.remarks || ''
  ];

  if (rowIndex > -1) {
    sheet.getRange(rowIndex, 1, 1, 5).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }
  return successResponse({});
}

function handleDeleteSignatory(data) {
  const sheet = getSheet('Signatories');
  const id = data.id;
  
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (parseInt(rows[i][0]) === id) {
      sheet.deleteRow(i + 1);
      return successResponse({});
    }
  }
  return successResponse({});
}

function handleGetRPTCollections(data) {
  const sheet = getSheet('RPT Collections');
  if (!sheet) return successResponse({ collections: [] });

  const rows = sheet.getDataRange().getValues();
  const collections = [];
  
  for (let i = 1; i < rows.length; i++) {
    if (rows[i].length < 1) continue;
    
    // Parse amount logic same as server
    let amount = 0;
    try {
      const cleaned = String(rows[i][8]).replace(/[^0-9.-]/g, '');
      amount = parseFloat(cleaned) || 0;
    } catch (e) {}

    collections.push({
      id: parseInt(rows[i][0]),
      af56Id: rows[i][1],
      orNumber: rows[i][2],
      payor: rows[i][3],
      barangay: rows[i][4],
      landName: rows[i][5],
      tdNumber: rows[i][6],
      yearsPaid: rows[i][7],
      amount: amount,
      date: rows[i][9],
      remarks: rows[i][10] || ''
    });
  }
  return successResponse({ collections });
}

function handleSaveRPTCollection(data) {
  const sheet = getSheet('RPT Collections');
  const collection = data.collection;
  
  const rows = sheet.getDataRange().getValues();
  let rowIndex = -1;
  
  if (collection.id) {
    for (let i = 1; i < rows.length; i++) {
      if (parseInt(rows[i][0]) === collection.id) {
        rowIndex = i + 1;
        break;
      }
    }
  }

  const rowData = [
    collection.id,
    collection.af56Id,
    collection.orNumber,
    collection.payor,
    collection.barangay,
    collection.landName,
    collection.tdNumber,
    collection.yearsPaid,
    collection.amount,
    collection.date,
    collection.remarks || ''
  ];

  if (rowIndex > -1) {
    sheet.getRange(rowIndex, 1, 1, 11).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }
  return successResponse({});
}

function handleDeleteRPTCollection(data) {
  const sheet = getSheet('RPT Collections');
  const id = data.id;
  
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (parseInt(rows[i][0]) === id) {
      sheet.deleteRow(i + 1);
      return successResponse({});
    }
  }
  return successResponse({});
}

// --- HELPERS ---

function getAppSpreadsheet() {
  const customId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (customId) {
    try {
      return SpreadsheetApp.openById(customId);
    } catch (e) {
      Logger.log('Failed to open spreadsheet by ID: ' + customId);
    }
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet(name) {
  const ss = getAppSpreadsheet();
  let lookupName = name;
  if (name === 'Signatories') {
    const tempSheet = ss.getSheetByName('Signatores');
    if (tempSheet) {
      lookupName = 'Signatores';
    }
  }
  let sheet = ss.getSheetByName(lookupName);
  if (!sheet) {
    sheet = ss.insertSheet(lookupName);
    // Initialize headers if new sheet
    if (lookupName === 'Users') {
      sheet.appendRow(['Email', 'Password', 'Name', 'Role']);
      sheet.appendRow(['admin@lgu.gov.ph', 'admin', 'Admin User', 'admin']);
    } else if (lookupName === 'Reports') {
      sheet.appendRow(['Date', 'Report Number', 'Collector', 'Fund Type', 'Total Collection', 'Status', 'JSON Data']);
    } else if (lookupName === 'Account Codes') {
      sheet.appendRow(['ID', 'Main Category', 'Sub Category', 'Code']);
    } else if (lookupName === 'Collections') {
      sheet.appendRow(['ID', 'AF No.', 'OR No.', 'Payor', 'Sub Category', 'Main Category', 'Account Code', 'Amount', 'Date', 'Remarks']);
    } else if (lookupName === 'Signatories' || lookupName === 'Signatores') {
      sheet.appendRow(['ID', 'Full Name', 'Position', 'Department', 'Remarks']);
    } else if (lookupName === 'RPT Collections') {
      sheet.appendRow(['ID', 'AF56 ID', 'OR Number', 'Payor', 'Barangay', 'Land Name', 'TD #', 'Years Paid', 'Amount', 'Date', 'Remarks']);
    }
  }
  return sheet;
}

function successResponse(data) {
  return { result: 'success', ...data };
}

function errorResponse(message) {
  return { result: 'error', message: message };
}

function handleGetSpreadsheet() {
  const customId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  const activeId = customId || SPREADSHEET_ID;
  return successResponse({
    spreadsheetId: activeId,
    spreadsheetUrl: 'https://docs.google.com/spreadsheets/d/' + activeId + '/edit'
  });
}

function handleUpdateSpreadsheet(data) {
  const urlOrId = data.spreadsheetUrl || data.spreadsheetId || data.url || data.id;
  if (!urlOrId) return errorResponse('Missing spreadsheet URL or ID');
  
  let extractedId = urlOrId.trim();
  const match = extractedId.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (match && match[1]) {
    extractedId = match[1];
  }
  
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', extractedId);
  return successResponse({
    message: 'Spreadsheet updated successfully',
    spreadsheetId: extractedId,
    spreadsheetUrl: 'https://docs.google.com/spreadsheets/d/' + extractedId + '/edit'
  });
}

