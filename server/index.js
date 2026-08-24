const express = require('express');
const cors = require('cors');
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');
require('dotenv').config();
const { ensureSpreadsheet } = require('./setup');

const app = express();
const PORT = process.env.PORT || 5000;

// Helper to extract Google Sheet ID from URL or raw ID
function extractSpreadsheetId(urlOrId) {
  if (!urlOrId || typeof urlOrId !== 'string') return null;
  const trimmed = urlOrId.trim();
  if (/^[a-zA-Z0-9-_]{25,100}$/.test(trimmed)) {
    return trimmed;
  }
  const match = trimmed.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (match && match[1]) {
    return match[1];
  }
  return null;
}

// Helper to write SPREADSHEET_ID to .env persistently
function updateEnvSpreadsheetId(newId) {
  try {
    const envPath = path.join(__dirname, '.env');
    let content = '';
    if (fs.existsSync(envPath)) {
      content = fs.readFileSync(envPath, 'utf8');
    }
    if (content.includes('SPREADSHEET_ID=')) {
      content = content.replace(/SPREADSHEET_ID=.*/g, `SPREADSHEET_ID=${newId}`);
    } else {
      content += `\nSPREADSHEET_ID=${newId}`;
    }
    fs.writeFileSync(envPath, content.trim() + '\n', 'utf8');
  } catch (err) {
    console.error('Failed to write to .env:', err);
  }
}

// Middleware
app.use(cors());
app.use(express.json({ type: ['application/json', 'text/plain'] }));
app.use((req, res, next) => {
  if (typeof req.body === 'string') {
    try {
      req.body = JSON.parse(req.body);
    } catch (e) {
      // ignore
    }
  }
  next();
});
app.use((req, res, next) => {
  console.log(`[${new Date().toISOString()}] ${req.method} ${req.url}`);
  next();
});

// Google Sheets Auth
// Expects 'credentials.json' in the root of the server directory
// OR GOOGLE_SERVICE_ACCOUNT_JSON env var
const SCOPES = ['https://www.googleapis.com/auth/spreadsheets'];

const auth = new google.auth.GoogleAuth({
  keyFile: process.env.GOOGLE_APPLICATION_CREDENTIALS || 'credentials.json',
  scopes: SCOPES,
});

const sheets = google.sheets({ version: 'v4', auth });
let SPREADSHEET_ID = process.env.SPREADSHEET_ID;

// Initialize
(async () => {
  try {
    console.log('Initializing server...');
    SPREADSHEET_ID = await ensureSpreadsheet(SPREADSHEET_ID);
    console.log(`Server initialized with Spreadsheet ID: ${SPREADSHEET_ID}`);
  } catch (error) {
    console.error('Failed to initialize spreadsheet:', error);
  }
})();

// Helper to get sheet data
async function getSheetData(range) {
  if (!SPREADSHEET_ID) throw new Error('Spreadsheet ID not initialized');
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range,
  });
  return response.data.values || [];
}

// Helper to append data
async function appendSheetData(range, values) {
  if (!SPREADSHEET_ID) throw new Error('Spreadsheet ID not initialized');
  const response = await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range,
    valueInputOption: 'USER_ENTERED',
    requestBody: {
      values: [values],
    },
  });
  return response.data.values;
}

// Routes

// 1. LOGIN
app.post('/api/login', async (req, res) => {
  try {
    const { email, password } = req.body;
    const rows = await getSheetData('Users!A:D'); // Email, Password, Name, Role
    
    // Skip header (index 0)
    const user = rows.slice(1).find(row => row[0] === email && row[1] === password);
    
    if (user) {
      res.json({
        result: 'success',
        email: user[0],
        name: user[2],
        role: user[3]
      });
    } else {
      res.status(401).json({ result: 'error', message: 'Invalid credentials' });
    }
  } catch (error) {
    console.error('Login error:', error);
    res.status(500).json({ result: 'error', message: error.message });
  }
});

// 2. GET REPORTS
app.get('/api/reports', async (req, res) => {
  try {
    const rows = await getSheetData('Reports!A:G'); // JSON is in G (index 6)
    const reports = [];
    
    // Read from bottom up, skip header
    // rows[0] is header
    for (let i = rows.length - 1; i >= 1; i--) {
      if (reports.length >= 50) break;
      try {
        const jsonStr = rows[i][6];
        if (jsonStr) {
          reports.push(JSON.parse(jsonStr));
        }
      } catch (e) {
        console.warn('Failed to parse report row:', i);
      }
    }
    
    res.json({ result: 'success', reports });
  } catch (error) {
    console.error('Get Reports error:', error);
    res.status(500).json({ result: 'error', message: error.message });
  }
});

// 3. SUBMIT REPORT
app.post('/api/reports', async (req, res) => {
  try {
    const { report } = req.body;
    if (!report) return res.status(400).json({ message: 'Missing report data' });

    const row = [
      report.date,
      report.reportNumber,
      report.collectorName,
      report.fundType,
      report.totalCollection,
      report.status,
      JSON.stringify(report)
    ];

    await appendSheetData('Reports!A:G', row);
    res.json({ result: 'success', message: 'Report saved' });
  } catch (error) {
    console.error('Submit Report error:', error);
    res.status(500).json({ result: 'error', message: error.message });
  }
});

// 4. GET ACCOUNT CODES
app.get('/api/account-codes', async (req, res) => {
  try {
    const rows = await getSheetData("'Account Codes'!A:D"); // ID, Main, Sub, Code
    const codes = [];
    
    // Skip header
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (row.length < 4) continue;
      codes.push({
        id: parseInt(row[0]),
        mainCategory: row[1],
        subCategory: row[2],
        code: row[3]
      });
    }
    
    res.json({ result: 'success', accountCodes: codes });
  } catch (error) {
    // If sheet doesn't exist, return empty
    if (error.message.includes('Unable to parse range')) {
      return res.json({ result: 'success', accountCodes: [] });
    }
    console.error('Get Account Codes error:', error);
    res.status(500).json({ result: 'error', message: error.message });
  }
});

// 5. SAVE ACCOUNT CODE (Upsert)
app.post('/api/account-codes', async (req, res) => {
  try {
    const { accountCode } = req.body;
    if (!accountCode) return res.status(400).json({ message: 'Missing data' });

    // We need to find if it exists to update it, or append if new.
    // However, Sheets API "update" requires knowing the specific range (Row number).
    // This is inefficient with simple "values.get". 
    // Optimization: Read all IDs, find index, update specific cell or append.
    
    const rows = await getSheetData("'Account Codes'!A:A"); // Get just IDs
    let rowIndex = -1;
    
    for (let i = 1; i < rows.length; i++) {
      if (parseInt(rows[i][0]) === accountCode.id) {
        rowIndex = i + 1; // 1-based index, +1 for header if we started at 1? No, rows includes header.
        // If rows includes header at 0. rows[1] is row 2.
        // rowIndex for API should be 2.
        rowIndex = i + 1;
        break;
      }
    }

    const values = [
      accountCode.id,
      accountCode.mainCategory,
      accountCode.subCategory,
      accountCode.code
    ];

    if (rowIndex > -1) {
      // Update existing
      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: `'Account Codes'!A${rowIndex}:D${rowIndex}`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values: [values] }
      });
    } else {
      // Append new
      await appendSheetData("'Account Codes'!A:D", values);
    }

    res.json({ result: 'success' });
  } catch (error) {
    console.error('Save Account Code error:', error);
    res.status(500).json({ result: 'error', message: error.message });
  }
});

app.post('/api/account-codes/delete', async (req, res) => {
  try {
    const { id } = req.body;
    
    const rows = await getSheetData("'Account Codes'!A:A");
    let rowIndex = -1;
    
    for (let i = 1; i < rows.length; i++) {
      if (parseInt(rows[i][0]) === id) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex > -1) {
      // Clear the row content
      await sheets.spreadsheets.values.clear({
        spreadsheetId: SPREADSHEET_ID,
        range: `'Account Codes'!A${rowIndex}:D${rowIndex}`,
      });
      // Note: This leaves a blank row. Proper deletion requires batchUpdate 'deleteDimension'.
      // For simplicity in this demo, clearing is often acceptable or we implement proper delete.
      // Let's implement proper delete to keep sheet clean.
      
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        requestBody: {
          requests: [
            {
              deleteDimension: {
                range: {
                  sheetId: await getSheetId('Account Codes'),
                  dimension: 'ROWS',
                  startIndex: rowIndex - 1,
                  endIndex: rowIndex
                }
              }
            }
          ]
        }
      });
    }

    res.json({ result: 'success' });
  } catch (error) {
    console.error('Delete Account Code error:', error);
    res.status(500).json({ result: 'error', message: error.message });
  }
});

// Helper to get Sheet ID by title
async function getSheetId(title) {
  const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  const sheet = meta.data.sheets.find(s => s.properties.title === title);
  return sheet ? sheet.properties.sheetId : 0;
}

// 6. COLLECTIONS
app.get('/api/collections', async (req, res) => {
  try {
    // Columns: ID, AF No., OR No., Payor, Sub Category, Main Category, Account Code, Amount, Date, Remarks
    const rows = await getSheetData("Collections!A:J");
    const entries = [];
    
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (row.length < 1) continue;
      entries.push({
        id: parseInt(row[0]),
        afNo: row[1],
        orNo: row[2],
        payor: row[3],
        subCategory: row[4],
        mainCategory: row[5],
        accountCode: row[6],
        amount: parseFloat(row[7] || 0),
        date: row[8],
        remarks: row[9]
      });
    }
    
    res.json({ result: 'success', entries });
  } catch (error) {
    console.error('Get Collections error:', error);
    res.status(500).json({ result: 'error', message: error.message });
  }
});

app.post('/api/collections', async (req, res) => {
  try {
    const { entry } = req.body;
    if (!entry) return res.status(400).json({ message: 'Missing entry data' });

    const row = [
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

    await appendSheetData("Collections!A:J", row);
    res.json({ result: 'success' });
  } catch (error) {
    console.error('Save Collection error:', error);
    res.status(500).json({ result: 'error', message: error.message });
  }
});

async function getNextCollectionId() {
  const rows = await getSheetData("Collections!A:A");
  let maxId = 0;
  for (let i = 1; i < rows.length; i++) {
    const val = parseInt(rows[i][0]);
    if (!isNaN(val)) {
      maxId = Math.max(maxId, val);
    }
  }
  return maxId + 1;
}

app.post('/api/collections/bulk', async (req, res) => {
  try {
    const { header, charges } = req.body;
    if (!header || !Array.isArray(charges) || charges.length === 0) {
      return res.status(400).json({ message: 'Missing header or charges' });
    }
    const startId = await getNextCollectionId();
    for (let i = 0; i < charges.length; i++) {
      const c = charges[i];
      const row = [
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
      ];
      await appendSheetData("Collections!A:J", row);
    }
    res.json({ result: 'success', startId, count: charges.length });
  } catch (error) {
    console.error('Save Collections Bulk error:', error);
    res.status(500).json({ result: 'error', message: error.message });
  }
});

// 7. UPDATE COLLECTION ENTRY
app.post('/api/collections/update', async (req, res) => {
  try {
    const { entry } = req.body;
    if (!entry || !entry.id) return res.status(400).json({ message: 'Missing entry data or ID' });

    const rows = await getSheetData("Collections!A:A");
    let rowIndex = -1;
    
    for (let i = 1; i < rows.length; i++) {
      if (parseInt(rows[i][0]) === entry.id) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex > -1) {
      const row = [
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

      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: `Collections!A${rowIndex}:J${rowIndex}`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values: [row] }
      });
      res.json({ result: 'success' });
    } else {
      res.status(404).json({ result: 'error', message: 'Entry not found' });
    }
  } catch (error) {
    console.error('Update Collection error:', error);
    res.status(500).json({ result: 'error', message: error.message });
  }
});

// 8. DELETE COLLECTION ENTRY
app.post('/api/collections/delete', async (req, res) => {
  try {
    const { id } = req.body;
    
    const rows = await getSheetData("Collections!A:A");
    let rowIndex = -1;
    
    for (let i = 1; i < rows.length; i++) {
      if (parseInt(rows[i][0]) === id) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex > -1) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        requestBody: {
          requests: [
            {
              deleteDimension: {
                range: {
                  sheetId: await getSheetId('Collections'),
                  dimension: 'ROWS',
                  startIndex: rowIndex - 1,
                  endIndex: rowIndex
                }
              }
            }
          ]
        }
      });
      res.json({ result: 'success' });
    } else {
      res.status(404).json({ result: 'error', message: 'Entry not found' });
    }
  } catch (error) {
    console.error('Delete Collection error:', error);
    res.status(500).json({ result: 'error', message: error.message });
  }
});

// 9. SIGNATORIES
app.get('/api/signatories', async (req, res) => {
  try {
    const rows = await getSheetData("Signatories!A:E"); // ID, Full Name, Position, Department, Remarks
    const signatories = [];
    
    // Skip header
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (row.length < 4) continue;
      signatories.push({
        id: parseInt(row[0]),
        fullName: row[1],
        position: row[2],
        department: row[3],
        remarks: row[4] || ''
      });
    }
    
    res.json({ result: 'success', signatories });
  } catch (error) {
    if (error.message.includes('Unable to parse range')) {
      return res.json({ result: 'success', signatories: [] });
    }
    console.error('Get Signatories error:', error);
    res.status(500).json({ result: 'error', message: error.message });
  }
});

app.post('/api/signatories', async (req, res) => {
  try {
    const { signatory } = req.body;
    if (!signatory) return res.status(400).json({ message: 'Missing data' });

    const rows = await getSheetData("Signatories!A:A");
    let rowIndex = -1;
    
    // Check for update
    if (signatory.id) {
        for (let i = 1; i < rows.length; i++) {
            if (parseInt(rows[i][0]) === signatory.id) {
                rowIndex = i + 1;
                break;
            }
        }
    }

    const values = [
      signatory.id,
      signatory.fullName,
      signatory.position,
      signatory.department,
      signatory.remarks || ''
    ];

    if (rowIndex > -1) {
      // Update
      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: `Signatories!A${rowIndex}:E${rowIndex}`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values: [values] }
      });
    } else {
      // Append
      await appendSheetData("Signatories!A:E", values);
    }

    res.json({ result: 'success' });
  } catch (error) {
    console.error('Save Signatory error:', error);
    res.status(500).json({ result: 'error', message: error.message });
  }
});

app.post('/api/signatories/delete', async (req, res) => {
  try {
    const { id } = req.body;
    
    const rows = await getSheetData("Signatories!A:A");
    let rowIndex = -1;
    
    for (let i = 1; i < rows.length; i++) {
      if (parseInt(rows[i][0]) === id) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex > -1) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        requestBody: {
          requests: [
            {
              deleteDimension: {
                range: {
                  sheetId: await getSheetId('Signatories'),
                  dimension: 'ROWS',
                  startIndex: rowIndex - 1,
                  endIndex: rowIndex
                }
              }
            }
          ]
        }
      });
    }

    res.json({ result: 'success' });
  } catch (error) {
    console.error('Delete Signatory error:', error);
    res.status(500).json({ result: 'error', message: error.message });
  }
});

// 10. RPT COLLECTIONS
app.get('/api/rpt-collections', async (req, res) => {
  try {
    // Columns: ID, AF56 ID, OR Number, Payor, Barangay, Land Name, TD #, Years paid, Amount, Date, Remarks
    const rows = await getSheetData("'RPT Collections'!A:K");
    const collections = [];
    
    // Skip header
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (row.length < 1) continue;
      
      // Helper to clean and parse amount string (removes currency symbols, commas, etc.)
      const parseAmount = (val) => {
        if (!val) return 0;
        // Convert to string, remove everything except digits, dots, and minus sign
        const cleaned = String(val).replace(/[^0-9.-]/g, '');
        return parseFloat(cleaned) || 0;
      };

      collections.push({
        id: parseInt(row[0]),
        af56Id: row[1],
        orNumber: row[2],
        payor: row[3],
        barangay: row[4],
        landName: row[5],
        tdNumber: row[6],
        yearsPaid: row[7],
        amount: parseAmount(row[8]),
        date: row[9],
        remarks: row[10] || ''
      });
    }
    
    res.json({ result: 'success', collections });
  } catch (error) {
    if (error.message.includes('Unable to parse range')) {
      return res.json({ result: 'success', collections: [] });
    }
    console.error('Get RPT Collections error:', error);
    res.status(500).json({ result: 'error', message: error.message });
  }
});

app.post('/api/rpt-collections', async (req, res) => {
  try {
    const { collection } = req.body;
    if (!collection) return res.status(400).json({ message: 'Missing data' });

    const rows = await getSheetData("'RPT Collections'!A:A");
    let rowIndex = -1;
    
    // Check for update
    if (collection.id) {
        for (let i = 1; i < rows.length; i++) {
            if (parseInt(rows[i][0]) === collection.id) {
                rowIndex = i + 1;
                break;
            }
        }
    }

    const values = [
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
      // Update
      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: `'RPT Collections'!A${rowIndex}:K${rowIndex}`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values: [values] }
      });
    } else {
      // Append
      await appendSheetData("'RPT Collections'!A:K", values);
    }

    res.json({ result: 'success' });
  } catch (error) {
    console.error('Save RPT Collection error:', error);
    res.status(500).json({ result: 'error', message: error.message });
  }
});

app.post('/api/rpt-collections/delete', async (req, res) => {
  try {
    const { id } = req.body;
    
    const rows = await getSheetData("'RPT Collections'!A:A");
    let rowIndex = -1;
    
    for (let i = 1; i < rows.length; i++) {
      if (parseInt(rows[i][0]) === id) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex > -1) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        requestBody: {
          requests: [
            {
              deleteDimension: {
                range: {
                  sheetId: await getSheetId('RPT Collections'),
                  dimension: 'ROWS',
                  startIndex: rowIndex - 1,
                  endIndex: rowIndex
                }
              }
            }
          ]
        }
      });
    }

    res.json({ result: 'success' });
  } catch (error) {
    console.error('Delete RPT Collection error:', error);
    res.status(500).json({ result: 'error', message: error.message });
  }
});

// 11. SETTINGS / SPREADSHEET CONFIG
app.get('/api/settings/spreadsheet', (req, res) => {
  res.json({
    result: 'success',
    spreadsheetId: SPREADSHEET_ID || '',
    spreadsheetUrl: SPREADSHEET_ID ? `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/edit` : ''
  });
});

app.post('/api/settings/spreadsheet', async (req, res) => {
  try {
    const { spreadsheetUrl, spreadsheetId: rawId } = req.body;
    const input = spreadsheetUrl || rawId;
    const extractedId = extractSpreadsheetId(input);

    if (!extractedId) {
      return res.json({ result: 'error', message: 'Invalid Google Sheet URL or ID format.' });
    }

    console.log(`Updating Google Sheet database to ID: ${extractedId}...`);
    
    // Verify access and setup sheets
    const verifiedId = await ensureSpreadsheet(extractedId);
    SPREADSHEET_ID = verifiedId;
    updateEnvSpreadsheetId(verifiedId);

    console.log(`Google Sheet database updated successfully: ${verifiedId}`);
    return res.json({
      result: 'success',
      message: 'Google Sheet database updated and structure verified successfully.',
      spreadsheetId: verifiedId,
      spreadsheetUrl: `https://docs.google.com/spreadsheets/d/${verifiedId}/edit`
    });
  } catch (error) {
    console.error('Update Spreadsheet error:', error.message);
    return res.json({
      result: 'error',
      message: error.message
    });
  }
});

// Action dispatcher for POST requests (GAS compatibility mode & REST API)
async function handleAction(req, res) {
  const { action } = req.body || {};
  if (!action) return false;

  try {
    switch (action) {
      case 'getSpreadsheet':
        res.json({
          result: 'success',
          spreadsheetId: SPREADSHEET_ID || '',
          spreadsheetUrl: SPREADSHEET_ID ? `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/edit` : ''
        });
        return true;

      case 'updateSpreadsheet':
      case 'setSpreadsheet': {
        const input = req.body.spreadsheetUrl || req.body.spreadsheetId || req.body.url || req.body.id;
        const extractedId = extractSpreadsheetId(input);
        if (!extractedId) {
          res.json({ result: 'error', message: 'Invalid Google Sheet URL or ID format.' });
          return true;
        }
        try {
          const verifiedId = await ensureSpreadsheet(extractedId);
          SPREADSHEET_ID = verifiedId;
          updateEnvSpreadsheetId(verifiedId);
          res.json({
            result: 'success',
            message: 'Google Sheet database updated and structure verified successfully.',
            spreadsheetId: verifiedId,
            spreadsheetUrl: `https://docs.google.com/spreadsheets/d/${verifiedId}/edit`
          });
        } catch (err) {
          console.error('Update Spreadsheet error:', err.message);
          res.json({
            result: 'error',
            message: err.message
          });
        }
        return true;
      }

      case 'login': {
        const { email, password } = req.body;
        const rows = await getSheetData('Users!A:D');
        const user = rows.slice(1).find(row => row[0] === email && row[1] === password);
        if (user) {
          res.json({ result: 'success', email: user[0], name: user[2], role: user[3] });
        } else {
          res.json({ result: 'error', message: 'Invalid credentials' });
        }
        return true;
      }

      case 'getReports': {
        const rows = await getSheetData('Reports!A:G');
        const reports = [];
        for (let i = rows.length - 1; i >= 1; i--) {
          if (reports.length >= 50) break;
          try {
            const jsonStr = rows[i][6];
            if (jsonStr) reports.push(JSON.parse(jsonStr));
          } catch (e) {}
        }
        res.json({ result: 'success', reports });
        return true;
      }

      case 'submitReport': {
        const { report } = req.body;
        if (!report) { res.json({ result: 'error', message: 'Missing report data' }); return true; }
        const row = [report.date, report.reportNumber, report.collectorName, report.fundType, report.totalCollection, report.status, JSON.stringify(report)];
        await appendSheetData('Reports!A:G', row);
        res.json({ result: 'success', message: 'Report saved' });
        return true;
      }

      case 'getAccountCodes': {
        try {
          const rows = await getSheetData("'Account Codes'!A:D");
          const codes = [];
          for (let i = 1; i < rows.length; i++) {
            const r = rows[i];
            if (r.length < 4) continue;
            codes.push({ id: parseInt(r[0]), mainCategory: r[1], subCategory: r[2], code: r[3] });
          }
          res.json({ result: 'success', accountCodes: codes });
        } catch (err) {
          res.json({ result: 'success', accountCodes: [] });
        }
        return true;
      }

      case 'saveAccountCode': {
        const { accountCode } = req.body;
        if (!accountCode) { res.json({ result: 'error', message: 'Missing data' }); return true; }
        const rows = await getSheetData("'Account Codes'!A:A");
        let rowIndex = -1;
        for (let i = 1; i < rows.length; i++) {
          if (parseInt(rows[i][0]) === accountCode.id) { rowIndex = i + 1; break; }
        }
        const values = [accountCode.id, accountCode.mainCategory, accountCode.subCategory, accountCode.code];
        if (rowIndex > -1) {
          await sheets.spreadsheets.values.update({
            spreadsheetId: SPREADSHEET_ID,
            range: `'Account Codes'!A${rowIndex}:D${rowIndex}`,
            valueInputOption: 'USER_ENTERED',
            requestBody: { values: [values] }
          });
        } else {
          await appendSheetData("'Account Codes'!A:D", values);
        }
        res.json({ result: 'success' });
        return true;
      }

      case 'deleteAccountCode': {
        const { id } = req.body;
        const rows = await getSheetData("'Account Codes'!A:A");
        let rowIndex = -1;
        for (let i = 1; i < rows.length; i++) {
          if (parseInt(rows[i][0]) === id) { rowIndex = i + 1; break; }
        }
        if (rowIndex > -1) {
          await sheets.spreadsheets.batchUpdate({
            spreadsheetId: SPREADSHEET_ID,
            requestBody: {
              requests: [{ deleteDimension: { range: { sheetId: await getSheetId('Account Codes'), dimension: 'ROWS', startIndex: rowIndex - 1, endIndex: rowIndex } } }]
            }
          });
        }
        res.json({ result: 'success' });
        return true;
      }

      case 'getCollections': {
        try {
          const rows = await getSheetData("Collections!A:J");
          const entries = [];
          for (let i = 1; i < rows.length; i++) {
            const r = rows[i];
            if (r.length < 1) continue;
            entries.push({
              id: parseInt(r[0]), afNo: r[1], orNo: r[2], payor: r[3], subCategory: r[4], mainCategory: r[5], accountCode: r[6], amount: parseFloat(r[7] || 0), date: r[8], remarks: r[9]
            });
          }
          res.json({ result: 'success', entries });
        } catch (err) {
          res.json({ result: 'success', entries: [] });
        }
        return true;
      }

      case 'saveCollection': {
        const { entry } = req.body;
        if (!entry) { res.json({ result: 'error', message: 'Missing entry' }); return true; }
        const row = [entry.id, entry.afNo, entry.orNo, entry.payor, entry.subCategory, entry.mainCategory, entry.accountCode, entry.amount, entry.date, entry.remarks];
        await appendSheetData("Collections!A:J", row);
        res.json({ result: 'success' });
        return true;
      }

      case 'saveCollectionBulk': {
        const { header, charges } = req.body;
        if (!header || !Array.isArray(charges)) { res.json({ result: 'error', message: 'Missing data' }); return true; }
        const startId = await getNextCollectionId();
        for (let i = 0; i < charges.length; i++) {
          const c = charges[i];
          const row = [startId + i, header.afNo, header.orNo, header.payor, c.subCategory, c.mainCategory, c.accountCode, c.amount, header.date, header.remarks];
          await appendSheetData("Collections!A:J", row);
        }
        res.json({ result: 'success', startId, count: charges.length });
        return true;
      }

      case 'updateCollection': {
        const { entry } = req.body;
        if (!entry || !entry.id) { res.json({ result: 'error', message: 'Missing entry' }); return true; }
        const rows = await getSheetData("Collections!A:A");
        let rowIndex = -1;
        for (let i = 1; i < rows.length; i++) {
          if (parseInt(rows[i][0]) === entry.id) { rowIndex = i + 1; break; }
        }
        if (rowIndex > -1) {
          const row = [entry.id, entry.afNo, entry.orNo, entry.payor, entry.subCategory, entry.mainCategory, entry.accountCode, entry.amount, entry.date, entry.remarks];
          await sheets.spreadsheets.values.update({
            spreadsheetId: SPREADSHEET_ID,
            range: `Collections!A${rowIndex}:J${rowIndex}`,
            valueInputOption: 'USER_ENTERED',
            requestBody: { values: [row] }
          });
          res.json({ result: 'success' });
        } else {
          res.json({ result: 'error', message: 'Entry not found' });
        }
        return true;
      }

      case 'deleteCollection': {
        const { id } = req.body;
        const rows = await getSheetData("Collections!A:A");
        let rowIndex = -1;
        for (let i = 1; i < rows.length; i++) {
          if (parseInt(rows[i][0]) === id) { rowIndex = i + 1; break; }
        }
        if (rowIndex > -1) {
          await sheets.spreadsheets.batchUpdate({
            spreadsheetId: SPREADSHEET_ID,
            requestBody: {
              requests: [{ deleteDimension: { range: { sheetId: await getSheetId('Collections'), dimension: 'ROWS', startIndex: rowIndex - 1, endIndex: rowIndex } } }]
            }
          });
        }
        res.json({ result: 'success' });
        return true;
      }

      case 'getSignatories': {
        try {
          const rows = await getSheetData("Signatories!A:E");
          const signatories = [];
          for (let i = 1; i < rows.length; i++) {
            const r = rows[i];
            if (r.length < 4) continue;
            signatories.push({ id: parseInt(r[0]), fullName: r[1], position: r[2], department: r[3], remarks: r[4] || '' });
          }
          res.json({ result: 'success', signatories });
        } catch (err) {
          res.json({ result: 'success', signatories: [] });
        }
        return true;
      }

      case 'saveSignatory': {
        const { signatory } = req.body;
        if (!signatory) { res.json({ result: 'error', message: 'Missing data' }); return true; }
        const rows = await getSheetData("Signatories!A:A");
        let rowIndex = -1;
        if (signatory.id) {
          for (let i = 1; i < rows.length; i++) {
            if (parseInt(rows[i][0]) === signatory.id) { rowIndex = i + 1; break; }
          }
        }
        const values = [signatory.id, signatory.fullName, signatory.position, signatory.department, signatory.remarks || ''];
        if (rowIndex > -1) {
          await sheets.spreadsheets.values.update({
            spreadsheetId: SPREADSHEET_ID,
            range: `Signatories!A${rowIndex}:E${rowIndex}`,
            valueInputOption: 'USER_ENTERED',
            requestBody: { values: [values] }
          });
        } else {
          await appendSheetData("Signatories!A:E", values);
        }
        res.json({ result: 'success' });
        return true;
      }

      case 'deleteSignatory': {
        const { id } = req.body;
        const rows = await getSheetData("Signatories!A:A");
        let rowIndex = -1;
        for (let i = 1; i < rows.length; i++) {
          if (parseInt(rows[i][0]) === id) { rowIndex = i + 1; break; }
        }
        if (rowIndex > -1) {
          await sheets.spreadsheets.batchUpdate({
            spreadsheetId: SPREADSHEET_ID,
            requestBody: {
              requests: [{ deleteDimension: { range: { sheetId: await getSheetId('Signatories'), dimension: 'ROWS', startIndex: rowIndex - 1, endIndex: rowIndex } } }]
            }
          });
        }
        res.json({ result: 'success' });
        return true;
      }

      case 'getRPTCollections': {
        try {
          const rows = await getSheetData("'RPT Collections'!A:K");
          const parseAmount = (val) => {
            if (!val) return 0;
            const cleaned = String(val).replace(/[^0-9.-]/g, '');
            return parseFloat(cleaned) || 0;
          };
          const collections = [];
          for (let i = 1; i < rows.length; i++) {
            const r = rows[i];
            if (r.length < 1) continue;
            collections.push({
              id: parseInt(r[0]), af56Id: r[1], orNumber: r[2], payor: r[3], barangay: r[4], landName: r[5], tdNumber: r[6], yearsPaid: r[7], amount: parseAmount(r[8]), date: r[9], remarks: r[10] || ''
            });
          }
          res.json({ result: 'success', collections });
        } catch (err) {
          res.json({ result: 'success', collections: [] });
        }
        return true;
      }

      case 'saveRPTCollection': {
        const { collection } = req.body;
        if (!collection) { res.json({ result: 'error', message: 'Missing data' }); return true; }
        const rows = await getSheetData("'RPT Collections'!A:A");
        let rowIndex = -1;
        if (collection.id) {
          for (let i = 1; i < rows.length; i++) {
            if (parseInt(rows[i][0]) === collection.id) { rowIndex = i + 1; break; }
          }
        }
        const values = [collection.id, collection.af56Id, collection.orNumber, collection.payor, collection.barangay, collection.landName, collection.tdNumber, collection.yearsPaid, collection.amount, collection.date, collection.remarks || ''];
        if (rowIndex > -1) {
          await sheets.spreadsheets.values.update({
            spreadsheetId: SPREADSHEET_ID,
            range: `'RPT Collections'!A${rowIndex}:K${rowIndex}`,
            valueInputOption: 'USER_ENTERED',
            requestBody: { values: [values] }
          });
        } else {
          await appendSheetData("'RPT Collections'!A:K", values);
        }
        res.json({ result: 'success' });
        return true;
      }

      case 'deleteRPTCollection': {
        const { id } = req.body;
        const rows = await getSheetData("'RPT Collections'!A:A");
        let rowIndex = -1;
        for (let i = 1; i < rows.length; i++) {
          if (parseInt(rows[i][0]) === id) { rowIndex = i + 1; break; }
        }
        if (rowIndex > -1) {
          await sheets.spreadsheets.batchUpdate({
            spreadsheetId: SPREADSHEET_ID,
            requestBody: {
              requests: [{ deleteDimension: { range: { sheetId: await getSheetId('RPT Collections'), dimension: 'ROWS', startIndex: rowIndex - 1, endIndex: rowIndex } } }]
            }
          });
        }
        res.json({ result: 'success' });
        return true;
      }

      default:
        return false;
    }
  } catch (err) {
    console.error(`Error handling action ${action}:`, err);
    res.status(500).json({ result: 'error', message: err.message });
    return true;
  }
}

app.use(async (req, res, next) => {
  if (req.method === 'POST') {
    const handled = await handleAction(req, res);
    if (handled) return;
  }
  next();
});

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});


