const { google } = require('googleapis');
require('dotenv').config();

const SCOPES = ['https://www.googleapis.com/auth/spreadsheets'];

let auth;
if (process.env.GOOGLE_SERVICE_ACCOUNT_JSON) {
  try {
    const creds = typeof process.env.GOOGLE_SERVICE_ACCOUNT_JSON === 'string'
      ? JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON)
      : process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
    auth = new google.auth.GoogleAuth({
      credentials: creds,
      scopes: SCOPES,
    });
  } catch (e) {
    console.error('Failed to parse GOOGLE_SERVICE_ACCOUNT_JSON env var:', e);
    auth = new google.auth.GoogleAuth({
      keyFile: process.env.GOOGLE_APPLICATION_CREDENTIALS || 'credentials.json',
      scopes: SCOPES,
    });
  }
} else {
  auth = new google.auth.GoogleAuth({
    keyFile: process.env.GOOGLE_APPLICATION_CREDENTIALS || 'credentials.json',
    scopes: SCOPES,
  });
}

const sheets = google.sheets({ version: 'v4', auth });

/**
 * Ensures a valid spreadsheet exists.
 * If the provided ID is invalid or missing, creates a new one.
 * @param {string} spreadsheetId 
 * @returns {Promise<string>} The valid spreadsheet ID
 */
async function ensureSpreadsheet(spreadsheetId) {
  let validId = spreadsheetId;

  // 1. Check if spreadsheet exists and is accessible
  if (validId) {
    try {
      console.log(`Checking spreadsheet access for ID: ${validId}...`);
      await sheets.spreadsheets.get({ spreadsheetId: validId });
      console.log('Spreadsheet found and accessible.');
    } catch (error) {
      const saEmail = 'credentials@rcd-lguconcepcion.iam.gserviceaccount.com';
      console.warn(`Spreadsheet access check failed for ID ${validId}:`, error.message);
      if (error.status === 403 || error.code === 403 || (error.message && error.message.toLowerCase().includes('permission'))) {
        throw new Error(`Google Sheet is not shared with the Service Account. Please share your Google Sheet with Editor access to: ${saEmail}`);
      }
      if (error.status === 404 || error.code === 404 || (error.message && error.message.toLowerCase().includes('not found'))) {
        throw new Error(`Google Sheet not found. Please verify the URL or ID.`);
      }
      throw new Error(`Unable to access Google Sheet (${error.message}). Please ensure it is shared with: ${saEmail}`);
    }
  }

  // 2. If no ID provided, create a default one
  if (!validId) {
    try {
      console.log('Creating new default spreadsheet...');
      const resource = { properties: { title: 'RCD System Database' } };
      const spreadsheet = await sheets.spreadsheets.create({
        resource,
        fields: 'spreadsheetId',
      });
      validId = spreadsheet.data.spreadsheetId;
      console.log(`New spreadsheet created with ID: ${validId}`);
    } catch (error) {
      console.error('Failed to create default spreadsheet:', error.message);
      throw new Error(`Failed to create new spreadsheet. Please create a sheet in Google Drive and paste its link above.`);
    }
  }

  // 3. Ensure Tabs and Structure
  await setupStructure(validId);

  return validId;
}

async function setupStructure(spreadsheetId) {
  try {
    console.log('Verifying sheet structure...');
    const meta = await sheets.spreadsheets.get({ spreadsheetId });
    const existingSheets = meta.data.sheets.map(s => s.properties.title);
    
    const sheetsToCreate = [
      { title: 'Users', header: ['Email', 'Password', 'Name', 'Role'] },
      { title: 'Reports', header: ['Date', 'Report Number', 'Collector', 'Fund Type', 'Total Collection', 'Status', 'JSON Data'] },
      { title: 'Account Codes', header: ['ID', 'Main Category', 'Sub Category', 'Account Code'] },
      { title: 'Collections', header: ['ID', 'AF No.', 'OR No.', 'Payor', 'Sub Category', 'Main Category', 'Account Code', 'Amount', 'Date', 'Remarks'] },
      { title: 'Signatories', header: ['ID', 'Full Name', 'Position', 'Department', 'Remarks'] },
      { title: 'RPT Collections', header: ['ID', 'AF56 ID', 'OR Number', 'Payor', 'Barangay', 'Land Name', 'TD Number', 'Years Paid', 'Amount', 'Date', 'Remarks'] }
    ];

    const requests = [];

    // Create missing sheets
    for (const sheet of sheetsToCreate) {
      if (!existingSheets.includes(sheet.title)) {
        console.log(`Preparing to create sheet: ${sheet.title}`);
        requests.push({
          addSheet: {
            properties: { title: sheet.title }
          }
        });
      }
    }

    if (requests.length > 0) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: { requests }
      });
      console.log('Created missing sheets.');
    }

    // Add headers if needed
    for (const sheet of sheetsToCreate) {
      const range = `${sheet.title}!A1`;
      const result = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range
      });
      
      if (!result.data.values || result.data.values.length === 0) {
        console.log(`Adding header to ${sheet.title}`);
        await sheets.spreadsheets.values.update({
          spreadsheetId,
          range: `${sheet.title}!A1`,
          valueInputOption: 'USER_ENTERED',
          requestBody: {
            values: [sheet.header]
          }
        });
        
        // Add default admin user
        if (sheet.title === 'Users') {
          console.log('Adding default admin user...');
          await sheets.spreadsheets.values.append({
            spreadsheetId,
            range: 'Users!A:D',
            valueInputOption: 'USER_ENTERED',
            requestBody: {
              values: [['admin@rcd.com', 'admin123', 'Admin User', 'admin']]
            }
          });
        }
      }
    }
    console.log('Structure verification complete.');
  } catch (error) {
    console.error('Structure setup failed:', error.message);
    throw error;
  }
}

// Allow standalone execution
if (require.main === module) {
  ensureSpreadsheet(process.env.SPREADSHEET_ID).then(id => {
    console.log(`Setup finished. Spreadsheet ID: ${id}`);
  });
}

module.exports = { ensureSpreadsheet };
