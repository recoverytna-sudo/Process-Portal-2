const express = require('express');
const cors = require('cors');
const { google } = require('googleapis');
require('dotenv').config();

const app = express();
app.use(cors());
app.use(express.json());

const auth = new google.auth.GoogleAuth({
  credentials: {
    client_email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
    private_key: process.env.GOOGLE_PRIVATE_KEY ? process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n') : '',
  },
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});

const spreadsheetId = process.env.GOOGLE_SPREADSHEET_ID;

async function getSheetData(sheetName) {
  const service = google.sheets({ version: 'v4', auth });
  const res = await service.spreadsheets.values.get({
    spreadsheetId,
    range: `${sheetName}!A:Z`,
  });
  const rows = res.data.values;
  if (!rows || rows.length === 0) return [];

  const headers = rows[0];
  return rows.slice(1).map(row => {
    let obj = {};
    headers.forEach((h, index) => {
      obj[h] = row[index] !== undefined ? row[index] : '';
    });
    return obj;
  });
}

async function appendSheetRow(sheetName, rowData) {
  const service = google.sheets({ version: 'v4', auth });
  const res = await service.spreadsheets.values.get({
    spreadsheetId,
    range: `${sheetName}!A1:Z1`,
  });
  const headers = res.data.values[0];
  const newRow = headers.map(h => rowData[h] !== undefined ? rowData[h] : '');

  await service.spreadsheets.values.append({
    spreadsheetId,
    range: `${sheetName}!A:A`,
    valueInputOption: 'USER_ENTERED',
    resource: { values: [newRow] },
  });
}

async function updateSheetRow(sheetName, idField, idValue, updatedData) {
  const service = google.sheets({ version: 'v4', auth });
  const res = await service.spreadsheets.values.get({
    spreadsheetId,
    range: `${sheetName}!A:Z`,
  });
  const rows = res.data.values;
  if (!rows || rows.length < 2) return false;

  const headers = rows[0];
  const idIndex = headers.indexOf(idField);
  if (idIndex === -1) return false;

  let rowIndex = -1;
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][idIndex] === idValue) {
      rowIndex = i + 1; // 1-based index for sheets
      break;
    }
  }

  if (rowIndex === -1) return false;

  const existingRow = rows[rowIndex - 1];
  const mergedRow = headers.map((h, index) => {
    if (updatedData[h] !== undefined) return updatedData[h];
    return existingRow[index] !== undefined ? existingRow[index] : '';
  });

  await service.spreadsheets.values.update({
    spreadsheetId,
    range: `${sheetName}!A${rowIndex}:Z${rowIndex}`,
    valueInputOption: 'USER_ENTERED',
    resource: { values: [mergedRow] },
  });
  return true;
}

// --- API Endpoints ---

// 1. Categories
app.get('/api/categories', async (req, res) => {
  try {
    const data = await getSheetData('Categories');
    res.json(data);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// 2. Processes GET All
app.get('/api/processes', async (req, res) => {
  try {
    const data = await getSheetData('Processes');
    res.json(data);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// 3. Process Details GET
app.get('/api/processes/:id', async (req, res) => {
  try {
    const data = await getSheetData('Processes');
    const item = data.find(p => p.ProcessID === req.params.id);
    if (!item) return res.status(404).json({ error: 'Process not found' });
    res.json(item);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// 4. Create Process (Leader)
app.post('/api/processes', async (req, res) => {
  try {
    const { password, ...body } = req.body;
    if (password !== (process.env.LEADER_PASSWORD || '1366')) {
      return res.status(401).json({ error: 'Unauthorized: Invalid password' });
    }

    const processID = 'PROC-' + Date.now().toString(36).toUpperCase();
    const timestamp = new Date().toISOString();
    const newRecord = {
      ProcessID: processID,
      Title: body.Title || '',
      CategoryID: body.CategoryID || '',
      Summary: body.Summary || '',
      ProcessContent: body.ProcessContent || '',
      AttachmentLink: body.AttachmentLink || '',
      SearchKeywords: body.SearchKeywords || '',
      Status: body.Status || 'Published',
      EffectiveDate: body.EffectiveDate || '',
      CreatedDate: timestamp,
      UpdatedDate: timestamp
    };

    await appendSheetRow('Processes', newRecord);
    res.json({ success: true, ProcessID: processID });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// 5. Update Process (Leader)
app.put('/api/processes/:id', async (req, res) => {
  try {
    const { password, ...body } = req.body;
    if (password !== (process.env.LEADER_PASSWORD || '1366')) {
      return res.status(401).json({ error: 'Unauthorized: Invalid password' });
    }

    const updatedData = {
      ...body,
      UpdatedDate: new Date().toISOString()
    };

    const success = await updateSheetRow('Processes', 'ProcessID', req.params.id, updatedData);
    if (!success) return res.status(404).json({ error: 'Process not found' });
    res.json({ success: true });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// 6. Archive Process (Leader)
app.post('/api/processes/:id/archive', async (req, res) => {
  try {
    const { password } = req.body;
    if (password !== (process.env.LEADER_PASSWORD || '1366')) {
      return res.status(401).json({ error: 'Unauthorized: Invalid password' });
    }

    const updatedData = {
      Status: 'Archived',
      UpdatedDate: new Date().toISOString()
    };

    const success = await updateSheetRow('Processes', 'ProcessID', req.params.id, updatedData);
    if (!success) return res.status(404).json({ error: 'Process not found' });
    res.json({ success: true });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// 7. Promotions GET All
app.get('/api/promotions', async (req, res) => {
  try {
    const data = await getSheetData('Promotions');
    res.json(data);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// 8. Promotion Details GET
app.get('/api/promotions/:id', async (req, res) => {
  try {
    const data = await getSheetData('Promotions');
    const item = data.find(p => p.PromotionID === req.params.id);
    if (!item) return res.status(404).json({ error: 'Promotion not found' });
    res.json(item);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// 9. Create Promotion (Leader)
app.post('/api/promotions', async (req, res) => {
  try {
    const { password, ...body } = req.body;
    if (password !== (process.env.LEADER_PASSWORD || '1366')) {
      return res.status(401).json({ error: 'Unauthorized: Invalid password' });
    }

    const promotionID = 'PROMO-' + Date.now().toString(36).toUpperCase();
    const timestamp = new Date().toISOString();
    const newRecord = {
      PromotionID: promotionID,
      PromotionName: body.PromotionName || '',
      CategoryID: body.CategoryID || '',
      Description: body.Description || '',
      PromotionRules: body.PromotionRules || '',
      ApplicableProducts: body.ApplicableProducts || '',
      StartDate: body.StartDate || '',
      EndDate: body.EndDate || '',
      AttachmentLink: body.AttachmentLink || '',
      Status: body.Status || 'Published',
      CreatedDate: timestamp,
      UpdatedDate: timestamp
    };

    await appendSheetRow('Promotions', newRecord);
    res.json({ success: true, PromotionID: promotionID });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// 10. Update Promotion (Leader)
app.put('/api/promotions/:id', async (req, res) => {
  try {
    const { password, ...body } = req.body;
    if (password !== (process.env.LEADER_PASSWORD || '1366')) {
      return res.status(401).json({ error: 'Unauthorized: Invalid password' });
    }

    const updatedData = {
      ...body,
      UpdatedDate: new Date().toISOString()
    };

    const success = await updateSheetRow('Promotions', 'PromotionID', req.params.id, updatedData);
    if (!success) return res.status(404).json({ error: 'Promotion not found' });
    res.json({ success: true });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// 11. Cancel Promotion (Leader)
app.post('/api/promotions/:id/cancel', async (req, res) => {
  try {
    const { password } = req.body;
    if (password !== (process.env.LEADER_PASSWORD || '1366')) {
      return res.status(401).json({ error: 'Unauthorized: Invalid password' });
    }

    const updatedData = {
      Status: 'Cancelled',
      UpdatedDate: new Date().toISOString()
    };

    const success = await updateSheetRow('Promotions', 'PromotionID', req.params.id, updatedData);
    if (!success) return res.status(404).json({ error: 'Promotion not found' });
    res.json({ success: true });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// 12. Archive Promotion (Leader)
app.post('/api/promotions/:id/archive', async (req, res) => {
  try {
    const { password } = req.body;
    if (password !== (process.env.LEADER_PASSWORD || '1366')) {
      return res.status(401).json({ error: 'Unauthorized: Invalid password' });
    }

    const updatedData = {
      Status: 'Archived',
      UpdatedDate: new Date().toISOString()
    };

    const success = await updateSheetRow('Promotions', 'PromotionID', req.params.id, updatedData);
    if (!success) return res.status(404).json({ error: 'Promotion not found' });
    res.json({ success: true });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Leader Login Verification Endpoint
app.post('/api/login', (req, res) => {
  const { password } = req.body;
  if (password === (process.env.LEADER_PASSWORD || '1366')) {
    res.json({ success: true });
  } else {
    res.status(401).json({ success: false, error: 'Invalid password' });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
