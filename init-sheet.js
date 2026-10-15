const { google } = require('googleapis');
require('dotenv').config();

const auth = new google.auth.GoogleAuth({
  credentials: {
    client_email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
    private_key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n'),
  },
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});

async function initSpreadsheet() {
  const service = google.sheets({ version: 'v4', auth });
  const spreadsheetId = process.env.GOOGLE_SPREADSHEET_ID;

  const requiredSheets = [
    {
      title: 'Categories',
      headers: ['CategoryID', 'CategoryName', 'Description', 'SortOrder', 'Status']
    },
    {
      title: 'Processes',
      headers: ['ProcessID', 'Title', 'CategoryID', 'Summary', 'ProcessContent', 'AttachmentLink', 'SearchKeywords', 'Status', 'EffectiveDate', 'CreatedDate', 'UpdatedDate']
    },
    {
      title: 'Promotions',
      headers: ['PromotionID', 'PromotionName', 'CategoryID', 'Description', 'PromotionRules', 'ApplicableProducts', 'StartDate', 'EndDate', 'AttachmentLink', 'Status', 'CreatedDate', 'UpdatedDate']
    }
  ];

  try {
    const spreadsheet = await service.spreadsheets.get({ spreadsheetId });
    const existingSheets = spreadsheet.data.sheets.map(s => s.properties.title);

    for (const sheet of requiredSheets) {
      if (!existingSheets.includes(sheet.title)) {
        await service.spreadsheets.batchUpdate({
          spreadsheetId,
          resource: {
            requests: [{ addSheet: { properties: { title: sheet.title } } }]
          }
        });
        console.log(`Created sheet: ${sheet.title}`);
      }

      // Check if headers exist, if empty add headers
      const res = await service.spreadsheets.values.get({
        spreadsheetId,
        range: `${sheet.title}!A1:Z1`
      });

      if (!res.data.values || res.data.values.length === 0) {
        await service.spreadsheets.values.update({
          spreadsheetId,
          range: `${sheet.title}!A1`,
          valueInputOption: 'RAW',
          resource: { values: [sheet.headers] }
        });
        console.log(`Initialized headers for: ${sheet.title}`);
      } else {
        console.log(`Sheet ${sheet.title} already has data/headers. Skipping header initialization.`);
      }
    }
    console.log('Database initialization completed successfully.');
  } catch (error) {
    console.error('Error initializing spreadsheet:', error);
  }
}

initSpreadsheet();
