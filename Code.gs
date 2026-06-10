/**
 * El Dorado Pool Rescue - Operations Web App backend
 *
 * IMPORTANT: This file is the source of truth for the Apps Script project
 * bound to / used by the Customers spreadsheet. After editing this file
 * here, copy its contents into the Apps Script editor (script.google.com)
 * and create a new deployment so the Web App URL picks up the changes.
 */

// Always read from this exact spreadsheet, regardless of which file
// the script project happens to be bound to.
const SPREADSHEET_ID = '1AjxwtbRM3mtNMlEu_xxxl8Bb27Le7jyXog7g3z3XM0A';

// Possible tab names for each dataset, in priority order. The first
// matching sheet name found in the spreadsheet is used. Update these
// lists if you rename tabs.
const CUSTOMERS_SHEET_NAMES = ['Customers'];
const SERVICE_REPORTS_SHEET_NAMES = ['Service Log', 'Service Reports', 'Visits', 'Reports'];

function getSpreadsheet_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function getSheetByNames_(ss, names) {
  for (const name of names) {
    const sheet = ss.getSheetByName(name);
    if (sheet) return sheet;
  }
  return null;
}

// Reads a sheet into an array of objects keyed by the header row.
function sheetToObjects_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(h => String(h).trim());
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (row.every(c => c === '' || c === null)) continue; // skip blank rows
    const obj = {};
    headers.forEach((h, idx) => { obj[h] = row[idx]; });
    rows.push(obj);
  }
  return rows;
}

function jsonResponse_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const action = e.parameter.action;

  if (action === 'getCustomers') {
    return jsonResponse_(getCustomers_());
  }

  if (action === 'getLastReadings') {
    return jsonResponse_(getLastReadings_(e.parameter.customer));
  }

  return jsonResponse_({ error: 'Unknown action: ' + action });
}

function doPost(e) {
  const params = e.parameter;

  if (!params.customer_name) {
    return jsonResponse_({ error: 'Missing customer_name' });
  }

  appendServiceReport_(params);
  return jsonResponse_({ status: 'ok' });
}

// ── Customers ──────────────────────────────────────────────────

function getCustomers_() {
  const ss = getSpreadsheet_();
  const sheet = getSheetByNames_(ss, CUSTOMERS_SHEET_NAMES);
  if (!sheet) return [];

  const rows = sheetToObjects_(sheet);

  return rows
    .filter(r => String(r['Active']).trim().toLowerCase() !== 'no')
    .map(r => ({
      id: r['Customer ID'] || '',
      name: r['Full Name'] || '',
      email: r['Email'] || '',
      phone: r['Phone'] || '',
      address: r['Service Address'] || '',
      gallons: Number(r['Pool Gallons']) || 0,
      salt: String(r['Salt System'] || '').trim().toLowerCase() === 'yes' ? 'Yes' : 'No',
      plan: r['Service Plan'] || '',
      rate: Number(r['Monthly Rate ($)']) || 0,
      customTasks: r['Custom Tasks'] || ''
    }));
}

// ── Last readings for a customer ─────────────────────────────────

function getLastReadings_(customerName) {
  if (!customerName) return {};

  const ss = getSpreadsheet_();
  const sheet = getSheetByNames_(ss, SERVICE_REPORTS_SHEET_NAMES);
  if (!sheet) return {};

  const rows = sheetToObjects_(sheet)
    .filter(r => String(r['Customer Name']).trim() === customerName.trim());

  if (rows.length === 0) return {};

  // Last row in the sheet for this customer = most recent visit.
  const last = rows[rows.length - 1];

  return {
    date: last['Date'] || '',
    chlorine: last['Chlorine'],
    ph: last['pH'],
    alkalinity: last['Alkalinity'],
    calcium: last['Calcium'],
    cya: last['CYA'],
    salt: last['Salt'],
    temp: last['Temp'],
    lsi: last['LSI Score'],
    notes: last['Tech Notes'] || ''
  };
}

// ── Append a new service report row ───────────────────────────

function appendServiceReport_(params) {
  const ss = getSpreadsheet_();
  const sheet = getSheetByNames_(ss, SERVICE_REPORTS_SHEET_NAMES);
  if (!sheet) {
    throw new Error('Could not find a Service Reports sheet. Checked: ' + SERVICE_REPORTS_SHEET_NAMES.join(', '));
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(h => String(h).trim());

  // Map header name -> value from the submitted form.
  const valueMap = {
    'Date': params.service_date || '',
    'Customer ID': params.customer_id || '',
    'Customer Name': params.customer_name || '',
    'Address': params.address || '',
    'Visit Type': params.visit_type || '',
    'Condition': params.condition || '',
    'Chlorine': params.chlorine || '',
    'pH': params.ph || '',
    'Alkalinity': params.alkalinity || '',
    'Calcium': params.calcium || '',
    'CYA': params.cya || '',
    'Salt': params.salt || '',
    'Temp': params.temp || '',
    'LSI Score': params.lsi_score || '',
    'Custom Tasks Done': params.tasks_done || '',
    'Tech Notes': params.notes || '',
    'Next Visit': params.next_visit || '',
    'Photos': params.photos || ''
  };

  const row = headers.map(h => (h in valueMap) ? valueMap[h] : '');
  sheet.appendRow(row);

  // Send confirmation email to the customer.
  if (params.customer_email) {
    try {
      MailApp.sendEmail({
        to: params.customer_email,
        subject: 'Pool Service Report - ' + (params.service_date || ''),
        body: 'Hi ' + params.customer_name + ',\n\n' +
          'Your pool was serviced on ' + (params.service_date || '') + '.\n' +
          (params.notes ? ('Notes: ' + params.notes + '\n\n') : '\n') +
          'Next visit: ' + (params.next_visit || 'TBD') + '\n\n' +
          '- El Dorado Pool Rescue'
      });
    } catch (err) {
      // Don't fail the whole request if email sending fails.
      console.error('Email send failed: ' + err);
    }
  }
}
