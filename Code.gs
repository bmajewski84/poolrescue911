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
const CHEMICAL_LOG_SHEET_NAMES = ['Chemical Log', 'Chemical Usage'];
const MONTHLY_INVOICE_SHEET_NAMES = ['Monthly Invoice'];

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

  if (action === 'getInvoiceData') {
    return jsonResponse_(getInvoiceData_(e.parameter.month));
  }

  return jsonResponse_({ error: 'Unknown action: ' + action });
}

function doPost(e) {
  const params = e.parameter;

  if (params.action === 'createInvoice') {
    return jsonResponse_(createInvoice_(params));
  }

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
    'Photos': params.photos || ''
  };

  const row = headers.map(h => (h in valueMap) ? valueMap[h] : '');
  sheet.appendRow(row);

  const chemicals = parseJsonSafe_(params.chemicals, []);
  appendChemicalLog_(params, chemicals);

  const photos = parseJsonSafe_(params.photos_data, []);
  sendServiceEmail_(params, chemicals, photos);
}

// ── Chemical Log ───────────────────────────────────────────────

function appendChemicalLog_(params, chemicals) {
  if (!chemicals.length) return;

  const ss = getSpreadsheet_();
  const sheet = getSheetByNames_(ss, CHEMICAL_LOG_SHEET_NAMES);
  if (!sheet) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(h => String(h).trim());

  const serviceDate = params.service_date ? new Date(params.service_date) : new Date();
  const billedMonth = Utilities.formatDate(serviceDate, Session.getScriptTimeZone(), 'MMMM yyyy');

  chemicals.forEach(chem => {
    const amount = Number(chem.amount) || 0;
    const costPerUnit = Number(chem.costPerUnit) || 0;
    const lineTotal = Math.round(amount * costPerUnit * 100) / 100;

    const valueMap = {
      'Date': params.service_date || '',
      'Customer ID': params.customer_id || '',
      'Customer Name': params.customer_name || '',
      'Chemical': chem.name || '',
      'Amount': amount,
      'Unit': chem.unit || '',
      'Cost/Unit ($)': costPerUnit,
      'Line Total ($)': lineTotal,
      'Billed Month': billedMonth
    };

    const row = headers.map(h => (h in valueMap) ? valueMap[h] : '');
    sheet.appendRow(row);
  });
}

// ── Email ────────────────────────────────────────────────────

function parseJsonSafe_(text, fallback) {
  if (!text) return fallback;
  try {
    return JSON.parse(text);
  } catch (err) {
    return fallback;
  }
}

function sendServiceEmail_(params, chemicals, photos) {
  if (!params.customer_email) return;

  try {
    const inlineImages = {};
    let photosHtml = '';

    photos.forEach((dataUrl, i) => {
      const match = String(dataUrl).match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.+)$/);
      if (!match) return;
      const contentType = match[1];
      const base64Data = match[2];
      const cid = 'photo' + i;
      inlineImages[cid] = Utilities.newBlob(Utilities.base64Decode(base64Data), contentType, cid + '.jpg');
      photosHtml += `<img src="cid:${cid}" style="max-width:280px;border-radius:6px;margin:4px;border:1px solid #ddd;">`;
    });

    const htmlBody = buildEmailBody_(params, chemicals, photosHtml);

    GmailApp.sendEmail(params.customer_email, 'Pool Service Report - ' + (params.service_date || ''), '', {
      htmlBody: htmlBody,
      inlineImages: inlineImages,
      name: 'El Dorado Pool Rescue'
    });
  } catch (err) {
    // Don't fail the whole request if email sending fails.
    console.error('Email send failed: ' + err);
  }
}

// ── Invoicing / Stripe ───────────────────────────────────────

function getCurrentBillingMonth_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM yyyy');
}

// Returns each active customer's monthly service fee, chemical charges
// for the given billing month, and the total due.
function getInvoiceData_(month) {
  const billingMonth = month || getCurrentBillingMonth_();

  const customers = getCustomers_();

  const ss = getSpreadsheet_();
  const chemSheet = getSheetByNames_(ss, CHEMICAL_LOG_SHEET_NAMES);
  const chemRows = chemSheet ? sheetToObjects_(chemSheet) : [];

  return customers.map(c => {
    const chemicalTotal = chemRows
      .filter(r => String(r['Customer ID']).trim() === String(c.id).trim()
                 && String(r['Billed Month']).trim() === billingMonth)
      .reduce((sum, r) => sum + (Number(r['Line Total ($)']) || 0), 0);

    return {
      id: c.id,
      name: c.name,
      email: c.email,
      plan: c.plan,
      serviceFee: c.rate,
      chemicalTotal: Math.round(chemicalTotal * 100) / 100,
      total: Math.round((c.rate + chemicalTotal) * 100) / 100,
      billingMonth: billingMonth
    };
  });
}

function getStripeKey_() {
  const key = PropertiesService.getScriptProperties().getProperty('STRIPE_SECRET_KEY');
  if (!key) throw new Error('STRIPE_SECRET_KEY is not set in Script Properties.');
  return key;
}

function stripeRequest_(path, payload) {
  const response = UrlFetchApp.fetch('https://api.stripe.com/v1/' + path, {
    method: 'post',
    headers: { Authorization: 'Bearer ' + getStripeKey_() },
    payload: payload,
    muteHttpExceptions: true
  });

  const body = JSON.parse(response.getContentText());
  if (response.getResponseCode() >= 300) {
    throw new Error('Stripe error: ' + (body.error && body.error.message ? body.error.message : response.getContentText()));
  }
  return body;
}

// Creates a one-off Stripe Price for the invoice total, then a Payment
// Link pointing at it, and logs the result to the Monthly Invoice tab.
function createInvoice_(params) {
  const customerName = params.customer_name || '';
  const billingMonth = params.billing_month || getCurrentBillingMonth_();
  const amount = Number(params.amount) || 0;

  if (amount <= 0) {
    return { error: 'Invoice amount must be greater than 0.' };
  }

  const amountCents = Math.round(amount * 100);

  const price = stripeRequest_('prices', {
    'currency': 'usd',
    'unit_amount': String(amountCents),
    'product_data[name]': 'Pool Service - ' + customerName + ' - ' + billingMonth
  });

  const paymentLink = stripeRequest_('payment_links', {
    'line_items[0][price]': price.id,
    'line_items[0][quantity]': '1'
  });

  logInvoice_(params, billingMonth, amount, paymentLink.url);

  return { url: paymentLink.url };
}

function logInvoice_(params, billingMonth, amount, url) {
  const ss = getSpreadsheet_();
  const sheet = getSheetByNames_(ss, MONTHLY_INVOICE_SHEET_NAMES);
  if (!sheet) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(h => String(h).trim());

  const valueMap = {
    'Date': Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    'Customer ID': params.customer_id || '',
    'Customer Name': params.customer_name || '',
    'Billing Month': billingMonth,
    'Service Fee': Number(params.service_fee) || 0,
    'Chemical Charges': Number(params.chemical_total) || 0,
    'Total': amount,
    'Payment Link': url,
    'Status': 'Sent'
  };

  const row = headers.map(h => (h in valueMap) ? valueMap[h] : '');
  sheet.appendRow(row);
}

function buildEmailBody_(params, chemicals, photosHtml) {
  const navy = '#0c1a2c';
  const gold = '#c8a548';

  const readingsRows = [
    ['Chlorine (ppm)', params.chlorine],
    ['pH', params.ph],
    ['Alkalinity (ppm)', params.alkalinity],
    ['Calcium (ppm)', params.calcium],
    ['CYA (ppm)', params.cya],
    ['Salt (ppm)', params.salt],
    ['Temp (°F)', params.temp],
    ['LSI Score', params.lsi_score]
  ].filter(([, v]) => v !== '' && v !== undefined && v !== null)
   .map(([label, v]) => `<tr><td style="padding:4px 12px 4px 0;color:#666;">${label}</td><td style="padding:4px 0;font-weight:600;">${v}</td></tr>`)
   .join('');

  const chemicalsHtml = chemicals.length
    ? `<h3 style="font-family:'Oswald',sans-serif;color:${navy};font-size:14px;letter-spacing:.05em;text-transform:uppercase;margin:20px 0 8px;">Chemicals Added</h3>
       <ul style="margin:0;padding-left:20px;color:#333;">
         ${chemicals.map(c => `<li>${c.name}: ${c.amount} ${c.unit}</li>`).join('')}
       </ul>`
    : '';

  return `
  <div style="font-family:'Inter',Arial,sans-serif;max-width:600px;margin:0 auto;color:#222;">
    <div style="background:${navy};padding:20px;text-align:center;">
      <h1 style="font-family:'Oswald',sans-serif;color:${gold};font-size:20px;letter-spacing:.08em;margin:0;">EL DORADO POOL RESCUE</h1>
      <p style="color:#fff;font-size:12px;letter-spacing:.1em;text-transform:uppercase;margin:4px 0 0;">Service Report</p>
    </div>
    <div style="padding:20px;">
      <p>Hi ${params.customer_name},</p>
      <p>Your pool was serviced on <strong>${params.service_date || ''}</strong> (${params.visit_type || ''}).</p>
      <p>Pool condition on arrival: <strong>${params.condition || '—'}</strong></p>

      <h3 style="font-family:'Oswald',sans-serif;color:${navy};font-size:14px;letter-spacing:.05em;text-transform:uppercase;margin:20px 0 8px;">Water Readings</h3>
      <table style="border-collapse:collapse;font-size:14px;">${readingsRows}</table>

      ${chemicalsHtml}

      ${params.tasks_done ? `<h3 style="font-family:'Oswald',sans-serif;color:${navy};font-size:14px;letter-spacing:.05em;text-transform:uppercase;margin:20px 0 8px;">Tasks Completed</h3><p style="color:#333;">${params.tasks_done}</p>` : ''}

      ${params.notes ? `<h3 style="font-family:'Oswald',sans-serif;color:${navy};font-size:14px;letter-spacing:.05em;text-transform:uppercase;margin:20px 0 8px;">Tech Notes</h3><p style="color:#333;">${params.notes}</p>` : ''}

      ${photosHtml ? `<h3 style="font-family:'Oswald',sans-serif;color:${navy};font-size:14px;letter-spacing:.05em;text-transform:uppercase;margin:20px 0 8px;">Photos</h3><div>${photosHtml}</div>` : ''}

      <p style="margin-top:24px;color:#666;font-size:12px;">Thanks for choosing El Dorado Pool Rescue!</p>
    </div>
  </div>`;
}
