// ===== GOOGLE APPS SCRIPT - BILLING SYSTEM API =====
// Paste this entire file into Extensions > Apps Script in your Google Sheet

// This function forces Gmail permission scope - run it ONCE manually to authorize
function authorizeEmailPermission() {
  MailApp.getRemainingDailyQuota();
  Logger.log('Email permission authorized. Remaining daily quota: ' + MailApp.getRemainingDailyQuota());
}

const SHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  try {
    const params = e.parameter;
    const action = params.action;
    let result;

    switch(action) {
      case 'getAll':
        result = getAllData(params.sheet);
        break;
      case 'add':
        result = addRow(params.sheet, JSON.parse(e.postData.contents));
        break;
      case 'update':
        result = updateRow(params.sheet, JSON.parse(e.postData.contents));
        break;
      case 'delete':
        result = deleteRow(params.sheet, params.id);
        break;
      case 'getConfig':
        result = getConfig();
        break;
      case 'setConfig':
        result = setConfig(JSON.parse(e.postData.contents));
        break;
      case 'sendStatement':
        result = sendStatement(JSON.parse(e.postData.contents));
        break;
      default:
        result = { error: 'Unknown action' };
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getAllData(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const row = {};
    for (let j = 0; j < headers.length; j++) {
      row[headers[j]] = data[i][j];
    }
    rows.push(row);
  }
  return rows;
}

function addRow(sheetName, data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return { error: 'Sheet not found: ' + sheetName };
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(h => data[h] || '');
  sheet.appendRow(row);
  return { success: true, id: data.id };
}

function updateRow(sheetName, data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return { error: 'Sheet not found' };
  const allData = sheet.getDataRange().getValues();
  const headers = allData[0];
  const idCol = headers.indexOf('id');
  for (let i = 1; i < allData.length; i++) {
    if (allData[i][idCol] == data.id) {
      const row = headers.map(h => data[h] !== undefined ? data[h] : allData[i][headers.indexOf(h)]);
      sheet.getRange(i + 1, 1, 1, headers.length).setValues([row]);
      return { success: true };
    }
  }
  return { error: 'Row not found' };
}

function deleteRow(sheetName, id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return { error: 'Sheet not found' };
  const allData = sheet.getDataRange().getValues();
  const headers = allData[0];
  const idCol = headers.indexOf('id');
  for (let i = 1; i < allData.length; i++) {
    if (allData[i][idCol] == id) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { error: 'Row not found' };
}

function getConfig() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  if (!sheet) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const newSheet = ss.insertSheet('Config');
    newSheet.getRange(1, 1, 1, 2).setValues([['key', 'value']]);
    return {};
  }
  const data = sheet.getDataRange().getValues();
  const config = {};
  for (let i = 1; i < data.length; i++) {
    config[data[i][0]] = data[i][1];
  }
  return config;
}

function setConfig(data) {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  if (!sheet) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    sheet = ss.insertSheet('Config');
    sheet.getRange(1, 1, 1, 2).setValues([['key', 'value']]);
  }
  const allData = sheet.getDataRange().getValues();
  for (const key in data) {
    let found = false;
    for (let i = 1; i < allData.length; i++) {
      if (allData[i][0] === key) {
        sheet.getRange(i + 1, 2).setValue(data[key]);
        found = true;
        break;
      }
    }
    if (!found) {
      sheet.appendRow([key, data[key]]);
    }
  }
  return { success: true };
}

function sendStatement(data) {
  // data: { to, clientName, monthDisplay, invoices: [{invoiceNumber, date, serveeName, reference, amount, balance}], totalDue }
  const subject = 'Monthly Statement - ' + data.monthDisplay + ' - Willingham Process Service';
  
  let invoiceRows = '';
  data.invoices.forEach(function(inv) {
    invoiceRows += '<tr>' +
      '<td style="padding:8px;border-bottom:1px solid #e5e7eb;">' + inv.invoiceNumber + '</td>' +
      '<td style="padding:8px;border-bottom:1px solid #e5e7eb;">' + inv.date + '</td>' +
      '<td style="padding:8px;border-bottom:1px solid #e5e7eb;">' + inv.serveeName + '</td>' +
      '<td style="padding:8px;border-bottom:1px solid #e5e7eb;">' + (inv.reference || '') + '</td>' +
      '<td style="padding:8px;border-bottom:1px solid #e5e7eb;">' + inv.amount + '</td>' +
      '<td style="padding:8px;border-bottom:1px solid #e5e7eb;font-weight:bold;">' + inv.balance + '</td>' +
      '</tr>';
  });

  const htmlBody = '<div style="font-family:Arial,sans-serif;max-width:700px;margin:0 auto;">' +
    '<div style="margin-bottom:20px;">' +
      '<strong>Nathan Willingham</strong><br>' +
      'Willingham Process Service<br>' +
      '65 Pine Ave. #418, Long Beach, CA 90802<br>' +
      '(714)-350-7775' +
    '</div>' +
    '<h2 style="text-align:center;margin-bottom:5px;">MONTHLY STATEMENT</h2>' +
    '<p style="text-align:center;color:#64748b;margin-bottom:20px;">' + data.monthDisplay + '</p>' +
    '<p><strong>' + data.clientName + '</strong></p>' +
    '<h4 style="margin-top:15px;">Unpaid Invoices</h4>' +
    '<table style="width:100%;border-collapse:collapse;margin-top:10px;">' +
      '<tr style="background:#f8fafc;">' +
        '<th style="padding:8px;text-align:left;border-bottom:2px solid #e5e7eb;">Invoice #</th>' +
        '<th style="padding:8px;text-align:left;border-bottom:2px solid #e5e7eb;">Date</th>' +
        '<th style="padding:8px;text-align:left;border-bottom:2px solid #e5e7eb;">Servee</th>' +
        '<th style="padding:8px;text-align:left;border-bottom:2px solid #e5e7eb;">Reference</th>' +
        '<th style="padding:8px;text-align:left;border-bottom:2px solid #e5e7eb;">Amount</th>' +
        '<th style="padding:8px;text-align:left;border-bottom:2px solid #e5e7eb;">Balance Due</th>' +
      '</tr>' +
      invoiceRows +
    '</table>' +
    '<p style="margin-top:15px;font-weight:bold;font-size:1.1em;">Total Balance Due: ' + data.totalDue + '</p>' +
    '<p style="margin-top:20px;color:#64748b;">Please remit payment at your earliest convenience. Thank you for your business.</p>' +
  '</div>';

  MailApp.sendEmail({
    to: data.to,
    subject: subject,
    htmlBody: htmlBody
  });

  return { success: true, to: data.to };
}
