// ===== GOOGLE APPS SCRIPT - BILLING SYSTEM API =====
// Paste this entire file into Extensions > Apps Script in your Google Sheet

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
