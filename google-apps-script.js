// === PASTE THIS INTO Google Apps Script (Extensions > Apps Script) ===
// Deploy > New Deployment > Web App > Anyone can access
// IMPORTANT: After pasting, click Deploy > Manage Deployments > Edit (pencil) > New Version > Deploy

function doGet(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
  if (!sheet) sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];

  // Check if this is a save/delete action via GET
  if (e && e.parameter && e.parameter.payload) {
    var payload = JSON.parse(decodeURIComponent(e.parameter.payload));
    return handleAction(sheet, payload);
  }

  // Otherwise return all deals
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    return ContentService.createTextOutput(JSON.stringify({})).setMimeType(ContentService.MimeType.JSON);
  }
  var headers = data[0];
  var deals = {};
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var deal = {};
    for (var j = 0; j < headers.length; j++) {
      deal[headers[j]] = row[j] || '';
    }
    if (deal.genBusiness) deals[deal.genBusiness] = deal;
  }
  return ContentService.createTextOutput(JSON.stringify(deals)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
  if (!sheet) sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var payload = JSON.parse(e.postData.contents);
  return handleAction(sheet, payload);
}

function handleAction(sheet, payload) {
  var action = payload.action;

  if (action === 'save') {
    var deal = payload.deal;
    var headers = ['genBusiness','genSeller','genSellerEmail','genAE','genPM','genPackage','genDSR','genRecording','genNotes','genSessionType','genGoLive','genAvailability','genAttendees','genDate','genHardware','genMenuStatus','genSpecial','genSourcePOS','genItemCount','genMenuType','genBPOActivity','genMenuText','genMenuNotes','timestamp','savedBy'];

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
    }

    var data = sheet.getDataRange().getValues();
    var found = -1;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === deal.genBusiness) { found = i + 1; break; }
    }

    var row = headers.map(function(h) { return deal[h] || ''; });

    if (found > 0) {
      sheet.getRange(found, 1, 1, row.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }

    return ContentService.createTextOutput(JSON.stringify({status:'ok'})).setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'delete') {
    var name = payload.name;
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === name) { sheet.deleteRow(i + 1); break; }
    }
    return ContentService.createTextOutput(JSON.stringify({status:'ok'})).setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput(JSON.stringify({status:'unknown'})).setMimeType(ContentService.MimeType.JSON);
}
