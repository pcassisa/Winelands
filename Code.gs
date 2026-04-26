function doPost(e) {
  return handleRequest(e);
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ result: 'ok', message: 'Sheet is live' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleRequest(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ss.getSheetByName('Service Reports');
    var xeroSheet = ss.getSheetByName('Xero Import');
    var data = JSON.parse(e.postData.contents);

    if (logSheet.getLastRow() === 0) {
      logSheet.appendRow(['Date','Estate','Client Name','Address','Mobile',
        'Technician','Chlorine','pH','Alkalinity','Stabilizer',
        'Water Clarity','Salt','Services','Chemicals','Parts','Notes']);
    }

    logSheet.appendRow([data.date, data.estate, data.clientName,
      data.clientAddress, data.clientMobile, data.technician,
      data.chlorine, data.pH, data.alkalinity, data.stabilizer,
      data.waterClarity, data.salt, data.services,
      data.chemicals, data.parts, data.notes]);

    var lastRow = logSheet.getLastRow();
    var invoiceNum = 'WPS-' + String(lastRow).padStart(4, '0');
    var dateParts = data.date.split('-');
    var xeroDate = dateParts[2] + '/' + dateParts[1] + '/' + dateParts[0];
    var d = new Date(data.date);
    d.setDate(d.getDate() + 7);
    var dueDate = String(d.getDate()).padStart(2,'0') + '/' +
                  String(d.getMonth()+1).padStart(2,'0') + '/' + d.getFullYear();

    if (xeroSheet.getLastRow() === 0) {
      xeroSheet.appendRow(['ContactName','POAddressLine1','DueDate',
        'InvoiceNumber','InvoiceDate','Description',
        'Quantity','AccountCode','TaxType','Currency']);
    }

    xeroSheet.appendRow([data.clientName, data.clientAddress,
      dueDate, invoiceNum, xeroDate,
      'Pool Service Visit - ' + data.estate,
      1, '200', 'Tax Exclusive', 'ZAR']);

    if (data.chemicals && data.chemicals.trim() !== '') {
      xeroSheet.appendRow([data.clientName, data.clientAddress,
        dueDate, invoiceNum, xeroDate,
        'Chemicals: ' + data.chemicals,
        1, '200', 'Tax Exclusive', 'ZAR']);
    }

    if (data.parts && data.parts.trim() !== '') {
      xeroSheet.appendRow([data.clientName, data.clientAddress,
        dueDate, invoiceNum, xeroDate,
        'Replacement Parts: ' + data.parts,
        1, '200', 'Tax Exclusive', 'ZAR']);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success', invoice: invoiceNum }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
