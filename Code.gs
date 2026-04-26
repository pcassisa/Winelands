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
      logSheet.appendRow([
        // A-F: Client
        'Date', 'Estate', 'Client Name', 'Address', 'Mobile', 'Technician',
        // G-L: Water Test
        'Chlorine', 'pH', 'Alkalinity', 'Stabilizer', 'Water Clarity', 'Salt',
        // M-S: Chemicals (7)
        'Chlorine', 'Alkalinity Increaser 100', 'Alkalinity Increaser 200',
        'Everblue Stabiliser', 'Bioguard Algishield', 'Super Clear Tabs', 'Pool Acid',
        // T-V: Services (3)
        'Vacuum, Skim, Brush, Weir & Pump Cleans',
        'Check Pump, Sump, APC, Timer',
        'Check Filter, Suction Leaks, APC hoses',
        // W-X: Fixed Parts (2)
        'Vacuum Weir Lid', 'Weir Basket',
        // Y-AB: Hose types (4)
        'Hose - Dominator 1.2m', 'Hose - Dominator Weir End',
        'Hose - Dominator APC End', 'Hose - Zodiac Click-On',
        // AC-AD: Diaphragm types (2)
        'Diaphragm - Zodiac Slant', 'Diaphragm - Zodiac Blunt',
        // AE-AF: Closeout (2)
        'Pool is set back to Filter', 'All taps/hoses are off',
        // AG: Notes
        'Notes'
      ]);
    }

    logSheet.appendRow([
      // Client
      data.date, data.estate, data.clientName, data.clientAddress,
      data.clientMobile, data.technician,
      // Water Test
      data.chlorine, data.pH, data.alkalinity, data.stabilizer,
      data.waterClarity, data.salt,
      // Chemicals (7)
      data.chem1, data.chem2, data.chem3, data.chem4,
      data.chem5, data.chem6, data.chem7,
      // Services (3)
      data.svc1, data.svc2, data.svc3,
      // Parts (8)
      data.part1, data.part2,
      data.part3, data.part4, data.part5, data.part6,
      data.part7, data.part8,
      // Closeout (2)
      data.co1, data.co2,
      // Notes
      data.notes
    ]);

    var lastRow = logSheet.getLastRow();
    var invoiceNum = 'WPS-' + String(lastRow).padStart(4, '0');
    var dateParts = data.date.split('-');
    var xeroDate = dateParts[2] + '/' + dateParts[1] + '/' + dateParts[0];
    var d = new Date(data.date);
    d.setDate(d.getDate() + 7);
    var dueDate = String(d.getDate()).padStart(2,'0') + '/' +
                  String(d.getMonth()+1).padStart(2,'0') + '/' + d.getFullYear();

    if (xeroSheet.getLastRow() === 0) {
      xeroSheet.appendRow([
        'ContactName', 'POAddressLine1', 'DueDate', 'InvoiceNumber',
        'InvoiceDate', 'Description', 'Quantity', 'AccountCode', 'TaxType', 'Currency'
      ]);
    }

    xeroSheet.appendRow([
      data.clientName, data.clientAddress, dueDate, invoiceNum, xeroDate,
      'Pool Service Visit - ' + data.estate, 1, '200', 'Tax Exclusive', 'ZAR'
    ]);

    var chemList = [
      data.chem1 !== '—' ? 'Chlorine x' + data.chem1 : '',
      data.chem2 !== '—' ? 'Alk Inc 100 x' + data.chem2 : '',
      data.chem3 !== '—' ? 'Alk Inc 200 x' + data.chem3 : '',
      data.chem4 !== '—' ? 'Everblue x' + data.chem4 : '',
      data.chem5 !== '—' ? 'Algishield ' + data.chem5 : '',
      data.chem6 !== '—' ? 'Super Clear x' + data.chem6 : '',
      data.chem7 !== '—' ? 'Pool Acid ' + data.chem7 : ''
    ].filter(function(x) { return x !== ''; }).join('; ');

    if (chemList) {
      xeroSheet.appendRow([
        data.clientName, data.clientAddress, dueDate, invoiceNum, xeroDate,
        'Chemicals: ' + chemList, 1, '200', 'Tax Exclusive', 'ZAR'
      ]);
    }

    var allParts = [
      data.part1 !== '—' ? 'Vacuum Weir Lid x' + data.part1 : '',
      data.part2 !== '—' ? 'Weir Basket x' + data.part2 : '',
      data.part3 !== '—' ? 'Dominator 1.2m x' + data.part3 : '',
      data.part4 !== '—' ? 'Dominator Weir End x' + data.part4 : '',
      data.part5 !== '—' ? 'Dominator APC End x' + data.part5 : '',
      data.part6 !== '—' ? 'Zodiac Click-On x' + data.part6 : '',
      data.part7 !== '—' ? 'Zodiac Slant x' + data.part7 : '',
      data.part8 !== '—' ? 'Zodiac Blunt x' + data.part8 : ''
    ].filter(function(x) { return x !== ''; }).join('; ');

    if (allParts) {
      xeroSheet.appendRow([
        data.clientName, data.clientAddress, dueDate, invoiceNum, xeroDate,
        'Replacement Parts: ' + allParts, 1, '200', 'Tax Exclusive', 'ZAR'
      ]);
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
