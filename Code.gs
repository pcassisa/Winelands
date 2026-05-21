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
        'Date','Estate','Client Name','Address','Mobile','Technician',
        'Chlorine','pH','Alkalinity','Stabilizer','Water Clarity','Salt',
        'Chlorine Large','Chlorine Small','Chlorine Cup',
        'Alkalinity Increaser 100','Ph Increaser 200',
        'Stabilizer',
        'Algaecide 1 litre','Algaecide 2 litre',
        'Super Clear Tabs','Pool Acid','Alum Powder',
        'Vacuum, Skim, Brush, Weir & Pump Cleans',
        'Check Pump, Sump, APC, Timer',
        'Check Filter, Suction Leaks, APC hoses',
        'Vacuum Weir Lid','Weir Basket',
        'Hose - Dominator 1.2m','Hose - Dominator Weir End',
        'Hose - Dominator APC End','Hose - Zodiac Click-On',
        'Diaphragm - Zodiac Slant','Diaphragm - Zodiac Blunt',
        'Pool is set back to Filter','All taps/hoses are off','Timer set to program',
        'Notes'
      ]);
    }

    logSheet.appendRow([
      data.date, data.estate, data.clientName, data.clientAddress,
      data.clientMobile, data.technician,
      data.chlorine, data.pH, data.alkalinity, data.stabilizer,
      data.waterClarity, data.salt,
      data.chem1a, data.chem1b, data.chem1c,
      data.chem2, data.chem3, data.chem4,
      data.chem5a, data.chem5b,
      data.chem6, data.chem7, data.chem8,
      data.svc1, data.svc2, data.svc3,
      data.part1, data.part2,
      data.part3, data.part4, data.part5, data.part6,
      data.part7, data.part8,
      data.co1, data.co2, data.co3,
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
        'ContactName','POAddressLine1','DueDate','InvoiceNumber',
        'InvoiceDate','Description','Quantity','AccountCode','TaxType','Currency'
      ]);
    }

    xeroSheet.appendRow([
      data.clientName, data.clientAddress, dueDate, invoiceNum, xeroDate,
      'Pool Service Visit - ' + data.estate, 1, '200', 'Tax Exclusive', 'ZAR'
    ]);

    var chemList = [
      data.chem1a !== '—' ? 'Chlorine Large x' + data.chem1a : '',
      data.chem1b !== '—' ? 'Chlorine Small x' + data.chem1b : '',
      data.chem1c !== '—' ? 'Chlorine Cup x' + data.chem1c : '',
      data.chem2  !== '—' ? 'Alk Inc 100 x' + data.chem2 : '',
      data.chem3  !== '—' ? 'Ph Inc 200 x' + data.chem3 : '',
      data.chem4  !== '—' ? 'Stabilizer x' + data.chem4 : '',
      data.chem5a !== '—' ? 'Algaecide 1L x' + data.chem5a : '',
      data.chem5b !== '—' ? 'Algaecide 2L x' + data.chem5b : '',
      data.chem6  !== '—' ? 'Super Clear x' + data.chem6 : '',
      data.chem7  !== '—' ? 'Pool Acid x' + data.chem7 : '',
      data.chem8  !== '—' ? 'Alum Powder x' + data.chem8 : ''
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
