// ============================================================
// BARAKA VET SHOP - Google Apps Script Backend v1.2
// Nakili code hii yote kwenye Google Apps Script
// barakashop2026@gmail.com
//
// MABADILIKO v1.2:
// - Ongezwa safu ya "Bei ya Punguzo (Discount)" kwenye Bidhaa
// - Ongezwa safu ya "Bei Punguzo" na "Jumla Punguzo" kwenye Mauzo
// - Backward compatible: data ya zamani haipotei
// ============================================================

// ---- MABADILIKO YA LAZIMA ----
// Weka Spreadsheet ID yako hapa (kutoka URL ya Google Sheet yako)
const SPREADSHEET_ID = '1-A_deTEYeK8EnccncG1lLcM8ofJdRqOzTz8N89OnZwE';

// Majina ya karatasi (sheets) - usibadilishe isipokuwa unataka
const SHEETS = {
  products: 'Bidhaa',
  customers: 'Wateja',
  sales: 'Mauzo',
  workers: 'Wafanyakazi',
  salaryPayments: 'Mishahara',
  logs: 'Kumbukumbu'
};

// ============================================================
// HEADERS ZA KILA SHEET
// MUHIMU: Bei ya Punguzo imeongezwa kwa Bidhaa na Mauzo
// ============================================================
const HEADERS = {
  products: [
    'ID', 'Jina la Bidhaa', 'Kategoria',
    'Bei ya Kuuza (Tsh)', 'Bei ya Punguzo (Tsh)',  // <-- Bei Punguzo mpya
    'Bei ya Kununulia (Tsh)', 'Idadi Stokini', 'Kitengo',
    'Tarehe ya Kuisha', 'Namba ya Batch', 'Maelezo', 'Tarehe Iliyoingizwa'
  ],
  customers: [
    'ID', 'Jina Kamili', 'Simu', 'Anuani', 'Aina ya Mifugo',
    'Jumla Nunuzi (Tsh)', 'Tarehe ya Ununuzi wa Mwisho', 'Maelezo', 'Tarehe Iliyoingizwa'
  ],
  sales: [
    'ID', 'Tarehe', 'Mteja', 'Bidhaa (JSON)',
    'Jumla Bidhaa', 'Jumla (Tsh)', 'Jumla Punguzo (Tsh)',  // <-- Jumla Punguzo mpya
    'Kilicholipwa (Tsh)', 'Chenji (Tsh)', 'ID ya Mteja'
  ],
  workers: [
    'ID', 'Jina Kamili', 'Nafasi', 'Simu', 'Mshahara (Tsh)',
    'Tarehe ya Kuanza', 'Benki/M-Pesa', 'Tarehe Iliyoingizwa'
  ],
  salaryPayments: [
    'ID', 'ID Mfanyakazi', 'Jina Mfanyakazi', 'Mwezi',
    'Kiasi (Tsh)', 'Njia ya Malipo', 'Maelezo', 'Tarehe ya Malipo'
  ],
  logs: ['Tarehe', 'Kitendo', 'Sheet', 'Maelezo']
};

// ============================================================
// ENTRY POINT - Inapokea POST requests kutoka mfumo
// ============================================================
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    const sheet = data.sheet;
    const payload = data.data;

    let result;
    if (action === 'sync') {
      result = syncSheet(sheet, payload);
    } else if (action === 'getAll') {
      result = getAllData();
    } else if (action === 'append') {
      result = appendRow(sheet, payload);
    } else {
      result = { success: false, message: 'Kitendo hakijulikani: ' + action };
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    logError('doPost', err.toString());
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================================
// GET - Kupata data (kwa ajili ya kuload mfumo)
// ============================================================
function doGet(e) {
  try {
    const action = e.parameter.action || 'getAll';

    if (action === 'getAll') {
      const result = getAllData();
      return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'getSheet') {
      const sheetName = e.parameter.sheet;
      const result = getSheetData(sheetName);
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, data: result }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ success: false, message: 'Kitendo hakijulikani' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================================
// SYNC SHEET - Andika data yote upya kwenye sheet moja
// Backward compatible: inabadilisha headers kama column mpya imeongezwa
// ============================================================
function syncSheet(sheetKey, dataArray) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetName = SHEETS[sheetKey];
    if (!sheetName) return { success: false, message: 'Sheet haijulikani: ' + sheetKey };

    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }

    // Futa data za zamani (headers na data)
    sheet.clearContents();

    // Andika headers mpya (zikiwemo za discount)
    const headers = HEADERS[sheetKey];
    if (headers) {
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setValues([headers]);
      styleHeaders(headerRange);
    }

    // Andika data
    if (dataArray && dataArray.length > 0) {
      const rows = dataArray.map(item => convertToRow(sheetKey, item));
      // Hakikisha kila row ina urefu sawa na headers
      const colCount = headers ? headers.length : rows[0].length;
      const normalizedRows = rows.map(row => {
        while (row.length < colCount) row.push('');
        return row.slice(0, colCount);
      });
      sheet.getRange(2, 1, normalizedRows.length, colCount).setValues(normalizedRows);
    }

    sheet.autoResizeColumns(1, headers ? headers.length : 10);
    logAction('sync', sheetKey, 'Rekodi ' + (dataArray ? dataArray.length : 0) + ' zimehifadhiwa');
    return { success: true, message: 'Sheet ' + sheetName + ' imesasishwa. Rekodi: ' + (dataArray ? dataArray.length : 0) };
  } catch (err) {
    logError('syncSheet:' + sheetKey, err.toString());
    return { success: false, error: err.toString() };
  }
}

// ============================================================
// CONVERT OBJECT → ROW ARRAY
// Discount price imeongezwa kwa products na sales
// ============================================================
function convertToRow(sheetKey, item) {
  const now = new Date().toLocaleString('sw-TZ');
  switch (sheetKey) {
    case 'products':
      return [
        item.id || '',
        item.name || '',
        item.category || '',
        item.price || 0,
        item.discountPrice || 0,           // Bei ya Punguzo (mpya - column 5)
        item.cost || 0,
        item.qty || 0,
        item.unit || '',
        item.expiry || '',
        item.batch || '',
        item.desc || '',
        now
      ];

    case 'customers':
      return [
        item.id || '',
        item.name || '',
        item.phone || '',
        item.address || '',
        item.animals || '',
        item.totalSpent || 0,
        item.lastPurchase ? new Date(item.lastPurchase).toLocaleString('sw-TZ') : '',
        item.notes || '',
        now
      ];

    case 'sales':
      return [
        item.id || '',
        item.date ? new Date(item.date).toLocaleString('sw-TZ') : '',
        item.customerName || 'Mteja Kawaida',
        JSON.stringify(item.items || []),
        item.items ? item.items.length : 0,
        item.total || 0,
        item.totalDiscount || 0,           // Jumla Punguzo (mpya - column 7)
        item.paid || 0,
        item.change || 0,
        item.customerId || ''
      ];

    case 'workers':
      return [
        item.id || '',
        item.name || '',
        item.role || '',
        item.phone || '',
        item.salary || 0,
        item.start || '',
        item.bank || '',
        now
      ];

    case 'salaryPayments':
      return [
        item.id || '',
        item.workerId || '',
        item.workerName || '',
        item.month || '',
        item.amount || 0,
        item.method || '',
        item.note || '',
        item.date ? new Date(item.date).toLocaleString('sw-TZ') : ''
      ];

    default:
      return Object.values(item);
  }
}

// ============================================================
// GET ALL DATA - Rudisha data zote kutoka sheets zote
// ============================================================
function getAllData() {
  try {
    const result = {};
    Object.keys(SHEETS).forEach(key => {
      if (key !== 'logs') {
        result[key] = getSheetData(key);
      }
    });
    return { success: true, data: result };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ============================================================
// GET SHEET DATA - Soma data kutoka sheet moja
// Backward compatible: inashughulikia sheets za zamani bila discount column
// ============================================================
function getSheetData(sheetKey) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetName = SHEETS[sheetKey];
    if (!sheetName) return [];

    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return [];

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow <= 1) return [];

    // Tumia lastCol badala ya headers.length ili kushughulikia sheets za zamani
    const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

    return data
      .filter(row => row[0] !== '')
      .map(row => convertRowToObject(sheetKey, row));
  } catch (err) {
    logError('getSheetData:' + sheetKey, err.toString());
    return [];
  }
}

// ============================================================
// CONVERT ROW → OBJECT
// Backward compatible: kama safu ya discount haipo, inatumia 0
// ============================================================
function convertRowToObject(sheetKey, row) {
  switch (sheetKey) {
    case 'products':
      // row[4] = Bei Punguzo (mpya). Kama sheet ya zamani haina, itakuwa undefined → 0
      // row[5] zamani ilikuwa cost, sasa imeshuka mstari mmoja
      // Angalia kama sheet ina columns 12 (mpya) au 11 (zamani)
      if (row.length >= 12) {
        // Sheet mpya na discount column
        return {
          id: row[0], name: row[1], category: row[2],
          price: parseFloat(row[3]) || 0,
          discountPrice: parseFloat(row[4]) || 0,     // Mpya
          cost: parseFloat(row[5]) || 0,
          qty: parseInt(row[6]) || 0,
          unit: row[7], expiry: row[8], batch: row[9], desc: row[10]
        };
      } else {
        // Sheet ya zamani (11 columns) - bila discount column
        return {
          id: row[0], name: row[1], category: row[2],
          price: parseFloat(row[3]) || 0,
          discountPrice: 0,                            // Default 0 kwa data za zamani
          cost: parseFloat(row[4]) || 0,
          qty: parseInt(row[5]) || 0,
          unit: row[6], expiry: row[7], batch: row[8], desc: row[9]
        };
      }

    case 'customers':
      return {
        id: row[0], name: row[1], phone: row[2],
        address: row[3], animals: row[4],
        totalSpent: parseFloat(row[5]) || 0,
        lastPurchase: row[6], notes: row[7]
      };

    case 'sales':
      let items = [];
      try { items = JSON.parse(row[3]); } catch(e) { items = []; }

      if (row.length >= 10) {
        // Sheet mpya na discount column
        return {
          id: row[0], date: row[1], customerName: row[2],
          items: items,
          total: parseFloat(row[5]) || 0,
          totalDiscount: parseFloat(row[6]) || 0,     // Mpya
          paid: parseFloat(row[7]) || 0,
          change: parseFloat(row[8]) || 0,
          customerId: row[9]
        };
      } else {
        // Sheet ya zamani (9 columns) - bila discount column
        return {
          id: row[0], date: row[1], customerName: row[2],
          items: items,
          total: parseFloat(row[5]) || 0,
          totalDiscount: 0,                            // Default 0 kwa data za zamani
          paid: parseFloat(row[6]) || 0,
          change: parseFloat(row[7]) || 0,
          customerId: row[8]
        };
      }

    case 'workers':
      return {
        id: row[0], name: row[1], role: row[2], phone: row[3],
        salary: parseFloat(row[4]) || 0,
        start: row[5], bank: row[6]
      };

    case 'salaryPayments':
      return {
        id: row[0], workerId: row[1], workerName: row[2],
        month: row[3],
        amount: parseFloat(row[4]) || 0,
        method: row[5], note: row[6], date: row[7]
      };

    default:
      return {};
  }
}

// ============================================================
// UNDA SHEETS ZOTE - Run mara moja tu mwanzoni
// Pia inaweza kutumika kuboresha sheets za zamani (column mpya itaongezwa)
// ============================================================
function setupSpreadsheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  Object.keys(SHEETS).forEach(key => {
    const sheetName = SHEETS[key];
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      Logger.log('Sheet mpya imeundwa: ' + sheetName);
    }

    // Weka/sasisha headers (discount columns zimejumuishwa)
    const headers = HEADERS[key];
    if (headers) {
      // Angalia kama headers za sasa ni tofauti
      const existingHeaders = sheet.getLastRow() > 0
        ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
        : [];

      // Sasa daima sasisha headers ili kujumuisha columns mpya
      const range = sheet.getRange(1, 1, 1, headers.length);
      range.setValues([headers]);
      styleHeaders(range);
      sheet.autoResizeColumns(1, headers.length);
    }
  });

  colorSheetTabs(ss);
  Logger.log('✅ Baraka Vet Shop v1.2 - Sheets zote zimeandaliwa!');
  SpreadsheetApp.getUi().alert(
    '✅ Mfumo v1.2 umewekwa vizuri!\n\n' +
    'Mabadiliko mapya:\n' +
    '• Bidhaa: Safu ya "Bei ya Punguzo" imeongezwa\n' +
    '• Mauzo: Safu ya "Jumla Punguzo" imeongezwa\n\n' +
    'Data ya zamani haijapotea!\n\n' +
    'Sheets:\n• Bidhaa\n• Wateja\n• Mauzo\n• Wafanyakazi\n• Mishahara\n• Kumbukumbu'
  );
}

// ============================================================
// STYLE HELPERS
// ============================================================
function styleHeaders(range) {
  range.setBackground('#0f2d40')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center');
}

function colorSheetTabs(ss) {
  const colors = {
    'Bidhaa': '#1a7a4a',
    'Wateja': '#2563eb',
    'Mauzo': '#f59e0b',
    'Wafanyakazi': '#7c3aed',
    'Mishahara': '#dc2626',
    'Kumbukumbu': '#64748b'
  };
  Object.entries(colors).forEach(([name, color]) => {
    const sheet = ss.getSheetByName(name);
    if (sheet) sheet.setTabColor(color);
  });
}

// ============================================================
// LOGGING
// ============================================================
function logAction(action, sheet, description) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let logSheet = ss.getSheetByName(SHEETS.logs);
    if (!logSheet) return;
    logSheet.appendRow([new Date().toLocaleString('sw-TZ'), action, sheet, description]);
  } catch(e) { /* Usisimame kwa sababu ya log */ }
}

function logError(context, message) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let logSheet = ss.getSheetByName(SHEETS.logs);
    if (!logSheet) return;
    logSheet.appendRow([new Date().toLocaleString('sw-TZ'), 'ERROR', context, message]);
  } catch(e) { /* silent */ }
}

// ============================================================
// RIPOTI YA KILA SIKU - Tuma kwa barua pepe (Optional)
// Weka trigger: Time-driven → Day timer → 11pm
// ============================================================
function sendDailyReport() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const salesSheet = ss.getSheetByName(SHEETS.sales);
    if (!salesSheet || salesSheet.getLastRow() <= 1) return;

    const today = new Date().toDateString();
    const allData = salesSheet.getDataRange().getValues();
    let todayTotal = 0, todayCount = 0, todayDiscount = 0;

    for (let i = 1; i < allData.length; i++) {
      const rowDate = new Date(allData[i][1]).toDateString();
      if (rowDate === today) {
        todayTotal += parseFloat(allData[i][5]) || 0;
        // Column 6 = Jumla Punguzo (mpya), au 0 kwa data za zamani
        todayDiscount += parseFloat(allData[i][6]) || 0;
        todayCount++;
      }
    }

    const subject = '📊 Ripoti ya Baraka Vet Shop - ' + new Date().toLocaleDateString('sw-TZ');
    const body = `
Habari,

Hapa ni muhtasari wa mauzo ya leo:

📅 Tarehe: ${new Date().toLocaleDateString('sw-TZ')}
🛒 Idadi ya Mauzo: ${todayCount}
💰 Jumla ya Mauzo: Tsh ${todayTotal.toLocaleString()}
🏷️ Jumla ya Punguzo: Tsh ${todayDiscount.toLocaleString()}

Tazama ripoti kamili hapa:
https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}

Baraka Vet Shop
barakashop2026@gmail.com
+255 744 065 662
`;
    GmailApp.sendEmail('barakashop2026@gmail.com', subject, body);
    Logger.log('Ripoti ya kila siku imetumwa!');
  } catch (err) {
    Logger.log('Hitilafu ya ripoti: ' + err.toString());
  }
}

// ============================================================
// MENU YA GOOGLE SHEETS (Inaonekana pale unafungua Sheet)
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🐾 Baraka Vet Shop')
    .addItem('⚙️ Weka/Sasisha Mfumo (v1.2)', 'setupSpreadsheet')
    .addSeparator()
    .addItem('📊 Ripoti ya Leo', 'showTodayReport')
    .addItem('📧 Tuma Ripoti kwa Barua Pepe', 'sendDailyReport')
    .addSeparator()
    .addItem('🗑️ Futa Kumbukumbu (Logs)', 'clearLogs')
    .addToUi();
}

// ============================================================
// RIPOTI YA LEO - Inaonekana kwenye Google Sheets
// ============================================================
function showTodayReport() {
  const today = new Date().toDateString();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const salesSheet = ss.getSheetByName(SHEETS.sales);

  if (!salesSheet || salesSheet.getLastRow() <= 1) {
    SpreadsheetApp.getUi().alert('Hakuna mauzo leo.');
    return;
  }

  const allData = salesSheet.getDataRange().getValues();
  let total = 0, count = 0, discount = 0;
  for (let i = 1; i < allData.length; i++) {
    try {
      const d = new Date(allData[i][1]).toDateString();
      if (d === today) {
        total += parseFloat(allData[i][5]) || 0;
        discount += parseFloat(allData[i][6]) || 0;
        count++;
      }
    } catch(e) {}
  }

  SpreadsheetApp.getUi().alert(
    '📊 Ripoti ya Leo\n\n' +
    '📅 Tarehe: ' + new Date().toLocaleDateString('sw-TZ') + '\n' +
    '🛒 Mauzo: ' + count + '\n' +
    '💰 Jumla: Tsh ' + total.toLocaleString() + '\n' +
    '🏷️ Jumla Punguzo: Tsh ' + discount.toLocaleString()
  );
}

// ============================================================
// FUTA LOGS
// ============================================================
function clearLogs() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const logSheet = ss.getSheetByName(SHEETS.logs);
  if (!logSheet) return;
  if (logSheet.getLastRow() > 1) {
    logSheet.deleteRows(2, logSheet.getLastRow() - 1);
  }
  SpreadsheetApp.getUi().alert('Kumbukumbu zimefutwa.');
}