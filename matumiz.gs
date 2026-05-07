/**
 * ═══════════════════════════════════════════════════════════════════
 *  UKM KITUO CHA MIFUGO — Mfumo wa Matumizi ya Duka
 *  Google Apps Script Backend
 *  Toleo: 1.0 | Tarehe: 2025
 * ═══════════════════════════════════════════════════════════════════
 *
 *  MAELEKEZO YA USANIDI (SETUP INSTRUCTIONS):
 *  ─────────────────────────────────────────────
 *  1. Fungua Google Sheets mpya: https://sheets.google.com
 *  2. Ipe jina "UKM - Matumizi ya Duka"
 *  3. Nenda: Extensions → Apps Script
 *  4. Futa msimbo wote wa awali, weka msimbo huu wote
 *  5. Bonyeza "Save" (Ctrl+S)
 *  6. Nenda: Deploy → New deployment
 *     - Type: Web app
 *     - Execute as: Me
 *     - Who has access: Anyone  ← (au "Anyone with Google account" kwa usalama zaidi)
 *  7. Bonyeza "Deploy" → Nakili URL inayotolewa
 *  8. Weka URL hiyo kwenye matumizi.html badala ya:
 *     const SCRIPT_URL = 'YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL_HERE';
 * ═══════════════════════════════════════════════════════════════════
 */

// ── Majina ya Laha (Sheet Names) ──────────────────────────────────
const SHEET_MATUMIZI  = 'Matumizi';
const SHEET_SUMMARY   = 'Muhtasari';
const SHEET_LOG       = 'Log ya Mfumo';

// ── Vichwa vya Jedwali (Column Headers) ───────────────────────────
const HEADERS_MATUMIZI = [
  'Tarehe',
  'Kumbukumbu',
  'Mtumiaji / Mfanyakazi',
  'Aina ya Matumizi',
  'Bidhaa / Huduma',
  'Kiasi',
  'Kitengo',
  'Bei kwa Kipande (TZS)',
  'Jumla ya Gharama (TZS)',
  'Chanzo cha Malipo',
  'Muuzaji / Mtoaji',
  'Idhini ya Mkurugenzi',
  'Hali ya Risiti',
  'Maelezo / Sababu',
  'Wakati wa Kuingiza',
  'Imeingizwa na (IP)',
];


// ═══════════════════════════════════════════════════════════════════
//  ENTRY POINTS — Mlango wa Ombi
// ═══════════════════════════════════════════════════════════════════

/**
 * GET handler — kupakia rekodi
 */
function doGet(e) {
  try {
    const action = e.parameter.action;
    
    if (action === 'getMatumizi') {
      return handleGetMatumizi(e);
    } else if (action === 'getSummary') {
      return handleGetSummary(e);
    } else if (action === 'getCategories') {
      return handleGetCategories(e);
    } else {
      return jsonResponse({ success: false, message: 'Hatua haijulikani: ' + action });
    }
  } catch (err) {
    logError('doGet', err);
    return jsonResponse({ success: false, message: 'Hitilafu ya mfumo: ' + err.message });
  }
}

/**
 * POST handler — kuingiza rekodi mpya
 */
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;

    if (action === 'addMatumizi') {
      return handleAddMatumizi(payload.data);
    } else if (action === 'updateMatumizi') {
      return handleUpdateMatumizi(payload.data);
    } else if (action === 'deleteMatumizi') {
      return handleDeleteMatumizi(payload.kumbukumbu);
    } else {
      return jsonResponse({ success: false, message: 'Hatua haijulikani: ' + action });
    }
  } catch (err) {
    logError('doPost', err);
    return jsonResponse({ success: false, message: 'Hitilafu ya mfumo: ' + err.message });
  }
}


// ═══════════════════════════════════════════════════════════════════
//  HANDLERS — Wasindikaji wa Vitendo
// ═══════════════════════════════════════════════════════════════════

/**
 * Ongeza rekodi mpya ya matumizi
 */
function handleAddMatumizi(data) {
  if (!data) return jsonResponse({ success: false, message: 'Hakuna data iliyotumwa.' });

  // Thibitisha sehemu zinazohitajika
  const required = ['tarehe', 'mtumiaji', 'aina', 'bidhaa', 'kiasi', 'jumla', 'chanzo'];
  for (const field of required) {
    if (!data[field] || String(data[field]).trim() === '') {
      return jsonResponse({ success: false, message: `Sehemu "${field}" inahitajika.` });
    }
  }

  const ss    = getOrCreateSpreadsheet();
  const sheet = getOrCreateSheet(ss, SHEET_MATUMIZI, HEADERS_MATUMIZI);

  // Hesabu ya kumbukumbu ya pekee
  const kumbukumbu = data.kumbukumbu || generateRef(sheet);

  // Angalia kama kumbukumbu tayari ipo
  if (referenceExists(sheet, kumbukumbu)) {
    return jsonResponse({ success: false, message: `Kumbukumbu ${kumbukumbu} tayari ipo.` });
  }

  // Andika safu mpya
  const newRow = [
    data.tarehe,
    kumbukumbu,
    sanitize(data.mtumiaji),
    sanitize(data.aina),
    sanitize(data.bidhaa),
    Number(data.kiasi) || 0,
    sanitize(data.kitengo || 'Vipande'),
    Number(data.bei) || 0,
    Number(data.jumla) || 0,
    sanitize(data.chanzo),
    sanitize(data.muuzaji || ''),
    sanitize(data.idhini || 'Haijaidhinishwa'),
    sanitize(data.risiti || 'Ipo'),
    sanitize(data.maelezo || ''),
    data.timestamp || new Date().toISOString(),
    Session.getActiveUser().getEmail() || 'Anonymous',
  ];

  sheet.appendRow(newRow);
  
  // Panga upya kulingana na tarehe (hiari — unaweza kuondoa kwa speed)
  sortSheetByDate(sheet);
  
  // Sasisisha muhtasari
  updateSummarySheet(ss);

  logActivity('ADD', kumbukumbu, data.mtumiaji);

  return jsonResponse({
    success: true,
    message: 'Matumizi yamehifadhiwa.',
    kumbukumbu: kumbukumbu,
    rowCount: sheet.getLastRow() - 1,
  });
}

/**
 * Pakia rekodi (hadi 200 za mwisho)
 */
function handleGetMatumizi(e) {
  const limit  = parseInt(e.parameter.limit)  || 50;
  const filter = e.parameter.filter || '';     // aina ya kuchuja
  const from   = e.parameter.from  || '';      // tarehe ya kuanzia
  const to     = e.parameter.to    || '';      // tarehe ya mwisho

  const ss    = getOrCreateSpreadsheet();
  const sheet = getOrCreateSheet(ss, SHEET_MATUMIZI, HEADERS_MATUMIZI);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return jsonResponse({ success: true, data: [], total: 0 });

  const allData = sheet.getRange(2, 1, lastRow - 1, HEADERS_MATUMIZI.length).getValues();

  // Chuja na upanga kwa utaratibu wa kupungua kwa tarehe
  let rows = allData
    .filter(r => r[0] !== '') // acha safu tupu
    .map(r => rowToObject(r))
    .filter(r => {
      if (filter && r.aina !== filter) return false;
      if (from   && r.tarehe < from)   return false;
      if (to     && r.tarehe > to)     return false;
      return true;
    })
    .reverse()                // mpya kwanza
    .slice(0, limit);

  return jsonResponse({ success: true, data: rows, total: rows.length });
}

/**
 * Pakia muhtasari wa takwimu
 */
function handleGetSummary(e) {
  const ss    = getOrCreateSpreadsheet();
  const sheet = getOrCreateSheet(ss, SHEET_MATUMIZI, HEADERS_MATUMIZI);
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return jsonResponse({ success: true, summary: emptySummary() });
  }

  const data = sheet.getRange(2, 1, lastRow - 1, HEADERS_MATUMIZI.length).getValues();
  const today = todayStr();
  const monthPrefix = today.substring(0, 7);
  const weekStart   = getWeekStart();

  let totalAll = 0, totalToday = 0, totalWeek = 0, totalMonth = 0;
  let countAll = 0, countToday = 0;
  const byCategory = {};

  data.forEach(r => {
    if (!r[0]) return;
    const tarehe = String(r[0]).substring(0, 10);
    const jumla  = Number(r[8]) || 0;
    const aina   = r[3] || 'Nyingine';

    totalAll += jumla; countAll++;
    if (tarehe === today)       { totalToday += jumla; countToday++; }
    if (tarehe >= weekStart)      totalWeek  += jumla;
    if (tarehe.startsWith(monthPrefix)) totalMonth += jumla;

    byCategory[aina] = (byCategory[aina] || 0) + jumla;
  });

  return jsonResponse({
    success: true,
    summary: {
      totalAll, totalToday, totalWeek, totalMonth,
      countAll, countToday, byCategory,
      lastUpdated: new Date().toISOString(),
    }
  });
}

/**
 * Sasisha rekodi iliyopo
 */
function handleUpdateMatumizi(data) {
  if (!data || !data.kumbukumbu) {
    return jsonResponse({ success: false, message: 'Kumbukumbu inahitajika.' });
  }

  const ss    = getOrCreateSpreadsheet();
  const sheet = getOrCreateSheet(ss, SHEET_MATUMIZI, HEADERS_MATUMIZI);
  const rowIdx = findRowByRef(sheet, data.kumbukumbu);

  if (rowIdx === -1) {
    return jsonResponse({ success: false, message: 'Rekodi haikupatikana: ' + data.kumbukumbu });
  }

  const updatedRow = [
    data.tarehe,
    data.kumbukumbu,
    sanitize(data.mtumiaji),
    sanitize(data.aina),
    sanitize(data.bidhaa),
    Number(data.kiasi) || 0,
    sanitize(data.kitengo || 'Vipande'),
    Number(data.bei) || 0,
    Number(data.jumla) || 0,
    sanitize(data.chanzo),
    sanitize(data.muuzaji || ''),
    sanitize(data.idhini || ''),
    sanitize(data.risiti || ''),
    sanitize(data.maelezo || ''),
    data.timestamp || new Date().toISOString(),
    Session.getActiveUser().getEmail() || 'Anonymous',
  ];

  sheet.getRange(rowIdx, 1, 1, updatedRow.length).setValues([updatedRow]);
  updateSummarySheet(ss);
  logActivity('UPDATE', data.kumbukumbu, data.mtumiaji);

  return jsonResponse({ success: true, message: 'Rekodi imesasishwa.' });
}

/**
 * Futa rekodi
 */
function handleDeleteMatumizi(kumbukumbu) {
  if (!kumbukumbu) return jsonResponse({ success: false, message: 'Kumbukumbu inahitajika.' });

  const ss    = getOrCreateSpreadsheet();
  const sheet = getOrCreateSheet(ss, SHEET_MATUMIZI, HEADERS_MATUMIZI);
  const rowIdx = findRowByRef(sheet, kumbukumbu);

  if (rowIdx === -1) {
    return jsonResponse({ success: false, message: 'Rekodi haikupatikana.' });
  }

  sheet.deleteRow(rowIdx);
  updateSummarySheet(ss);
  logActivity('DELETE', kumbukumbu, 'N/A');

  return jsonResponse({ success: true, message: 'Rekodi imefutwa.' });
}

/**
 * Pata orodha ya aina zilizopo
 */
function handleGetCategories() {
  const ss    = getOrCreateSpreadsheet();
  const sheet = getOrCreateSheet(ss, SHEET_MATUMIZI, HEADERS_MATUMIZI);
  const lastRow = sheet.getLastRow();
  
  if (lastRow < 2) return jsonResponse({ success: true, categories: [] });

  const ainaCol = sheet.getRange(2, 4, lastRow - 1, 1).getValues().flat();
  const unique  = [...new Set(ainaCol.filter(Boolean))].sort();

  return jsonResponse({ success: true, categories: unique });
}


// ═══════════════════════════════════════════════════════════════════
//  SHEET MANAGEMENT — Usimamizi wa Laha
// ═══════════════════════════════════════════════════════════════════

function getOrCreateSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getOrCreateSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    setupSheet(sheet, headers, name);
  } else if (sheet.getLastRow() === 0) {
    setupSheet(sheet, headers, name);
  }
  return sheet;
}

function setupSheet(sheet, headers, name) {
  // Weka vichwa
  const headerRow = sheet.getRange(1, 1, 1, headers.length);
  headerRow.setValues([headers]);

  // Pangua header
  headerRow.setBackground('#1b4332');
  headerRow.setFontColor('#f4c945');
  headerRow.setFontWeight('bold');
  headerRow.setFontSize(10);
  headerRow.setHorizontalAlignment('center');

  // Ganda (freeze) mstari wa kwanza
  sheet.setFrozenRows(1);

  // Upana wa safu wima
  if (name === SHEET_MATUMIZI) {
    const widths = [90, 100, 140, 120, 160, 60, 80, 100, 110, 110, 130, 110, 100, 200, 140, 120];
    widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  }

  // Alternate row colors (ukanda)
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=MOD(ROW(),2)=0')
    .setBackground('#f0f7f4')
    .setRanges([sheet.getDataRange()])
    .build();
  sheet.setConditionalFormatRules([rule]);
}

/**
 * Unda/sasisha laha ya muhtasari
 */
function updateSummarySheet(ss) {
  let sumSheet = ss.getSheetByName(SHEET_SUMMARY);
  if (!sumSheet) sumSheet = ss.insertSheet(SHEET_SUMMARY);
  sumSheet.clearContents();

  const matSheet = ss.getSheetByName(SHEET_MATUMIZI);
  if (!matSheet || matSheet.getLastRow() < 2) return;

  const data = matSheet.getRange(2, 1, matSheet.getLastRow() - 1, HEADERS_MATUMIZI.length).getValues();

  const today  = todayStr();
  const month  = today.substring(0, 7);
  const wStart = getWeekStart();

  let totalAll = 0, totalToday = 0, totalWeek = 0, totalMonth = 0, countAll = 0;
  const byAina = {};

  data.forEach(r => {
    if (!r[0]) return;
    const tarehe = String(r[0]).substring(0, 10);
    const jumla  = Number(r[8]) || 0;
    const aina   = r[3] || 'Nyingine';

    totalAll += jumla; countAll++;
    if (tarehe === today)        totalToday += jumla;
    if (tarehe >= wStart)        totalWeek  += jumla;
    if (tarehe.startsWith(month)) totalMonth += jumla;

    byAina[aina] = (byAina[aina] || 0) + jumla;
  });

  const summaryData = [
    ['UKM KITUO CHA MIFUGO — MUHTASARI WA MATUMIZI', ''],
    ['Imesasishwa', new Date().toLocaleString('sw-TZ')],
    ['', ''],
    ['KIPINDI', 'JUMLA (TZS)'],
    ['Leo', totalToday],
    ['Wiki Hii', totalWeek],
    ['Mwezi Huu', totalMonth],
    ['Jumla Yote', totalAll],
    ['Idadi ya Rekodi', countAll],
    ['', ''],
    ['AINA YA MATUMIZI', 'JUMLA (TZS)'],
    ...Object.entries(byAina).sort((a, b) => b[1] - a[1]),
  ];

  sumSheet.getRange(1, 1, summaryData.length, 2).setValues(summaryData);

  // Pantikia kichwa
  sumSheet.getRange(1, 1, 1, 2).merge()
    .setBackground('#1b4332').setFontColor('#f4c945')
    .setFontWeight('bold').setFontSize(12)
    .setHorizontalAlignment('center');
}


// ═══════════════════════════════════════════════════════════════════
//  HELPER FUNCTIONS — Msaada wa Kazi
// ═══════════════════════════════════════════════════════════════════

function rowToObject(r) {
  return {
    tarehe:    String(r[0]).substring(0, 10),
    kumbukumbu: r[1],
    mtumiaji:   r[2],
    aina:       r[3],
    bidhaa:     r[4],
    kiasi:      r[5],
    kitengo:    r[6],
    bei:        r[7],
    jumla:      r[8],
    chanzo:     r[9],
    muuzaji:    r[10],
    idhini:     r[11],
    risiti:     r[12],
    maelezo:    r[13],
    timestamp:  r[14],
  };
}

function generateRef(sheet) {
  const lastRow = sheet.getLastRow();
  const num     = String(lastRow).padStart(4, '0');
  const date    = todayStr().replace(/-/g, '').substring(2);
  return `MAT-${date}-${num}`;
}

function referenceExists(sheet, ref) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  const refs = sheet.getRange(2, 2, lastRow - 1, 1).getValues().flat();
  return refs.includes(ref);
}

function findRowByRef(sheet, ref) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const refs = sheet.getRange(2, 2, lastRow - 1, 1).getValues().flat();
  const idx  = refs.indexOf(ref);
  return idx === -1 ? -1 : idx + 2;
}

function sortSheetByDate(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 2) return;
  sheet.getRange(2, 1, lastRow - 1, HEADERS_MATUMIZI.length).sort({ column: 1, ascending: true });
}

function sanitize(str) {
  if (str === null || str === undefined) return '';
  return String(str).replace(/[<>\"']/g, '').trim().substring(0, 500);
}

function todayStr() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function getWeekStart() {
  const d = new Date();
  d.setDate(d.getDate() - d.getDay());
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function emptySummary() {
  return { totalAll: 0, totalToday: 0, totalWeek: 0, totalMonth: 0, countAll: 0, countToday: 0, byCategory: {} };
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


// ═══════════════════════════════════════════════════════════════════
//  LOGGING — Kumbukumbu ya Mfumo
// ═══════════════════════════════════════════════════════════════════

function logActivity(action, ref, user) {
  try {
    const ss    = getOrCreateSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_LOG) || ss.insertSheet(SHEET_LOG);
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, 5).setValues([['Wakati', 'Kitendo', 'Kumbukumbu', 'Mtumiaji', 'Maelezo']]);
      sheet.getRange(1, 1, 1, 5).setBackground('#1b4332').setFontColor('#f4c945').setFontWeight('bold');
    }
    sheet.appendRow([new Date().toISOString(), action, ref, user, 'OK']);
  } catch (e) {
    // Usiruhusu hitilafu ya log kuzuia mtiririko mkuu
  }
}

function logError(context, err) {
  try {
    const ss    = getOrCreateSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_LOG) || ss.insertSheet(SHEET_LOG);
    sheet.appendRow([new Date().toISOString(), 'ERROR', context, err.message, err.stack || '']);
  } catch (e) {
    // Kimya
  }
}


// ═══════════════════════════════════════════════════════════════════
//  SETUP TRIGGER — Kazi ya Kusanidi (Endesha Mara Moja)
// ═══════════════════════════════════════════════════════════════════

/**
 * Endesha kazi hii mara moja kwenye Apps Script kwa kubonyeza Run → initializeSystem
 * kabla ya kufanya deploy. Itaunda laha zote na muundo sahihi.
 */
function initializeSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.rename('UKM Kituo cha Mifugo — Matumizi ya Duka');

  getOrCreateSheet(ss, SHEET_MATUMIZI, HEADERS_MATUMIZI);
  getOrCreateSheet(ss, SHEET_SUMMARY, []);
  getOrCreateSheet(ss, SHEET_LOG, []);

  updateSummarySheet(ss);

  SpreadsheetApp.getUi().alert(
    '✅ Mfumo Umesanidiwa!\n\n' +
    'Laha zifuatazo zimeundwa:\n' +
    '• Matumizi\n• Muhtasari\n• Log ya Mfumo\n\n' +
    'Sasa fanya Deploy → New Deployment kama Web App.'
  );
}


// ═══════════════════════════════════════════════════════════════════
//  MENU — Menyu ya Custom kwenye Google Sheets
// ═══════════════════════════════════════════════════════════════════

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🐾 UKM Mfumo')
    .addItem('Sanidi Mfumo (Mara ya Kwanza)', 'initializeSystem')
    .addItem('Sasisha Muhtasari', 'refreshSummary')
    .addItem('Angalia Rekodi za Leo', 'showTodayRecords')
    .addSeparator()
    .addItem('Hamisha CSV (Mwezi Huu)', 'exportCurrentMonth')
    .addToUi();
}

function refreshSummary() {
  updateSummarySheet(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert('Muhtasari umesasishwa! ✅');
}

function showTodayRecords() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_MATUMIZI);
  if (!sheet || sheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('Hakuna rekodi za leo.');
    return;
  }
  const today = todayStr();
  const data  = sheet.getRange(2, 1, sheet.getLastRow()-1, HEADERS_MATUMIZI.length).getValues();
  const todayRows = data.filter(r => String(r[0]).substring(0,10) === today);
  const total = todayRows.reduce((s, r) => s + (Number(r[8]) || 0), 0);

  SpreadsheetApp.getUi().alert(
    `📊 Rekodi za Leo (${today})\n\n` +
    `Idadi: ${todayRows.length}\n` +
    `Jumla: TZS ${total.toLocaleString()}`
  );
}

function exportCurrentMonth() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const sheet  = ss.getSheetByName(SHEET_MATUMIZI);
  if (!sheet || sheet.getLastRow() < 2) return;

  const month  = todayStr().substring(0, 7);
  const allData = sheet.getRange(2, 1, sheet.getLastRow()-1, HEADERS_MATUMIZI.length).getValues();
  const monthRows = allData.filter(r => String(r[0]).substring(0,7) === month);

  if (!monthRows.length) {
    SpreadsheetApp.getUi().alert('Hakuna rekodi za mwezi huu.');
    return;
  }

  // Unda laha mpya ya muda
  const exportSheet = ss.insertSheet('Export_' + month);
  exportSheet.getRange(1, 1, 1, HEADERS_MATUMIZI.length).setValues([HEADERS_MATUMIZI]);
  exportSheet.getRange(2, 1, monthRows.length, HEADERS_MATUMIZI.length).setValues(monthRows);

  SpreadsheetApp.getUi().alert(
    `✅ Laha "Export_${month}" imeundwa na rekodi ${monthRows.length}.\n` +
    'Unaweza kuisahihisha au kuipakua kama CSV.'
  );
}
