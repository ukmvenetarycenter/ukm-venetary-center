/**
 * ══════════════════════════════════════════════════════════════════
 *  UKM Kituo cha Mifugo — Madeni API
 *  Google Apps Script (Web App)
 *
 *  JINSI YA KUTUMIA:
 *  1. Fungua script.google.com → New Project
 *  2. Copy code hii yote → ubadilishe content ya Code.gs
 *  3. Bonyeza Deploy → New Deployment
 *     • Type: Web App
 *     • Execute as: Me
 *     • Who has access: Anyone
 *  4. Copy Web App URL → iweke kwenye madeni.html API field
 * ══════════════════════════════════════════════════════════════════
 */

// ─── CONFIGURATION ────────────────────────────────────────────────
const SHEET_NAME   = 'Madeni';      // Jina la sheet
const SHEET_ID     = '';            // (Hiari) ID ya Spreadsheet maalum
                                    // Kama wazi, itatumia spreadsheet iliyounganishwa na script

// ─── COLUMN MAP ────────────────────────────────────────────────────
// Nafasi za safu (columns) katika sheet (1-based)
const COL = {
  id:           1,   // A — ID ya kipekee (auto-generated)
  jina:         2,   // B — Jina la Mteja
  simu:         3,   // C — Namba ya Simu
  anuani:       4,   // D — Anuani / Mtaa
  bidhaa:       5,   // E — Bidhaa / Huduma
  jumla:        6,   // F — Jumla ya Deni
  kilicholipwa: 7,   // G — Kilicholipwa
  kilichobaki:  8,   // H — Kinachobaki (calculated)
  tarehe:       9,   // I — Tarehe ya Deni
  deadline:     10,  // J — Tarehe ya Mwisho
  hali:         11,  // K — Hali ya Malipo
  maelezo:      12,  // L — Maelezo / Kumbukumbu
  iliyoandikwa: 13,  // M — Timestamp ya kuandika
  iliyosasishwa:14,  // N — Timestamp ya kusasishwa
};
const TOTAL_COLS = 14;


// ══════════════════════════════════════════════════════════════════
//  GET — Handle GET requests
// ══════════════════════════════════════════════════════════════════
function doGet(e) {
  try {
    const action = e.parameter.action || 'getMadeni';

    if (action === 'getMadeni') {
      return jsonResponse(getMadeni());
    }

    if (action === 'getSummary') {
      return jsonResponse(getSummary());
    }

    return jsonResponse({ success: false, message: 'Hatua isiyojulikana: ' + action });

  } catch (err) {
    return jsonResponse({ success: false, message: err.message });
  }
}


// ══════════════════════════════════════════════════════════════════
//  POST — Handle POST requests
// ══════════════════════════════════════════════════════════════════
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action  = payload.action;

    if (action === 'addMadeni') {
      return jsonResponse(addMadeni(payload));
    }

    if (action === 'updateMadeni') {
      return jsonResponse(updateMadeni(payload));
    }

    if (action === 'deleteMadeni') {
      return jsonResponse(deleteMadeni(payload.id));
    }

    return jsonResponse({ success: false, message: 'Hatua isiyojulikana: ' + action });

  } catch (err) {
    return jsonResponse({ success: false, message: err.message });
  }
}


// ══════════════════════════════════════════════════════════════════
//  GET ALL MADENI
// ══════════════════════════════════════════════════════════════════
function getMadeni() {
  const sheet = getSheet();
  const data  = sheet.getDataRange().getValues();

  if (data.length <= 1) {
    return { success: true, data: [], total: 0 };
  }

  // Skip header row (row 0)
  const records = data.slice(1)
    .filter(row => row[COL.id - 1] !== '')   // skip empty rows
    .map(row => rowToObject(row));

  // Sort by date descending (newest first)
  records.sort((a, b) => new Date(b.tarehe) - new Date(a.tarehe));

  return { success: true, data: records, total: records.length };
}


// ══════════════════════════════════════════════════════════════════
//  ADD MADENI
// ══════════════════════════════════════════════════════════════════
function addMadeni(payload) {
  validatePayload(payload);

  const sheet     = getSheet();
  const id        = generateId();
  const now       = new Date();
  const jumla     = parseFloat(payload.jumla)        || 0;
  const kilipwa   = parseFloat(payload.kilicholipwa) || 0;
  const kinabaki  = jumla - kilipwa;

  const newRow = new Array(TOTAL_COLS).fill('');
  newRow[COL.id           - 1] = id;
  newRow[COL.jina         - 1] = payload.jina.trim();
  newRow[COL.simu         - 1] = payload.simu.trim();
  newRow[COL.anuani       - 1] = (payload.anuani  || '').trim();
  newRow[COL.bidhaa       - 1] = payload.bidhaa.trim();
  newRow[COL.jumla        - 1] = jumla;
  newRow[COL.kilicholipwa - 1] = kilipwa;
  newRow[COL.kilichobaki  - 1] = kinabaki;
  newRow[COL.tarehe       - 1] = payload.tarehe;
  newRow[COL.deadline     - 1] = payload.deadline  || '';
  newRow[COL.hali         - 1] = payload.hali      || determineHali(jumla, kilipwa);
  newRow[COL.maelezo      - 1] = (payload.maelezo  || '').trim();
  newRow[COL.iliyoandikwa - 1] = Utilities.formatDate(now, 'Africa/Nairobi', 'yyyy-MM-dd HH:mm:ss');
  newRow[COL.iliyosasishwa- 1] = '';

  sheet.appendRow(newRow);

  // Auto-format the new row
  formatSheet(sheet);

  return {
    success: true,
    message: 'Deni limeongezwa kikamilifu.',
    id: id,
    data: rowToObject(newRow)
  };
}


// ══════════════════════════════════════════════════════════════════
//  UPDATE MADENI
// ══════════════════════════════════════════════════════════════════
function updateMadeni(payload) {
  if (!payload.id) throw new Error('ID inahitajika kusasisha rekodi.');
  validatePayload(payload);

  const sheet = getSheet();
  const rowIndex = findRowById(sheet, payload.id);

  if (!rowIndex) {
    return { success: false, message: 'Rekodi yenye ID "' + payload.id + '" haikupatikana.' };
  }

  const now      = new Date();
  const jumla    = parseFloat(payload.jumla)        || 0;
  const kilipwa  = parseFloat(payload.kilicholipwa) || 0;
  const kinabaki = jumla - kilipwa;

  // Update individual cells
  sheet.getRange(rowIndex, COL.jina        ).setValue(payload.jina.trim());
  sheet.getRange(rowIndex, COL.simu        ).setValue(payload.simu.trim());
  sheet.getRange(rowIndex, COL.anuani      ).setValue((payload.anuani || '').trim());
  sheet.getRange(rowIndex, COL.bidhaa      ).setValue(payload.bidhaa.trim());
  sheet.getRange(rowIndex, COL.jumla       ).setValue(jumla);
  sheet.getRange(rowIndex, COL.kilicholipwa).setValue(kilipwa);
  sheet.getRange(rowIndex, COL.kilichobaki ).setValue(kinabaki);
  sheet.getRange(rowIndex, COL.tarehe      ).setValue(payload.tarehe);
  sheet.getRange(rowIndex, COL.deadline    ).setValue(payload.deadline || '');
  sheet.getRange(rowIndex, COL.hali        ).setValue(payload.hali || determineHali(jumla, kilipwa));
  sheet.getRange(rowIndex, COL.maelezo     ).setValue((payload.maelezo || '').trim());
  sheet.getRange(rowIndex, COL.iliyosasishwa).setValue(
    Utilities.formatDate(now, 'Africa/Nairobi', 'yyyy-MM-dd HH:mm:ss')
  );

  return {
    success: true,
    message: 'Rekodi imesasishwa kikamilifu.',
    id: payload.id
  };
}


// ══════════════════════════════════════════════════════════════════
//  DELETE MADENI
// ══════════════════════════════════════════════════════════════════
function deleteMadeni(id) {
  if (!id) throw new Error('ID inahitajika kufuta rekodi.');

  const sheet    = getSheet();
  const rowIndex = findRowById(sheet, id);

  if (!rowIndex) {
    return { success: false, message: 'Rekodi yenye ID "' + id + '" haikupatikana.' };
  }

  sheet.deleteRow(rowIndex);

  return {
    success: true,
    message: 'Rekodi imefutwa kikamilifu.',
    id: id
  };
}


// ══════════════════════════════════════════════════════════════════
//  GET SUMMARY STATISTICS
// ══════════════════════════════════════════════════════════════════
function getSummary() {
  const result = getMadeni();
  if (!result.success) return result;

  const records = result.data;
  const summary = {
    jumla_wateja:  records.length,
    jumla_deni:    0,
    haijalipiwa:   { idadi: 0, kiasi: 0 },
    sehemu:        { idadi: 0, kiasi: 0 },
    imelipwa:      { idadi: 0, kiasi: 0 },
  };

  records.forEach(r => {
    const jumla   = parseFloat(r.jumla)        || 0;
    const kilipwa = parseFloat(r.kilicholipwa) || 0;
    const baki    = jumla - kilipwa;

    summary.jumla_deni += jumla;

    if (r.hali === 'Imelipwa') {
      summary.imelipwa.idadi++;
      summary.imelipwa.kiasi += jumla;
    } else if (r.hali === 'Imelipwa Sehemu') {
      summary.sehemu.idadi++;
      summary.sehemu.kiasi += baki;
    } else {
      summary.haijalipiwa.idadi++;
      summary.haijalipiwa.kiasi += baki;
    }
  });

  return { success: true, summary };
}


// ══════════════════════════════════════════════════════════════════
//  SHEET HELPER FUNCTIONS
// ══════════════════════════════════════════════════════════════════

/**
 * Get or create the Madeni sheet with proper headers
 */
function getSheet() {
  let ss;
  if (SHEET_ID && SHEET_ID !== '') {
    ss = SpreadsheetApp.openById(SHEET_ID);
  } else {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }

  let sheet = ss.getSheetByName(SHEET_NAME);

  // Create sheet and headers if it doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    setupHeaders(sheet);
  }

  return sheet;
}


/**
 * Create professional headers for the sheet
 */
function setupHeaders(sheet) {
  const headers = [
    'ID',
    'Jina la Mteja',
    'Namba ya Simu',
    'Anuani / Mtaa',
    'Bidhaa / Huduma',
    'Jumla ya Deni (TZS)',
    'Kilicholipwa (TZS)',
    'Kinachobaki (TZS)',
    'Tarehe ya Deni',
    'Tarehe ya Mwisho',
    'Hali ya Malipo',
    'Maelezo / Kumbukumbu',
    'Tarehe Iliyoandikwa',
    'Tarehe Iliyosasishwa'
  ];

  const headerRange = sheet.getRange(1, 1, 1, TOTAL_COLS);
  headerRange.setValues([headers]);

  // Style headers
  headerRange
    .setBackground('#1b4332')
    .setFontColor('#f4c945')
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);

  sheet.setRowHeight(1, 42);

  // Set column widths
  sheet.setColumnWidth(COL.id,            100);
  sheet.setColumnWidth(COL.jina,          180);
  sheet.setColumnWidth(COL.simu,          130);
  sheet.setColumnWidth(COL.anuani,        150);
  sheet.setColumnWidth(COL.bidhaa,        200);
  sheet.setColumnWidth(COL.jumla,         130);
  sheet.setColumnWidth(COL.kilicholipwa,  130);
  sheet.setColumnWidth(COL.kilichobaki,   130);
  sheet.setColumnWidth(COL.tarehe,        110);
  sheet.setColumnWidth(COL.deadline,      110);
  sheet.setColumnWidth(COL.hali,          130);
  sheet.setColumnWidth(COL.maelezo,       200);
  sheet.setColumnWidth(COL.iliyoandikwa,  160);
  sheet.setColumnWidth(COL.iliyosasishwa, 160);

  // Freeze header row
  sheet.setFrozenRows(1);

  // Add data validation for Hali column
  const haliRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Haijalipiwa', 'Imelipwa Sehemu', 'Imelipwa'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, COL.hali, sheet.getMaxRows() - 1, 1).setDataValidation(haliRule);
}


/**
 * Apply conditional formatting & number formats to data rows
 */
function formatSheet(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // Format currency columns
  const currencyCols = [COL.jumla, COL.kilicholipwa, COL.kilichobaki];
  currencyCols.forEach(c => {
    sheet.getRange(2, c, lastRow - 1, 1)
      .setNumberFormat('#,##0');
  });

  // Hali column conditional colors
  const haliRange = sheet.getRange(2, COL.hali, lastRow - 1, 1);
  const rules = sheet.getConditionalFormatRules();

  // Clear old conditional format rules for hali column to avoid duplicates
  const filteredRules = rules.filter(r => {
    const ranges = r.getRanges();
    return !ranges.some(range =>
      range.getColumn() === COL.hali && range.getRow() > 1
    );
  });

  const paidRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Imelipwa')
    .setBackground('#e8f5ee')
    .setFontColor('#1b8a4b')
    .setBold(true)
    .setRanges([haliRange])
    .build();

  const partialRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Imelipwa Sehemu')
    .setBackground('#fef3e2')
    .setFontColor('#e67e22')
    .setBold(true)
    .setRanges([haliRange])
    .build();

  const unpaidRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Haijalipiwa')
    .setBackground('#fdecea')
    .setFontColor('#c0392b')
    .setBold(true)
    .setRanges([haliRange])
    .build();

  filteredRules.push(paidRule, partialRule, unpaidRule);
  sheet.setConditionalFormatRules(filteredRules);

  // Alternate row colors for readability
  for (let i = 2; i <= lastRow; i++) {
    const rowRange = sheet.getRange(i, 1, 1, TOTAL_COLS);
    if (i % 2 === 0) {
      rowRange.setBackground('#fdf8f0');
    } else {
      rowRange.setBackground('#ffffff');
    }
  }

  // Bold the "Kinachobaki" column if > 0 (highlight outstanding balance)
  const bakiRange = sheet.getRange(2, COL.kilichobaki, lastRow - 1, 1);
  const values    = bakiRange.getValues();
  values.forEach((row, i) => {
    const cell = sheet.getRange(i + 2, COL.kilichobaki);
    if (parseFloat(row[0]) > 0) {
      cell.setFontColor('#c0392b').setFontWeight('bold');
    } else {
      cell.setFontColor('#1b8a4b').setFontWeight('normal');
    }
  });
}


// ══════════════════════════════════════════════════════════════════
//  UTILITY FUNCTIONS
// ══════════════════════════════════════════════════════════════════

function rowToObject(row) {
  return {
    id:           String(row[COL.id           - 1] || ''),
    jina:         String(row[COL.jina         - 1] || ''),
    simu:         String(row[COL.simu         - 1] || ''),
    anuani:       String(row[COL.anuani       - 1] || ''),
    bidhaa:       String(row[COL.bidhaa       - 1] || ''),
    jumla:        parseFloat(row[COL.jumla        - 1]) || 0,
    kilicholipwa: parseFloat(row[COL.kilicholipwa - 1]) || 0,
    kilichobaki:  parseFloat(row[COL.kilichobaki  - 1]) || 0,
    tarehe:       formatCellDate(row[COL.tarehe    - 1]),
    deadline:     formatCellDate(row[COL.deadline  - 1]),
    hali:         String(row[COL.hali         - 1] || 'Haijalipiwa'),
    maelezo:      String(row[COL.maelezo      - 1] || ''),
    iliyoandikwa: String(row[COL.iliyoandikwa - 1] || ''),
    iliyosasishwa:String(row[COL.iliyosasishwa- 1] || ''),
  };
}

function formatCellDate(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, 'Africa/Nairobi', 'yyyy-MM-dd');
  }
  return String(val);
}

function findRowById(sheet, id) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][COL.id - 1]) === String(id)) {
      return i + 1; // 1-based row index
    }
  }
  return null;
}

function generateId() {
  const now = new Date();
  const ts  = Utilities.formatDate(now, 'Africa/Nairobi', 'yyyyMMddHHmmss');
  const rnd = Math.floor(Math.random() * 1000).toString().padStart(3, '0');
  return 'UKM-' + ts + '-' + rnd;
}

function determineHali(jumla, kilipwa) {
  if (kilipwa <= 0)        return 'Haijalipiwa';
  if (kilipwa >= jumla)    return 'Imelipwa';
  return 'Imelipwa Sehemu';
}

function validatePayload(p) {
  if (!p.jina   || !p.jina.trim())   throw new Error('Jina la mteja linahitajika.');
  if (!p.simu   || !p.simu.trim())   throw new Error('Namba ya simu inahitajika.');
  if (!p.bidhaa || !p.bidhaa.trim()) throw new Error('Jina la bidhaa / huduma linahitajika.');
  if (!p.jumla  || isNaN(+p.jumla))  throw new Error('Jumla ya deni inahitajika (namba).');
  if (!p.tarehe)                     throw new Error('Tarehe ya deni inahitajika.');
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}


// ══════════════════════════════════════════════════════════════════
//  MANUAL SETUP (run once from Script Editor to initialize)
// ══════════════════════════════════════════════════════════════════

/**
 * Run this function ONCE from the Apps Script editor
 * to set up the sheet with proper headers and formatting.
 * Menu: Run → initializeSheet
 */
function initializeSheet() {
  const sheet = getSheet();
  setupHeaders(sheet);
  SpreadsheetApp.getUi().alert(
    '✅ Sheet "' + SHEET_NAME + '" imewekwa kikamilifu!\n\n' +
    'Sasa unaweza:\n' +
    '1. Deploy script kama Web App\n' +
    '2. Copy URL → iweke kwenye madeni.html'
  );
}

/**
 * Add sample test data — useful for development
 * Run from Apps Script editor: Run → addSampleData
 */
function addSampleData() {
  const samples = [
    { jina:'Hamisi Mwalimu',  simu:'0744111222', anuani:'Nida, Songea',    bidhaa:'Dawa ya minyoo x5, Vitamin B12',    jumla:85000,  kilicholipwa:0,     tarehe:'2025-04-10', hali:'Haijalipiwa',    maelezo:'Mteja wa kawaida' },
    { jina:'Fatuma Saidi',    simu:'0755333444', anuani:'Mji Mpya, Songea', bidhaa:'Chanjo ya FMD ng\'ombe 8',         jumla:120000, kilicholipwa:60000, tarehe:'2025-04-15', hali:'Imelipwa Sehemu', maelezo:'Atalipa mwezi ujao' },
    { jina:'John Kapinga',    simu:'0622555666', anuani:'Peramiho',         bidhaa:'Pembejeo mbegu, Mbolea NPK',       jumla:200000, kilicholipwa:200000,tarehe:'2025-03-20', hali:'Imelipwa',        maelezo:'' },
    { jina:'Mariam Abdallah', simu:'0688777888', anuani:'Ruhekei',          bidhaa:'Dawa ya kupe, Betamec pour-on',    jumla:45000,  kilicholipwa:20000, tarehe:'2025-04-20', hali:'Imelipwa Sehemu', maelezo:'Baki italipwa Ijumaa' },
    { jina:'Paulo Mwasongwe', simu:'0712999000', anuani:'Matogoro',         bidhaa:'Matibabu dharura ng\'ombe 2, dawa',jumla:300000, kilicholipwa:0,     tarehe:'2025-04-25', hali:'Haijalipiwa',    maelezo:'Ameahidi kulipa mwezi wa 5' },
  ];

  const sheet = getSheet();
  const now   = Utilities.formatDate(new Date(), 'Africa/Nairobi', 'yyyy-MM-dd HH:mm:ss');

  samples.forEach(s => {
    const id      = generateId();
    const baki    = s.jumla - s.kilicholipwa;
    const row     = new Array(TOTAL_COLS).fill('');
    row[COL.id           - 1] = id;
    row[COL.jina         - 1] = s.jina;
    row[COL.simu         - 1] = s.simu;
    row[COL.anuani       - 1] = s.anuani;
    row[COL.bidhaa       - 1] = s.bidhaa;
    row[COL.jumla        - 1] = s.jumla;
    row[COL.kilicholipwa - 1] = s.kilicholipwa;
    row[COL.kilichobaki  - 1] = baki;
    row[COL.tarehe       - 1] = s.tarehe;
    row[COL.deadline     - 1] = '';
    row[COL.hali         - 1] = s.hali;
    row[COL.maelezo      - 1] = s.maelezo;
    row[COL.iliyoandikwa - 1] = now;
    row[COL.iliyosasishwa- 1] = '';
    sheet.appendRow(row);
    Utilities.sleep(100);
  });

  formatSheet(sheet);
  SpreadsheetApp.getUi().alert('✅ Data ya mfano 5 imeongezwa!');
}

