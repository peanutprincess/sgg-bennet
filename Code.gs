// ─────────────────────────────────────────
// BENNET SCHLESINGER — Commissions & Inquiries
// Separate Apps Script deployment
// ─────────────────────────────────────────

const BENNET_VERSION    = '1.9';

const COMMISSIONS_SS_ID = '1j3dymNa0u_ajdAtUZax8sQAkEQJpi1BBhOJWMpTcBJ4';
const INQUIRIES_SS_ID   = '1nx30zFjY8TFplbI7CTAgtusLNbUDdZ0tPltvqJ8F6bM';
const COMMISSIONS_SHEET = 'Comissions'; // note: one 'm'
const INQUIRIES_SHEET   = 'Sheet1';
const INVENTORY_SHEET   = 'Inventory';

// Column indices (0-indexed)
const COL = {
  CLIENT:          0,  // A
  IMAGE:           1,  // B
  TITLE:           2,  // C
  YEAR:            3,  // D
  DESCRIPTION:     4,  // E — SKU
  DIMENSIONS:      5,  // F
  MATERIAL:        6,  // G
  SG_INVOICE:      7,  // H
  PRICE:           8,  // I
  SALE_PRICE:      9,  // J
  BENNET_PAYOUT:  10,  // K
  ARTIST_PAID:    11,  // L — Bennet Payment (status)
  PAYOUT_RECORDS: 12,  // M — Bennet Payout Records
  GALLERY_PAID:   13,  // N — Gallery Payment (status)
  PROD_STATUS:    14,  // O
  // P (15) = smiley/done indicator — not read or written
  DUE_DATE:       16,  // Q
  DAYS_TIL_DUE:   17,  // R — formula column
  NOTES:          18,  // S
  SHIP_ADDRESS:   19,  // T
  SHIPPER:        20,  // U
  SHIP_REF:       21,  // V
  FILE_REF:       22,  // W
  TRACKING:       23,  // X
  TRACKING_SENT:  24,  // Y
  POC:            25,  // Z
};

// Column numbers for setValue (1-indexed)
const WRITE_COL = {
  client:         1,
  title:          3,
  year:           4,
  description:    5,
  dimensions:     6,
  material:       7,
  sgInvoice:      8,
  price:          9,
  salePrice:      10,
  bennetPayout:   11,
  artistPaid:     12,
  payoutRecords:  13,
  galleryPaid:    14,
  status:         15,
  dueDate:        17,
  notes:          19,
  shipAddress:    20,
  shipper:        21,
  shipRef:        22,
  fileRef:        23,
  tracking:       24,
  trackingSent:   25,
  poc:            26,
};

// ─────────────────────────────────────────
// ─────────────────────────────────────────
// POST — log a payable against a commission row
// ─────────────────────────────────────────
function logPayable(data) {
  try {
    const ss    = SpreadsheetApp.openById(COMMISSIONS_SS_ID);
    const sheet = ss.getSheetByName(COMMISSIONS_SHEET);
    if (!sheet) throw new Error('Sheet not found');

    const d = data.data || {};
    const bsc = String(d.bscNumber || '').trim();
    if (!bsc) throw new Error('BSC number is required');

    // ── Find matching row by SKU (description column) ──
    const rows = sheet.getDataRange().getValues();
    let targetRow = -1;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][COL.DESCRIPTION] || '').trim() === bsc) {
        targetRow = i + 1; // 1-indexed
        break;
      }
    }
    if (targetRow === -1) throw new Error('SKU not found in sheet: ' + bsc);

    // ── Upload PDF to Drive ──
    let driveLink = '';
    if (d.pdfBase64 && d.filename) {
      const folder = _getOrCreateFolder('Bennet Payables');
      const blob   = Utilities.newBlob(
        Utilities.base64Decode(d.pdfBase64), 'application/pdf', d.filename
      );
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      driveLink = file.getUrl();
    }

    // ── Update Column K (BENNET_PAYOUT) — add to existing ──
    const existingPayout = String(rows[targetRow - 1][COL.BENNET_PAYOUT] || '').trim();
    const newAmt = parseFloat(String(d.amount || '0').replace(/[^0-9.]/g, '')) || 0;
    const oldAmt = parseFloat(existingPayout.replace(/[^0-9.]/g, '')) || 0;
    const totalAmt = oldAmt + newAmt;
    sheet.getRange(targetRow, WRITE_COL.bennetPayout).setValue(totalAmt.toFixed(2));

    // ── Update Column M (PAYOUT_RECORDS) — append entry ──
    const existingRecords = String(rows[targetRow - 1][COL.PAYOUT_RECORDS] || '').trim();
    const linkPart  = driveLink ? driveLink : '(no PDF)';
    const notesPart = d.notes   ? d.notes   : '';
    const newEntry  = [d.payNumber, d.date, linkPart, notesPart].filter(Boolean).join(' | ');
    const updated   = existingRecords ? existingRecords + '\n' + newEntry : newEntry;
    sheet.getRange(targetRow, WRITE_COL.payoutRecords).setValue(updated);

    SpreadsheetApp.flush();
    return ContentService
      .createTextOutput(JSON.stringify({ success: true, driveLink: driveLink, total: totalAmt }))
      .setMimeType(ContentService.MimeType.TEXT);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.TEXT);
  }
}

function _getOrCreateFolder(name) {
  const folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}

// ONE-TIME: insert "Bennet Payout Records" column at M (run once from editor)
function insertPayoutRecordsColumn() {
  const ss    = SpreadsheetApp.openById(COMMISSIONS_SS_ID);
  const sheet = ss.getSheetByName(COMMISSIONS_SHEET);
  sheet.insertColumnBefore(13); // inserts before current col M → becomes new M
  sheet.getRange(1, 13).setValue('Bennet Payout Records');
  SpreadsheetApp.flush();
  Logger.log('Column M inserted: Bennet Payout Records');
}

// ONE-TIME: rename Description → SKU in header row (run once, then delete)
function renameDescriptionToSKU() {
  const ss    = SpreadsheetApp.openById(COMMISSIONS_SS_ID);
  const sheet = ss.getSheetByName(COMMISSIONS_SHEET);
  const row1  = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  for (let i = 0; i < row1.length; i++) {
    if (String(row1[i]).trim().toLowerCase() === 'description') {
      sheet.getRange(1, i + 1).setValue('SKU');
      SpreadsheetApp.flush();
      Logger.log('Renamed column ' + (i + 1) + ' → SKU');
      return;
    }
  }
  Logger.log('Description column not found');
}

// ─────────────────────────────────────────
function doGet(e) {
  const action = e && e.parameter && e.parameter.action;
  if (action === 'commissions')     return serveCommissions();
  if (action === 'inquiries')       return serveInquiries();
  if (action === 'tristan')         return serveTristandEntries();
  if (action === 'inventory-sales') return serveInventorySales();
  return ContentService
    .createTextOutput(JSON.stringify({ error: 'Unknown action' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.type === 'update_commission')    return updateCommission(data);
    if (data.type === 'create_commission')    return createCommission(data);
    if (data.type === 'create_inquiry')       return createInquiry(data);
    if (data.type === 'update_inquiry')       return updateInquiry(data);
    if (data.type === 'log_payable')          return logPayable(data);
    if (data.type === 'create_tristan')       return createTristandEntry(data);
    if (data.type === 'create_inventory_sale') return createInventorySale(data);
    if (data.type === 'update_inventory_sale') return updateInventorySale(data);
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: 'Unknown type' }))
      .setMimeType(ContentService.MimeType.TEXT);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.TEXT);
  }
}

// ─────────────────────────────────────────
// GET commissions
// ─────────────────────────────────────────
function serveCommissions() {
  try {
    const ss    = SpreadsheetApp.openById(COMMISSIONS_SS_ID);
    const sheet = ss.getSheetByName(COMMISSIONS_SHEET);
    if (!sheet) throw new Error('Sheet "' + COMMISSIONS_SHEET + '" not found');

    const data = sheet.getDataRange().getValues();
    const commissions = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!String(row[COL.CLIENT] || '').trim() && !String(row[COL.TITLE] || '').trim()) continue;

      const dueVal = row[COL.DUE_DATE];
      const dueDate = dueVal instanceof Date
        ? Utilities.formatDate(dueVal, 'America/New_York', 'yyyy-MM-dd')
        : String(dueVal || '').trim();

      commissions.push({
        rowIndex:      i + 1,
        client:        String(row[COL.CLIENT]        || '').trim(),
        title:         String(row[COL.TITLE]         || '').trim(),
        year:          String(row[COL.YEAR]          || '').trim(),
        description:   String(row[COL.DESCRIPTION]   || '').trim(),
        dimensions:    String(row[COL.DIMENSIONS]    || '').trim(),
        material:      String(row[COL.MATERIAL]      || '').trim(),
        sgInvoice:     String(row[COL.SG_INVOICE]    || '').trim(),
        price:         String(row[COL.PRICE]         || '').trim(),
        salePrice:     String(row[COL.SALE_PRICE]    || '').trim(),
        bennetPayout:  String(row[COL.BENNET_PAYOUT] || '').trim(),
        artistPaid:    String(row[COL.ARTIST_PAID]    || '').trim(),
        payoutRecords: String(row[COL.PAYOUT_RECORDS] || '').trim(),
        galleryPaid:   String(row[COL.GALLERY_PAID]   || '').trim(),
        status:        String(row[COL.PROD_STATUS]   || '').trim(),
        dueDate:       dueDate,
        notes:         String(row[COL.NOTES]         || '').trim(),
        daysTilDue:    String(row[COL.DAYS_TIL_DUE]  || '').trim(),
        shipAddress:   String(row[COL.SHIP_ADDRESS]  || '').trim(),
        shipper:       String(row[COL.SHIPPER]       || '').trim(),
        shipRef:       String(row[COL.SHIP_REF]      || '').trim(),
        fileRef:       String(row[COL.FILE_REF]      || '').trim(),
        tracking:      String(row[COL.TRACKING]      || '').trim(),
        trackingSent:  String(row[COL.TRACKING_SENT] || '').trim(),
        poc:           String(row[COL.POC]           || '').trim(),
      });
    }

    return ContentService
      .createTextOutput(JSON.stringify({ commissions: commissions, version: BENNET_VERSION }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ─────────────────────────────────────────
// GET inquiries — last 10 rows, newest first
// ─────────────────────────────────────────
function serveInquiries() {
  try {
    const ss    = SpreadsheetApp.openById(INQUIRIES_SS_ID);
    const sheet = ss.getSheetByName(INQUIRIES_SHEET);
    if (!sheet) throw new Error('Sheet "' + INQUIRIES_SHEET + '" not found');

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return ContentService
      .createTextOutput(JSON.stringify({ inquiries: [] }))
      .setMimeType(ContentService.MimeType.JSON);

    // Find first non-empty row as the header (handles blank leading rows)
    let headerIdx = 0;
    for (let i = 0; i < data.length; i++) {
      if (data[i].some(c => String(c).trim() !== '')) { headerIdx = i; break; }
    }
    const headers  = data[headerIdx].map(h => String(h).trim());
    const dataRows = data.slice(headerIdx + 1);

    const inquiries = [];
    dataRows.forEach((row, i) => {
      if (!row.some(c => c !== '')) return; // skip blank rows
      const obj = { rowIndex: headerIdx + i + 2 }; // 1-indexed sheet row
      headers.forEach((h, j) => { if (h) obj[h] = String(row[j] || '').trim(); });
      inquiries.push(obj);
    });

    return ContentService
      .createTextOutput(JSON.stringify({ inquiries: inquiries }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ─────────────────────────────────────────
// POST — update commission row
// ─────────────────────────────────────────
function updateCommission(data) {
  try {
    const ss    = SpreadsheetApp.openById(COMMISSIONS_SS_ID);
    const sheet = ss.getSheetByName(COMMISSIONS_SHEET);
    if (!sheet) throw new Error('Sheet not found');

    const rowIndex = parseInt(data.rowIndex);
    if (!rowIndex || rowIndex < 2) throw new Error('Invalid row index: ' + rowIndex);

    const updates = data.updates || {};
    Object.entries(updates).forEach(([field, value]) => {
      const col = WRITE_COL[field];
      if (col) sheet.getRange(rowIndex, col).setValue(value);
    });

    SpreadsheetApp.flush();
    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.TEXT);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.TEXT);
  }
}

// ─────────────────────────────────────────
// POST — create new commission row
// ─────────────────────────────────────────
function createCommission(data) {
  try {
    const ss    = SpreadsheetApp.openById(COMMISSIONS_SS_ID);
    const sheet = ss.getSheetByName(COMMISSIONS_SHEET);
    if (!sheet) throw new Error('Sheet not found');

    const d = data.data || {};
    // Build a new row aligned to column order (A–Z); leave B (image) blank
    const newRow = new Array(26).fill('');
    newRow[COL.CLIENT]        = d.client        || '';
    newRow[COL.TITLE]         = d.title         || '';
    newRow[COL.YEAR]          = d.year          || '';
    newRow[COL.DESCRIPTION]   = d.description   || '';
    newRow[COL.DIMENSIONS]    = d.dimensions    || '';
    newRow[COL.MATERIAL]      = d.material      || '';
    newRow[COL.SG_INVOICE]    = d.sgInvoice     || '';
    newRow[COL.PRICE]         = d.price         || '';
    newRow[COL.SALE_PRICE]    = d.salePrice     || '';
    newRow[COL.BENNET_PAYOUT] = d.bennetPayout  || '';
    newRow[COL.ARTIST_PAID]   = d.artistPaid    || '';
    newRow[COL.GALLERY_PAID]  = d.galleryPaid   || '';
    newRow[COL.PROD_STATUS]   = d.status        || 'Inquiry';
    newRow[COL.DUE_DATE]      = d.dueDate       || '';
    newRow[COL.NOTES]         = d.notes         || '';
    newRow[COL.SHIP_ADDRESS]  = d.shipAddress   || '';
    newRow[COL.SHIPPER]       = d.shipper       || '';
    newRow[COL.SHIP_REF]      = d.shipRef       || '';
    newRow[COL.FILE_REF]      = d.fileRef       || '';
    newRow[COL.TRACKING]      = d.tracking      || '';
    newRow[COL.TRACKING_SENT] = d.trackingSent  || '';
    newRow[COL.POC]           = d.poc           || '';

    sheet.appendRow(newRow);
    SpreadsheetApp.flush();
    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.TEXT);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.TEXT);
  }
}

// ─────────────────────────────────────────
// POST — create new inquiry row
// ─────────────────────────────────────────
function createInquiry(data) {
  try {
    const ss    = SpreadsheetApp.openById(INQUIRIES_SS_ID);
    const sheet = ss.getSheetByName(INQUIRIES_SHEET);
    if (!sheet) throw new Error('Inquiries sheet not found');

    const d = data.data || {};
    // Find first non-empty row to use as headers (handles blank leading rows)
    const allData = sheet.getDataRange().getValues();
    let headerRow = allData[0] || [];
    for (let i = 0; i < allData.length; i++) {
      if (allData[i].some(c => String(c).trim() !== '')) { headerRow = allData[i]; break; }
    }
    const headers = headerRow.map(h => String(h).trim().toLowerCase());

    const newRow = new Array(headers.length).fill('');
    const set = (...keys) => {
      for (const k of keys) {
        const idx = headers.findIndex(h => h.includes(k));
        if (idx > -1) return idx;
      }
      return -1;
    };

    const apply = (val, ...keys) => {
      const idx = set(...keys);
      if (idx > -1) newRow[idx] = val || '';
    };

    apply(d.date    || Utilities.formatDate(new Date(), 'America/New_York', 'MM/dd/yyyy'), 'timestamp','date','time');
    apply(d.name,    'name','first','contact');
    apply(d.org,     'org','gallery','company','organization');
    apply(d.email,   'email','mail');
    apply(d.artwork, 'artwork','title','piece');
    apply(d.notes,   'notes','message','comment','inquiry');

    sheet.appendRow(newRow.length ? newRow : [d.date || '', d.name || '', d.org || '', d.email || '', d.artwork || '', d.notes || '']);
    SpreadsheetApp.flush();
    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.TEXT);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.TEXT);
  }
}

// ─────────────────────────────────────────
// POST — update existing inquiry row
// ─────────────────────────────────────────
function updateInquiry(data) {
  try {
    const ss    = SpreadsheetApp.openById(INQUIRIES_SS_ID);
    const sheet = ss.getSheetByName(INQUIRIES_SHEET);
    if (!sheet) throw new Error('Inquiries sheet not found');

    const rowIndex = parseInt(data.rowIndex);
    if (!rowIndex || rowIndex < 2) throw new Error('Invalid row index: ' + rowIndex);

    const d = data.updates || {};
    // Find headers to map field names to columns
    const allData = sheet.getDataRange().getValues();
    let headerRow = allData[0] || [];
    for (let i = 0; i < allData.length; i++) {
      if (allData[i].some(c => String(c).trim() !== '')) { headerRow = allData[i]; break; }
    }
    const headers = headerRow.map(h => String(h).trim().toLowerCase());

    const findCol = (...keys) => {
      for (const k of keys) {
        const idx = headers.findIndex(h => h.includes(k));
        if (idx > -1) return idx + 1; // 1-indexed for setValue
      }
      return -1;
    };

    const setVal = (val, ...keys) => {
      const col = findCol(...keys);
      if (col > 0 && val !== undefined) sheet.getRange(rowIndex, col).setValue(val || '');
    };

    setVal(d.date,    'timestamp', 'date', 'time');
    setVal(d.name,    'name', 'first', 'contact');
    setVal(d.org,     'org', 'gallery', 'company', 'organization');
    setVal(d.email,   'email', 'mail');
    setVal(d.artwork, 'artwork', 'title', 'piece');
    setVal(d.notes,   'notes', 'message', 'comment', 'inquiry');

    SpreadsheetApp.flush();
    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.TEXT);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.TEXT);
  }
}

// ─────────────────────────────────────────
// TRISTAN UNRAU INTEREST LOG
// ─────────────────────────────────────────
const TRISTAN_SHEET = 'Tristan Interests';

function serveTristandEntries() {
  try {
    const ss    = SpreadsheetApp.openById(INQUIRIES_SS_ID);
    let sheet   = ss.getSheetByName(TRISTAN_SHEET);
    if (!sheet) {
      sheet = ss.insertSheet(TRISTAN_SHEET);
      sheet.appendRow(['Timestamp', 'Name', 'Organization', 'Email', 'Artwork', 'Notes']);
    }
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return ContentService
      .createTextOutput(JSON.stringify({ entries: [] }))
      .setMimeType(ContentService.MimeType.JSON);

    const headers  = data[0].map(h => String(h).trim());
    const entries  = data.slice(1)
      .filter(row => row.some(c => c !== ''))
      .map(row => {
        const obj = {};
        headers.forEach((h, i) => { if (h) obj[h] = String(row[i] || '').trim(); });
        return obj;
      });

    return ContentService
      .createTextOutput(JSON.stringify({ entries: entries }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function createTristandEntry(data) {
  try {
    const ss    = SpreadsheetApp.openById(INQUIRIES_SS_ID);
    let sheet   = ss.getSheetByName(TRISTAN_SHEET);
    if (!sheet) {
      sheet = ss.insertSheet(TRISTAN_SHEET);
      sheet.appendRow(['Timestamp', 'Name', 'Organization', 'Email', 'Artwork', 'Notes']);
    }
    const ts = data.date || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    sheet.appendRow([ts, data.name || '', data.org || '', data.email || '', data.artwork || '', data.notes || '']);
    SpreadsheetApp.flush();
    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.TEXT);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.TEXT);
  }
}

// ─────────────────────────────────────────
// INVENTORY SALES
// ─────────────────────────────────────────
function serveInventorySales() {
  try {
    const ss    = SpreadsheetApp.openById(COMMISSIONS_SS_ID);
    const sheet = ss.getSheetByName(INVENTORY_SHEET);
    if (!sheet) return ContentService
      .createTextOutput(JSON.stringify({ sales: [] }))
      .setMimeType(ContentService.MimeType.JSON);

    const data = sheet.getDataRange().getValues();
    const sales = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!String(row[COL.CLIENT] || '').trim() && !String(row[COL.TITLE] || '').trim()) continue;

      const dueVal  = row[COL.DUE_DATE];
      const dueDate = dueVal instanceof Date
        ? Utilities.formatDate(dueVal, 'America/New_York', 'yyyy-MM-dd')
        : String(dueVal || '').trim();

      sales.push({
        rowIndex:      i + 1,
        client:        String(row[COL.CLIENT]         || '').trim(),
        title:         String(row[COL.TITLE]          || '').trim(),
        year:          String(row[COL.YEAR]           || '').trim(),
        description:   String(row[COL.DESCRIPTION]    || '').trim(),
        dimensions:    String(row[COL.DIMENSIONS]     || '').trim(),
        material:      String(row[COL.MATERIAL]       || '').trim(),
        sgInvoice:     String(row[COL.SG_INVOICE]     || '').trim(),
        price:         String(row[COL.PRICE]          || '').trim(),
        salePrice:     String(row[COL.SALE_PRICE]     || '').trim(),
        bennetPayout:  String(row[COL.BENNET_PAYOUT]  || '').trim(),
        artistPaid:    String(row[COL.ARTIST_PAID]    || '').trim(),
        payoutRecords: String(row[COL.PAYOUT_RECORDS] || '').trim(),
        galleryPaid:   String(row[COL.GALLERY_PAID]   || '').trim(),
        status:        String(row[COL.PROD_STATUS]    || '').trim(),
        dueDate:       dueDate,
        notes:         String(row[COL.NOTES]          || '').trim(),
        shipAddress:   String(row[COL.SHIP_ADDRESS]   || '').trim(),
        shipper:       String(row[COL.SHIPPER]        || '').trim(),
        shipRef:       String(row[COL.SHIP_REF]       || '').trim(),
        fileRef:       String(row[COL.FILE_REF]       || '').trim(),
        tracking:      String(row[COL.TRACKING]       || '').trim(),
        trackingSent:  String(row[COL.TRACKING_SENT]  || '').trim(),
        poc:           String(row[COL.POC]            || '').trim(),
      });
    }

    return ContentService
      .createTextOutput(JSON.stringify({ sales: sales }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function createInventorySale(data) {
  try {
    const ss    = SpreadsheetApp.openById(COMMISSIONS_SS_ID);
    const sheet = ss.getSheetByName(INVENTORY_SHEET);
    if (!sheet) throw new Error('Sheet "' + INVENTORY_SHEET + '" not found');

    const d      = data.data || {};
    const newRow = new Array(26).fill('');
    newRow[COL.CLIENT]        = d.client        || '';
    newRow[COL.TITLE]         = d.title         || '';
    newRow[COL.YEAR]          = d.year          || '';
    newRow[COL.DESCRIPTION]   = d.description   || '';
    newRow[COL.DIMENSIONS]    = d.dimensions    || '';
    newRow[COL.MATERIAL]      = d.material      || '';
    newRow[COL.SG_INVOICE]    = d.sgInvoice     || '';
    newRow[COL.PRICE]         = d.price         || '';
    newRow[COL.SALE_PRICE]    = d.salePrice     || '';
    newRow[COL.BENNET_PAYOUT] = d.bennetPayout  || '';
    newRow[COL.ARTIST_PAID]   = d.artistPaid    || '';
    newRow[COL.GALLERY_PAID]  = d.galleryPaid   || '';
    newRow[COL.PROD_STATUS]   = d.status        || '';
    newRow[COL.DUE_DATE]      = d.dueDate       || '';
    newRow[COL.NOTES]         = d.notes         || '';
    newRow[COL.SHIP_ADDRESS]  = d.shipAddress   || '';
    newRow[COL.SHIPPER]       = d.shipper       || '';
    newRow[COL.SHIP_REF]      = d.shipRef       || '';
    newRow[COL.FILE_REF]      = d.fileRef       || '';
    newRow[COL.TRACKING]      = d.tracking      || '';
    newRow[COL.TRACKING_SENT] = d.trackingSent  || '';
    newRow[COL.POC]           = d.poc           || '';

    sheet.appendRow(newRow);
    SpreadsheetApp.flush();
    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.TEXT);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.TEXT);
  }
}

function updateInventorySale(data) {
  try {
    const ss    = SpreadsheetApp.openById(COMMISSIONS_SS_ID);
    const sheet = ss.getSheetByName(INVENTORY_SHEET);
    if (!sheet) throw new Error('Sheet "' + INVENTORY_SHEET + '" not found');

    const rowIndex = parseInt(data.rowIndex);
    if (!rowIndex || rowIndex < 2) throw new Error('Invalid row index: ' + rowIndex);

    const updates = data.updates || {};
    Object.entries(updates).forEach(([field, value]) => {
      const col = WRITE_COL[field];
      if (col) sheet.getRange(rowIndex, col).setValue(value || '');
    });

    SpreadsheetApp.flush();
    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.TEXT);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.TEXT);
  }
}
