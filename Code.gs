// ─────────────────────────────────────────
// BENNET SCHLESINGER — Commissions & Inquiries
// Separate Apps Script deployment
// ─────────────────────────────────────────

const COMMISSIONS_SS_ID = '1j3dymNa0u_ajdAtUZax8sQAkEQJpi1BBhOJWMpTcBJ4';
const INQUIRIES_SS_ID   = '1nx30zFjY8TFplbI7CTAgtusLNbUDdZ0tPltvqJ8F6bM';
const COMMISSIONS_SHEET = 'Comissions'; // note: one 'm'
const INQUIRIES_SHEET   = 'Sheet1';

// Column indices (0-indexed)
const COL = {
  CLIENT:        0,  // A
  IMAGE:         1,  // B
  TITLE:         2,  // C
  YEAR:          3,  // D
  DESCRIPTION:   4,  // E
  DIMENSIONS:    5,  // F
  MATERIAL:      6,  // G
  SG_INVOICE:    7,  // H
  PRICE:         8,  // I
  SALE_PRICE:    9,  // J
  BENNET_PAYOUT: 10, // K
  ARTIST_PAID:   11, // L
  GALLERY_PAID:  12, // M
  PAYMENT_NOTES: 13, // N
  PROD_STATUS:   14, // O
  DUE_DATE:      15, // P
  NOTES:         16, // Q
  DAYS_TIL_DUE:  17, // R
  SHIP_ADDRESS:  18, // S
  SHIPPER:       19, // T
  SHIP_REF:      20, // U
  FILE_REF:      21, // V
  TRACKING:      22, // W
  TRACKING_SENT: 23, // X
  POC:           24, // Y
};

// Column numbers for setValue (1-indexed)
const WRITE_COL = {
  client:       1,
  title:        3,
  year:         4,
  description:  5,
  dimensions:   6,
  material:     7,
  sgInvoice:    8,
  price:        9,
  salePrice:    10,
  bennetPayout: 11,
  artistPaid:   12,
  galleryPaid:  13,
  paymentNotes: 14,
  status:       15,
  dueDate:      16,
  notes:        17,
  shipAddress:  19,
  shipper:      20,
  shipRef:      21,
  fileRef:      22,
  tracking:     23,
  trackingSent: 24,
  poc:          25,
};

// ─────────────────────────────────────────
function doGet(e) {
  const action = e && e.parameter && e.parameter.action;
  if (action === 'commissions') return serveCommissions();
  if (action === 'inquiries')   return serveInquiries();
  return ContentService
    .createTextOutput(JSON.stringify({ error: 'Unknown action' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.type === 'update_commission') return updateCommission(data);
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
        artistPaid:    String(row[COL.ARTIST_PAID]   || '').trim(),
        galleryPaid:   String(row[COL.GALLERY_PAID]  || '').trim(),
        paymentNotes:  String(row[COL.PAYMENT_NOTES] || '').trim(),
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
      .createTextOutput(JSON.stringify({ commissions: commissions }))
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

    const headers = data[0].map(h => String(h).trim());
    const rows    = data.slice(1).filter(r => r.some(c => c !== ''));

    const inquiries = rows.map(row => {
      const obj = {};
      headers.forEach((h, i) => { if (h) obj[h] = String(row[i] || '').trim(); });
      return obj;
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
