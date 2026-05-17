// =========================================================
//  CONFIGURATION
// =========================================================
const SPREADSHEET_ID = '1R_O4llA1K43Y97GAgkK97WMvWbqg-tftz_FXpcUSZPU'; // আপনার শিট আইডি
const SHEET_NAME     = 'Examiner Information';

// ── Column indices (0-based) ──────────────────────────────
const COL = {
  SL:          0,   // A
  NAME:        1,   // B (Nick Name)
  TPIN:        3,   // D
  INST:        4,   // E
  DEPT:        5,   // F
  BATCH:       6,   // G (HSC Batch)
  RM:          7,   // H
  REMARKED_BY: 8,   // I
  MOB1:        9,   // J (Mobile Number)
  ALT:         10,  // K (Alternate)
  NAGAD:       11,  // L (Mobile Banking / Nagad)
  EN:          61,  // BJ English(%)
  BN:          64,  // BM Bangla(%)
  PHY:         67,  // BP Physics(%)
  CHEM:        70,  // BS Chemistry(%)
  MATH:        73,  // BV Math(%)
  BIO:         76,  // BY Biology(%)
  ICT:         79,  // CB ICT(%)
  TRAIN:       82,  // CE Training Report
  TRAIN_DATE:  83,  // CF Training Date
  CAMPUS:      88,  // CK Campus
  REMARK_RAW:  92   // CQ Remark
};

const ALLOW = {
  ENGLISH:   55,
  BANGLA:    48,
  PHYSICS:   48,
  ENVIRONMENT: 48,
  CHEMISTRY: 48,
  MATH:      48,
  BIOLOGY:   48,
  ICT:       48
};

// =========================================================
//  CORE LOGIC (Optimized for Speed)
// =========================================================

function doGet(e) {
  const action = (e && e.parameter) ? e.parameter.action : null;
  const query  = (e && e.parameter) ? e.parameter.query : '';
  
  let result = { success: false, error: 'No action specified' };

  if (action === 'lookup') {
    result = lookupByQuery(query);
  } else if (action === 'filterOptions') {
    result = getAllDataForSync();
  } else if (action === 'ping') {
    result = { success: true, pong: true, time: new Date().toISOString() };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  return doGet(e); // Handle POST same as GET
}

// শিট থেকে সব ডেটা একসাথে নিয়ে আসার জন্য (Background Sync)
function getAllDataForSync() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error('Sheet not found');

    const data = sheet.getDataRange().getValues();
    const rows = data.slice(1); // Header বাদ দিয়ে

    return {
      success: true,
      rowCount: rows.length,
      rows: rows, // এই rows ই আপনার অ্যাপে Instant Result দিবে
      allow: ALLOW,
      timestamp: Date.now()
    };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function lookupByQuery(query) {
  try {
    const q = String(query || '').trim().toLowerCase();
    if (!q) return { success: true, found: false };

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const body = data.slice(1);

    // শিটে খুঁজে দেখা (Fallback if local sync fails)
    for (let i = 0; i < body.length; i++) {
      const r = body[i];
      if (normalize_(r[COL.TPIN]) === q || 
          normalize_(r[COL.MOB1]) === q || 
          normalize_(r[COL.ALT]) === q) {
        return { success: true, found: true, data: mapRowToExaminer(r) };
      }
    }

    return { success: true, found: false };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// =========================================================
//  HELPERS
// =========================================================

function normalize_(s) {
  let val = String(s || '').trim().toLowerCase();
  // If it's a mobile number, take last 10 digits
  if (/^\d+$/.test(val.replace(/\D/g, ''))) {
     return val.replace(/\D/g, '').slice(-10);
  }
  return val;
}

function mapRowToExaminer(r) {
  return {
    sl: String(r[COL.SL]),
    name: String(r[COL.NAME]),
    tpin: String(r[COL.TPIN]),
    inst: String(r[COL.INST]),
    dept: String(r[COL.DEPT]),
    batch: String(r[COL.BATCH]),
    mobile: String(r[COL.MOB1]),
    alternate: String(r[COL.ALT]),
    nagad: String(r[COL.NAGAD]),
    // thresholds অনুযায়ী পাস/ফেল নির্ধারণ
    english:   { score: r[COL.EN],   allowed: r[COL.EN] >= ALLOW.ENGLISH },
    bangla:    { score: r[COL.BN],   allowed: r[COL.BN] >= ALLOW.BANGLA },
    physics:   { score: r[COL.PHY],  allowed: r[COL.PHY] >= ALLOW.PHYSICS },
    chemistry: { score: r[COL.CHEM], allowed: r[COL.CHEM] >= ALLOW.CHEMISTRY },
    math:      { score: r[COL.MATH], allowed: r[COL.MATH] >= ALLOW.MATH },
    biology:   { score: r[COL.BIO],  allowed: r[COL.BIO] >= ALLOW.BIOLOGY },
    ict:       { score: r[COL.ICT],  allowed: r[COL.ICT] >= ALLOW.ICT }
  };
}
