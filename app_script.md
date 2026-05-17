# Google Apps Script Setup (API JSON)

Here is the updated App Script code that handles both returning the web HTML (if you use it as a Web App directly) and answering JSON API requests for our Vite + React App.

### Instructions:
1. Open your Google Sheet
2. Go to **Extensions > Apps Script**
3. Select `Code.gs` and replace all its contents with the code below.
4. Click **Deploy > New deployment**.
5. Select type: **Web app**.
6. Execute as: **Me**.
7. Who has access: **Anyone** (This is required so the React app can fetch it without authentication issues).
8. Click **Deploy**.
9. Copy the Web App URL and update `SCRIPT_URL` in `src/App.tsx` (already updated to `AKfycbxun3tEo...`).

```javascript
// =========================================================
//  CONFIGURATION
// =========================================================
const SPREADSHEET_ID = '1R_O4llA1K43Y97GAgkK97WMvWbqg-tftz_FXpcUSZPU';
const SHEET_NAME     = 'Examiner Information';

// ── Column indices (0-based) ──────────────────────────────
const COL = {
  SL:          0,   // A
  NAME:        1,   // B  (Nick Name)
  TPIN:        3,   // D
  INST:        4,   // E
  DEPT:        5,   // F
  BATCH:       6,   // G  (HSC Batch)
  RM:          7,   // H
  REMARKED_BY: 8,   // I
  MOB1:        9,   // J  (Mobile Number)
  ALT:         10,  // K  (Alternate)
  NAGAD:       11,  // L  (Mobile Banking / Nagad)

  EN:          61,  // BJ  English(%)
  BN:          64,  // BM  Bangla(%)
  PHY:         67,  // BP  Physics(%)
  CHEM:        70,  // BS  Chemistry(%)
  MATH:        73,  // BV  Math(%)
  BIO:         76,  // BY  Biology(%)
  ICT:         79,  // CB  ICT(%)

  TRAIN:       82,  // CE  Training Report
  TRAIN_DATE:  83,  // CF  Training Date  ← fixed: space replaced with underscore
  CAMPUS:      88,  // CK  Campus
};

const ALLOW = {
  ENGLISH:   55,
  BANGLA:    48,
  PHYSICS:   48,
  CHEMISTRY: 48,
  MATH:      48,
  BIOLOGY:   48,
  ICT:       48
};

// =========================================================
//  HELPER FUNCTIONS
// =========================================================

function normalize_(s) {
  return String(s || '').trim().toLowerCase().replace(/\s+/g, '');
}

// Google Sheets strips leading 0 from numeric cells.
// Bangladeshi numbers: 01XXXXXXXXX stored as 1XXXXXXXXX (10 digits) → restore.
function padMobile_(v) {
  const s = String(v || '').trim().replace(/\.0$/, '');
  if (/^\d{10}$/.test(s)) return '0' + s;
  return s;
}

// =========================================================
//  CACHE & STORE
// =========================================================
function buildStore_() {
  const cache    = CacheService.getScriptCache();
  const cacheKey = 'examiner_store_v1';
  const cached   = cache.get(cacheKey);

  if (cached) {
    try {
      const parsed = JSON.parse(cached);
      const index  = new Map(parsed.indexEntries);
      const rows   = parsed.rows;
      return { index, rows, rowCount: parsed.rowCount };
    } catch (e) {
      // Cache corrupted — fall through to rebuild
    }
  }

  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    throw new Error('Sheet "' + SHEET_NAME + '" not found. Check SHEET_NAME constant.');
  }

  const data = sheet.getDataRange().getValues();
  const body = data.slice(1); // skip header row

  const indexEntries = [];

  function tryAdd(rawKey, rowIdx) {
    const k = normalize_(rawKey);
    if (!k) return;
    const existing = indexEntries.findIndex(([ek]) => ek === k);
    if (existing === -1) {
      indexEntries.push([k, rowIdx]);
    }
  }

  for (let i = 0; i < body.length; i++) {
    const r = body[i];
    tryAdd(r[COL.TPIN],             i);
    tryAdd(padMobile_(r[COL.MOB1]), i);
    tryAdd(padMobile_(r[COL.ALT]),  i);
  }

  const payload = JSON.stringify({ indexEntries, rows: body, rowCount: body.length });
  if (payload.length < 90000) {
    try { cache.put(cacheKey, payload, 360); } catch (e) { }
  }

  const index = new Map(indexEntries);
  return { index, rows: body, rowCount: body.length };
}

// =========================================================
//  doGet
// =========================================================
function doGet(e) {
  if (e && e.parameter && e.parameter.action) {
    let result;
    const action = e.parameter.action;
    const query  = e.parameter.query || '';

    if (action === 'lookup') {
      result = lookupExaminerByTPin(query);
    } else if (action === 'filterOptions') {
      result = getFilterOptionsFast();
    } else if (action === 'ping') {
      result = { success: true, pong: true };
    } else {
      result = { success: false, error: 'Unknown action: ' + action };
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  try {
    return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Examiner Assessment Test Checker')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    return HtmlService.createHtmlOutput(
      '<h2 style="font-family:sans-serif;color:#c00">Setup Error</h2>' +
      '<p style="font-family:sans-serif">' + err.message + '</p>' +
      '<p style="font-family:sans-serif">Please check that <code>Index.html</code> exists in this Apps Script project.</p>'
    ).setTitle('Error — Examiner Checker');
  }
}

// =========================================================
//  doPost
// =========================================================
function doPost(e) {
  if (!e) {
    return ContentService
      .createTextOutput('Manual run detected. Please call as Web App.')
      .setMimeType(ContentService.MimeType.TEXT);
  }

  try {
    let action, query;

    if (e.postData && e.postData.contents) {
      // JSON body
      const params = JSON.parse(e.postData.contents);
      action = params.action || '';
      query  = params.query  || '';
    } else {
      // URL-encoded form params
      action = (e.parameter && e.parameter.action) || '';
      query  = (e.parameter && e.parameter.query)  || '';
    }

    let result;
    if (action === 'lookup') {
      result = lookupExaminerByTPin(query);
    } else if (action === 'filterOptions') {
      result = getFilterOptionsFast();
    } else {
      result = { success: false, error: 'Invalid action: ' + action };
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// =========================================================
//  getFilterOptionsFast
// =========================================================
function getFilterOptionsFast() {
  try {
    const store = buildStore_();
    const result = { success: true, rowCount: store.rowCount, allow: ALLOW };
    
    // Attempt to include rows in the initial sync for "Ultra Search"
    try {
      const rowsPayload = JSON.stringify(store.rows);
      if (rowsPayload.length < 90000) {
        result.rows = store.rows;
      }
    } catch (e) { }
    
    return result;
  } catch (e) {
    Logger.log('getFilterOptionsFast error: ' + e.message);
    return { success: false, error: e.message };
  }
}

// =========================================================
//  lookupExaminerByTPin
// =========================================================
function lookupExaminerByTPin(query) {
  try {
    const raw = String(query || '').trim();
    if (!raw) return { success: true, found: false };

    const { index, rows } = buildStore_();

    // Try multiple variants to handle leading-zero mismatch
    const variants = [normalize_(raw)];
    if (/^\d{10}$/.test(raw))  variants.push(normalize_('0' + raw));
    if (/^0\d{10}$/.test(raw)) variants.push(normalize_(raw.slice(1)));

    let rowIdx = null;
    for (const v of variants) {
      if (index.has(v)) { rowIdx = index.get(v); break; }
    }
    if (rowIdx === null) return { success: true, found: false };

    const r = rows[rowIdx];

    function pct(v) {
      const n = parseFloat(v);
      return Number.isFinite(n) ? n : '';
    }

    const data = {
      sl:          r[COL.SL]           || '',
      name:        r[COL.NAME]         || '',
      tpin:        r[COL.TPIN]         || '',
      inst:        r[COL.INST]         || '',
      dept:        r[COL.DEPT]         || '',
      batch:       r[COL.BATCH]        || '',
      rm:          r[COL.RM]           || '',
      mobile:      padMobile_(r[COL.MOB1]),
      alternate:   padMobile_(r[COL.ALT]),
      nagad:       padMobile_(r[COL.NAGAD]),

      english:     { score: pct(r[COL.EN]),   allowed: pct(r[COL.EN])   >= ALLOW.ENGLISH   },
      bangla:      { score: pct(r[COL.BN]),   allowed: pct(r[COL.BN])   >= ALLOW.BANGLA    },
      physics:     { score: pct(r[COL.PHY]),  allowed: pct(r[COL.PHY])  >= ALLOW.PHYSICS   },
      chemistry:   { score: pct(r[COL.CHEM]), allowed: pct(r[COL.CHEM]) >= ALLOW.CHEMISTRY },
      math:        { score: pct(r[COL.MATH]), allowed: pct(r[COL.MATH]) >= ALLOW.MATH      },
      biology:     { score: pct(r[COL.BIO]),  allowed: pct(r[COL.BIO])  >= ALLOW.BIOLOGY   },
      ict:         { score: pct(r[COL.ICT]),  allowed: pct(r[COL.ICT])  >= ALLOW.ICT       },

      training:    r[COL.TRAIN]        || '',
      trainDate:   r[COL.TRAIN_DATE]   || '',   
      campus:      r[COL.CAMPUS]       || '',
      remarkedBy:  r[COL.REMARKED_BY]  || '',
    };

    return { success: true, found: true, data };

  } catch (e) {
    Logger.log('lookupExaminerByTPin error: ' + e.message);
    return { success: false, found: false, error: e.message };
  }
}

// =========================================================
//  CACHE UTILITIES
// =========================================================
function invalidateCache() {
  CacheService.getScriptCache().remove('examiner_store_v1');
  Logger.log('Cache cleared.');
}
```
