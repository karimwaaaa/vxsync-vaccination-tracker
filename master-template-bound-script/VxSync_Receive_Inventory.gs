/**
 * VxSync — Receive Inventory (in-Sheet modal dialog).
 *
 * Companion to the Dashboard.html "Receive Inventory" tab in the
 * standalone web app (Code.gs) — same idea (stage a batch of newly
 * received lots, review, then write every physical unit as its own
 * Lot_Expiry_Master row in one batched write), but opened from inside
 * the Sheet itself via VxSync ▸ 📷 Receive Inventory…, for whoever's
 * already working directly in the spreadsheet rather than the web app.
 *
 * This is a SEPARATE Apps Script project from the standalone Code.gs —
 * they cannot share code at runtime, so the column maps, schedule
 * lookup, and batch-write logic below are deliberately duplicated
 * (matching the same VX_LOT / TRK_SCHED conventions already used
 * elsewhere in this bound project) rather than imported from anywhere.
 * If Lot_Expiry_Master's or Vaccine_Schedule_Master's column layout
 * ever changes, both this file AND Code.gs's equivalent functions
 * (getReceivingFormOptions / getBrandsForReceiving /
 * receiveLotInventoryRows) need updating together.
 */

function vxReceiveShowDialog() {
  const html = HtmlService.createHtmlOutputFromFile('VxSync_Receive_Dialog')
    .setWidth(720)
    .setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(html, '📷 Receive Inventory');
}

// Same Vaccine_Schedule_Master column convention as TRK_SCHED
// (VxSync_Tracker_Backfill.gs) and SCHEDULE_COL (Code.gs) — 1-based here
// since this file works in Range coordinates throughout, same as the
// rest of this bound project.
const VXR_SCHED = Object.freeze({
  sheet: 'Vaccine_Schedule_Master', firstRow: 2,
  vaccineType: 1, brand: 3, active: 21
});

// Same physical Lot_Expiry_Master column layout as VX_LOT
// (VxSync_Lot_Dropdown.gs) and LOT (Code.gs) — A/B/H stay blank on a
// newly-received row until a claim stamps them; see the header comment
// on LOT in Code.gs for the full rationale, not repeated here.
const VXR_LOT = Object.freeze({
  sheet: 'Lot_Expiry_Master', firstRow: 6,
  sessionRef: 1, vaccinationDate: 2, site: 3, vaccineType: 4,
  vaccineBrand: 5, lotNumber: 6, expiryDate: 7, availability: 8
});

const VXR_MAX_QTY_PER_LINE = 2000; // same fat-finger guard as the standalone web app's receiveLotInventoryRows

function vxReceiveGetFormOptions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return {
    vaccineTypes: vxReceiveGetVaccineTypes_(ss),
    sites: vxReceiveGetKnownSites_(ss)
  };
}

function vxReceiveGetBrands(vaccineType) {
  if (!vaccineType) return [];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return vxReceiveGetBrandsForType_(ss, String(vaccineType).trim());
}

function vxReceiveGetVaccineTypes_(ss) {
  const sheet = ss.getSheetByName(VXR_SCHED.sheet);
  if (!sheet) throw new Error(VXR_SCHED.sheet + ' sheet not found.');
  const lastRow = sheet.getLastRow();
  if (lastRow < VXR_SCHED.firstRow) return [];
  const data = sheet.getRange(VXR_SCHED.firstRow, 1, lastRow - VXR_SCHED.firstRow + 1, VXR_SCHED.active).getValues();
  const types = new Set();
  data.forEach(function (row) {
    const type = row[VXR_SCHED.vaccineType - 1];
    const active = String(row[VXR_SCHED.active - 1] || '').trim().toLowerCase() === 'yes';
    if (type && active) types.add(type);
  });
  return Array.from(types).sort();
}

function vxReceiveGetBrandsForType_(ss, vaccineType) {
  const sheet = ss.getSheetByName(VXR_SCHED.sheet);
  if (!sheet) throw new Error(VXR_SCHED.sheet + ' sheet not found.');
  const lastRow = sheet.getLastRow();
  if (lastRow < VXR_SCHED.firstRow) return [];
  const data = sheet.getRange(VXR_SCHED.firstRow, 1, lastRow - VXR_SCHED.firstRow + 1, VXR_SCHED.active).getValues();
  const brands = new Set();
  data.forEach(function (row) {
    const type = row[VXR_SCHED.vaccineType - 1];
    const brand = row[VXR_SCHED.brand - 1];
    const active = String(row[VXR_SCHED.active - 1] || '').trim().toLowerCase() === 'yes';
    if (type === vaccineType && active && brand) brands.add(brand);
  });
  return Array.from(brands).sort();
}

// Distinct "Site Name - PROVINCE" strings already in use, pulled from
// BOTH Vaccination_Tracker's Vaccination Location column (TRK.location —
// see VxSync_Tracker_Backfill.gs) and Lot_Expiry_Master's own Site of
// Vaccination column, so a site that's only ever been used for
// receiving (never yet an actual vaccination) still shows up here too.
function vxReceiveGetKnownSites_(ss) {
  const sites = new Set();

  const trackerSheet = ss.getSheetByName('Vaccination_Tracker');
  if (trackerSheet) {
    const lastRow = trackerSheet.getLastRow();
    if (lastRow >= (typeof TRK !== 'undefined' ? TRK.firstRow : 2)) {
      const locationCol = (typeof TRK !== 'undefined') ? TRK.location : 25; // fallback matches TRK.location's known value if TRK isn't loaded for some reason
      const firstRow = (typeof TRK !== 'undefined') ? TRK.firstRow : 2;
      const data = trackerSheet.getRange(firstRow, locationCol, lastRow - firstRow + 1, 1).getValues();
      data.forEach(function (row) { if (row[0]) sites.add(String(row[0]).trim()); });
    }
  }

  const lotSheet = ss.getSheetByName(VXR_LOT.sheet);
  if (lotSheet) {
    const lastRow = lotSheet.getLastRow();
    if (lastRow >= VXR_LOT.firstRow) {
      const data = lotSheet.getRange(VXR_LOT.firstRow, VXR_LOT.site, lastRow - VXR_LOT.firstRow + 1, 1).getValues();
      data.forEach(function (row) { if (row[0]) sites.add(String(row[0]).trim()); });
    }
  }

  return Array.from(sites).sort();
}

// Mirrors parseDateOnlyLocal_ in Code.gs exactly (same reasoning: a
// plain new Date("2027-08-01") parses as UTC midnight, which can land on
// the WRONG calendar day once displayed in a local timezone behind UTC —
// this constructs the Date from its parts instead, always local midnight).
function vxReceiveParseDateOnly_(dateStr) {
  if (!dateStr) return null;
  const s = String(dateStr).trim();
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) {
    const fallback = new Date(s);
    return isNaN(fallback.getTime()) ? null : fallback;
  }
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}

// `records`: array of { site, vaccineType, vaccineBrand, lotNumber,
// expiryDate (YYYY-MM-DD string), quantity }. Same validate-everything-
// before-writing-anything approach as the standalone web app's
// receiveLotInventoryRows in Code.gs — see that function's comment for
// the full reasoning, not repeated here.
function vxReceiveSubmitBatch(records) {
  if (!Array.isArray(records) || records.length === 0) {
    throw new Error('Nothing to submit — the batch is empty.');
  }

  const newRows = [];
  records.forEach(function (rec, idx) {
    const n = idx + 1;
    const site = String((rec && rec.site) || '').trim();
    const vaccineType = String((rec && rec.vaccineType) || '').trim();
    const vaccineBrand = String((rec && rec.vaccineBrand) || '').trim();
    const lotNumber = String((rec && rec.lotNumber) || '').trim();
    const expiryStr = String((rec && rec.expiryDate) || '').trim();
    const qtyRaw = (rec && rec.quantity !== undefined && rec.quantity !== null && rec.quantity !== '') ? rec.quantity : 1;
    const quantity = Math.floor(Number(qtyRaw));

    if (!site) throw new Error('Line ' + n + ': Site of Vaccination is required.');
    if (!vaccineType) throw new Error('Line ' + n + ': Vaccine Type is required.');
    if (!vaccineBrand) throw new Error('Line ' + n + ': Vaccine Brand is required.');
    if (!lotNumber) throw new Error('Line ' + n + ': Lot Number is required.');
    if (!expiryStr) throw new Error('Line ' + n + ': Expiry Date is required.');
    const expiryDate = vxReceiveParseDateOnly_(expiryStr);
    if (!expiryDate) throw new Error('Line ' + n + ': Expiry Date "' + expiryStr + '" could not be understood.');
    if (!Number.isFinite(quantity) || quantity < 1) {
      throw new Error('Line ' + n + ': Quantity must be a whole number of at least 1.');
    }
    if (quantity > VXR_MAX_QTY_PER_LINE) {
      throw new Error('Line ' + n + ': Quantity (' + quantity + ') is over the ' + VXR_MAX_QTY_PER_LINE +
        '-per-line safety cap. If this is really correct, split it across more than one line.');
    }

    for (let u = 0; u < quantity; u++) {
      newRows.push(['', '', site, vaccineType, vaccineBrand, lotNumber, expiryDate, '']);
    }
  });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = ss.getSheetByName(VXR_LOT.sheet);
    if (!sheet) throw new Error(VXR_LOT.sheet + ' sheet not found.');
    const firstBlankRow = Math.max(sheet.getLastRow() + 1, VXR_LOT.firstRow);
    sheet.getRange(firstBlankRow, 1, newRows.length, 8).setValues(newRows);
    try { sheet.getRange(firstBlankRow, VXR_LOT.expiryDate, newRows.length, 1).setNumberFormat('yyyy-mm-dd'); } catch (fmtErr) { /* cosmetic only, never block the write */ }
    // Same defense-in-depth reasoning as onOpen's refreshVxSyncLotDropdowns_
    // call: a freshly-received lot with a brand-new Vaccine Type/Brand
    // combination should still get correct dropdown validation on sight,
    // not only after the next manual edit or Sheet reopen.
    try { refreshVxSyncLotDropdowns_(ss); } catch (ddErr) { console.error('vxReceiveSubmitBatch: dropdown refresh failed:', ddErr); }
    return { ok: true, rowsAdded: newRows.length };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
//  AI PHOTO FALLBACK (Gemini) — for when the barcode scanner can't read
//  a box at all: damaged packaging, a smudged label, a barcode format
//  ZXing doesn't support, or stock that never had one printed. This is
//  a SEPARATE button in the dialog (see handleAiScan in
//  VxSync_Receive_Dialog.html) — never triggered automatically after a
//  failed barcode scan, so an AI call only ever happens when the person
//  deliberately chooses to spend one.
//
//  ONE-TIME SETUP REQUIRED, by hand, before this will work: open this
//  project in the Apps Script editor → Project Settings → Script
//  Properties → add a property named exactly GEMINI_API_KEY with your
//  own Gemini API key as the value (from Google AI Studio — a
//  DIFFERENT credential than anything already used by this Sheet).
//  Nobody else viewing this dialog's page source can ever see that key;
//  it's read here, server-side, and never sent to the browser.
//
//  CONFIDENCE: the prompt below explicitly instructs Gemini to return
//  null for anything it isn't confident about or can't clearly read —
//  never a guess, never placeholder text. vxReceiveParseGeminiFields_
//  then double-checks that independently: anything that isn't a
//  plausible non-empty string is forced to null before it's ever
//  allowed to reach the dialog, so a malformed or unexpected model
//  response can degrade to "nothing prefilled," never to "wrong value
//  silently prefilled."
//
//  IMAGE HANDLING: the resized photo arrives here as one base64 string,
//  is used only to build the single request below, and is never
//  written to Drive, a Sheet, or a Property — it goes out of scope the
//  moment this function returns. Nothing about this call keeps a copy.
//
//  MODEL / ENDPOINT: gemini-2.5-flash was current and free-tier-eligible
//  at the time this was written, but both the exact model name AND the
//  free tier's request limits are the kind of thing Google changes
//  without much notice — if this starts failing outright (not just
//  rate-limiting), check https://ai.google.dev/gemini-api/docs/models
//  for whether the model name needs updating.
// ============================================================

const VXR_GEMINI_MODEL = 'gemini-2.5-flash';
const VXR_GEMINI_ENDPOINT = 'https://generativelanguage.googleapis.com/v1beta/models/' + VXR_GEMINI_MODEL + ':generateContent';

const VXR_GEMINI_PROMPT =
  'You are reading a photo of a vaccine box or vial label for a vaccination inventory system. ' +
  'Extract exactly these four fields and respond with ONLY a single JSON object, no other text, no markdown fencing:\n' +
  '{\n' +
  '  "lotNumber": string or null,\n' +
  '  "expiryDate": string in YYYY-MM-DD format, or null,\n' +
  '  "brand": string or null,\n' +
  '  "vaccineType": string or null\n' +
  '}\n' +
  'Rules:\n' +
  '- If you are not confident about a field, or it is not clearly visible or legible in the photo, set it to null.\n' +
  '- Never guess. Never output placeholder text like "N/A", "unknown", or "not visible" — use null for that, not a string.\n' +
  '- "expiryDate" must be null unless you can express it as a real YYYY-MM-DD date (if only a month/year is printed, e.g. "2027-08", set it to null rather than inventing a day).\n' +
  '- Respond with the JSON object only.';

// `base64Jpeg`: raw base64 (no "data:image/..." prefix) of an already
// client-side-resized photo — see resizeImageToBase64_ in
// VxSync_Receive_Dialog.html. `mimeType`: e.g. "image/jpeg".
function vxReceiveExtractFromPhotoAI(base64Jpeg, mimeType) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    return {
      ok: false,
      message: 'Gemini API key isn\'t set up yet for this Sheet. An admin needs to add a GEMINI_API_KEY ' +
        'Script Property (Apps Script editor → Project Settings → Script Properties) — see the comment ' +
        'above vxReceiveExtractFromPhotoAI in VxSync_Receive_Inventory.gs.'
    };
  }
  if (!base64Jpeg) {
    return { ok: false, message: 'No photo data was received.' };
  }

  const payload = {
    contents: [{
      parts: [
        { text: VXR_GEMINI_PROMPT },
        { inline_data: { mime_type: mimeType || 'image/jpeg', data: base64Jpeg } }
      ]
    }],
    generationConfig: { responseMimeType: 'application/json' }
  };

  let response;
  try {
    response = UrlFetchApp.fetch(VXR_GEMINI_ENDPOINT + '?key=' + encodeURIComponent(apiKey), {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
  } catch (networkErr) {
    return { ok: false, message: 'Could not reach the Gemini API: ' + networkErr.message };
  }

  const code = response.getResponseCode();
  const bodyText = response.getContentText();

  if (code === 429) {
    return { ok: false, rateLimited: true, message: 'Gemini API rate limit hit (HTTP 429).' };
  }
  if (code !== 200) {
    // RESOURCE_EXHAUSTED is Google's own error status for a quota/rate
    // limit being hit even on a non-429 code in some cases — checked as
    // a substring rather than a strict parse since the exact error body
    // shape isn't worth being fragile about here.
    const rateLimited = bodyText.indexOf('RESOURCE_EXHAUSTED') !== -1 || bodyText.indexOf('quota') !== -1;
    return {
      ok: false,
      rateLimited: rateLimited,
      message: 'Gemini API returned HTTP ' + code + (rateLimited ? ' (usage limit).' : ': ' + bodyText.slice(0, 200))
    };
  }

  let parsedResponse;
  try {
    parsedResponse = JSON.parse(bodyText);
  } catch (parseErr) {
    return { ok: false, message: 'Could not understand the Gemini API\'s response.' };
  }

  const rawText = vxReceiveExtractGeminiText_(parsedResponse);
  if (!rawText) {
    return { ok: false, message: 'Gemini did not return any readable content for that photo.' };
  }

  const fields = vxReceiveParseGeminiFields_(rawText);
  if (!fields) {
    return { ok: false, message: 'Gemini\'s response wasn\'t in the expected format — try a clearer photo, or enter manually.' };
  }

  return { ok: true, fields: fields };
}

// Pulls the model's text out of the standard generateContent response
// shape. Defensive about structure — an unexpected shape here should
// fail as "nothing extracted," never throw and break the whole dialog.
function vxReceiveExtractGeminiText_(parsedResponse) {
  try {
    const candidate = parsedResponse.candidates && parsedResponse.candidates[0];
    const part = candidate && candidate.content && candidate.content.parts && candidate.content.parts[0];
    return (part && part.text) ? String(part.text).trim() : '';
  } catch (e) {
    return '';
  }
}

// Belt-and-suspenders on top of the prompt's own "use null, never
// guess" instruction: even if Gemini's JSON is well-formed, this still
// independently refuses to pass through anything that isn't a plausible
// non-empty string, and validates expiryDate's shape rather than
// trusting it blindly. Returns null (not a partially-filled object) if
// the response can't be parsed as JSON at all.
function vxReceiveParseGeminiFields_(rawText) {
  // responseMimeType: 'application/json' should mean rawText IS the raw
  // JSON already, but strip a markdown code fence defensively in case
  // the model wraps it anyway (seen often enough across model versions
  // to be worth guarding rather than assuming strict compliance).
  const cleaned = rawText.replace(/^```(?:json)?/i, '').replace(/```$/, '').trim();

  let obj;
  try {
    obj = JSON.parse(cleaned);
  } catch (e) {
    return null;
  }
  if (!obj || typeof obj !== 'object') return null;

  function cleanString_(v) {
    if (typeof v !== 'string') return null;
    const t = v.trim();
    if (!t) return null;
    const lower = t.toLowerCase();
    if (lower === 'null' || lower === 'n/a' || lower === 'na' || lower === 'unknown' || lower === 'not visible') return null;
    return t;
  }

  const lotNumber = cleanString_(obj.lotNumber);
  const brand = cleanString_(obj.brand);
  const vaccineType = cleanString_(obj.vaccineType);

  let expiryDate = cleanString_(obj.expiryDate);
  if (expiryDate && !/^\d{4}-\d{2}-\d{2}$/.test(expiryDate)) {
    expiryDate = null; // don't guess at reformatting a differently-shaped date — safer to leave it blank
  }

  return { lotNumber: lotNumber, expiryDate: expiryDate, brand: brand, vaccineType: vaccineType };
}
