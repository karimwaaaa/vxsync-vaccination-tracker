/**
 * VxSync Vaccination_Tracker backfill assist.
 *
 * Lets logistics hand-enter an EXISTING/historical vaccination record
 * directly into Vaccination_Tracker (a dose given before this system
 * existed, or administered somewhere the web app wasn't used) with the
 * same typo-proofing and derived-field computation the web app gives a
 * record entered through it — so a backfilled row is a complete row,
 * not a half-finished one that quietly breaks whatever reads Next Dose /
 * Recommended Date / Series Status later.
 *
 * Column maps below are 1-based (A=1) and confirmed against the real
 * workbook headers, not guessed:
 *   Vaccination_Tracker header row 1, data starts row 2.
 *   Recipients_Master header row 1, data starts row 2.
 *   Vaccinators_Master header row 1, data starts row 2.
 *   Vaccine_Schedule_Master header row 1, data starts row 2 (same sheet
 *     Code.gs's SCHEDULE_COL already reads, same 0-based columns +1).
 *   Disposition_Master header row 1, data starts row 2 (already drives
 *     Vaccination_Entry's own Reason Category / Specific Reason cascade
 *     via a SORT/UNIQUE/FILTER array formula — this reuses the exact
 *     same table for Vaccination_Tracker's cascade instead).
 *
 * Wired into the shared onEdit(e) in MasterTemplate_Code.gs — see the
 * "VACCINATION TRACKER BACKFILL" block added there. Not a separate
 * installed trigger, for the same reason VxSync_Lot_Dropdown.gs isn't:
 * a bare onEdit needs no authorization and survives a Drive copy
 * automatically, an installed trigger doesn't.
 *
 * KNOWN LIMITATION — File > Import into Recipients_Master (as opposed
 * to typing or copy-pasting cell values directly into it) is not
 * guaranteed to fire onEdit; Google's own behavior here is undocumented
 * and inconsistent across import modes (replace sheet / new sheet /
 * append rows). If a CSV is ever loaded into Recipients_Master via
 * File > Import and the Recipient ID dropdown on Vaccination_Tracker
 * doesn't update, that's this gap, not a bug — either reopen the
 * spreadsheet (fires onOpen, full rebuild) or edit any single cell in
 * Recipients_Master (fires onEdit, targeted rebuild) to force it.
 * Confirmed workflow as of 2026-09-29: admin always types or
 * copy-pastes directly into Recipients_Master, never uses File >
 * Import, so this has not been made automatic on purpose — a
 * time-based trigger would close the gap but was judged unnecessary
 * overhead for a path this system doesn't actually use.
 */

const TRK = Object.freeze({
  sheet: 'Vaccination_Tracker', firstRow: 2,
  recordId: 1, recipientId: 2, recipientName: 3, email: 4, dob: 5, category: 6,
  company: 7, deptId: 8, dept: 9, site: 10,
  vaccineType: 11, intendedDose: 12, disposition: 13, reasonCategory: 14,
  specificReason: 15, statusDate: 16, reviewDate: 17,
  brand: 18, doseNumber: 19, vaccinationDate: 20, lotNumber: 21, expiryDate: 22,
  adminSite: 23, route: 24, location: 25, vaccinator: 26, vaccinatorLicense: 27,
  scheduleCode: 28, nextDose: 29, recommendedDate: 30, scheduledAppt: 31,
  reminderDate: 32, scheduleStatus: 33, reminderStatus: 34, reminderSentDate: 35,
  seriesStatus: 36, scheduleOverride: 37, remarks: 38
});

const TRK_REC = Object.freeze({
  sheet: 'Recipients_Master', firstRow: 2,
  id: 1, employeeId: 2, lastName: 3, firstName: 4, middleName: 5, dob: 6, sex: 7,
  category: 8, company: 9, deptId: 10, dept: 11, site: 12, email: 13, mobile: 14,
  enrollSource: 15, dateAdded: 16, active: 17, remarks: 18, fullName: 19
});

const TRK_VACC = Object.freeze({
  sheet: 'Vaccinators_Master', firstRow: 2,
  id: 1, lastName: 2, firstName: 3, middleName: 4, title: 5, licenseType: 6,
  licenseNumber: 7, ptr: 8, tin: 9, email: 10, mobile: 11, active: 12,
  remarks: 13, displayName: 14
});

// Same sheet/columns Code.gs's SCHEDULE_COL already reads — 0-based there,
// +1 here since this file works in 1-based Range coordinates throughout.
const TRK_SCHED = Object.freeze({
  sheet: 'Vaccine_Schedule_Master', firstRow: 2,
  vaccineType: 1, brand: 3, scheduleCode: 5, scheduleDisplay: 6,
  scheduleType: 8, currentDose: 9, nextAction: 10, intervalValue: 11,
  intervalUnit: 12, totalSeriesDoses: 16, active: 21
});

const TRK_DISP = Object.freeze({
  sheet: 'Disposition_Master', firstRow: 2,
  disposition: 1, reasonCategory: 2, specificReason: 3
});

// Fixed lists — NOT derived from a master sheet, confirmed against the
// real Data Validation already installed on Vaccination_Entry (B12's
// Intended Dose dropdown is this exact 6-value set, not a per-schedule
// computed list) and EntryForm.html's own Admin Site / Route /
// Disposition <option> lists, so the Tracker's dropdowns say the same
// thing the web app's do — no separate list to drift out of sync with.
const TRK_DOSE_NUMBER_OPTIONS = ['Dose 1', 'Dose 2', 'Dose 3', 'Booster', 'Annual Dose', 'Not Applicable'];
const TRK_ADMIN_SITE_OPTIONS = ['Left Deltoid', 'Right Deltoid', 'Left Anterolateral Thigh', 'Right Anterolateral Thigh', 'Other / Clinical Review'];
const TRK_ROUTE_OPTIONS = ['IM', 'SC', 'ID', 'Oral'];
const TRK_DISPOSITION_OPTIONS = ['Administered', 'Deferred', 'Declined', 'No-show', 'Contraindicated – Temporary', 'Contraindicated – Permanent'];

// Recipient ID values look like "VAC-000102" — Sheets' default column
// width clips that behind the dropdown arrow. Widened automatically
// every time this column's dropdown is (re)built, so nobody has to
// remember to drag it wider by hand, on this sheet or any future one.
const TRK_RECIPIENT_ID_COL_WIDTH = 140;

// ============================================================
//  ENTRY POINT — called from MasterTemplate_Code.gs's shared onEdit(e)
// ============================================================
function vxSyncTrackerEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  const ss = e.source;
  if (sheet.getName() !== TRK.sheet) return;

  const row = e.range.getRow();
  const col = e.range.getColumn();
  const lastRow = e.range.getLastRow();
  if (lastRow < TRK.firstRow) return;

  // ---- Record ID auto-generation (whole edited range, paste-safe) ----
  trkAssignRecordIds_(sheet, Math.max(TRK.firstRow, row), lastRow);

  // A multi-row paste (e.g. this whole test kit) touches many columns at
  // once — re-run every per-row refresh below across the full pasted
  // range rather than assuming a single-cell edit, same reasoning as
  // Recipients_Master's own paste handling in MasterTemplate_Code.gs.
  const first = Math.max(TRK.firstRow, row);
  const last = lastRow;
  const endCol = e.range.getLastColumn();

  const touches = (c) => col <= c && endCol >= c;

  if (touches(TRK.recipientId)) trkPrefillFromRecipient_(sheet, ss, first, last);
  if (touches(TRK.vaccineType)) trkRefreshBrandOptions_(sheet, ss, first, last);
  if (touches(TRK.brand)) trkRefreshScheduleCodeOptions_(sheet, ss, first, last);
  if (touches(TRK.vaccinator)) trkPrefillVaccinatorLicense_(sheet, ss, first, last);
  if (touches(TRK.disposition)) trkRefreshReasonCategoryOptions_(sheet, ss, first, last);
  if (touches(TRK.reasonCategory)) trkRefreshSpecificReasonOptions_(sheet, ss, first, last);

  if (touches(TRK.vaccinationDate) || touches(TRK.scheduleCode) || touches(TRK.doseNumber)) {
    trkRecomputeFollowUp_(sheet, ss, first, last);
  }

  // Lot inventory claim — see trkClaimLotInventory_ above for full
  // rationale. Any edit to one of these columns could be what finally
  // makes a row "complete enough" to identify a target unit; the
  // function itself no-ops on an incomplete row and is idempotent on an
  // already-claimed one, so re-checking on every relevant edit is safe.
  if (touches(TRK.disposition) || touches(TRK.vaccineType) || touches(TRK.brand) ||
      touches(TRK.lotNumber) || touches(TRK.location) || touches(TRK.vaccinationDate)) {
    trkClaimLotInventory_(sheet, ss, first, last);
  }
}

// Also called from the shared onEdit(e) when Vaccine_Schedule_Master,
// Disposition_Master, Recipients_Master, or Vaccinators_Master is
// edited (Active flipped, a row added, etc.) — the SOURCE table
// changed, so every Tracker row's dropdowns fed by that specific table
// need re-checking, not just whichever row was last touched. Only
// refreshes what that source actually feeds, not everything, since a
// full re-validate of every column on every row is unnecessary work for
// e.g. a Vaccinators_Master edit that has nothing to do with Schedule
// Codes.
function vxSyncTrackerSourceEdit(e, sourceSheetName) {
  if (!e || !e.range) return;
  const ss = e.source;
  const sheet = ss.getSheetByName(TRK.sheet);
  if (!sheet) return;
  // Deliberately getMaxRows(), NOT getLastRow() — a brand-new client's
  // Vaccination_Tracker is completely empty (header only, getLastRow()
  // === 1) until someone starts encoding into it, but the dropdowns need
  // to be ready and waiting in every blank row from the very first time
  // Recipients_Master/Vaccinators_Master/Schedule/Disposition gets
  // populated, not only once Tracker itself has rows. getLastRow() here
  // was the actual bug behind three separate "ran setup manually and it
  // worked" reports — it silently no-op'd on every fresh client until
  // trkRefreshAllDropdowns_ (which already used getMaxRows()) got run
  // some other way, via onOpen or the manual setup function.
  const lastRow = sheet.getMaxRows();
  if (lastRow < TRK.firstRow) return;
  const first = TRK.firstRow;

  if (sourceSheetName === TRK_SCHED.sheet) {
    trkRefreshBrandOptions_(sheet, ss, first, lastRow);
    trkRefreshScheduleCodeOptions_(sheet, ss, first, lastRow);
    trkRecomputeFollowUp_(sheet, ss, first, lastRow);
    const types = [...new Set(trkGetActiveScheduleData_(ss).map(function (r) { return r.type; }))].sort((a, b) => a.localeCompare(b));
    sheet.getRange(first, TRK.vaccineType, lastRow - first + 1, 1).setDataValidation(trkListRule_(types));
  } else if (sourceSheetName === TRK_DISP.sheet) {
    trkRefreshReasonCategoryOptions_(sheet, ss, first, lastRow);
    trkRefreshSpecificReasonOptions_(sheet, ss, first, lastRow);
  } else if (sourceSheetName === TRK_REC.sheet) {
    const recIds = [...trkGetRecipientsById_(ss).keys()].sort((a, b) => a.localeCompare(b));
    sheet.getRange(first, TRK.recipientId, lastRow - first + 1, 1).setDataValidation(trkListRule_(recIds));
    sheet.setColumnWidth(TRK.recipientId, TRK_RECIPIENT_ID_COL_WIDTH);
  } else if (sourceSheetName === TRK_VACC.sheet) {
    const vaccNames = [...trkGetVaccinatorsByDisplayName_(ss).keys()].sort((a, b) => a.localeCompare(b));
    sheet.getRange(first, TRK.vaccinator, lastRow - first + 1, 1).setDataValidation(trkListRule_(vaccNames));
  }
}

// ============================================================
//  RECORD ID AUTO-GENERATION (VAX-###### — same scheme as
//  getNextTrackerId_ in Code.gs, so the web app and manual sheet entry
//  never produce colliding IDs). Locked the same way
//  MasterTemplate_Code.gs's Recipients_Master ID stamping already is —
//  both sides read/increment "the last number used", so both need to
//  hold the SAME lock while they do it.
// ============================================================
function trkAssignRecordIds_(sheet, first, last) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const idRange = sheet.getRange(TRK.firstRow, TRK.recordId, sheet.getLastRow() - TRK.firstRow + 1, 1);
    const ids = idRange.getValues();
    let maxNum = 0;
    ids.forEach(function (r) {
      const m = /(\d+)$/.exec(String(r[0] || ''));
      if (m) maxNum = Math.max(maxNum, parseInt(m[1], 10));
    });
    for (let row = first; row <= last; row++) {
      const cell = sheet.getRange(row, TRK.recordId);
      if (cell.getValue()) continue; // never overwrite an existing ID
      // Skip a fully blank row (nothing entered yet) — don't burn an ID
      // number on a row nobody has actually started filling in.
      const rowHasContent = sheet.getRange(row, TRK.recipientId, 1, TRK.remarks - TRK.recipientId + 1)
        .getValues()[0].some(function (v) { return v !== '' && v !== null; });
      if (!rowHasContent) continue;
      maxNum += 1;
      cell.setValue('VAX-' + String(maxNum).padStart(6, '0'));
    }
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
//  RECIPIENT ID -> PREFILL (Name, Email, DOB, Category, Company,
//  Department ID/Name, Assigned Site). Recipient ID is the lookup key,
//  never Name — two different people can share a name (this system
//  already treats that as expected, not a data error, elsewhere), so
//  matching on Name instead of ID risks silently attaching one
//  recipient's dose history to a different person entirely.
// ============================================================
function trkGetRecipientsById_(ss) {
  const sheet = ss.getSheetByName(TRK_REC.sheet);
  if (!sheet) throw new Error('Missing worksheet: ' + TRK_REC.sheet);
  const last = sheet.getLastRow();
  const byId = new Map();
  if (last < TRK_REC.firstRow) return byId;
  const data = sheet.getRange(TRK_REC.firstRow, 1, last - TRK_REC.firstRow + 1, TRK_REC.fullName).getDisplayValues();
  data.forEach(function (row) {
    const id = row[TRK_REC.id - 1];
    if (id) byId.set(id, row);
  });
  return byId;
}

function trkPrefillFromRecipient_(sheet, ss, first, last) {
  const byId = trkGetRecipientsById_(ss);
  const ids = sheet.getRange(first, TRK.recipientId, last - first + 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    const row = first + i;
    const id = ids[i][0];
    if (!id) continue; // blank Recipient ID — leave whatever's there, nothing to prefill from
    const rec = byId.get(id);
    if (!rec) continue; // an ID that doesn't exist in Recipients_Master — leave as-is, don't guess
    sheet.getRange(row, TRK.recipientName).setValue(rec[TRK_REC.fullName - 1] ||
      (rec[TRK_REC.lastName - 1] + ', ' + rec[TRK_REC.firstName - 1]));
    sheet.getRange(row, TRK.email).setValue(rec[TRK_REC.email - 1] || '');
    sheet.getRange(row, TRK.dob).setValue(rec[TRK_REC.dob - 1] || '');
    sheet.getRange(row, TRK.category).setValue(rec[TRK_REC.category - 1] || '');
    sheet.getRange(row, TRK.company).setValue(rec[TRK_REC.company - 1] || '');
    sheet.getRange(row, TRK.deptId).setValue(rec[TRK_REC.deptId - 1] || '');
    sheet.getRange(row, TRK.dept).setValue(rec[TRK_REC.dept - 1] || '');
    sheet.getRange(row, TRK.site).setValue(rec[TRK_REC.site - 1] || '');
  }
  // Recipient ID itself is validated as a dropdown, not free text — see
  // trkRefreshStaticDropdowns_/refreshVxSyncLotDropdowns_-style full
  // refresh below, which (re)builds this list from Recipients_Master.
}

// ============================================================
//  VACCINATOR -> PREFILL LICENSE NUMBER (same ID-not-name principle,
//  though here the dropdown is keyed on Display Name, matching how
//  getVaccinators() in Code.gs already builds the web app's own
//  Vaccinator dropdown — Display Name is what's actually stored in
//  Tracker column Z today, not a separate Vaccinator ID column, so this
//  matches existing convention rather than inventing a new one).
// ============================================================
function trkGetVaccinatorsByDisplayName_(ss) {
  const sheet = ss.getSheetByName(TRK_VACC.sheet);
  if (!sheet) throw new Error('Missing worksheet: ' + TRK_VACC.sheet);
  const last = sheet.getLastRow();
  const byName = new Map();
  if (last < TRK_VACC.firstRow) return byName;
  const data = sheet.getRange(TRK_VACC.firstRow, 1, last - TRK_VACC.firstRow + 1, TRK_VACC.displayName).getDisplayValues();
  data.forEach(function (row) {
    const active = String(row[TRK_VACC.active - 1] || '').trim().toLowerCase() === 'yes';
    if (!active || !row[TRK_VACC.lastName - 1]) return;
    const display = row[TRK_VACC.displayName - 1] ||
      (row[TRK_VACC.lastName - 1] + ', ' + row[TRK_VACC.firstName - 1] +
        (row[TRK_VACC.middleName - 1] ? ' ' + row[TRK_VACC.middleName - 1] : '') +
        (row[TRK_VACC.title - 1] ? ', ' + row[TRK_VACC.title - 1] : ''));
    byName.set(display, row);
  });
  return byName;
}

function trkPrefillVaccinatorLicense_(sheet, ss, first, last) {
  const byName = trkGetVaccinatorsByDisplayName_(ss);
  const names = sheet.getRange(first, TRK.vaccinator, last - first + 1, 1).getValues();
  for (let i = 0; i < names.length; i++) {
    const row = first + i;
    const name = names[i][0];
    if (!name) continue;
    const v = byName.get(name);
    if (!v) continue;
    sheet.getRange(row, TRK.vaccinatorLicense).setValue(v[TRK_VACC.licenseNumber - 1] || '');
  }
}

// ============================================================
//  VACCINE TYPE -> BRAND -> SCHEDULE CODE cascade, same Active=Yes
//  filtering Code.gs's own getVaccineTypes()/getBrands() already use —
//  reuses the SAME Vaccine_Schedule_Master columns, just read directly
//  here since this runs inside the bound script, not via a client call.
// ============================================================
function trkGetActiveScheduleData_(ss) {
  const sheet = ss.getSheetByName(TRK_SCHED.sheet);
  if (!sheet) throw new Error('Missing worksheet: ' + TRK_SCHED.sheet);
  const last = sheet.getLastRow();
  const rows = [];
  if (last >= TRK_SCHED.firstRow) {
    const data = sheet.getRange(TRK_SCHED.firstRow, 1, last - TRK_SCHED.firstRow + 1, TRK_SCHED.active).getDisplayValues();
    data.forEach(function (r) {
      if (String(r[TRK_SCHED.active - 1] || '').trim().toLowerCase() !== 'yes') return;
      rows.push({
        type: r[TRK_SCHED.vaccineType - 1].trim(),
        brand: r[TRK_SCHED.brand - 1].trim(),
        scheduleCode: r[TRK_SCHED.scheduleCode - 1].trim(),
        scheduleDisplay: r[TRK_SCHED.scheduleDisplay - 1],
        scheduleType: r[TRK_SCHED.scheduleType - 1],
        currentDose: Number(r[TRK_SCHED.currentDose - 1]),
        nextAction: r[TRK_SCHED.nextAction - 1],
        intervalValue: Number(r[TRK_SCHED.intervalValue - 1]),
        intervalUnit: String(r[TRK_SCHED.intervalUnit - 1] || '').trim(),
        totalSeriesDoses: Number(r[TRK_SCHED.totalSeriesDoses - 1])
      });
    });
  }
  return rows;
}

function trkListRule_(values) {
  return values && values.length ? SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true).setAllowInvalid(false).build() : null;
}

function trkRefreshBrandOptions_(sheet, ss, first, last) {
  const schedule = trkGetActiveScheduleData_(ss);
  const byType = new Map();
  schedule.forEach(function (r) {
    if (!byType.has(r.type)) byType.set(r.type, new Set());
    byType.get(r.type).add(r.brand);
  });
  const types = sheet.getRange(first, TRK.vaccineType, last - first + 1, 1).getDisplayValues();
  const rules = types.map(function (r) {
    const brands = byType.get(r[0]);
    return [trkListRule_(brands ? [...brands].sort((a, b) => a.localeCompare(b)) : null)];
  });
  sheet.getRange(first, TRK.brand, rules.length, 1).setDataValidations(rules);
}

function trkRefreshScheduleCodeOptions_(sheet, ss, first, last) {
  const schedule = trkGetActiveScheduleData_(ss);
  const byTypeBrand = new Map();
  schedule.forEach(function (r) {
    const key = r.type + '\u0001' + r.brand;
    if (!byTypeBrand.has(key)) byTypeBrand.set(key, new Set());
    byTypeBrand.get(key).add(r.scheduleCode);
  });
  const typesAndBrands = sheet.getRange(first, TRK.vaccineType, last - first + 1, TRK.brand - TRK.vaccineType + 1).getDisplayValues();
  const rules = typesAndBrands.map(function (r) {
    const key = r[0] + '\u0001' + (r[TRK.brand - TRK.vaccineType] || '');
    const codes = byTypeBrand.get(key);
    return [trkListRule_(codes ? [...codes].sort((a, b) => a.localeCompare(b)) : null)];
  });
  sheet.getRange(first, TRK.scheduleCode, rules.length, 1).setDataValidations(rules);
}

// ============================================================
//  DISPOSITION -> REASON CATEGORY -> SPECIFIC REASON, reusing
//  Disposition_Master — the exact table already driving this same
//  cascade on Vaccination_Entry via its own array formula. 'Administered'
//  has no reason to pick (mirrors onEdit's existing B14:B15 clear logic
//  for Vaccination_Entry) so it gets no Reason Category options at all.
// ============================================================
function trkGetDispositionData_(ss) {
  const sheet = ss.getSheetByName(TRK_DISP.sheet);
  if (!sheet) throw new Error('Missing worksheet: ' + TRK_DISP.sheet);
  const last = sheet.getLastRow();
  const rows = [];
  if (last >= TRK_DISP.firstRow) {
    const data = sheet.getRange(TRK_DISP.firstRow, 1, last - TRK_DISP.firstRow + 1, TRK_DISP.specificReason).getDisplayValues();
    data.forEach(function (r) {
      if (!r[TRK_DISP.disposition - 1]) return;
      rows.push({
        disposition: r[TRK_DISP.disposition - 1].trim(),
        reasonCategory: r[TRK_DISP.reasonCategory - 1].trim(),
        specificReason: r[TRK_DISP.specificReason - 1].trim()
      });
    });
  }
  return rows;
}

function trkRefreshReasonCategoryOptions_(sheet, ss, first, last) {
  const disp = trkGetDispositionData_(ss);
  const byDisposition = new Map();
  disp.forEach(function (r) {
    if (!byDisposition.has(r.disposition)) byDisposition.set(r.disposition, new Set());
    if (r.reasonCategory) byDisposition.get(r.disposition).add(r.reasonCategory);
  });
  const dispositions = sheet.getRange(first, TRK.disposition, last - first + 1, 1).getDisplayValues();
  const rules = dispositions.map(function (r) {
    if (r[0] === 'Administered' || !r[0]) return [null];
    const cats = byDisposition.get(r[0]);
    return [trkListRule_(cats ? [...cats].sort((a, b) => a.localeCompare(b)) : null)];
  });
  sheet.getRange(first, TRK.reasonCategory, rules.length, 1).setDataValidations(rules);
}

function trkRefreshSpecificReasonOptions_(sheet, ss, first, last) {
  const disp = trkGetDispositionData_(ss);
  const byKey = new Map();
  disp.forEach(function (r) {
    const key = r.disposition + '\u0001' + r.reasonCategory;
    if (!byKey.has(key)) byKey.set(key, new Set());
    if (r.specificReason) byKey.get(key).add(r.specificReason);
  });
  const dispAndCat = sheet.getRange(first, TRK.disposition, last - first + 1, TRK.reasonCategory - TRK.disposition + 1).getDisplayValues();
  const rules = dispAndCat.map(function (r) {
    const key = r[0] + '\u0001' + (r[TRK.reasonCategory - TRK.disposition] || '');
    const reasons = byKey.get(key);
    return [trkListRule_(reasons ? [...reasons].sort((a, b) => a.localeCompare(b)) : null)];
  });
  sheet.getRange(first, TRK.specificReason, rules.length, 1).setDataValidations(rules);
}

// ============================================================
//  FOLLOW-UP COMPUTATION — ported from previewFollowUpSchedule_ in
//  Code.gs, same matching rule (Vaccine Type + Schedule Code + the
//  numeric position derived from Dose Number), same interval math, same
//  Series Status text. This is the piece that makes a backfilled record
//  a COMPLETE record instead of one with blank Next Dose/Recommended
//  Date/Series Status — deliberately kept logically identical to the
//  web app's version rather than "close enough", since anything reading
//  these columns later shouldn't be able to tell which path wrote them.
// ============================================================
function trkDeriveCurrentDoseNumber_(doseNumberText) {
  if (!doseNumberText) return null;
  const text = String(doseNumberText).trim();
  const m = /^Dose\s+(\d+)$/i.exec(text);
  if (m) return parseInt(m[1], 10);
  if (/^(Annual Dose|Seasonal Dose)$/i.test(text)) return 1;
  return null;
}

function trkComputeFollowUp_(schedule, vaccineType, scheduleCode, doseNumberText, vaccinationDate) {
  const currentNum = trkDeriveCurrentDoseNumber_(doseNumberText);
  if (currentNum === null) return { nextDose: '', recommendedDate: '', reminderDate: '', seriesStatus: '' };

  let matchRow = null;
  for (let i = 0; i < schedule.length; i++) {
    const r = schedule[i];
    if (r.type !== vaccineType || r.scheduleCode !== scheduleCode) continue;
    if (r.currentDose === currentNum) { matchRow = r; break; }
  }
  if (!matchRow) return { nextDose: '', recommendedDate: '', reminderDate: '', seriesStatus: '' };

  const nextDose = matchRow.nextAction || '';
  let recommendedDate = '';
  if (vaccinationDate && nextDose && !isNaN(matchRow.intervalValue) && matchRow.intervalUnit) {
    const baseDate = new Date(vaccinationDate);
    if (!isNaN(baseDate.getTime())) {
      let d = new Date(baseDate.getTime());
      if (matchRow.intervalUnit === 'Days') d.setDate(d.getDate() + matchRow.intervalValue);
      else if (matchRow.intervalUnit === 'Weeks') d.setDate(d.getDate() + matchRow.intervalValue * 7);
      else if (matchRow.intervalUnit === 'Months') d.setMonth(d.getMonth() + matchRow.intervalValue);
      else if (matchRow.intervalUnit === 'Years') d.setFullYear(d.getFullYear() + matchRow.intervalValue);
      else d = null;
      if (d) recommendedDate = d;
    }
  }

  let reminderDate = '';
  if (recommendedDate) {
    reminderDate = new Date(recommendedDate.getTime());
    reminderDate.setDate(reminderDate.getDate() - 7);
  }

  // Mirrors previewFollowUpSchedule_'s COMPLETE_VALUES exactly.
  const COMPLETE_VALUES = ['Complete', 'Complete / Future Booster', 'Complete / Future Guidance', 'Complete for Current Pregnancy'];
  let seriesStatus = '';
  if (nextDose) {
    if (nextDose === 'Next Seasonal Dose') seriesStatus = 'Complete - Recurring';
    else if (COMPLETE_VALUES.indexOf(nextDose) !== -1) seriesStatus = 'Complete';
    else seriesStatus = 'In Progress';
  }

  return { nextDose: nextDose, recommendedDate: recommendedDate, reminderDate: reminderDate, seriesStatus: seriesStatus };
}

function trkRecomputeFollowUp_(sheet, ss, first, last) {
  const schedule = trkGetActiveScheduleData_(ss);
  const cols = [TRK.vaccineType, TRK.scheduleCode, TRK.doseNumber, TRK.vaccinationDate];
  const minCol = Math.min.apply(null, cols), maxCol = Math.max.apply(null, cols);
  const data = sheet.getRange(first, minCol, last - first + 1, maxCol - minCol + 1).getDisplayValues();
  for (let i = 0; i < data.length; i++) {
    const row = first + i;
    const vaccineType = data[i][TRK.vaccineType - minCol];
    const scheduleCode = data[i][TRK.scheduleCode - minCol];
    const doseNumberText = data[i][TRK.doseNumber - minCol];
    const vaccinationDate = data[i][TRK.vaccinationDate - minCol];
    if (!vaccineType || !scheduleCode || !doseNumberText) continue;
    const result = trkComputeFollowUp_(schedule, vaccineType, scheduleCode, doseNumberText, vaccinationDate);
    sheet.getRange(row, TRK.nextDose).setValue(result.nextDose);
    sheet.getRange(row, TRK.recommendedDate).setValue(result.recommendedDate || '');
    sheet.getRange(row, TRK.reminderDate).setValue(result.reminderDate || '');
    sheet.getRange(row, TRK.seriesStatus).setValue(result.seriesStatus);
  }
}

// ============================================================
//  LOT INVENTORY CLAIM (closes the "Backfill double-claim" gap) —
//  reuses claimLotInventoryForBoundSheet_ from MasterTemplate_Code.gs
//  (same bound Apps Script project, so this is a normal function call,
//  not a duplicated implementation — that function already carries the
//  full province-alias-aware region matching the web app uses, ported
//  once already for the Sheets-native submitVaccination() path; no need
//  to port it a second time here).
//
//  WHY THIS EXISTS: before this, a dose entered through Backfill never
//  touched Lot_Expiry_Master at all. The physical unit it actually used
//  stayed marked "Ready" (or whatever it already was) forever, so the
//  live web app could still offer that same already-used unit to a
//  different, real patient later — a real double-claim risk, not a
//  hypothetical one.
//
//  MIRRORS THE WEB APP'S OWN RULE exactly: "Administered" -> "Used Up"
//  (permanently removed from selection), anything else (Deferred,
//  Declined, No-show, Contraindicated) -> "Hold" (still selectable —
//  the physical dose was never actually given).
//
//  IDEMPOTENT, using the EXISTING Session ID / Order Ref column — no new
//  column added anywhere. Before attempting a claim, this checks whether
//  any Lot_Expiry_Master row already has THIS Tracker row's own Record ID
//  stamped in Session ID / Order Ref; if so, this row was already claimed
//  by an earlier edit and is left alone. Also means: editing a row's Lot
//  Number AFTER it was already successfully claimed does NOT re-claim a
//  different unit or release the original one — same limitation the live
//  web app already has (its own Vaccination_Tracker rows aren't
//  editable-and-reclaimable after submission either). If an admin needs
//  to correct a backfilled row's Lot Number after the fact, the original
//  Lot_Expiry_Master row needs fixing by hand too.
//
//  Only ATTEMPTS a claim once a row has enough filled in to identify a
//  target unit (Vaccine Type, Brand, Lot Number, Vaccination Location,
//  Disposition, and — implicitly, since trkAssignRecordIds_ always runs
//  first in vxSyncTrackerEdit — a Record ID). A row still being filled in
//  is silently left alone, not flagged as an error.
//
//  ON FAILURE this does NOT throw/block — there's no "save" to abort
//  here, the cell edit already happened. What happens next depends on
//  WHY it failed:
//
//    - Lot Number doesn't exist anywhere in Lot_Expiry_Master at all
//      (result.notFound): treated as a plain historical record, not an
//      error. Backfilled/bulk-imported rows can legitimately reference a
//      dose that was never PQ HealthShield's own live-tracked stock — a
//      different provider entirely, or an older PQH program that
//      predates live lot tracking for that lot. Since it was never
//      entered through the live entry form's inventory-aware flow, there
//      is nothing to reconcile it against. No note, no toast — the row
//      is simply left untouched.
//    - Every other failure (wrong region, missing region text in
//      Vaccination Location, or the lot already recorded Used Up): the
//      lot DOES exist in Lot_Expiry_Master, so this is a genuine
//      data-entry or logistics mismatch worth surfacing. It's flagged
//      two ways so it's never silently lost: an immediate toast (if
//      someone's actively looking at the sheet right now) AND a
//      persistent note on that row's Lot Number cell (so it's still
//      discoverable later, e.g. after a bulk paste nobody was watching).
//
//  A later SUCCESSFUL claim (e.g. after the Lot Number gets corrected)
//  clears any old note either way.
// ============================================================
function trkClaimLotInventory_(sheet, ss, first, last) {
  const lotSheet = ss.getSheetByName('Lot_Expiry_Master');
  if (!lotSheet) return; // client doesn't use live lot inventory — nothing to do, same opt-in behavior as everywhere else this feature touches

  const cols = [TRK.recordId, TRK.disposition, TRK.vaccineType, TRK.brand, TRK.lotNumber, TRK.vaccinationDate, TRK.location];
  const minCol = Math.min.apply(null, cols), maxCol = Math.max.apply(null, cols);
  const data = sheet.getRange(first, minCol, last - first + 1, maxCol - minCol + 1).getDisplayValues();

  for (let i = 0; i < data.length; i++) {
    const row = first + i;
    const recordId = data[i][TRK.recordId - minCol];
    const disposition = data[i][TRK.disposition - minCol];
    const vaccineType = data[i][TRK.vaccineType - minCol];
    const brand = data[i][TRK.brand - minCol];
    const lotNumber = data[i][TRK.lotNumber - minCol];
    const vaccinationDate = data[i][TRK.vaccinationDate - minCol];
    const location = data[i][TRK.location - minCol];
    if (!recordId || !disposition || !vaccineType || !brand || !lotNumber || !location) continue; // not complete enough yet

    if (trkLotAlreadyClaimedBy_(lotSheet, recordId)) continue; // idempotent — already claimed by this exact row

    const finalStatus = disposition === 'Administered' ? 'Used Up' : 'Hold';
    const result = claimLotInventoryForBoundSheet_(ss, vaccineType, brand, lotNumber, location, recordId, vaccinationDate, finalStatus);

    const lotNumberCell = sheet.getRange(row, TRK.lotNumber);
    if (result && result.ok) {
      lotNumberCell.clearNote(); // clear any stale warning from a previously-failed attempt on this row
    } else if (result && result.notFound) {
      // Lot Number doesn't exist anywhere in Lot_Expiry_Master. For a
      // backfilled/bulk-imported row this is expected and NOT an error:
      // the dose may have been given by a different provider entirely, or
      // by PQ HealthShield itself under an older program that predates
      // live lot tracking. Either way it was never entered through the
      // live entry form's inventory-aware flow, so there is nothing to
      // reconcile against — this row is just a historical record, left
      // completely untouched. No note, no toast, no claim side effect.
      //
      // KNOWN TRADEOFF: this also silences a plain typo of a real Lot
      // Number — there is no way to tell "genuinely external dose" from
      // "fat-fingered an existing lot" from inside this function alone,
      // since both look identical here (no match found). That's not lost
      // information though: getInventoryBalanceReport (Code.gs) sources
      // Administered counts from Vaccination_Tracker independently of
      // this claim step, so ANY (site, lot) with administered doses but
      // zero Lot_Expiry_Master rows — typo or genuine external record —
      // still shows up there under the noLotRowsButAdministered flag. If
      // inventory math ever looks off, that report is where to look
      // first; a human still has to judge whether a flagged lot text is
      // a near-miss of a real one or clearly a different system's lot.
      lotNumberCell.clearNote(); // also clears any stale note from before this distinction existed
    } else if (result && !result.skipped) {
      // Every other failure (wrong region, missing region text, lot
      // already Used Up) means the lot DOES exist in Lot_Expiry_Master
      // and something about this row genuinely doesn't line up — that's
      // still worth flagging, unlike the notFound case above.
      const message = 'Lot inventory claim failed for this backfilled record: ' + result.message;
      lotNumberCell.setNote(message);
      try {
        SpreadsheetApp.getActiveSpreadsheet().toast(message, 'Vaccination_Tracker — Lot Claim', 10);
      } catch (err) {
        // toast() can be unavailable in some trigger contexts — the cell
        // note above is the durable record either way, so this is fine
        // to swallow rather than let it break the rest of the edit.
      }
    }
  }
}

// Session ID / Order Ref is column A (LOT.SESSION_REF in the standalone
// Code.gs's own LOT map — same column here, header row 5, data row 6).
function trkLotAlreadyClaimedBy_(lotSheet, recordId) {
  const FIRST_ROW = 6;
  const lastRow = lotSheet.getLastRow();
  if (lastRow < FIRST_ROW) return false;
  const refs = lotSheet.getRange(FIRST_ROW, 1, lastRow - FIRST_ROW + 1, 1).getDisplayValues();
  return refs.some(function (r) { return r[0] === recordId; });
}

// ============================================================
//  STATIC / MASTER-SOURCED DROPDOWN SETUP — call once (or let onOpen
//  call it) to install every Tracker column's validation across all
//  rows: Recipient ID (from Recipients_Master), Disposition/Admin
//  Site/Route/Dose Number (fixed lists), Vaccinator (from
//  Vaccinators_Master), plus Vaccine Type and the cascading
//  Brand/Schedule Code/Reason Category/Specific Reason columns already
//  handled per-row above.
// ============================================================
function trkRefreshAllDropdowns_(ss) {
  const sheet = ss.getSheetByName(TRK.sheet);
  if (!sheet) throw new Error('Missing worksheet: ' + TRK.sheet);
  const count = sheet.getMaxRows() - TRK.firstRow + 1;
  if (count < 1) return;
  const first = TRK.firstRow, last = sheet.getMaxRows();

  const recIds = [...trkGetRecipientsById_(ss).keys()].sort((a, b) => a.localeCompare(b));
  sheet.getRange(first, TRK.recipientId, count, 1).setDataValidation(trkListRule_(recIds));
  sheet.setColumnWidth(TRK.recipientId, TRK_RECIPIENT_ID_COL_WIDTH);

  const vaccNames = [...trkGetVaccinatorsByDisplayName_(ss).keys()].sort((a, b) => a.localeCompare(b));
  sheet.getRange(first, TRK.vaccinator, count, 1).setDataValidation(trkListRule_(vaccNames));

  const schedule = trkGetActiveScheduleData_(ss);
  const types = [...new Set(schedule.map(function (r) { return r.type; }))].sort((a, b) => a.localeCompare(b));
  sheet.getRange(first, TRK.vaccineType, count, 1).setDataValidation(trkListRule_(types));

  sheet.getRange(first, TRK.disposition, count, 1).setDataValidation(trkListRule_(TRK_DISPOSITION_OPTIONS));
  // Both dose-number columns share the same fixed 6-value list: Intended
  // Dose Number (what was planned) and Dose Number (what was actually
  // given, and the one trkRecomputeFollowUp_ reads). Missing this line
  // for Intended Dose Number was a real gap — free-typing that column
  // let anything through, including values that don't match this list at
  // all, with zero connection to the fact that Dose Number was already a
  // proper dropdown right next to it.
  sheet.getRange(first, TRK.intendedDose, count, 1).setDataValidation(trkListRule_(TRK_DOSE_NUMBER_OPTIONS));
  sheet.getRange(first, TRK.doseNumber, count, 1).setDataValidation(trkListRule_(TRK_DOSE_NUMBER_OPTIONS));
  sheet.getRange(first, TRK.adminSite, count, 1).setDataValidation(trkListRule_(TRK_ADMIN_SITE_OPTIONS));
  sheet.getRange(first, TRK.route, count, 1).setDataValidation(trkListRule_(TRK_ROUTE_OPTIONS));

  trkRefreshBrandOptions_(sheet, ss, first, last);
  trkRefreshScheduleCodeOptions_(sheet, ss, first, last);
  trkRefreshReasonCategoryOptions_(sheet, ss, first, last);
  trkRefreshSpecificReasonOptions_(sheet, ss, first, last);
}

// Manual "force a full setup/refresh now" utility, same role as
// setupVxSyncLotDropdowns() — never required for the automation to
// work, just useful for debugging or a first-time manual sanity check.
function setupVxSyncTrackerBackfill() {
  trkRefreshAllDropdowns_(SpreadsheetApp.getActiveSpreadsheet());
}
