/**************************************************************
 * CORPSHIELD VxSync
 * Vaccination Program Tracking System
 *
 * CURRENT Vaccination_Entry MAP
 * ------------------------------------------------------------
 * B2  Recipient Name
 * B3  Recipient ID
 * B4  Email Address
 * B5  Date of Birth
 * B6  Recipient Category
 * B7  Company / Organization
 * B8  Department / Unit
 * B9  Assigned Site / Location
 *
 * B11 Vaccine Type
 * B12 Intended Dose Number
 * B13 Dose Disposition
 * B14 Reason Category
 * B15 Specific Reason
 * B16 Status Date / Assessment Date
 * B17 Review / Reschedule Date
 *
 * B19 Vaccine Brand
 * B20 Schedule
 * B21 Vaccination Date
 * B22 Lot Number
 * B23 Expiry Date
 * B24 Administration Site
 * B25 Route of Administration
 * B26 Vaccination Location
 * B27 Vaccinator
 * B28 Vaccinator License Number
 * B29 Remarks
 *
 * B31 Next Dose
 * B32 Recommended Next Dose Date
 * B33 Reminder Date
 * B34 Series Status
 * B35 Next Appointment Date
 **************************************************************/


/**************************************************************
 * ON EDIT
 *
 * Handles:
 * 1. Recipients_Master auto Recipient ID / Date Added / Active
 * 2. Vaccination_Entry disposition workflow
 * 3. Clears stale dependent dropdown values
 **************************************************************/
function onEdit(e) {

  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  const cell = e.range.getA1Notation();


  // ==========================================================
  // RECIPIENTS MASTER
  // ==========================================================
  if (sheetName === 'Recipients_Master') {

    var startRow = e.range.getRow();
    var numRows = e.range.getNumRows();

    if (startRow < 2) {
      // The edited range touched row 1 (header) — if it was a multi-row
      // paste starting there, still process row 2 onward from that same
      // paste instead of ignoring the whole thing.
      numRows -= (2 - startRow);
      startRow = 2;
      if (numRows <= 0) return;
    }

    // Locked for the WHOLE edited range, not just one row: a multi-row
    // paste (the normal way a client-provided list gets pre-registered)
    // fires onEdit ONCE for the entire pasted range, not once per row —
    // the previous version only ever looked at e.range.getRow(), the
    // first row of whatever was pasted, so every row after the first
    // silently got no ID at all.
    var lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      var lastRow = sheet.getLastRow();
      if (lastRow < 2) return;

      var endRow = Math.min(startRow + numRows - 1, lastRow);

      // Scan the id column once for the current max, then hand out
      // sequential numbers locally — one generateRecipientId_-style
      // rescan per row would get slower the larger the pasted list is.
      var allIds = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      var maxNum = 0;
      allIds.forEach(function (idRow) {
        var m = /^VAC-(\d+)$/.exec(String(idRow[0] || ''));
        if (m) maxNum = Math.max(maxNum, parseInt(m[1], 10));
      });

      for (var row = startRow; row <= endRow; row++) {
        var recipientIdCell = sheet.getRange(row, 1);   // A
        var lastName = sheet.getRange(row, 3).getValue();   // C
        var firstName = sheet.getRange(row, 4).getValue();  // D
        var dateAddedCell = sheet.getRange(row, 16);    // P
        var activeCell = sheet.getRange(row, 17);       // Q

        if (!lastName || !firstName) continue;

        if (!recipientIdCell.getValue()) {
          maxNum += 1;
          recipientIdCell.setValue('VAC-' + String(maxNum).padStart(6, '0'));
        }

        if (!dateAddedCell.getValue()) {
          dateAddedCell.setValue(new Date());
          dateAddedCell.setNumberFormat('MM/dd/yyyy');
        }

        // Active (column Q) must be exactly 'Yes' or the standalone
        // Code.gs's getRecipients() silently filters this row out of the
        // entry form's recipient search (isYes_(row[16]) check) — no
        // error, it just never shows up. The in-app "Add Recipient"
        // modal already sets this itself; a recipient typed straight
        // into the Sheet only ever got an ID + Date Added from this
        // trigger, never Active, which is exactly what made this
        // invisible. Only fills it in when blank, so a recipient
        // someone has deliberately marked inactive is never silently
        // reactivated by this trigger.
        if (!activeCell.getValue()) {
          activeCell.setValue('Yes');
        }
      }
    } finally {
      lock.releaseLock();
    }

    // A recipient row changed (new one added, ID edited elsewhere) —
    // Vaccination_Tracker's Recipient ID dropdown is fed straight from
    // this table, see VxSync_Tracker_Backfill.gs.
    vxSyncTrackerSourceEdit(e, TRK_REC.sheet);

    return;
  }


  // ==========================================================
  // LOT INVENTORY DROPDOWNS (Lot_Expiry_Master / Vaccine_Schedule_Master)
  // See VxSync_Lot_Dropdown.gs — kept as its own file for clarity, but
  // invoked directly from THIS shared bare onEdit(e) simple trigger
  // rather than through a second, separately-installed trigger. A
  // simple trigger needs no authorization and survives a Drive copy
  // automatically, exactly like the rest of this function already
  // does — that's what makes the dropdown refresh work for every
  // future auto-provisioned client with zero manual setup, not just a
  // one-click shortcut to the same manual setup.
  // ==========================================================
  if (sheetName === VX_LOT.lots || sheetName === VX_LOT.master) {
    vxSyncLotDropdownEdit(e);
    // Vaccine_Schedule_Master also feeds Vaccination_Tracker's own
    // Brand/Schedule Code cascade and its Next Dose/Recommended
    // Date/Series Status computation — see VxSync_Tracker_Backfill.gs.
    // Lot_Expiry_Master edits don't affect Tracker at all, so this is a
    // no-op (fast return inside the function) for that sheet.
    if (sheetName === VX_LOT.master) vxSyncTrackerSourceEdit(e, TRK_SCHED.sheet);
    return;
  }


  // ==========================================================
  // VACCINATION TRACKER BACKFILL (manual/historical record entry)
  // See VxSync_Tracker_Backfill.gs. Handles Record ID auto-generation,
  // Recipient ID -> prefill, Vaccine Type -> Brand -> Schedule Code
  // cascade, Disposition -> Reason Category -> Specific Reason cascade,
  // Vaccinator -> License prefill, and Next Dose/Recommended
  // Date/Reminder Date/Series Status computation — the same derived
  // fields the web app computes at submission time, so a hand-entered
  // historical record ends up complete, not half-finished.
  // ==========================================================
  if (sheetName === TRK.sheet) {
    vxSyncTrackerEdit(e);
    return;
  }
  if (sheetName === TRK_DISP.sheet) {
    vxSyncTrackerSourceEdit(e, TRK_DISP.sheet);
    return;
  }
  if (sheetName === TRK_VACC.sheet) {
    vxSyncTrackerSourceEdit(e, TRK_VACC.sheet);
    return;
  }


  // ==========================================================
  // VACCINATION ENTRY
  // ==========================================================
  if (sheetName !== 'Vaccination_Entry') return;
  ensureDispositionValidations_();


  // ----------------------------------------------------------
  // VACCINE TYPE CHANGED
  // Clear stale Brand + Schedule
  // ----------------------------------------------------------
  if (cell === 'B11') {

    sheet.getRange('B19:B20').clearContent();

    return;
  }


  // ----------------------------------------------------------
  // DOSE DISPOSITION CHANGED
  // ----------------------------------------------------------
  if (cell === 'B13') {

    const disposition = String(e.value || '').trim();

    // Always remove previous patient's / previous selection's
    // reason values.
    sheet.getRange('B14:B15').clearContent();

    // Stamp assessment/status date if blank.
    if (disposition && !sheet.getRange('B16').getValue()) {
      sheet.getRange('B16')
        .setValue(new Date())
        .setNumberFormat('mmmm d, yyyy');
    }

    // ADMINISTERED
    if (disposition === 'Administered') {

      // No reason or reschedule information required.
      sheet.getRange('B14:B15').clearContent();
      sheet.getRange('B17').clearContent();

    } else {

      // A non-administered disposition must not retain
      // vaccination-product details from an earlier patient.
      sheet.getRange('B19:B27').clearContent();

      // B28 License No. may contain a lookup formula — preserve.
      sheet.getRange('B29').clearContent();

      // Next Appointment applies to succeeding administered dose.
      sheet.getRange('B35').clearContent();
    }

    return;
  }


  // ----------------------------------------------------------
  // REASON CATEGORY CHANGED
  // Clear stale Specific Reason
  // ----------------------------------------------------------
  if (cell === 'B14') {

    sheet.getRange('B15').clearContent();

    return;
  }


  // ----------------------------------------------------------
  // VACCINE BRAND CHANGED
  // Clear stale Schedule
  // ----------------------------------------------------------
  if (cell === 'B19') {

    sheet.getRange('B20').clearContent();

    return;
  }
}



/**************************************************************
 * LIVE LOT INVENTORY — ENFORCED FROM THIS SHEETS-NATIVE PATH TOO
 *
 * Context: the web app (EntryForm.html + the separate STANDALONE Code.gs
 * it talks to — NOT this bound script) is the intended, primary, and
 * expected-to-be-only way encoders submit vaccination records. It
 * already enforces Live Lot Inventory end-to-end: region-scoped Vaccine
 * Type -> Brand -> Lot Number dropdowns, single-option auto-fill, and an
 * atomic claim that flips a lot to "Used Up" (or "Hold" for a
 * non-administered disposition) the moment it's used.
 *
 * This bound script's submitVaccination() is a separate, Sheets-native
 * path that predates that feature and has no dropdown, no cascade, no
 * UI of its own. It's expected to be used rarely if ever — but "rarely"
 * is not "never," and once a client turns Live Lot Inventory on, the
 * INVARIANT that matters (a real dose is never recorded as Administered
 * against a lot that's already gone, or against the wrong region's
 * stock) has to hold no matter which path someone used to submit it.
 * A softer "best effort, let unmatched lots through" version of this
 * check would leave that invariant broken for anyone using this path —
 * which defeats the entire point of turning the feature on for that
 * client — so this does NOT take that shortcut.
 *
 * Behavior once Lot_Expiry_Master has any rows for this client:
 *  - Vaccination Location (B26) has no "- REGION" suffix -> BLOCKED.
 *    Same rule the web app enforces on every submission, applied here
 *    for the same reason: nothing downstream can be matched to a region
 *    that was never stated.
 *  - No row in Lot_Expiry_Master matches this exact Vaccine Type /
 *    Brand / Lot Number, in this exact region -> BLOCKED. An untracked
 *    or mistyped lot is no longer waved through — it gets the same
 *    treatment the web app gives it (it simply wouldn't appear in that
 *    dropdown at all).
 *  - A matching row exists but only in a DIFFERENT region -> BLOCKED,
 *    with a message saying so specifically (distinct from "not found
 *    at all", since it's almost always a Vaccination Location typo).
 *  - A matching, correct-region row is found but already "Used Up" ->
 *    BLOCKED — the specific case that started this conversation.
 *  - A matching, correct-region row is found and still selectable
 *    (blank / Ready / Hold) -> claimed: stamped with this Vaccination
 *    Record ID and date, flipped to "Used Up".
 *
 * Deliberately NOT replicated here (this is an integrity gate, not a
 * UX rebuild): the dropdown/cascade/auto-fill experience, and the
 * shared PROVINCE_ALIASES table the web app uses to bucket spelling
 * variants (e.g. "QC" / "Quezon City") into one region — region text
 * is normalized through the SAME province-alias bucketing the web app
 * uses (see BOUND_PROVINCE_ALIASES below) — "QC" and "Metro Manila" are
 * treated as the same region here too, not just an exact literal-text
 * match. An earlier version of this file compared literal text only;
 * that was a real bug (a lot logged under "Metro Manila" would never
 * match a site typed as "...- QC"), fixed below.
 *
 * Column layout matches the standalone Code.gs's LOT map exactly
 * (Lot_Expiry_Master header row 5, data starting row 6):
 *   A Session ID/Order Ref, B Vaccination Date, C Site of Vaccination,
 *   D Vaccine Type, E Vaccine Brand, F Lot Number, G Expiry Date,
 *   H Availability.
 **************************************************************/

// ============================================================
// PROVINCE ALIASES — DUPLICATED from the standalone Code.gs (the
// "VxSync — [ClientName]" project deployed as the web app, NOT this
// bound script). Apps Script projects can't share code between them, so
// this is a manual copy, not a shared library. IF the standalone
// Code.gs's PROVINCE_ALIASES table is ever updated (a new alias added,
// a typo-correction added), this table needs the identical update by
// hand, or region matching will quietly drift apart between the web app
// and this bound script again — exactly the bug this replaces. Kept as
// one block so a future sync is a straight copy-paste of the object
// literal, not a diff to reconstruct by memory.
// ============================================================
const BOUND_PROVINCE_ALIASES = {
  "metro manila": "Metro Manila", "mm": "Metro Manila", "ncr": "Metro Manila",
  "qc": "Metro Manila", "q.c.": "Metro Manila", "quezon city": "Metro Manila",
  "manila": "Metro Manila", "makati": "Metro Manila", "makati city": "Metro Manila",
  "pasig": "Metro Manila", "pasig city": "Metro Manila", "marikina": "Metro Manila",
  "marikina city": "Metro Manila", "taguig": "Metro Manila", "bgc": "Metro Manila",
  "paranaque": "Metro Manila", "paranaque city": "Metro Manila", "pasay": "Metro Manila",
  "pasay city": "Metro Manila", "alabang": "Metro Manila", "muntinlupa": "Metro Manila",
  "mandaluyong": "Metro Manila", "san juan": "Metro Manila", "caloocan": "Metro Manila",
  "malabon": "Metro Manila", "navotas": "Metro Manila", "valenzuela": "Metro Manila",
  "las pinas": "Metro Manila", "laspinas": "Metro Manila", "camanava": "Metro Manila",
  "qc north": "Metro Manila",
  "davao del sur": "Davao del Sur", "daval del sur": "Davao del Sur",
  "davao": "Davao del Sur", "davao city": "Davao del Sur",
  "agusan del norte": "Agusan del Norte", "agusan del norte ": "Agusan del Norte",
  "butuan city": "Agusan del Norte", "butuan": "Agusan del Norte",
  "pangasinan": "Pangasinan", "pangsinan": "Pangasinan", "lingayen": "Pangasinan",
  "urdaneta": "Pangasinan",
  "pampanga": "Pampanga", "angeles": "Pampanga", "angeles city": "Pampanga",
  "agneles city": "Pampanga", "agneles": "Pampanga",
  "cavite": "Cavite", "bacoor": "Cavite", "bacoor city": "Cavite", "gma": "Cavite",
  "general mariano alvarez": "Cavite", "dasmarinas": "Cavite", "imus": "Cavite",
  "laguna": "Laguna", "san pablo": "Laguna", "san pablo city": "Laguna",
  "san pedro": "Laguna", "calamba": "Laguna", "sta rosa": "Laguna", "santa rosa": "Laguna",
  "batangas": "Batangas", "lipa": "Batangas", "lipa city": "Batangas",
  "tanauan": "Batangas", "tanauan batangas": "Batangas", "darasa": "Batangas",
  "nueva ecija": "Nueva Ecija", "cabanatuan": "Nueva Ecija", "cabanatuan city": "Nueva Ecija",
  "palawan": "Palawan", "puerto princesa": "Palawan", "puerto princesa city": "Palawan",
  "camarines sur": "Camarines Sur", "naga": "Camarines Sur", "naga city": "Camarines Sur",
  "camarines norte": "Camarines Norte",
  "leyte": "Leyte", "palo": "Leyte", "tacloban": "Leyte", "tacloban city": "Leyte",
  "albay": "Albay", "legazpi": "Albay", "legazpi city": "Albay", "legaspi": "Albay",
  "western visayas": "Iloilo", "iloilo": "Iloilo", "ilo-ilo": "Iloilo", "ilo ilo": "Iloilo",
  "ilo-ilo city": "Iloilo", "iloilo city": "Iloilo",
  "misamis oriental": "Misamis Oriental", "cagayan de oro": "Misamis Oriental",
  "cagayan de oro city": "Misamis Oriental", "cdo": "Misamis Oriental",
  "cagayan": "Cagayan", "tuguegarao": "Cagayan", "tuguegarao city": "Cagayan",
  "isabela": "Isabela", "cauayan": "Isabela", "cauayan city": "Isabela",
  "zamboanga": "Zamboanga", "zamboanga city": "Zamboanga",
  "south cotabato": "South Cotabato", "general santos": "South Cotabato", "gensan": "South Cotabato",
  "bulacan": "Bulacan", "san rafael": "Bulacan", "malolos": "Bulacan",
  "oriental mindoro": "Oriental Mindoro", "calapan": "Oriental Mindoro", "calapan city": "Oriental Mindoro",
  "zambales": "Zambales", "olongapo": "Zambales", "olongapo city": "Zambales",
  "cebu": "Cebu", "cebu city": "Cebu", "mandaue": "Cebu", "lapu-lapu": "Cebu"
};

function boundNormalizeLookupKey_(text) {
  return String(text || '').trim().toLowerCase().replace(/\s+/g, ' ');
}

// Same fallback design as the web app's normalizeProvince_: an
// unrecognized region never hard-fails — it self-buckets under a
// title-cased version of whatever was typed, so a brand-new region
// nobody's typed before still gets a consistent bucket instead of
// silently matching nothing.
function boundNormalizeProvince_(rawText) {
  const key = boundNormalizeLookupKey_(rawText);
  if (!key) return null;
  if (BOUND_PROVINCE_ALIASES[key]) return BOUND_PROVINCE_ALIASES[key];
  return String(rawText).trim().replace(/\w\S*/g, function (w) {
    return w.charAt(0).toUpperCase() + w.slice(1).toLowerCase();
  });
}

// Same "Site Name - PROVINCE" suffix convention as the web app's
// extractProvinceFromVaccinationSite_ (last "-", non-empty text after
// it) — now returns the NORMALIZED province ("Metro Manila"), not the
// raw literal text, so "QC" and "Metro Manila" correctly match.
function extractRegionFromSiteText_(siteText) {
  if (!siteText) return null;
  const idx = String(siteText).lastIndexOf('-');
  if (idx === -1) return null;
  const provincePart = String(siteText).substring(idx + 1).trim();
  if (!provincePart) return null;
  return boundNormalizeProvince_(provincePart);
}

// `finalStatus` defaults to 'Used Up' so the existing call below (this
// Sheets-native path only ever calls this for isAdministered === true)
// keeps behaving exactly as before with no change needed at that call
// site. Added so a second caller — Vaccination_Tracker's Backfill tool,
// see VxSync_Tracker_Backfill.gs's trkClaimLotInventory_ — can pass
// 'Hold' for a non-Administered disposition, matching the SAME
// Used-Up-vs-Hold rule the standalone web app's claimLotInventoryUnit_
// already applies (an Administered dose permanently removes a unit from
// selection; anything else just flags that unit as already-touched
// without excluding it).
function claimLotInventoryForBoundSheet_(ss, vaccineType, vaccineBrand, lotNumber, siteText, recordId, vaccinationDate, finalStatus) {
  const claimedStatus = finalStatus || 'Used Up';

  const sheet = ss.getSheetByName('Lot_Expiry_Master');
  if (!sheet) return { ok: true, skipped: true };

  const FIRST_ROW = 6;
  const lastRow = sheet.getLastRow();
  if (lastRow < FIRST_ROW) return { ok: true, skipped: true };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const data = sheet.getRange(FIRST_ROW, 1, lastRow - FIRST_ROW + 1, 8).getValues();

    const hasAnyInventory = data.some(function (r) { return !!r[3]; }); // column D, Vaccine Type
    if (!hasAnyInventory) return { ok: true, skipped: true };

    const targetRegion = extractRegionFromSiteText_(siteText);
    if (!targetRegion) {
      return {
        ok: false,
        message: 'Live inventory tracking is on for this client. Vaccination Location (B26) must include the region ' +
          'after a dash (e.g. "MTC White Plains - QC") before an Administered dose can be matched against ' +
          'Lot_Expiry_Master. Please correct it and try again.'
      };
    }

    const targetType = String(vaccineType || '').trim();
    const targetBrand = String(vaccineBrand || '').trim();
    const targetLot = String(lotNumber || '').trim();

    let foundWrongRegionOnly = false;

    for (let i = 0; i < data.length; i++) {
      const r = data[i];
      const type = String(r[3] || '').trim();   // D
      const brand = String(r[4] || '').trim();  // E
      const lot = String(r[5] || '').trim();    // F
      if (type !== targetType || brand !== targetBrand || lot !== targetLot) continue;

      const rowRegion = extractRegionFromSiteText_(r[2]); // C Site of Vaccination
      if (rowRegion !== targetRegion) { foundWrongRegionOnly = true; continue; }

      const availability = String(r[7] || '').trim(); // H — blank counts as Ready, same convention as the web app
      if (availability.toLowerCase() === 'used up') {
        return {
          ok: false,
          message: 'This vaccine lot ("' + lotNumber + '") is already recorded as "Used Up" in Lot_Expiry_Master — ' +
            'it cannot be administered again.\n\nPlease verify the Lot Number, or check with logistics.'
        };
      }

      const rowNum = FIRST_ROW + i;
      sheet.getRange(rowNum, 1).setValue(recordId);              // A Session ID / Order Ref
      sheet.getRange(rowNum, 2).setValue(vaccinationDate || '');  // B Vaccination Date
      sheet.getRange(rowNum, 8).setValue(claimedStatus);          // H Availability
      return { ok: true, claimed: true };
    }

    if (foundWrongRegionOnly) {
      return {
        ok: false,
        message: 'This vaccine lot ("' + lotNumber + '") exists in Lot_Expiry_Master, but only for a different region ' +
          'than "' + targetRegion + '". Please verify the Lot Number and Vaccination Location.'
      };
    }

    // Live inventory IS active for this client, and nothing in
    // Lot_Expiry_Master matches this Type/Brand/Lot in this region at
    // all — the web app's dropdown would never have offered it either.
    //
    // `notFound: true` distinguishes this specific case ("this lot simply
    // isn't in Lot_Expiry_Master at all") from every other failure above
    // (wrong region, missing region text, already Used Up) where the lot
    // DOES exist and something else is genuinely off. Callers that only
    // ever see live, dropdown-sourced lot numbers (the web app) never hit
    // this branch anyway, since the dropdown can't offer an unlisted lot.
    // But a caller ingesting bulk/backfilled historical records — where a
    // dose may legitimately be external-provider or predate this client's
    // live lot tracking — can use this flag to treat "no such lot" as an
    // ordinary, silent non-match instead of an error worth surfacing.
    return {
      ok: false,
      notFound: true,
      message: 'Live inventory tracking is on for this client, but no matching Ready/Hold lot was found for\n' +
        vaccineType + ' / ' + vaccineBrand + ' / Lot "' + lotNumber + '" in region "' + targetRegion + '".\n\n' +
        'Please verify the Lot Number, Vaccine Brand, and Vaccination Location, or check with logistics.'
    };
  } finally {
    lock.releaseLock();
  }
}



/**************************************************************
 * SUBMIT VACCINATION / ASSESSMENT RECORD
 **************************************************************/
function submitVaccination() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const entrySheet = ss.getSheetByName('Vaccination_Entry');
  const trackerSheet = ss.getSheetByName('Vaccination_Tracker');
  const scheduleSheet = ss.getSheetByName('Vaccine_Schedule_Master');
  const recipientSheet = ss.getSheetByName('Recipients_Master');


  // ==========================================================
  // VERIFY REQUIRED WORKSHEETS
  // ==========================================================
  if (
    !entrySheet ||
    !trackerSheet ||
    !scheduleSheet ||
    !recipientSheet
  ) {

    ui.alert(
      'Required worksheet missing.\n\n' +
      'Please verify Vaccination_Entry, Vaccination_Tracker, ' +
      'Vaccine_Schedule_Master, and Recipients_Master.'
    );

    return;
  }


  // ==========================================================
  // RECIPIENT INFORMATION
  // ==========================================================
  const recipientDisplayName = entrySheet.getRange('B2').getValue();
  const recipientId = entrySheet.getRange('B3').getValue();
  const recipientName = getCleanRecipientName_(recipientDisplayName);

  const email = entrySheet.getRange('B4').getValue();
  const dob = entrySheet.getRange('B5').getValue();
  const recipientCategory = entrySheet.getRange('B6').getValue();

  const companyOrganization = entrySheet.getRange('B7').getValue();
  const departmentUnit = entrySheet.getRange('B8').getValue();
  const assignedSiteLocation = entrySheet.getRange('B9').getValue();


  // ==========================================================
  // GET DEPARTMENT / UNIT ID
  // ==========================================================
  const recipientData = recipientSheet.getDataRange().getValues();

  let departmentUnitId = '';

  for (let i = 1; i < recipientData.length; i++) {

    if (
      String(recipientData[i][0]).trim() ===
      String(recipientId).trim()
    ) {

      departmentUnitId = recipientData[i][9]; // Column J
      break;
    }
  }


  // ==========================================================
  // DOSE ASSESSMENT & DISPOSITION
  // ==========================================================
  const vaccineType = entrySheet.getRange('B11').getValue();
  const intendedDoseNumber = entrySheet.getRange('B12').getValue();
  const doseDisposition = entrySheet.getRange('B13').getValue();

  let reasonCategory = entrySheet.getRange('B14').getValue();
  let specificReason = entrySheet.getRange('B15').getValue();

  const statusDate = entrySheet.getRange('B16').getValue();
  const reviewRescheduleDate = entrySheet.getRange('B17').getValue();


  // ==========================================================
  // VACCINATION DETAILS
  // ==========================================================
  const vaccineBrand = entrySheet.getRange('B19').getValue();
  const scheduleDisplay = entrySheet.getRange('B20').getValue();

  const vaccinationDate = entrySheet.getRange('B21').getValue();
  const lotNumber = entrySheet.getRange('B22').getValue();
  const expiryDate = entrySheet.getRange('B23').getValue();

  const administrationSite = entrySheet.getRange('B24').getValue();
  const route = entrySheet.getRange('B25').getValue();
  const vaccinationLocation = entrySheet.getRange('B26').getValue();

  const vaccinator = entrySheet.getRange('B27').getValue();
  const professionalLicenseNo = entrySheet.getRange('B28').getValue();

  const remarks = entrySheet.getRange('B29').getValue();


  // ==========================================================
  // FOLLOW-UP & SCHEDULE
  // ==========================================================
  const nextDose = entrySheet.getRange('B31').getValue();
  const recommendedNextDoseDate =
    entrySheet.getRange('B32').getValue();

  let reminderDate = entrySheet.getRange('B33').getValue();

  const seriesStatus = entrySheet.getRange('B34').getValue();
  const scheduledAppointmentDate =
    entrySheet.getRange('B35').getValue();


  // ==========================================================
  // NORMALIZE DISPOSITION
  // ==========================================================
  const disposition =
    String(doseDisposition || '').trim();

  const isAdministered =
    disposition.toLowerCase() === 'administered';


  // Administered vaccinations should never carry stale reasons.
  if (isAdministered) {

    reasonCategory = '';
    specificReason = '';

    entrySheet.getRange('B14:B15').clearContent();
    entrySheet.getRange('B17').clearContent();
  }


  // ==========================================================
  // VALIDATE CORE REQUIRED FIELDS
  // ==========================================================
  const missingFields = [];

  if (!recipientName) {
    missingFields.push('Recipient Name');
  }

  if (!recipientId) {
    missingFields.push('Recipient ID');
  }

  if (!vaccineType) {
    missingFields.push('Vaccine Type');
  }

  if (!intendedDoseNumber) {
    missingFields.push('Intended Dose Number');
  }

  if (!doseDisposition) {
    missingFields.push('Dose Disposition');
  }

  if (!statusDate) {
    missingFields.push('Status Date');
  }


  // ==========================================================
  // ADMINISTERED-SPECIFIC VALIDATION
  // ==========================================================
  if (isAdministered) {

    if (!vaccineBrand) {
      missingFields.push('Vaccine Brand');
    }

    if (!scheduleDisplay) {
      missingFields.push('Schedule');
    }

    if (!vaccinationDate) {
      missingFields.push('Vaccination Date');
    }

    if (!lotNumber) {
      missingFields.push('Lot Number');
    }

    if (!expiryDate) {
      missingFields.push('Expiry Date');
    }

    if (!administrationSite) {
      missingFields.push('Administration Site');
    }

    if (!route) {
      missingFields.push('Route of Administration');
    }

    if (!vaccinationLocation) {
      missingFields.push('Vaccination Location');
    }

    if (!vaccinator) {
      missingFields.push('Vaccinator');
    }

    if (!professionalLicenseNo) {
      missingFields.push('Vaccinator License Number');
    }
  }


  // ==========================================================
  // NON-ADMINISTERED VALIDATION
  // ==========================================================
  if (!isAdministered) {

    if (!reasonCategory) {
      missingFields.push('Reason Category');
    }

    if (!specificReason) {
      missingFields.push('Specific Reason');
    }
  }


  // ==========================================================
  // DISPLAY MISSING FIELDS
  // ==========================================================
  if (missingFields.length > 0) {

    ui.alert(
      'Please complete the following required fields:\n\n' +
      missingFields.join('\n')
    );

    return;
  }


  // ==========================================================
  // VALIDATE EXPIRY DATE
  // Only applies when vaccine was administered
  // ==========================================================
  if (
    isAdministered &&
    expiryDate instanceof Date &&
    vaccinationDate instanceof Date &&
    expiryDate < vaccinationDate
  ) {

    ui.alert(
      'The vaccine expiry date is earlier than the ' +
      'vaccination date.\n\nPlease verify the vaccine details.'
    );

    return;
  }


  // ==========================================================
  // ACTUAL DOSE NUMBER
  //
  // Encoder enters Intended Dose once.
  // Actual Dose Number exists only when administered.
  // ==========================================================
  const doseNumber =
    isAdministered ? intendedDoseNumber : '';


  // ==========================================================
  // FIND SCHEDULE CODE
  //
  // Only needed for an administered vaccination.
  // ==========================================================
  let scheduleCode = '';

  if (isAdministered) {

    const scheduleData =
      scheduleSheet.getDataRange().getValues();

    const masterDoseValue =
      normalizeDoseForMaster_(intendedDoseNumber);

    for (let i = 1; i < scheduleData.length; i++) {

      const row = scheduleData[i];

      const masterVaccineType = row[0];   // A
      const masterBrand = row[2];         // C
      const masterScheduleCode = row[4];  // E
      const masterScheduleName = row[5];  // F
      const masterDose = row[8];          // I
      const active = row[20];             // U

    if (
  normalizeTextForMatch_(masterVaccineType) ===
    normalizeTextForMatch_(vaccineType) &&

  normalizeTextForMatch_(masterBrand) ===
    normalizeTextForMatch_(vaccineBrand) &&

  normalizeTextForMatch_(masterScheduleName) ===
    normalizeTextForMatch_(scheduleDisplay) &&

  String(masterDose).trim() ===
    String(masterDoseValue).trim() &&

  normalizeTextForMatch_(active) === 'yes'
) {
  scheduleCode = masterScheduleCode;
  break;
}

}


    if (!scheduleCode) {

      ui.alert(
        'No matching active schedule was found in ' +
        'Vaccine_Schedule_Master.\n\n' +
        'Please verify Vaccine Type, Vaccine Brand, Schedule, ' +
        'and Intended Dose Number.'
      );

      return;
    }
  }


  // ==========================================================
  // DUPLICATE ADMINISTERED DOSE SAFEGUARD
  //
  // Important:
  // - Applies ONLY to administered doses.
  // - Brand change cannot bypass the safeguard.
  // - Schedule change cannot bypass the safeguard.
  // - Deferred / declined / no-show records are allowed for
  //   the same intended dose.
  // ==========================================================
  if (isAdministered) {

    const duplicate =
      findDuplicateAdministeredDose_(
        trackerSheet,
        recipientId,
        vaccineType,
        intendedDoseNumber,
        vaccinationDate
      );


    if (duplicate.found) {

      ui.alert(
        'Duplicate Dose Detected\n\n' +
        'This recipient already has this vaccination dose ' +
        'recorded as administered.\n\n' +

        'Recipient: ' + recipientName +
        '\nVaccine: ' + vaccineType +
        '\nDose: ' + intendedDoseNumber +
        '\nExisting Record ID: ' + duplicate.recordId +

        '\n\nPlease verify the recipient vaccination history ' +
        'before proceeding.'
      );

      return;
    }
  }


  // ==========================================================
  // GENERATE RECORD ID
  // ==========================================================
  const recordId =
    generateVaccinationRecordId_(trackerSheet);


  // ==========================================================
  // LIVE LOT INVENTORY — ENFORCED HERE TOO
  // See claimLotInventoryForBoundSheet_ above for full scope/rationale.
  // Runs AFTER the duplicate-dose check and BEFORE the confirmation
  // dialog, so a blocked claim aborts immediately rather than wasting
  // the encoder's time confirming a submission that's about to fail.
  // ==========================================================
  if (isAdministered && lotNumber) {

    const lotResult = claimLotInventoryForBoundSheet_(
      ss,
      vaccineType,
      vaccineBrand,
      lotNumber,
      vaccinationLocation,
      recordId,
      vaccinationDate instanceof Date ? vaccinationDate : (vaccinationDate || new Date())
    );

    if (!lotResult.ok) {
      ui.alert(lotResult.message);
      return;
    }
  }


  // ==========================================================
  // DETERMINE SCHEDULE / REMINDER STATUS
  // ==========================================================
  let scheduleStatus = '';
  let reminderStatus = '';


  if (isAdministered) {

    if (
      String(seriesStatus).trim().toLowerCase() ===
      'complete'
    ) {

      scheduleStatus = 'Complete';
      reminderStatus = 'Not Required';

    } else if (scheduledAppointmentDate) {

      scheduleStatus = 'Scheduled';
      reminderStatus = 'Pending';

    } else if (recommendedNextDoseDate) {

      scheduleStatus = 'Due Date Generated';
      reminderStatus = 'Pending';

    } else {

      scheduleStatus = 'Clinical Review';
      reminderStatus = 'Not Required';
    }

  } else {

    const dispositionLower =
      disposition.toLowerCase();


    if (dispositionLower === 'deferred') {

      scheduleStatus =
        reviewRescheduleDate
          ? 'Deferred - Review Scheduled'
          : 'Deferred - Review Pending';

      reminderStatus =
        reviewRescheduleDate
          ? 'Pending'
          : 'Not Scheduled';


    } else if (dispositionLower === 'no-show') {

      scheduleStatus =
        reviewRescheduleDate
          ? 'No-show - Rescheduled'
          : 'No-show - Reschedule Pending';

      reminderStatus =
        reviewRescheduleDate
          ? 'Pending'
          : 'Not Scheduled';


    } else if (
      dispositionLower ===
      'contraindicated - temporary'
    ) {

      scheduleStatus =
        reviewRescheduleDate
          ? 'Temporary Contraindication - Review Scheduled'
          : 'Temporary Contraindication - Review Pending';

      reminderStatus =
        reviewRescheduleDate
          ? 'Pending'
          : 'Not Scheduled';


    } else if (
      dispositionLower ===
      'contraindicated - permanent'
    ) {

      scheduleStatus =
        'Permanent Contraindication';

      reminderStatus =
        'Not Required';


    } else if (
      dispositionLower === 'declined'
    ) {

      scheduleStatus = 'Declined';
      reminderStatus = 'Not Required';


    } else {

      scheduleStatus = disposition;
      reminderStatus = 'Not Required';
    }
  }


  // ==========================================================
  // FOLLOW-UP REMINDER FOR NON-ADMINISTERED RECORD
  //
  // If a Review / Reschedule Date exists and B33 is blank,
  // set reminder 7 days before.
  // ==========================================================
  if (
    !isAdministered &&
    reviewRescheduleDate instanceof Date &&
    !reminderDate
  ) {

    reminderDate =
      new Date(reviewRescheduleDate);

    reminderDate.setDate(
      reminderDate.getDate() - 7
    );
  }


  // ==========================================================
  // CONFIRM BEFORE SAVING
  // ==========================================================
  let confirmationMessage =

    'Recipient: ' + recipientName +
    '\nRecipient ID: ' + recipientId +
    '\nOrganization: ' + companyOrganization +
    '\nDepartment / Unit: ' + departmentUnit +
    '\nAssigned Site: ' + assignedSiteLocation +

    '\n\nVaccine Type: ' + vaccineType +
    '\nIntended Dose: ' + intendedDoseNumber +
    '\nDisposition: ' + doseDisposition +

    '\nStatus Date: ' +
    formatDateForDisplay_(statusDate, ss);


  if (isAdministered) {

    confirmationMessage +=

      '\n\nVaccine Brand: ' + vaccineBrand +
      '\nSchedule: ' + scheduleDisplay +
      '\nVaccination Date: ' +
      formatDateForDisplay_(vaccinationDate, ss) +

      '\nRoute: ' + route +
      '\nAdministration Site: ' + administrationSite +
      '\nVaccinator: ' + vaccinator;

  } else {

    confirmationMessage +=

      '\n\nReason Category: ' + reasonCategory +
      '\nSpecific Reason: ' + specificReason;

    if (reviewRescheduleDate) {

      confirmationMessage +=

        '\nReview / Reschedule Date: ' +
        formatDateForDisplay_(
          reviewRescheduleDate,
          ss
        );
    }
  }


  confirmationMessage +=
    '\n\nSave this VxSync record?';


  const response = ui.alert(
    'Confirm VxSync Record',
    confirmationMessage,
    ui.ButtonSet.YES_NO
  );


  if (response !== ui.Button.YES) {
    return;
  }


  // ==========================================================
  // WRITE TO VACCINATION TRACKER
  //
  // A:AL = 38 columns
  // ==========================================================
  const newRow =
    trackerSheet.getLastRow() + 1;


  const record = [[

    recordId,                       // A
    recipientId,                    // B
    recipientName,                  // C
    email,                          // D
    dob,                            // E
    recipientCategory,              // F

    companyOrganization,            // G
    departmentUnitId,               // H
    departmentUnit,                 // I
    assignedSiteLocation,           // J

    vaccineType,                    // K
    intendedDoseNumber,             // L
    doseDisposition,                // M
    reasonCategory,                 // N
    specificReason,                 // O
    statusDate,                     // P
    reviewRescheduleDate,           // Q

    isAdministered ? vaccineBrand : '',       // R
    doseNumber,                               // S
    isAdministered ? vaccinationDate : '',    // T
    isAdministered ? lotNumber : '',          // U
    isAdministered ? expiryDate : '',         // V

    isAdministered ? administrationSite : '', // W
    isAdministered ? route : '',              // X
    isAdministered ? vaccinationLocation : '',// Y
    isAdministered ? vaccinator : '',         // Z
    isAdministered ? professionalLicenseNo : '', // AA

    scheduleCode,                    // AB
    isAdministered ? nextDose : '',  // AC

    isAdministered
      ? recommendedNextDoseDate
      : '',                          // AD

    isAdministered
      ? scheduledAppointmentDate
      : reviewRescheduleDate,        // AE

    reminderDate,                    // AF

    scheduleStatus,                  // AG
    reminderStatus,                  // AH
    '',                              // AI Reminder Sent Date

    isAdministered
      ? seriesStatus
      : 'In Progress',               // AJ

    'No',                            // AK Schedule Override
    remarks                          // AL
  ]];


  trackerSheet
    .getRange(
      newRow,
      1,
      1,
      record[0].length
    )
    .setValues(record);


  // ==========================================================
  // FORMAT TRACKER DATES
  // ==========================================================
  const dateColumns = [
    5,   // E DOB
    16,  // P Status Date
    17,  // Q Review / Reschedule
    20,  // T Vaccination Date
    22,  // V Expiry Date
    30,  // AD Recommended Next Dose
    31,  // AE Scheduled Appointment
    32   // AF Reminder Date
  ];


  dateColumns.forEach(column => {

    trackerSheet
      .getRange(newRow, column)
      .setNumberFormat('mmm d, yyyy');
  });


  // ==========================================================
  // CLEAR ENTRY FORM
  // ==========================================================
  clearVaccinationEntry_();


  // ==========================================================
  // SUCCESS MESSAGE
  // ==========================================================
  ui.alert(
    'VxSync record successfully saved.\n\n' +
    'Record ID: ' + recordId
  );
}



/**************************************************************
 * FIND DUPLICATE ADMINISTERED DOSE
 *
 * New Tracker columns:
 * B Recipient ID
 * K Vaccine Type
 * L Intended Dose
 * M Disposition
 * S Actual Dose Number
 * T Vaccination Date
 *
 * Historical rows have blank disposition but contain an
 * actual Dose Number + Vaccination Date; these are treated
 * as administered records.
 **************************************************************/
function findDuplicateAdministeredDose_(
  trackerSheet,
  recipientId,
  vaccineType,
  intendedDoseNumber,
  vaccinationDate
) {

  const data =
    trackerSheet.getDataRange().getValues();

  const intended =
    String(intendedDoseNumber || '').trim()
      .toLowerCase();


  for (let i = 1; i < data.length; i++) {

    const row = data[i];

    const existingRecordId = row[0];     // A
    const existingRecipientId = row[1];  // B
    const existingVaccineType = row[10]; // K

    const existingIntendedDose = row[11];// L
    const existingDisposition = row[12]; // M

    const existingActualDose = row[18];  // S
    const existingVaccDate = row[19];    // T


    const sameRecipient =
      String(existingRecipientId).trim() ===
      String(recipientId).trim();

    const sameVaccine =
      String(existingVaccineType).trim() ===
      String(vaccineType).trim();


    if (!sameRecipient || !sameVaccine) {
      continue;
    }


    // New records use M = Administered.
    // Historical records have blank M but have an actual
    // vaccination date / actual dose.
    const existingWasAdministered =
      String(existingDisposition).trim()
        .toLowerCase() === 'administered' ||
      (
        !existingDisposition &&
        existingVaccDate
      );


    if (!existingWasAdministered) {
      continue;
    }


    // Prefer actual administered dose. Fall back to intended.
    const existingDose =
      existingActualDose || existingIntendedDose;


    const sameDose =
      normalizeDoseLabel_(existingDose) ===
      normalizeDoseLabel_(intendedDoseNumber);


    if (!sameDose) {
      continue;
    }


    // --------------------------------------------------------
    // ANNUAL DOSES
    //
    // Permit another annual dose in another calendar year.
    // Prevent two annual administrations in the same year.
    // --------------------------------------------------------
    if (intended.includes('annual')) {

      if (
        vaccinationDate instanceof Date &&
        existingVaccDate instanceof Date &&
        vaccinationDate.getFullYear() !==
          existingVaccDate.getFullYear()
      ) {

        continue;
      }
    }


    return {
      found: true,
      recordId: existingRecordId
    };
  }


  return {
    found: false,
    recordId: ''
  };
}



/**************************************************************
 * NORMALIZE DOSE FOR VACCINE SCHEDULE MASTER
 *
 * "Dose 1" -> 1
 * "Dose 2" -> 2
 * "Dose 3" -> 3
 *
 * Other labels such as:
 * Annual Dose / Booster
 * remain unchanged.
 **************************************************************/
/************************************************************
 * NORMALIZE TEXT FOR MASTER MATCHING
 ************************************************************/
function normalizeTextForMatch_(value) {
  return String(value || '')
    .replace(/ /g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function normalizeDoseForMaster_(dose) {

  const value =
    String(dose || '').trim();

  // Standard numbered doses:
  // Dose 1 -> 1
  // Dose 2 -> 2
  // Dose 3 -> 3
  const match =
    value.match(/^Dose\s+(\d+)$/i);

  if (match) {
    return Number(match[1]);
  }

  // Recurring / annual single-dose schedules
  // Vaccine_Schedule_Master stores Current Dose = 1
  if (
    /^Annual Dose$/i.test(value) ||
    /^Seasonal Dose$/i.test(value)
  ) {
    return 1;
  }

  return value;
}



/**************************************************************
 * NORMALIZE DOSE LABEL FOR DUPLICATE COMPARISON
 **************************************************************/
function normalizeDoseLabel_(dose) {

  const value =
    String(dose || '')
      .trim()
      .toLowerCase();

  const match =
    value.match(/^dose\s+(\d+)$/i);

  if (match) {
    return 'dose ' + Number(match[1]);
  }

  // Historical tracker may contain numeric 1, 2, 3
  if (/^\d+$/.test(value)) {
    return 'dose ' + Number(value);
  }

  return value;
}



/**************************************************************
 * CLEAN RECIPIENT DISPLAY NAME
 **************************************************************/
function getCleanRecipientName_(displayName) {

  if (!displayName) return '';

  return String(displayName)
    .replace(
      /\s*[—-]\s*VAC-\d+\s*$/i,
      ''
    )
    .trim();
}



/**************************************************************
 * CLEAR VACCINATION ENTRY
 *
 * Important:
 * Clear CONTENT only.
 * Dropdown validation remains intact.
 * Formula cells are preserved.
 **************************************************************/
function clearVaccinationEntry_() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName('Vaccination_Entry');

  if (!sheet) return;


  const cellsToClear = [

    'B2',   // Recipient Name

    // Dose Assessment
    'B11',  // Vaccine Type
    'B12',  // Intended Dose Number
    'B13',  // Dose Disposition
    'B14',  // Reason Category
    'B15',  // Specific Reason
    'B16',  // Status Date
    'B17',  // Review / Reschedule Date

    // Vaccination Details
    'B19',  // Vaccine Brand
    'B20',  // Schedule
    'B21',  // Vaccination Date
    'B22',  // Lot Number
    'B23',  // Expiry Date
    'B24',  // Administration Site
    'B25',  // Route
    'B26',  // Vaccination Location
    'B27',  // Vaccinator

    // B28 License No. may be a formula — preserve.

    'B29',  // Remarks

    // B31:B34 are formula/system fields — preserve.

    'B35'   // Next Appointment Date
  ];


  cellsToClear.forEach(cell => {

    sheet
      .getRange(cell)
      .clearContent();
  });
}


/**************************************************************
 * GENERATE RECIPIENT ID
 **************************************************************/
function generateRecipientId_(sheet) {

  const lastRow =
    Math.max(sheet.getLastRow(), 2);

  const ids = sheet
    .getRange(
      2,
      1,
      lastRow - 1,
      1
    )
    .getValues()
    .flat();


  let highestNumber = 0;


  ids.forEach(id => {

    const match =
      String(id || '')
        .match(/^VAC-(\d+)$/);

    if (match) {

      highestNumber =
        Math.max(
          highestNumber,
          parseInt(match[1], 10)
        );
    }
  });


  return (
    'VAC-' +
    String(highestNumber + 1)
      .padStart(6, '0')
  );
}



/**************************************************************
 * GENERATE VACCINATION / VxSync RECORD ID
 **************************************************************/
function generateVaccinationRecordId_(
  trackerSheet
) {

  const lock =
    LockService.getScriptLock();

  lock.waitLock(10000);


  try {

    const lastRow =
      Math.max(
        trackerSheet.getLastRow(),
        2
      );


    const ids =
      trackerSheet
        .getRange(
          2,
          1,
          lastRow - 1,
          1
        )
        .getValues()
        .flat();


    let highestNumber = 0;


    ids.forEach(id => {

      const match =
        String(id || '')
          .match(/^VAX-(\d+)$/);


      if (match) {

        const number =
          parseInt(match[1], 10);

        if (number > highestNumber) {
          highestNumber = number;
        }
      }
    });


    return (
      'VAX-' +
      String(highestNumber + 1)
        .padStart(6, '0')
    );


  } finally {

    lock.releaseLock();
  }
}



/**************************************************************
 * FORMAT DATE FOR CONFIRMATION DIALOG
 **************************************************************/
function formatDateForDisplay_(
  dateValue,
  spreadsheet
) {

  if (!dateValue) return '';

  try {

    return Utilities.formatDate(
      new Date(dateValue),
      spreadsheet.getSpreadsheetTimeZone(),
      'MMM d, yyyy'
    );

  } catch (error) {

    return String(dateValue);
  }
}

function ensureDispositionValidations_() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const entrySheet =
    ss.getSheetByName('Vaccination_Entry');

  const dispositionSheet =
    ss.getSheetByName('Disposition_Master');

  if (!entrySheet || !dispositionSheet) return;


  // B14 = Reason Category
  const reasonCategoryRule =
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(
        dispositionSheet.getRange('E2:E'),
        true
      )
      .setAllowInvalid(false)
      .build();


  // B15 = Specific Reason
  const specificReasonRule =
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(
        dispositionSheet.getRange('F2:F'),
        true
      )
      .setAllowInvalid(false)
      .build();


  entrySheet
    .getRange('B14')
    .setDataValidation(reasonCategoryRule);

  entrySheet
    .getRange('B15')
    .setDataValidation(specificReasonRule);
}
