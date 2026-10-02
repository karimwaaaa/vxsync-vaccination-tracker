/**
 * VxSync inline format guidance — cell notes on header cells.
 *
 * Every note below documents a column where the underlying code depends
 * on a SPECIFIC format that fails SILENTLY when violated — no error
 * message anywhere, the affected feature just quietly stops working
 * (a recipient never shows up in any province, a schedule's Recommended
 * Date never computes, a lot behaves as already-claimed). Each of these
 * was found the hard way, in a real live sheet, this session — see the
 * comment on each note for exactly what broke and why.
 *
 * Idempotent — setNote() simply overwrites, so calling this repeatedly
 * (every onOpen, on every client) is harmless. Runs automatically via
 * the onOpen already in VxSync_Lot_Dropdown.gs, same zero-manual-step
 * principle as the rest of this project. installVxSyncFormatNotes()
 * below is kept only as a manual "apply/refresh notes right now" utility,
 * same role as setupVxSyncLotDropdowns() / setupVxSyncTrackerBackfill().
 */

function installVxSyncFormatNotes() {
  installVxSyncFormatNotes_(SpreadsheetApp.getActiveSpreadsheet());
}

function installVxSyncFormatNotes_(ss) {
  setHeaderNote_(ss, 'Recipients_Master', 12, // L — Assigned Site/Location
    'FORMAT: "PROVINCE - descriptor" — e.g. "QC - MTC White Plains", ' +
    '"Cebu - Mission". Province must come FIRST, opposite order from ' +
    'Lot_Expiry_Master\'s Site of Vaccination column. This field isn\'t ' +
    'shown anywhere in the web app, but it drives which province bucket ' +
    'this recipient shows up in on the encoding form. Get the order ' +
    'backwards and the recipient silently never appears for ANY Site of ' +
    'Vaccination in this province — no error, the dropdown just shows ' +
    '"0 recipients."');

  setHeaderNote_(ss, 'Recipients_Master', 17, // Q — Active
    'Must be exactly "Yes" for this recipient to appear anywhere in the ' +
    'web app. Blank, "Active", "Y", "TRUE", or anything else is treated ' +
    'as inactive and silently excluded — no error.');

  setHeaderNote_(ss, 'Lot_Expiry_Master', 1, // A — Session ID / Order Ref
    'SYSTEM-ONLY — leave blank when adding a new lot row. This gets ' +
    'stamped automatically the moment a real dose is claimed against ' +
    'this lot. A row with a Session ID already in it is treated as ' +
    'claimed/protected and won\'t be auto-cleared on a Type edit. ' +
    'Typing anything here yourself makes this lot behave as if it\'s ' +
    'already been used.');

  setHeaderNote_(ss, 'Lot_Expiry_Master', 3, // C — Site of Vaccination
    'FORMAT: "Site Name - PROVINCE" — e.g. "MTC White Plains - QC". ' +
    'Province must come LAST, after the final hyphen — opposite order ' +
    'from Recipients_Master\'s Assigned Site/Location column. Drives ' +
    'region-based lot matching in the web app\'s Lot picker.');

  setHeaderNote_(ss, 'Vaccine_Schedule_Master', 21, // U — Active
    'Must be exactly "Yes" for this Type/Brand/Schedule combination to ' +
    'appear anywhere — the Lot_Expiry_Master Type/Brand dropdowns, the ' +
    'web app\'s Vaccine Type/Brand cascade, and Vaccination_Tracker\'s ' +
    'own cascade are all filtered on this. Blank or anything else = ' +
    'silently excluded.');

  setHeaderNote_(ss, 'Vaccine_Schedule_Master', 12, // L — Interval Unit
    'Must be spelled EXACTLY (case-sensitive): Days, Weeks, Months, or ' +
    'Years. "days", "Day", "month", etc. are not recognized — ' +
    'Recommended Next Dose Date will silently compute blank instead of ' +
    'showing an error, for both web-app submissions and Tracker ' +
    'backfill rows.');

  setHeaderNote_(ss, 'Vaccinators_Master', 12, // L — Active
    'Must be exactly "Yes" for this vaccinator to appear in the ' +
    'Vaccinator dropdown (web app or Vaccination_Tracker). Blank or ' +
    'anything else = silently excluded, no error.');

  setHeaderNote_(ss, 'Disposition_Master', 1, // A — Dose Disposition
    'Must exactly match one of the six values used everywhere else in ' +
    'this system: Administered, Deferred, Declined, No-show, ' +
    'Contraindicated – Temporary, Contraindicated – Permanent. ' +
    'The two Contraindicated values use an EN DASH (–), not a ' +
    'regular hyphen (-) — they look almost identical but are different ' +
    'characters. A mismatched value here silently breaks the Reason ' +
    'Category / Specific Reason dropdown cascade for that row, with no ' +
    'error.');
}

function setHeaderNote_(ss, sheetName, col, text) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return; // sheet not present in this client's copy — skip quietly
  sheet.getRange(1, col).setNote(text);
}
