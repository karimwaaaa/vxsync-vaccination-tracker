/**
 * VxSync Lot_Expiry_Master dropdowns for Google Sheets.
 * Source: Vaccine_Schedule_Master, A = Vaccine Type,
 * C = Brand / Example, U = Active (Yes).
 * Destination: Lot_Expiry_Master, D = Vaccine Type,
 * E = Vaccine Brand, H = Availability. Data begins on row 6.
 *
 * FULLY AUTOMATIC — no manual step, ever, for any client.
 *
 * Earlier versions of this file wired the per-edit refresh
 * (vxSyncLotDropdownEdit) through an INSTALLABLE trigger
 * (ScriptApp.newTrigger().onEdit()...). That requires a human to
 * personally authorize this exact spreadsheet's script copy first —
 * fine for one existing client (Latter Day Saints), a real manual step
 * for every future auto-provisioned one.
 *
 * This version calls vxSyncLotDropdownEdit(e) directly from the bare
 * onEdit(e) SIMPLE trigger that already lives in MasterTemplate_Code.gs
 * (see the "LOT INVENTORY DROPDOWNS" block added there). A simple
 * trigger needs no authorization and no installation — it just works,
 * the instant any future client's Sheet is Drive-copied, exactly like
 * the rest of that file's onEdit logic already does. That is what
 * actually removes the manual step, not a menu shortcut to the same
 * manual step.
 *
 * onOpen(e) below is now pure defense-in-depth: a full refresh of every
 * row's dropdowns whenever the Sheet is opened, in case
 * Vaccine_Schedule_Master was ever changed some other way that doesn't
 * fire onEdit (e.g. a bulk paste-as-values from an external import, or
 * a programmatic write via the Sheets API rather than the UI — neither
 * of those fires Apps Script edit events). Needs no authorization
 * either, since it only reads/writes ranges on this spreadsheet.
 *
 * setupVxSyncLotDropdowns() is kept only as a manual "force a full
 * refresh right now" utility — useful for debugging, never required
 * for the automation itself to work.
 */
const VX_LOT = Object.freeze({
  master: 'Vaccine_Schedule_Master',
  lots: 'Lot_Expiry_Master',
  firstRow: 6,
  availability: ['Ready', 'Hold', 'Used Up']
});

function setupVxSyncLotDropdowns() {
  refreshVxSyncLotDropdowns_(SpreadsheetApp.getActiveSpreadsheet());
}

// Called directly from MasterTemplate_Code.gs's shared onEdit(e) —
// see the "LOT INVENTORY DROPDOWNS" block there. Not registered as its
// own trigger anymore; this is now a plain helper function.
function vxSyncLotDropdownEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  const ss = e.source;
  const name = sheet.getName();
  const startCol = e.range.getColumn();
  const endCol = e.range.getLastColumn();

  // Refresh the options when Vaccine Type, Brand, or Active is edited.
  if (name === VX_LOT.master &&
      ((startCol <= 1 && endCol >= 1) ||
       (startCol <= 3 && endCol >= 3) ||
       (startCol <= 21 && endCol >= 21))) {
    refreshVxSyncLotDropdowns_(ss);
    return;
  }

  if (name !== VX_LOT.lots || endCol < 4 || startCol > 4 ||
      e.range.getLastRow() < VX_LOT.firstRow) return;

  const byType = getVxSyncActiveBrands_(ss);
  const first = Math.max(VX_LOT.firstRow, e.range.getRow());
  const last = e.range.getLastRow();
  const chosenTypes = sheet.getRange(first, 4, last - first + 1, 1)
    .getDisplayValues();
  const previousBrands = sheet.getRange(first, 5, last - first + 1, 1)
    .getDisplayValues();
  // Session ID (column A) is stamped only by an actual claim — see
  // claimLotInventoryUnit_ / claimLotInventoryForBoundSheet_ in the
  // web-app and bound-script inventory code. A stamped row represents a
  // real administered (or held) dose, not a blank row logistics is
  // still filling in.
  const sessionRefs = sheet.getRange(first, 1, last - first + 1, 1)
    .getDisplayValues();
  const rules = chosenTypes.map(([type]) => [brandRule_(byType.get(type))]);
  sheet.getRange(first, 5, rules.length, 1).setDataValidations(rules);

  // A previously selected brand cannot remain attached to a different type.
  // Its lot details are also cleared so an old lot is not reused by mistake
  // — but NEVER for a row a real claim has already stamped (Session ID
  // present in column A). Wiping Availability there would silently read
  // as "Ready" again downstream and let an already-administered dose look
  // unused.
  previousBrands.forEach(([brand], i) => {
    if (sessionRefs[i][0]) return; // already claimed by a real record — leave it alone
    const options = byType.get(chosenTypes[i][0]) || [];
    if (brand && !options.includes(brand)) {
      sheet.getRange(first + i, 5, 1, 6).clearContent(); // E:J
    }
  });
}

function refreshVxSyncLotDropdowns_(ss) {
  const sheet = ss.getSheetByName(VX_LOT.lots);
  if (!sheet) throw new Error('Missing worksheet: ' + VX_LOT.lots);
  const byType = getVxSyncActiveBrands_(ss);
  const types = [...byType.keys()].sort((a, b) => a.localeCompare(b));
  if (!types.length) throw new Error('No active vaccine types found in Vaccine_Schedule_Master.');

  const count = sheet.getMaxRows() - VX_LOT.firstRow + 1;
  if (count < 1) throw new Error('Lot_Expiry_Master needs rows below its header.');

  const typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(types, true)
    .setAllowInvalid(false)
    .setHelpText('Choose an active vaccine type from the schedule master.')
    .build();
  sheet.getRange(VX_LOT.firstRow, 4, count, 1).setDataValidation(typeRule);

  const selected = sheet.getRange(VX_LOT.firstRow, 4, count, 1)
    .getDisplayValues();
  const brandRules = selected.map(([type]) => [brandRule_(byType.get(type))]);
  sheet.getRange(VX_LOT.firstRow, 5, count, 1)
    .setDataValidations(brandRules);

  const availabilityRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(VX_LOT.availability, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(VX_LOT.firstRow, 8, count, 1)
    .setDataValidation(availabilityRule);
}

function getVxSyncActiveBrands_(ss) {
  const master = ss.getSheetByName(VX_LOT.master);
  if (!master) throw new Error('Missing worksheet: ' + VX_LOT.master);
  const last = master.getLastRow();
  const byType = new Map();
  if (last < 2) return byType;

  const records = master.getRange(2, 1, last - 1, 21).getDisplayValues();
  records.forEach(row => {
    const type = row[0].trim();
    const brand = row[2].trim();
    const active = row[20].trim().toLowerCase();
    if (!type || !brand || active !== 'yes') return;
    if (!byType.has(type)) byType.set(type, new Set());
    byType.get(type).add(brand);
  });

  return new Map([...byType].map(([type, brands]) => [
    type, [...brands].sort((a, b) => a.localeCompare(b))
  ]));
}

function brandRule_(brands) {
  return brands && brands.length ? SpreadsheetApp.newDataValidation()
    .requireValueInList(brands, true)
    .setAllowInvalid(false)
    .setHelpText('Choose a brand for the selected vaccine type.')
    .build() : null;
}

// ============================================================
//  ONE REMAINING SIMPLE TRIGGER — safe alongside MasterTemplate_Code.gs
//  (that file has onEdit but no onOpen, so this doesn't collide).
//  Pure defense-in-depth refresh on Sheet open; see file header.
// ============================================================
function onOpen(e) {
  try {
    refreshVxSyncLotDropdowns_(SpreadsheetApp.getActiveSpreadsheet());
  } catch (err) {
    console.error('onOpen: refreshVxSyncLotDropdowns_ failed:', err);
  }
  try {
    // See VxSync_Tracker_Backfill.gs — same defense-in-depth reasoning:
    // catches any Recipients_Master/Vaccinators_Master/Schedule/
    // Disposition change made some way that doesn't fire onEdit.
    trkRefreshAllDropdowns_(SpreadsheetApp.getActiveSpreadsheet());
  } catch (err) {
    console.error('onOpen: trkRefreshAllDropdowns_ failed:', err);
  }
  try {
    // See VxSync_Format_Notes.gs — inline header-cell guidance for
    // columns whose required format fails silently (no error) when
    // violated. Idempotent, so calling it on every open is harmless.
    installVxSyncFormatNotes_(SpreadsheetApp.getActiveSpreadsheet());
  } catch (err) {
    console.error('onOpen: installVxSyncFormatNotes_ failed:', err);
  }
  try {
    // See VxSync_Receive_Inventory.gs — this is the ONLY custom menu in
    // this whole bound project, so it lives here rather than risking a
    // second onOpen(e) function elsewhere silently overwriting this one
    // (Apps Script concatenates every .gs file's top-level declarations
    // into one global scope — two functions named onOpen in the same
    // project is a redefinition, not an error, and whichever file's
    // definition loads last silently wins).
    SpreadsheetApp.getUi()
      .createMenu('VxSync')
      .addItem('📷 Receive Inventory…', 'vxReceiveShowDialog')
      .addToUi();
  } catch (err) {
    console.error('onOpen: VxSync menu creation failed:', err);
  }
}
