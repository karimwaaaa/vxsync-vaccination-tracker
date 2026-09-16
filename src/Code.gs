/**************************************************************
 * VxSync – Backend (Code.gs)
 * [COMPANY_NAME] Worksite Vaccination Program
 *
 * ARCHITECTURE NOTES (read before editing):
 *
 * 1. Vaccination_Entry is NOT a log/table. It is a single-record
 *    "scratchpad" that recalculates one record at a time via
 *    formulas anchored to fixed cells (B2, B11-B29 = inputs;
 *    B3-B9, B28, B31-B34 = formula outputs looked up from
 *    Recipients_Master / Vaccine_Schedule_Master / Vaccinators_Master).
 *
 * 2. Vaccination_Tracker is the real append-only log. Every
 *    submitted record must become ONE NEW ROW there. The
 *    Vaccination_Report dashboard reads ONLY from
 *    Vaccination_Tracker + Recipients_Master. It never looks at
 *    Vaccination_Entry. (The previous version of this file wrote
 *    to Vaccination_Entry and never touched Vaccination_Tracker —
 *    every submission was invisible to the dashboard.)
 *
 * 3. Because we don't want to re-implement the workbook's dose
 *    schedule math (regex dose parsing, interval units, series
 *    completion, etc.) in Apps Script and risk it drifting from
 *    the sheet's own logic, saveVaccinationRecord() uses the
 *    sheet itself as the calculation engine:
 *       write inputs -> Vaccination_Entry -> flush() -> read
 *       back the computed outputs -> append full row to
 *       Vaccination_Tracker -> clear the scratchpad.
 *    This MUST be wrapped in a lock (see LockService below) or
 *    two nurses submitting at the same moment will corrupt each
 *    other's computed values.
 **************************************************************/

// ============================================================
//  CONFIG — the ONLY block that should differ between client deployments.
//  This project follows the "separate Sheet + Apps Script deployment per
//  client" model: this whole bound-script file is copied into a fresh
//  Apps Script project bound to a fresh copy of the workbook for each new
//  client, and this CONFIG block (plus Index.html's logo/branding markup)
//  is the only thing that needs to change to stand up a new client. Every
//  function below reads from CONFIG (or the derived constants right after
//  it, kept for readability at their call sites) — nothing else in the
//  file should ever hard-code a client name, email, or site.
//
//  Checklist for onboarding a new client (see also the naming convention
//  note below):
//   1. Make a copy of the master template Sheet (never edit the template
//      itself) and a copy of this Apps Script project bound to it.
//   2. Fill in every field below with the new client's real values.
//   3. Update Index.html's brand header (client name/logo) to match.
//   4. Deploy > New deployment > Web app, name it "VxSync — [Client Name]"
//      so it's identifiable later.
//   5. Log the new deployment (URL, client name, go-live date, Sheet ID)
//      in the external multi-client tracking spreadsheet.
// ============================================================
const CONFIG = {
  clientName: '[COMPANY_NAME]',

  // --- Branding shown in Index.html's header. All optional — sensible
  // defaults kick in if left blank, so a new client deployment only needs
  // to touch clientName/poweredByText/poweredByLogoBase64 to rebrand. ---
  appName: 'VxSync',
  appTagline: 'Vaccination Tracking System',
  // Text next to the second ("powered by") logo. Defaults to
  // "Powered by " + clientName when left blank.
  poweredByText: '',

  // --- Hard‑coded user lists (update these per client) ---
  // nurseEmails is now a MANUAL OVERRIDE/ADDITION list, not the primary
  // source of Nurse/encoder authorization — per client instruction
  // ("Authorize all the vaccinators in the website already, their gmails
  // are in the sheets"), every ACTIVE row in Vaccinators_Master that has
  // an email address is automatically granted the Nurse role (see
  // getVaccinatorEmailSet_/getUserRole below). Add an email here only for
  // someone who needs encoder access but isn't in Vaccinators_Master for
  // some reason.
  nurseEmails: ['nurse1@company.com', 'nurse2@company.com'],
  adminEmails: ['admin1@company.com', 'admin2@company.com'],
  clientEmails: ['client@company.com'],

  // --- Optional starting site suggestion per nurse. ---
  // NOTE: as of the "type it, remember it" change to Site of Vaccination,
  // this map is NO LONGER enforced/locked server-side — it is only used as
  // a fallback first-run suggestion for an encoder who has never typed a
  // site before (see getEncodingSite() below). Once an encoder types a
  // site once, their own remembered value (PropertiesService) takes over
  // and this map is never consulted again for that account. Safe to leave
  // sparse or even empty per client.
  // DUMMY VALUES FOR TESTING — these three sites are picked from the messy
  // free-text values already sitting in Vaccination_Tracker column Y
  // ("MTC Whiteplains" / "MTC WHITEPLAINS" / "MTC Quezon City" / "MTC QC" /
  // "Site Office A" / "Site Office A (Main)" / "Main Office" — all typed by
  // hand before this interface existed).
  encoderSiteMap: {
    'nurse1@company.com': 'Site Office A',
    'nurse2@company.com': 'Site Office B'
  },

  // --- Hub SSO integration ---
  // hubApiUrl/hubAnonKey are the same values as the Hub's own BACKEND_URL /
  // SUPABASE_ANON_KEY. hubClientId is the id of THIS client's row in the
  // Hub's Supabase `clients` table (Table Editor -> clients -> id column) —
  // the one value you set by hand per client copy, same spirit as the rest
  // of CONFIG.
  hubApiUrl: 'https://YOUR-PROJECT-REF.supabase.co/functions/v1/api',
  hubAnonKey: 'YOUR-ANON-KEY',
  hubClientId: 0,

  // Where to send someone who opens this VxSync URL directly instead of
  // through the Hub (see verifyHubToken_ below) - your Hub's real
  // Netlify URL, e.g. 'https://pq-healthshield-hub.netlify.app'. Leave
  // blank and the "please go through the Hub" page just shows text with
  // no link/button instead of failing.
  hubLoginUrl: 'https://YOUR-HUB-SITE.netlify.app',

  // --- Spreadsheet targeting ---
  // Leave BLANK for a manually-copied, container-bound deployment (the
  // original onboarding model - this script lives inside the Sheet
  // itself, so it always knows which Sheet it's the container of).
  //
  // Set to a real Sheet ID ONLY for a client provisioned by the
  // automated onboarding flow, which deploys this same script as a
  // STANDALONE script (not bound to any Sheet) and fills this in so it
  // knows which Sheet to open. See getSheet_() right below CONFIG.
  sheetId: ''
};

// Derived constants — kept so the rest of the file (and anyone reading it)
// doesn't need to rewrite every reference to CONFIG.xyz. These are the
// ONLY place NURSE_EMAILS/ADMIN_EMAILS/CLIENT_EMAILS/ENCODER_SITE_MAP are
// still defined; do not hard-code new values here per client — edit
// CONFIG above instead.
const NURSE_EMAILS = CONFIG.nurseEmails;
const ADMIN_EMAILS = CONFIG.adminEmails;
const CLIENT_EMAILS = CONFIG.clientEmails;
const ENCODER_SITE_MAP = CONFIG.encoderSiteMap;

// Every SpreadsheetApp.getActiveSpreadsheet() call in this file (there
// were ~29 of them) was replaced with getSheet_() so the exact same
// source works for BOTH deployment models without a single other line
// changing:
//   - Container-bound (CONFIG.sheetId left blank): falls back to
//     getActiveSpreadsheet(), i.e. exactly the old behavior - every
//     already-deployed client keeps working untouched.
//   - Standalone (CONFIG.sheetId set to a real Sheet ID by the
//     provisioning automation): opens that Sheet explicitly, since a
//     standalone script has no bound container to fall back to.
// Cached per execution since openById() is a real remote call and
// several functions call getSheet_() more than once.
var _cachedSheet_ = null;
function getSheet_() {
  if (_cachedSheet_) return _cachedSheet_;
  _cachedSheet_ = CONFIG.sheetId
    ? SpreadsheetApp.openById(CONFIG.sheetId)
    : SpreadsheetApp.getActiveSpreadsheet();
  return _cachedSheet_;
}

// --- doGet: serve ONE shell, with server-side role gating ---
// Content the role isn't allowed to see is never sent to the browser
// (it's stripped by the template engine before the response leaves
// the server) — hiding a tab with CSS is not the same thing and is
// not treated as access control anywhere in this file.
//
// TWO independent gates run here, in order, and BOTH must pass:
//  1. verifyHubToken_ — did this request arrive with a hubToken the
//     Hub itself just issued (i.e. did the person click "Launch" from
//     inside the Hub a moment ago)? This does NOT identify who the
//     person is — it only proves the click came from the Hub, not from
//     someone who bookmarked/forwarded this Apps Script URL directly.
//  2. getUserRole — the existing real Google-account check (admin/
//     nurse/client lists, Hub assignment lookup, sheet fallback). This
//     still runs exactly as before and still decides WHAT the person
//     can see. Gate 1 does not replace this, it only adds a "you have
//     to come from the Hub" requirement in front of it.
function doGet(e) {
  const hubToken = e.parameter.hubToken;
  const hubSession = verifyHubToken_(hubToken);

  if (!hubSession) {
    return htmlWithFavicon_(
      '<div style="font-family:sans-serif; max-width:480px; margin:80px auto; text-align:center; color:#333;">' +
      '<h1 style="color:#1a4d8f;">Please open VxSync from the Hub</h1>' +
      '<p>This link only works when you launch it from inside the [COMPANY_NAME] Hub - ' +
      'log in there and use the "Launch Entry Form" / "Admin Dashboard" button for this ' +
      'client instead of opening this address directly.</p>' +
      (CONFIG.hubLoginUrl && CONFIG.hubLoginUrl.indexOf('YOUR-HUB-SITE') === -1
        ? '<p><a href="' + CONFIG.hubLoginUrl + '" style="display:inline-block; margin-top:12px; padding:10px 20px; background:#1a4d8f; color:#fff; text-decoration:none; border-radius:6px;">Go to the Hub</a></p>'
        : '') +
      '</div>',
      'VxSync – Sign in through the Hub'
    );
  }

  const userEmail = Session.getActiveUser().getEmail();
  console.log('🚨 SERVER RECEIVED EMAIL:', userEmail);
  const role = getUserRole(userEmail);

  if (role === 'None') {
    // Restored from the original doGet(): lets someone signed into the
    // wrong Google account switch accounts right from the denial page,
    // instead of just telling them "no" with no way out.
    const selfUrl = ScriptApp.getService().getUrl();
    const switchUrl = 'https://accounts.google.com/AccountChooser?continue=' + encodeURIComponent(selfUrl);
    const safeEmail = String(userEmail || '(no Google account detected)')
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    return htmlWithFavicon_(
      '<div style="font-family:sans-serif; max-width:480px; margin:60px auto; text-align:center; color:#444;">' +
      '<h1>Access Denied</h1>' +
      '<p>The Google account <b>' + safeEmail + '</b> is not authorised to use this application.</p>' +
      '<p>If a different Google account IS authorised, switch to it below:</p>' +
      '<p><a href="' + switchUrl + '" style="display:inline-block; margin-top:12px; padding:10px 20px; background:#1a4d8f; color:#fff; text-decoration:none; border-radius:6px;">Switch Google Account</a></p>' +
      '<p style="font-size:0.85em; color:#888;">Otherwise, contact your program administrator if you believe this is an error.</p>' +
      '</div>',
      'VxSync – Access Denied'
    );
  }

  // Log this access to the Hub's Activity Log (the "Dashboard Viewed" /
  // entry-form-accessed requirement). This is a best-effort, fire-and-
  // forget ping — see logVxSyncActivity_ below — a failed log call must
  // NEVER block someone from actually using VxSync, so its result is
  // deliberately ignored here.
  //
  // View distinction is best-effort, not exact: doGet serves ONE shell
  // that can contain both the Entry Form and Dashboard tabs (Admins see
  // both, switching client-side via vxShowTab — see Index.html), so there
  // is no server-side way to know which tab an Admin will actually land
  // on/use without a client-side round trip. What CAN be known here is
  // which panel is active on load per role (see the `role === 'Client' ?
  // 'active' : ...` logic in Index.html): Encoders always land on Entry,
  // Clients always land on Dashboard, Admins land on Entry by default.
  // If a truly per-tab log entry is wanted later, EntryForm.html/
  // VxSyncDashboard.html would need their own onload call to a
  // logVxSyncActivity(view) client-facing wrapper instead of doing it
  // here in doGet.
  logVxSyncActivity_(role === 'Client' ? 'dashboard' : 'entry', hubSession.displayName, userEmail);

  const template = HtmlService.createTemplateFromFile('Index');
  template.role = role;
  template.userEmail = userEmail;
  template.encodingSite = ENCODER_SITE_MAP[userEmail] || '';
  template.appName = CONFIG.appName || 'VxSync';
  template.appTagline = CONFIG.appTagline || 'Vaccination Tracking System';
  template.poweredByText = CONFIG.poweredByText || ('Powered by ' + CONFIG.clientName);
  // Also handed to Index.html so it can render the favicon link itself -
  // see FAVICON_BASE64_PNG below and the paste-in snippet delivered
  // alongside this file for the one line Index.html's <head> needs.
  template.faviconBase64 = FAVICON_BASE64_PNG;
  return template.evaluate()
    .setTitle('VxSync – ' + CONFIG.clientName)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

// Confirms a hubToken came from a real, still-valid Hub login - this is
// the SAME session token the Hub already issues at login and already
// uses to gate the Central Ops Dashboard (see DashboardCode.gs's
// verifyHubToken_ / the Hub's own "verifySession" backend action). Reused
// here rather than inventing a second token type, since it's already
// built, already tested, and already expires/invalidates exactly the way
// a Hub session should. Returns null for missing/invalid/expired tokens
// AND when the Hub can't be reached - unlike checkHubAccess_ below, this
// one has NO fallback: no valid token here means no entry, full stop.
// This only proves "a real Hub session opened this," it does not by
// itself decide role/authorization - getUserRole() still runs right
// after this and is what actually decides Admin/Nurse/Client/None.
function verifyHubToken_(token) {
  if (!token) return null;
  try {
    var res = UrlFetchApp.fetch(CONFIG.hubApiUrl, {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'Authorization': 'Bearer ' + CONFIG.hubAnonKey,
        'apikey': CONFIG.hubAnonKey
      },
      payload: JSON.stringify({ action: 'verifySession', token: token }),
      muteHttpExceptions: true
    });
    var data = JSON.parse(res.getContentText());
    if (!data || data.success !== true) return null;
    return { role: data.role, displayName: data.displayName };
  } catch (err) {
    Logger.log('verifyHubToken_ failed: ' + err);
    return null;
  }
}

// The VxSync favicon (cropped from the VxSync cloud+checkmark logo,
// background stripped) as a base64 PNG, so it works from a plain
// HtmlService.createHtmlOutput() string (the "please open through the
// Hub" / "Access Denied" pages) with zero external hosting - Apps Script
// can't serve an uploaded image file as a URL the way Netlify does, but a
// data: URL needs no hosting at all. htmlWithFavicon_ below wraps any raw
// HTML string with this in the <head>; Index.html gets the same string
// via template.faviconBase64 (see the paste-in snippet delivered
// alongside this file).
const FAVICON_BASE64_PNG = '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';

// htmlWithFavicon_ wraps a plain HTML string with a <head><link
// rel="icon" ...></head> so the deny/redirect pages above also show the
// VxSync icon in the browser tab instead of a blank one.
function htmlWithFavicon_(bodyHtml, title) {
  return HtmlService.createHtmlOutput(
    '<head><link rel="icon" type="image/png" href="data:image/png;base64,' + FAVICON_BASE64_PNG + '" /></head>' + bodyHtml
  ).setTitle(title || 'VxSync');
}

// Vaccinators_Master column indexes (0-based). CONFIRMED against the
// latest workbook — this layout shifted from an earlier version of the
// sheet (two new columns, "PTR No." and "TIN NO.", were inserted before
// Email Address), which silently broke every hard-coded index that used
// to point at Email/Active/Display Name. Referencing them by name here
// instead of bare numbers is exactly so a future column insertion doesn't
// quietly break role authorization again without a loud, findable error.
const VACCINATOR_COL = {
  ID: 0, LAST_NAME: 1, FIRST_NAME: 2, MIDDLE_NAME: 3, TITLE: 4,
  LICENSE_TYPE: 5, LICENSE_NUMBER: 6, PTR_NO: 7, TIN_NO: 8,
  EMAIL: 9, MOBILE: 10, ACTIVE: 11, REMARKS: 12, DISPLAY_NAME: 13,
  ADDRESS: 14, CITY: 15, PROVINCE: 16, PROVINCE_REGION: 17
};

function emailListHas_(list, email) {
  const target = String(email || '').trim().toLowerCase();
  if (!target) return false;
  return (list || []).some(function (e) { return String(e).trim().toLowerCase() === target; });
}

// Every ACTIVE vaccinator/encoder with an email address in
// Vaccinators_Master is authorized as a Nurse automatically — per client
// instruction ("Authorize all the vaccinators in the website already,
// their gmails are in the sheets"). This means CONFIG.nurseEmails no
// longer needs the whole roster hand-copied into it; that list still
// works as a manual override/addition (e.g. an encoder who isn't in
// Vaccinators_Master for some reason), but the source of truth for "is
// this person an encoder" is now the sheet itself. Cached briefly
// (CacheService, 5 minutes) since getUserRole() runs on nearly every
// server call in this file — without caching, a busy shift would re-read
// all ~100+ Vaccinators_Master rows on every single RPC.
function getVaccinatorEmailSet_() {
  const cache = CacheService.getScriptCache();
  try {
    const cached = cache.get('vx_vaccinator_emails');
    if (cached) return new Set(JSON.parse(cached));
  } catch (e) { /* cache miss/corrupt — fall through and rebuild */ }

  const emails = [];
  const ss = getSheet_();
  const sheet = ss.getSheetByName('Vaccinators_Master');
  if (sheet) {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (isYes_(row[VACCINATOR_COL.ACTIVE]) && row[VACCINATOR_COL.EMAIL]) {
        emails.push(String(row[VACCINATOR_COL.EMAIL]).trim().toLowerCase());
      }
    }
  }
  try { cache.put('vx_vaccinator_emails', JSON.stringify(emails), 300); } catch (e) { /* sheet too large for cache — non-fatal, just recomputed next call */ }
  return new Set(emails);
}

// Calls the Hub's checkVxSyncAccess endpoint to ask "is this email
// currently assigned to THIS client project in the Hub's Encoder
// Assignments?" Returns null (not false) when the Hub can't be reached at
// all, so the caller can fall back to the sheet-based check instead of
// hard-denying a legitimate nurse over a network hiccup.
function checkHubAccess_(email) {
  try {
    var res = UrlFetchApp.fetch(CONFIG.hubApiUrl, {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'Authorization': 'Bearer ' + CONFIG.hubAnonKey,
        'apikey': CONFIG.hubAnonKey
      },
      payload: JSON.stringify({
        action: 'checkVxSyncAccess',
        payload: { clientId: CONFIG.hubClientId, email: email }
      }),
      muteHttpExceptions: true
    });
    var data = JSON.parse(res.getContentText());
    if (!data || data.success !== true) return null;
    return { authorized: !!data.authorized, role: data.role || null };
  } catch (err) {
    Logger.log('checkHubAccess_ failed: ' + err);
    return null;
  }
}

// ------------------------------------------------------------
// Generic server-to-server call into the Hub's Supabase Edge Function,
// using the SAME trust model as checkHubAccess_/verifyHubToken_ above:
// no Hub session token, authorized purely by the anon key (which Supabase
// itself verifies before the Edge Function code even runs) plus this
// client's own CONFIG.hubClientId identifying which client the call is
// about. Every one of the newer VxSync -> Hub actions below
// (logVxSyncActivity, submitHelpTicketFromVxSync, syncClientVxSyncData,
// reportSyncFailure) is this same shape, so it's pulled out once instead
// of copy-pasted per call site. Returns null on any network/parse
// failure or on { success:false } from the Hub — callers decide whether
// that's fatal or safe to swallow (a failed activity-log ping should
// never block the encoder's actual work, for instance).
function callHubApi_(action, payload) {
  try {
    var res = UrlFetchApp.fetch(CONFIG.hubApiUrl, {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'Authorization': 'Bearer ' + CONFIG.hubAnonKey,
        'apikey': CONFIG.hubAnonKey
      },
      payload: JSON.stringify({ action: action, payload: payload || {} }),
      muteHttpExceptions: true
    });
    var data = JSON.parse(res.getContentText());
    if (!data || data.success !== true) return null;
    return data;
  } catch (err) {
    Logger.log('callHubApi_(' + action + ') failed: ' + err);
    return null;
  }
}

// Pings the Hub's Activity Log so admins can see "so-and-so accessed
// VxSync" without leaving the Hub. Server-to-server (see callHubApi_) -
// VxSync users authenticate via Google, not a Hub session, so there is
// no session token to attach here; the Hub's logVxSyncActivity action
// trusts this call the same way it trusts checkVxSyncAccess (anon key +
// this client's own hubClientId). Deliberately swallows failures - never
// throws, never awaited by the caller for its result - because a broken
// or slow logging call must not be allowed to break or delay someone
// actually opening VxSync.
function logVxSyncActivity_(view, actorName, actorEmail) {
  try {
    callHubApi_('logVxSyncActivity', {
      clientId: CONFIG.hubClientId,
      clientName: CONFIG.clientName,
      actorName: actorName || actorEmail || '(unknown)',
      actorEmail: actorEmail || '',
      view: view || 'entry'
    });
  } catch (err) {
    Logger.log('logVxSyncActivity_ failed: ' + err);
  }
}

// Lets VxSync's "Contact IT Support" help-widget fallback file a ticket
// straight into the Hub's IT Support inbox without the person needing a
// Hub login. Server-to-server, same trust model as above. The Hub tags
// the resulting ticket source: 'vxsync' + source_client_name so the IT
// Support tab can show which system it came from (see
// submitHelpTicketFromVxSync in api_index.ts / getHelpTickets's
// source/sourceClientName fields).
//
// NOT YET WIRED TO A UI — this is the backend half only. The floating
// "Help" button + "Contact IT Support" fallback form in Index.html/
// EntryForm.html/VxSyncDashboard.html is a separate, not-yet-built piece;
// once it exists, its submit handler should call
// google.script.run.submitVxSyncHelpTicket(message, priority).
function submitVxSyncHelpTicket(message, priority) {
  const role = getUserRole(Session.getActiveUser().getEmail());
  if (role === 'None') throw new Error('Not authorised.');
  const email = Session.getActiveUser().getEmail();
  const result = callHubApi_('submitHelpTicketFromVxSync', {
    clientId: CONFIG.hubClientId,
    clientName: CONFIG.clientName,
    submitterName: email, // VxSync has no separate "display name" for the caller beyond their Google account — good enough to identify who filed it
    submitterEmail: email,
    message: message || '',
    priority: priority || 'normal'
  });
  if (!result) {
    throw new Error('Could not reach the Hub to submit this ticket. Please try again in a moment, or contact IT Support directly.');
  }
  return { success: true };
}

// Bug reports are deliberately kept OUT of the Hub's help_tickets table -
// they're a pure email-to-IT-Support path (see reportBugFromVxSync in
// api_index.ts). No ticket is created, no status to track — just an
// immediate email with Reply-To set to the reporter.
function reportBugFromVxSync(message, priority) {
  const role = getUserRole(Session.getActiveUser().getEmail());
  if (role === 'None') throw new Error('Not authorised.');
  const email = Session.getActiveUser().getEmail();
  const result = callHubApi_('reportBugFromVxSync', {
    clientId: CONFIG.hubClientId,
    clientName: CONFIG.clientName,
    submitterName: email,
    submitterEmail: email,
    message: message || '',
    priority: priority || 'normal'
  });
  if (!result) {
    throw new Error('Could not reach the Hub to submit this bug report. Please try again in a moment, or contact IT Support directly.');
  }
  return { success: true };
}

function getUserRole(email) {
  if (!email) return 'None';
  if (emailListHas_(ADMIN_EMAILS, email)) return 'Admin';
  if (emailListHas_(CLIENT_EMAILS, email)) return 'Client';
  if (emailListHas_(NURSE_EMAILS, email)) return 'Encoder'; // manual override/addition list — unchanged

  // Hub is now the source of truth for Encoder authorization, whenever
  // it's reachable — see Encoder Assignments in the Hub's Client Vault
  // tab. NOTE: role naming was standardized Hub-side from "nurse" to
  // "encoder" (see migration_rename_nurse_to_encoder.sql) — this file is
  // catching up to that same vocabulary. The role STRING returned by the
  // Hub was always the lowercase 'admin'/'encoder' pair, so this
  // comparison didn't actually depend on the old 'Nurse' spelling and
  // needed no change beyond the literal below.
  var hub = checkHubAccess_(email);
  if (hub !== null) {
    if (hub.authorized) return hub.role === 'admin' ? 'Admin' : 'Encoder';
    return 'None'; // Hub reached, and said this email isn't assigned here
  }

  // Hub unreachable (network/Supabase hiccup) — fall back to the sheet
  // check rather than locking out a legitimate encoder mid-shift.
  if (getVaccinatorEmailSet_().has(String(email).trim().toLowerCase())) return 'Encoder';
  return 'None';
}

// Used by templates to pull in partial HTML files.
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Every server function that returns report/report-adjacent data
// re-checks the caller's role itself. This is defense in depth:
// even if someone found a way to load report JS in an Encoder session,
// the server would still refuse to hand back the data.
function assertCanViewReport_() {
  const role = getUserRole(Session.getActiveUser().getEmail());
  if (role !== 'Admin' && role !== 'Client') {
    throw new Error('You are not authorised to view the program dashboard.');
  }
  return role;
}
function assertCanEncode_() {
  const role = getUserRole(Session.getActiveUser().getEmail());
  if (role !== 'Admin' && role !== 'Encoder') {
    throw new Error('You are not authorised to submit vaccination records.');
  }
  return role;
}

// ============================================================
//  API FUNCTIONS (called from client‑side JavaScript)
// ============================================================

// Tolerant "is this Yes/Active?" check. Real sheets accumulate junk over
// time — trailing spaces from copy-paste, "yes"/"YES" instead of "Yes",
// a checkbox that stores boolean true instead of text. Comparing with
// === 'Yes' breaks silently (returns an empty list, no error) the moment
// any of that happens, which is exactly what produced the "No active
// recipients found" report even though Recipients_Master has rows in it.
function isYes_(value) {
  if (value === true) return true;
  return String(value || '').trim().toLowerCase() === 'yes';
}

// google.script.run has a known failure mode where a raw Date object
// nested inside an array of objects can silently fail to serialize across
// the RPC bridge — the client receives `null` instead of a catchable
// error, even though the server function completed successfully and
// Logger/manual-Run show the correct return value (manual Run never goes
// through the RPC bridge, so it never surfaces this). Fix: never hand a
// Date object across google.script.run — convert to a plain string first.
function formatDateForClient_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return Utilities.formatDate(value, Session.getScriptTimeZone() || 'Asia/Manila', 'yyyy-MM-dd');
  }
  return String(value);
}

// --- DIAGNOSTIC: run this manually from the Apps Script editor (Run menu)
// when a master-data lookup comes back empty. It never runs from the web
// app itself. Check View > Logs / Executions for the output — it will show
// you the exact header row and first 3 data rows exactly as Apps Script
// reads them, so you can see whether "Active" really contains the text
// "Yes", a boolean, or something with hidden whitespace.
function debugDumpRecipients_() {
  const sheet = getSheet_().getSheetByName('Recipients_Master');
  const data = sheet.getDataRange().getValues();
  Logger.log('Header (row 1): %s', JSON.stringify(data[0]));
  for (let i = 1; i < Math.min(4, data.length); i++) {
    Logger.log('Row %s: %s', i + 1, JSON.stringify(data[i]));
    Logger.log('  -> column Q (Active) value = %s (type %s)', data[i][16], typeof data[i][16]);
  }
}

// --- Recipients ---
function getRecipients() {
  const callerRole = assertCanEncode_();
  Logger.log('getRecipients called by %s (resolved role: %s), spreadsheet: %s',
    Session.getActiveUser().getEmail() || '(blank — see note below)',
    callerRole,
    getSheet_().getUrl());
  try {
    const ss = getSheet_();
    const sheet = ss.getSheetByName('Recipients_Master');
    if (!sheet) throw new Error('Recipients_Master sheet not found');
    const data = sheet.getDataRange().getValues();
    Logger.log('Recipients_Master has %s data rows (excluding header)', data.length - 1);
    const recipients = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!isYes_(row[16])) continue; // Active = 'Yes'
      if (!row[0]) continue;
      // "province" is derived from Assigned Site/Location (Recipients_Master
      // column L) purely so the entry form can default-filter this
      // recipient into the right province bucket — CHANGED per client
      // instruction, Assigned Site/Location itself is no longer shown
      // anywhere in the web app; Department/Unit is the field encoders
      // now see. See extractProvinceFromRecipientSite_ for the "text
      // before the first hyphen is the province" parsing rule.
      recipients.push({
        id: row[0],
        name: row[2] + ', ' + row[3] + (row[4] ? ' ' + row[4] : ''),
        fullName: row[18] || row[2] + ', ' + row[3],
        email: row[12],
        dob: formatDateForClient_(row[5]),
        category: row[7],
        company: row[8],
        dept: row[10],
        province: extractProvinceFromRecipientSite_(row[11]) || '',
        mobile: row[13]
      });
    }
    Logger.log('getRecipients returning %s active recipients', recipients.length);
    return recipients;
  } catch (e) {
    Logger.log('getRecipients threw: %s', e.message);
    console.error('getRecipients error:', e);
    throw e;
  }
}

function getRecipientDetails(recipientId) {
  assertCanEncode_();
  try {
    const ss = getSheet_();
    const sheet = ss.getSheetByName('Recipients_Master');
    if (!sheet) throw new Error('Recipients_Master sheet not found');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0] === recipientId) {
        return {
          id: row[0],
          // Exact "Last, First Middle" string as stored in Recipients_Master.
          // This is what gets written into Vaccination_Entry!B2 — we use the
          // authoritative string tied to the ID the nurse actually picked,
          // rather than re-typing it, to avoid typos breaking the lookup.
          nameForLookup: row[2] + ', ' + row[3] + (row[4] ? ' ' + row[4] : ''),
          name: row[2] + ', ' + row[3] + (row[4] ? ' ' + row[4] : ''),
          email: row[12],
          dob: formatDateForClient_(row[5]),
          category: row[7],
          company: row[8],
          dept: row[10],
          province: extractProvinceFromRecipientSite_(row[11]) || '',
          // Kept ONLY so saveVaccinationRecord can still write Vaccination_
          // Tracker column J (Assigned Site/Location) for continuity with
          // existing records — CHANGED per client instruction, this value
          // is never displayed or editable in the web app anymore.
          rawAssignedSite: row[11] || '',
          mobile: row[13],
          sex: row[6]
        };
      }
    }
    return null;
  } catch (e) {
    console.error('getRecipientDetails error:', e);
    throw e;
  }
}

// --- Vaccine Schedule ---
function getVaccineTypes() {
  assertCanEncode_();
  try {
    const ss = getSheet_();
    const sheet = ss.getSheetByName('Vaccine_Schedule_Master');
    if (!sheet) throw new Error('Vaccine_Schedule_Master sheet not found');
    const data = sheet.getDataRange().getValues();
    const types = new Set();
    for (let i = 1; i < data.length; i++) {
      const type = data[i][0];
      if (type && isYes_(data[i][20])) types.add(type);
    }
    return Array.from(types).sort();
  } catch (e) {
    console.error('getVaccineTypes error:', e);
    throw e;
  }
}

function getBrands(vaccineType) {
  assertCanEncode_();
  try {
    const ss = getSheet_();
    const sheet = ss.getSheetByName('Vaccine_Schedule_Master');
    if (!sheet) throw new Error('Vaccine_Schedule_Master sheet not found');
    const data = sheet.getDataRange().getValues();
    const brands = new Set();
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0] === vaccineType && isYes_(row[20]) && row[2]) {
        brands.add(row[2]);
      }
    }
    return Array.from(brands).sort();
  } catch (e) {
    console.error('getBrands error:', e);
    throw e;
  }
}

function getSchedules(vaccineType, brand) {
  assertCanEncode_();
  try {
    const ss = getSheet_();
    const sheet = ss.getSheetByName('Vaccine_Schedule_Master');
    if (!sheet) throw new Error('Vaccine_Schedule_Master sheet not found');
    const data = sheet.getDataRange().getValues();
    const schedules = new Set();
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0] === vaccineType && row[2] === brand && isYes_(row[20]) && row[5]) {
        schedules.add(row[5]);
      }
    }
    return Array.from(schedules).sort();
  } catch (e) {
    console.error('getSchedules error:', e);
    throw e;
  }
}

// CONFIRMED against the real workbook's own data validation on
// Vaccination_Entry!B12: the Intended Dose Number field is NOT a
// per-schedule computed list — it is this exact fixed set of 6 values,
// the same for every vaccine type/brand/schedule combination:
//   Dose 1, Dose 2, Dose 3, Booster, Annual Dose, Not Applicable
// (An earlier version of this function derived a per-schedule dose list
// from Vaccine_Schedule_Master's "Current Dose" column instead — that
// doesn't match how the sheet's own dropdown actually works, so any
// vaccine needing a 4th/5th dose, e.g. DTaP or Twinrix Rapid, simply
// falls under "Booster" here rather than a "Dose 4"/"Dose 5" that the
// sheet has no way to select in the first place.)
const INTENDED_DOSE_OPTIONS = ['Dose 1', 'Dose 2', 'Dose 3', 'Booster', 'Annual Dose', 'Not Applicable'];

function getDoseNumbers(vaccineType, brand, scheduleDisplay) {
  assertCanEncode_();
  return INTENDED_DOSE_OPTIONS.slice();
}

// --- Disposition ---
function getDispositions() {
  assertCanEncode_();
  try {
    const ss = getSheet_();
    const sheet = ss.getSheetByName('Disposition_Master');
    if (!sheet) throw new Error('Disposition_Master sheet not found');
    const data = sheet.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0]) {
        result.push({ disposition: row[0], category: row[1], reason: row[2] });
      }
    }
    return result;
  } catch (e) {
    console.error('getDispositions error:', e);
    throw e;
  }
}

// --- Site of Vaccination (where THIS dose was actually given) ---
// CHANGED per client instruction: this is now a free-typed field, not a
// server-locked dropdown. ENCODER_SITE_MAP is no longer used to enforce a
// site — it stays in the file only as an optional seed value the first
// time a given encoder ever opens the form (see fallback below), so an
// existing deployment doesn't regress to a blank field on day one.
// PropertiesService.getUserProperties() is scoped automatically to
// whichever Google account is calling this (no email needs to be passed
// in) and persists across sessions/devices for that account, so "the site
// this encoder typed last time" survives a closed browser tab, a
// different day, even a different computer signed into the same Google
// account — exactly the "so they won't have to type it again" behaviour
// asked for.
function getEncodingSite() {
  const role = getUserRole(Session.getActiveUser().getEmail());
  if (role !== 'Encoder' && role !== 'Admin') throw new Error('Not authorised.');
  const remembered = PropertiesService.getUserProperties().getProperty('vx_last_site');
  if (remembered) return remembered;
  return ENCODER_SITE_MAP[Session.getActiveUser().getEmail()] || '';
}

// Called after a successful save (see saveVaccinationRecord) so the next
// time this same encoder opens the form, their site is already filled in.
function rememberEncodingSite_(site) {
  if (!site) return;
  PropertiesService.getUserProperties().setProperty('vx_last_site', site);
}

// Distinct site names pulled from Recipients_Master (authoritative
// vocabulary already in use) plus anything in ENCODER_SITE_MAP, for the
// Admin's editable site dropdown.
// NOTE: this is the ENCODING site list (where a dose was actually given —
// Vaccination_Tracker column Y, "Vaccination Location"), NOT the
// recipient's home "Assigned Site / Location" (Recipients_Master column L)
// — those are two different concepts in the real data. An earlier version
// of this function pulled from Recipients_Master column L by mistake,
// which would have offered Admins a list of recipients' home
// affiliations (e.g. "Lingayen - Mission") instead of actual encoding
// sites (e.g. "MTC Whiteplains"). Confirmed against the real workbook:
// Vaccination Location values are currently messy/inconsistent free text
// ("MTC Whiteplains" vs "MTC WHITEPLAINS" vs "MTC QC" vs "Mission Trainin
// Center, Quezon City") because nurses were typing them by hand — this is
// exactly the problem ENCODER_SITE_MAP (locking a nurse's site server-side)
// is meant to fix going forward. Existing messy values are still included
// here so Admins can see/select historical sites, but new nurse accounts
// should always get a clean value in ENCODER_SITE_MAP rather than relying
// on free text.
function getKnownSites() {
  assertCanEncode_();
  const ss = getSheet_();
  const trackerSheet = ss.getSheetByName('Vaccination_Tracker');
  const sites = new Set();
  if (trackerSheet) {
    const data = trackerSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][TRACKER_COL.VACCINATION_LOCATION]) sites.add(data[i][TRACKER_COL.VACCINATION_LOCATION]);
    }
  }
  Object.keys(ENCODER_SITE_MAP).forEach(function (email) {
    if (ENCODER_SITE_MAP[email] && ENCODER_SITE_MAP[email].indexOf('REPLACE_WITH') !== 0) {
      sites.add(ENCODER_SITE_MAP[email]);
    }
  });
  return Array.from(sites).sort();
}

// ============================================================
//  DOSE HISTORY / SERIES CONTINUATION
//  Vaccination_Tracker is the real log — this reads that directly to find
//  a recipient's most recent ADMINISTERED dose of a given vaccine type,
//  so the entry form can auto-continue their series instead of asking the
//  nurse to re-derive "what dose comes next" by hand. Only 'Administered'
//  rows count: a Deferred/Declined/No-show attempt did not actually
//  advance the series.
// ============================================================

// Vaccination_Tracker column indices (0-based), for readability below.
const TRACKER_COL = {
  RECORD_ID: 0, RECIPIENT_ID: 1, RECIPIENT_NAME: 2, DEPARTMENT: 8, ASSIGNED_SITE: 9,
  VACCINE_TYPE: 10, INTENDED_DOSE: 11, DISPOSITION: 12,
  REASON_CATEGORY: 13, SPECIFIC_REASON: 14, STATUS_DATE: 15, REVIEW_DATE: 16,
  VACCINE_BRAND: 17, DOSE_NUMBER: 18, VACCINATION_DATE: 19,
  LOT_NUMBER: 20, EXPIRY_DATE: 21,
  VACCINATION_LOCATION: 24,
  SCHEDULE_CODE: 27, NEXT_DOSE: 28, RECOMMENDED_DATE: 29, SCHEDULED_APPOINTMENT: 30,
  SCHEDULE_STATUS: 32, REMINDER_STATUS: 33, SERIES_STATUS: 35
};

// Vaccine_Schedule_Master column indexes (0-based), confirmed against the
// real workbook (CorpShield_VxSync). Referenced by name below instead of
// bare numbers so the dose-history / starting-dose logic reads the same
// way the sheet's own Vaccination_Entry!B31 "Next Dose" formula does.
const SCHEDULE_COL = {
  VACCINE_TYPE: 0, BRAND: 2, SCHEDULE_CODE: 4, SCHEDULE_DISPLAY: 5,
  SCHEDULE_TYPE: 7, CURRENT_DOSE: 8, NEXT_ACTION: 9,
  INTERVAL_VALUE: 10, INTERVAL_UNIT: 11,
  TOTAL_SERIES_DOSES: 15, SCHEDULE_NOTES: 19, ACTIVE: 20
};

// ============================================================
//  ROUTE-OF-ADMINISTRATION HINT FROM SCHEDULE NOTES (col T) — per client
//  instruction, prefill Route of Administration from "the vaccine type's
//  rules in Vaccine_Schedule_Master." CONFIRMED against real column T
//  values: it is free-form prose written for humans (clinical warnings,
//  scheduling reminders, dose-sequencing rules, configuration
//  placeholders — see real examples like "Do not auto-schedule
//  generically." or "Dose 1 must be Tdap. Dose 2 may be Td or Tdap."),
//  NOT a structured field. Only a small minority of rows mention a route
//  at all (e.g. "...administered subcutaneously.", "...administered SC."),
//  and essentially none mention an Administration Site — so this can only
//  ever be a best-effort ROUTE hint, never a Site hint. Administration
//  Site continues to be prefilled ONLY from the recipient's own last-dose
//  history (see efPrefillSelectValue('adminSite', hist.lastAdminSite) in
//  EntryForm.html) — there's nothing in this sheet to source it from.
//
//  DESIGN CHOICE: this deliberately only ever returns a route when the
//  match is UNAMBIGUOUS — a single route mentioned, consistently, across
//  every row for that vaccine type. If notes mention more than one route
//  (or disagree row-to-row) or mention none, this returns null and the
//  field is simply left for the encoder to fill in themselves, same as
//  today. Given the route printed on the wrong dose is a genuine patient-
//  safety concern (see checkRouteWarning above), a confident "no answer"
//  is much better than a guess dressed up as a fact.
// ============================================================
const ROUTE_NOTE_PATTERNS = [
  { re: /\bsub-?cutaneous(ly)?\b/i, route: 'SC' },
  { re: /\bintramuscular(ly)?\b/i, route: 'IM' },
  { re: /\bintradermal(ly)?\b/i, route: 'ID' },
  { re: /\borally?\b/i, route: 'Oral' },
  { re: /\bsc\b/i, route: 'SC' },
  { re: /\bim\b/i, route: 'IM' },
  { re: /\bid\b/i, route: 'ID' },
  { re: /\bpo\b/i, route: 'Oral' } // "PO" — per os, common shorthand for oral
];

// Scans one Schedule Notes string for a route mention. Returns the
// canonical route code (matching the Route <select>'s own option values:
// 'IM' | 'SC' | 'ID' | 'Oral') only if EXACTLY ONE distinct route is
// mentioned in the text — if the text mentions none, or contradicts
// itself by mentioning more than one, this returns null rather than
// picking one arbitrarily.
function extractRouteHintFromNotes_(notesText) {
  if (!notesText) return null;
  const text = String(notesText);
  const matched = {};
  ROUTE_NOTE_PATTERNS.forEach(function (p) {
    if (p.re.test(text)) matched[p.route] = true;
  });
  const routes = Object.keys(matched);
  return routes.length === 1 ? routes[0] : null;
}

// Checks EVERY Vaccine_Schedule_Master row for the given vaccine type
// (there's normally one row per dose/brand/schedule combination) and only
// returns a route if every row that mentions one at all agrees on the
// SAME route — one row saying "IM" and another saying "SC" for the same
// vaccine type is a reason to stay silent, not to guess which one wins.
function getRouteHintForVaccineType_(vaccineType) {
  const sheet = getSheet_().getSheetByName('Vaccine_Schedule_Master');
  if (!sheet || !vaccineType) return null;
  const data = sheet.getDataRange().getValues();
  const found = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[SCHEDULE_COL.VACCINE_TYPE] !== vaccineType) continue;
    const hint = extractRouteHintFromNotes_(row[SCHEDULE_COL.SCHEDULE_NOTES]);
    if (hint) found[hint] = true;
  }
  const routes = Object.keys(found);
  return routes.length === 1 ? routes[0] : null;
}

// Client-facing: best-effort Route of Administration hint for a vaccine
// type, derived from Vaccine_Schedule_Master's Schedule Notes (col T).
// Non-blocking/non-critical by design — same posture as checkSiteTypo.
function getRouteHintForVaccine(vaccineType) {
  assertCanEncode_();
  try {
    return { route: getRouteHintForVaccineType_(vaccineType) };
  } catch (e) {
    console.error('getRouteHintForVaccine error:', e);
    return { route: null };
  }
}

// ============================================================
//  LOCATION INTELLIGENCE
//  Everything in this block exists to answer one question reliably:
//  "what PROVINCE is this piece of free text talking about?" — because
//  every filtering feature the client asked for (recipient lookup scoped
//  to an encoder's area, vaccinator lookup scoped to an encoder's area,
//  the dashboard's site-of-vaccination filter) is built on top of that
//  one answer.
//
//  PROVINCE_ALIASES was seeded from the REAL Vaccinators_Master data in
//  the workbook the client sent — including the typos actually sitting in
//  it today (confirmed programmatically): "Pangsinan" vs "Pangasinan",
//  "Daval del Sur" / "Davao Del Sur" / "Davao del Sur" (three spellings of
//  one province), "Agusan del NOrte", "Agneles City" (Angeles City),
//  "Ilo-Ilo City" vs "Iloilo City". Cebu is included even though no
//  vaccinator is currently based there, because Recipients_Master already
//  has recipients assigned to "Cebu - Mission".
//
//  DESIGN CHOICE: normalizeProvince_() never hard-fails on an unrecognized
//  input. If nothing matches, it title-cases the raw text and returns
//  THAT as a new province bucket. A brand-new province nobody has typed
//  before doesn't crash the filter or get dumped into an "Unknown" catch-
//  all — it just becomes its own bucket immediately, and Tier 2 below
//  starts learning it from real usage. This is deliberately the opposite
//  of a hard-coded exhaustive list of all ~82 PH provinces: an exhaustive
//  list would still miss a client's typo of a province THAT IS on the
//  list (that's the whole problem — "Pangsinan" is not going to self-
//  correct just because "Pangasinan" is in some master list), so the real
//  fix is the alias table plus the learning tier below, not more rows.
// ============================================================

const PROVINCE_ALIASES = {
  "metro manila": "Metro Manila",
  "mm": "Metro Manila",
  "ncr": "Metro Manila",
  "qc": "Metro Manila",
  "q.c.": "Metro Manila",
  "quezon city": "Metro Manila",
  "manila": "Metro Manila",
  "makati": "Metro Manila",
  "makati city": "Metro Manila",
  "pasig": "Metro Manila",
  "pasig city": "Metro Manila",
  "marikina": "Metro Manila",
  "marikina city": "Metro Manila",
  "taguig": "Metro Manila",
  "bgc": "Metro Manila",
  "paranaque": "Metro Manila",
  "paranaque city": "Metro Manila",
  "pasay": "Metro Manila",
  "pasay city": "Metro Manila",
  "alabang": "Metro Manila",
  "muntinlupa": "Metro Manila",
  "mandaluyong": "Metro Manila",
  "san juan": "Metro Manila",
  "caloocan": "Metro Manila",
  "malabon": "Metro Manila",
  "navotas": "Metro Manila",
  "valenzuela": "Metro Manila",
  "las pinas": "Metro Manila",
  "laspinas": "Metro Manila",
  "camanava": "Metro Manila",
  "qc north": "Metro Manila",
  "davao del sur": "Davao del Sur",
  "daval del sur": "Davao del Sur",
  "davao": "Davao del Sur",
  "davao city": "Davao del Sur",
  "agusan del norte": "Agusan del Norte",
  "agusan del norte ": "Agusan del Norte",
  "butuan city": "Agusan del Norte",
  "butuan": "Agusan del Norte",
  "pangasinan": "Pangasinan",
  "pangsinan": "Pangasinan",
  "lingayen": "Pangasinan",
  "urdaneta": "Pangasinan",
  "pampanga": "Pampanga",
  "angeles": "Pampanga",
  "angeles city": "Pampanga",
  "agneles city": "Pampanga",
  "agneles": "Pampanga",
  "cavite": "Cavite",
  "bacoor": "Cavite",
  "bacoor city": "Cavite",
  "gma": "Cavite",
  "general mariano alvarez": "Cavite",
  "dasmarinas": "Cavite",
  "imus": "Cavite",
  "laguna": "Laguna",
  "san pablo": "Laguna",
  "san pablo city": "Laguna",
  "san pedro": "Laguna",
  "calamba": "Laguna",
  "sta rosa": "Laguna",
  "santa rosa": "Laguna",
  "batangas": "Batangas",
  "lipa": "Batangas",
  "lipa city": "Batangas",
  "tanauan": "Batangas",
  "tanauan batangas": "Batangas",
  "darasa": "Batangas",
  "nueva ecija": "Nueva Ecija",
  "cabanatuan": "Nueva Ecija",
  "cabanatuan city": "Nueva Ecija",
  "palawan": "Palawan",
  "puerto princesa": "Palawan",
  "puerto princesa city": "Palawan",
  "camarines sur": "Camarines Sur",
  "naga": "Camarines Sur",
  "naga city": "Camarines Sur",
  "camarines norte": "Camarines Norte",
  "leyte": "Leyte",
  "palo": "Leyte",
  "tacloban": "Leyte",
  "tacloban city": "Leyte",
  "albay": "Albay",
  "legazpi": "Albay",
  "legazpi city": "Albay",
  "legaspi": "Albay",
  "western visayas": "Iloilo",
  "iloilo": "Iloilo",
  "ilo-ilo": "Iloilo",
  "ilo ilo": "Iloilo",
  "ilo-ilo city": "Iloilo",
  "iloilo city": "Iloilo",
  "misamis oriental": "Misamis Oriental",
  "cagayan de oro": "Misamis Oriental",
  "cagayan de oro city": "Misamis Oriental",
  "cdo": "Misamis Oriental",
  "cagayan": "Cagayan",
  "tuguegarao": "Cagayan",
  "tuguegarao city": "Cagayan",
  "isabela": "Isabela",
  "cauayan": "Isabela",
  "cauayan city": "Isabela",
  "zamboanga": "Zamboanga",
  "zamboanga city": "Zamboanga",
  "south cotabato": "South Cotabato",
  "general santos": "South Cotabato",
  "gensan": "South Cotabato",
  "bulacan": "Bulacan",
  "san rafael": "Bulacan",
  "malolos": "Bulacan",
  "oriental mindoro": "Oriental Mindoro",
  "calapan": "Oriental Mindoro",
  "calapan city": "Oriental Mindoro",
  "zambales": "Zambales",
  "olongapo": "Zambales",
  "olongapo city": "Zambales",
  "cebu": "Cebu",
  "cebu city": "Cebu",
  "mandaue": "Cebu",
  "lapu-lapu": "Cebu"
};

// Common, high-confidence typo fixes applied SILENTLY (with a small
// notice so the encoder isn't confused by their own text changing under
// them) — Tier 1 of the three-tier system. Keys are lowercased/trimmed.
// This is intentionally a SHORT list of near-certain fixes, not a dumping
// ground — anything less than "obviously the same place" belongs in Tier
// 2 (learned) or Tier 3 (suggested), not here.
const TIER1_FIXED_CORRECTIONS = {
  'qc': 'QC', 'q.c.': 'QC', 'quezon city': 'Quezon City', 'mm': 'Metro Manila', 'ncr': 'Metro Manila',
  'mtc whitepalins': 'MTC Whiteplains', 'mtc whiteplans': 'MTC Whiteplains',
  'mtc whiteplains': 'MTC Whiteplains', 'mtc wp': 'MTC Whiteplains',
  'mtc qc': 'MTC Quezon City', 'mtc quezon city': 'MTC Quezon City',
  'pangsinan': 'Pangasinan', 'davao del sur': 'Davao del Sur',
  'daval del sur': 'Davao del Sur', 'agusan del norte': 'Agusan del Norte',
  'ilo-ilo': 'Iloilo', 'ilo ilo': 'Iloilo', 'iloilo': 'Iloilo'
};

// ============================================================
//  DATE-ONLY PARSING/FORMATTING — fixes a real bug: Date of Birth,
//  Status Date, Vaccination Date, Review Date and Expiry Date are all
//  meant to be DATES ONLY, but every one of them was written to the
//  sheet as `new Date(dateOnlyString)`. Parsing a bare "YYYY-MM-DD"
//  string that way is interpreted as UTC MIDNIGHT — which, once Google
//  Sheets displays it in the spreadsheet's own timezone, can land on a
//  time other than local midnight (e.g. 08:00 for a UTC+8 spreadsheet),
//  so the cell shows a time component even though nobody entered one.
//  parseDateOnlyLocal_ builds the Date directly from the LOCAL
//  year/month/day components instead, so it's always exactly midnight in
//  the spreadsheet's own timezone — never anything to display.
// ============================================================
function todayDateOnly_() {
  const n = new Date();
  return new Date(n.getFullYear(), n.getMonth(), n.getDate());
}

function parseDateOnlyLocal_(dateStr) {
  if (!dateStr) return null;
  const s = String(dateStr).trim();
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/); // matches a plain date or the date part of an old datetime-local value
  if (!m) {
    const fallback = new Date(s);
    return isNaN(fallback.getTime()) ? null : fallback;
  }
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}

// Belt-and-suspenders: explicitly force a cell/range to a date-only number
// format, so a time component can never display even if some other code
// path (or someone typing directly into the sheet) puts a non-midnight
// value there. Never throws — formatting is cosmetic, not worth failing a
// save or a one-time setup step over.
function stampDateOnlyFormat_(range) {
  try { range.setNumberFormat('yyyy-mm-dd'); } catch (e) { /* non-critical */ }
}

function normalizeLookupKey_(text) {
  return String(text || '').trim().toLowerCase().replace(/\s+/g, ' ');
}

// Resolves free text to a canonical province name. Handles two input
// shapes: a bare place name ("QC", "Angeles"), or the two site-string
// conventions used elsewhere in this file — a "PROVINCE - description"
// prefix (Recipients_Master's existing Assigned Site/Location values,
// e.g. "Cebu - Mission") or a "Site Name - PROVINCE" suffix (the NEW
// Site of Vaccination convention the client is introducing, e.g.
// "MTC White Plains - QC"). Callers that already know which shape they
// have should use extractProvinceFromRecipientSite_ /
// extractProvinceFromVaccinationSite_ below instead — this function is
// the shared final lookup step both of those call into.
function normalizeProvince_(rawText) {
  const key = normalizeLookupKey_(rawText);
  if (!key) return null;
  if (PROVINCE_ALIASES[key]) return PROVINCE_ALIASES[key];
  // Not recognized — self-bucket under a title-cased version of whatever
  // was typed, rather than failing. Confirmed deliberate design choice,
  // see the comment above PROVINCE_ALIASES.
  return String(rawText).trim().replace(/\w\S*/g, function (w) {
    return w.charAt(0).toUpperCase() + w.slice(1).toLowerCase();
  });
}

// Recipients_Master's Assigned Site/Location convention (pre-existing,
// confirmed against real data): province comes FIRST, e.g.
// "Cebu - Mission", "QC North - Mission", "Angeles - Mission".
function extractProvinceFromRecipientSite_(assignedSiteText) {
  if (!assignedSiteText) return null;
  const parts = String(assignedSiteText).split('-');
  const provincePart = parts[0].trim();
  if (!provincePart) return null;
  return normalizeProvince_(provincePart);
}

// Site of Vaccination's NEW convention (per client instruction): province
// comes LAST, after the final " - ", e.g. "MTC White Plains - QC". If
// there's no " - " at all, we can't reliably detect a province from this
// text — return null rather than guessing off the whole string.
function extractProvinceFromVaccinationSite_(siteText) {
  if (!siteText) return null;
  const idx = String(siteText).lastIndexOf('-');
  if (idx === -1) return null;
  const provincePart = String(siteText).substring(idx + 1).trim();
  if (!provincePart) return null;
  return normalizeProvince_(provincePart);
}

// Mirror of extractProvinceFromVaccinationSite_ above, but returns the
// SITE NAME half of the "Site Name - PROVINCE" convention (the text
// BEFORE the last "-") rather than the province. Per client instruction:
// the dashboard's Site-of-Vaccination filter's second dropdown should
// show/match on just the site name (e.g. "MTC Whiteplains"), not the full
// raw string ("MTC Whiteplains - QC") — since the same site name can
// legitimately appear against more than one raw string over time (typos,
// spacing), while still meaning the same physical site.
function extractSiteNameFromVaccinationSite_(siteText) {
  if (!siteText) return null;
  const idx = String(siteText).lastIndexOf('-');
  if (idx === -1) return String(siteText).trim() || null;
  const namePart = String(siteText).substring(0, idx).trim();
  return namePart || null;
}

// ---- Levenshtein distance (no built-in in Apps Script) ----
// Used by Tier 2/Tier 3 below to find "the known value this typed text is
// probably trying to say" — e.g. "MTC WP" vs "MTC Whiteplains" have very
// different lengths but the abbreviation is a real-world shorthand, so
// plain edit distance alone under-detects it; see acronymMatches_ below
// for that case. For same-ish-length typos ("Whitepalins" vs
// "Whiteplains"), edit distance is exactly the right tool.
function levenshtein_(a, b) {
  a = String(a || ''); b = String(b || '');
  const m = a.length, n = b.length;
  if (m === 0) return n;
  if (n === 0) return m;
  let prev = new Array(n + 1);
  let curr = new Array(n + 1);
  for (let j = 0; j <= n; j++) prev[j] = j;
  for (let i = 1; i <= m; i++) {
    curr[0] = i;
    for (let j = 1; j <= n; j++) {
      const cost = (a.charAt(i - 1) === b.charAt(j - 1)) ? 0 : 1;
      curr[j] = Math.min(curr[j - 1] + 1, prev[j] + 1, prev[j - 1] + cost);
    }
    const tmp = prev; prev = curr; curr = tmp;
  }
  return prev[n];
}

function similarityRatio_(a, b) {
  a = String(a || ''); b = String(b || '');
  const maxLen = Math.max(a.length, b.length);
  if (maxLen === 0) return 1;
  return 1 - (levenshtein_(a.toLowerCase(), b.toLowerCase()) / maxLen);
}

// Catches abbreviation-style shorthand ("MTC WP" -> "MTC Whiteplains")
// that edit distance alone misses because the strings are very different
// lengths: true when every word in the short text is a prefix of the
// corresponding word in the long text, in order (or the whole short text
// is literally the initials of the long text's words).
function acronymMatches_(shortText, longText) {
  const shortWords = String(shortText).trim().toLowerCase().split(/\s+/);
  const longWords = String(longText).trim().toLowerCase().split(/\s+/);
  if (shortWords.length && shortWords.length === longWords.length) {
    let allPrefix = true;
    for (let i = 0; i < shortWords.length; i++) {
      if (longWords[i].indexOf(shortWords[i]) !== 0) { allPrefix = false; break; }
    }
    if (allPrefix) return true;
  }
  // Initials check: "MTC WP" -> "M","T","C","W","P" vs first letters of
  // every word in the long text run together.
  const initials = longWords.map(function (w) { return w.charAt(0); }).join('');
  const shortJoined = shortWords.join('');
  if (shortJoined.length >= 2 && initials.indexOf(shortJoined) !== -1) return true;
  return false;
}

// ---- Three-tier autocorrect — WORD BY WORD ----
// Applied to free-typed location fields (Site of Vaccination today; built
// generically enough to reuse elsewhere).
//
// CHANGED (real bug fix): this used to compare the ENTIRE typed phrase
// against entire known phrases with one whole-string edit-distance
// similarity score. That's what produced a confirmed bad suggestion in
// practice: typing "MTC Whiteplains - Quezzzon Cityy" got "Did you mean
// 'Site Office A - Quezon City'?" — a completely unrelated site — because
// whole-string similarity only cares about OVERALL character overlap and
// length, not which specific part is actually similar to what. Two long
// strings can share enough characters/structure (" - ", "City", similar
// length) to clear a 50%+ similarity bar despite naming different places
// entirely.
//
// The fix: check EACH WORD independently against a vocabulary built from
// words that appear in known values, and only ever flag/correct the
// specific word(s) that look mistyped. "MTC" and "Whiteplains" are already
// correct as-typed (exact vocabulary matches) and are never touched, no
// matter what the rest of the phrase says; only "Quezzzon" and "Cityy" get
// evaluated, correctly landing near "Quezon" and "City".
//
// knownValues is an array of { value, count } — full phrases (e.g. past
// Site of Vaccination entries), count = how many times that exact phrase
// has been used. alwaysKnownPhrases is an optional array of plain strings
// (no usage history required) whose words are folded into the vocabulary
// at a synthetic "well-established" weight — e.g. canonical province names
// and known correct site names, so a word is recognized as correct even
// before anyone has typed it correctly yet.
//
// Returns:
//   {
//     tier: 0 | 1 | 2 | 3,                 — highest tier seen across all words
//     correctedValue: '...' | null,          — original phrase with tier 1/2 words silently swapped in (null if nothing to auto-apply)
//     autoCorrections: [{ original, corrected, tier }],  — words actually changed (apply silently, notify which word(s))
//     suggestions: [{ original, suggestion }],           — words merely flagged (tier 3) — DO NOT apply, just hint at that word
//     original: '...'
//   }
const WORD_TIER2_MIN_USAGE = 3;        // "used before" threshold to trust a learned per-word correction
const WORD_TIER2_THRESHOLD = 0.72;     // 0..1, higher = stricter match required to auto-apply
const WORD_TIER3_THRESHOLD = 0.55;     // looser bar — enough to hint, not enough to auto-apply
const WORD_MIN_LENGTH_TO_CHECK = 3;    // words shorter than this (e.g. "of", "a") are too short to fuzzy-match reliably

// Word-level fixed corrections — single-token, near-certain fixes only.
// Multi-word TIER1_FIXED_CORRECTIONS entries (e.g. "mtc whitepalins") are
// intentionally NOT duplicated here — per-word Tier 2/3 checking below
// catches "whitepalins" -> "Whiteplains" on its own, using the vocabulary
// built from real known values instead of a hardcoded phrase list.
const WORD_FIXED_CORRECTIONS = { 'qc': 'QC', 'q.c.': 'QC', 'mm': 'Metro Manila', 'ncr': 'Metro Manila' };

// Splits a phrase into word tokens for typo-checking, dropping tokens with
// no letters/digits (bare punctuation like a lone "-").
function tokenizeWords_(text) {
  return String(text || '').trim().split(/\s+/).filter(function (w) { return /[a-z0-9]/i.test(w); });
}

// Builds a per-word vocabulary: lowercased word -> { word: <canonical
// casing seen>, count: <weight> }. Words from knownValues are weighted by
// how many times their PHRASE was used; words from alwaysKnownPhrases get
// a flat synthetic weight of WORD_TIER2_MIN_USAGE so they always count as
// "established" even with zero real usage yet.
function buildWordVocabulary_(knownValues, alwaysKnownPhrases) {
  const vocab = {};
  function addPhrase(phrase, weight) {
    tokenizeWords_(phrase).forEach(function (w) {
      const key = w.toLowerCase();
      if (!vocab[key]) vocab[key] = { word: w, count: 0 };
      vocab[key].count += weight;
    });
  }
  (knownValues || []).forEach(function (kv) { addPhrase(kv.value, kv.count); });
  (alwaysKnownPhrases || []).forEach(function (p) { addPhrase(p, WORD_TIER2_MIN_USAGE); });
  return vocab;
}

// Checks a single word against the vocabulary. Also tries the existing
// acronymMatches_ helper against multi-word vocabulary entries is NOT
// applicable per-word (acronyms are inherently multi-word shorthand), so
// that check is intentionally not used at this level — a real acronym like
// "MTC WP" is short enough (2 tokens) that Tier 3's looser threshold still
// surfaces it as a suggestion rather than silently mangling it.
function checkWordTypo_(word, vocab) {
  const key = word.toLowerCase();
  if (WORD_FIXED_CORRECTIONS[key] && WORD_FIXED_CORRECTIONS[key] !== word) {
    return { tier: 1, corrected: WORD_FIXED_CORRECTIONS[key], original: word };
  }
  if (vocab[key]) return { tier: 0, original: word }; // exact known word already — nothing to do
  if (key.length < WORD_MIN_LENGTH_TO_CHECK) return { tier: 0, original: word };

  let bestTier2 = null;
  Object.keys(vocab).forEach(function (vk) {
    const entry = vocab[vk];
    if (entry.count < WORD_TIER2_MIN_USAGE) return;
    if (Math.abs(vk.length - key.length) > 3) return; // cheap prefilter before running edit distance
    const sim = similarityRatio_(key, vk);
    if (sim >= WORD_TIER2_THRESHOLD && (!bestTier2 || entry.count > bestTier2.count)) bestTier2 = entry;
  });
  if (bestTier2) return { tier: 2, corrected: bestTier2.word, original: word, matchCount: bestTier2.count };

  let bestTier3 = null;
  Object.keys(vocab).forEach(function (vk) {
    if (Math.abs(vk.length - key.length) > 4) return;
    const sim = similarityRatio_(key, vk);
    if (sim >= WORD_TIER3_THRESHOLD && (!bestTier3 || sim > bestTier3.sim)) bestTier3 = { word: vocab[vk].word, sim: sim };
  });
  if (bestTier3) return { tier: 3, suggestion: bestTier3.word, original: word };

  return { tier: 0, original: word };
}

function checkTypoMultiWord_(typedValue, knownValues, alwaysKnownPhrases) {
  const typed = String(typedValue || '').trim();
  if (!typed) return { tier: 0, correctedValue: null, autoCorrections: [], suggestions: [], original: typed };

  const vocab = buildWordVocabulary_(knownValues, alwaysKnownPhrases);
  // Split on whitespace but KEEP the whitespace/punctuation tokens (odd
  // indices from a capturing-group split) so the corrected phrase can be
  // reassembled exactly, spacing and dashes included.
  const tokens = typed.split(/(\s+)/);
  const autoCorrections = [];
  const suggestions = [];
  let highestTier = 0;

  const rebuilt = tokens.map(function (tok) {
    if (!/[a-z0-9]/i.test(tok)) return tok; // whitespace / bare "-" — pass through untouched
    const result = checkWordTypo_(tok, vocab);
    if (result.tier === 1 || result.tier === 2) {
      autoCorrections.push({ original: tok, corrected: result.corrected, tier: result.tier });
      if (result.tier > highestTier) highestTier = result.tier;
      return result.corrected;
    }
    if (result.tier === 3) {
      suggestions.push({ original: tok, suggestion: result.suggestion });
      if (highestTier < 3) highestTier = 3;
    }
    return tok;
  });

  const correctedValue = autoCorrections.length ? rebuilt.join('') : null;
  return { tier: highestTier, correctedValue: correctedValue, autoCorrections: autoCorrections, suggestions: suggestions, original: typed };
}

// Builds the { value, count } list checkTypo_ needs for Site of
// Vaccination specifically: every distinct value already used in
// Vaccination_Tracker (Tier 2's "learned from past entries" source),
// weighted by how many times it's been used.
function getKnownSiteUsageCounts_() {
  const ss = getSheet_();
  const trackerSheet = ss.getSheetByName('Vaccination_Tracker');
  const counts = {}; // lowercased key -> { value, count }
  if (trackerSheet) {
    const data = trackerSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const v = data[i][TRACKER_COL.VACCINATION_LOCATION];
      if (!v) continue;
      const key = normalizeLookupKey_(v);
      if (!counts[key]) counts[key] = { value: v, count: 0 };
      counts[key].count++;
    }
  }
  return Object.keys(counts).map(function (k) { return counts[k]; });
}

// Canonical spellings that should always count as "known vocabulary" even
// with zero Vaccination_Tracker usage history yet — every province name
// PROVINCE_ALIASES resolves to, plus the correct-spelling side of every
// TIER1_FIXED_CORRECTIONS entry (site-name fixes like "Whiteplains").
function getAlwaysKnownSiteWords_() {
  const fromProvinces = Object.keys(PROVINCE_ALIASES).map(function (k) { return PROVINCE_ALIASES[k]; });
  const fromTier1 = Object.keys(TIER1_FIXED_CORRECTIONS).map(function (k) { return TIER1_FIXED_CORRECTIONS[k]; });
  return fromProvinces.concat(fromTier1);
}

// Client-facing: check a typed Site of Vaccination value for typos before
// save. Called on blur/change, not on every keystroke — cheap enough
// (reads Vaccination_Tracker once) but no need to run it 12 times while
// someone is still typing. Checks WORD BY WORD (see checkTypoMultiWord_
// above) so a suggestion or auto-correction only ever names the specific
// mistyped word(s), never swaps in an unrelated whole phrase.
function checkSiteTypo(typedValue) {
  assertCanEncode_();
  try {
    return checkTypoMultiWord_(typedValue, getKnownSiteUsageCounts_(), getAlwaysKnownSiteWords_());
  } catch (e) {
    console.error('checkSiteTypo error:', e);
    return { tier: 0, correctedValue: null, autoCorrections: [], suggestions: [], original: String(typedValue || '') }; // non-critical — never block the encoder over this failing
  }
}

// Client-facing: resolve a typed Site of Vaccination string to a province,
// using the NEW "Site Name - PROVINCE" suffix convention (see
// extractProvinceFromVaccinationSite_). Drives the Recipient/Vaccinator
// default-filtering in EntryForm.html — kept as a thin wrapper so the
// (large) province/alias table stays server-side only, never duplicated
// into client JS.
function getProvinceForSite(siteText) {
  assertCanEncode_();
  try {
    return extractProvinceFromVaccinationSite_(siteText) || '';
  } catch (e) {
    console.error('getProvinceForSite error:', e);
    return '';
  }
}


// Mirrors the regex the sheet itself uses (Vaccination_Entry!B31) to turn
// an Intended-Dose-Number TEXT value into the numeric "Current Dose"
// position it corresponds to in Vaccine_Schedule_Master:
//   - "Dose 3"                    -> 3
//   - "Annual Dose"/"Seasonal Dose" -> 1 (recurring schedules always key off Current Dose 1)
//   - anything else (e.g. a free-text booster label) -> null, meaning
//     "not a numbered dose we can mechanically advance from"
function deriveCurrentDoseNumber_(intendedDoseText) {
  if (!intendedDoseText) return null;
  const text = String(intendedDoseText).trim();
  const m = /^Dose\s+(\d+)$/i.exec(text);
  if (m) return parseInt(m[1], 10);
  if (/^(Annual Dose|Seasonal Dose)$/i.test(text)) return 1;
  return null;
}

// Recurring schedules (Schedule Type = "Recurring", e.g. Influenza) use
// "Annual Dose" as their dose label instead of "Dose N" — confirmed
// against real Vaccination_Tracker rows (Influenza rows use "Annual Dose"
// as both Intended Dose Number and Dose Number). "Seasonal Dose" is NOT
// one of the fixed options in the sheet's own Intended-Dose dropdown
// (Vaccination_Entry!B12 data validation lists exactly: Dose 1, Dose 2,
// Dose 3, Booster, Annual Dose, Not Applicable) — so recurring vaccines
// always map to "Annual Dose" here, even ones described as "seasonal" in
// their schedule name, to stay a selectable value.
function getRecurringDoseLabel_(scheduleDisplayName) {
  return 'Annual Dose';
}

// The sheet's fixed Intended Dose list only goes up to "Dose 3" before
// falling back to the generic "Booster" label (see INTENDED_DOSE_OPTIONS
// above) — so a schedule's 4th, 5th, etc. dose (e.g. DTaP's 5-dose
// pediatric series, Twinrix Rapid's 4-dose series) has no dedicated
// selectable value and is represented as "Booster" instead.
function doseLabelForPosition_(currentDoseNumber) {
  return (currentDoseNumber <= 3) ? ('Dose ' + currentDoseNumber) : 'Booster';
}

// NOTE ON BRAND: every value looked up from Vaccine_Schedule_Master below —
// Schedule Display Name, Total Series Doses, Interval Value/Unit, Next
// Action — is confirmed IDENTICAL across every brand that shares the same
// (Vaccine Type, Schedule Code) in the real workbook (checked
// programmatically across all 126 rows: zero brand-driven variance,
// zero cases where a schedule code maps to more than one Total Series
// Doses / interval / display name). What DOES vary Total Series Doses is
// the SCHEDULE CODE itself — e.g. Tdap's "Adult Primary / Catch-up" code
// is a 3-dose series while its "Pregnancy" code is a 1-dose series, same
// vaccine type, same brands. So brand is dropped entirely from the
// functions below (per the client's explicit instruction: "brand would
// not matter in counting the dose of the type vaccine"), but Schedule
// Code is deliberately KEPT in the match — it is what actually encodes
// the patient's series length, not a cosmetic detail.
function lookupScheduleDisplayByCode_(vaccineType, code) {
  if (!code) return '';
  const sheet = getSheet_().getSheetByName('Vaccine_Schedule_Master');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[SCHEDULE_COL.VACCINE_TYPE] === vaccineType && row[SCHEDULE_COL.SCHEDULE_CODE] === code) {
      return row[SCHEDULE_COL.SCHEDULE_DISPLAY];
    }
  }
  return '';
}

// Works out the next SELECTABLE Intended Dose label for a recipient
// continuing a series: parse the last administered dose's numeric
// position, then look up the row one position further in the same
// (vaccine type, schedule code) group — brand is NOT part of the match
// (see note above the previous function). This is deterministic and
// matches actual selectable dose values — unlike the sheet's own "Next
// Action" text, which is a human-readable advisory ("Booster if
// Continuing Risk", "Complete / Td or Tdap Booster in 10 Years") that
// often is NOT itself a value the Intended Dose dropdown offers. Returns
// { label, complete }: label is null when the last dose wasn't a numbered
// dose we can mechanically advance from (e.g. it was already a free-text
// booster entry) — in that case show the sheet's own advisory text to the
// nurse instead and leave the field open.
function computeNextIntendedDose_(vaccineType, scheduleCode, lastDoseNumberText) {
  const currentNum = deriveCurrentDoseNumber_(lastDoseNumberText);
  if (currentNum === null) return { label: null, complete: false };

  const sheet = getSheet_().getSheetByName('Vaccine_Schedule_Master');
  const data = sheet.getDataRange().getValues();
  let currentRow = null;
  let nextRow = null;
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[SCHEDULE_COL.VACCINE_TYPE] !== vaccineType || row[SCHEDULE_COL.SCHEDULE_CODE] !== scheduleCode) continue;
    if (Number(row[SCHEDULE_COL.CURRENT_DOSE]) === currentNum) currentRow = row;
    if (Number(row[SCHEDULE_COL.CURRENT_DOSE]) === currentNum + 1) nextRow = row;
  }
  if (!currentRow) return { label: null, complete: false };

  // Recurring schedules (annual flu, etc.) repeat the same dose forever —
  // there is no "next row", the current row IS the next visit.
  if (String(currentRow[SCHEDULE_COL.SCHEDULE_TYPE]).trim() === 'Recurring') {
    return { label: getRecurringDoseLabel_(currentRow[SCHEDULE_COL.SCHEDULE_DISPLAY]), complete: false };
  }

  const totalSeriesDoses = Number(currentRow[SCHEDULE_COL.TOTAL_SERIES_DOSES]);
  if (!isNaN(totalSeriesDoses) && currentNum >= totalSeriesDoses) {
    return { label: null, complete: true }; // series already fully administered
  }
  if (!nextRow) return { label: null, complete: false }; // no further row defined — can't determine mechanically

  const label = (String(nextRow[SCHEDULE_COL.SCHEDULE_TYPE]).trim() === 'Recurring')
    ? getRecurringDoseLabel_(nextRow[SCHEDULE_COL.SCHEDULE_DISPLAY])
    : doseLabelForPosition_(Number(nextRow[SCHEDULE_COL.CURRENT_DOSE]));
  return { label: label, complete: false };
}

function getRecipientVaccineHistory(recipientId, vaccineType) {
  assertCanEncode_();
  if (!recipientId || !vaccineType) return { hasHistory: false };
  const ss = getSheet_();
  const sheet = ss.getSheetByName('Vaccination_Tracker');
  if (!sheet) throw new Error('Vaccination_Tracker sheet not found');
  const data = sheet.getDataRange().getValues();

  let latest = null;
  let latestRecordNum = -1;
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[TRACKER_COL.RECIPIENT_ID] !== recipientId) continue;
    if (row[TRACKER_COL.VACCINE_TYPE] !== vaccineType) continue;
    if (row[TRACKER_COL.DISPOSITION] !== 'Administered') continue;
    const m = /(\d+)$/.exec(String(row[TRACKER_COL.RECORD_ID]));
    const recNum = m ? parseInt(m[1], 10) : i; // fall back to row order if the ID doesn't parse
    if (recNum > latestRecordNum) {
      latestRecordNum = recNum;
      latest = row;
    }
  }
  if (!latest) return { hasHistory: false };

  const lastBrand = latest[TRACKER_COL.VACCINE_BRAND];
  const lastScheduleCode = latest[TRACKER_COL.SCHEDULE_CODE];
  const lastSchedule = lookupScheduleDisplayByCode_(vaccineType, lastScheduleCode);
  const lastDoseNumberText = latest[TRACKER_COL.DOSE_NUMBER];

  // Brand-independent per the client's instruction — see the note above
  // computeNextIntendedDose_. lastBrand is still returned below so the UI
  // can show the nurse what was used last time, it just no longer
  // participates in figuring out what dose comes next.
  const nextInfo = computeNextIntendedDose_(vaccineType, lastScheduleCode, lastDoseNumberText);

  return {
    hasHistory: true,
    lastBrand: lastBrand,
    lastSchedule: lastSchedule,
    lastDoseNumber: lastDoseNumberText,
    lastVaccinationDate: formatDateForClient_(latest[TRACKER_COL.VACCINATION_DATE]),
    seriesStatus: latest[TRACKER_COL.SERIES_STATUS],
    // Human-readable advisory straight from the sheet's own Next Action
    // text (e.g. "Booster if Continuing Risk") — shown to the nurse for
    // context even when it's not a mechanically selectable dose value.
    nextDoseAdvisory: latest[TRACKER_COL.NEXT_DOSE],
    // Deterministic, selectable "Dose N" / "Annual Dose" value (or null —
    // see computeNextIntendedDose_ doc comment above).
    nextIntendedDose: nextInfo.label,
    seriesComplete: nextInfo.complete,
    // Administration Site (col W, index 22) / Route of Administration
    // (col X, index 23) — not in TRACKER_COL since nothing else in this
    // file reads them, but the client wants these PREFILLED (same as
    // Brand/Schedule) and left UNLOCKED (unlike Intended Dose), since
    // inventory availability can change which site/route is actually used
    // this time even though it was consistent last time.
    lastAdminSite: latest[22] || '',
    lastRoute: latest[23] || ''
  };
}

// For a recipient with NO prior administered dose of this vaccine type:
// the correct STARTING dose label, derived from Vaccine_Schedule_Master —
// "Dose 1" for standard/booster/complete-style series, "Annual Dose" /
// "Seasonal Dose" for Recurring schedules (confirmed against real
// Influenza rows in Vaccination_Tracker, which use "Annual Dose" rather
// than "Dose 1").
function getStartingDoseInfo_(vaccineType, brand, scheduleDisplay) {
  const sheet = getSheet_().getSheetByName('Vaccine_Schedule_Master');
  const data = sheet.getDataRange().getValues();
  let firstRow = null;
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[SCHEDULE_COL.VACCINE_TYPE] === vaccineType && row[SCHEDULE_COL.BRAND] === brand && row[SCHEDULE_COL.SCHEDULE_DISPLAY] === scheduleDisplay) {
      if (!firstRow || Number(row[SCHEDULE_COL.CURRENT_DOSE]) < Number(firstRow[SCHEDULE_COL.CURRENT_DOSE])) firstRow = row;
    }
  }
  if (!firstRow) return { label: 'Dose 1', recurring: false };
  const scheduleType = String(firstRow[SCHEDULE_COL.SCHEDULE_TYPE]).trim();
  const recurring = scheduleType === 'Recurring';
  const label = recurring ? getRecurringDoseLabel_(scheduleDisplay) : doseLabelForPosition_(Number(firstRow[SCHEDULE_COL.CURRENT_DOSE]));
  return { label: label, recurring: recurring, nextActionHint: firstRow[SCHEDULE_COL.NEXT_ACTION] || '' };
}

function getStartingDoseInfo(vaccineType, brand, scheduleDisplay) {
  assertCanEncode_();
  return getStartingDoseInfo_(vaccineType, brand, scheduleDisplay);
}

// ============================================================
//  LIVE FOLLOW-UP & SCHEDULE PREVIEW
//  Mirrors Vaccination_Entry!B31:B34 exactly, so the encoder sees Next
//  Dose / Recommended Date / Reminder Date / Series Status update live
//  as they fill the form — instead of only after saving. This does NOT
//  touch the Vaccination_Entry scratchpad at all (no writes, no flush,
//  no lock) — it is a pure read of Vaccine_Schedule_Master, safe to call
//  on every keystroke/field change without risking a race with another
//  nurse's saveVaccinationRecord().
//
//  B31's real formula filters Vaccine_Schedule_Master by
//  (Vaccine Type = B11, Brand = B19, Schedule Display = B20,
//   Current Dose = numeric position parsed from B12, Active = Yes) and
//  returns that row's Next Action (col J). B32 uses the SAME matched
//  row's Interval Value/Unit (cols K/L) applied to the vaccination date.
//  Brand is part of the sheet's own filter — but confirmed against all
//  126 rows of the real Vaccine_Schedule_Master, Next Action / Interval
//  Value / Interval Unit never differ by brand for a given
//  (Vaccine Type, Schedule Code, Current Dose) — so matching by Schedule
//  Code instead of Brand+Schedule-Display here gives an identical result
//  and doesn't require the client to have picked a Brand yet before a
//  preview can be shown. If the caller has a brand/schedule pair instead
//  of a code, pass scheduleCode='' and scheduleDisplay — either is enough
//  to resolve the row.
function previewFollowUpSchedule_(vaccineType, scheduleCode, scheduleDisplay, intendedDoseText, vaccinationDateStr) {
  const currentNum = deriveCurrentDoseNumber_(intendedDoseText);
  const sheet = getSheet_().getSheetByName('Vaccine_Schedule_Master');
  const data = sheet.getDataRange().getValues();

  let matchRow = null;
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[SCHEDULE_COL.VACCINE_TYPE] !== vaccineType) continue;
    if (!isYes_(row[SCHEDULE_COL.ACTIVE])) continue;
    if (scheduleCode) {
      if (row[SCHEDULE_COL.SCHEDULE_CODE] !== scheduleCode) continue;
    } else if (scheduleDisplay) {
      if (row[SCHEDULE_COL.SCHEDULE_DISPLAY] !== scheduleDisplay) continue;
    } else {
      continue;
    }
    if (currentNum === null) {
      // "Not Applicable" or an unparseable dose label — the sheet's own
      // formula would fall through to the FILTER's literal-text branch
      // (matching Current Dose against the raw B12 text itself), which
      // never matches a numeric Current Dose column. Net effect: no row,
      // blank Next Dose. Mirror that rather than guessing.
      continue;
    }
    if (Number(row[SCHEDULE_COL.CURRENT_DOSE]) === currentNum) { matchRow = row; break; }
  }

  if (!matchRow) {
    return { nextDose: '', recommendedDate: '', reminderDate: '', seriesStatus: '' };
  }

  const nextDose = matchRow[SCHEDULE_COL.NEXT_ACTION] || '';

  let recommendedDate = '';
  if (vaccinationDateStr && nextDose) {
    const intervalValue = Number(matchRow[SCHEDULE_COL.INTERVAL_VALUE]);
    const intervalUnit = String(matchRow[SCHEDULE_COL.INTERVAL_UNIT] || '').trim();
    const baseDate = new Date(vaccinationDateStr);
    if (!isNaN(baseDate.getTime()) && !isNaN(intervalValue) && intervalUnit) {
      let d = null;
      if (intervalUnit === 'Days') { d = new Date(baseDate.getTime()); d.setDate(d.getDate() + intervalValue); }
      else if (intervalUnit === 'Weeks') { d = new Date(baseDate.getTime()); d.setDate(d.getDate() + intervalValue * 7); }
      else if (intervalUnit === 'Months') { d = new Date(baseDate.getTime()); d.setMonth(d.getMonth() + intervalValue); }
      else if (intervalUnit === 'Years') { d = new Date(baseDate.getTime()); d.setFullYear(d.getFullYear() + intervalValue); }
      if (d) recommendedDate = d;
    }
  }

  let reminderDate = '';
  if (recommendedDate) {
    reminderDate = new Date(recommendedDate.getTime());
    reminderDate.setDate(reminderDate.getDate() - 7);
  }

  // Mirrors B34 exactly.
  const COMPLETE_VALUES = ['Complete', 'Complete / Future Booster', 'Complete / Future Guidance', 'Complete for Current Pregnancy'];
  let seriesStatus = '';
  if (nextDose) {
    if (nextDose === 'Next Seasonal Dose') seriesStatus = 'Complete - Recurring';
    else if (COMPLETE_VALUES.indexOf(nextDose) !== -1) seriesStatus = 'Complete';
    else seriesStatus = 'In Progress';
  }

  return {
    nextDose: nextDose,
    recommendedDate: formatDateForClient_(recommendedDate),
    reminderDate: formatDateForClient_(reminderDate),
    seriesStatus: seriesStatus
  };
}

// Client-facing wrapper. vaccineType/scheduleDisplay/intendedDose/
// vaccinationDate come straight from the live form state; scheduleCode is
// optional (resolved from vaccineType+brand+scheduleDisplay via
// lookupScheduleCode_ if the caller has it, otherwise pass '').
function getFollowUpPreview(vaccineType, scheduleCode, scheduleDisplay, intendedDose, vaccinationDate) {
  assertCanEncode_();
  try {
    if (!vaccineType || !intendedDose) {
      return { nextDose: '', recommendedDate: '', reminderDate: '', seriesStatus: '' };
    }
    return previewFollowUpSchedule_(vaccineType, scheduleCode || '', scheduleDisplay || '', intendedDose, vaccinationDate || '');
  } catch (e) {
    console.error('getFollowUpPreview error:', e);
    throw e;
  }
}

// --- Vaccinators ---
// Returns the FULL active roster in one call — deliberately not
// paginated or server-filtered. There are ~100 rows today; filtering by
// the encoder's province and free-text search both happen client-side
// (see EntryForm.html) so switching between "my area" and "search
// everywhere" is instant, no round-trip per keystroke. Each vaccinator
// carries a `province` tag (derived from the sheet's own Province column,
// canonicalized — see normalizeProvince_) so the client can do that
// filtering without re-deriving it itself.
function getVaccinators() {
  assertCanEncode_();
  try {
    const ss = getSheet_();
    const sheet = ss.getSheetByName('Vaccinators_Master');
    if (!sheet) throw new Error('Vaccinators_Master sheet not found');
    const data = sheet.getDataRange().getValues();
    const vaccinators = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (isYes_(row[VACCINATOR_COL.ACTIVE]) && row[VACCINATOR_COL.LAST_NAME]) {
        const display = row[VACCINATOR_COL.DISPLAY_NAME] ||
          (row[VACCINATOR_COL.LAST_NAME] + ', ' + row[VACCINATOR_COL.FIRST_NAME] +
            (row[VACCINATOR_COL.MIDDLE_NAME] ? ' ' + row[VACCINATOR_COL.MIDDLE_NAME] : '') +
            (row[VACCINATOR_COL.TITLE] ? ', ' + row[VACCINATOR_COL.TITLE] : ''));
        const rawProvince = row[VACCINATOR_COL.PROVINCE] || row[VACCINATOR_COL.PROVINCE_REGION] || '';
        vaccinators.push({
          id: row[VACCINATOR_COL.ID],
          displayName: display,
          licenseNumber: row[VACCINATOR_COL.LICENSE_NUMBER],
          email: row[VACCINATOR_COL.EMAIL] || '',
          city: row[VACCINATOR_COL.CITY] || '',
          province: rawProvince ? normalizeProvince_(rawProvince) : ''
        });
      }
    }
    return vaccinators;
  } catch (e) {
    console.error('getVaccinators error:', e);
    throw e;
  }
}

// ============================================================
//  SAVE VACCINATION RECORD
//  Writes to the Vaccination_Entry scratchpad so the workbook's
//  own formulas compute Recipient ID / next dose / recommended
//  date / series status, reads those back, then appends the full
//  row to Vaccination_Tracker (the sheet Vaccination_Report
//  actually reads from).
// ============================================================

// Row/column map for Vaccination_Entry (1-indexed).
const ENTRY = {
  RECIPIENT_NAME: 2,
  RECIPIENT_ID: 3,
  EMAIL: 4,
  DOB: 5,
  CATEGORY: 6,
  COMPANY: 7,
  DEPT: 8,
  SITE: 9,
  VACCINE_TYPE: 11,
  INTENDED_DOSE: 12,
  DISPOSITION: 13,
  REASON_CATEGORY: 14,
  SPECIFIC_REASON: 15,
  STATUS_DATE: 16,
  REVIEW_DATE: 17,
  VACCINE_BRAND: 19,
  SCHEDULE: 20,
  VACCINATION_DATE: 21,
  LOT_NUMBER: 22,
  EXPIRY_DATE: 23,
  ADMIN_SITE: 24,
  ROUTE: 25,
  LOCATION: 26,
  VACCINATOR: 27,
  VACCINATOR_LICENSE: 28,
  REMARKS: 29,
  NEXT_DOSE: 31,
  RECOMMENDED_DATE: 32,
  REMINDER_DATE: 33,
  SERIES_STATUS: 34
};
// ============================================================
//  LOT NUMBER CONSISTENCY CHECK
//  Per client instruction: "lot number in vaccination_tracker should be
//  able to track if the lot number currently entered has already been
//  used/entered during previous entries or from past records. submission
//  should not push through if this is wrong and should also alert the
//  encoder immediately."
//
//  INTERPRETATION — flagged explicitly rather than silently guessed,
//  since the literal instruction is ambiguous: a lot number is
//  legitimately reused across MANY doses drawn from the same vial/batch,
//  so treating "this lot number was used before" as itself an error would
//  block completely normal, correct data entry every single day. What IS
//  actually a data-entry error is the SAME lot number being recorded
//  against a DIFFERENT Vaccine Brand or a DIFFERENT Expiry Date than it
//  was recorded with before — a real vaccine lot has exactly one brand
//  and one expiry date, so any variance there means the Lot Number,
//  Brand, or Expiry Date was mistyped on this entry or a prior one. That
//  mismatch is what this checks and blocks on. If a different meaning was
//  intended, this is a contained, one-function change.
// ============================================================
function checkLotNumberConsistency_(lotNumber, brand, expiryDate) {
  if (!lotNumber) return { ok: true };
  const ss = getSheet_();
  const trackerSheet = ss.getSheetByName('Vaccination_Tracker');
  if (!trackerSheet) return { ok: true };
  const data = trackerSheet.getDataRange().getValues();
  const normLot = String(lotNumber).trim().toUpperCase();
  const normBrand = String(brand || '').trim();
  const expiryTime = expiryDate ? new Date(expiryDate).getTime() : null;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowLot = row[TRACKER_COL.LOT_NUMBER];
    if (!rowLot || String(rowLot).trim().toUpperCase() !== normLot) continue;

    const rowBrand = String(row[TRACKER_COL.VACCINE_BRAND] || '').trim();
    if (normBrand && rowBrand && rowBrand !== normBrand) {
      return {
        ok: false,
        reason: 'brand',
        message: 'Lot Number "' + lotNumber + '" was previously recorded under Vaccine Brand "' + rowBrand +
          '", not "' + normBrand + '". A vaccine lot has exactly one brand — please double-check the Lot Number and Brand for typos before submitting.'
      };
    }

    const rowExpiry = row[TRACKER_COL.EXPIRY_DATE];
    if (expiryTime && rowExpiry instanceof Date && !isNaN(expiryTime)) {
      // Compare by calendar day, not exact timestamp — a Date cell can
      // carry a time-of-day component that isn't meaningful for an
      // expiry date.
      const rowDay = new Date(rowExpiry.getFullYear(), rowExpiry.getMonth(), rowExpiry.getDate()).getTime();
      const thisExpiry = new Date(expiryTime);
      const thisDay = new Date(thisExpiry.getFullYear(), thisExpiry.getMonth(), thisExpiry.getDate()).getTime();
      if (rowDay !== thisDay) {
        return {
          ok: false,
          reason: 'expiry',
          message: 'Lot Number "' + lotNumber + '" was previously recorded with Expiry Date ' +
            Utilities.formatDate(rowExpiry, Session.getScriptTimeZone(), 'MMM d, yyyy') +
            ', which does not match the expiry date entered now. A vaccine lot has exactly one expiry date — please double-check for typos before submitting.'
        };
      }
    }
  }
  return { ok: true };
}

// Client-facing: live check as the encoder fills in Lot Number (called on
// blur/change from EntryForm.html), so a mismatch is flagged immediately
// rather than only discovered at submission time.
function checkLotNumber(lotNumber, brand, expiryDate) {
  assertCanEncode_();
  try {
    return checkLotNumberConsistency_(lotNumber, brand, expiryDate);
  } catch (e) {
    console.error('checkLotNumber error:', e);
    throw e;
  }
}

// ------------------------------------------------------------
// Route-of-Administration WARNING (not a hard block) — per client
// instruction: "you can add a rule that warns the encoder if they try to
// save an MMR vaccine as 'IM'." Live vaccines given by the wrong route is
// a real clinical concern worth flagging, but unlike the Lot Number
// check above this is intentionally NOT enforced in saveVaccinationRecord
// — a genuine clinical override (or a vaccine type this list hasn't
// caught up to) must still be saveable, just with the encoder having
// seen the warning first. This mirrors checkSiteTypo's non-blocking
// tier-3 pattern rather than checkLotNumberConsistency_'s hard block.
//
// VACCINE_ROUTE_HINTS below is deliberately small and conservative: only
// vaccine types where SC-vs-IM is a well-established, unambiguous rule
// are listed (see the corrected LDS_Project_corrected.xlsx rows this
// session — MMR and Varicella are SC; Imojev is SC per this workbook's
// own Vaccine_Schedule_Master notes). Do NOT add Hepatitis A/Havrix here
// — that vaccine's route depends on which brand/formulation, which this
// simple type-keyed lookup can't distinguish; see this session's note on
// why the Havrix rows were deliberately left uncorrected.
const VACCINE_ROUTE_HINTS = {
  'MMR': 'SC',
  'Varicella': 'SC',
  'Imojev': 'SC'
};
function checkRouteWarning(vaccineType, route) {
  assertCanEncode_();
  if (!vaccineType || !route) return { warn: false };
  const expected = VACCINE_ROUTE_HINTS[String(vaccineType).trim()];
  if (!expected) return { warn: false };
  if (String(route).trim().toUpperCase() === expected) return { warn: false };
  return {
    warn: true,
    message: vaccineType + ' is conventionally given ' + expected + ', not ' + route +
      '. Please double-check the route before saving — this is a warning only, the record can still be saved as entered.'
  };
}

const ENTRY_INPUT_ROWS = [
  ENTRY.RECIPIENT_NAME, ENTRY.VACCINE_TYPE, ENTRY.INTENDED_DOSE, ENTRY.DISPOSITION,
  ENTRY.REASON_CATEGORY, ENTRY.SPECIFIC_REASON, ENTRY.STATUS_DATE, ENTRY.REVIEW_DATE,
  ENTRY.VACCINE_BRAND, ENTRY.SCHEDULE, ENTRY.VACCINATION_DATE, ENTRY.LOT_NUMBER,
  ENTRY.EXPIRY_DATE, ENTRY.ADMIN_SITE, ENTRY.ROUTE, ENTRY.LOCATION, ENTRY.VACCINATOR,
  ENTRY.REMARKS
];

// ------------------------------------------------------------
// SELF-HEALING DATA VALIDATION for Vaccination_Entry input cells.
//
// Real bug this fixes: several Vaccination_Entry cells (confirmed:
// Vaccine Brand / B19) carry a native Google Sheets dropdown validation
// rule set to "Reject input", built directly into the workbook template —
// Code.gs does not create or manage this rule. That static list is a
// SEPARATE, unsynchronized copy of the same choices the web app already
// computes live from Vaccine_Schedule_Master (getBrands()/getSchedules(),
// both filtered to "active" rows only).
//
// This becomes a real, patient-data-losing bug specifically BECAUSE of an
// earlier, deliberate feature: brand prefill intentionally offers a
// recipient's own previously-used brand for a vaccine type even if that
// brand has since been deactivated/discontinued in Vaccine_Schedule_Master
// (a patient continuing an existing series on a brand no longer offered
// to NEW patients is a legitimate, correct case — see the "(last used —
// you can pick a different brand below)" option in EntryForm.html). The
// web app correctly allows that historical value through. The sheet's own
// hard-reject dropdown does not know about it, throws mid-write, and
// because the throw happens partway through saveVaccinationRecord's write
// sequence, the whole record is lost with no data saved anywhere — not a
// cosmetic validation warning, a silent full save failure.
//
// Fix: rather than trying to keep two independently-maintained lists in
// sync forever (fragile, and this will recur the next time
// Vaccine_Schedule_Master changes), relax any existing "reject" rule on
// these cells to "warn" (setAllowInvalid(true)) the moment before we
// write to them. This preserves the dropdown/typo-catching UI for anyone
// editing the sheet BY HAND directly, while guaranteeing the web app's
// own already-validated submissions (validated against the live master
// data one layer up, in getBrands_/getSchedules_) can never be silently
// blocked and lose a patient's record. If a cell has no validation rule
// at all, this is a harmless no-op.
// ------------------------------------------------------------
function ensureRangeValidationAllowsInvalid_(range) {
  try {
    const rule = range.getDataValidation();
    if (!rule) return;
    if (rule.getAllowInvalid()) return; // already non-blocking — nothing to do
    const rebuilt = SpreadsheetApp.newDataValidation()
      .withCriteria(rule.getCriteriaType(), rule.getCriteriaValues())
      .setAllowInvalid(true)
      .setHelpText(rule.getHelpText() || '')
      .build();
    range.setDataValidation(rebuilt);
  } catch (e) {
    // Never let a validation-repair problem block an actual save — this is
    // a defensive best-effort relax, not a required step.
    console.error('ensureRangeValidationAllowsInvalid_ error:', e);
  }
}

function relaxEntryCellValidation_(entrySheet) {
  ENTRY_INPUT_ROWS.forEach(function (row) {
    ensureRangeValidationAllowsInvalid_(entrySheet.getRange(row, 2));
  });
}

function saveVaccinationRecord(record) {
  const callerRole = assertCanEncode_();

  // CHANGED per client instruction: Site of Vaccination is now typed by
  // the encoder on every submission rather than locked server-side from
  // ENCODER_SITE_MAP. This is a deliberate trade-off worth being explicit
  // about — the old behaviour meant a nurse could not accidentally (or
  // deliberately) log a record under the wrong site, because the server
  // ignored whatever the browser sent and substituted the mapped value.
  // That guarantee is gone now: the dashboard's site grouping is only as
  // reliable as what each encoder actually types. What we keep instead is
  // (a) the field is still required, not optional, and (b) whatever an
  // encoder types is saved to THEIR OWN account via rememberEncodingSite_
  // below so they aren't retyping it every time, not shared or writable
  // by anyone else.
  if (!record.location) {
    throw new Error('Please enter the site of vaccination.');
  }

  const lock = LockService.getScriptLock();
  const gotLock = lock.tryLock(30000); // wait up to 30s
  if (!gotLock) {
    throw new Error('The system is busy processing another record. Please try again in a few seconds.');
  }
  try {
    const ss = getSheet_();
    const entrySheet = ss.getSheetByName('Vaccination_Entry');
    const trackerSheet = ss.getSheetByName('Vaccination_Tracker');
    const scheduleSheet = ss.getSheetByName('Vaccine_Schedule_Master');
    if (!entrySheet) throw new Error('Vaccination_Entry sheet not found');
    if (!trackerSheet) throw new Error('Vaccination_Tracker sheet not found');
    if (!scheduleSheet) throw new Error('Vaccine_Schedule_Master sheet not found');

    // 0) Lot Number consistency — hard block. Re-checked here (not just
    // client-side on blur) under the script lock, both because the
    // encoder could bypass/skip the live check and because this is the
    // authoritative point where "submission should not push through if
    // this is wrong" is actually enforced. See checkLotNumberConsistency_
    // above for what "wrong" means here.
    if (record.lotNumber) {
      const lotCheck = checkLotNumberConsistency_(record.lotNumber, record.vaccineBrand, record.expiryDate);
      if (!lotCheck.ok) {
        throw new Error(lotCheck.message);
      }
    }

    // 1) Write raw inputs into the Entry scratchpad.
    //    Relax any hard-reject sheet-level dropdown validation on these
    //    cells first — see relaxEntryCellValidation_ above for why this is
    //    necessary (a legitimately-prefilled historical brand/schedule can
    //    otherwise be rejected by a stale native validation rule, silently
    //    losing the whole record).
    relaxEntryCellValidation_(entrySheet);
    entrySheet.getRange(ENTRY.RECIPIENT_NAME, 2).setValue(record.recipientNameForLookup || record.recipientName);
    entrySheet.getRange(ENTRY.VACCINE_TYPE, 2).setValue(record.vaccineType);
    entrySheet.getRange(ENTRY.INTENDED_DOSE, 2).setValue(record.intendedDose);
    entrySheet.getRange(ENTRY.DISPOSITION, 2).setValue(record.doseDisposition);
    entrySheet.getRange(ENTRY.REASON_CATEGORY, 2).setValue(record.reasonCategory || '');
    entrySheet.getRange(ENTRY.SPECIFIC_REASON, 2).setValue(record.specificReason || '');
    entrySheet.getRange(ENTRY.STATUS_DATE, 2).setValue(record.statusDate ? parseDateOnlyLocal_(record.statusDate) : todayDateOnly_());
    stampDateOnlyFormat_(entrySheet.getRange(ENTRY.STATUS_DATE, 2));
    entrySheet.getRange(ENTRY.REVIEW_DATE, 2).setValue(record.reviewDate ? parseDateOnlyLocal_(record.reviewDate) : '');
    stampDateOnlyFormat_(entrySheet.getRange(ENTRY.REVIEW_DATE, 2));
    entrySheet.getRange(ENTRY.VACCINE_BRAND, 2).setValue(record.vaccineBrand);
    entrySheet.getRange(ENTRY.SCHEDULE, 2).setValue(record.schedule);
    entrySheet.getRange(ENTRY.VACCINATION_DATE, 2).setValue(record.vaccinationDate ? parseDateOnlyLocal_(record.vaccinationDate) : '');
    stampDateOnlyFormat_(entrySheet.getRange(ENTRY.VACCINATION_DATE, 2));
    entrySheet.getRange(ENTRY.LOT_NUMBER, 2).setValue(record.lotNumber || '');
    entrySheet.getRange(ENTRY.EXPIRY_DATE, 2).setValue(record.expiryDate ? parseDateOnlyLocal_(record.expiryDate) : '');
    stampDateOnlyFormat_(entrySheet.getRange(ENTRY.EXPIRY_DATE, 2));
    entrySheet.getRange(ENTRY.ADMIN_SITE, 2).setValue(record.adminSite || '');
    entrySheet.getRange(ENTRY.ROUTE, 2).setValue(record.route || '');
    entrySheet.getRange(ENTRY.LOCATION, 2).setValue(record.location || '');
    entrySheet.getRange(ENTRY.VACCINATOR, 2).setValue(record.vaccinator || '');
    entrySheet.getRange(ENTRY.REMARKS, 2).setValue(record.remarks || '');

    // 2) Force the workbook to recalculate, then read back computed cells.
    SpreadsheetApp.flush();

    const computedRecipientId = entrySheet.getRange(ENTRY.RECIPIENT_ID, 2).getValue();
    const computedVaccinatorLicense = entrySheet.getRange(ENTRY.VACCINATOR_LICENSE, 2).getValue();
    const nextDose = entrySheet.getRange(ENTRY.NEXT_DOSE, 2).getValue();
    const recommendedDate = entrySheet.getRange(ENTRY.RECOMMENDED_DATE, 2).getValue();
    const reminderDate = entrySheet.getRange(ENTRY.REMINDER_DATE, 2).getValue();
    const seriesStatus = entrySheet.getRange(ENTRY.SERIES_STATUS, 2).getValue();

    if (record.recipientId && computedRecipientId && String(computedRecipientId) !== String(record.recipientId)) {
      // The name-based lookup on the sheet resolved to a different Recipient ID
      // than the one the nurse actually selected in the dropdown — almost
      // certainly two recipients share the exact same Last/First/Middle name.
      // Refuse to save rather than silently attaching the record to the
      // wrong person.
      throw new Error(
        'Recipient name lookup mismatch: selected ID ' + record.recipientId +
        ' but the sheet resolved "' + (record.recipientNameForLookup || record.recipientName) +
        '" to ' + computedRecipientId + '. Likely a duplicate name in Recipients_Master. ' +
        'Record NOT saved — please contact your administrator.'
      );
    }

    // 3) Schedule Code lookup (Vaccine_Schedule_Master col E), not exposed on Entry.
    const scheduleCode = lookupScheduleCode_(scheduleSheet, record.vaccineType, record.vaccineBrand, record.schedule);

    // 4) Derive Schedule Status / Reminder Status.
    //    NOTE: These two columns are NOT computed anywhere in Vaccination_Entry.
    //    The rule below is reverse-engineered from the sample rows in
    //    Vaccination_Tracker for "Administered", "Deferred" and "Declined".
    //    "No-show" and "Contraindicated – Temporary/Permanent" do not appear
    //    in any sample row — CONFIRM these two cases with [COMPANY_NAME]
    //    before go-live; the labels below are a reasonable placeholder, not
    //    a verified spec.
    const statusPair = deriveScheduleAndReminderStatus_(record.doseDisposition, nextDose, recommendedDate);

    // 5) Build the Vaccination_Tracker row (36 columns, A:AJ) and append.
    const nextId = getNextTrackerId_(trackerSheet);
    const row = [
      nextId,                                   // A  Vaccination Record ID
      computedRecipientId || record.recipientId,// B  Recipient ID
      record.recipientNameForLookup || record.recipientName, // C Recipient Name
      record.email || '',                       // D  Email Address
      record.dob || '',                         // E  Date of Birth
      record.category || '',                    // F  Recipient Category
      record.company || '',                     // G  Company / Organization
      '',                                        // H  Department / Unit ID — CHANGED per client instruction: this column is deliberately never populated by the web app anymore.
      record.department || '',                  // I  Department / Unit
      record.rawAssignedSite || '',              // J  Assigned Site/Location — carried through silently from the recipient's own record for continuity with historical rows; not shown or editable anywhere in the web app (client instruction: "take Assigned Site/Location out of the website").
      record.vaccineType,                       // K  Vaccine Type
      record.intendedDose,                      // L  Intended Dose Number
      record.doseDisposition,                   // M  Dose Disposition
      record.reasonCategory || '',              // N  Reason Category
      record.specificReason || '',              // O  Specific Reason
      record.statusDate ? parseDateOnlyLocal_(record.statusDate) : todayDateOnly_(), // P Status Date
      record.reviewDate ? parseDateOnlyLocal_(record.reviewDate) : '', // Q Review / Reschedule Date
      record.vaccineBrand,                      // R  Vaccine Brand
      record.intendedDose || '',                // S Dose Number (mirrors Intended Dose Number text exactly — "Dose 1", "Annual Dose", etc. — confirmed against real Vaccination_Tracker rows; this is NOT a bare number)
      record.vaccinationDate ? parseDateOnlyLocal_(record.vaccinationDate) : '', // T Vaccination Date
      record.lotNumber || '',                   // U  Lot Number
      record.expiryDate ? parseDateOnlyLocal_(record.expiryDate) : '', // V Expiry Date
      record.adminSite || '',                   // W  Administration Site
      record.route || '',                       // X  Route of Administration
      record.location || '',                    // Y  Vaccination Location
      record.vaccinator || '',                  // Z  Vaccinator
      computedVaccinatorLicense || '',          // AA Vaccinator License Number
      scheduleCode || '',                       // AB Schedule Code
      nextDose || '',                           // AC Next Dose
      recommendedDate || '',                    // AD Recommended Next Dose Date
      '',                                        // AE Scheduled Appointment Date (set later, not at intake)
      reminderDate || '',                       // AF Reminder Date
      statusPair.scheduleStatus,                // AG Schedule Status
      statusPair.reminderStatus,                // AH Reminder Status
      '',                                        // AI Reminder Sent Date
      seriesStatus || '',                       // AJ Series Status
      'No',                                      // AK Schedule Override
      record.remarks || ''                      // AL Remarks
    ];
    trackerSheet.appendRow(row);
    // Force the date-only columns of the row we just appended to a
    // date-only display format, regardless of what the column's existing
    // format was — belt-and-suspenders alongside parseDateOnlyLocal_ above,
    // see the comment on stampDateOnlyFormat_.
    (function () {
      var newRow = trackerSheet.getLastRow();
      [TRACKER_COL.STATUS_DATE, TRACKER_COL.REVIEW_DATE, TRACKER_COL.VACCINATION_DATE, TRACKER_COL.EXPIRY_DATE].forEach(function (colIdx) {
        stampDateOnlyFormat_(trackerSheet.getRange(newRow, colIdx + 1));
      });
    })();

    // 6) Clear the scratchpad so the next nurse starts from a blank form
    //    and stale values from this submission don't linger on screen for
    //    anyone else with the sheet open.
    ENTRY_INPUT_ROWS.forEach(function (r) {
      entrySheet.getRange(r, 2).setValue('');
    });
    SpreadsheetApp.flush();

    // 7) Remember this encoder's typed site for next time (see
    //    getEncodingSite/rememberEncodingSite_ above). Deliberately AFTER
    //    the append succeeds, not before — if the save throws above, we
    //    don't want to have already overwritten a good remembered value
    //    with a typo the encoder never actually submitted.
    rememberEncodingSite_(record.location);

    return {
      success: true,
      message: 'Record saved successfully.',
      recordId: nextId,
      nextDose: nextDose,
      // Same Date-across-the-bridge precaution as getRecipients()/
      // getDashboardData() — recommendedDate here comes straight off a
      // formula cell and can be a real Date object.
      recommendedDate: formatDateForClient_(recommendedDate),
      reminderDate: formatDateForClient_(reminderDate),
      seriesStatus: seriesStatus
    };
  } catch (e) {
    console.error('saveVaccinationRecord error:', e);
    throw e;
  } finally {
    lock.releaseLock();
  }
}

function lookupScheduleCode_(scheduleSheet, vaccineType, brand, scheduleDisplay) {
  const data = scheduleSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0] === vaccineType && row[2] === brand && row[5] === scheduleDisplay) {
      return row[4]; // column E, Schedule Code
    }
  }
  return '';
}

function deriveScheduleAndReminderStatus_(disposition, nextDose, recommendedDate) {
  const COMPLETE_VALUES = ['Complete', 'Complete / Future Booster', 'Complete / Future Guidance', 'Complete for Current Pregnancy'];
  if (!disposition) return { scheduleStatus: '', reminderStatus: '' };
  if (disposition === 'Deferred') return { scheduleStatus: 'Deferred - Review Scheduled', reminderStatus: 'Pending' };
  if (disposition === 'No-show') return { scheduleStatus: 'No-show - Reschedule Required', reminderStatus: 'Pending' };
  if (disposition === 'Declined') return { scheduleStatus: 'Declined', reminderStatus: 'Not Required' };
  if (String(disposition).indexOf('Contraindicated') !== -1) {
    return { scheduleStatus: 'Contraindicated - Clinical Hold', reminderStatus: 'Not Required' };
  }
  // Administered
  if (recommendedDate) return { scheduleStatus: 'Due Date Generated', reminderStatus: 'Pending' };
  if (COMPLETE_VALUES.indexOf(nextDose) !== -1) return { scheduleStatus: 'Complete', reminderStatus: 'Not Required' };
  return { scheduleStatus: 'Clinical Review', reminderStatus: 'Not Required' };
}

function getNextTrackerId_(trackerSheet) {
  const lastRow = trackerSheet.getLastRow();
  if (lastRow < 2) return 'VAX-000001';
  const lastId = trackerSheet.getRange(lastRow, 1).getValue();
  const match = /(\d+)$/.exec(String(lastId));
  const nextNum = match ? (parseInt(match[1], 10) + 1) : (lastRow); // fallback
  return 'VAX-' + String(nextNum).padStart(6, '0');
}

// ------------------------------------------------------------
// Recipient ID auto-generation — per client instruction: "follow the
// existing VAC-000001 pattern and extend it." Same scan-the-last-row-
// and-increment approach as getNextTrackerId_ above (Vaccination_Tracker
// uses VAX-######, Recipients_Master uses VAC-######) — deliberately NOT
// unified into one shared function even though the logic is identical,
// because the two sheets/prefixes must never cross-contaminate each
// other's numbering if a row is ever deleted from just one of them.
function generateRecipientId_(recipientsSheet) {
  const lastRow = recipientsSheet.getLastRow();
  if (lastRow < 2) return 'VAC-000001';
  const lastId = recipientsSheet.getRange(lastRow, 1).getValue();
  const match = /(\d+)$/.exec(String(lastId));
  const nextNum = match ? (parseInt(match[1], 10) + 1) : (lastRow); // fallback, same as getNextTrackerId_
  return 'VAC-' + String(nextNum).padStart(6, '0');
}

// Backs the "Add New Recipient" modal in EntryForm.html.
// Appends one new row to Recipients_Master and returns the generated ID
// so the entry form can immediately select the new recipient without a
// second round trip.
//
// COLUMN NOTE: every Recipients_Master column is now collected/populated
// here except column P (Date Added), which is system-generated (see
// below) rather than typed by the encoder. Column B (Employee/Client ID)
// is the only one treated as optional per client instruction — everything
// else is asked for in the modal, though only Last/First Name are
// actually enforced as required (matching this function's own validation
// below).
// Duplicate-name check — per client instruction ("make a duplicate
// check"). Deliberately a WARN-AND-CONFIRM, not a hard block: two
// genuinely different people CAN share the exact same name (this is the
// same reason saveVaccinationRecord's own duplicate-name guard above
// refuses to silently pick one instead of erroring). Matches on
// normalized Last+First+Middle only — case/whitespace-insensitive via
// normalizeLookupKey_, same helper the typo/site matching already uses.
function findDuplicateRecipientNames_(sheet, lastName, firstName, middleName) {
  const targetKey = normalizeLookupKey_(lastName + ', ' + firstName + (middleName ? ' ' + middleName : ''));
  const data = sheet.getDataRange().getValues();
  const matches = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    const existingName = row[2] + ', ' + row[3] + (row[4] ? ' ' + row[4] : '');
    if (normalizeLookupKey_(existingName) === targetKey) {
      matches.push({ id: row[0], name: existingName, company: row[8] || '', dept: row[10] || '' });
    }
  }
  return matches;
}

function addRecipient(fields) {
  const callerRole = assertCanEncode_();
  if (!fields || !fields.lastName || !fields.firstName) {
    throw new Error('Last name and first name are required.');
  }
  const lock = LockService.getScriptLock();
  const gotLock = lock.tryLock(30000);
  if (!gotLock) {
    throw new Error('The system is busy. Please try again in a few seconds.');
  }
  try {
    const ss = getSheet_();
    const sheet = ss.getSheetByName('Recipients_Master');
    if (!sheet) throw new Error('Recipients_Master sheet not found');

    // Run the duplicate check INSIDE the lock, not just client-side before
    // calling this function — that's what makes it authoritative rather
    // than racy (two encoders adding the same name within the same second
    // would otherwise both pass a client-side-only check). If matches are
    // found and the caller hasn't already confirmed, refuse to write and
    // hand back the matches so the UI can show them and let a person
    // decide — see fields.confirmDuplicate below.
    const dupes = findDuplicateRecipientNames_(sheet, fields.lastName, fields.firstName, fields.middleName);
    if (dupes.length && !fields.confirmDuplicate) {
      return { success: false, duplicate: true, matches: dupes };
    }

    const newId = generateRecipientId_(sheet);
    const fullName = fields.lastName + ', ' + fields.firstName + (fields.middleName ? ' ' + fields.middleName : '');

    // Columns, by index — full map now confirmed (client-supplied for
    // J/O/P/R): A id, B employee/client ID (optional), C last, D first,
    // E middle, F dob, G sex, H category, I company, J department/unit ID,
    // K department/unit (name), L assigned site, M email, N mobile,
    // O enrollment source, P date added, Q active, R remarks, S full name.
    const row = [];
    row[0] = newId;
    row[1] = fields.employeeId || ''; // Employee / Client ID — optional, per client instruction
    row[2] = fields.lastName;
    row[3] = fields.firstName;
    row[4] = fields.middleName || '';
    row[5] = fields.dob ? parseDateOnlyLocal_(fields.dob) : '';
    row[6] = fields.sex || '';
    row[7] = fields.category || '';
    row[8] = fields.company || '';
    row[9] = fields.deptId || ''; // Department / Unit ID
    row[10] = fields.dept || '';
    row[11] = fields.assignedSite || '';
    row[12] = fields.email || '';
    row[13] = fields.mobile || '';
    row[14] = fields.enrollmentSource || ''; // Enrollment Source
    // Date Added is system-generated, not collected from the form — a
    // "when was this recipient actually added" audit field should reflect
    // the real creation moment, not whatever the encoder might type.
    row[15] = new Date();
    row[16] = 'Yes'; // Active — required for the new recipient to show up in getRecipients()
    row[17] = fields.remarks || ''; // Remarks
    row[18] = fullName;

    sheet.appendRow(row);
    // DOB must never display a time component — force the format on the
    // cell we just wrote regardless of what the column's existing format
    // was (see stampDateOnlyFormat_ / parseDateOnlyLocal_ above).
    stampDateOnlyFormat_(sheet.getRange(sheet.getLastRow(), 6));
    return { success: true, recipientId: newId, name: fullName };
  } catch (e) {
    console.error('addRecipient error:', e);
    throw e;
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// AUTO-ID ON DIRECT SHEET EDIT — for bulk pre-registration: admins paste
// or type a client-provided recipient list straight into
// Recipients_Master (before the program even starts, well outside the
// Entry Form), and each qualifying row should get its VAC-###### the
// moment it's real rather than someone typing IDs in by hand one at a
// time. Fires once a row has BOTH Last Name (col C) and First Name
// (col D) filled and no ID yet in col A — same minimum addRecipient()
// itself requires, so a row counts as "real" at the same point either
// entry path would consider it real. DOB and everything else can be
// filled in before, after, or never — they don't gate ID assignment.
//
// ONE-TIME SETUP REQUIRED, PER CLIENT PROJECT: unlike a same-project
// bound script, this project is a STANDALONE Apps Script (see CONFIG.
// sheetId / getSheet_() above) once it's gone through auto-provisioning,
// and a plain function named onEdit() ONLY auto-fires for a script's own
// container-bound spreadsheet — it does nothing at all for a Sheet a
// standalone script merely opens by ID. There is also no REST API call
// that can install this trigger remotely the way provisionVxSyncClient
// installs the web app deployment; Google only allows an installable
// trigger like this to be created by actually running code inside the
// project. So: after this Code.gs is deployed to a client (new or
// updated), open that project in the Apps Script editor, pick
// setupRecipientAutoIdTrigger from the function dropdown, and click Run
// once (approving the permissions prompt that appears). That's a
// one-time action per project — it does not need to be repeated on every
// later Code.gs update, only if the trigger is ever deleted or the
// project is recreated from scratch.
// ============================================================
function setupRecipientAutoIdTrigger() {
  var ss = getSheet_();
  // Remove any previous copy of this exact trigger first, so re-running
  // this (e.g. after recreating the project) never leaves duplicates
  // that would each assign a row two IDs.
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'recipientsMasterOnEdit_') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('recipientsMasterOnEdit_')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  // Also fix date-only columns' display format for the WHOLE column, not
  // just rows the web app writes — this is what makes DOB never show a
  // time even when someone types it directly into Recipients_Master
  // rather than through the Entry/Add-Recipient forms. Applied to the
  // entire column (not just existing rows) so it keeps holding for every
  // future row too, typed or app-written. Piggybacks on this same
  // one-time "run me once per project" setup step rather than adding a
  // second one for the client to remember.
  var recipientsSheet = ss.getSheetByName('Recipients_Master');
  if (recipientsSheet) {
    stampDateOnlyFormat_(recipientsSheet.getRange('F2:F')); // Date of Birth
  }
  var entrySheet = ss.getSheetByName('Vaccination_Entry');
  if (entrySheet) {
    stampDateOnlyFormat_(entrySheet.getRange(ENTRY.STATUS_DATE, 2));
    stampDateOnlyFormat_(entrySheet.getRange(ENTRY.REVIEW_DATE, 2));
    stampDateOnlyFormat_(entrySheet.getRange(ENTRY.VACCINATION_DATE, 2));
    stampDateOnlyFormat_(entrySheet.getRange(ENTRY.EXPIRY_DATE, 2));
  }
  var trackerSheet = ss.getSheetByName('Vaccination_Tracker');
  if (trackerSheet) {
    [TRACKER_COL.STATUS_DATE, TRACKER_COL.REVIEW_DATE, TRACKER_COL.VACCINATION_DATE, TRACKER_COL.EXPIRY_DATE].forEach(function (colIdx) {
      stampDateOnlyFormat_(trackerSheet.getRange(2, colIdx + 1, Math.max(trackerSheet.getMaxRows() - 1, 1), 1));
    });
  }
}

function recipientsMasterOnEdit_(e) {
  try {
    if (!e || !e.range) return; // e.g. someone ran this manually from the editor rather than a real edit firing it
    var sheet = e.range.getSheet();
    if (sheet.getName() !== 'Recipients_Master') return;

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

    var lock = LockService.getScriptLock();
    // A short wait, not the 30s addRecipient() uses — this is a simple/
    // installable trigger reacting to something that already happened on
    // the sheet, not a person waiting on a submit button. If another
    // edit's trigger run is mid-flight, this one's rows are already
    // sitting on the sheet and will get picked up by the next edit to
    // touch them (or a manual re-save) rather than making someone wait.
    if (!lock.tryLock(10000)) return;
    try {
      assignMissingRecipientIds_(sheet, startRow, numRows);
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    // Must never throw back into the edit itself — an uncaught error
    // here would look to the person typing like their edit failed.
    console.error('recipientsMasterOnEdit_ error:', err);
  }
}

function assignMissingRecipientIds_(sheet, startRow, numRows) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // A running max over the WHOLE id column, not just "the last row" —
  // unlike the Entry Form (which always appends at the bottom in order),
  // a bulk paste can land rows out of physical order, so the true next
  // number has to be the highest one anywhere in the column.
  var allIds = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var maxNum = 0;
  for (var i = 0; i < allIds.length; i++) {
    var m = /VAC-(\d+)$/.exec(String(allIds[i][0] || ''));
    if (m) {
      var n = parseInt(m[1], 10);
      if (n > maxNum) maxNum = n;
    }
  }

  // Read every existing name once up front for the duplicate flag below,
  // rather than re-scanning the whole sheet per row.
  var allNames = sheet.getRange(2, 3, lastRow - 1, 3).getValues(); // C,D,E = last, first, middle

  var endRow = Math.min(startRow + numRows - 1, lastRow);
  for (var r = startRow; r <= endRow; r++) {
    var idCell = sheet.getRange(r, 1);
    if (idCell.getValue()) continue; // already has an ID — never reassign or remove it

    var last = sheet.getRange(r, 3).getValue();
    var first = sheet.getRange(r, 4).getValue();
    if (!last || !first) continue; // not yet a "real" row

    maxNum += 1;
    idCell.setValue('VAC-' + String(maxNum).padStart(6, '0'));

    // Duplicate flag — non-blocking (a visible note, not a hard stop):
    // there's no one to click "confirm anyway" on an unattended sheet
    // edit the way the Entry Form's modal allows, and two different
    // people CAN genuinely share a name.
    var middle = sheet.getRange(r, 5).getValue();
    var targetKey = normalizeLookupKey_(last + ', ' + first + (middle ? ' ' + middle : ''));
    var dupeRow = -1;
    for (var j = 0; j < allNames.length; j++) {
      var otherRow = j + 2;
      if (otherRow === r) continue;
      var oLast = allNames[j][0], oFirst = allNames[j][1], oMiddle = allNames[j][2];
      if (!oLast || !oFirst) continue;
      if (normalizeLookupKey_(oLast + ', ' + oFirst + (oMiddle ? ' ' + oMiddle : '')) === targetKey) {
        dupeRow = otherRow;
        break;
      }
    }
    idCell.setNote(dupeRow > 0
      ? '⚠ Possible duplicate: same name as row ' + dupeRow + ' in this sheet. Verify before treating these as two different people.'
      : null);
  }
}


// ============================================================
//  DASHBOARD API
//  FINAL split, per explicit client instruction: "the dashboard should
//  rely in vaccination_report's formulas, the filters shouldn't since
//  there's no existing formula for that in the sheets."
//    - The DEFAULT/unfiltered view (top stats, Action Required, Upcoming
//      Doses, Site/Location Summary) reads Vaccination_Report's own
//      pre-built formula cells — see getDashboardStats/getActionRequired/
//      getUpcomingDoses/getSiteSummary below. The sheet is the source of
//      truth there; this file does not recompute those numbers.
//    - The Department filter and the Site-of-Vaccination filter have no
//      corresponding sheet formula (Department/Unit grouping and the
//      site-of-vaccination cascade are both new), so those two modes are
//      computed live from Vaccination_Tracker + Recipients_Master via
//      computeDashboardData_ below. The always-on Department Follow-Up
//      Summary table is treated the same way — it replaced the sheet's
//      per-Assigned-Site summary and has no formula backing of its own,
//      so it stays live-computed via getDepartmentSummary too.
//
//  Filtering modes:
//    'none'       — everything, unfiltered (sheet-based)
//    'department' — scoped to one Department/Unit value (live-computed)
//    'site'       — scoped to one Site-of-Vaccination SITE NAME, i.e. the
//                   text before "-" in "Site Name - PROVINCE" (live-
//                   computed; province is only used client-side to build
//                   the cascading dropdown, via getFilterOptionsV2 below)
// ============================================================

const COMPLETE_SERIES_STATUSES = ['Complete', 'Complete / Future Booster', 'Complete / Future Guidance', 'Complete for Current Pregnancy', 'Complete - Recurring'];
const ACTION_DISPOSITIONS = ['Deferred', 'No-show'];

function inferRequiredAction_(disposition) {
  if (disposition === 'Deferred' || disposition === 'No-show') return 'Review / Reassess';
  if (disposition === 'Declined') return 'Documented — No Further Action';
  if (String(disposition).indexOf('Contraindicated') !== -1) return 'Clinical Hold — Reassess Eligibility';
  return '';
}

// Populates: the Department dropdown (flat), and the Site-of-Vaccination
// cascade (province -> sites logged under that province). Site->province
// resolution reuses extractProvinceFromVaccinationSite_ — a site typed
// before the "- PROVINCE" convention existed (all of today's real
// historical data) falls into the "Unspecified / Legacy" bucket rather
// than being silently dropped or mis-bucketed; see the note on that
// function for why province is never guessed from the whole string.
function getFilterOptionsV2() {
  assertCanViewReport_();
  try {
    const ss = getSheet_();
    const departments = new Set();
    const recipientsSheet = ss.getSheetByName('Recipients_Master');
    if (recipientsSheet) {
      const data = recipientsSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][10]) departments.add(data[i][10]); // Department / Unit
      }
    }

    const sitesByProvince = {}; // province -> Set of site strings
    const trackerSheet = ss.getSheetByName('Vaccination_Tracker');
    if (trackerSheet) {
      const data = trackerSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (row[TRACKER_COL.DEPARTMENT]) departments.add(row[TRACKER_COL.DEPARTMENT]); // safety net: a department that only ever shows up on a tracker row
        const site = row[TRACKER_COL.VACCINATION_LOCATION];
        if (!site) continue;
        const province = extractProvinceFromVaccinationSite_(site) || 'Unspecified / Legacy';
        // Per client instruction: the site dropdown shows/matches the
        // SITE NAME only (text before "-", e.g. "MTC Whiteplains"), not
        // the full raw "Site Name - PROVINCE" string.
        const siteName = extractSiteNameFromVaccinationSite_(site) || site;
        if (!sitesByProvince[province]) sitesByProvince[province] = new Set();
        sitesByProvince[province].add(siteName);
      }
    }
    const sitesByProvinceOut = {};
    Object.keys(sitesByProvince).forEach(function (p) {
      sitesByProvinceOut[p] = Array.from(sitesByProvince[p]).sort();
    });

    return {
      departments: Array.from(departments).sort(),
      provinces: Object.keys(sitesByProvinceOut).sort(),
      sitesByProvince: sitesByProvinceOut
    };
  } catch (e) {
    console.error('getFilterOptionsV2 error:', e);
    throw e;
  }
}

// The one live calculation engine behind every dashboard view —
// unfiltered, department-filtered, or site-filtered. mode/filterValue:
//   ('none', null)                 — everything
//   ('department', 'Mission Center') — one Department/Unit
//   ('site', 'MTC White Plains - QC') — one Site of Vaccination
function computeDashboardData_(mode, filterValue) {
  const ss = getSheet_();
  const trackerSheet = ss.getSheetByName('Vaccination_Tracker');
  if (!trackerSheet) throw new Error('Vaccination_Tracker sheet not found');
  const data = trackerSheet.getDataRange().getValues();

  // Recipients_Master lookup — used for the Active-recipient count, which
  // (for 'none'/'department') is independent of whether a recipient has
  // any Tracker rows yet. 'site' has no such independent meaning (a
  // recipient isn't "assigned" to an encoding site) — approximated below
  // as "active recipients with at least one record at that site", same
  // approach as the pre-existing site-filter logic this replaces.
  const recipientActive = {};
  const recipientDept = {};
  const recipientsSheet = ss.getSheetByName('Recipients_Master');
  if (recipientsSheet) {
    const rdata = recipientsSheet.getDataRange().getValues();
    for (let i = 1; i < rdata.length; i++) {
      const r = rdata[i];
      if (!r[0]) continue;
      recipientActive[r[0]] = isYes_(r[16]);
      recipientDept[r[0]] = r[10];
    }
  }

  function rowMatches(row) {
    if (mode === 'department') return row[TRACKER_COL.DEPARTMENT] === filterValue;
    // 'site' filterValue is now a SITE-NAME-ONLY value (text before the
    // "-" in "Site Name - PROVINCE", e.g. "MTC Whiteplains") per client
    // instruction — match every raw Vaccination_Tracker row whose site
    // name extracts to the same value, not by exact full-string equality
    // (a site can be logged as "MTC Whiteplains - QC" one day and "MTC
    // Whiteplains  -  QC" another and still be the same site).
    if (mode === 'site') {
      const loc = row[TRACKER_COL.VACCINATION_LOCATION];
      const name = extractSiteNameFromVaccinationSite_(loc) || loc;
      return name === filterValue;
    }
    return true; // 'none'
  }

  const today = new Date();
  const activeRecipientsSet = new Set();
  const vaccinatedSet = new Set();
  let dosesAdministered = 0, seriesCompleted = 0, actionRequiredCount = 0;
  let overdue = 0, due7 = 0, due30 = 0, due90 = 0, future = 0;
  const actionRows = [];
  const upcomingRows = [];

  if (mode === 'none' || mode === 'department') {
    Object.keys(recipientActive).forEach(function (id) {
      if (!recipientActive[id]) return;
      if (mode === 'department' && recipientDept[id] !== filterValue) return;
      activeRecipientsSet.add(id);
    });
  }

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!rowMatches(row)) continue;
    const disposition = row[TRACKER_COL.DISPOSITION];
    if (!disposition) continue;
    const recipientId = row[TRACKER_COL.RECIPIENT_ID];

    if (mode === 'site' && recipientActive[recipientId]) activeRecipientsSet.add(recipientId);

    if (disposition === 'Administered') {
      vaccinatedSet.add(recipientId);
      dosesAdministered++;
      if (COMPLETE_SERIES_STATUSES.indexOf(row[TRACKER_COL.SERIES_STATUS]) !== -1) seriesCompleted++;
      const dueDate = row[TRACKER_COL.RECOMMENDED_DATE];
      if (dueDate instanceof Date) {
        const daysUntil = Math.floor((dueDate.getTime() - today.getTime()) / 86400000);
        let dueStatus;
        if (daysUntil < 0) { overdue++; dueStatus = 'OVERDUE'; }
        else if (daysUntil <= 7) { due7++; dueStatus = 'DUE ≤7 DAYS'; }
        else if (daysUntil <= 30) { due30++; dueStatus = 'DUE 8–30 DAYS'; }
        else if (daysUntil <= 90) { due90++; dueStatus = 'DUE 31–90 DAYS'; }
        else { future++; dueStatus = 'FUTURE'; }
        upcomingRows.push({
          recipientId: recipientId,
          name: row[TRACKER_COL.RECIPIENT_NAME],
          department: row[TRACKER_COL.DEPARTMENT],
          site: row[TRACKER_COL.VACCINATION_LOCATION],
          vaccineType: row[TRACKER_COL.VACCINE_TYPE],
          nextDose: row[TRACKER_COL.NEXT_DOSE],
          recommendedDate: formatDateForClient_(dueDate),
          scheduledAppointment: formatDateForClient_(row[TRACKER_COL.SCHEDULED_APPOINTMENT]),
          daysUntilDue: daysUntil,
          dueStatus: dueStatus,
          scheduleStatus: row[TRACKER_COL.SCHEDULE_STATUS],
          reminderStatus: row[TRACKER_COL.REMINDER_STATUS]
        });
      }
    } else if (ACTION_DISPOSITIONS.indexOf(disposition) !== -1 || String(disposition).indexOf('Contraindicated') !== -1) {
      actionRequiredCount++;
      const reasonStatus = [row[TRACKER_COL.REASON_CATEGORY], row[TRACKER_COL.SPECIFIC_REASON]].filter(Boolean).join(' — ');
      actionRows.push({
        recipientId: recipientId,
        name: row[TRACKER_COL.RECIPIENT_NAME],
        department: row[TRACKER_COL.DEPARTMENT],
        site: row[TRACKER_COL.VACCINATION_LOCATION],
        vaccineType: row[TRACKER_COL.VACCINE_TYPE],
        intendedDose: row[TRACKER_COL.INTENDED_DOSE],
        disposition: disposition,
        reasonStatus: reasonStatus,
        reviewDueDate: formatDateForClient_(row[TRACKER_COL.REVIEW_DATE]),
        nextDose: row[TRACKER_COL.NEXT_DOSE],
        recommendedDate: formatDateForClient_(row[TRACKER_COL.RECOMMENDED_DATE]),
        scheduleStatus: row[TRACKER_COL.SCHEDULE_STATUS],
        requiredAction: inferRequiredAction_(disposition)
      });
    }
  }

  return {
    stats: {
      activeRecipients: activeRecipientsSet.size,
      vaccinatedRecipients: vaccinatedSet.size,
      dosesAdministered: dosesAdministered,
      nextDosesDue: overdue + due7 + due30 + due90 + future,
      seriesCompleted: seriesCompleted,
      actionRequired: actionRequiredCount,
      overdue: overdue, due7: due7, due30: due30, due90: due90, future: future
    },
    actionRequired: actionRows,
    upcomingDoses: upcomingRows
  };
}

// ------------------------------------------------------------
// REVERTED per client instruction: "the dashboard should rely in
// vaccination_report's formulas, the filters shouldn't since there's no
// existing formula for that in the sheets." So the DEFAULT/unfiltered
// view goes back to reading Vaccination_Report's own pre-built formula
// cells (these four functions), while Department/Site-of-Vaccination
// filtering — which the sheet has no formula for — stays on the
// live-computed computeDashboardData_ path above. This reintroduces the
// two-path split the comment above used to warn against, but that's the
// explicit tradeoff asked for: sheet-of-record for the default view,
// custom logic only where the sheet genuinely has nothing to read.
// ------------------------------------------------------------
function getDashboardStats() {
  assertCanViewReport_();
  try {
    const ss = getSheet_();
    const reportSheet = ss.getSheetByName('Vaccination_Report');
    if (!reportSheet) throw new Error('Vaccination_Report sheet not found');
    return {
      activeRecipients: reportSheet.getRange('A5').getValue(),
      vaccinatedRecipients: reportSheet.getRange('C5').getValue(),
      dosesAdministered: reportSheet.getRange('E5').getValue(),
      nextDosesDue: reportSheet.getRange('G5').getValue(),
      seriesCompleted: reportSheet.getRange('I5').getValue(),
      actionRequired: reportSheet.getRange('K5').getValue(),
      overdue: reportSheet.getRange('A9').getValue(),
      due7: reportSheet.getRange('C9').getValue(),
      due30: reportSheet.getRange('E9').getValue(),
      due90: reportSheet.getRange('G9').getValue(),
      future: reportSheet.getRange('I9').getValue()
    };
  } catch (e) {
    console.error('getDashboardStats error:', e);
    throw e;
  }
}

// ============================================================
//  HUB SYNC — pushes a stats snapshot to the Hub on a time-driven
//  trigger, per client instruction: "The Hub to VxSync sync is something
//  you'll build now, its not something that already exists."
//
//  Deliberately does NOT go through getDashboardStats() above:
//  that function calls assertCanViewReport_(), which checks
//  Session.getActiveUser() against Admin/Client roles — correct for a
//  real person clicking around the dashboard, wrong for a scheduled
//  trigger, which runs as the script owner with no "role" in this
//  system's sense at all. Reads the same Vaccination_Report cells
//  directly instead, with no role gate, since this is a system job, not
//  a user action — the Hub side call it lands on
//  (syncClientVxSyncData/reportSyncFailure) is what enforces trust there
//  (server-to-server, anon-key-gated, no session token — see
//  callHubApi_ above).
// ============================================================
function readDashboardStatsForSync_() {
  const ss = getSheet_();
  const reportSheet = ss.getSheetByName('Vaccination_Report');
  if (!reportSheet) throw new Error('Vaccination_Report sheet not found');
  const trackerSheet = ss.getSheetByName('Vaccination_Tracker');
  return {
    stats: {
      activeRecipients: reportSheet.getRange('A5').getValue(),
      vaccinatedRecipients: reportSheet.getRange('C5').getValue(),
      dosesAdministered: reportSheet.getRange('E5').getValue(),
      nextDosesDue: reportSheet.getRange('G5').getValue(),
      seriesCompleted: reportSheet.getRange('I5').getValue(),
      actionRequired: reportSheet.getRange('K5').getValue(),
      overdue: reportSheet.getRange('A9').getValue(),
      due7: reportSheet.getRange('C9').getValue(),
      due30: reportSheet.getRange('E9').getValue(),
      due90: reportSheet.getRange('G9').getValue(),
      future: reportSheet.getRange('I9').getValue()
    },
    // recordsSynced = total rows in Vaccination_Tracker (minus header) —
    // the simplest honest measure of "how much data did this push
    // represent," shown on the Hub's client card as "(N records)."
    recordsSynced: trackerSheet ? Math.max(0, trackerSheet.getLastRow() - 1) : 0
  };
}

// The actual push. Call this from a time-driven trigger (see
// installHubSyncTrigger below) or manually from the Apps Script editor
// (Run > syncStatsToHub) to test/force a sync on demand.
//
// NAMING NOTE: this function, installHubSyncTrigger, and
// uninstallHubSyncTrigger deliberately do NOT end in an underscore, even
// though the rest of this file uses a trailing underscore as its "private
// helper" convention (see checkHubAccess_, verifyHubToken_, etc.).
// Apps Script's editor hides any function ending in "_" from the Run
// dropdown AND from the "choose which function to run" picker in the
// Triggers UI — so a trailing underscore here would make these three
// functions genuinely un-clickable, not just stylistically "private."
// These three are meant to be run/selected by a person, so the
// underscore is dropped on purpose.
function syncStatsToHub() {
  try {
    const snapshot = readDashboardStatsForSync_();
    const result = callHubApi_('syncClientVxSyncData', {
      clientId: CONFIG.hubClientId,
      recordsSynced: snapshot.recordsSynced,
      stats: snapshot.stats
    });
    if (!result) {
      // callHubApi_ already logged the underlying network/parse error —
      // still tell the Hub explicitly that this attempt failed, so
      // "never heard from this client" and "heard from it, it failed"
      // aren't indistinguishable on the Hub's sync-status view.
      callHubApi_('reportSyncFailure', {
        clientId: CONFIG.hubClientId,
        errorMessage: 'Could not reach the Hub (network error or Hub returned success:false).'
      });
      Logger.log('syncStatsToHub: push failed, reported to Hub as a failure.');
      return { success: false };
    }
    Logger.log('syncStatsToHub: pushed ' + snapshot.recordsSynced + ' records successfully.');
    return { success: true };
  } catch (e) {
    // Reading the sheet itself threw (e.g. Vaccination_Report missing) —
    // still try to tell the Hub, best-effort, before giving up.
    try {
      callHubApi_('reportSyncFailure', {
        clientId: CONFIG.hubClientId,
        errorMessage: String(e && e.message || e)
      });
    } catch (err2) { /* nothing more we can do */ }
    Logger.log('syncStatsToHub error: ' + e);
    return { success: false, error: String(e && e.message || e) };
  }
}

// One-time setup, per client deployment: run this ONCE from the Apps
// Script editor (select installHubSyncTrigger in the function dropdown,
// click Run) to install the recurring push. Safe to re-run — it clears
// any previous copy of this trigger first, so running it twice doesn't
// create two competing triggers.
function installHubSyncTrigger() {
  uninstallHubSyncTrigger();
  ScriptApp.newTrigger('syncStatsToHub')
    .timeBased()
    .everyHours(6)
    .create();
  Logger.log('Hub sync trigger installed: syncStatsToHub will run every 6 hours.');
}

// Removes the trigger installed above, in case a client copy needs to
// stop pushing (e.g. being decommissioned) without hunting through
// Apps Script's Triggers UI by hand.
function uninstallHubSyncTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'syncStatsToHub') {
      ScriptApp.deleteTrigger(t);
    }
  });
}

function getActionRequired() {
  assertCanViewReport_();
  try {
    const ss = getSheet_();
    const reportSheet = ss.getSheetByName('Vaccination_Report');
    if (!reportSheet) throw new Error('Vaccination_Report sheet not found');
    // Header row 11, data starts row 12. Range is generous because this
    // section is a spilled array formula whose length grows with the data.
    const range = reportSheet.getRange('A12:L200');
    const data = range.getValues();
    const result = [];
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (row[0] && row[0] !== '') {
        result.push({
          recipientId: row[0],
          name: row[1],
          // No Department column exists in Vaccination_Report's own
          // formulas (confirmed against the workbook) — left blank in
          // this sheet-based default view. Only the live-computed
          // 'department' filter mode (and the always-on Department
          // Follow-Up Summary table) has a per-row Department value.
          department: '',
          site: row[2],
          vaccineType: row[3],
          intendedDose: row[4],
          disposition: row[5],
          reasonStatus: row[6],
          // These sheet cells can hold real Date objects. google.script.run
          // has a documented failure mode where a Date nested inside an
          // array of objects silently serializes to null on the client
          // instead of throwing — see formatDateForClient_.
          reviewDueDate: formatDateForClient_(row[7]),
          nextDose: row[8],
          recommendedDate: formatDateForClient_(row[9]),
          scheduleStatus: row[10],
          requiredAction: row[11]
        });
      }
    }
    return result;
  } catch (e) {
    console.error('getActionRequired error:', e);
    throw e;
  }
}

function getUpcomingDoses() {
  assertCanViewReport_();
  try {
    const ss = getSheet_();
    const reportSheet = ss.getSheetByName('Vaccination_Report');
    if (!reportSheet) throw new Error('Vaccination_Report sheet not found');
    // Header row 20, data starts row 21.
    const range = reportSheet.getRange('A21:K200');
    const data = range.getValues();
    const result = [];
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (row[0] && row[0] !== '') {
        result.push({
          recipientId: row[0],
          name: row[1],
          department: '', // see getActionRequired note above
          site: row[2],
          vaccineType: row[3],
          nextDose: row[4],
          recommendedDate: formatDateForClient_(row[5]),
          scheduledAppointment: formatDateForClient_(row[6]),
          daysUntilDue: row[7],
          dueStatus: row[8],
          scheduleStatus: row[9],
          reminderStatus: row[10]
        });
      }
    }
    return result;
  } catch (e) {
    console.error('getUpcomingDoses error:', e);
    throw e;
  }
}

function getSiteSummary() {
  assertCanViewReport_();
  try {
    const ss = getSheet_();
    const reportSheet = ss.getSheetByName('Vaccination_Report');
    if (!reportSheet) throw new Error('Vaccination_Report sheet not found');
    // Header row 46, data starts row 47. Column B is a blank spacer between
    // the site name (A) and the numeric columns (C:G) — do not read A:F as a
    // contiguous block, it silently drops the "Future Doses" column and
    // shifts every other value one column to the left.
    const range = reportSheet.getRange('A47:G100');
    const data = range.getValues();
    const result = [];
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (row[0] && row[0] !== '') {
        result.push({
          site: row[0],
          recipientsWithDue: row[2],
          totalDueDoses: row[3],
          dueWithin30: row[4],
          due31to90: row[5],
          futureDoses: row[6]
        });
      }
    }
    return result;
  } catch (e) {
    console.error('getSiteSummary error:', e);
    throw e;
  }
}

// Client-facing entry point for the top stats / Action Required /
// Upcoming Doses tables. mode: 'none' | 'department' | 'site'.
// 'none' reads Vaccination_Report's own formulas (sheet is the source of
// truth for the default view); 'department'/'site' use the live
// computation above, since the sheet has no formula for either grouping.
function getDashboardData(mode, filterValue) {
  assertCanViewReport_();
  try {
    const m = mode || 'none';
    if (m === 'none') {
      return {
        stats: getDashboardStats(),
        actionRequired: getActionRequired(),
        upcomingDoses: getUpcomingDoses()
      };
    }
    return computeDashboardData_(m, filterValue);
  } catch (e) {
    console.error('getDashboardData error:', e);
    throw e;
  }
}

// The always-visible, always-UNFILTERED "Department Follow-Up Summary"
// table — the department-based replacement for the old per-Assigned-Site
// summary. Same self-computed approach as everything else above.
function getDepartmentSummary() {
  assertCanViewReport_();
  try {
    const ss = getSheet_();
    const trackerSheet = ss.getSheetByName('Vaccination_Tracker');
    if (!trackerSheet) throw new Error('Vaccination_Tracker sheet not found');
    const data = trackerSheet.getDataRange().getValues();
    const today = new Date();
    const byDept = {}; // department -> { recipientsWithDue: Set, totalDueDoses, dueWithin30, due31to90, futureDoses }

    function bucket(dept) {
      if (!byDept[dept]) byDept[dept] = { recipientsWithDue: new Set(), totalDueDoses: 0, dueWithin30: 0, due31to90: 0, futureDoses: 0 };
      return byDept[dept];
    }

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[TRACKER_COL.DISPOSITION] !== 'Administered') continue;
      const dueDate = row[TRACKER_COL.RECOMMENDED_DATE];
      if (!(dueDate instanceof Date)) continue;
      const dept = row[TRACKER_COL.DEPARTMENT] || '(No Department)';
      const b = bucket(dept);
      const daysUntil = Math.floor((dueDate.getTime() - today.getTime()) / 86400000);
      b.recipientsWithDue.add(row[TRACKER_COL.RECIPIENT_ID]);
      b.totalDueDoses++;
      if (daysUntil <= 30) b.dueWithin30++;
      else if (daysUntil <= 90) b.due31to90++;
      else b.futureDoses++;
    }

    return Object.keys(byDept).sort().map(function (dept) {
      const b = byDept[dept];
      return {
        department: dept,
        recipientsWithDue: b.recipientsWithDue.size,
        totalDueDoses: b.totalDueDoses,
        dueWithin30: b.dueWithin30,
        due31to90: b.due31to90,
        futureDoses: b.futureDoses
      };
    });
  } catch (e) {
    console.error('getDepartmentSummary error:', e);
    throw e;
  }
}

// ============================================================
//  PDF EXPORT — data source for VxSyncDashboard.html's "Export PDF"
//  button. The PDF itself (charts, tables, layout) is built entirely
//  client-side (jsPDF + jspdf-autotable + Chart.js, loaded from CDN in
//  the dashboard partial) — this function's only job is to hand back
//  clean, privacy-filtered JSON. Apps Script has no good native charting
//  or PDF-layout story for something this detailed, so building the
//  document itself in the browser (same pattern the rest of this app
//  already leans on — the sheet/backend does data, the browser does
//  presentation) is more maintainable than fighting Slides/Docs export.
//
//  PRIVACY: per client instruction, contact info (email — Recipients_
//  Master's mobile number isn't even in this sheet, so it can't leak
//  here either) and clinical/reason detail (Reason Category, Specific
//  Reason — the closest thing this workbook has to a "health condition"
//  field, since deferral/contraindication reasons are often clinical)
//  are BOTH excluded from the recipient table and exceptions list by
//  default, and only included if the caller explicitly opts in via
//  payload.includeContactInfo / payload.includeClinicalDetails. This
//  function is the enforcement point — the checkboxes in the dashboard
//  UI are just how that choice gets made, not where the exclusion
//  actually happens.
function getPdfReportData(payload) {
  assertCanViewReport_();
  try {
    payload = payload || {};
    const includeContactInfo = !!payload.includeContactInfo;
    const includeClinicalDetails = !!payload.includeClinicalDetails;
    // Inclusive of the full end day (23:59:59), so a report "through
    // June 30" actually includes everything recorded ON June 30, not just
    // up to midnight at its start.
    const startDate = payload.startDate ? new Date(payload.startDate + 'T00:00:00') : null;
    const endDate = payload.endDate ? new Date(payload.endDate + 'T23:59:59') : null;

    const ss = getSheet_();
    const trackerSheet = ss.getSheetByName('Vaccination_Tracker');
    if (!trackerSheet) throw new Error('Vaccination_Tracker sheet not found');
    const data = trackerSheet.getDataRange().getValues();

    const EXCEPTION_DISPOSITIONS = ['Deferred', 'No-show', 'Contraindicated – Temporary', 'Contraindicated – Permanent'];
    const dispositionCounts = {};
    const vaccineTypeCounts = {};
    const siteCounts = {};
    const departmentCounts = {};
    const recipients = [];
    const exceptions = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[TRACKER_COL.RECORD_ID]) continue;

      // Scope by Vaccination Date when present; a Deferred/Declined/
      // No-show/Contraindicated record often never gets a vaccination
      // date at all (nothing was administered), so those fall back to
      // Status Date (column P, index 15) instead of being silently
      // dropped from a date-ranged report just because the "main" date
      // field is blank.
      const vaxDate = row[TRACKER_COL.VACCINATION_DATE];
      const statusDate = row[15];
      const effectiveDate = (vaxDate instanceof Date) ? vaxDate : (statusDate instanceof Date ? statusDate : null);
      if (startDate && (!effectiveDate || effectiveDate < startDate)) continue;
      if (endDate && (!effectiveDate || effectiveDate > endDate)) continue;

      const disposition = row[TRACKER_COL.DISPOSITION] || '(blank)';
      const vaccineType = row[TRACKER_COL.VACCINE_TYPE] || '(blank)';
      const site = row[TRACKER_COL.VACCINATION_LOCATION] || '(blank)';
      const department = row[TRACKER_COL.DEPARTMENT] || '(blank)';
      dispositionCounts[disposition] = (dispositionCounts[disposition] || 0) + 1;
      vaccineTypeCounts[vaccineType] = (vaccineTypeCounts[vaccineType] || 0) + 1;
      // Same population as vaccineTypeCounts/dispositionCounts above (every
      // in-range record, not just "Administered") — a deferred/no-show
      // record still belongs to a site/department, and dropping it would
      // make a site's exception rate invisible from this breakdown.
      siteCounts[site] = (siteCounts[site] || 0) + 1;
      departmentCounts[department] = (departmentCounts[department] || 0) + 1;

      const recipientRow = {
        recipientId: row[TRACKER_COL.RECIPIENT_ID],
        name: row[TRACKER_COL.RECIPIENT_NAME],
        category: row[5] || '',
        company: row[6] || '',
        department: row[TRACKER_COL.DEPARTMENT] || '',
        vaccineType: vaccineType,
        brand: row[TRACKER_COL.VACCINE_BRAND] || '',
        doseNumber: row[TRACKER_COL.DOSE_NUMBER] || '',
        disposition: disposition,
        vaccinationDate: formatDateForClient_(vaxDate),
        site: row[TRACKER_COL.VACCINATION_LOCATION] || '',
        seriesStatus: row[TRACKER_COL.SERIES_STATUS] || ''
      };
      if (includeContactInfo) {
        recipientRow.email = row[3] || '';
      }
      if (includeClinicalDetails) {
        recipientRow.reasonCategory = row[TRACKER_COL.REASON_CATEGORY] || '';
        recipientRow.specificReason = row[TRACKER_COL.SPECIFIC_REASON] || '';
      }
      recipients.push(recipientRow);

      if (EXCEPTION_DISPOSITIONS.indexOf(disposition) !== -1) {
        exceptions.push({
          recipientId: row[TRACKER_COL.RECIPIENT_ID],
          name: row[TRACKER_COL.RECIPIENT_NAME],
          vaccineType: vaccineType,
          disposition: disposition,
          reasonCategory: includeClinicalDetails ? (row[TRACKER_COL.REASON_CATEGORY] || '') : '',
          statusDate: formatDateForClient_(statusDate),
          reviewDate: formatDateForClient_(row[TRACKER_COL.REVIEW_DATE])
        });
      }
    }

    // Executive Summary reuses the SAME sheet-formula-driven snapshot the
    // live Program Dashboard shows (Vaccination_Report), rather than
    // recomputing "active recipients" / "series completed" scoped to the
    // date range — consistent with this file's existing "the sheet is the
    // source of truth for the default view" principle (see the DASHBOARD
    // API section). These figures are a CURRENT snapshot, not date-range
    // scoped — the report says so explicitly (see reportScopeNote) so
    // nobody reads "Active Recipients: 240" as being about the selected
    // date range when it isn't.
    let summaryStats = null;
    try { summaryStats = getDashboardStats(); } catch (e) { summaryStats = null; }

    return {
      success: true,
      metadata: {
        clientName: CONFIG.clientName,
        generatedBy: Session.getActiveUser().getEmail(),
        generatedAt: new Date().toISOString(),
        startDate: payload.startDate || null,
        endDate: payload.endDate || null,
        includeContactInfo: includeContactInfo,
        includeClinicalDetails: includeClinicalDetails,
        recordCount: recipients.length,
        reportScopeNote: 'Executive Summary figures reflect the CURRENT overall program status (same source as the live Program Dashboard). The Detailed Recipient Table, Visual Analytics, and Exceptions below are scoped to the selected date range.'
      },
      summaryStats: summaryStats,
      vaccineTypeCounts: vaccineTypeCounts,
      dispositionCounts: dispositionCounts,
      siteCounts: siteCounts,
      departmentCounts: departmentCounts,
      recipients: recipients,
      exceptions: exceptions
    };
  } catch (e) {
    console.error('getPdfReportData error:', e);
    throw e;
  }
}