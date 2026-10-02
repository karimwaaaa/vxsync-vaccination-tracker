# Internal Handoff Notes — VxSync / Hub

**Not part of the public portfolio README.** This file is for whoever picks up
active development next (a person, or another AI assistant being handed this
codebase cold). It assumes no memory of prior conversations — everything
you need to not repeat solved work or walk into known traps is below.

Last updated: 2026-10-02 (docs/index.html demo brought back in line with the real app: lot picker parity, Used Up behavior fix, Inventory sub-tab added, FAQ-render bug fixed; also lot search bar + expiry-date fix + PII leak fix in this repo from 2026-10-01).

---

## 1. What this system actually is

Two live, deployed pieces, plus one static demo:

- **Hub** (`pq-healthshield-hub` repo) — a Supabase-backed admin portal.
  Provisions a brand-new VxSync deployment per client (copies a master
  Google Sheet template, spins up a standalone Apps Script project, deploys
  it as a web app), tracks a "Client Vault" with per-client status
  (Ongoing / Active / Complete), issues SSO tokens, manages encoder
  accounts, help tickets, activity logs.
- **VxSync** (this repo) — the actual vaccination-tracking web app. One
  deployment per client. `src/Code.gs` is server-side Apps Script;
  `src/EntryForm.html` / `src/Dashboard.html` / `src/Index.html` are the
  client-side pages, rendered via `HtmlService`.
- **`docs/index.html`** — a static, self-contained GitHub Pages demo with
  fake baked-in data. It is NOT the real app and deliberately simplifies
  some things (see §5). Don't treat behavior differences between it and
  the real `src/` files as bugs.

**The Master Template Sheet itself is not in either repo** — new clients
still get their Sheet by a native Drive copy of a template sheet, not by
cloning this repo. But as of this session, the template's bound Apps
Script **source is now tracked for reference** under
`master-template-bound-script/` (see §7): `Code.gs` (the bound project's
own file, distinct from `src/Code.gs` above — both projects have a file
literally named `Code.gs`, they just can't share code at runtime),
`VxSync_Lot_Dropdown.gs`, `VxSync_Tracker_Backfill.gs`,
`VxSync_Format_Notes.gs`, `VxSync_Receive_Inventory.gs`,
`VxSync_Receive_Dialog.html`. **Tracking it here does not make it
deploy anywhere** — it's still a manually-pasted codebase on the actual
template Sheet's Apps Script editor, and a change made in this repo's
copy does nothing in production until someone pastes it into the real
bound project by hand. If you're told "the dropdowns aren't working on a
new client" and this repo's `src/Code.gs` looks fine, check whether the
*real* template's bound script (not just this repo's copy of it) actually
got the fix pasted in.

**Every automation runs off bare `onEdit(e)` / `onOpen(e)` simple
triggers**, never `ScriptApp.newTrigger()`. This is deliberate: simple
triggers need no per-copy authorization and survive a Drive copy
automatically, which is what makes zero-manual-step provisioning possible.
If you're asked to add a new automated behavior, wire it into the existing
shared `onEdit`/`onOpen`, don't install a new trigger type.

---

## 2. Format conventions that fail silently (read this before touching data)

These aren't validated by code — get them wrong and nothing errors, a
feature just quietly stops working:

- `Recipients_Master` → **Assigned Site/Location**: `"PROVINCE - descriptor"`
  (province FIRST, e.g. `"QC - MTC White Plains"`).
- `Lot_Expiry_Master` → **Site of Vaccination**: `"Site Name - PROVINCE"`
  (province LAST, e.g. `"MTC White Plains - QC"`). **This is the opposite
  order from the line above, on purpose (real, confirmed client data), and
  mixing them up is the single most common silent-failure bug in this
  system** — a recipient whose Assigned Site/Location uses the wrong order
  will show "0 recipients" for every province in the web app, with no error
  anywhere. This exact bug happened with real test data this session.
- `Active` columns (`Recipients_Master`, `Vaccine_Schedule_Master`,
  `Vaccinators_Master`): must be exactly `"Yes"`. Anything else — blank,
  `"Active"`, `"TRUE"` — is silently treated as inactive.
- `Vaccine_Schedule_Master` → **Interval Unit**: exactly `Days` / `Weeks` /
  `Months` / `Years`, case-sensitive.
- `Disposition_Master` → **Dose Disposition**: the two `Contraindicated`
  values use an EN DASH (`–`), not a hyphen (`-`). They look almost
  identical.
- `Lot_Expiry_Master` → **Session ID / Order Ref** (col A): system-stamped
  only. Typing anything here manually makes that lot look already-claimed.

All seven of these have an actual cell note on the header explaining the
format (see `VxSync_Format_Notes.gs`, template-side, manually pasted per
client — not in this repo since it's part of the untracked template
script).

`getLastRow()` vs `getMaxRows()`: **always use `getMaxRows()`** when sizing
a refresh over `Vaccination_Tracker`. A brand-new client's Tracker is
completely empty (`getLastRow()` returns 1), so anything sized off
`getLastRow()` silently no-ops on every edit until something else (manual
setup, or `onOpen`) happens to run a full refresh. This was a real, shipped
bug this session (`vxSyncTrackerSourceEdit` in the template's
`VxSync_Tracker_Backfill.gs`) that took three separate "dropdown missing on
a new client" reports to actually diagnose — the fix was one word.

---

## 3. What changed this session (2026-09)

- Vaccination_Tracker's Recipient ID column now auto-widens (140px) on
  every dropdown refresh.
- Intended Dose Number is now a real dropdown (was silently free-typed).
- The `getLastRow()`/`getMaxRows()` bug above — fixed.
- `EntryForm.html`: Recipient Name dropdown no longer appends a
  `(province)` suffix.
- `EntryForm.html`: the site-typo/autocorrect notice is now dismissible
  (✕ button), and a new, independent **unknown-province soft warning**
  fires when the parsed province isn't in `PROVINCE_ALIASES` — never a hard
  block, since a genuinely new province is valid.
- Help (❓) FAQ panel: 4 new entries added, matching the above.
- `docs/index.html` demo: removed its own (separate, fake-data) recipient
  dropdown parenthetical; added an **illustrative** lot-inventory card
  picker (Ready/Hold/Used Up, "Your site" badge) — see §5, this is NOT a
  full port of the real Live Lot Inventory feature.
- Full patch notes with a testing-discipline disclaimer:
  `VxSync_Patch_Notes.md` (delivered to the user separately, not in-repo).

**Testing discipline**: the user's practice sandbox is a sheet called
"Practice Client." Real clients (including one referred to as "LDS") and
the Master Template must never be used for testing changes.

---

## 3a. What changed this session (2026-09-30, bound-script side)

This is a second, later round of changes in the same session, entirely on
the **bound-script / Master Template side** — none of it touches `src/`.
**None of this has been live-tested yet.** The user is planning a real
dry run on "Practice Client" (scan a real barcode, try the AI fallback,
submit a 5–10 unit batch, confirm the rows land correctly in
`Lot_Expiry_Master`) — treat everything below as "built and syntax-checked,
not yet proven in the Sheet" until that happens.

- **Backfill `notFound` silent-skip rule.** A Backfill-imported
  `Vaccination_Tracker` row whose lot isn't in `Lot_Expiry_Master` at all is
  now treated as an inert historical record — no claim attempt, no error
  note, no toast — instead of being flagged the same way a genuine
  mismatch (wrong region, already Used Up) is. Changed in
  `master-template-bound-script/VxSync_Tracker_Backfill.gs`
  (`trkClaimLotInventory_`) and `master-template-bound-script/Code.gs`
  (`claimLotInventoryForBoundSheet_`, the `notFound: true` flag). **Known,
  accepted tradeoff**: a genuine typo in a Backfill row's lot number is now
  indistinguishable from an intentional pre-VxSync historical record — both
  silently skip. See the extended comments at both call sites, and the
  `noLotRowsButAdministered` flag in `src/Code.gs`'s
  `getInventoryBalanceReport` for the one place this can still be surfaced
  after the fact (inventory math looking off is the symptom to watch for,
  not an error anywhere).

- **New "Receive Inventory" feature — Sheets-only, not the Dashboard.**
  Added a `VxSync` custom menu (`VxSync_Lot_Dropdown.gs`'s existing
  `onOpen(e)` — deliberately not a second `onOpen`, see that file's
  comments) with one item, **VxSync ▸ 📷 Receive Inventory…**, opening a
  modal dialog (new files: `VxSync_Receive_Inventory.gs` +
  `VxSync_Receive_Dialog.html`). Lets someone receiving new vaccine stock
  build up a batch (Site, Vaccine Type, Brand, Lot Number, Expiry Date,
  Quantity per line) and submit it as one locked, batched write — still
  **one row per physical claimable unit** in `Lot_Expiry_Master`, same data
  model as always, nothing about that changed.
  - Two independent ways to fill a line, in order of preference, both
    optional (typing by hand always works): **(1) barcode scan from a
    photo** — decodes a GS1 barcode client-side (ZXing-js off a static
    image, not a live camera feed, to sidestep `getUserMedia` inside an
    Apps Script `HtmlService` iframe) and parses AI 01/17/10 (GTIN/expiry/
    lot). **(2) Gemini AI photo fallback** — a visually separate second
    button, never auto-triggered after a failed scan, for when the barcode
    itself won't scan (damaged box, smudged label, no barcode printed).
    Reads the label's printed text via `gemini-2.5-flash`. The dialog has
    inline help text above both buttons explaining this decision order.
  - **This was deliberately built Sheets-only, not in the Dashboard.** An
    earlier version of this same feature was built into `src/Dashboard.html`
    and `src/Code.gs` and then **fully removed** once the user decided the
    Sheet was the better fit — if you ever see a stray reference to a
    Dashboard "Receive Inventory" tab in an old export or chat history,
    that's dead, superseded direction, not a rollback target.
  - **One manual one-time setup step, per Master Template (not per
    client)**: a `GEMINI_API_KEY` Script Property must be set on the bound
    project (`Project Settings ▸ Script Properties` in the Apps Script
    editor) for the AI fallback button to work. The barcode-scan button and
    manual typing both work with zero setup. The key is never read
    client-side and the uploaded photo is never persisted — resized to a
    max 1280px/JPEG q0.82, sent once, discarded when the function returns.
  - **Known real limitation, not a bug**: custom Apps Script menus (and
    therefore this whole feature) **do not render in the native Sheets
    mobile app** on phones — a genuine, longstanding Google Apps Script
    platform limitation, not something specific to this build. Confirmed
    against Google's own docs/samples. Workaround if a client insists on
    phone use: open Sheets in a mobile **browser** with "Request Desktop
    Site" enabled, which loads the full editor including custom menus. For
    a desk-bound PC/laptop workflow (the realistic case here), a USB/
    Bluetooth keyboard-wedge barcode scanner was recommended as a better
    long-term fit than photo-based scanning — **not built**, just scoped
    and offered; would plug into the same client-side `parseGS1()` parser
    if built later.

---

## 3b. What changed this session (2026-10-01, standalone web app side)

Back on `src/Code.gs` / `src/EntryForm.html` — the live lot-inventory picker
in the encoding form.

- **Search bar added above the lot card list** (`EntryForm.html`). Plain
  client-side filter against the already-loaded cards, no extra server
  round trip (`efLotSearchTerm` + a filter inside `efRenderLotCards`).
  Only shows once there's more than one lot to search through; resets on
  every fresh Vaccine Type/Brand fetch so a leftover search term can't
  silently hide a brand-new list.
- **Real fix for "Expiry Date doesn't fill in when I pick a lot."** Root
  cause: `formatDateForClient_` (`Code.gs`) only reformatted a cell's date
  when the sheet held a real JS `Date` object. A Lot_Expiry_Master Expiry
  Date cell typed by hand as text — e.g. `29/08/2027` — was passed through
  completely unconverted. `<input type="date">` in every browser silently
  refuses to populate from anything that isn't exactly `yyyy-mm-dd` — no
  error, nothing in the console — so some lots filled the date and others
  didn't, with zero visible reason why. Fixed by having
  `formatDateForClient_` also recognize `d/m/yyyy`/`dd/mm/yyyy` text and
  convert it to ISO before it ever reaches the client. **Flagged to the
  client as a follow-up, not yet confirmed**: worth checking whether those
  cells are stored as text or as real dates with a misleading display
  format — if other code elsewhere ever reads that same column a
  different way, it could still disagree with what this fix now shows.

### ⚠️ Real finding this session, not hypothetical: a prior sync of this repo leaked real client PII

When re-syncing `src/Code.gs` for the two fixes above, a check that should
have been run the FIRST time this file was wholesale-copied into the
portfolio repo (earlier in the 2026-09-30 session, see §3a) was not done
carefully enough: the copy-in only checked for the generic
`REPLACE_WITH`-style placeholder pattern, not for real embedded values.
**`CONFIG.nurseEmails` / `CONFIG.adminEmails` / `CONFIG.encoderSiteMap` in
that copy held the real client's actual email addresses and real staff
names** (a `mypqh.com` domain, a personal Gmail, a `.edu.ph` address, real
first/last names) — not placeholders. `hubApiUrl` / `hubAnonKey` /
`hubLoginUrl` / `sheetId` were already correctly placeholdered; only the
three email-list fields and the site map leaked.

**This means the `vxsync-repo.zip` delivered earlier in this same session,
before this fix, contained real client PII.** If that zip (or its contents)
has been pushed anywhere public — GitHub, a portfolio site, shared with
anyone outside this engagement — those real emails and names need to be
scrubbed from that history too; replacing the file in a fresh zip does not
retroactively fix anything already published. **This has been told to the
user directly — don't treat it as quietly resolved just because this repo
copy is now clean.**

Fixed in this repo copy only (the real, deployed `Code.gs` is correct as-is
— it's SUPPOSED to have the real client's data, it's just never supposed to
leave this engagement): `nurseEmails`/`adminEmails`/`clientEmails` replaced
with generic `@example.com` placeholders, `encoderSiteMap` replaced with a
matching generic example. **Lesson for next time, written down so it
doesn't get re-learned the hard way**: "no `REPLACE_WITH` pattern found" is
NOT the same check as "no real embedded value found" — before any future
wholesale copy of a real client file into this repo, actually grep for
email-address patterns (`@[a-z0-9.-]+\.[a-z]{2,}`) and real-looking proper
nouns, not just the placeholder convention.

---

## 3c. What changed this session (2026-10-02, docs/index.html demo)

The user flagged that the demo no longer matched the real app and asked for
it to be brought back in line. Two separate things were wrong, of
different severity:

- **Lot picker visually stale.** The real `EntryForm.html` picked up a
  two-column card grid with a search box above it (see §3b), plus
  relabeled badges ("✓ Ready" / "⚠ N on hold" / "Your site" / "Same
  region"). The demo's `vx-lot-grid` was still the older single-column
  flex-wrap layout with no search box and the old badge text. Brought it
  back in step: same grid CSS, `#lotSearchBox` added above
  `#lotPickerGrid`, `efRenderLotPicker()` split into a reset variant
  (brand/site actually changed — clears selection + search) and a plain
  redraw variant (search box typed into — filters only, never touches the
  current selection), same split the real app makes.
- **The demo was also quietly showing a state the real app never shows at
  all**, independent of anything changed this session: a lot the real
  system marks "Used Up" is never sent to the client (`isLotSelectable_`
  in `Code.gs` filters it server-side) — it just isn't in the list. The
  demo instead rendered it as a greyed-out, disabled card. Fixed by
  filtering "Used Up" `DEMO_LOTS` entries out of the render entirely, so a
  claimed lot now disappears from the picker the same way it does for
  real, not greys out. Updated the FAQ entry that had been describing the
  wrong (greyed-out) behavior.
- **The whole Inventory sub-tab was missing.** `Dashboard.html` has had a
  second sub-tab — "💉 Inventory" / the Lot Inventory Balance Report
  (Site filter, Received/Administered/Remaining/Hold/Flags table) —
  for a while, and the demo never got an equivalent at all; its "Program
  Dashboard" was Program Monitoring only. Added a matching `vx-tabs` /
  `vx-view` sub-tab split (mirroring `Dashboard.html`'s own
  `vxSwitchTab`), a `DEMO_INVENTORY` illustrative dataset (separate from
  `DEMO_LOTS` — same distinction the real app makes between
  `getInventoryLots()` for the entry form and `getInventoryBalanceReport()`
  for the dashboard), a site filter, and a flagged-row example (including
  one deliberately showing a negative Remaining, same red-text treatment
  as the real report, to make the "Flags" column legible at a glance
  rather than just theoretical).
- **Found and fixed a real, pre-existing, unrelated bug while verifying
  all this in a headless browser** (not something introduced this
  session): `vxRenderFaqs()` was called immediately at top-level script
  execution, but `#vxFaqList` lives later in the document than the
  `<script>` block — so `document.getElementById('vxFaqList')` returned
  `null` on every single page load, threw, and because it's an uncaught
  error in top-level synchronous script execution, silently aborted every
  statement after it in that same `<script>` block. In practice: the help
  panel's FAQ list was **always empty** on this demo, and the whole
  help-ticket submission flow defined after that line never properly
  initialized (`vxHelpTicketType` stayed `undefined` until a user actually
  clicked a type toggle). Fixed by deferring the call to
  `DOMContentLoaded` instead of calling it inline. Verified via a headless
  Playwright pass clicking through every tab, the lot picker, the search
  box, both dashboard sub-tabs, and opening the help panel — zero page
  errors, 10 FAQ entries actually render now.

---

## 4. Known bugs / gotchas already found and fixed — don't re-diagnose these

- "Recipient ID dropdown missing" on a brand-new client → §2/§3,
  `getLastRow()` bug. Fixed. If it recurs, check whether the template's
  *bound* script actually got the fix pasted in (see §1) before assuming a
  new bug.
- "0 recipients" showing for every province in the web app → almost always
  the Assigned Site/Location province-order bug in §2, or a stale
  client-side cache (`efLoadRecipients()` in `EntryForm.html` populates
  `efAllRecipients` once at page load and does not auto-refresh on a sheet
  edit — tell the user to hard-reload the tab before assuming a data bug).
- Test kit files (`VxSync_Tracker_Backfill_Test_Kit.xlsx`,
  `VxSync_Lot_Inventory_Test_Kit.xlsx`) exist for regression testing; not
  committed to this repo, delivered directly to the user.

---

## 5. What the `docs/index.html` demo does and doesn't cover

It's a static, self-contained mockup with fake baked-in data — no Sheets/
Apps Script/Hub backend. Keep these gaps in mind before assuming a demo
change ports 1:1 to the real app or vice versa:

- **Site of Vaccination is a dropdown in the demo**, with an on-page tip
  explaining the real app uses free-text + typo correction. The demo
  intentionally has no equivalent of the dismissible typo notice or the
  unknown-province warning — don't add one without also removing that
  explanatory tip, or the two will contradict each other.
- **The lot-inventory picker added this session is illustrative only** —
  about 10 hand-written `DEMO_LOTS` entries, not a simulation of a real
  `Lot_Expiry_Master` sheet with true province/region logic (the demo's
  site list has no separate province field to base that on honestly).
  Good enough to show *how the feature behaves* to someone reviewing the
  portfolio; not a spec to copy for real inventory logic.

---

## 6. Open / pending decisions — NOT resolved, don't guess on these

### 6a. Booster dose recurrence (Vaccination_Tracker backfill)

A backfilled record whose last recorded Dose Number is `"Booster"`
currently computes nothing (Next Dose / Recommended Date / Series Status
all stay blank) — `trkDeriveCurrentDoseNumber_` / `trkComputeFollowUp_`
don't know how to mechanically advance from a non-numeric dose label. This
may be correct (Booster may genuinely have no fixed next step) or it may
need to be treated as recurring the same way Annual/Seasonal Dose already
is, if a real `Vaccine_Schedule_Master` recurring schedule exists for it.
**Ask the client before changing this** — don't assume either direction.

### 6b. New feature idea: per-client "time to completion" for billing/scheduling

The user is considering pricing/scheduling an engagement based on how long
it takes a client's employees to finish most/all of their dose series.
**Nothing for this exists yet** — this is a design discussion only, not a
built feature. Key findings from investigating feasibility, so whoever
builds this doesn't have to re-derive them:

- **Per-recipient completion data already exists and is reliable.** Each
  `Vaccination_Tracker` row computes a real `Series Status` (`In Progress`
  / `Complete` / `Complete - Recurring`, see `COMPLETE_SERIES_STATUSES` in
  `Code.gs`), and the row a recipient's series completes on carries a real
  date. This part needs no new engineering.
- **There is no "program start date" anywhere.** Not in this repo, not in
  the Hub's `clients` table (which has no `created_at` column at all — see
  `hub-repo/backend/schema.sql`), not anywhere durable.
- **The user's proposed fix — use the Client Vault's Ongoing → Active
  status transition as the clock start — is sound and preferred over a
  separate manually-entered date field.** Ongoing is provisioning/data-entry
  time and shouldn't count against completion speed; Active is the
  deliberate admin judgment call that encoders are actually live. This is
  a Hub-side change: `hub-repo/backend/api_index.ts`'s `updateClientStatus`
  handler needs to write a dedicated timestamp column (e.g.
  `active_started_at`) on the `clients` table at the moment status becomes
  `Active` — a small migration + one extra write, not a big lift.
- **Do NOT source that timestamp from `audit_log`, even though it already
  logs the status change with a real timestamp.** `hub-repo/backend/
  migrations/migration_log_retention.sql` schedules a nightly cron job that
  prunes `audit_log` rows older than 180 days. A client rollout that runs
  longer than 6 months would silently lose its own start date. This is a
  real trap, not a hypothetical — check for it if anyone proposes building
  this off the existing log instead of a new column.
- **Do NOT use the timestamp of the status flipping to `Complete` as the
  completion end-date either**, if `Complete` stays a manual admin action
  (it currently is, same `updateClientStatus` endpoint, no automatic
  trigger). That measures admin reaction time, not actual clinical
  completion speed. Prefer computing the real end date directly from each
  recipient's own last-`Administered` / Series-Status-Complete row in
  `Vaccination_Tracker` — that data is already exact and already exists.
  `Complete` status can stay a separate, decoupled business/contractual
  signal.
- **Also currently missing**: the Hub's `client_vxsync_sync` table (added
  in `migration_activity_sync_ticketsource.sql`) only tracks sync
  freshness/success — it does NOT persist the actual stats payload a
  client's VxSync pushes on each sync (`syncClientVxSyncData` in
  `api_index.ts` receives `stats` and echoes it back in the response but
  never writes it to a column). There is currently no historical trend of
  completion % anywhere, Hub or Sheet side — only whatever the live
  dashboard can compute right now, for whichever client's Sheet you happen
  to open.
- **Still an open question, not yet decided by the client**: if a client's
  status goes `Active → Ongoing → Active` again (paused rollout, mistake,
  stock-out), should the clock reset, or should `active_started_at` be
  "first time ever" and stay fixed? This changes the schema (a single
  timestamp column vs. something that can be recomputed) — get an answer
  before implementing either way.

---

## 7. Repo/file map quick reference

```
vxsync-repo/
  src/Code.gs           — standalone web app, server-side logic (real, current)
  src/EntryForm.html     — encoding form (real, current)
  src/Dashboard.html     — ops dashboard (real, current — no Receive
                           Inventory tab; that feature lives in the Sheet
                           now, see §3a, not here)
  src/Index.html         — app shell/router (unchanged this session)
  master-template-bound-script/  — reference copy of the Master Template
                           Sheet's bound Apps Script project (see §1 — NOT
                           what deploys; the real one is pasted by hand):
    Code.gs                        — bound project's own Code.gs (distinct
                                      file from src/Code.gs above)
    VxSync_Lot_Dropdown.gs         — dropdown automation + the one onOpen
                                      that also builds the VxSync menu
    VxSync_Tracker_Backfill.gs     — Backfill claim logic incl. notFound
                                      silent-skip (§3a)
    VxSync_Format_Notes.gs         — inline header-cell format guidance
    VxSync_Receive_Inventory.gs    — Receive Inventory feature, server side
    VxSync_Receive_Dialog.html     — Receive Inventory modal dialog, UI
  docs/index.html        — static demo (see §5)
  README.md              — public portfolio README, keep clean/polished
  HANDOFF.md             — this file

hub-repo/
  backend/api_index.ts       — Supabase Edge Function, all backend logic
  backend/schema.sql          — base schema
  backend/migrations/*.sql    — chronological migrations (READ THESE before
                                 assuming what columns/tables exist — several
                                 things referenced in §6b, like
                                 client_vxsync_sync, only exist because of a
                                 migration, not the base schema.sql)
  frontend/hub.html           — entire frontend SPA
  docs/index.html             — static demo
  README.md                   — public portfolio README, keep clean/polished
```

Not tracked in either repo: any real client's live Sheet data, and the
actual deployed state of the Master Template Sheet's bound Apps Script
project — only a reference copy of its source lives here now (§1, §7),
kept in sync by hand.
