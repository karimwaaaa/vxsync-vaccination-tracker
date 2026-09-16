# VxSync — Vaccination Tracking System

*Built while working at PQ Healthshield Inc.*

A per-client vaccination program tracker I built on Google Apps Script +
Google Sheets, deployed as its own web app for each client a healthcare
logistics company runs an on-site vaccination program for. Originally a
single hand-maintained deployment, later rebuilt so a fresh, fully working
copy could be spun up for a brand-new client in minutes instead of days (see
the companion [Hub](https://github.com/karimwaaaa/pq-healthshield-hub)
project, which automates that provisioning).

> This repo is a sanitized copy of a real internal system. Real names,
> emails, company identifiers, and Sheet/API IDs have been replaced with
> placeholders (`[COMPANY_NAME]`, `[COMPANY_DOMAIN]`, `YOUR-PROJECT-REF`,
> example emails, etc.). It won't run as-is without your own Google Sheet
> set up with the expected tabs/columns and your own Hub backend to talk to.

## Live Demo

**[https://karimwaaaa.github.io/vxsync-vaccination-tracker/](https://karimwaaaa.github.io/vxsync-vaccination-tracker/)**

A fully static, self-contained mockup (`docs/index.html`) that recreates the
Entry Form and Program Dashboard UI with fake sample data baked in — no
Google Sheets/Apps Script backend or auth required, since none of that runs
on GitHub Pages. Enable via Settings → Pages → Source: Deploy from branch →
main → /docs.

## What it does

- **Bulk + one-at-a-time recipient registration.** Admins can paste in a
  client-provided recipient list directly into the underlying Sheet ahead of
  a program launch, or register someone on the spot through the Entry Form
  if they weren't on the original list. Either path gets the same
  auto-generated, never-reused `VAC-######` recipient ID.
- **Guided vaccination entry form** (`EntryForm.html`) — cascading
  dropdowns (vaccine type → brand → schedule → dose number) driven entirely
  by data already in the Sheet, so adding a new vaccine or schedule to the
  "master" tabs is enough to make it available in the form with zero code
  changes.
- **Typo-tolerant site matching, checked word by word.** Vaccination sites
  were historically typed by hand by different people ("MTC Whiteplains" /
  "MTC WHITEPLAINS" / "MTC WP" / ...) long before this interface existed.
  `Code.gs` includes a three-tier autocorrect layer (fixed corrections,
  learned-from-history matches, and soft suggestions) that checks each
  *word* independently rather than comparing the whole typed phrase against
  whole known values — an earlier whole-string version could match a typo'd
  site against a completely unrelated real site just because the two
  strings happened to share overall length/structure. Word-level matching
  means a correctly-typed word is never swept into a suggestion because
  some other word nearby was misspelled, and a suggestion only ever names
  the specific word that looks wrong.
- **Live operations dashboard** (`Dashboard.html`) — active/vaccinated
  recipient counts, doses administered, upcoming/overdue dose follow-ups
  bucketed by urgency, and a PDF export of the whole report.
- **Role-based access without its own login system** — a vaccinator's own
  email (pulled live from the `Vaccinators_Master` tab) grants them Nurse/
  Encoder access automatically; no separate account provisioning step
  needed when staff changes.
- **Single sign-on from the Hub.** This app doesn't have its own login
  screen — it verifies a signed token handed to it by the Hub and looks up
  the visiting user's role from there, so access control lives in one place
  instead of being duplicated per client.

## Architecture

```
Hub (SSO token) ──▶ Index.html ──▶ EntryForm.html / Dashboard.html
                        │
                        ▼
                 Code.gs (server-side Apps Script)
                        │
                        ▼
        Bound Google Sheet: Recipients_Master, Vaccination_Tracker,
        Vaccinators_Master, Vaccine_Schedule_Master, Disposition_Master
```

- **`src/Code.gs`** — all server-side logic: role/session verification
  against the Hub, recipient/vaccination CRUD, the site-name fuzzy matcher,
  dashboard stat aggregation, and PDF report generation.
- **`src/Index.html`** — the app's shell/router (decides whether to show
  the Entry Form or Dashboard based on the visiting user's role).
- **`src/EntryForm.html`** — the vaccination encoding form.
- **`src/Dashboard.html`** — the operations/reporting dashboard.
- **`src/favicons/`** — app icon assets.

## Notable engineering decisions

- **Every async lookup chain is generation-guarded.** Picking a recipient,
  vaccine type, or brand kicks off a cascade of `google.script.run` round
  trips (dose history -> brand list -> schedule list -> starting-dose
  lookup). Nothing stopped an older, slower response from landing *after*
  a newer one and silently overwriting the form with stale data — e.g. a
  user picks vaccine type A, changes their mind to type B before A's
  lookup returns, and A's late response quietly clobbers what's now
  showing for B, so what gets submitted no longer matches what's on
  screen. `EntryForm.html` now stamps every lookup with a generation
  counter and every callback checks it's still current before touching
  the DOM, so a superseded response is simply dropped instead of applied.
- **Recipient IDs are never reused, on purpose.** A deleted recipient's old
  `VAC-######` ID is retired permanently rather than being handed to the
  next new registration — reusing it risks silently reattaching an old
  person's real vaccination history to a different new person, since
  historical records reference the ID, not the row.
- **Auto-ID generation supports two very different real workflows.**
  Vaccination IDs are assigned automatically the moment a row in
  `Recipients_Master` has both a last and first name — whether that row
  was typed one at a time through the Entry Form, or pasted in bulk ahead
  of a program launch from a client-provided list. Both paths converge on
  the same ID-assignment logic instead of needing two.
- **Fuzzy site matching over a fixed dropdown.** Locking site names to a
  hard-coded list would have meant a code change every time a client added
  a physical site. Instead, `Code.gs` normalizes free-text input against
  known aliases and flags likely-new entries, so the list of valid sites
  grows from real usage instead of a deployment.
- **No separate login system.** Piggy-backing on the Hub's session token
  instead of building VxSync's own auth avoided maintaining two sources of
  truth for "who is this person and what can they do" across every client
  deployment.

## Tech stack

Google Apps Script (server-side JS on V8 runtime) · Google Sheets (as the
data store) · vanilla HTML/CSS/JS (client-side, template-rendered via
Apps Script's `HtmlService`) · Google Drive API + Apps Script API (for
automated provisioning — see the Hub repo)

## Part of a 3-repo system

This was one of three connected projects I built for the same employer, each
kept as its own repo here since they're independently useful and readable
on their own:

- **VxSync** (this repo) — the per-client vaccination tracker.
- **[Hub](https://github.com/karimwaaaa/pq-healthshield-hub)** — the admin
  portal that provisions a new VxSync deployment per client and issues the
  SSO token this repo verifies.
- **[Ops Dashboard](https://github.com/karimwaaaa/central-ops-dashboard)**
  — a separate AR/order-management/inventory system for the same company,
  authenticated the same way.

They share one sign-on flow, which is the main thing that ties them
together architecturally.

## Setup (if you want to run your own copy)

1. Create a Google Sheet with tabs named `Recipients_Master`,
   `Vaccination_Tracker`, `Vaccinators_Master`, `Vaccine_Schedule_Master`,
   `Disposition_Master`, and `Vaccination_Report` — see the column names
   referenced throughout `Code.gs` (search for `.getRange(` calls) for the
   exact expected columns.
2. Create a new Apps Script project, either bound to that Sheet directly,
   or standalone with `CONFIG.sheetId` in `Code.gs` set to the Sheet's ID
   (the standalone route is what the Hub's auto-provisioning uses, since a
   bound script's ID can't be discovered reliably after a Drive copy).
3. Fill in `CONFIG` at the top of `Code.gs` — client name/branding,
   authorized emails, and (if integrating with a Hub-style SSO backend)
   `hubApiUrl` / `hubAnonKey` / `hubClientId` / `hubLoginUrl`.
4. Deploy as a web app (Execute as: the deploying account or the accessing
   user, depending on whether you want a shared or per-user Google
   identity; Who has access: your choice).
