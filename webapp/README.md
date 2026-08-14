# Xplor Data Migration Tools — Web

A browser-based rewrite of the Xplor Data Migration Tools, built for deployment on
Vercel. This replaces the original Streamlit/Python app (kept in the repo root for
reference).

**All 5 tools are ported and live:** Parent and Child Import, Bookings Import,
Payment Plans Import, Balance Adjustments, and Room & Fee Names checking. Each is
a 1:1 behavioural port of its Python source (`validator_v2.py`,
`prepare_bookings_import.py`, `payment_plan_checker.py`,
`process_balance_adjustments.py`, `check_names.py`), verified against synthetic
fixtures — see Verification below.

## Why this exists

The original app processed real childcare-enrolment PII (children's medical data,
parents' contact details) entirely server-side, with no access control and
indefinite server-side caching of uploaded files. See the project memory /
conversation history for the full audit. This rewrite's core design decision:

**All parsing and validation runs client-side, in the browser.** Uploaded files
are never sent to a server — there is no API route, no server action, and no
server-side code path that ever sees file contents. The app ships as static
assets; `next build` prerenders every route to static HTML/JS (confirmed — see
Verification below).

## Data-handling guarantees

- **No network calls carrying file data.** Parsing (PapaParse/ExcelJS), every
  tool's processing pipeline, and report generation (ExcelJS, JSZip) all run in
  the browser's JS runtime. Verified with Playwright on every tool: zero
  non-localhost network requests fire during upload → run → download. (Balance
  Adjustments does one same-origin `fetch()` for its bundled `.xlsx` template —
  a static app asset, not user data, so this doesn't change the guarantee.)
- **No logging of row/file content.** No `console.log` of parsed data anywhere in
  `src/lib/`.
- **No persistence of PII.** Uploaded files and results live only in React
  component state — gone on refresh. The only thing that persists (in
  `sessionStorage`, cleared when the tab closes) is the service ID mapping —
  centre names and IDs, not enrolment PII — mirroring the original app's
  "remember the service map across tools" convenience. It can be entered manually
  (up to 5 services) or uploaded as a CSV.
- **No third-party analytics or telemetry** in the shipped app.
- **No secrets required.** None of these tools make any API calls at all.
  `.env.example` is a placeholder in case a future tool needs one;
  root `.gitignore` already blocks `.env*`.

## Getting started

```bash
npm install
npm run dev
```

Open [http://localhost:3000](http://localhost:3000).

## Verification scripts

Playwright- and Node-driven scripts were used during development to exercise each
tool end-to-end against synthetic data (never real client data) — logic checks run
against hand-computed expected results, and browser runs assert zero console
errors and zero non-localhost network requests. These were ad hoc scripts run
during the build and are not part of the committed test suite; recreate similarly
shaped scripts under `scripts/` if you need to re-verify a change (start the dev
server first, drive the page with Playwright's `chromium`, assert on the
`Download all outputs` flow).

## Deploying to Vercel

1. Push this repo (or just the `webapp/` subtree) to GitHub and import it in
   Vercel, with **Root Directory** set to `webapp/`.
2. **Set up an access gate before sharing the URL** — this handles real client
   PII and should not be a public, unauthenticated deployment:
   - Vercel dashboard → your project → **Settings → Deployment Protection**.
   - Enable **Standard Protection** (requires Vercel-account login for anyone on
     your team) or **Password Protection** (a single shared password — simplest
     for a small team of onboarding staff). Password Protection is on Vercel Pro
     and above.
   - This is enforced at Vercel's edge, before any request reaches the app — it
     can't be bypassed by anything in the client code.
3. Deploy. No environment variables are required for any of these tools.

## Project structure

```
src/
  app/
    page.tsx                     — landing page (tool cards)
    parent-child-import/page.tsx — Parent and Child Import
    bookings-import/page.tsx     — Bookings Import
    payment-plans-import/page.tsx — Payment Plans Import
    balance-adjustments/page.tsx — Balance Adjustments
    room-fee-names/page.tsx      — Room & Fee Names checking
  components/       — FileUploader, ServiceMapUploader, ResultsTable, SummaryMetrics, shadcn/ui primitives
  lib/
    validator/            — Parent and Child Import's engine (constants, parse, transforms,
                             validators, cross-row checks, pipeline, report writers) —
                             ported 1:1 from validator_v2.py
    tools/
      roomFeeNames/        — fee/room name mismatch logic + Excel report (check_names.py)
      paymentPlans/        — column-mapped CSV validation, date parsing, split CSVs,
                             3-sheet coloured error report (payment_plan_checker.py)
      balanceAdjustments/  — CSV/XLSX/HTML-table parsing, styled-template cloning via
                             ExcelJS, consolidated + duplicate reports (process_balance_adjustments.py)
      bookingsImport/      — duplicate detection, recurring-schedule overlap (Union-Find),
                             casual/recurring overlap, split-by-service CSVs, styled
                             multi-sheet reports (prepare_bookings_import.py — the most
                             complex tool; its CLI-only `_main_cli` fallback was not
                             ported, since it duplicates `main()` with no browser equivalent)
    state/
      serviceMapStore.ts   — Zustand + sessionStorage, shared service map across every tool
    download.ts            — generic client-side file/blob download helpers

public/
  balance-adjustments-template.xlsx — bundled Excel template (fetched client-side; not user data)
```

The business rules in `lib/validator/` and `lib/tools/*/` are a deliberate 1:1
port of the Python scripts — see each Python file's docstrings for rule
rationale. Changes to validation/transform logic should be made in both places
until the Python app is retired, or flagged to whoever owns the Onboarding team's
import rules.

### Known minor divergences from the Python originals

- **Room & Fee Names / Payment Plans / Bookings Import** now resolve QK Service
  IDs via the shared `ServiceMapping` class, which also matches on `QKDBID` (the
  original per-tool scripts only matched on `QKServiceID`). This only adds
  successful matches — it never breaks a match that worked before.
- **Bookings Import**'s flexible date-parsing fallback (for dates that don't
  match any of the 5 strict formats) is a best-effort re-implementation of
  pandas' `dayfirst=True` inference, not a literal port — verified against the
  formats actually seen in QK exports, but arbitrary verbose date strings may
  parse slightly differently than pandas would.
