# AGENTS.md — runbook for the weekly FCC → CRA lead pipeline

Codex reads this file automatically. It is the operating manual for the agent
*and* the human. Setup instructions live in `START-HERE.md`.

Paths below use Windows form. On macOS/Linux swap
`.venv\Scripts\python.exe` → `./.venv/bin/python`.

---

## HARD RULES — these are not negotiable

1. **Never send outreach email, ever.** Not to a lead, not to a prospect, not
   "just a draft to test". A human reviews and sends every message from their
   own mailbox. (CAN-SPAM / GDPR + deliverability.) You may *write drafts to a
   file*. You may not transmit anything.
2. **Never modify `Nemko_CRA_DATA_Transfer.xlsx`** (the internal Nemko leads
   list). Read-only, for dedupe only. Never merge its rows into the tracker.
3. **US-based applicants only.** Non-US rows are dropped at scrape time.
4. **Never try to defeat the FCC's bot detection.** No TLS-impersonation tools
   (curl-impersonate and friends), no stealth browser plugins, no UA spoofing.
   Driving an ordinary browser is fine; evading a defence is not. If the site
   says no, take the no and fall back to the manual path.
5. **Do not retry CLI HTTP against `apps.fcc.gov`.** It times out by design —
   see "Transport" below. Retrying wastes the run.
6. **A company you cannot research is not left blank** — set gate G1 = TRUE and
   DQ it. Never leave a scoring row half-filled.
7. **Never rerun `ingest --force` after a successful ingest.** The leads are in
   the master by then, so it produces an empty weekly file.

---

## Transport: how to get data out of the FCC (read this before scraping)

`apps.fcc.gov` is behind **Akamai Bot Manager**. Hard-won facts — do not
re-litigate them, each cost hours:

- **Plain HTTP is blackholed.** `curl`, Python `urllib`, `requests` — all time
  out, even a bare `HEAD`. Cookie bootstrap, forced HTTP/1.1, and browser
  User-Agents were each tested and each failed.
- **In-page `fetch()` to `GenericSearchResult.cfm` returns 503**, even from a
  session that just loaded the search page normally.
- **What works: submitting the page's REAL `<form>`** so the browser performs a
  genuine top-level navigation. Navigate fresh to `GenericSearch.cfm`, set
  `document.forms[0]` fields, call `.submit()`.
- **`fetch()` DOES work for Form-731 report pages** (`/tcb/GetTcb731Report.do`).
  The block is scoped to the search-results endpoint, not the whole origin. So
  step 1f can be batched in one in-page async loop at ~1.6 s spacing.
- **Use `outputformat=HTML` and read `#rsTable`. Never XML.** Submitting the
  real form with `outputformat=XML` makes the browser *download* the response
  (`net::ERR_ABORTED`); the page is left unchanged and it looks exactly like a
  block but is not one.
- **Pagination is the hidden field `FromRec`** (1-based). The form's own
  `fetchfrom` field is **ignored by the server** — a scraper that advances
  `fetchfrom` silently re-reads page 1 forever. Advance `FromRec` by the number
  of rows *actually served*; stop when a page repeats or comes back empty.
- `show_records=500` works. A weekly (5–7 day) window is normally one page per
  class.
- Empty result looks like: no `#rsTable`, and page text
  "There are no applications on file that match the search criteria".

`scripts/browser_extract.js` implements all of the above. Load it into the page
and drive `FCC.reset()` / `FCC.submit()` / `FCC.grab()` / `FCC.payload()`.

Two supported ways to run it — see `START-HERE.md` §4:
- **A. Manual browser** (always works): DevTools console, paste, drive by hand.
- **B. Playwright** (`scripts/fcc_playwright.py`, **unvalidated** — run
  `--smoke-test` first): same JS, driven from Python.

---

## The weekly loop

### STEP 1 — SCRAPE

Window: `Config!last_run_date` → today.

For each equipment class in `CLASS_IDS` (`scripts/fcc_pull.py`, mirrored in
`browser_extract.js`), search with `application_status=GI` (Grant Issued),
`application_purpose=O` (Original grant), `outputformat=HTML`,
`show_records=500`, paginating with `FromRec`.

From each `#rsTable` row capture: applicant name, city, state, country, FCC ID,
application purpose, grant date, and the row's `GetTcb731Report` +
`ViewExhibitReport` hrefs. Keep only `country = United States`. Dedupe by FCC ID
across classes (this matters a lot — see "Yield" below).

Produce JSON:

```json
{"window": {"from": "2026-09-06", "to": "2026-09-13"},
 "rows": [{"company": "...", "city": "...", "state": "...", "country": "United States",
           "fcc_id": "...", "grant_date": "...", "purpose": "...",
           "equipment_class": "DTS", "form731": "https://...", "exhibits": "https://..."}]}
```

> **`window.from` / `window.to` MUST be ISO `YYYY-MM-DD`.** `ingest` builds the
> weekly filename from them, so `mm/dd/yyyy` creates bogus subpaths and
> `save()` raises `FileNotFoundError`. The FCC *form fields*
> (`grant_date_from` / `grant_date_to`) still take `mm/dd/yyyy` —
> `browser_extract.js` converts for you. On this failure nothing is changed
> (the weekly save happens before the master append) — fix the dates and rerun
> `ingest` **without** `--force`.

Then:

```
.venv\Scripts\python.exe scripts\weekly_leads.py ingest scraped.json
```

This dedupes (master ledger + internal workbook + the Blacklist tab), creates
the weekly workbook, appends the new leads to the master, and prints each new
lead's Form-731 URL.

**Enrich (step 1f).** Fetch each printed Form-731 URL (`fetch()` is fine here,
~1.6 s apart) and parse the page's innerText:
- Contact block **"Person at the applicant's address to receive grant or for
  contact"** → First Name / Last Name / Title / Telephone Number / Email.
  If empty, fall back to the **"Non Technical Contact"** block.
  **IGNORE "Technical Contact" — that is the test lab, not the customer.**
- **"Description of product as it is marketed"** → the product (skip the
  "(NOTE:…" line).

Build `[{fcc_id, product, contact_first, contact_last, contact_title,
contact_phone, contact_email, contact_source}]`, save as `forms.json`, then:

```
.venv\Scripts\python.exe scripts\weekly_leads.py enrich731 --week "Weekly Leads\Week of ... .xlsx" forms.json
```

If the FCC is unreachable even in a real browser: report and stop. Nothing is
changed on failure.

### STEP 2 — SCORE

Research **every** lead on the weekly file's **CRA Scoring** sheet and follow
the master workbook's **Scoring Model** tab exactly.

- Gates **G1–G6** first. Any TRUE → `DISQUALIFIED? = "DQ"` and
  `GRADE = "DQ (gate)"`.
- Then **Q1–Q13**. Points: Q1 15, Q2 15, Q3 10, Q4 6, Q5 4, Q6 5, Q7 5, Q8 5,
  Q9 8, Q10 6, Q11 6, Q12 5, Q13 10 (= 100).
- Grades: **80–100** `A — Pursue now` · **60–79** `B — Strong` ·
  **40–59** `C — Qualify` · **<40** `D — Deprioritize`.
- The script pre-sets Q1/Q6/Q7/Q10 = TRUE (31 pts) — that is only what the FCC
  grant itself proves. Evaluate the rest from research: company website,
  LinkedIn, product pages.
- **No findable information about the company → G1 = TRUE → DQ.** Never leave
  a row incomplete.
- Fill `Key evidence`, `Confidence`, and a one-line `Research note` per row.
- Score **only the current week's leads.**
- Edit with openpyxl: **load once, save once.** If a save fails because the
  file is open in Excel, report it — do not retry in a loop.

### STEP 3 — CONTACTS

```
.venv\Scripts\python.exe scripts\weekly_leads.py finalize --week "Weekly Leads\Week of ... .xlsx"
```

Copies contact info for qualifying grades (`Config!qualify_grades`, currently
`A,B`) to the **Contacts** sheet and syncs all scores back to the master ledger.

Optional per-lead helper (browser session only):
`https://apps.fcc.gov/OETLabServices/getFCCIDList?fccId=<FCCID>` returns
canonical grantee / address / purpose.

**Delivery is the file.** The finished workbook in `Weekly Leads\` *is* the
deliverable — no email step. The operator opens it, reviews the Contacts sheet,
and sends outreach manually.

If `email-template.md` has a voice sample below its marker line, also draft a
first-touch email per Contacts row into a companion `*-drafts.md` file next to
the weekly workbook: reference the FCC filing (product + grant date), lead with
the CRA deadline, soft secondary hook to other Nemko services (safety, EMC,
radio), target the Program/Project/Product Manager. **Write the file. Never
send.**

### REPORT

End every run with: grants scraped · non-US dropped · new leads · DQ count and
reason distribution · A/B leads with contact info this week · what is waiting on
the operator (review Contacts sheet, send emails).

---

## Yield notes (save yourself queries)

- Cross-class dedupe is **near-total** for the overlap classes. In the
  2026-08-30 run, 1,233 scraped rows → 476 US → **100 unique FCC IDs**; NII,
  PCE and every 6xx class contributed **zero** net-new unique IDs over
  DTS/DSS/DXX (dual-band devices file under both).
- Practical schedule: weekly runs on **DTS, DSS, DXX** (+ PCE, DCD) lose almost
  nothing; do a full 15-class sweep monthly.
- `6CD`, `6PP`, `6VL` have returned zero results across runs — drop candidates.
- Query high-yield classes first so a run cut short still banks the value.
- Company-level dedupe is a 30-day window (`Config!company_dedupe_days`). This
  is why daily runs are pointless — they return the same leads for 7× the query
  volume.

## Etiquette

~35 POSTs over ~15 minutes with 1.5–2 s spacing has never drawn throttling.
Keep the spacing. Don't hammer.

---

## Config (master workbook → Config tab, no code edits needed)

| Setting | Current | Meaning |
|---|---|---|
| `last_run_date` | set by `ingest` | start of the next window (YYYY-MM-DD) |
| `lookback_days_default` | 7 | used when `last_run_date` is blank |
| `company_dedupe_days` | 30 | skip a company added within N days |
| `us_only` | TRUE | keep only United States applicants |
| `internal_list_path` | `Nemko_CRA_DATA_Transfer.xlsx` | read-only dedupe source; blank = skip |
| `include_equipment_classes` | 15 classes | which classes to pull |
| `qualify_grades` | A,B | grades that reach the Contacts sheet |
| `weekly_folder` | `Weekly Leads` | where weekly workbooks are written |

## Files

| Path | Role |
|---|---|
| `Nemko CRA Leads.xlsx` | **Master ledger — the dedupe state.** Tabs: Scoring Model, All Leads, Contact Sheet, Config, Instructions, Blacklist |
| `Nemko_CRA_DATA_Transfer.xlsx` | Internal Nemko leads — READ ONLY |
| `Weekly Leads\*.xlsx` | The deliverables (3 sheets each) |
| `scripts\weekly_leads.py` | `ingest` / `enrich731` / `finalize` |
| `scripts\fcc_pull.py` | Shared helpers + constants. Its network path is dead (Akamai) — imported for `CLASS_IDS`, scoring, dedupe |
| `scripts\browser_extract.js` | The scrape logic, transport-agnostic |
| `scripts\fcc_playwright.py` | Optional Playwright driver — **unvalidated**, smoke-test first |
| `scripts\receiver.py` | One-shot localhost JSON receiver. Only needed if your browser can't hand JSON back directly |
