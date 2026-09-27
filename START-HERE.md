# START HERE — FCC → CRA Lead-Gen System

A weekly loop that pulls new FCC equipment-authorization grants, CRA-scores the
companies behind them, and produces a spreadsheet of A/B-grade leads with
contact details for you to email manually.

This package is **self-contained and independent**. It has no link back to the
machine it came from — you run your own copy, your own ledger, your own
schedule.

Built and run on macOS with Claude Code. **You are on Windows with Codex**, so
§2 and §3 cover exactly what changes. The Python itself is already
cross-platform — the differences are in the *harness*, not the code.

Read in this order: this file → `AGENTS.md` (the weekly runbook).

---

## 1. What's in the box

| Path | Role |
|---|---|
| `AGENTS.md` | **The runbook.** Codex reads this automatically. Every operating rule and hard-won gotcha is in it |
| `Nemko CRA Leads.xlsx` | **Master ledger — the dedupe state.** Ships with 5 fictitious demo leads. Tabs: Scoring Model, All Leads, Contact Sheet, Config, Instructions, Blacklist |
| `Weekly Leads\*.xlsx` | Created by `ingest` — the weekly deliverables (git-ignored) |
| `sample\scraped.json`, `sample\forms.json` | Fictitious inputs in the exact scrape / Form-731 formats — see README |
| `scripts\weekly_leads.py` | The pipeline: `ingest` / `enrich731` / `finalize` |
| `scripts\fcc_pull.py` | Shared helpers, `CLASS_IDS`, scoring model, dedupe logic |
| `scripts\browser_extract.js` | **The scrape logic.** Works in DevTools, in Playwright, or driven by an agent |
| `scripts\fcc_playwright.py` | Optional automation — **unvalidated, smoke-test it first** (§4B) |
| `scripts\receiver.py` | One-shot localhost JSON receiver; only needed for some browser setups |
| `setup.ps1` | Windows setup: builds the venv, verifies the ledger |
| `requirements.txt` | `openpyxl==3.1.5` — the only dependency |

**Not included, by design:** the internal Nemko leads workbook and the email
delivery step. See §7.

---

## 2. What was Claude-specific, and what it becomes on Codex

The Python does not care what agent drives it. Four things do.

| # | On Claude Code | What it actually needs | On Codex |
|---|---|---|---|
| 1 | Scheduled task at `~/.claude/scheduled-tasks/<name>/SKILL.md`, fired by the app | Something to run a prompt weekly | **Windows Task Scheduler** launching Codex with the prompt (§6). Codex has no built-in scheduler |
| 2 | Built-in in-app browser (the Browser pane) doing a real form submit | A *real browser* to beat Akamai | **The hard one.** Three options in §4 — manual DevTools, Playwright MCP, or `fcc_playwright.py` |
| 3 | Claude memory files carrying the scraper gotchas | Persistent project knowledge | **`AGENTS.md`** — Codex reads it automatically. Everything from memory is already baked into it |
| 4 | `WebSearch` / `WebFetch` for scoring research | Web access during step 2 | Codex's **web search** — enable it (`--search`, or `tools.web_search` in `~\.codex\config.toml`). Flag names move between versions; check `codex --help` |

**The `SKILL.md` prompt is gone** — its content is now `AGENTS.md`, which is
better anyway because it is version-controlled next to the code instead of
living in a hidden app directory.

**Nothing else was Claude-exclusive.** `weekly_leads.py`, `fcc_pull.py` and
`receiver.py` are pure standard library + `openpyxl`, use `pathlib` throughout,
and contain no shell-outs and no platform branches. They were compile-checked
for this handoff.

---

## 3. Windows setup

### 3.1 Python
Windows doesn't ship it. Install **Python 3.9 or newer** from
[python.org](https://www.python.org/downloads/windows/) and tick
**"Add python.exe to PATH"** during install.

### 3.2 Put the folder somewhere without OneDrive
Pick a plain local path — `C:\Nemko\Lead Gen` is ideal.

**Avoid OneDrive-synced `Desktop` or `Documents`.** OneDrive's Files-On-Demand
can evict the `.xlsx` files to cloud-only stubs, and the scripts will fail with
confusing I/O errors mid-run. If you must keep it there, right-click the folder
→ **Always keep on this device**.

### 3.3 Build the environment
In PowerShell, from the folder:

```powershell
.\setup.ps1
```

That creates `.venv`, installs `openpyxl`, smoke-tests the scripts, and prints
the ledger row count. If PowerShell refuses to run the script:

```powershell
powershell -ExecutionPolicy Bypass -File .\setup.ps1
```

Manual equivalent:

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

### 3.4 The path difference to remember

| macOS/Linux | Windows |
|---|---|
| `./.venv/bin/python` | `.venv\Scripts\python.exe` |
| `python3` | `python` (or `py -3`) |

Every command in `AGENTS.md` is already written in Windows form. If you copy a
command out of an old note, translate it.

### 3.5 Windows things that bite

- **Close the workbooks in Excel before running anything.** Excel takes an
  exclusive lock; the scripts detect it and exit with a clear message rather
  than corrupting anything, but the run stops. This bites harder on Windows
  than on macOS.
- **Paths with spaces need quotes:**
  `--week "Weekly Leads\Week of 2026-09-06 to 2026-09-13.xlsx"`
- **Controlled Folder Access** (Windows Security → Ransomware protection) can
  silently block writes to `Documents`/`Desktop`. If saves fail for no visible
  reason, either move the folder or allow `python.exe` through.
- macOS needed Full Disk Access for scheduled jobs. **Windows has no
  equivalent** — nothing to do here.

---

## 4. The scrape transport — the one genuinely hard part

**`apps.fcc.gov` is behind Akamai Bot Manager.** Plain HTTP is blackholed
(curl, `requests`, `urllib` — all time out). In-page `fetch()` to the
search-results endpoint returns 503. The only thing that works is a **real
browser performing a genuine top-level form submit**.

`scripts\browser_extract.js` contains all the scrape logic and is
transport-agnostic. Pick how you drive it:

### A. Manual browser — always works, ~10 min/week. Start here.

1. Open `https://apps.fcc.gov/oetcf/eas/reports/GenericSearch.cfm` in Chrome or
   Edge.
2. Open DevTools (**F12**) → Console. Paste the whole of
   `scripts\browser_extract.js` and press Enter.
3. Drive it:
   ```js
   FCC.reset('2026-09-06','2026-09-13')   // ISO dates — this matters
   FCC.submit({cls:'DTS'})                // page navigates to results
   // re-paste the script (navigation clears it), then:
   FCC.grab('DTS')                        // {served, added, nonUS, uniqueTotal}
   FCC.submit({cls:'DSS'}); /* re-paste */ FCC.grab('DSS')
   // ...DXX, PCE, DCD. Then:
   copy(FCC.payload())                    // JSON now on your clipboard
   ```
   State lives in `localStorage`, so it survives every navigation — only the
   `FCC` helper itself needs re-pasting.
4. Paste into `scraped.json` and run `ingest`.

> Tip: put the script in a DevTools **Snippet** (Sources → Snippets) and it's
> one click to re-run instead of a re-paste.

### B. Playwright — automated, but **unvalidated**

```powershell
.\.venv\Scripts\python.exe -m pip install playwright
.\.venv\Scripts\python.exe -m playwright install chromium
.\.venv\Scripts\python.exe scripts\fcc_playwright.py --smoke-test
```

The smoke test submits one narrow query and tells you in about 30 seconds
whether Akamai accepts a Playwright-driven Chromium.

- **PASS** → `.\.venv\Scripts\python.exe scripts\fcc_playwright.py --run --from 2026-09-06 --to 2026-09-13`
- **FAIL** → use option A. **Do not** reach for stealth plugins or
  TLS-impersonation tools to get around it. This is a government site;
  automating an ordinary browser is fine, defeating a bot defence is not.

Nobody has run this successfully end-to-end — it was written for this handoff
and could not be tested without hitting the FCC. Treat a PASS as good news, not
an expectation.

### C. Playwright MCP inside Codex — most Claude-like

Add a browser MCP server to `~\.codex\config.toml` so Codex can drive the
browser itself:

```toml
[mcp_servers.playwright]
command = "npx"
args = ["-y", "@playwright/mcp@latest"]
```

Then Codex can navigate, run `browser_extract.js`, and read results the way
Claude's built-in browser did. Same Akamai caveat as B — it's still Playwright
underneath, so validate it with one query before trusting a full run. (Check
current Playwright-MCP docs for the exact package name and tool set; this
ecosystem moves fast.)

---

## 5. Codex configuration

1. **`AGENTS.md` is already in the folder root** — Codex picks it up with no
   configuration. That is your runbook and your memory.
2. **Enable web search.** Step 2 scores every lead from live research. Without
   web access, every company hits the "no information found → G1 → DQ" rule and
   your week's output is empty. This is not a nice-to-have.
3. **Optional: a browser MCP** (§4C).
4. **Approvals.** For a weekly run you'll want Codex able to run
   `.venv\Scripts\python.exe` and edit files in the folder without prompting on
   every call. Use whatever the current Codex approval/sandbox setting is for
   "workspace write" — but keep it scoped to this folder.

---

## 6. Scheduling on Windows

Codex has no built-in scheduler. Use **Task Scheduler**:

1. Task Scheduler → **Create Task**
2. *Triggers* → New → Weekly, Monday, 08:30
3. *Actions* → New → Program: `powershell.exe`, Arguments:
   ```
   -NoProfile -Command "cd 'C:\Nemko\Lead Gen'; codex exec 'Run the weekly FCC CRA lead pipeline per AGENTS.md' | Tee-Object -FilePath .\run.log"
   ```
4. *Conditions* → untick "Start only if on AC power" if it's a laptop
5. *Settings* → tick **"Run task as soon as possible after a scheduled start is
   missed"**

Verify the exact non-interactive invocation against `codex --help` — the
subcommand for unattended runs varies by version.

**Honestly: run it manually for the first month.** The scrape step needs
judgement when the FCC misbehaves, and you want to have watched the loop a few
times before you let it run unattended. A calendar reminder every Monday beats
a scheduled task that fails silently.

---

## 7. What's deliberately not in this package

### The internal leads workbook
`Nemko_CRA_DATA_Transfer.xlsx` is Nemko-internal lead data. It is **not**
included — source your own copy through Nemko rather than receiving it
second-hand.

The pipeline reads it **read-only**, purely so companies already being worked
internally don't get re-added as "new" leads. To wire yours in: drop the file in
the folder root and set `Config!internal_list_path` to its filename. It reads
the `Company Name` column of an `All Leads` tab and the `Company` column of a
`Contact Sheet` tab.

**Without it the pipeline runs fine** — `internal_companies()` returns an empty
set and that one dedupe layer is skipped. The master-ledger dedupe and the
Blacklist tab still work. You'll just occasionally surface a company a colleague
already has.

### The email delivery step
The original had a background job that emailed the finished workbook to the
operator via a personal mail relay. **That's gone.** Delivery is now just the
file: `finalize` writes the workbook into `Weekly Leads\` and you open it. No
credentials, no relay, no daemon.

(Outreach email was *never* automated and must never be — see the hard rules.)

---

## 8. You are now independent — and so is the other system

Your copy and the original share a starting point and nothing else. That has
one consequence worth planning around:

**You will both surface the same companies.** The FCC grant list is the same for
everyone. If both systems keep running, the same manufacturer can get outreach
from two Nemko people in the same week — which looks bad to the prospect.

Decide up front, with whoever else is running this:

- **Only one system runs** (cleanest — the other is retired), or
- **Split the source**: one takes DTS/NII/DSS, the other DXX/PCE/DCD, via
  `Config!include_equipment_classes`, or
- **Split by territory** and filter on the `State` column, or
- **Reconcile in the CRM** before anyone sends — treat Salesforce as the real
  dedupe layer and the ledgers as local caches.

The master ledger you were given is a **snapshot**, not a live feed. It stops
reflecting the other system the moment you start running.

---

## 9. First-run validation

Do all of this before trusting a single lead.

```powershell
# 1. dependency
.\.venv\Scripts\python.exe -c "import openpyxl; print(openpyxl.__version__)"   # 3.1.5

# 2. the pipeline loads
.\.venv\Scripts\python.exe scripts\weekly_leads.py --help
.\.venv\Scripts\python.exe scripts\weekly_leads.py ingest --help
.\.venv\Scripts\python.exe scripts\weekly_leads.py enrich731 --help
.\.venv\Scripts\python.exe scripts\weekly_leads.py finalize --help

# 3. the ledger opens and carries its history
.\.venv\Scripts\python.exe -c "from openpyxl import load_workbook as L; w=L('Nemko CRA Leads.xlsx'); print(w.sheetnames); print(w['All Leads'].max_row)"
# -> ['Scoring Model','All Leads','Contact Sheet','Config','Instructions','Blacklist']
# -> 6          (5 demo leads + header row)
```

4. **Transport smoke test** (§4). Get *one* successful search returning an
   `#rsTable`. Until this passes, nothing else matters.
5. **Set the window.** Open `Nemko CRA Leads.xlsx` → Config tab → set
   `last_run_date` to the day you're starting from. Save and close Excel.
6. **One full manual cycle over a short window**: scrape → `ingest` →
   enrich → score 1–2 leads by hand → `finalize`. Confirm a new workbook appears
   in `Weekly Leads\` with all three sheets and that Contacts is populated for
   any A/B lead. Compare its shape against the one produced from `sample\` (README).

Only then consider scheduling it.

---

## 10. The hard rules

Full versions in `AGENTS.md`. The short list:

1. **Never send outreach email automatically.** A human reviews and sends every
   message, from their own mailbox.
2. **Never modify the internal leads workbook.** Read-only, never merged.
3. **US-based applicants only.**
4. **Never try to defeat the FCC's bot detection.** No impersonation tooling, no
   stealth plugins. If the site says no, take the no.
5. **Unresearchable company → G1 = TRUE → DQ.** Never leave a scoring row blank.
