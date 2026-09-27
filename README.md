# Nemko CRA Lead-Gen System

A ~$0/month lead-gen loop for Nemko CRA (Cyber Resilience Act) outreach, driven
by new FCC equipment-authorization grants.

**New here? Read [`START-HERE.md`](START-HERE.md) first** (setup), then
[`AGENTS.md`](AGENTS.md) (how to run the weekly loop).

## The idea

An FCC equipment-authorization grant is public proof that a company just
finished a **connected product** — which is exactly the CRA-relevant universe.
Every week the system pulls the new grants, works out which companies are worth
approaching, finds a contact, and hands you a spreadsheet. You do the judgement
and the sending.

## The weekly loop

1. **Scrape** — pull the week's grants per CRA-relevant equipment class from the
   FCC EAS search, US applicants only. Dedupe against the master ledger, the
   internal leads list, and the Blacklist tab (mega-OEMs and repeat filers that
   are a poor sales fit). → `weekly_leads.py ingest` creates the weekly workbook
   and appends to the master. Each lead's **Form 731** is then fetched for the
   product description and the applicant's contact. → `weekly_leads.py enrich731`
2. **Score** — the agent researches every lead against the **Scoring Model** tab
   and fills the CRA Scoring sheet (gates G1–G6, questions Q1–Q13, TOTAL /100,
   GRADE). A company with no findable information is never left incomplete:
   **G1 = TRUE → DQ**.
3. **Contacts** — `weekly_leads.py finalize` copies A/B-grade leads' contact
   info to the Contacts sheet and syncs scores back to the master.

You review the Contacts sheet and send outreach from your own mailbox.
**No automated sending, ever** (CAN-SPAM / GDPR + deliverability).

## Scoring

Knockout gates G1–G6 (any TRUE = DQ) plus Q1–Q13 summing to 100.
Grades: **A** 80–100 Pursue now · **B** 60–79 Strong · **C** 40–59 Qualify ·
**D** <40 Deprioritize · **DQ (gate)**.

The scrape pre-fills only what an FCC grant proves on its own — Q1 (digital/
connected), Q6 (wireless), Q7 (recent filing), Q10 (holds a grant) = 31 pts,
grade D. Everything above that comes from research.

## Honest limitations

- **The FCC contact is usually not the PM.** It's often a compliance agent or
  the test lab. Finding the Program/Project/Product Manager stays a manual
  LinkedIn step.
- **FCC is a late signal.** By the time a grant issues, the design is frozen and
  a test lab is often already engaged. It's excellent for *volume* and for
  cross-selling other Nemko services — it is not the "catch them before they
  choose a vendor" early window.
- **Cross-class overlap is heavy.** Dual-band devices file under several
  equipment classes; expect ~1,200 scraped rows to collapse to ~100 unique
  leads. See `AGENTS.md` → Yield notes.
- **The scrape needs a real browser.** The FCC is behind Akamai Bot Manager.
  See `START-HERE.md` §4.

## Try it with the sample data

`Nemko CRA Leads.xlsx` ships with fictitious demo leads, and `sample\` holds
fictitious inputs in the exact formats the scraper and Form-731 step produce.
No FCC access needed:

```powershell
.\setup.ps1
.\.venv\Scripts\python.exe scripts\weekly_leads.py ingest sample\scraped.json
.\.venv\Scripts\python.exe scripts\weekly_leads.py enrich731 --week "Weekly Leads\Week of 2026-09-19 to 2026-09-26.xlsx" sample\forms.json
.\.venv\Scripts\python.exe scripts\weekly_leads.py finalize --week "Weekly Leads\Week of 2026-09-19 to 2026-09-26.xlsx"
```

`ingest` should report 2 new, 3 dup, 1 non-US, 1 blacklist. The 7 sample rows
cover a cross-class duplicate, a company added within 30 days, an FCC ID
already in the ledger, a blacklisted company and a non-US applicant. Grade a
lead A or B on the CRA Scoring sheet before `finalize` to see it reach
Contacts. **Running this changes the ledger** — `git checkout -- "Nemko CRA Leads.xlsx"`
resets it before real use.

## Files

Full manifest in `START-HERE.md` §1. The two that matter most:
`Nemko CRA Leads.xlsx` (the master ledger — this *is* the dedupe state) and
`AGENTS.md` (the runbook and the rules).
