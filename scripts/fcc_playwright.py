#!/usr/bin/env python3
"""
Drive the FCC EAS search with Playwright — the non-Claude scrape transport.

  !!  UNVALIDATED  !!
  This script has NEVER been run successfully end-to-end against apps.fcc.gov.
  It is a working starting point, not a proven path. The FCC sits behind Akamai
  Bot Manager, and whether Akamai accepts a Playwright-driven Chromium is an
  open question that only a live test answers.

  RUN THE SMOKE TEST FIRST:
      python scripts/fcc_playwright.py --smoke-test

  It submits one narrow query and tells you whether a results table came back.
  * PASS -> use --run for the full weekly scrape.
  * FAIL -> do NOT try to defeat the detection. Use the manual browser path
            in START-HERE.md instead (open the search in your normal browser,
            paste scripts/browser_extract.js into DevTools, drive it by hand).
            It always works and costs about ten minutes a week.

  Deliberately NO stealth/evasion plugins, no TLS impersonation, no UA spoofing.
  This is a government site; automating an ordinary browser is fine, defeating
  a bot defence is not. If it says no, take the no.

Setup:
    python -m pip install playwright
    python -m playwright install chromium

Output: a scraped.json in the shape weekly_leads.py ingest expects.
"""
import argparse
import json
import sys
import time
from datetime import date, datetime, timedelta
from pathlib import Path

FOLDER = Path(__file__).resolve().parent.parent
EXTRACT_JS = Path(__file__).resolve().parent / "browser_extract.js"
SEARCH_URL = "https://apps.fcc.gov/oetcf/eas/reports/GenericSearch.cfm"
PROFILE_DIR = FOLDER / ".browser-profile"   # persistent profile: cookies survive runs

# Highest-yield classes first, so a run cut short still banks most of the value.
# Cross-class dedupe is near-total for NII/PCE/6xx over DTS/DSS/DXX — see
# SCRAPER-ISSUES.md #13b/#13d. A weekly run of the first six loses almost nothing.
DEFAULT_CLASSES = ["DTS", "NII", "DSS", "DXX", "PCE", "DCD",
                   "TNB", "6XD", "6ID", "6FC", "6FX", "6SD", "6CD", "6PP", "6VL"]
PAGE_SIZE = 500
PAUSE = 2.0          # seconds between queries — be a polite guest


def _playwright():
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        sys.exit("ERROR: playwright not installed.\n"
                 "  python -m pip install playwright\n"
                 "  python -m playwright install chromium")
    return sync_playwright


def _open(pw, headless):
    """Persistent, ordinary Chromium. No fingerprint tampering."""
    ctx = pw.chromium.launch_persistent_context(
        str(PROFILE_DIR), headless=headless, viewport={"width": 1400, "height": 950})
    # Re-inject the helper on every navigation (form.submit() destroys the page).
    ctx.add_init_script(EXTRACT_JS.read_text(encoding="utf-8"))
    return ctx


def _submit(page, cls, from_rec=1):
    """Fill + submit the real form, then wait for the navigation to settle."""
    with page.expect_navigation(wait_until="domcontentloaded", timeout=120_000):
        page.evaluate("([c, f]) => FCC.submit({cls: c, fromRec: f, showRecords: %d})"
                      % PAGE_SIZE, [cls, from_rec])


def smoke_test(headless):
    """One narrow query. Answers the only question that matters: does it work?"""
    to = date.today()
    frm = to - timedelta(days=3)
    sync_playwright = _playwright()
    with sync_playwright() as pw:
        ctx = _open(pw, headless)
        page = ctx.new_page()
        print(f"opening {SEARCH_URL}")
        page.goto(SEARCH_URL, wait_until="domcontentloaded", timeout=120_000)
        if not page.query_selector("form"):
            print("FAIL: the search page itself did not load a form.")
            ctx.close()
            return 1
        print(f"submitting DTS for {frm} .. {to}")
        page.evaluate("([a, b]) => FCC.reset(a, b)", [frm.isoformat(), to.isoformat()])
        try:
            _submit(page, "DTS")
        except Exception as e:
            print(f"FAIL: navigation after submit failed ({e.__class__.__name__}).")
            ctx.close()
            return 1
        try:
            res = page.evaluate("() => FCC.grab('DTS')")
        except Exception as e:
            body = (page.inner_text("body") or "")[:400]
            print(f"FAIL: no results table ({e}).\n--- page said ---\n{body}")
            ctx.close()
            return 1
        ctx.close()
        if res.get("empty"):
            print("PASS (transport works): search ran, but that window had no grants. "
                  "Widen --from/--to and rerun to see rows.")
            return 0
        print(f"PASS: {res['served']} rows served, {res['added']} US rows kept.")
        print("Playwright gets through. Use --run for the weekly scrape.")
        return 0


def run(frm, to, classes, out, headless):
    sync_playwright = _playwright()
    with sync_playwright() as pw:
        ctx = _open(pw, headless)
        page = ctx.new_page()
        page.goto(SEARCH_URL, wait_until="domcontentloaded", timeout=120_000)
        page.evaluate("([a, b]) => FCC.reset(a, b)", [frm.isoformat(), to.isoformat()])

        for cls in classes:
            from_rec, guard = 1, 0
            while guard < 20:                     # max_pages safety
                guard += 1
                page.goto(SEARCH_URL, wait_until="domcontentloaded", timeout=120_000)
                try:
                    _submit(page, cls, from_rec)
                    res = page.evaluate("(c) => FCC.grab(c)", cls)
                except Exception as e:
                    print(f"  {cls}: stopped ({e.__class__.__name__}: {e})")
                    break
                served, added = res.get("served", 0), res.get("added", 0)
                print(f"  {cls} FromRec={from_rec}: {served} served, "
                      f"{added} new unique US, {res.get('nonUS', 0)} non-US")
                if served < PAGE_SIZE:            # last page
                    break
                from_rec += served                # advance by rows ACTUALLY served
                time.sleep(PAUSE)
            time.sleep(PAUSE)

        payload = page.evaluate("() => FCC.payload()")
        ctx.close()

    Path(out).write_text(payload, encoding="utf-8")
    data = json.loads(payload)
    print(f"\n{len(data['rows'])} unique US rows -> {out}")
    print(f'Next:  python scripts/weekly_leads.py ingest "{out}"')
    return 0


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--smoke-test", action="store_true",
                    help="one narrow query; reports whether Playwright gets through")
    ap.add_argument("--run", action="store_true", help="full scrape")
    ap.add_argument("--from", dest="frm", help="YYYY-MM-DD (default: 7 days ago)")
    ap.add_argument("--to", help="YYYY-MM-DD (default: today)")
    ap.add_argument("--classes", help="comma-separated equipment classes")
    ap.add_argument("--out", default="scraped.json")
    ap.add_argument("--headless", action="store_true",
                    help="run without a visible window (more likely to be blocked)")
    a = ap.parse_args()

    if a.smoke_test:
        return smoke_test(a.headless)
    if not a.run:
        ap.error("pass --smoke-test (do this first) or --run")

    to = datetime.strptime(a.to, "%Y-%m-%d").date() if a.to else date.today()
    frm = (datetime.strptime(a.frm, "%Y-%m-%d").date() if a.frm
           else to - timedelta(days=7))
    classes = ([c.strip().upper() for c in a.classes.split(",") if c.strip()]
               if a.classes else DEFAULT_CLASSES)
    print(f"window {frm} .. {to}; classes: {', '.join(classes)}")
    return run(frm, to, classes, a.out, a.headless)


if __name__ == "__main__":
    sys.exit(main())
