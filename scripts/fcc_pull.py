#!/usr/bin/env python3
"""
Weekly FCC equipment-authorization pull for the Nemko CRA lead tracker.

Pull new grants from the official FCC OET EAS search since the last run,
filter to CRA-relevant equipment classes and (by default) US-based
applicants, dedupe against the tracker AND the internal leads workbook,
score per the internal scoring model, and append to the "All Leads" tab
of "Nemko CRA Leads.xlsx".

The internal workbook (Nemko_CRA_DATA_Transfer.xlsx) is READ-ONLY here:
its companies are skipped so the two lists stay separate without
double-adding; it is never modified.

Run:  ./.venv/bin/python scripts/fcc_pull.py
      ./.venv/bin/python scripts/fcc_pull.py --from 2026-07-01 --to 2026-07-08
      ./.venv/bin/python scripts/fcc_pull.py --fixture path/to/saved_results.html

Fails loudly and changes nothing on network/parse errors (exit code 1).
No email is ever sent by this script.
"""
import argparse
import html
import re
import sys
import urllib.parse
import urllib.request
from datetime import date, datetime, timedelta
from html.parser import HTMLParser
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Font

# --------------------------------------------------------------- constants
# Official FCC OET Equipment Authorization Search (public, free).
# Form spec verified against the live GenericSearch.cfm form on 2026-07-11:
# the form POSTs to GenericSearchResult.cfm; application_status "GI" = Grant
# Issued; application_purpose "O" = Original Grant; outputformat XML returns
# every matching row (show_records is ignored); equipment_class takes the
# numeric IDs in CLASS_IDS below.
# CAVEAT: apps.fcc.gov rejects/blackholes CLI HTTP clients (curl AND python
# urllib time out) while real browsers work — see SCRAPER-ISSUES.md #1. If
# this direct fetch fails, do a browser-assisted pull and feed the saved XML
# to --fixture.
SEARCH_URL = ("https://apps.fcc.gov/oetcf/eas/reports/GenericSearchResult.cfm"
              "?RequestTimeout=500")
BASE_PARAMS = {
    "calledFromFrame": "N",
    "application_status": "GI",   # Grant Issued
    "application_purpose": "O",   # Original Grant (new products)
    "outputformat": "XML",
    "show_records": "500",
    "fetchfrom": "1",
}
PARAM_DATE_FROM = "grant_date_from"  # MM/DD/YYYY ("Final Action Date" range)
PARAM_DATE_TO = "grant_date_to"
# equipment_class numeric form IDs, read from the live form 2026-07-11
CLASS_IDS = {
    "DTS": "103", "NII": "102", "DSS": "15", "DXX": "18", "DCD": "12",
    "TNB": "56", "PCE": "41",
    # 6 GHz WiFi 6E/7 classes
    "6CD": "246", "6FC": "252", "6FX": "251", "6ID": "243", "6PP": "244",
    "6SD": "250", "6VL": "255", "6XD": "245",
}
USER_AGENT = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7)"
TIMEOUT = 90

FOLDER = Path(__file__).resolve().parent.parent
WORKBOOK = FOLDER / "Nemko CRA Leads.xlsx"

# "All Leads" columns (1-based). Internal-template block first (29 cols),
# then the automation block. Must match the workbook header order.
G_FIRST, G_LAST = 7, 12       # G1..G6
DQ_COL = 13                   # DISQUALIFIED?
Q_FIRST, Q_LAST = 14, 26      # Q1..Q13
COL = {
    "status": 1, "account_owner": 2, "company": 3, "office": 4,
    "total": 5, "grade": 6, "key_evidence": 27, "confidence": 28, "notes": 29,
    "lead_id": 30, "date_added": 31, "source": 32, "fcc_id": 33,
    "grant_date": 34, "equipment_class": 35, "product": 36, "country": 37,
    "linkedin": 38,
}

# Internal scoring model: points per Q1..Q13 (sums to 100)
Q_POINTS = [15, 15, 10, 6, 4, 5, 5, 5, 8, 6, 6, 5, 10]
# What an FCC grant proves on its own: Q1 digital/connected, Q6 wireless,
# Q7 FCC filing <12mo, Q10 holds an FCC grant. Everything else is for
# enrichment per the Scoring Model tab's "How to test" column.
SCRIPT_TRUE_QS = {1, 6, 7, 10}

GRADE_BANDS = [(80, "A — Pursue now"), (60, "B — Strong"),
               (40, "C — Qualify"), (0, "D — Deprioritize")]

US_NAMES = {"united states", "usa", "us", "u s a", "u s",
            "united states of america"}

# Equipment class -> rough connectivity label
CLASS_INFO = {
    "DTS": "WiFi", "NII": "WiFi", "TNB": "Cellular", "PCE": "Cellular",
    "DSS": "Bluetooth", "DXX": "BLE/Zigbee/sub-GHz",
    "DCD": "Low-power <1705kHz",
    "6CD": "WiFi 6E/7", "6FC": "WiFi 6E/7", "6FX": "WiFi 6E/7",
    "6ID": "WiFi 6E/7", "6PP": "WiFi 6E/7", "6SD": "WiFi 6E/7",
    "6VL": "WiFi 6E/7", "6XD": "WiFi 6E/7",
}


# ------------------------------------------------------------ HTML parsing
class TableGrabber(HTMLParser):
    """Collect every <table> as rows of cell text, and all anchors.

    Tables on the FCC pages are nested (layout tables wrap the results
    table), so open tables are tracked as a stack; each table's rows/cells
    only ever touch the innermost open table.
    """

    def __init__(self):
        super().__init__()
        self.tables, self.links = [], []
        self._stack = []  # one {"rows": [], "row": None, "cell": None} per open table
        self._href = None

    def handle_starttag(self, tag, attrs):
        if tag == "table":
            self._stack.append({"rows": [], "row": None, "cell": None})
        elif tag == "tr" and self._stack:
            self._stack[-1]["row"] = []
        elif tag in ("td", "th") and self._stack and self._stack[-1]["row"] is not None:
            self._stack[-1]["cell"] = []
        elif tag == "a":
            self._href = dict(attrs).get("href")
            if self._href is not None:
                self.links.append([self._href, ""])

    def handle_endtag(self, tag):
        if not self._stack:
            return
        top = self._stack[-1]
        if tag == "table":
            self._flush_cell(top)
            if top["row"]:
                top["rows"].append(top["row"])
            if top["rows"]:
                self.tables.append(top["rows"])
            self._stack.pop()
        elif tag == "tr" and top["row"] is not None:
            self._flush_cell(top)
            if top["row"]:
                top["rows"].append(top["row"])
            top["row"] = None
        elif tag in ("td", "th"):
            self._flush_cell(top)
        elif tag == "a":
            self._href = None

    @staticmethod
    def _flush_cell(top):
        if top["cell"] is not None and top["row"] is not None:
            top["row"].append(" ".join("".join(top["cell"]).split()))
            top["cell"] = None

    def handle_data(self, data):
        if self._stack and self._stack[-1]["cell"] is not None:
            self._stack[-1]["cell"].append(data)
        if self._href is not None and self.links:
            self.links[-1][1] += data


def parse_results_xml(xml_text):
    """Parse the EAS search's outputformat=XML response.

    Row schema (observed live 2026-07-11): applicant_name, address, city,
    state, country, zip_code, fcc_id, application_purpose, grant_date,
    lower/upper_freq_mhz. The XML has NO equipment class or product
    description; an <equipment_class> element is injected per row by our
    browser-assisted fetch, and defaults to "" for raw server output.
    """
    # The FCC's XML is NOT well-formed (unescaped & in company names), so a
    # strict parser would silently truncate. Regex extraction per <Row> is
    # deliberate — see SCRAPER-ISSUES.md #2.
    def unesc(s):
        return s.replace("&lt;", "<").replace("&gt;", ">").replace("&amp;", "&")

    grants = []
    for m in re.finditer(r"<Row>(.*?)</Row>", xml_text, re.S):
        block = m.group(1)

        def val(tag):
            t = re.search(rf"<{tag}>(.*?)</{tag}>", block, re.S)
            return unesc(t.group(1).strip()) if t else ""
        fcc_id = val("fcc_id")
        if not fcc_id:
            continue
        state = val("state")
        grants.append({
            "fcc_id": fcc_id,
            "company": val("applicant_name"),
            "equipment_class": val("equipment_class"),
            "grant_date": val("grant_date"),
            "country": val("country"),
            "description": val("product_description"),
            "city": val("city").rstrip(","),
            "state": "" if state.upper() in ("N/A", "") else state,
        })
    return grants, None


def parse_results(text):
    """Return (list of grant dicts, next-page URL or None).

    Auto-detects XML (outputformat=XML responses / saved fixtures) vs HTML.
    HTML parsing is header-driven: finds the table whose header row contains
    'FCC ID', then maps columns by header name so column order changes don't
    break us.
    """
    if text.lstrip().startswith("<?xml") or text.lstrip().startswith("<Results"):
        return parse_results_xml(text)
    html_text = text
    p = TableGrabber()
    p.feed(html_text)
    grants = []
    for tbl in p.tables:
        header_idx = next(
            (i for i, row in enumerate(tbl)
             if any("fcc id" in c.lower() for c in row)), None)
        if header_idx is None:
            continue
        headers = [c.lower() for c in tbl[header_idx]]

        def col(*needles, exclude=()):
            for i, h in enumerate(headers):
                if any(n in h for n in needles) and not any(x in h for x in exclude):
                    return i
            return None

        i_fccid = col("fcc id", exclude=("composite", "related"))
        i_company = col("applicant", "grantee name")
        i_class = col("equipment class")
        # rsTable renders the header as "FinalActionDate" (no spaces survive
        # cell-text join), hence the "finalaction" needle
        i_date = col("grant date", "final action date", "finalaction")
        i_country = col("country")
        i_desc = col("description")
        i_city = col("city")
        i_state = col("state", exclude=("united",))
        if i_fccid is None or i_company is None:
            continue

        def cell(row, idx):
            return row[idx].strip() if idx is not None and idx < len(row) else ""

        for row in tbl[header_idx + 1:]:
            fcc_id = cell(row, i_fccid)
            if not fcc_id or "fcc id" in fcc_id.lower():
                continue
            grants.append({
                "fcc_id": fcc_id,
                "company": html.unescape(cell(row, i_company)),
                "equipment_class": cell(row, i_class),
                "grant_date": cell(row, i_date),
                "country": cell(row, i_country),
                "description": html.unescape(cell(row, i_desc)),
                "city": cell(row, i_city),
                "state": cell(row, i_state),
            })
        break  # first matching table is the results grid
    next_url = next(
        (href for href, text in p.links if "next" in text.lower()), None)
    if next_url:
        next_url = urllib.parse.urljoin(SEARCH_URL, next_url)
    return grants, next_url


# ------------------------------------------------------------------ config
def read_config(wb):
    cfg = {}
    for row in wb["Config"].iter_rows(min_row=2, max_col=2, values_only=True):
        if row[0]:
            cfg[str(row[0]).strip()] = row[1]
    return cfg


def csv_set(value):
    return {v.strip().upper() for v in str(value or "").split(",") if v.strip()}


def truthy(value):
    return str(value).strip().lower() in ("true", "yes", "y", "1")


# ------------------------------------------------------------------- fetch
def fetch_post(params):
    data = urllib.parse.urlencode(params).encode()
    req = urllib.request.Request(
        SEARCH_URL, data=data,
        headers={"User-Agent": USER_AGENT,
                 "Content-Type": "application/x-www-form-urlencoded"})
    with urllib.request.urlopen(req, timeout=TIMEOUT) as resp:
        return resp.read().decode("utf-8", errors="replace")


def pull_grants(date_from, date_to, include_classes):
    """One XML POST per equipment class; the class tag is injected into each
    grant since the XML output doesn't carry it.

    Each class is fetched twice and the larger row set wins: the server
    occasionally returns a truncated-but-complete-looking response (see
    SCRAPER-ISSUES.md #3).
    """
    import time
    grants = []
    for cls in sorted(include_classes):
        class_id = CLASS_IDS.get(cls)
        if not class_id:
            print(f"WARNING: no form ID known for equipment class {cls}; skipped")
            continue
        params = dict(BASE_PARAMS)
        params["equipment_class"] = class_id
        params[PARAM_DATE_FROM] = date_from.strftime("%m/%d/%Y")
        params[PARAM_DATE_TO] = date_to.strftime("%m/%d/%Y")
        best = []
        for _ in range(2):
            rows = parse_results(fetch_post(params))[0]
            if len(rows) > len(best):
                best = rows
            time.sleep(1.5)
        for g in best:
            g["equipment_class"] = g["equipment_class"] or cls
            grants.append(g)
    return grants


# ----------------------------------------------------------------- scoring
def class_code(equipment_class):
    """'DTS - Digital Transmission System' or 'DTS' -> 'DTS'."""
    m = re.match(r"\s*([A-Z0-9]{2,4})\b", str(equipment_class).upper())
    return m.group(1) if m else ""


def parse_grant_date(s):
    for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%m/%d/%y"):
        try:
            return datetime.strptime(str(s).strip(), fmt).date()
        except (ValueError, AttributeError):
            continue
    return None


def grade_for(total):
    for floor, label in GRADE_BANDS:
        if total >= floor:
            return label
    return GRADE_BANDS[-1][1]


def norm_company(name):
    n = re.sub(r"[^a-z0-9 ]", " ", str(name or "").lower())
    n = re.sub(r"\b(inc|incorporated|llc|llp|ltd|co|corp|corporation|company|"
               r"gmbh|bv|sa|srl|kk|kabushiki|kaisha|limited|holdings|group|"
               r"technologies|technology|electronics|international)\b", " ", n)
    return " ".join(n.split())


def is_us(country):
    return re.sub(r"[^a-z ]", "", str(country or "").lower()).strip() in US_NAMES


def linkedin_url(company):
    kw = urllib.parse.quote(
        f'"{company}" program manager OR project manager OR product manager')
    return f"https://www.linkedin.com/search/results/people/?keywords={kw}"


def internal_companies(cfg):
    """Companies already in the internal leads workbook (read-only)."""
    path = FOLDER / str(cfg.get("internal_list_path") or "").strip()
    names = set()
    if not path.name or not path.exists():
        return names
    src = load_workbook(path, read_only=True, data_only=True)
    for sheet, col_header in (("All Leads", "Company Name"), ("Contact Sheet", "Company")):
        if sheet not in src.sheetnames:
            continue
        ws = src[sheet]
        rows = ws.iter_rows(values_only=True)
        headers = [str(h).strip() if h else "" for h in next(rows, [])]
        if col_header not in headers:
            continue
        idx = headers.index(col_header)
        for row in rows:
            if idx < len(row) and row[idx]:
                names.add(norm_company(row[idx]))
    src.close()
    return names


def blacklisted_companies(wb):
    """Companies on the master workbook's 'Blacklist' tab (mega-OEMs, repeat
    filers that are a poor Nemko sales fit) -- skipped entirely on ingest,
    same tier as the internal-workbook skip."""
    names = set()
    if "Blacklist" not in wb.sheetnames:
        return names
    ws = wb["Blacklist"]
    rows = ws.iter_rows(values_only=True)
    headers = [str(h).strip() if h else "" for h in next(rows, [])]
    if "Company" not in headers:
        return names
    idx = headers.index("Company")
    for row in rows:
        if idx < len(row) and row[idx]:
            names.add(norm_company(row[idx]))
    return names


# ------------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--from", dest="date_from", help="YYYY-MM-DD (default: last_run_date or lookback)")
    ap.add_argument("--to", dest="date_to", help="YYYY-MM-DD (default: today)")
    ap.add_argument("--fixture", help="parse a saved HTML file instead of hitting the FCC (parser test)")
    ap.add_argument("--dry-run", action="store_true", help="report what would be appended, don't save")
    args = ap.parse_args()

    if not WORKBOOK.exists():
        sys.exit(f"ERROR: workbook not found at {WORKBOOK}")
    wb = load_workbook(WORKBOOK)
    cfg = read_config(wb)
    today = date.today()

    # date window
    if args.date_from:
        date_from = datetime.strptime(args.date_from, "%Y-%m-%d").date()
    elif cfg.get("last_run_date"):
        v = cfg["last_run_date"]
        date_from = v.date() if isinstance(v, datetime) else datetime.strptime(str(v)[:10], "%Y-%m-%d").date()
    else:
        date_from = today - timedelta(days=int(cfg.get("lookback_days_default") or 7))
    date_to = datetime.strptime(args.date_to, "%Y-%m-%d").date() if args.date_to else today

    # fetch
    include = csv_set(cfg.get("include_equipment_classes"))
    try:
        if args.fixture:
            grants, _ = parse_results(Path(args.fixture).read_text(errors="replace"))
        else:
            grants = pull_grants(date_from, date_to, include)
    except Exception as e:
        sys.exit(f"ERROR: FCC fetch failed ({e.__class__.__name__}: {e}). "
                 f"Nothing was changed. apps.fcc.gov may be down or blocking; retry later.")
    print(f"Fetched {len(grants)} grant rows for {date_from} .. {date_to}")

    # filters
    excl_kw = {k.lower() for k in csv_set(cfg.get("exclude_keywords"))}
    us_only = truthy(cfg.get("us_only"))
    kept, dropped_class, dropped_nonus = [], 0, 0
    for g in grants:
        code = class_code(g["equipment_class"])
        if (include and code not in include) or (
                excl_kw and any(k in g["description"].lower() for k in excl_kw)):
            dropped_class += 1
            continue
        if us_only and not is_us(g["country"]):
            dropped_nonus += 1
            continue
        kept.append(g)
    print(f"{len(kept)} rows kept ({dropped_class} filtered by class/keyword, "
          f"{dropped_nonus} non-US dropped)")

    # existing data for dedupe
    ws = wb["All Leads"]
    seen_ids, recent_companies = set(), {}
    dedupe_days = int(cfg.get("company_dedupe_days") or 30)
    for row in ws.iter_rows(min_row=2, values_only=True):
        if len(row) >= COL["fcc_id"] and row[COL["fcc_id"] - 1]:
            seen_ids.add(str(row[COL["fcc_id"] - 1]).strip().upper())
        comp = norm_company(row[COL["company"] - 1] if len(row) >= COL["company"] else "")
        added = row[COL["date_added"] - 1] if len(row) >= COL["date_added"] else None
        if comp and added:
            d = added.date() if isinstance(added, datetime) else None
            if d is None:
                try:
                    d = datetime.strptime(str(added)[:10], "%Y-%m-%d").date()
                except ValueError:
                    continue
            recent_companies[comp] = max(recent_companies.get(comp, d), d)
    internal = internal_companies(cfg)
    print(f"Cross-file dedupe list: {len(internal)} companies from the internal workbook")

    # next Lead ID
    max_id = 0
    for row in ws.iter_rows(min_row=2, min_col=COL["lead_id"], max_col=COL["lead_id"], values_only=True):
        m = re.match(r"L(\d+)", str(row[0] or ""))
        if m:
            max_id = max(max_id, int(m.group(1)))

    # scoring prefill: what the FCC record proves
    base_total = sum(pts for i, pts in enumerate(Q_POINTS, 1) if i in SCRIPT_TRUE_QS)

    added_n, skipped_dup, skipped_internal = 0, 0, 0
    for g in kept:
        fid = g["fcc_id"].upper()
        comp = norm_company(g["company"])
        if fid in seen_ids or (
                comp in recent_companies and (today - recent_companies[comp]).days <= dedupe_days):
            skipped_dup += 1
            continue
        if comp in internal:
            skipped_internal += 1
            continue
        seen_ids.add(fid)
        recent_companies[comp] = today

        code = class_code(g["equipment_class"])
        gd = parse_grant_date(g["grant_date"])
        office = ", ".join(x for x in (g["city"], g["state"]) if x)
        max_id += 1
        r = ws.max_row + 1
        evidence = (f"{g['description'] or CLASS_INFO.get(code, 'radio device')}; "
                    f"FCC grant {gd.isoformat() if gd else g['grant_date']} ({g['fcc_id']})")
        values = {
            "status": "Not Contacted", "company": g["company"], "office": office,
            "total": base_total, "grade": grade_for(base_total),
            "key_evidence": evidence, "confidence": "Low",
            "lead_id": f"L{max_id:04d}", "date_added": today.isoformat(),
            "source": "FCC", "fcc_id": g["fcc_id"],
            "grant_date": gd.isoformat() if gd else g["grant_date"],
            "equipment_class": code or g["equipment_class"],
            "product": g["description"], "country": g["country"],
        }
        if not args.dry_run:
            for key, v in values.items():
                ws.cell(row=r, column=COL[key], value=v)
            for c in range(G_FIRST, G_LAST + 1):
                ws.cell(row=r, column=c, value=False)
            ws.cell(row=r, column=DQ_COL, value="OK")
            for i, c in enumerate(range(Q_FIRST, Q_LAST + 1), 1):
                ws.cell(row=r, column=c, value=i in SCRIPT_TRUE_QS)
            link = ws.cell(row=r, column=COL["linkedin"], value="Find PM")
            link.hyperlink = linkedin_url(g["company"])
            link.font = Font(name="Arial", size=10, color="0563C1", underline="single")
        added_n += 1

    if args.dry_run:
        print(f"DRY RUN: would append {added_n} leads ({skipped_dup} duplicates, "
              f"{skipped_internal} already in internal workbook)")
        return

    # grow the Excel table to cover appended rows, stamp last_run_date, save
    if added_n:
        ws.tables["LeadsTable"].ref = f"A1:AL{ws.max_row}"
    for row in wb["Config"].iter_rows(min_row=2, max_col=2):
        if str(row[0].value).strip() == "last_run_date":
            row[1].value = date_to.isoformat()
            break
    try:
        wb.save(WORKBOOK)
    except PermissionError:
        sys.exit("ERROR: could not save — close 'Nemko CRA Leads.xlsx' in Excel and rerun.")
    print(f"Appended {added_n} new leads ({skipped_dup} duplicates, "
          f"{skipped_internal} already in internal workbook). "
          f"last_run_date -> {date_to.isoformat()}")


if __name__ == "__main__":
    main()
