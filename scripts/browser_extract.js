/* ===========================================================================
 * browser_extract.js — FCC EAS search helper, transport-agnostic.
 *
 * Paste this into the browser's DevTools console on
 *   https://apps.fcc.gov/oetcf/eas/reports/GenericSearch.cfm
 * ...or pass it to Playwright's page.evaluate() / an agent's JS tool.
 * It defines a global `FCC` object and touches nothing else.
 *
 * WHY THIS EXISTS
 * The FCC search sits behind Akamai Bot Manager. A plain HTTP client is
 * blackholed, and in-page fetch() to GenericSearchResult.cfm returns 503.
 * The ONLY thing that works is submitting the page's real <form> element so
 * the browser does a genuine top-level navigation. FCC.submit() does exactly
 * that; FCC.grab() then reads the results table the navigation produced.
 *
 * TYPICAL RUN (one equipment class):
 *   FCC.reset('2026-08-30','2026-09-06')   // once, at the start of a run
 *   FCC.submit({cls:'DTS'})                // page navigates to the results
 *   FCC.grab('DTS')                        // after it loads
 *   // more rows? FCC.submit({cls:'DTS', fromRec: FCC.lastRows+1}) then grab
 *   // next class:  FCC.submit({cls:'NII'}) ... FCC.grab('NII')
 *   FCC.count()                            // unique FCC IDs so far
 *   FCC.payload()                          // JSON string -> save -> ingest
 *
 * State survives navigations via localStorage, so you can submit/grab across
 * as many page loads as the run needs.
 * ========================================================================= */
(function () {
  'use strict';

  var KEY = 'fccRun';
  var SEARCH_URL = 'https://apps.fcc.gov/oetcf/eas/reports/GenericSearch.cfm';

  // equipment_class numeric form IDs, read from the live form 2026-07-11.
  // Keep in sync with CLASS_IDS in scripts/fcc_pull.py.
  var CLASS_IDS = {
    DTS:'103', NII:'102', DSS:'15', DXX:'18', DCD:'12', TNB:'56', PCE:'41',
    '6CD':'246', '6FC':'252', '6FX':'251', '6ID':'243', '6PP':'244',
    '6SD':'250', '6VL':'255', '6XD':'245'
  };

  // Column positions observed on the live rsTable (2026). Used only as a
  // fallback when the header text can't be matched.
  var FALLBACK_IDX = { fcc_id: 12 };

  function store() {
    try { return JSON.parse(localStorage.getItem(KEY)) || null; }
    catch (e) { return null; }
  }
  function save(s) { localStorage.setItem(KEY, JSON.stringify(s)); }

  function need() {
    var s = store();
    if (!s) throw new Error('No run in progress — call FCC.reset(fromISO, toISO) first.');
    return s;
  }

  function isoToUS(iso) {                       // '2026-09-06' -> '09/06/2026'
    var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(iso).trim());
    if (!m) throw new Error('Dates must be ISO YYYY-MM-DD, got: ' + iso);
    return m[2] + '/' + m[3] + '/' + m[1];
  }

  function txt(el) { return (el && el.textContent || '').replace(/\s+/g, ' ').trim(); }

  /* --- locate the results table and map its columns by header text ------- */
  function resultsTable() {
    var t = document.getElementById('rsTable');
    if (t) return t;
    // fallback: first table whose header row mentions "FCC ID"
    var tables = document.getElementsByTagName('table');
    for (var i = 0; i < tables.length; i++) {
      var r = tables[i].rows[0];
      if (r && /fcc\s*id/i.test(txt(r))) return tables[i];
    }
    return null;
  }

  function headerMap(table) {
    var head = table.rows[0];
    var map = {}, i, h;
    if (!head) return map;
    for (i = 0; i < head.cells.length; i++) {
      h = txt(head.cells[i]).toLowerCase();
      if (!h) continue;
      // "FinalActionDate" arrives with no spaces, hence the loose needles
      if (map.fcc_id      == null && /fcc\s*id/.test(h) && !/composite|related/.test(h)) map.fcc_id = i;
      if (map.company     == null && /applicant|grantee\s*name/.test(h))                 map.company = i;
      if (map.grant_date  == null && /grant\s*date|final\s*action|finalaction/.test(h))  map.grant_date = i;
      if (map.country     == null && /country/.test(h))                                  map.country = i;
      if (map.city        == null && /city/.test(h))                                     map.city = i;
      if (map.state       == null && /state/.test(h) && !/united/.test(h))               map.state = i;
      if (map.purpose     == null && /purpose/.test(h))                                  map.purpose = i;
    }
    if (map.fcc_id == null) map.fcc_id = FALLBACK_IDX.fcc_id;
    return map;
  }

  function cell(row, idx) {
    return (idx != null && idx < row.cells.length) ? txt(row.cells[idx]) : '';
  }

  function linksIn(row) {
    var out = { form731: '', exhibits: '' };
    var as = row.getElementsByTagName('a');
    for (var i = 0; i < as.length; i++) {
      var href = as[i].href || '';
      if (!out.form731  && /GetTcb731Report/i.test(href))   out.form731  = href;
      if (!out.exhibits && /ViewExhibitReport/i.test(href)) out.exhibits = href;
    }
    return out;
  }

  var FCC = {
    lastRows: 0,          // rows served by the most recent grab() — use for FromRec
    CLASS_IDS: CLASS_IDS,

    /* Begin a run. from/to are ISO YYYY-MM-DD — ingest names the weekly file
       from them, so mm/dd/yyyy here silently produces a broken path. */
    reset: function (fromISO, toISO) {
      isoToUS(fromISO); isoToUS(toISO);                    // validate both
      save({ window: { from: fromISO, to: toISO }, rows: [], ids: [] });
      this.lastRows = 0;
      return 'run started: ' + fromISO + ' .. ' + toISO;
    },

    /* Fill the page's REAL form and submit it (genuine navigation).
       opts: {cls, fromRec, status, purpose, showRecords} */
    submit: function (opts) {
      opts = opts || {};
      var s = need();
      var form = document.forms[0];
      if (!form) throw new Error('No form on this page — navigate to ' + SEARCH_URL + ' first.');
      var id = CLASS_IDS[opts.cls];
      if (!id) throw new Error('Unknown equipment class: ' + opts.cls);

      var fields = {
        grant_date_from:    isoToUS(s.window.from),
        grant_date_to:      isoToUS(s.window.to),
        application_status: opts.status  || 'GI',   // Grant Issued
        application_purpose:opts.purpose || 'O',    // Original grant
        equipment_class:    id,
        outputformat:       'HTML',                 // NOT XML — see below
        show_records:       String(opts.showRecords || 500),
        calledFromFrame:    'N',
        FromRec:            String(opts.fromRec || 1)
      };
      // outputformat=XML makes the browser DOWNLOAD the response instead of
      // navigating (net::ERR_ABORTED) — looks like a block but isn't.
      // Pagination is the hidden field FromRec; the form's own 'fetchfrom'
      // field is ignored by the server. See SCRAPER-ISSUES.md #5.
      for (var name in fields) {
        var el = form.elements[name];
        if (!el) {                                   // create missing hidden fields
          el = document.createElement('input');
          el.type = 'hidden'; el.name = name;
          form.appendChild(el);
        }
        el.value = fields[name];
      }
      form.submit();
      return 'submitted ' + opts.cls + ' FromRec=' + fields.FromRec;
    },

    /* Read the results page currently loaded, merge new US rows into the run.
       Pass the equipment class you queried — the table doesn't carry it. */
    grab: function (classCode) {
      var s = need();
      var table = resultsTable();
      if (!table) {
        this.lastRows = 0;
        var body = (document.body && document.body.innerText) || '';
        if (/no applications on file that match/i.test(body)) return { served: 0, added: 0, empty: true };
        throw new Error('No results table on this page (blocked, or search not run yet).');
      }
      var map = headerMap(table);
      var seen = {}, i;
      for (i = 0; i < s.ids.length; i++) seen[s.ids[i]] = 1;

      var served = 0, added = 0, nonUS = 0;
      for (i = 1; i < table.rows.length; i++) {           // row 0 = header
        var row = table.rows[i];
        var fid = cell(row, map.fcc_id).toUpperCase();
        if (!fid || /fcc\s*id/i.test(fid)) continue;      // repeated header row
        served++;
        var country = cell(row, map.country) || 'United States';
        if (!/^\s*(united states|usa|u\.?s\.?a?\.?)\s*$/i.test(country)) { nonUS++; continue; }
        if (seen[fid]) continue;                          // cross-class dedupe
        seen[fid] = 1; s.ids.push(fid);
        var l = linksIn(row);
        s.rows.push({
          company:        cell(row, map.company),
          city:           cell(row, map.city).replace(/,+$/, ''),
          state:          cell(row, map.state),
          country:        country,
          fcc_id:         fid,
          grant_date:     cell(row, map.grant_date),
          purpose:        cell(row, map.purpose),
          equipment_class: classCode || '',
          form731:        l.form731,
          exhibits:       l.exhibits
        });
        added++;
      }
      save(s);
      this.lastRows = served;
      // served === show_records means there is probably another page:
      // re-submit with fromRec = (previous FromRec + served).
      return { served: served, added: added, nonUS: nonUS, uniqueTotal: s.rows.length };
    },

    count:  function () { var s = store(); return s ? s.rows.length : 0; },
    status: function () { var s = store(); return s ? { window: s.window, unique: s.rows.length } : 'no run'; },

    /* The JSON to hand to:  weekly_leads.py ingest <file> */
    payload: function () { return JSON.stringify(need(), null, 1); },

    /* Form-731 URLs for the enrich step, in scrape order. */
    formUrls: function () { return need().rows.map(function (r) { return r.form731; }).filter(Boolean); },

    clear: function () { localStorage.removeItem(KEY); return 'cleared'; }
  };

  window.FCC = FCC;
  return 'FCC helper ready — FCC.reset(from,to) / FCC.submit({cls}) / FCC.grab(cls) / FCC.payload()';
})();
