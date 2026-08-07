/* ============================================================
   AMORE Dashboard — Google Sheets Data Loader
   ------------------------------------------------------------
   Fetches live data from spreadsheet
   "[LIVE] Dashboard AMORE Nasional - For MD"
   and exposes it in the exact shape expected by index.html.
   ============================================================ */

(function (global) {
  'use strict';

  const SPREADSHEET_ID = '1uU3mG1tu-d9aW1Z-eL-ocOQHhMzg2xPWAN4DvSYmo1o';

  const SHEETS = {
    dashboard: 'Dashboard',
    dataMentah: 'Data Mentah',
    rekap: 'Rekap Result All MD',
  };

  const MONTHS_SHORT = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const PROC_LABELS_INTERNAL = ['Overall','Flinkote','Regulator','Timer','MarkingSelang','MarkingFB'];
  const PROC_LABELS_DISPLAY = ['Flinkote','Regulator','Timer','Marking Selang','Marking FB'];

  const PROC_DATA_COL_INDEX = {
    'Flinkote':       3,
    'Regulator':      4,
    'Timer':          5,
    'Marking Selang': 6,
    'Marking FB':     7,
  };

  const OK_CODE_PREFIX = {
    'Flinkote':       '1',
    'Regulator':      '23',
    'Timer':          '1',
    'Marking Selang': '1',
    'Marking FB':     '1',
  };

  // MD-to-Region fallback map used if the raw-data sheet has no row for that MD yet
  // (e.g. very early in the month before any report comes in). Kept in sync with
  // the "Cek Region" tab of the source workbook.
  const MD_REGION_FALLBACK = {
    B10:'REGION 1', B3Z:'REGION 1', C10:'REGION 1', C3Z:'REGION 1',
    D2Z:'REGION 1', D3Z:'REGION 1', E20:'REGION 1', G01:'REGION 1',
    G02:'REGION 1', G5Z:'REGION 1', H2Z:'REGION 1',
    I01:'REGION 2', I3Z:'REGION 2', J10:'REGION 2', J20:'REGION 2',
    K0Z:'REGION 2', L01:'REGION 2', M2Z:'REGION 2', M3Z:'REGION 2',
    N01:'REGION 2', N02:'REGION 2',
    Q01:'REGION 3', R4Z:'REGION 3', R5Z:'REGION 3', T10:'REGION 3',
    U10:'REGION 3', V2Z:'REGION 3', W01:'REGION 3', Z11:'REGION 3',
  };

  function csvUrl(sheetName) {
    const enc = encodeURIComponent(sheetName);
    // headers=0 stops gviz from auto-treating the first "data-like" row as column labels,
    // so metadata rows above the real header (Periode Data, Last Update, ...) don't shift
    // and the real header ("Timestamp", "MD") comes through as an ordinary data cell.
    return `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/gviz/tq?tqx=out:csv&headers=0&sheet=${enc}&t=${Date.now()}`;
  }

  async function fetchSheet(sheetName) {
    const res = await fetch(csvUrl(sheetName), { cache: 'no-store' });
    if (!res.ok) {
      throw new Error(`Gagal fetch sheet "${sheetName}" (HTTP ${res.status}). Pastikan spreadsheet di-share sebagai "Anyone with the link can view".`);
    }
    const text = await res.text();
    if (text.startsWith('<')) {
      throw new Error(`Sheet "${sheetName}" tidak dapat diakses. Pastikan spreadsheet di-share sebagai "Anyone with the link can view" dan tab "${sheetName}" ada.`);
    }
    // gviz sometimes wraps the CSV in a JSONP callback if it thinks you asked for JSON.
    // Detect and unwrap "google.visualization.Query.setResponse({...})"
    if (/^\s*google\.visualization/.test(text)) {
      throw new Error(`Sheet "${sheetName}" dikembalikan sebagai JSONP, bukan CSV. Cek URL endpoint gviz.`);
    }
    const rows = parseCSV(text);
    // Enable verbose fetch logging by running `AmoreData.debug = true` in the browser console.
    if (global.AmoreData && global.AmoreData.debug) {
      try {
        console.log(`[AmoreData] sheet "${sheetName}": ${rows.length} rows, first row =`, rows[0]);
      } catch (_) {}
    }
    return rows;
  }

  // Normalize a cell: strip surrounding whitespace, replace non-breaking spaces, lowercase optional.
  function normCell(s) {
    if (s == null) return '';
    return String(s).replace(/ /g, ' ').trim();
  }
  function normLower(s) {
    return normCell(s).toLowerCase();
  }
  // Preview the first N rows for debug messages.
  function previewRows(rows, n) {
    return rows.slice(0, n).map((r, i) => `[${i}] ${r.slice(0, 4).map(c => JSON.stringify(normCell(c))).join(' | ')}`).join('\n');
  }

  function parseCSV(text) {
    const rows = [];
    let cur = [];
    let field = '';
    let inQ = false;
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (inQ) {
        if (ch === '"') {
          if (text[i + 1] === '"') { field += '"'; i++; }
          else { inQ = false; }
        } else {
          field += ch;
        }
      } else {
        if (ch === '"') { inQ = true; }
        else if (ch === ',') { cur.push(field); field = ''; }
        else if (ch === '\n') { cur.push(field); rows.push(cur); cur = []; field = ''; }
        else if (ch === '\r') { /* skip */ }
        else { field += ch; }
      }
    }
    if (field.length || cur.length) { cur.push(field); rows.push(cur); }
    return rows;
  }

  function parsePct(str) {
    if (str == null) return null;
    const s = String(str).trim();
    if (!s || s === '-' || s === '#DIV/0!' || s === '#N/A') return null;
    const n = parseFloat(s.replace('%', '').replace(',', '.'));
    return isFinite(n) ? n : null;
  }
  function parseInt0(str) {
    if (str == null) return 0;
    const s = String(str).trim();
    if (!s || s === '-') return 0;
    const n = parseInt(s.replace(/[^\d-]/g, ''), 10);
    return isFinite(n) ? n : 0;
  }

  /* ------------------------------------------------------------
     Parse "Dashboard" sheet → md rows, region/nasional aggregates,
     mdMeta (LCR & Sample), and period info.
     ------------------------------------------------------------ */
  function parseDashboard(rows) {
    // Find header row: first cell === "MD" (case-insensitive, whitespace-tolerant)
    let headerIdx = -1;
    for (let i = 0; i < rows.length; i++) {
      if (normLower(rows[i][0]) === 'md') { headerIdx = i; break; }
    }
    if (headerIdx < 0) throw new Error('Header row "MD" tidak ditemukan di sheet Dashboard. Preview:\n' + previewRows(rows, 8));

    // Find period (row above with month info) and last update
    let period = '';
    let lastUpdate = '';
    for (let i = 0; i < headerIdx; i++) {
      const c0 = (rows[i][0] || '').trim();
      const c1 = (rows[i][1] || '').trim();
      if (!period && c1 && /\d{4}/.test(c1)) period = c1;
      if (c0.toLowerCase().startsWith('last update') && c1) lastUpdate = c1;
    }

    // Data rows follow header until blank MD or Region/Nasional row
    const md = [];
    const mdMeta = {};
    const regionAggNow = {};
    let nasionalNow = [0,0,0,0,0,0];

    for (let i = headerIdx + 1; i < rows.length; i++) {
      const r = rows[i];
      const name = (r[0] || '').trim();
      if (!name) continue;
      const lcr = parseInt0(r[1]);
      const sample = parseInt0(r[2]);
      const overall  = parsePct(r[4]);
      const flin     = parsePct(r[7]);
      const regul    = parsePct(r[10]);
      const timer    = parsePct(r[13]);
      const mSelang  = parsePct(r[16]);
      const mFB      = parsePct(r[19]);
      const vals = [overall, flin, regul, timer, mSelang, mFB].map(v => v == null ? 0 : v);

      if (/^Region\s*\d/i.test(name)) {
        const key = 'REGION ' + name.match(/\d/)[0];
        regionAggNow[key] = vals;
        mdMeta[key] = { lcr, sample };
      } else if (/^Nasional/i.test(name)) {
        nasionalNow = vals;
        mdMeta['Nasional'] = { lcr, sample };
      } else if (/^[A-Z]\d{1,2}[A-Z0-9]?\s*-\s*/i.test(name)) {
        // "B10 - PT Indako Trading Co."
        const m = name.match(/^([A-Z0-9]{3})\s*-\s*(.+)$/i);
        if (!m) continue;
        const code = m[1].toUpperCase();
        const mdName = m[2].trim();
        md.push([code, mdName, null, overall || 0, flin || 0, regul || 0, timer || 0, mSelang || 0, mFB || 0]);
        mdMeta[`${code} - ${mdName}`] = { lcr, sample };
      }
    }

    return { md, mdMeta, regionAggNow, nasionalNow, period, lastUpdate };
  }

  /* ------------------------------------------------------------
     Parse "Rekap Result All MD" — 6 stacked sections
     (Overall, Flinkote, Regulator, Timer, MarkingSelang, MarkingFB),
     each with monthly values Jan..Dec + Gr M vs M-1 column.
     Output shape: { "<md/region/nasional>": { Overall:[], Flinkote:[], ... } }
     ------------------------------------------------------------ */
  function parseRekap(rows) {
    const sectionOrder = ['Overall','Flinkote','Regulator','Timer','MarkingSelang','MarkingFB'];
    const sectionNames = {
      'OVERALL PROCESS': 'Overall',
      'OVERALL': 'Overall',
      'FLINKOTE': 'Flinkote',
      'REGULATOR': 'Regulator',
      'TIMER': 'Timer',
      'MARKING SELANG': 'MarkingSelang',
      'MARKING FB': 'MarkingFB',
    };

    // Find every "MD | Jan-26 | Feb-26 | ..." header. Each precedes a data block.
    const sections = []; // {sectionKey, headerIdx, monthCount}
    let currentSectionKey = 'Overall'; // first section has no explicit label above
    let sawFirstHeader = false;

    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      const first = (r[0] || '').trim();
      const upper = first.toUpperCase();
      // Section labels between blocks
      if (sectionNames[upper]) {
        currentSectionKey = sectionNames[upper];
      }
      // Header row
      if (first === 'MD' && (r[1] || '').toLowerCase().startsWith('jan')) {
        // Count month columns
        let monthCount = 0;
        for (let c = 1; c < r.length; c++) {
          const v = (r[c] || '').trim();
          if (/^[A-Z][a-z]{2}-\d{2}$/.test(v)) monthCount++;
          else break;
        }
        sections.push({ key: sawFirstHeader ? currentSectionKey : 'Overall', headerIdx: i, monthCount });
        sawFirstHeader = true;
      }
    }

    const out = {};
    function ensure(entityKey) {
      if (!out[entityKey]) out[entityKey] = { Overall:[], Flinkote:[], Regulator:[], Timer:[], MarkingSelang:[], MarkingFB:[] };
      return out[entityKey];
    }

    sections.forEach(sec => {
      const monthCount = sec.monthCount;
      // Data rows follow headerIdx until blank / non-entity row
      for (let i = sec.headerIdx + 1; i < rows.length; i++) {
        const r = rows[i];
        const name = (r[0] || '').trim();
        if (!name) break;
        // Stop if this is a section label (uppercase word like "FLINKOTE")
        if (sectionNames[name.toUpperCase()]) break;
        // Also stop if this is a new "MD" header
        if (name === 'MD') break;

        const vals = [];
        for (let c = 1; c <= monthCount; c++) {
          vals.push(parsePct(r[c]));
        }
        // Trim trailing nulls
        while (vals.length && vals[vals.length - 1] == null) vals.pop();
        if (!vals.length) continue;

        let entityKey = name;
        // Normalize "Region 1" / "Nasional" as-is; MDs are "B10 - PT Indako..." matches hardcoded shape.
        ensure(entityKey)[sec.key] = vals.map(v => v == null ? 0 : v);
      }
    });

    return out;
  }

  /* ------------------------------------------------------------
     Parse "Data Mentah" — raw report rows.
     Build:
       - dataMentah = { headers:[...], rows:[[...]] } for XLSX download
       - detailProblemData = per-MD × per-process counts of OK/NG + problem codes
     ------------------------------------------------------------ */
  function parseDataMentah(rows) {
    // Find header row. gviz auto-types column A as DateTime because it's full of
    // timestamps, then drops every non-date value in that column — including the
    // literal header cell "Timestamp" and the metadata rows above it. So col 0 of
    // the header row is empty at fetch time. We anchor detection on "Main Dealer"
    // (col 1) instead, then re-insert "Timestamp" as the col 0 header.
    let headerIdx = -1;
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i] || [];
      if (normLower(r[0]).startsWith('timestamp')) { headerIdx = i; break; }
      if (normLower(r[1]) === 'main dealer') { headerIdx = i; break; }
    }
    if (headerIdx < 0) throw new Error('Header row Data Mentah tidak ditemukan (mencari "Timestamp" atau "Main Dealer"). Preview:\n' + previewRows(rows, 8));

    const headers = rows[headerIdx].map(s => normCell(s));
    if (!headers[0]) headers[0] = 'Timestamp';
    // Truncate any trailing empty header columns
    while (headers.length && headers[headers.length - 1] === '') headers.pop();

    const dataRows = [];
    for (let i = headerIdx + 1; i < rows.length; i++) {
      const r = rows[i];
      const ts = (r[0] || '').trim();
      const md = (r[1] || '').trim();
      if (!ts && !md) continue;
      // Keep only rows with a Main Dealer
      if (!md) continue;
      const trimmed = r.slice(0, headers.length).map(v => (v == null ? '' : v));
      dataRows.push(trimmed);
    }

    // Build detailProblemData
    const detailProblemData = {};
    const codeMdRegion = {};

    dataRows.forEach(r => {
      const mdKey = (r[1] || '').trim();
      const regionRaw = (r[headers.length - 1] || '').trim(); // "Region" column is last
      if (!mdKey) return;

      if (!detailProblemData[mdKey]) {
        detailProblemData[mdKey] = { region: normalizeRegion(regionRaw), processes: {} };
        PROC_LABELS_DISPLAY.forEach(p => {
          detailProblemData[mdKey].processes[p] = { sample: 0, ok: 0, ng: 0, problems: {} };
        });
      } else if (!detailProblemData[mdKey].region) {
        detailProblemData[mdKey].region = normalizeRegion(regionRaw);
      }

      PROC_LABELS_DISPLAY.forEach(procLabel => {
        const colIdx = PROC_DATA_COL_INDEX[procLabel];
        const rawVal = (r[colIdx] || '').trim();
        if (!rawVal) return;
        const bucket = detailProblemData[mdKey].processes[procLabel];
        bucket.sample += 1;
        const codeMatch = rawVal.match(/^(\d+)\s*-\s*/);
        const code = codeMatch ? codeMatch[1] : null;
        const isOK = code === OK_CODE_PREFIX[procLabel];
        if (isOK) {
          bucket.ok += 1;
        } else {
          bucket.ng += 1;
          bucket.problems[rawVal] = (bucket.problems[rawVal] || 0) + 1;
        }
      });

      const code = mdKey.split(' - ')[0];
      codeMdRegion[code] = normalizeRegion(regionRaw);
    });

    return { dataMentah: { headers, rows: dataRows }, detailProblemData, codeMdRegion };
  }

  function normalizeRegion(s) {
    if (!s) return null;
    const t = s.toUpperCase();
    if (t.includes('REGION 1')) return 'REGION 1';
    if (t.includes('REGION 2')) return 'REGION 2';
    if (t.includes('REGION 3')) return 'REGION 3';
    return null;
  }

  /* ------------------------------------------------------------
     Master loader.
     ------------------------------------------------------------ */
  async function loadAll() {
    const [dashRows, rekapRows, dmRows] = await Promise.all([
      fetchSheet(SHEETS.dashboard),
      fetchSheet(SHEETS.rekap),
      fetchSheet(SHEETS.dataMentah),
    ]);

    const dash = parseDashboard(dashRows);
    const rekapAllMD = parseRekap(rekapRows);
    const dm = parseDataMentah(dmRows);

    // Attach region info: prefer Data Mentah (authoritative), fall back to hardcoded map
    dash.md.forEach(row => {
      const code = row[0];
      row[2] = dm.codeMdRegion[code] || MD_REGION_FALLBACK[code] || null;
    });
    const md = dash.md.filter(r => r[2]);

    return {
      md,
      mdMeta: dash.mdMeta,
      regionAggNow: dash.regionAggNow,
      nasionalNow: dash.nasionalNow,
      detailProblemData: dm.detailProblemData,
      rekapAllMD,
      dataMentah: dm.dataMentah,
      period: dash.period,
      lastUpdate: dash.lastUpdate,
      fetchedAt: new Date(),
    };
  }

  global.AmoreData = { loadAll, SPREADSHEET_ID };
})(window);
