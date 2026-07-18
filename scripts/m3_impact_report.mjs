#!/usr/bin/env node
/*
 * M3 Release Impact Analysis generator — Node.js port (no Python required).
 *
 * Cross-references an Infor M3 Release Report against an API Release Management
 * workbook (the APIs/components a customer actually uses) and produces a single
 * self-contained interactive HTML report with four views: Impact, By Component,
 * Release Landscape and Opportunities, plus a customer-facing dashboard.
 *
 * Usage:
 *   node m3_impact_report.mjs \
 *       --release-report  M3_Release_Report_*.xlsx \
 *       --api-release     API_Release_Mgt_M3_*.xlsx \
 *       [--cutoff 2026.06]   # default = current year.month
 *       [--out M3_Release_Impact_Analysis.html]
 *       [--emit-impacts impacts.json] [--summaries summaries.json]
 *       [--dashboard dash.html] [--no-dashboard]
 *
 * Matching rules (the core logic — keep these stable as the skill evolves):
 *   * A release item is considered only if it mentions a PROGRAM the customer uses.
 *   * If the item NAMES a transaction (attached PROG/Txn | PROG.Txn, or prose e.g.
 *     "transaction AddBatchLine in OIS100MI"), it is SPECIFIC: only components
 *     calling that exact transaction (and only transactions the customer actually
 *     uses) are flagged. If none are used -> not an impact.
 *   * If no transaction is named, it is GENERAL (program-level): every component
 *     on that program is flagged.
 *
 * Only third-party dependency is SheetJS ("xlsx"), used exactly like openpyxl was:
 * to read the two workbooks. Everything else is plain JS (regex, Map/Set, JSON,
 * string templating), so this runs anywhere Node runs.
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const require = createRequire(import.meta.url);
// Prefer the vendored single-file SheetJS (no npm install needed); fall back to an
// installed "xlsx" if someone removed the vendored copy.
let XLSX;
try {
  XLSX = require('./vendor/xlsx.mini.min.js');
} catch {
  XLSX = require('xlsx');
}

// ---------------------------------------------------------------------------
// Small helpers to mirror Python semantics exactly.
// ---------------------------------------------------------------------------

// Python str(x): None -> 'None'. openpyxl empty cells come through as null here.
const pystr = (x) => (x === null || x === undefined ? 'None' : String(x));

// Read a worksheet as an array-of-arrays (like openpyxl iter_rows(values_only=True)).
// Empty cells become null; short rows are left short (callers guard on length).
function sheetRows(ws) {
  if (!ws || !ws['!ref']) return [];
  return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true });
}

// ---------------------------------------------------------------------------
// parse_workbook_name
// ---------------------------------------------------------------------------
function parseWorkbookName(p) {
  let base = path.basename(p).replace(/\.xls[xm]?$/i, '');
  if (base.includes('~')) {
    const parts = base.split('~').map((s) => s.trim());
    let anchor = null;
    for (let i = 0; i < parts.length; i++) {
      if (/api.?release.?mgt/i.test(parts[i])) { anchor = i; break; }
    }
    if (anchor !== null && anchor >= 2) return [parts[anchor - 2], parts[anchor - 1]];
    if (parts.length >= 2) return [parts[0], parts[1]];
  }
  const m = base.match(/(.*?)_API[_ ]?Release[_ ]?Mgt/i);
  if (m) {
    const toks = m[1].split('_').filter(Boolean);
    const ENVS = new Set(['PRD', 'PROD', 'DEV', 'TST', 'TEST', 'TRN', 'TRAIN', 'DEM', 'DEMO', 'TRL',
      'TRIAL', 'PRO', 'ACC', 'UAT', 'EDU', 'SND', 'SAND', 'QA', 'STG', 'STAGE', 'DR']);
    for (let i = 1; i < toks.length - 1; i++) {
      if (ENVS.has(toks[i].toUpperCase())) {
        return [toks.slice(0, i + 1).join('_'), toks.slice(i + 1).join(' ')];
      }
    }
    if (toks.length >= 2) return [toks[0], toks.slice(1).join(' ')];
    if (toks.length) return [toks[0], ''];
  }
  return ['', ''];
}

// ---------------------------------------------------------------------------
// transaction prose detection — KEEP VERBS MAINTAINED (see SKILL.md)
// ---------------------------------------------------------------------------
const VERBS = 'Add|Chg|Del|Get|Lst|List|Upd|Sel|Cnv|Set|Mov|Cre|Rtv|Rel|Cancel|Close|' +
  'Open|Approve|Confirm|Calc|Check|Copy|Print|Send|Split|Trigger|Receive|' +
  'Import|Export|Validate|Generate|Retrieve|Update|Change|Create|Delete|' +
  'Select|Move|Release|Reset|Refresh|Load|Save|Run|Reverse|Connect|' +
  'Disconnect|Activate|Deactivate|Recalc|Reprice|Post|Settle|Allocate|Deallocate';
const ATTACHED_RE = new RegExp('\\b([A-Z]{2,4}\\d{3}MI)[\\./]([A-Za-z][A-Za-z0-9]+)', 'g');
const PROSE_TXN_RE = new RegExp('\\b(?:' + VERBS + ')[A-Z][A-Za-z0-9]+\\b', 'g');
const PROG_RE = new RegExp('\\b([A-Z]{2,4}\\d{3}MI)\\b', 'g');
const MI_PROG = /^[A-Z]{2,8}\d{0,4}MI$/;      // validates a parsed program code
const MI_SEARCH = /[A-Z]{2,8}\d{0,4}MI/;       // finds MI-program tokens inside a cell
const MONTH_NAMES = ['', 'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December'];
const KB_RE = /KB\d+/g;
const OPP_RE = /new API|new MI|Add MI Program|new transaction|new output field|new field|new input field|now possible|new option/i;

const findAll1 = (re, s) => [...s.matchAll(re)].map((m) => m[1]); // capture group 1
const findAll0 = (re, s) => [...s.matchAll(re)].map((m) => m[0]); // whole match

function moduleOf(comp) {
  const c = String(comp);
  if (c.includes('Finance') || c.includes('Financial Business Messages')) return 'Finance';
  if (c.includes('Sales Management') || c.includes('Sales Hub') || c.includes('Cash Desk')) return 'Sales';
  if (c.includes('Supply Chain Execution')) return 'Warehouse & Logistics';
  if (c.includes('Supply Chain Planning')) return 'Supply Chain Planning';
  if (c.includes('Procurement') || c.includes('Grower Contract')) return 'Procurement';
  if (c.includes('Maintenance')) return 'Maintenance / EAM';
  if (c.includes('Manufacturing')) return 'Manufacturing';
  if (c.includes('Product Data') || c.includes('Product Configurator') || c.includes('Attribute Control')) return 'Product Data';
  if (c.includes('Rental')) return 'Rental';
  if (c.includes('Project')) return 'Project';
  if (c.includes('Application Foundation') || c.includes('Job Scheduler')) return 'Foundation';
  if (c.startsWith('M3 Core Technology') || c.includes('Integrations and BODs') || c.includes('Portals') ||
      c.includes('Event Hub') || c.includes('Enterprise Collaborator') || c.includes('Data Lake') || c.includes('ACM')) {
    return 'Technology & Integration';
  }
  return 'Other / Add-ons';
}

// relkey -> comparable integer y*100+m, or null (mirrors tuple comparison).
function relkey(rel) {
  const m = String(rel).match(/^(\d{4})\.(\d{2})/);
  return m ? parseInt(m[1], 10) * 100 + parseInt(m[2], 10) : null;
}

function apiColumn(rows) {
  const width = rows.reduce((w, r) => Math.max(w, r.length), 0);
  let bestI = null, bestN = 0;
  for (let ci = 0; ci < width; ci++) {
    let n = 0;
    for (const r of rows) if (r.length > ci && r[ci] && MI_SEARCH.test(String(r[ci]))) n++;
    if (n > bestN) { bestI = ci; bestN = n; }
  }
  return bestN ? bestI : null;
}

// ---------------------------------------------------------------------------
// load_usage
// ---------------------------------------------------------------------------
function loadUsage(p) {
  const wb = XLSX.read(fs.readFileSync(p), { type: 'buffer', cellDates: false });
  const names = wb.SheetNames;

  let master = names.find((n) => n.trim().toLowerCase() === 'distinct apis') || null;
  if (master === null) master = names.find((n) => /distinct.*api/i.test(n)) || null;

  const progToTxns = new Map(); // prog -> Set(txn)
  const addTxn = (prog, t) => { if (!progToTxns.has(prog)) progToTxns.set(prog, new Set()); progToTxns.get(prog).add(t); };

  if (master) {
    const rows = sheetRows(wb.Sheets[master]);
    for (const r of rows.slice(1)) {
      if (!r || !r[0]) continue;
      const full = String(r[0]).trim();
      const prog = full.split('/')[0].trim();
      if (!MI_PROG.test(prog)) continue;
      const t = full.includes('/') ? full.split('/')[1].trim() : '';
      if (t) addTxn(prog, t);
    }
  }

  const compMap = new Map(); // prog -> [{type,name,api}]
  const compTypes = [];
  const skipped = [];
  for (const n of names) {
    if (n === master) continue;
    const all = sheetRows(wb.Sheets[n]);
    const sample = all.slice(0, 40);
    if (sample.length === 0) { skipped.push([n, 'empty']); continue; }
    const header = (sample[0] || []).map((c) => (c ? String(c).trim().toLowerCase() : ''));
    let apiCol = header.findIndex((h) => h.includes('api'));
    if (apiCol === -1) apiCol = null;
    if (apiCol === null) apiCol = apiColumn(sample.slice(1));
    if (apiCol === null) { skipped.push([n, 'no API/Transaction column found']); continue; }
    const nameCol = apiCol !== 0 ? 0 : 1;
    const stripVer = nameCol < header.length ? header[nameCol].includes('version') : false;
    compTypes.push(n);
    for (const r of all.slice(1)) {
      if (!r || r.length <= apiCol || !r[apiCol]) continue;
      const api = String(r[apiCol]).trim();
      const prog = api.split('/')[0].trim();
      if (!MI_PROG.test(prog)) continue;
      let name = (r.length > nameCol && r[nameCol]) ? String(r[nameCol]).trim() : api;
      if (stripVer) name = name.replace(/\s*\([^)]*\)\s*$/, '').trim() || name;
      if (!compMap.has(prog)) compMap.set(prog, []);
      compMap.get(prog).push({ type: n, name, api });
    }
  }

  // De-duplicate each program's component list by (type,name,api).
  for (const p2 of [...compMap.keys()]) {
    const seen = new Set(); const uniq = [];
    for (const c of compMap.get(p2)) {
      const k = c.type + '\u0000' + c.name + '\u0000' + c.api;
      if (seen.has(k)) continue;
      seen.add(k); uniq.push(c);
    }
    compMap.set(p2, uniq);
  }

  // Fallback: derive used set from components if there was no master list.
  if (progToTxns.size === 0) {
    for (const [p2, lst] of compMap) {
      for (const c of lst) {
        const t = c.api.includes('/') ? c.api.split('/')[1].trim() : '';
        if (t) addTxn(p2, t);
      }
    }
  }

  let usedApiCount = 0;
  for (const v of progToTxns.values()) usedApiCount += v.size;
  return { progToTxns, compMap, compTypes, usedApiCount, master, skipped };
}

// ---------------------------------------------------------------------------
// find_report_columns
// ---------------------------------------------------------------------------
const REPORT_HEADERS = {
  release: ['release'], component: ['component'], type: ['issue type', 'type'],
  summary: ['summary'], overview: ['overview'], detailed: ['detailed'],
  ref: ['reference'], expiry: ['expiration'],
};
const REPORT_FALLBACK = { release: 0, component: 2, type: 3, summary: 4, overview: 5, detailed: 6, ref: 16, expiry: 13 };

function findReportColumns(rows) {
  let headerRow = null, header = null;
  const limit = Math.min(15, rows.length);
  for (let i = 1; i <= limit; i++) {
    const r = rows[i - 1] || [];
    const cells = r.map((c) => (c ? String(c).trim().toLowerCase() : ''));
    if (cells.some((c) => c === 'release' || c.startsWith('release'))) { headerRow = i; header = cells; break; }
  }
  const cols = { ...REPORT_FALLBACK };
  if (header) {
    for (const [key, subs] of Object.entries(REPORT_HEADERS)) {
      for (let ci = 0; ci < header.length; ci++) {
        const cell = header[ci];
        if (cell && subs.some((s) => cell.includes(s))) { cols[key] = ci; break; }
      }
    }
  }
  return { headerRow: headerRow || 1, cols };
}

// ---------------------------------------------------------------------------
// build
// ---------------------------------------------------------------------------
function build(reportPath, progToTxns, compMap, cutoff) {
  const usedProgs = new Set(progToTxns.keys());
  const wb = XLSX.read(fs.readFileSync(reportPath), { type: 'buffer', cellDates: false });
  const sheetName = wb.SheetNames.includes('Report') ? 'Report' : wb.SheetNames[wb.SheetNames.length - 1];
  const rows = sheetRows(wb.Sheets[sheetName]);
  const { headerRow, cols: C } = findReportColumns(rows);

  const cell = (r, key) => {
    const i = C[key];
    return (i !== null && i !== undefined && r.length > i) ? r[i] : null;
  };

  // Dataset = Set A ∪ Set B, built before analysis (see SKILL.md).
  const items = [];
  for (let ri = headerRow; ri < rows.length; ri++) {
    const r = rows[ri];
    if (!r) continue;
    const rel = cell(r, 'release');
    if (rel === null || rel === 'Release' || String(rel).startsWith('\u00a9')) continue;
    const relk = relkey(rel);
    const expv = cell(r, 'expiry');
    const expk = (expv !== null && expv !== '') ? relkey(expv) : null;
    const toggleOpen = expk !== null && expk >= cutoff;
    const inSetA = (expk === null || expk < cutoff) && (relk !== null && relk >= cutoff);
    if (toggleOpen || inSetA) items.push(r);
  }

  const matches = [], landscape = [];
  for (const r of items) {
    const rel = cell(r, 'release'), comp = cell(r, 'component'), typ = cell(r, 'type');
    const summ = cell(r, 'summary'), ovr = cell(r, 'overview'), det = cell(r, 'detailed'), ref = cell(r, 'ref');
    const expv = cell(r, 'expiry');
    const expiry = (expv !== null && expv !== '' && String(expv).trim() !== 'None') ? String(expv).trim() : '';
    const overview = (ovr && String(ovr) !== 'None') ? String(ovr) : '';
    const detailed = (det && String(det) !== 'None') ? String(det) : '';
    const searchBlob = r.filter((c) => c && String(c) !== 'None').map((c) => String(c)).join(' ');
    const progs = new Set(findAll1(PROG_RE, searchBlob));
    const kbs = [...new Set(findAll0(KB_RE, detailed))].sort();
    const mod = moduleOf(comp);
    const isOpp = OPP_RE.test(searchBlob) && typ === 'Enhancement';
    const relk = relkey(rel);
    const major = !!(relk && (Math.floor(relk % 100) === 4 || Math.floor(relk % 100) === 10));

    const progsSorted = [...progs].sort();
    const usedProg = progsSorted.some((p) => usedProgs.has(p));
    landscape.push({
      release: rel, module: mod, component: String(comp), type: typ,
      summary: String(summ), overview, detailed, kbs, ref: String(ref), expiry, major,
      progs: progsSorted, used_prog: usedProg, opp: isOpp,
    });

    const hit = progsSorted.filter((p) => usedProgs.has(p));
    if (hit.length === 0) continue;

    const named = new Map(); // prog -> Set(txn)
    for (const m of searchBlob.matchAll(ATTACHED_RE)) {
      if (!named.has(m[1])) named.set(m[1], new Set());
      named.get(m[1]).add(m[2]);
    }
    const prose = new Set(findAll0(PROSE_TXN_RE, searchBlob).filter((t) => t.length >= 5));
    const specific = named.size > 0 || prose.size > 0;

    let impacted = []; const usedTxns = new Set();
    let scope;
    if (specific) {
      for (const p of hit) {
        const cand = new Set([...(named.get(p) || new Set()), ...prose]);
        const ptx = progToTxns.get(p) || new Set();
        for (const t of cand) if (ptx.has(t)) usedTxns.add(`${p}/${t}`);
        for (const c of (compMap.get(p) || [])) {
          if (usedTxns.has(c.api)) impacted.push({ ...c, exact: true });
        }
      }
      scope = 'specific';
    } else {
      for (const p of hit) {
        for (const c of (compMap.get(p) || [])) impacted.push({ ...c, exact: false });
      }
      scope = 'general';
    }

    const seen = new Set(); const ded = [];
    for (const c of impacted) {
      const key = c.type + '\u0000' + c.name + '\u0000' + c.api;
      if (seen.has(key)) continue;
      seen.add(key); ded.push(c);
    }
    if (scope === 'specific' && ded.length === 0) continue; // names unused txns -> landscape only

    let sev;
    if (usedTxns.size && typ === 'Defect') sev = 'High';
    else if (usedTxns.size) sev = 'Medium';
    else if (typ === 'Defect') sev = 'Medium';
    else sev = 'Low';

    matches.push({
      release: rel, module: mod, component: String(comp), type: typ,
      summary: String(summ), overview, detailed, kbs, ref: String(ref), expiry, major,
      progs: hit.slice().sort(),
      txns: [...usedTxns].sort(), scope, severity: sev,
      affected: ded, n_components: ded.length,
    });
  }

  return { matches, landscape, total: items.length };
}

// ---------------------------------------------------------------------------
// HTML rendering
// ---------------------------------------------------------------------------
const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const PALETTE = ['#4263eb', '#7048e8', '#0ca678', '#f08c00', '#e8590c', '#1098ad', '#ae3ec9', '#37b24d'];

const readTemplate = (name) => fs.readFileSync(path.join(__dirname, '..', 'templates', name), 'utf8');
// split/join avoids $-substitution and replaces the single occurrence safely.
const inject = (tpl, token, value) => tpl.split(token).join(value);

function renderHtml(matches, landscape, totalFiltered, usedApiCount, usedProgCount, cutoffStr, compTypes, tenant, customer) {
  let monthLabel;
  try {
    const mm = parseInt(cutoffStr.split('.')[1], 10); const yy = cutoffStr.split('.')[0];
    monthLabel = `${MONTH_NAMES[mm]} ${yy}`;
  } catch { monthLabel = cutoffStr; }
  const payload = JSON.stringify({
    generated: new Date().toISOString().slice(0, 10), cutoff: cutoffStr,
    month_label: monthLabel, tenant, customer,
    total_filtered: totalFiltered, used_api_count: usedApiCount,
    used_prog_count: usedProgCount, comp_types: compTypes || [],
    matches, landscape,
  });
  return inject(readTemplate('analyst.html'), '__PAYLOAD__', payload);
}

// Mirror Python's str(float): integer-valued floats print with a trailing ".0".
const pyfloat = (x) => (Number.isInteger(x) ? `${x}.0` : String(x));
// Mirror Python f"{x:.2f}": preserves negative zero as "-0.00".
const py2f = (x) => (Object.is(x, -0) ? '-0.00' : x.toFixed(2));

function donut(segments, centerNum, centerLabel, size = 168, stroke = 26) {
  const total = segments.reduce((s, seg) => s + seg[1], 0) || 1;
  const r = pyfloat((size - stroke) / 2);
  const cxN = size / 2, cyN = size / 2;
  const cx = pyfloat(cxN), cy = pyfloat(cyN);
  const circ = 2 * Math.PI * ((size - stroke) / 2);
  const parts = [`<circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="#eef1f5" stroke-width="${stroke}"/>`];
  let offset = 0.0;
  for (const [, value, color] of segments) {
    if (value <= 0) continue;
    const dash = (value / total) * circ;
    parts.push(`<circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="${color}" stroke-width="${stroke}" ` +
      `stroke-dasharray="${py2f(dash)} ${py2f(circ - dash)}" stroke-dashoffset="${py2f(-offset)}" ` +
      `transform="rotate(-90 ${cx} ${cy})"/>`);
    offset += dash;
  }
  parts.push(`<text x="${cx}" y="${pyfloat(cyN - 1)}" text-anchor="middle" class="dnum">${centerNum}</text>`);
  parts.push(`<text x="${cx}" y="${pyfloat(cyN + 17)}" text-anchor="middle" class="dlbl">${centerLabel}</text>`);
  return `<svg viewBox="0 0 ${size} ${size}" width="${size}" height="${size}">${parts.join('')}</svg>`;
}

function legend(segments) {
  const rows = segments.filter(([, v]) => v > 0)
    .map(([label, v, c]) => `<div class="lg"><span class="dot" style="background:${c}"></span>${esc(label)}<b>${v}</b></div>`)
    .join('');
  return `<div class="legend">${rows}</div>`;
}

function chartCard(title, segments, centerNum, centerLabel) {
  return `<div class="panel chart"><h2>${title}</h2>` +
    `<div class="chartwrap">${donut(segments, centerNum, centerLabel)}${legend(segments)}</div></div>`;
}

function renderDashboard(matches, landscape, cutoff, cutoffStr, usedApiCount, usedProgCount, tenant, customer) {
  const cur = landscape.filter((x) => relkey(x.release) === cutoff);
  let imp = matches.filter((m) => relkey(m.release) === cutoff && m.scope === 'specific');
  const sevRank = { High: 0, Medium: 1, Low: 2 };
  imp.sort((a, b) => sevRank[a.severity] - sevRank[b.severity]);
  imp = imp.map((m, i) => [m, i]).sort((a, b) => (a[0].release < b[0].release ? -1 : a[0].release > b[0].release ? 1 : a[1] - b[1])).map((x) => x[0]);
  const cutM = Math.floor(cutoff % 100);
  const isMajor = cutM === 4 || cutM === 10;
  const monthLabel = `${MONTH_NAMES[cutM]} ${Math.floor(cutoff / 100)}`;
  const versions = [...new Set(cur.map((x) => x.release))].sort();
  const defects = cur.filter((x) => x.type === 'Defect').length;
  const enh = cur.length - defects;
  const compsSet = new Set();
  for (const m of imp) for (const c of m.affected) compsSet.add(c.type + '\u0000' + c.name);
  const sevCounts = { High: 0, Medium: 0, Low: 0 };
  for (const m of imp) sevCounts[m.severity] += 1;

  const kpi = (n, label, cls = '') => `<div class="kpi ${cls}"><div class="n">${n}</div><div class="l">${label}</div></div>`;
  const kpis = [
    kpi(cur.length, 'Release items this month'),
    kpi(imp.length, 'Affecting your applications', 'accent'),
    kpi(compsSet.size, 'Components affected'),
    kpi(defects, 'Defects fixed'),
    kpi(enh, 'Enhancements'),
  ].join('');

  const sevSeg = [['High', sevCounts.High, '#e03131'], ['Medium', sevCounts.Medium, '#f08c00'], ['Low', sevCounts.Low, '#2f9e44']];
  const typeSeg = [['Defects', defects, '#fa5252'], ['Enhancements', enh, '#4263eb']];
  const modCounts = new Map();
  for (const x of cur) modCounts.set(x.module, (modCounts.get(x.module) || 0) + 1);
  const ordered = [...modCounts.entries()].sort((a, b) => b[1] - a[1]);
  const modSeg = ordered.slice(0, 6).map(([m, c], i) => [m, c, PALETTE[i % PALETTE.length]]);
  const other = ordered.slice(6).reduce((s, [, c]) => s + c, 0);
  if (other) modSeg.push(['Other', other, '#adb5bd']);

  const charts = chartCard('Your exposure by severity', sevSeg, imp.length, 'impacts') +
    chartCard('Change type this month', typeSeg, cur.length, 'items') +
    chartCard('Where Infor changed', modSeg, cur.length, 'items');

  let cards = '';
  for (const m of imp) {
    const ctypeCounts = new Map();
    for (const c of m.affected) ctypeCounts.set(c.type, (ctypeCounts.get(c.type) || 0) + 1);
    const ctags = [...ctypeCounts.entries()].map(([t, n]) => `<span class="ctag">${esc(t)} \u00d7${n}</span>`).join(' ');
    const isum = m.impact_summary ? `<div class="isum">${esc(m.impact_summary)}</div>` : '';
    const exp = m.expiry ? `<span class="exp">\u23fb toggle expires ${m.expiry}</span>` : '';
    const kbtags = (m.kbs || []).map((k) => `<span class="kb">${esc(k)}</span>`).join(' ');
    const ref = (m.ref && m.ref !== 'None') ? `<span class="refbadge">${esc(m.ref)}</span>` : '';
    const n = m.affected.length;
    const colsN = n <= 6 ? 1 : (n <= 20 ? 2 : 3);
    const CAP = 60;
    const shownAff = m.affected.slice(0, CAP);
    let itemsHtml = shownAff.map((c) =>
      `<div class="affitem ${c.exact ? 'exact' : ''}">` +
      `<div class="affapi">${esc(c.api)}</div>` +
      `<div class="affsub"><span class="afftype">${esc(c.type)}</span>` +
      `<span class="affname">${esc(c.name)}</span></div></div>`).join('');
    if (n > CAP) itemsHtml += `<div class="affmore">+ ${n - CAP} more not shown (${n} total)</div>`;
    const afflist = itemsHtml
      ? `<details class="affwrap"><summary>Impacted components &amp; transactions (${n})</summary>` +
        `<div class="afflist c${colsN}">${itemsHtml}</div></details>`
      : '';
    cards +=
      `<div class="card ${m.severity}"><div class="chead">` +
      `<span class="rel">${m.release}</span>${ref}${kbtags}` +
      `<span class="sev ${m.severity}">${m.severity}</span>` +
      `<span class="typ">${m.type}</span>${exp}` +
      `<span class="mod">${esc(m.module)}</span></div>` +
      `<div class="ctitle">${esc(m.summary)}</div>${isum}` +
      `<div class="ctags">Affects ${m.affected.length} component(s): ${ctags}</div>${afflist}</div>`;
  }

  const kindNote = isMajor
    ? 'This is a <b>major release</b> \u2014 April and October carry Infor CloudSuite\u2019s major feature releases, so expect broader change.'
    : 'This is a <b>minor (monthly) release</b>. Infor CloudSuite\u2019s major feature releases ship in April and October.';
  const pct = Math.round(100 * imp.length / Math.max(cur.length, 1));

  let idtags = '';
  if (customer) idtags += `<span class="idtag"><span class="lbl">Prepared for</span><b>${esc(customer)}</b></span>`;
  if (tenant) idtags += `<span class="idtag"><span class="lbl">Tenant</span><b>${esc(tenant)}</b></span>`;
  const custline = idtags ? `<div class="herometa">${idtags}</div>` : '';

  const fields = {
    month: monthLabel, kind_note: kindNote, kpis, charts, pct: String(pct),
    n_imp: String(imp.length), n_cur: String(cur.length), nver: String(versions.length), custline,
    footprint: `${usedApiCount} transactions across ${usedProgCount} programs`,
    cards: cards || '<div class="muted">No items affecting your applications this month.</div>',
    generated: new Date().toISOString().slice(0, 10),
  };
  return pyformat(readTemplate('dashboard.html'), fields);
}

// Mirror Python str.format on the dashboard template: replace {field} tokens, then
// collapse the doubled CSS braces ({{ -> {, }} -> }). Field values contain no braces.
function pyformat(tpl, fields) {
  let out = tpl;
  for (const [k, v] of Object.entries(fields)) out = out.split('{' + k + '}').join(v);
  out = out.split('{{').join('{').split('}}').join('}');
  return out;
}

// ---------------------------------------------------------------------------
// emit_impacts
// ---------------------------------------------------------------------------
function emitImpacts(matches, p) {
  const out = matches.map((m) => ({
    ref: m.ref, release: m.release, severity: m.severity, type: m.type,
    component: m.component, scope: m.scope, expiry: m.expiry || '',
    summary: m.summary, overview: m.overview, detailed: m.detailed,
    progs: m.progs, txns: m.txns,
    affected: m.affected.map((c) => ({ type: c.type, api: c.api, name: c.name })),
  }));
  fs.writeFileSync(p, JSON.stringify(out, null, 2));
}

// ---------------------------------------------------------------------------
// CLI
// ---------------------------------------------------------------------------
function parseArgs(argv) {
  const a = { out: 'M3_Release_Impact_Analysis.html', no_dashboard: false };
  const map = {
    '--release-report': 'release_report', '--api-release': 'api_release', '--cutoff': 'cutoff',
    '--out': 'out', '--emit-impacts': 'emit_impacts', '--summaries': 'summaries', '--dashboard': 'dashboard',
  };
  for (let i = 0; i < argv.length; i++) {
    const t = argv[i];
    if (t === '--no-dashboard') { a.no_dashboard = true; continue; }
    if (map[t]) { a[map[t]] = argv[++i]; continue; }
    throw new Error(`unknown argument: ${t}`);
  }
  if (!a.release_report || !a.api_release) {
    throw new Error('required: --release-report <xlsx> --api-release <xlsx>');
  }
  return a;
}

function main() {
  const a = parseArgs(process.argv.slice(2));
  let cutoff, cutoffStr;
  if (a.cutoff) {
    const m = a.cutoff.match(/(\d{4})\.(\d{2})/);
    cutoff = parseInt(m[1], 10) * 100 + parseInt(m[2], 10);
    cutoffStr = a.cutoff;
  } else {
    const t = new Date();
    cutoff = t.getFullYear() * 100 + (t.getMonth() + 1);
    cutoffStr = `${t.getFullYear()}.${String(t.getMonth() + 1).padStart(2, '0')}`;
  }

  const { progToTxns, compMap, compTypes, usedApiCount, master, skipped } = loadUsage(a.api_release);
  const [tenant, customer] = parseWorkbookName(a.api_release);
  if (tenant || customer) console.log(`workbook identity: customer '${customer}' | tenant '${tenant}'`);
  console.log(`API workbook: master '${master === null ? 'None' : master}' | component sheets (${compTypes.length}): [${compTypes.map((s) => `'${s}'`).join(', ')}]`);
  if (skipped.length) {
    console.log('  NOTE — sheets NOT included (no API/Transaction column): ' +
      skipped.map(([n, why]) => `'${n}' (${why})`).join(', ') +
      ". If any of these should be analyzed, add an 'API / Transaction' column header.");
  }

  const { matches, landscape, total } = build(a.release_report, progToTxns, compMap, cutoff);

  if (a.emit_impacts) {
    emitImpacts(matches, a.emit_impacts);
    console.log(`cutoff ${cutoffStr} | impacts ${matches.length} -> ${a.emit_impacts}`);
  }

  const summaries = a.summaries ? JSON.parse(fs.readFileSync(a.summaries, 'utf8')) : {};
  for (const mt of matches) {
    const s = summaries[String(mt.ref)] ?? summaries[String(mt.summary)];
    if (s) mt.impact_summary = s;
  }

  const html = renderHtml(matches, landscape, total, usedApiCount, progToTxns.size, cutoffStr, compTypes, tenant, customer);
  fs.writeFileSync(a.out, html);
  const nSum = matches.filter((m) => m.impact_summary).length;
  const nOpp = landscape.filter((l) => l.opp).length;
  console.log(`cutoff ${cutoffStr} | scanned ${total} | impacts ${matches.length} (${nSum} summarised) ` +
    `| landscape ${landscape.length} | opportunities ${nOpp} | component types [${compTypes.map((s) => `'${s}'`).join(', ')}]`);
  console.log(`-> ${a.out}`);

  if (!a.no_dashboard) {
    const dashPath = a.dashboard || a.out.replace(/\.html?$/, '') + '_customer_dashboard.html';
    const dash = renderDashboard(matches, landscape, cutoff, cutoffStr, usedApiCount, progToTxns.size, tenant, customer);
    fs.writeFileSync(dashPath, dash);
    const curN = landscape.filter((l) => relkey(l.release) === cutoff).length;
    const curImp = matches.filter((m) => relkey(m.release) === cutoff && m.scope === 'specific').length;
    console.log(`customer dashboard (${cutoffStr}): ${curN} month items, ${curImp} specific impacts shown -> ${dashPath}`);
  }
}

main();
