/**
 * ============================================================
 *  InsuranceDekho Partner Dashboard — Google Apps Script
 * ============================================================
 *
 *  HOW TO USE:
 *  1. Open your Google Sheet
 *  2. Extensions → Apps Script
 *  3. Delete ALL existing code, paste this entire file
 *  4. Click Save (Ctrl+S)
 *  5. Run → testGetData  (approve permissions when asked)
 *  6. Deploy → New Deployment → Web App
 *     • Execute as : Me
 *     • Who has access : Anyone
 *  7. Copy the /exec URL → paste into index.html (see comment there)
 */

// ✅ ALREADY SET TO YOUR CORRECT SHEET NAME
const SHEET_NAME = "master";

// Month columns — must match your sheet headers exactly
const MONTH_COLS = [
  "April'2025","May'25","June'25","July'25","August'25",
  "September'25","October'25","November'25","December'25",
  "January'26","Feburary'26","March'26","April'26"
];

// ─────────────────────────────────────────────
function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || "getData";
    if (action === "getData")    return sendJSON(getPartnerData());
    if (action === "getSummary") return sendJSON(getSummaryStats());
    if (action === "ping")       return sendJSON({ status:"ok", time: new Date().toISOString() });
    return sendJSON({ error:"Unknown action" });
  } catch(err) {
    return sendJSON({ error: err.message });
  }
}

// ─────────────────────────────────────────────
function getPartnerData() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    const names = ss.getSheets().map(s => s.getName()).join(", ");
    throw new Error('Sheet "' + SHEET_NAME + '" not found. Available: ' + names);
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 3) return { headers:[], rows:[], total:0, monthCols:MONTH_COLS, lastUpdated: new Date().toISOString() };

  // Row 2 = headers (row 1 is the banner)
  const headers = sheet.getRange(2, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const raw     = sheet.getRange(3, 1, lastRow - 2, lastCol).getValues();

  const rows = [];
  raw.forEach(row => {
    if (row.every(c => c === "" || c === null || c === undefined)) return;
    const obj = {};
    headers.forEach((h, i) => { if (h) obj[h] = (row[i] === null || row[i] === undefined) ? "" : row[i]; });

    // Monthly data
    const md = {};
    let total = 0, months = 0;
    MONTH_COLS.forEach(m => {
      const v = pNum(obj[m]); md[m] = v;
      if (v > 0) { total += v; months++; }
    });
    obj._monthlyData        = md;
    obj._totalBusiness      = total;
    obj._avgMonthlyBusiness = months > 0 ? Math.round(total / months) : 0;

    // Growth
    const gRaw = String(obj["Growth/Degrowth"] || "").replace(/%/g,"").trim();
    obj._growthPct = parseFloat(gRaw) || 0;

    // Active flag
    const statusRaw = String(obj["Active/Inactive"] || "").toLowerCase();
    obj._isActive = statusRaw.includes("active") && !statusRaw.includes("inactive");

    rows.push(obj);
  });

  return { headers, rows, total: rows.length, monthCols: MONTH_COLS, lastUpdated: new Date().toISOString() };
}

// ─────────────────────────────────────────────
function getSummaryStats() {
  const d = getPartnerData();
  const rows = d.rows;
  const stats = {
    totalPartners:0, activePartners:0, inactivePartners:0,
    totalBusiness:0, growingPartners:0, degrowing:0,
    zoneBreakdown:{}, ownerBreakdown:{}, stateBreakdown:{},
    topPartners:[], lastUpdated: d.lastUpdated
  };
  rows.forEach(r => {
    stats.totalPartners++;
    if (r._isActive) stats.activePartners++; else stats.inactivePartners++;
    stats.totalBusiness += r._totalBusiness || 0;
    if (r._growthPct > 0) stats.growingPartners++; else if (r._growthPct < 0) stats.degrowing++;
    const z = String(r["Zone"]  || "Unknown").trim(); stats.zoneBreakdown[z]  = (stats.zoneBreakdown[z]  || 0) + 1;
    const o = String(r["Owner"] || "Unknown").trim(); stats.ownerBreakdown[o] = (stats.ownerBreakdown[o] || 0) + 1;
    const s = String(r["STATE"] || "Unknown").trim(); stats.stateBreakdown[s] = (stats.stateBreakdown[s] || 0) + 1;
  });
  stats.topPartners = rows
    .filter(r => r._totalBusiness > 0)
    .sort((a,b) => b._totalBusiness - a._totalBusiness)
    .slice(0,10)
    .map(r => ({ name: r["NAME"]||"—", gid: r["GID/GCD"]||"—", zone: r["Zone"]||"—", total: r._totalBusiness }));
  return stats;
}

// ─────────────────────────────────────────────
function pNum(v) {
  if (v === null || v === undefined || v === "") return 0;
  const n = parseFloat(String(v).replace(/,/g,""));
  return isNaN(n) ? 0 : n;
}

function sendJSON(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─────────────────────────────────────────────
// TEST — run this first to verify everything works
function testGetData() {
  const r = getPartnerData();
  Logger.log("Total rows : " + r.total);
  Logger.log("Headers    : " + r.headers.slice(0,10).join(", "));
  if (r.rows.length > 0) Logger.log("First row  : " + JSON.stringify(r.rows[0]).substring(0,300));
}
