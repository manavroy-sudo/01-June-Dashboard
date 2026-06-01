/**
 * ============================================================
 *  InsuranceDekho Partner Dashboard — Apps Script v3
 *  Role-Based Access Control Edition
 * ============================================================
 *
 *  SETUP:
 *  1. Paste this entire file into Apps Script (replace all)
 *  2. Run setupUsersSheet() ONCE to create the Users tab
 *  3. Run testGetData() to verify data reads correctly
 *  4. Deploy → New Deployment → Web App
 *     • Execute as : Me
 *     • Who has access : Anyone
 *  5. Copy /exec URL → paste into index.html
 */

// ── CONFIG ───────────────────────────────────────────────────
const SHEET_NAME   = "master";       // Your data sheet tab
const USERS_SHEET  = "DASHBOARD_USERS"; // Auto-created by setupUsersSheet()
const ADMIN_EMAIL  = "manav.roy@insurancedekho.com"; // Central admin

// Column indices in your master sheet (1-based)
const COL_DESIG    = 7;   // Column G = Designation (ZH/RH/SH/RM/AM)
const COL_OWNER    = 8;   // Column H = Owner Name

const MONTH_COLS = [
  "April'2025","May'25","June'25","July'25","August'25",
  "September'25","October'25","November'25","December'25",
  "January'26","Feburary'26","March'26","April'26"
];

// Role hierarchy — what each role can see
const ROLE_HIERARCHY = {
  "central": ["ZH","RH","SH","RM","AM"],
  "ZH":      ["RH","SH","RM","AM"],
  "RH":      ["SH","RM","AM"],
  "SH":      ["RM","AM"],
  "RM":      ["AM"],
  "AM":      []
};

// ── ROUTER ───────────────────────────────────────────────────
function doGet(e) {
  try {
    const p      = e && e.parameter ? e.parameter : {};
    const action = p.action || "ping";
    const token  = p.token  || "";

    if (action === "ping")        return sendJSON({ status:"ok", time: new Date().toISOString() });
    if (action === "login")       return sendJSON(handleLogin(p.email, p.password));
    if (action === "setPassword") return sendJSON(handleSetPassword(p.email, p.password, p.newPassword));
    if (action === "resetRequest")return sendJSON(handleResetRequest(p.email));

    // All other actions require a valid token
    const user = verifyToken(token);
    if (!user) return sendJSON({ error:"AUTH_REQUIRED", message:"Invalid or expired session." });

    if (action === "getData")      return sendJSON(getData(user, p));
    if (action === "getSummary")   return sendJSON(getSummary(user));
    if (action === "getHierarchy") return sendJSON(getHierarchy(user));

    // Admin-only actions
    if (!isAdmin(user)) return sendJSON({ error:"FORBIDDEN" });
    if (action === "getUsers")     return sendJSON(getUsers());
    if (action === "createUser")   return sendJSON(createUser(p));
    if (action === "updateUser")   return sendJSON(updateUser(p));
    if (action === "deleteUser")   return sendJSON(deleteUser(p.email));
    if (action === "getOwnerList") return sendJSON(getOwnerList());

    return sendJSON({ error:"Unknown action" });
  } catch(err) {
    return sendJSON({ error: "SERVER_ERROR", message: err.message });
  }
}

// ── AUTH ─────────────────────────────────────────────────────
function handleLogin(email, password) {
  if (!email || !password) return { error:"MISSING_FIELDS" };
  email = email.toLowerCase().trim();

  const users = getUsersSheet();
  const data  = users.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const uEmail = String(row[0]).toLowerCase().trim();
    if (uEmail !== email) continue;

    const status = String(row[5]).trim(); // Column F: status
    if (status === "disabled") return { error:"ACCOUNT_DISABLED" };

    const storedHash = String(row[4]).trim(); // Column E: password_hash
    if (!storedHash || storedHash === "SETUP_REQUIRED") {
      return { error:"SETUP_REQUIRED", email: email };
    }

    if (hashPassword(password) !== storedHash) {
      return { error:"WRONG_PASSWORD" };
    }

    // Generate token and store it
    const token = generateToken(email);
    users.getRange(i+1, 8).setValue(token);           // Col H: token
    users.getRange(i+1, 9).setValue(new Date().toISOString()); // Col I: last_login

    return {
      success: true,
      token,
      user: {
        email:       String(row[0]),
        name:        String(row[1]),
        role:        String(row[2]),
        designation: String(row[3]),
        ownerName:   String(row[6] || ""),
        zone:        String(row[7] || "")
      }
    };
  }
  return { error:"USER_NOT_FOUND" };
}

function handleSetPassword(email, currentOrTemp, newPassword) {
  if (!email || !newPassword) return { error:"MISSING_FIELDS" };
  email = email.toLowerCase().trim();
  const users = getUsersSheet();
  const data  = users.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const uEmail = String(data[i][0]).toLowerCase().trim();
    if (uEmail !== email) continue;

    const stored = String(data[i][4]).trim();
    // Allow set if: SETUP_REQUIRED, or current password matches
    if (stored !== "SETUP_REQUIRED" && hashPassword(currentOrTemp) !== stored) {
      return { error:"WRONG_CURRENT_PASSWORD" };
    }

    if (newPassword.length < 6) return { error:"PASSWORD_TOO_SHORT" };

    const newHash = hashPassword(newPassword);
    users.getRange(i+1, 5).setValue(newHash);
    users.getRange(i+1, 6).setValue("active");

    // Auto-login
    const token = generateToken(email);
    users.getRange(i+1, 8).setValue(token);
    users.getRange(i+1, 9).setValue(new Date().toISOString());

    return {
      success: true, token,
      user: {
        email:       String(data[i][0]),
        name:        String(data[i][1]),
        role:        String(data[i][2]),
        designation: String(data[i][3]),
        ownerName:   String(data[i][6] || ""),
        zone:        String(data[i][7] || "")
      }
    };
  }
  return { error:"USER_NOT_FOUND" };
}

function handleResetRequest(email) {
  // Just marks SETUP_REQUIRED so admin can reset
  if (!email) return { error:"MISSING_FIELDS" };
  email = email.toLowerCase().trim();
  const users = getUsersSheet();
  const data  = users.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase().trim() === email) {
      users.getRange(i+1, 5).setValue("SETUP_REQUIRED");
      return { success:true };
    }
  }
  return { error:"USER_NOT_FOUND" };
}

function verifyToken(token) {
  if (!token || token.length < 10) return null;
  const users = getUsersSheet();
  const data  = users.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][7]).trim() === token) {
      return {
        email:       String(data[i][0]),
        name:        String(data[i][1]),
        role:        String(data[i][2]),
        designation: String(data[i][3]),
        ownerName:   String(data[i][6] || ""),
        zone:        String(data[i][7+1] || "") // col J
      };
    }
  }
  return null;
}

function isAdmin(user) {
  return user && (user.role === "central" || user.email === ADMIN_EMAIL);
}

// ── DATA FETCHING ─────────────────────────────────────────────
function getData(user, params) {
  const allRows = readMasterSheet();
  const filtered = filterByRole(allRows, user);
  return {
    rows:     filtered,
    total:    filtered.length,
    monthCols: MONTH_COLS,
    user:     { name:user.name, role:user.role, designation:user.designation },
    lastUpdated: new Date().toISOString()
  };
}

function getSummary(user) {
  const rows = filterByRole(readMasterSheet(), user);
  const summary = {
    totalPartners:0, activePartners:0, inactivePartners:0,
    totalBusiness:0, growingPartners:0, degrowing:0,
    zoneBreakdown:{}, ownerBreakdown:{}, stateBreakdown:{},
    designationBreakdown:{}, topPartners:[],
    lastUpdated: new Date().toISOString()
  };
  rows.forEach(r => {
    summary.totalPartners++;
    if (r._isActive) summary.activePartners++; else summary.inactivePartners++;
    summary.totalBusiness += r._totalBusiness || 0;
    if (r._growthPct > 0) summary.growingPartners++; else if (r._growthPct < 0) summary.degrowing++;
    const z = String(r["Zone"] || "—").trim();   summary.zoneBreakdown[z]  = (summary.zoneBreakdown[z]||0)+1;
    const o = String(r["Owner"]|| "—").trim();   summary.ownerBreakdown[o] = (summary.ownerBreakdown[o]||0)+1;
    const s = String(r["STATE"]|| "—").trim();   summary.stateBreakdown[s] = (summary.stateBreakdown[s]||0)+1;
    const d = String(r["Designation"] || r["Owner"]||"—").trim(); summary.designationBreakdown[d]=(summary.designationBreakdown[d]||0)+1;
  });
  summary.topPartners = rows
    .filter(r=>r._totalBusiness>0)
    .sort((a,b)=>b._totalBusiness-a._totalBusiness)
    .slice(0,10)
    .map(r=>({ name:r["NAME"]||"—", gid:r["GID/GCD"]||"—", zone:r["Zone"]||"—", total:r._totalBusiness }));
  return summary;
}

function getHierarchy(user) {
  // Returns the org tree visible to this user
  const rows = filterByRole(readMasterSheet(), user);
  const tree = {};
  rows.forEach(r => {
    const zone  = String(r["Zone"]||"—").trim();
    const desig = String(r["Designation"]||"—").trim();
    const owner = String(r["Owner"]||"—").trim();
    if (!tree[zone]) tree[zone]={};
    if (!tree[zone][desig]) tree[zone][desig]=new Set();
    tree[zone][desig].add(owner);
  });
  // Convert Sets to arrays
  const result = {};
  Object.keys(tree).forEach(z=>{
    result[z]={};
    Object.keys(tree[z]).forEach(d=>{ result[z][d]=Array.from(tree[z][d]).sort(); });
  });
  return { hierarchy: result };
}

// ── ROLE FILTER ───────────────────────────────────────────────
function filterByRole(rows, user) {
  if (user.role === "central") return rows; // sees everything

  return rows.filter(r => {
    const rowDesig = String(r["Designation"] || "").trim().toUpperCase();
    const rowOwner = String(r["Owner"] || "").trim().toLowerCase();
    const userOwner = String(user.ownerName || "").trim().toLowerCase();
    const userDesig = String(user.designation || "").trim().toUpperCase();
    const userZone  = String(user.zone || "").trim().toLowerCase();

    if (user.role === "ZH") {
      // ZH sees all rows in their zone
      return String(r["Zone"] || "").trim().toLowerCase() === userZone ||
             rowOwner === userOwner;
    }
    if (user.role === "RH" || user.role === "SH") {
      // Sees rows where owner matches them or rows they supervise
      return rowOwner === userOwner || String(r["STATE"]||"").trim().toLowerCase() === String(user.zone||"").trim().toLowerCase();
    }
    if (user.role === "RM" || user.role === "AM") {
      return rowOwner === userOwner;
    }
    return false;
  });
}

// ── READ MASTER SHEET ─────────────────────────────────────────
function readMasterSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Sheet "' + SHEET_NAME + '" not found.');

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 3) return [];

  const headers = sheet.getRange(2, 1, 1, lastCol).getValues()[0].map(h=>String(h).trim());
  const raw     = sheet.getRange(3, 1, lastRow-2, lastCol).getValues();

  // Get designation and owner from their columns
  const rows = [];
  raw.forEach(row => {
    if (row.every(c=>c===""||c===null||c===undefined)) return;
    const obj = {};
    headers.forEach((h,i)=>{ if(h) obj[h] = (row[i]===null||row[i]===undefined)?'':row[i]; });

    // Map col G and H explicitly
    obj["Designation"] = String(row[COL_DESIG-1]||"").trim();
    obj["Owner"]       = String(row[COL_OWNER-1]||"").trim();

    // Monthly aggregation
    const md={}; let tot=0, cnt=0;
    MONTH_COLS.forEach(m=>{ const v=pNum(obj[m]); md[m]=v; if(v>0){tot+=v;cnt++;} });
    obj._monthlyData        = md;
    obj._totalBusiness      = tot;
    obj._avgMonthlyBusiness = cnt>0?Math.round(tot/cnt):0;

    // Growth
    obj._growthPct = parseFloat(String(obj["Growth/Degrowth"]||"").replace(/%/g,""))||0;

    // Active
    const s = String(obj["Active/Inactive"]||"").toLowerCase();
    obj._isActive = s.includes("active") && !s.includes("inactive");

    rows.push(obj);
  });
  return rows;
}

// ── USER MANAGEMENT ───────────────────────────────────────────
function getUsers() {
  const sheet = getUsersSheet();
  const data  = sheet.getDataRange().getValues();
  const users = [];
  for (let i=1;i<data.length;i++) {
    const r=data[i];
    if (!r[0]) continue;
    users.push({
      email:       String(r[0]),
      name:        String(r[1]),
      role:        String(r[2]),
      designation: String(r[3]),
      status:      String(r[5]),
      ownerName:   String(r[6]||""),
      zone:        String(r[10]||""),
      hasPassword: String(r[4]).trim() !== "" && String(r[4]).trim() !== "SETUP_REQUIRED",
      lastLogin:   String(r[9]||"")
    });
  }
  return { users };
}

function createUser(p) {
  if (!p.email||!p.name||!p.role) return { error:"MISSING_FIELDS" };
  const sheet = getUsersSheet();
  const data  = sheet.getDataRange().getValues();
  // Check duplicate
  for(let i=1;i<data.length;i++){
    if(String(data[i][0]).toLowerCase()===p.email.toLowerCase()) return {error:"EMAIL_EXISTS"};
  }
  sheet.appendRow([
    p.email.toLowerCase().trim(),  // A: email
    p.name.trim(),                 // B: name
    p.role,                        // C: role
    p.designation||p.role,         // D: designation
    "SETUP_REQUIRED",              // E: password (user sets on first login)
    "active",                      // F: status
    p.ownerName||"",               // G: ownerName (must match Owner column exactly)
    "",                            // H: token
    "",                            // I: last_login
    "",                            // J: zone
    p.zone||""                     // K: zone filter
  ]);
  return { success:true };
}

function updateUser(p) {
  if (!p.email) return { error:"MISSING_FIELDS" };
  const sheet = getUsersSheet();
  const data  = sheet.getDataRange().getValues();
  for(let i=1;i<data.length;i++){
    if(String(data[i][0]).toLowerCase()===p.email.toLowerCase()){
      if(p.name)        sheet.getRange(i+1,2).setValue(p.name);
      if(p.role)        sheet.getRange(i+1,3).setValue(p.role);
      if(p.designation) sheet.getRange(i+1,4).setValue(p.designation);
      if(p.status)      sheet.getRange(i+1,6).setValue(p.status);
      if(p.ownerName)   sheet.getRange(i+1,7).setValue(p.ownerName);
      if(p.zone)        sheet.getRange(i+1,11).setValue(p.zone);
      if(p.resetPassword==="true") sheet.getRange(i+1,5).setValue("SETUP_REQUIRED");
      return { success:true };
    }
  }
  return { error:"USER_NOT_FOUND" };
}

function deleteUser(email) {
  if (!email) return { error:"MISSING_FIELDS" };
  const sheet = getUsersSheet();
  const data  = sheet.getDataRange().getValues();
  for(let i=1;i<data.length;i++){
    if(String(data[i][0]).toLowerCase()===email.toLowerCase()){
      sheet.deleteRow(i+1);
      return { success:true };
    }
  }
  return { error:"USER_NOT_FOUND" };
}

function getOwnerList() {
  // Returns unique owner names + designations from master sheet for easy user creation
  const rows = readMasterSheet();
  const map  = {};
  rows.forEach(r=>{
    const o=String(r["Owner"]||"").trim();
    const d=String(r["Designation"]||"").trim();
    const z=String(r["Zone"]||"").trim();
    if(o && !map[o]) map[o]={ownerName:o,designation:d,zone:z};
  });
  return { owners: Object.values(map).sort((a,b)=>a.ownerName.localeCompare(b.ownerName)) };
}

// ── USERS SHEET SETUP ─────────────────────────────────────────
function setupUsersSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(USERS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(USERS_SHEET);
    // Headers
    sheet.getRange(1,1,1,11).setValues([[
      "Email","Name","Role","Designation","PasswordHash","Status",
      "OwnerName","Token","LastLogin","Reserved","Zone"
    ]]);
    sheet.getRange(1,1,1,11).setFontWeight("bold").setBackground("#1a2840").setFontColor("#ffffff");
    sheet.setColumnWidth(1,220); sheet.setColumnWidth(2,160); sheet.setColumnWidth(5,200); sheet.setColumnWidth(8,300);

    // Create the central admin account
    sheet.appendRow([
      ADMIN_EMAIL, "Central Admin", "central", "central",
      "SETUP_REQUIRED", "active", "", "", "", "", ""
    ]);
    Logger.log("✅ Users sheet created. Admin: " + ADMIN_EMAIL + " — first login will prompt password setup.");
  } else {
    Logger.log("Users sheet already exists.");
  }
}

// ── UTILS ─────────────────────────────────────────────────────
function hashPassword(pw) {
  // Simple deterministic hash using Apps Script's Utilities
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pw + "ID_SALT_2026");
  return bytes.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

function generateToken(email) {
  const rand = Utilities.getUuid();
  return Utilities.base64Encode(email + ":" + rand + ":" + Date.now()).replace(/[+/=]/g,"").substring(0,48);
}

function getUsersSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(USERS_SHEET);
  if (!sheet) throw new Error('Users sheet not found. Run setupUsersSheet() first.');
  return sheet;
}

function pNum(v) {
  if(v===null||v===undefined||v==='') return 0;
  const n=parseFloat(String(v).replace(/,/g,''));
  return isNaN(n)?0:n;
}

function sendJSON(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── TEST FUNCTIONS ────────────────────────────────────────────
function testSetup() {
  setupUsersSheet();
  Logger.log("Setup complete.");
}

function testGetData() {
  const rows = readMasterSheet();
  Logger.log("Total rows: " + rows.length);
  if(rows.length>0) {
    Logger.log("Designation col: " + rows[0]["Designation"]);
    Logger.log("Owner col: " + rows[0]["Owner"]);
    Logger.log("Headers sample: " + Object.keys(rows[0]).slice(0,8).join(", "));
  }
}

function testOwnerList() {
  const r = getOwnerList();
  Logger.log("Unique owners: " + r.owners.length);
  r.owners.slice(0,5).forEach(o=>Logger.log(JSON.stringify(o)));
}
