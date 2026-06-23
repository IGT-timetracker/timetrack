// ============================================================
//  TimeTrack PWA — app.js
//  Microsoft Graph API + SharePoint Excel as database
// ============================================================

// ── MSAL Configuration ──────────────────────────────────────
// Using the Microsoft common endpoint so any work/school
// account can sign in without a custom Azure App Registration.
// For production, replace CLIENT_ID with your own App Registration.
const MSAL_CONFIG = {
  auth: {
    clientId: "d3590ed6-52b3-4102-aeff-aad2292ab01c", // Microsoft Office public client ID
    authority: "https://login.microsoftonline.com/common",
    redirectUri: window.location.origin + window.location.pathname,
    navigateToLoginRequestUrl: false,
  },
  cache: {
    cacheLocation: "localStorage",
    storeAuthStateInCookie: true,
  },
};

const GRAPH_SCOPES = ["Files.ReadWrite", "Sites.ReadWrite.All", "User.Read"];

let msalInstance;
let currentUser = null;
let accessToken = null;
let employees = [];
let clockEntries = [];
let settings = {};
let selectedEmp = null;
let editingEmpKey = null;
let isSyncing = false;

// ── Startup ─────────────────────────────────────────────────
async function startApp() {
  msalInstance = new msal.PublicClientApplication(MSAL_CONFIG);
  await msalInstance.initialize();

  // Handle redirect after login
  try {
    const resp = await msalInstance.handleRedirectPromise();
    if (resp) {
      currentUser = resp.account;
      accessToken = resp.accessToken;
    }
  } catch (e) {
    console.error("Redirect error:", e);
  }

  // Check if already signed in
  const accounts = msalInstance.getAllAccounts();
  if (accounts.length > 0 && !currentUser) {
    currentUser = accounts[0];
    await refreshToken();
  }

  if (currentUser) {
    showApp();
  } else {
    document.getElementById("login-screen").style.display = "flex";
  }

  // PWA install prompt
  window.addEventListener("beforeinstallprompt", (e) => {
    e.preventDefault();
    window._installPrompt = e;
    showInstallBanner();
  });

  // Offline detection
  window.addEventListener("online", () => updateOnlineStatus());
  window.addEventListener("offline", () => updateOnlineStatus());
  updateOnlineStatus();
}

function updateOnlineStatus() {
  const el = document.getElementById("offline-indicator");
  if (el) el.style.display = navigator.onLine ? "none" : "inline-block";
}

// ── Auth ─────────────────────────────────────────────────────
async function signIn() {
  document.getElementById("login-status").textContent = "Redirecting to Microsoft login…";
  try {
    await msalInstance.loginRedirect({ scopes: GRAPH_SCOPES });
  } catch (e) {
    document.getElementById("login-status").textContent = "Sign-in failed: " + e.message;
  }
}

async function refreshToken() {
  try {
    const resp = await msalInstance.acquireTokenSilent({
      scopes: GRAPH_SCOPES,
      account: currentUser,
    });
    accessToken = resp.accessToken;
    return true;
  } catch (e) {
    try {
      await msalInstance.acquireTokenRedirect({ scopes: GRAPH_SCOPES, account: currentUser });
    } catch (e2) {
      console.error("Token refresh failed:", e2);
    }
    return false;
  }
}

function signOut() {
  msalInstance.logoutRedirect({ account: currentUser });
}

// ── Show App ─────────────────────────────────────────────────
async function showApp() {
  document.getElementById("login-screen").style.display = "none";
  document.getElementById("app").style.display = "block";

  // Set user info
  const name = currentUser.name || currentUser.username;
  document.getElementById("user-name").textContent = name.split(" ")[0] || name;
  document.getElementById("user-avatar").textContent = initials(name);

  // Load local data first for instant UI
  loadLocal();
  renderAll();
  startClock();

  document.getElementById("report-date").value = today();

  // Then sync from SharePoint
  if (settings.siteName && settings.filePath) {
    await syncFromSharePoint(true);
  } else {
    showConfigNotice();
  }
}

function showConfigNotice() {
  const el = document.getElementById("config-notice");
  if (el) el.style.display = "block";
}

// ── Local Storage ─────────────────────────────────────────────
function loadLocal() {
  employees = JSON.parse(localStorage.getItem("tt_employees") || "[]");
  clockEntries = JSON.parse(localStorage.getItem("tt_entries") || "[]");
  settings = JSON.parse(localStorage.getItem("tt_settings") || "{}");

  if (!employees.length) {
    employees = [
      { key: "e1", name: "Alex Chen", empId: "EMP001", area: "Production", startTime: "09:00", endTime: "17:00", hours: 8 },
      { key: "e2", name: "Jordan Smith", empId: "EMP002", area: "Warehouse", startTime: "08:00", endTime: "16:00", hours: 8 },
      { key: "e3", name: "Sam Patel", empId: "EMP003", area: "Office", startTime: "07:00", endTime: "15:00", hours: 8 },
    ];
    saveLocal();
  }
}

function saveLocal() {
  localStorage.setItem("tt_employees", JSON.stringify(employees));
  localStorage.setItem("tt_entries", JSON.stringify(clockEntries));
  localStorage.setItem("tt_settings", JSON.stringify(settings));
}

// ── SharePoint / Graph API ────────────────────────────────────
function graphHeaders() {
  return {
    Authorization: `Bearer ${accessToken}`,
    "Content-Type": "application/json",
  };
}

async function getSharePointFileId() {
  const site = settings.siteName;
  const path = settings.filePath;
  if (!site || !path) return null;

  // Get site ID
  const tenantDomain = currentUser.username.split("@")[1];
  const siteRes = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${tenantDomain}:/sites/${site}`,
    { headers: graphHeaders() }
  );
  if (!siteRes.ok) throw new Error("Site not found: " + site);
  const siteData = await siteRes.json();
  const siteId = siteData.id;

  // Get file drive item
  const fileRes = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/${path}`,
    { headers: graphHeaders() }
  );
  if (!fileRes.ok) throw new Error("File not found: " + path);
  const fileData = await fileRes.json();
  return { siteId, driveId: fileData.parentReference.driveId, itemId: fileData.id };
}

// Read the whole Excel workbook from SharePoint
async function readWorkbookFromSharePoint() {
  await refreshToken();
  const ids = await getSharePointFileId();
  if (!ids) throw new Error("SharePoint not configured");

  const { driveId, itemId } = ids;
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/content`,
    { headers: { Authorization: `Bearer ${accessToken}` } }
  );
  if (!res.ok) throw new Error("Failed to download file");

  const arrayBuffer = await res.arrayBuffer();
  const wb = XLSX.read(arrayBuffer, { type: "array", cellDates: true });
  return wb;
}

// Write the whole workbook back to SharePoint
async function writeWorkbookToSharePoint(wb) {
  await refreshToken();
  const ids = await getSharePointFileId();
  if (!ids) throw new Error("SharePoint not configured");

  const { driveId, itemId } = ids;
  const wbOut = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  const blob = new Blob([wbOut], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });

  const res = await fetch(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/content`,
    {
      method: "PUT",
      headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/octet-stream" },
      body: blob,
    }
  );
  if (!res.ok) throw new Error("Failed to upload file");
  return true;
}

// Parse workbook sheets into app data
function parseWorkbook(wb) {
  const result = { employees: [], entries: [], settings: {} };

  if (wb.SheetNames.includes("Employees")) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets["Employees"], { defval: "" });
    result.employees = rows.map((r) => ({
      key: r.Key || "e" + Date.now() + Math.random(),
      name: r.Name || "",
      empId: r.EmployeeID || "",
      area: r.Area || "",
      startTime: r.StartTime || "09:00",
      endTime: r.EndTime || "17:00",
      hours: parseFloat(r.HoursPerDay) || 8,
    })).filter(e => e.name);
  }

  if (wb.SheetNames.includes("Attendance")) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets["Attendance"], { defval: "" });
    result.entries = rows.map((r) => ({
      empKey: r.EmpKey || "",
      empId: r.EmployeeID || "",
      name: r.Name || "",
      area: r.Area || "",
      date: r.Date || "",
      timeIn: r.TimeIn || null,
      timeOut: r.TimeOut || null,
      stdStart: r.StdStart || "",
      stdEnd: r.StdEnd || "",
      stdHours: parseFloat(r.StdHours) || 8,
    })).filter(e => e.date && e.name);
  }

  if (wb.SheetNames.includes("Settings")) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets["Settings"], { defval: "" });
    rows.forEach((r) => {
      if (r.Key && r.Value !== undefined) result.settings[r.Key] = r.Value;
    });
  }

  return result;
}

// Build workbook from app data
function buildWorkbook() {
  const wb = XLSX.utils.book_new();

  // Employees sheet
  const empData = employees.map((e) => ({
    Key: e.key,
    EmployeeID: e.empId,
    Name: e.name,
    Area: e.area,
    StartTime: e.startTime,
    EndTime: e.endTime,
    HoursPerDay: e.hours,
  }));
  const wsEmp = XLSX.utils.json_to_sheet(empData);
  wsEmp["!cols"] = [{ wch: 14 }, { wch: 12 }, { wch: 22 }, { wch: 16 }, { wch: 11 }, { wch: 11 }, { wch: 13 }];
  XLSX.utils.book_append_sheet(wb, wsEmp, "Employees");

  // Attendance sheet
  const entData = clockEntries.map((e) => ({
    EmpKey: e.empKey,
    EmployeeID: e.empId,
    Name: e.name,
    Area: e.area,
    Date: e.date,
    TimeIn: e.timeIn || "",
    TimeOut: e.timeOut || "",
    StdStart: e.stdStart,
    StdEnd: e.stdEnd,
    StdHours: e.stdHours,
  }));
  const wsAtt = XLSX.utils.json_to_sheet(entData);
  wsAtt["!cols"] = [
    { wch: 14 }, { wch: 12 }, { wch: 22 }, { wch: 16 },
    { wch: 12 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 12 },
  ];
  XLSX.utils.book_append_sheet(wb, wsAtt, "Attendance");

  // Settings sheet
  const settingsArr = Object.entries(settings).map(([k, v]) => ({ Key: k, Value: v }));
  const wsSet = XLSX.utils.json_to_sheet(settingsArr);
  XLSX.utils.book_append_sheet(wb, wsSet, "Settings");

  return wb;
}

// Sync FROM SharePoint into local
async function syncFromSharePoint(silent = false) {
  if (isSyncing) return;
  if (!settings.siteName || !settings.filePath) {
    if (!silent) toast("Configure SharePoint settings first", "error");
    return;
  }
  isSyncing = true;
  setSyncStatus("Syncing…");
  try {
    const wb = await readWorkbookFromSharePoint();
    const data = parseWorkbook(wb);
    if (data.employees.length) employees = data.employees;
    if (data.entries.length) {
      // Merge: keep local entries not in SharePoint (offline clock-ins)
      const spKeys = new Set(data.entries.map(e => e.empKey + e.date));
      const localOnly = clockEntries.filter(e => !spKeys.has(e.empKey + e.date));
      clockEntries = [...data.entries, ...localOnly];
    }
    if (Object.keys(data.settings).length) {
      settings = { ...settings, ...data.settings };
    }
    saveLocal();
    renderAll();
    setSyncStatus("Synced " + new Date().toLocaleTimeString());
    if (!silent) toast("Pulled latest data from SharePoint");
    logSync("✓ Pulled from SharePoint at " + new Date().toLocaleTimeString());
  } catch (e) {
    setSyncStatus("Sync failed");
    if (!silent) toast("Sync failed: " + e.message, "error");
    logSync("✗ Pull failed: " + e.message);
    console.error(e);
  }
  isSyncing = false;
}

async function syncToSharePoint() {
  if (isSyncing) return;
  if (!settings.siteName || !settings.filePath) {
    toast("Configure SharePoint settings first", "error");
    return;
  }
  isSyncing = true;
  setSyncStatus("Uploading…");
  try {
    const wb = buildWorkbook();
    await writeWorkbookToSharePoint(wb);
    setSyncStatus("Saved " + new Date().toLocaleTimeString());
    toast("Pushed to SharePoint successfully");
    logSync("✓ Pushed to SharePoint at " + new Date().toLocaleTimeString());
  } catch (e) {
    setSyncStatus("Upload failed");
    toast("Upload failed: " + e.message, "error");
    logSync("✗ Push failed: " + e.message);
    console.error(e);
  }
  isSyncing = false;
}

async function testConnection() {
  toast("Testing SharePoint connection…");
  try {
    await refreshToken();
    await getSharePointFileId();
    toast("✓ Connection successful! File found in SharePoint.");
    logSync("✓ Connection test passed at " + new Date().toLocaleTimeString());
  } catch (e) {
    toast("Connection failed: " + e.message, "error");
    logSync("✗ Connection test failed: " + e.message);
  }
}

function setSyncStatus(msg) {
  const el = document.getElementById("sync-status");
  if (el) el.textContent = msg;
}

function logSync(msg) {
  const el = document.getElementById("sync-log");
  if (el) el.textContent = msg;
}

// ── Clock Logic ───────────────────────────────────────────────
function today() {
  return new Date().toISOString().slice(0, 10);
}

function startClock() {
  function tick() {
    const n = new Date();
    const el = document.getElementById("live-clock");
    if (el) el.textContent = n.toLocaleTimeString("en-AU", { hour: "2-digit", minute: "2-digit", second: "2-digit", hour12: false });
    const de = document.getElementById("clock-date");
    if (de) de.textContent = n.toLocaleDateString("en-AU", { weekday: "long", year: "numeric", month: "long", day: "numeric" });
    const tl = document.getElementById("today-date-label");
    if (tl) tl.textContent = n.toLocaleDateString("en-AU", { weekday: "long", day: "numeric", month: "long" });
  }
  tick();
  setInterval(tick, 1000);
}

function getClockedInEntry(empKey) {
  return clockEntries.find(e => e.empKey === empKey && e.date === today() && e.timeIn && !e.timeOut);
}

async function clockIn(empKey) {
  const emp = employees.find(e => e.key === empKey);
  const timeIn = new Date().toLocaleTimeString("en-AU", { hour: "2-digit", minute: "2-digit", hour12: false });
  clockEntries.push({
    empKey, date: today(), timeIn, timeOut: null,
    name: emp.name, area: emp.area, empId: emp.empId,
    stdStart: emp.startTime, stdEnd: emp.endTime, stdHours: emp.hours,
  });
  saveLocal();
  toast(`${emp.name} clocked in at ${timeIn}`);
  renderAll();
  // Auto-sync to SharePoint
  syncToSharePoint();
}

async function clockOut(empKey) {
  const emp = employees.find(e => e.key === empKey);
  const timeOut = new Date().toLocaleTimeString("en-AU", { hour: "2-digit", minute: "2-digit", hour12: false });
  const idx = clockEntries.findIndex(e => e.empKey === empKey && e.date === today() && e.timeIn && !e.timeOut);
  if (idx >= 0) clockEntries[idx].timeOut = timeOut;
  saveLocal();
  selectedEmp = null;
  toast(`${emp.name} clocked out at ${timeOut}`);
  renderAll();
  syncToSharePoint();
}

// ── Render ────────────────────────────────────────────────────
function renderAll() {
  renderEmpGrid();
  renderClockAction();
  renderTodayTable();
  renderActiveBanner();
  renderEmpList();
  loadSettingsForm();
  renderReportRecipient();
  genReport();
}

const AVATAR_COLORS = [
  ["#E6F1FB", "#185FA5"], ["#EAF3DE", "#3B6D11"], ["#FAEEDA", "#854F0B"],
  ["#EEEDFE", "#534AB7"], ["#FAECE7", "#993C1D"], ["#E1F5EE", "#085041"],
];

function initials(n) {
  return (n || "?").split(" ").map(x => x[0]).join("").toUpperCase().slice(0, 2);
}

function avatarStyle(i) {
  const c = AVATAR_COLORS[i % AVATAR_COLORS.length];
  return `background:${c[0]};color:${c[1]}`;
}

function renderEmpGrid() {
  const g = document.getElementById("emp-select-grid");
  if (!g) return;
  if (!employees.length) {
    g.innerHTML = '<div style="font-size:13px;color:var(--text2)">No employees. Add them in Admin.</div>';
    return;
  }
  g.innerHTML = employees.map((e, i) => `
    <div class="emp-tile${selectedEmp === e.key ? " selected" : ""}" onclick="selectEmp('${e.key}')">
      <div class="emp-avatar" style="${avatarStyle(i)};width:40px;height:40px;font-size:14px;margin:0 auto">${initials(e.name)}</div>
      <div class="emp-tile-name">${e.name}</div>
      <div class="emp-tile-area">${e.area}</div>
      <div class="emp-tile-time">${e.startTime}–${e.endTime}</div>
    </div>`).join("");
}

function selectEmp(key) {
  selectedEmp = key;
  renderEmpGrid();
  renderClockAction();
}

function renderClockAction() {
  const area = document.getElementById("clock-action-area");
  if (!area) return;
  if (!selectedEmp) {
    area.innerHTML = '<div style="color:var(--text2);font-size:13px;text-align:center">Select an employee above to clock in or out.</div>';
    return;
  }
  const entry = getClockedInEntry(selectedEmp);
  const emp = employees.find(e => e.key === selectedEmp);
  if (!emp) return;
  if (entry) {
    area.innerHTML = `<div style="text-align:center">
      <div style="font-size:13px;color:var(--text2);margin-bottom:.75rem">${emp.name} clocked in at <strong>${entry.timeIn}</strong></div>
      <button class="btn btn-danger" style="margin:0 auto" onclick="clockOut('${selectedEmp}')">⏹ Clock out — ${emp.name}</button>
    </div>`;
  } else {
    const alreadyDone = clockEntries.find(e => e.empKey === selectedEmp && e.date === today() && e.timeOut);
    area.innerHTML = `<div style="text-align:center">
      ${alreadyDone ? `<div style="font-size:12px;color:var(--text2);margin-bottom:.5rem">Already completed today (${alreadyDone.timeIn}–${alreadyDone.timeOut})</div>` : ""}
      <button class="btn btn-success" style="margin:0 auto" onclick="clockIn('${selectedEmp}')">▶ Clock in — ${emp.name}</button>
    </div>`;
  }
}

function renderActiveBanner() {
  const d = today();
  const active = clockEntries.filter(e => e.date === d && e.timeIn && !e.timeOut);
  const el = document.getElementById("active-banner");
  if (!el) return;
  if (!active.length) { el.innerHTML = ""; return; }
  el.innerHTML = `<div class="active-banner">
    <div style="font-size:20px">🟢</div>
    <div>
      <div style="font-weight:600;color:#085041;font-size:13px">${active.length} employee${active.length > 1 ? "s" : ""} currently clocked in</div>
      <div style="font-size:12px;color:#0F6E56;margin-top:2px">${active.map(e => `${e.name} (since ${e.timeIn})`).join(" · ")}</div>
    </div>
  </div>`;
}

function renderTodayTable() {
  const d = today();
  const entries = clockEntries.filter(e => e.date === d);
  const wrap = document.getElementById("today-table-wrap");
  if (!wrap) return;
  if (!entries.length) {
    wrap.innerHTML = '<div class="empty">⏰<br><br>No clock-ins recorded today</div>';
    return;
  }
  wrap.innerHTML = `<div style="overflow-x:auto"><table>
    <thead><tr><th>Employee</th><th>Area</th><th>Clock in</th><th>Clock out</th><th>Hours</th><th>Status</th></tr></thead>
    <tbody>${entries.map(e => {
      const hrs = calcHours(e.timeIn, e.timeOut);
      return `<tr>
        <td><div class="emp-row"><div class="emp-avatar" style="${avatarStyle(employees.findIndex(x => x.key === e.empKey))};width:30px;height:30px;font-size:11px">${initials(e.name)}</div><div><div style="font-weight:500">${e.name}</div><div style="font-size:11px;color:var(--text2)">${e.empId}</div></div></div></td>
        <td><span class="tag">${e.area}</span></td>
        <td>${e.timeIn || "—"}</td>
        <td>${e.timeOut || "—"}</td>
        <td>${hrs !== null ? hrs.toFixed(1) + "h" : "—"}</td>
        <td>${e.timeOut ? '<span class="badge badge-green">✓ Done</span>' : '<span class="badge badge-amber">● Active</span>'}</td>
      </tr>`;
    }).join("")}</tbody>
  </table></div>`;
}

// ── Report ────────────────────────────────────────────────────
function calcHours(tin, tout) {
  if (!tin || !tout) return null;
  const [h1, m1] = tin.split(":").map(Number);
  const [h2, m2] = tout.split(":").map(Number);
  return ((h2 * 60 + m2) - (h1 * 60 + m1)) / 60;
}

function timeDiffStr(t1, t2) {
  if (!t1 || !t2) return null;
  const [h1, m1] = t1.split(":").map(Number);
  const [h2, m2] = t2.split(":").map(Number);
  const diff = (h2 * 60 + m2) - (h1 * 60 + m1);
  const sign = diff < 0 ? "-" : "+";
  const abs = Math.abs(diff);
  return `${sign}${Math.floor(abs / 60)}h ${abs % 60}m`;
}

function renderReportRecipient() {
  const el = document.getElementById("report-recipient-info");
  if (!el) return;
  const name = settings.recipientName || "—";
  const email = settings.recipientEmail || "—";
  el.innerHTML = `📧 Report recipient: <strong>${name}</strong> &lt;${email}&gt;`;
}

function genReport() {
  const dateVal = document.getElementById("report-date")?.value || today();
  const entries = clockEntries.filter(e => e.date === dateVal);
  const wrap = document.getElementById("report-content");
  if (!wrap) return;

  if (!entries.length) {
    wrap.innerHTML = `<div class="card"><div class="empty">📅<br><br>No records for ${dateVal}</div></div>`;
    return;
  }

  const rows = entries.map(e => {
    const actual = calcHours(e.timeIn, e.timeOut);
    const diff = actual !== null ? actual - e.stdHours : null;
    const inVar = timeDiffStr(e.stdStart, e.timeIn);
    const outVar = e.timeOut ? timeDiffStr(e.stdEnd, e.timeOut) : null;
    const status = e.timeOut ? (diff !== null && diff >= 0 ? "On time" : "Short") : e.timeIn ? "In progress" : "Absent";
    return { ...e, actual, diff, inVar, outVar, status };
  });

  const totalStd = entries.reduce((s, e) => s + e.stdHours, 0);
  const totalActual = rows.reduce((s, r) => s + (r.actual || 0), 0);
  const onTime = rows.filter(r => r.status === "On time").length;

  wrap.innerHTML = `
    <div class="grid3" style="margin-bottom:1rem">
      <div class="stat-card"><div class="stat-label">Employees</div><div class="stat-value">${entries.length}</div></div>
      <div class="stat-card"><div class="stat-label">Std hours</div><div class="stat-value">${totalStd}h</div></div>
      <div class="stat-card"><div class="stat-label">Actual hours</div><div class="stat-value">${totalActual.toFixed(1)}h</div></div>
    </div>
    <div class="card">
      <div style="font-weight:600;font-size:15px;margin-bottom:.75rem">
        Timesheet — ${dateVal}
        <span class="tag" style="margin-left:6px">${settings.company || ""}</span>
        <span class="badge badge-green" style="margin-left:6px">${onTime}/${entries.length} on time</span>
      </div>
      <div style="overflow-x:auto">
      <table>
        <thead><tr>
          <th>Employee</th><th>ID</th><th>Area</th>
          <th>Std start</th><th>Actual in</th><th>Variance</th>
          <th>Std end</th><th>Actual out</th><th>Variance</th>
          <th>Std hrs</th><th>Actual hrs</th><th>Diff</th><th>Status</th>
        </tr></thead>
        <tbody>${rows.map(r => `<tr>
          <td><div class="emp-row"><div class="emp-avatar" style="${avatarStyle(employees.findIndex(x => x.key === r.empKey))};width:28px;height:28px;font-size:11px">${initials(r.name)}</div><div style="font-weight:500">${r.name}</div></div></td>
          <td style="color:var(--text2)">${r.empId}</td>
          <td><span class="tag">${r.area}</span></td>
          <td>${r.stdStart}</td>
          <td><strong>${r.timeIn || "—"}</strong></td>
          <td class="${r.inVar ? (r.inVar.startsWith("+") ? "time-diff-neg" : "time-diff-pos") : ""}">${r.inVar || "—"}</td>
          <td>${r.stdEnd}</td>
          <td><strong>${r.timeOut || "—"}</strong></td>
          <td class="${r.outVar ? (r.outVar.startsWith("-") ? "time-diff-neg" : "time-diff-pos") : ""}">${r.outVar || "—"}</td>
          <td>${r.stdHours}h</td>
          <td>${r.actual !== null ? r.actual.toFixed(1) + "h" : "—"}</td>
          <td class="${r.diff === null ? "" : r.diff >= 0 ? "time-diff-pos" : "time-diff-neg"}">${r.diff === null ? "—" : (r.diff >= 0 ? "+" : "") + r.diff.toFixed(1) + "h"}</td>
          <td>${statusBadge(r.status)}</td>
        </tr>`).join("")}</tbody>
      </table>
      </div>
    </div>`;
}

function statusBadge(s) {
  if (s === "On time") return '<span class="badge badge-green">✓ On time</span>';
  if (s === "Short") return '<span class="badge badge-red">⚠ Short</span>';
  if (s === "In progress") return '<span class="badge badge-amber">● Active</span>';
  return '<span class="badge badge-gray">Absent</span>';
}

function exportExcel() {
  const dateVal = document.getElementById("report-date")?.value || today();
  const entries = clockEntries.filter(e => e.date === dateVal);

  const wsData = [
    [`${settings.company || "Company"} — Daily Timesheet Report`, "", "", "", "", "", "", "", "", "", "", ""],
    [`Date: ${dateVal}`, "", "Report for:", `${settings.recipientName || ""} <${settings.recipientEmail || ""}>`, "", "", "", "", "", "", "", ""],
    [],
    ["Employee", "Employee ID", "Work Area", "Std Start", "Actual Clock In", "Start Variance", "Std End", "Actual Clock Out", "End Variance", "Std Hours", "Actual Hours", "Difference", "Status"],
    ...entries.map(e => {
      const actual = calcHours(e.timeIn, e.timeOut);
      const diff = actual !== null ? +(actual - e.stdHours).toFixed(2) : null;
      const inVar = timeDiffStr(e.stdStart, e.timeIn);
      const outVar = e.timeOut ? timeDiffStr(e.stdEnd, e.timeOut) : null;
      const status = e.timeOut ? (diff !== null && diff >= 0 ? "On time" : "Short hours") : e.timeIn ? "In progress" : "Absent";
      return [e.name, e.empId, e.area, e.stdStart, e.timeIn || "", inVar || "", e.stdEnd, e.timeOut || "", outVar || "", e.stdHours, actual !== null ? +actual.toFixed(2) : "", diff !== null ? diff : "", status];
    }),
    [],
    ["", "", "", "", "", "", "", "", "", `=SUM(J5:J${entries.length + 4})`, `=SUM(K5:K${entries.length + 4})`, `=SUM(L5:L${entries.length + 4})`, ""],
    ["", "", "", "", "", "", "", "", "", "Total Std Hrs", "Total Actual Hrs", "Total Diff", ""],
  ];

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(wsData);
  ws["!cols"] = [{ wch: 22 }, { wch: 12 }, { wch: 16 }, { wch: 11 }, { wch: 16 }, { wch: 14 }, { wch: 11 }, { wch: 16 }, { wch: 14 }, { wch: 11 }, { wch: 13 }, { wch: 11 }, { wch: 14 }];
  ws["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 12 } }];
  XLSX.utils.book_append_sheet(wb, ws, "Daily Timesheet");

  // Full history sheet
  const allData = [
    ["Employee", "Employee ID", "Area", "Date", "Time In", "Time Out", "Std Hours", "Actual Hours", "Difference"],
    ...clockEntries.map(e => {
      const a = calcHours(e.timeIn, e.timeOut);
      return [e.name, e.empId, e.area, e.date, e.timeIn || "", e.timeOut || "", e.stdHours, a !== null ? +a.toFixed(2) : "", a !== null ? +(a - e.stdHours).toFixed(2) : ""];
    }),
  ];
  const ws2 = XLSX.utils.aoa_to_sheet(allData);
  ws2["!cols"] = [{ wch: 22 }, { wch: 12 }, { wch: 16 }, { wch: 12 }, { wch: 10 }, { wch: 10 }, { wch: 12 }, { wch: 13 }, { wch: 12 }];
  XLSX.utils.book_append_sheet(wb, ws2, "Full History");

  XLSX.writeFile(wb, `Timesheet_${dateVal}.xlsx`);
  toast("Excel report downloaded!");
}

// ── Admin ─────────────────────────────────────────────────────
function renderEmpList() {
  const el = document.getElementById("emp-list");
  if (!el) return;
  if (!employees.length) {
    el.innerHTML = '<div class="card"><div class="empty">No employees added yet</div></div>';
    return;
  }
  el.innerHTML = employees.map((e, i) => `
    <div class="card" style="margin-bottom:8px;padding:1rem">
      <div style="display:flex;align-items:center;gap:10px">
        <div class="emp-avatar" style="${avatarStyle(i)}">${initials(e.name)}</div>
        <div style="flex:1">
          <div style="font-weight:600;font-size:14px">${e.name} <span class="tag">${e.empId}</span></div>
          <div style="font-size:12px;color:var(--text2)">${e.area} · ${e.startTime}–${e.endTime} · ${e.hours}h/day</div>
        </div>
        <div style="display:flex;gap:6px">
          <button class="btn" onclick="openEmpModal('${e.key}')" style="padding:6px 10px">✏</button>
          <button class="btn btn-danger" onclick="deleteEmp('${e.key}')" style="padding:6px 10px">🗑</button>
        </div>
      </div>
    </div>`).join("");
}

function openEmpModal(key) {
  editingEmpKey = key || null;
  const areas = (settings.areas || "Production,Warehouse,Office,Kitchen").split(",").map(a => a.trim());
  document.getElementById("emp-area").innerHTML = areas.map(a => `<option>${a}</option>`).join("");
  if (key) {
    const e = employees.find(x => x.key === key);
    document.getElementById("modal-title").textContent = "Edit employee";
    document.getElementById("emp-name").value = e.name;
    document.getElementById("emp-id-field").value = e.empId;
    document.getElementById("emp-area").value = e.area;
    document.getElementById("emp-start").value = e.startTime;
    document.getElementById("emp-end").value = e.endTime;
    document.getElementById("emp-hours").value = e.hours;
  } else {
    document.getElementById("modal-title").textContent = "Add employee";
    document.getElementById("emp-name").value = "";
    document.getElementById("emp-id-field").value = "";
    document.getElementById("emp-start").value = "09:00";
    document.getElementById("emp-end").value = "17:00";
    document.getElementById("emp-hours").value = "8";
  }
  document.getElementById("emp-modal").classList.add("open");
}

function closeEmpModal() {
  document.getElementById("emp-modal").classList.remove("open");
}

async function saveEmployee() {
  const name = document.getElementById("emp-name").value.trim();
  const empId = document.getElementById("emp-id-field").value.trim();
  const area = document.getElementById("emp-area").value;
  const startTime = document.getElementById("emp-start").value;
  const endTime = document.getElementById("emp-end").value;
  const hours = parseFloat(document.getElementById("emp-hours").value);
  if (!name || !empId) { toast("Name and ID required", "error"); return; }
  if (editingEmpKey) {
    const idx = employees.findIndex(e => e.key === editingEmpKey);
    employees[idx] = { ...employees[idx], name, empId, area, startTime, endTime, hours };
  } else {
    employees.push({ key: "e" + Date.now(), name, empId, area, startTime, endTime, hours });
  }
  saveLocal();
  closeEmpModal();
  toast(editingEmpKey ? "Employee updated" : "Employee added");
  renderAll();
  syncToSharePoint();
}

async function deleteEmp(key) {
  if (!confirm("Remove this employee?")) return;
  employees = employees.filter(e => e.key !== key);
  saveLocal();
  toast("Employee removed");
  renderAll();
  syncToSharePoint();
}

function loadSettingsForm() {
  document.getElementById("cfg-site").value = settings.siteName || "APACManufacturingOperationsTeam";
  document.getElementById("cfg-path").value = settings.filePath || "General/ATTENDANCE/Attendance.xlsx";
  document.getElementById("cfg-areas").value = settings.areas || "Production,Warehouse,Office,Kitchen";
  document.getElementById("cfg-company").value = settings.company || "";
  document.getElementById("cfg-recipient-name").value = settings.recipientName || "";
  document.getElementById("cfg-recipient-email").value = settings.recipientEmail || "";
}

async function saveSettings() {
  settings = {
    ...settings,
    siteName: document.getElementById("cfg-site").value.trim(),
    filePath: document.getElementById("cfg-path").value.trim(),
    areas: document.getElementById("cfg-areas").value,
    company: document.getElementById("cfg-company").value.trim(),
    recipientName: document.getElementById("cfg-recipient-name").value.trim(),
    recipientEmail: document.getElementById("cfg-recipient-email").value.trim(),
  };
  saveLocal();
  document.getElementById("config-notice").style.display = "none";
  toast("Settings saved");
  renderReportRecipient();
}

// ── UI Helpers ────────────────────────────────────────────────
function showSection(id, btn) {
  document.querySelectorAll(".section").forEach(s => s.classList.remove("active"));
  document.querySelectorAll(".nav-btn").forEach(b => b.classList.remove("active"));
  document.getElementById("sec-" + id).classList.add("active");
  if (btn) btn.classList.add("active");
  if (id === "report") genReport();
}

function toast(msg, type = "") {
  const t = document.getElementById("toast");
  t.textContent = msg;
  t.className = "toast" + (type ? " " + type : "");
  t.classList.add("show");
  clearTimeout(t._timer);
  t._timer = setTimeout(() => t.classList.remove("show"), 3000);
}

function showInstallBanner() {
  const wrap = document.getElementById("install-banner-wrap");
  if (!wrap) return;
  wrap.innerHTML = `<div class="install-banner">
    <div style="font-size:13px;color:#3B6D11;font-weight:500">📲 Install TimeTrack as a desktop app for quick access</div>
    <div style="display:flex;gap:8px">
      <button class="btn btn-primary" onclick="installApp()" style="padding:7px 14px;font-size:13px">Install app</button>
      <button class="btn" onclick="this.closest('.install-banner').parentElement.innerHTML=''" style="padding:7px 10px">✕</button>
    </div>
  </div>`;
}

async function installApp() {
  if (window._installPrompt) {
    window._installPrompt.prompt();
    const { outcome } = await window._installPrompt.userChoice;
    if (outcome === "accepted") {
      document.getElementById("install-banner-wrap").innerHTML = "";
      toast("TimeTrack installed!");
    }
    window._installPrompt = null;
  }
}

// ── Boot ──────────────────────────────────────────────────────
startApp();
