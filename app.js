// ============================================================
//  IGT TimeTrack — app.js (PIN-based, no Microsoft auth)
// ============================================================

let employees = [], clockEntries = [], settings = {};
let selectedEmpKey = null, pinBuffer = "", adminPinBuffer = "";
let editingEmpKey = null, isAdminUnlocked = false;

function init() {
  loadLocal();
  renderEmpGrid();
  renderAll();
  startClock();
  document.getElementById("report-date").value = today();
  updateOnlineStatus();
  window.addEventListener("online", updateOnlineStatus);
  window.addEventListener("offline", updateOnlineStatus);
  // Start auto-save
  startAutoSave();
  // Init SharePoint sync
  initSpSync();
}

function updateOnlineStatus() {
  const el = document.getElementById("offline-indicator");
  if (el) el.style.display = navigator.onLine ? "none" : "inline-block";
}

function loadLocal() {
  employees = JSON.parse(localStorage.getItem("tt_employees") || "[]");
  clockEntries = JSON.parse(localStorage.getItem("tt_entries") || "[]");
  settings = JSON.parse(localStorage.getItem("tt_settings") || "{}");
  if (!employees.length) {
    employees = [
      { key:"e1", name:"Alex Chen",    empId:"EMP001", area:"Production", startTime:"09:00", endTime:"17:00", hours:8, pin:"1234" },
      { key:"e2", name:"Jordan Smith", empId:"EMP002", area:"Warehouse",  startTime:"08:00", endTime:"16:00", hours:8, pin:"2345" },
      { key:"e3", name:"Sam Patel",    empId:"EMP003", area:"Office",     startTime:"07:00", endTime:"15:00", hours:8, pin:"3456" },
    ];
    settings = { adminPin:"0000", areas:"Gaming Assembly,Fintech Assembly,Repair Centre,Warehouse,Operations Support", company:"IGT APAC Manufacturing", recipientName:"Operations Manager", recipientEmail:"manager@igt.com", siteName:"APACManufacturingOperationsTeam", filePath:"General/ATTENDANCE/Attendance.xlsx", defaultLunch:30 };
    saveLocal();
  }
}

function saveLocal() {
  localStorage.setItem("tt_employees", JSON.stringify(employees));
  localStorage.setItem("tt_entries", JSON.stringify(clockEntries));
  localStorage.setItem("tt_settings", JSON.stringify(settings));
}


const AVATAR_COLORS = [["#e6eef9","#0047BB"],["#e6f4ed","#1a7a4a"],["#fff0e8","#c0390b"],["#fff3e0","#e65100"],["#f3e8ff","#6b21a8"],["#e0f2fe","#0369a1"]];
function initials(n) { return (n||"?").split(" ").map(x=>x[0]).join("").toUpperCase().slice(0,2); }
function avatarStyle(i) { const c=AVATAR_COLORS[i%AVATAR_COLORS.length]; return `background:${c[0]};color:${c[1]}`; }

function renderEmpGrid(filter) {
  const g = document.getElementById("emp-grid");
  if (!g) return;
  // Show nothing until user starts typing
  if (filter === undefined || filter === "") {
    g.innerHTML = "";
    return;
  }
  const q = filter.toLowerCase().trim();
  if (!q) { g.innerHTML = ""; return; }
  const filtered = employees.filter(e =>
    e.name.toLowerCase().includes(q) || e.empId.toLowerCase().includes(q) || e.area.toLowerCase().includes(q)
  );
  if (!filtered.length) {
    g.innerHTML = '<div class="emp-list-wrap"><div class="emp-empty">No employees match your search.</div></div>';
    return;
  }
  const rows = filtered.map(e => {
    const i = employees.indexOf(e);
    const active = getClockedInEntry(e.key);
    const done = clockEntries.find(en => en.empKey === e.key && en.date === today() && en.timeOut);
    const statusBadge = active
      ? '<span class="badge badge-green" style="font-size:11px">● Clocked in</span>'
      : done
      ? '<span class="badge badge-gray" style="font-size:11px">✓ Done</span>'
      : '<span class="badge badge-amber" style="font-size:11px">○ Not in</span>';
    return `<div class="emp-list-item" onclick="selectEmployee('${e.key}')">
      <div class="emp-avatar" style="${avatarStyle(i)};width:40px;height:40px;font-size:14px;flex-shrink:0">${initials(e.name)}</div>
      <div class="emp-item-info">
        <div class="emp-item-name">${highlight(e.name, q)}</div>
        <div class="emp-item-meta">${e.empId} · ${e.area} · ${e.startTime}–${e.endTime}</div>
      </div>
      <div class="emp-item-status">${statusBadge}</div>
    </div>`;
  }).join("");
  g.innerHTML = `<div class="emp-list-wrap">${rows}</div>`;
}

function highlight(text, q) {
  if (!q) return text;
  const idx = text.toLowerCase().indexOf(q);
  if (idx === -1) return text;
  return text.slice(0, idx) + `<mark style="background:#fff0c0;border-radius:2px;padding:0 1px">${text.slice(idx, idx + q.length)}</mark>` + text.slice(idx + q.length);
}

function filterEmpList() {
  const q = document.getElementById("emp-search")?.value || "";
  renderEmpGrid(q);
}

function selectEmployee(key) {
  selectedEmpKey = key;
  const emp = employees.find(e => e.key === key);

  // Check if already clocked in
  const active = getClockedInEntry(key);
  if (active) {
    document.getElementById("warning-message").innerHTML =
      `<strong>${emp.name}</strong> is already clocked in since <strong>${active.timeIn}</strong>.<br><br>
      If you continue, you will be clocking <strong>out</strong>.`;
    document.getElementById("clockin-warning-modal").classList.add("open");
    return;
  }

  // Check if already completed a shift today
  const done = clockEntries.find(e => e.empKey === key && e.date === today() && e.timeOut);
  if (done) {
    document.getElementById("done-message").innerHTML =
      `<strong>${emp.name}</strong> has already completed a shift today.<br><br>
      <strong>Clocked in:</strong> ${done.timeIn}<br>
      <strong>Clocked out:</strong> ${done.timeOut}<br>
      <strong>Total hours:</strong> ${calcHours(done.timeIn, done.timeOut)?.toFixed(1) || "—"}h<br><br>
      Do you need to clock in again for a second shift?`;
    document.getElementById("done-warning-modal").classList.add("open");
    return;
  }

  openPinScreen();
}

function openPinScreen() {
  const emp = employees.find(e => e.key === selectedEmpKey);
  const i = employees.indexOf(emp);
  document.getElementById("pin-name").textContent = emp.name;
  document.getElementById("pin-area").textContent = `${emp.area} · ${emp.startTime}–${emp.endTime}`;
  const av = document.getElementById("pin-avatar");
  av.style.cssText = avatarStyle(i) + ";width:60px;height:60px;font-size:22px;font-weight:700;border-radius:50%;display:flex;align-items:center;justify-content:center;margin:0 auto .75rem";
  av.textContent = initials(emp.name);
  pinBuffer = "";
  updatePinDots("pin-dots", 0, "");
  document.getElementById("pin-error").textContent = "";
  showScreen("screen-pin");
}

function proceedToPin() {
  document.getElementById("clockin-warning-modal").classList.remove("open");
  openPinScreen();
}

function closeWarningModal() {
  document.getElementById("clockin-warning-modal").classList.remove("open");
  selectedEmpKey = null;
}

function proceedToPinForce() {
  document.getElementById("done-warning-modal").classList.remove("open");
  openPinScreen();
}

function closeDoneModal() {
  document.getElementById("done-warning-modal").classList.remove("open");
  selectedEmpKey = null;
}

function updatePinDots(id, len, state) {
  document.getElementById(id).querySelectorAll(".pin-dot").forEach((d,i) => {
    d.className = "pin-dot";
    if (i < len) d.classList.add(state==="error"?"error":"filled");
  });
}

function pinPress(d) {
  if (pinBuffer.length>=4) return;
  pinBuffer += d;
  updatePinDots("pin-dots", pinBuffer.length, "");
  document.getElementById("pin-error").textContent = "";
  if (pinBuffer.length===4) setTimeout(verifyPin, 150);
}

function pinDel() {
  if (!pinBuffer.length) return;
  pinBuffer = pinBuffer.slice(0,-1);
  updatePinDots("pin-dots", pinBuffer.length, "");
  document.getElementById("pin-error").textContent = "";
}

function verifyPin() {
  const emp = employees.find(e=>e.key===selectedEmpKey);
  if (!emp) return;
  if (pinBuffer === emp.pin) {
    showScreen("screen-app");
    showSection("clock", document.querySelector(".nav-btn"));
    performClockAction(emp.key);
  } else {
    pinBuffer = "";
    updatePinDots("pin-dots", 4, "error");
    document.getElementById("pin-error").textContent = "Incorrect PIN. Try again.";
    setTimeout(() => { updatePinDots("pin-dots",0,""); document.getElementById("pin-error").textContent=""; }, 1200);
  }
}

function showAdminLogin() {
  adminPinBuffer = "";
  updatePinDots("admin-pin-dots", 0, "");
  document.getElementById("admin-pin-error").textContent = "";
  document.getElementById("admin-modal").classList.add("open");
}

function closeAdminModal() { document.getElementById("admin-modal").classList.remove("open"); adminPinBuffer=""; }

function adminPinPress(d) {
  if (adminPinBuffer.length>=4) return;
  adminPinBuffer += d;
  updatePinDots("admin-pin-dots", adminPinBuffer.length, "");
  document.getElementById("admin-pin-error").textContent = "";
  if (adminPinBuffer.length===4) setTimeout(verifyAdminPin, 150);
}

function adminPinDel() {
  if (!adminPinBuffer.length) return;
  adminPinBuffer = adminPinBuffer.slice(0,-1);
  updatePinDots("admin-pin-dots", adminPinBuffer.length, "");
}

function verifyAdminPin() {
  if (adminPinBuffer === (settings.adminPin||"0000")) {
    closeAdminModal();
    isAdminUnlocked = true;
    document.querySelectorAll(".admin-only").forEach(el => el.style.display = "");
    document.getElementById("admin-signout-btn").style.display = "";
    showScreen("screen-app");
    // Find admin nav button by its onclick attribute
    const adminBtn = [...document.querySelectorAll(".nav-btn")].find(b => b.getAttribute("onclick")?.includes("admin"));
    showSection("admin", adminBtn);
    toast("Admin access granted", "success");
  } else {
    adminPinBuffer = "";
    updatePinDots("admin-pin-dots", 4, "error");
    document.getElementById("admin-pin-error").textContent = "Incorrect PIN.";
    setTimeout(() => { updatePinDots("admin-pin-dots",0,""); document.getElementById("admin-pin-error").textContent=""; }, 1200);
  }
}

function adminSignOut() {
  isAdminUnlocked = false;
  document.querySelectorAll(".admin-only").forEach(el=>el.style.display="none");
  document.getElementById("admin-signout-btn").style.display="none";
  showSection("clock", document.querySelector(".nav-btn"));
  toast("Admin locked");
}

function today() { return new Date().toISOString().slice(0,10); }

function startClock() {
  function tick() {
    const n = new Date();
    const el=document.getElementById("live-clock"); if(el) el.textContent=n.toLocaleTimeString("en-AU",{hour:"2-digit",minute:"2-digit",second:"2-digit",hour12:false});
    const de=document.getElementById("clock-date"); if(de) de.textContent=n.toLocaleDateString("en-AU",{weekday:"long",year:"numeric",month:"long",day:"numeric"});
    const tl=document.getElementById("today-date-label"); if(tl) tl.textContent=n.toLocaleDateString("en-AU",{weekday:"long",day:"numeric",month:"long"});
  }
  tick(); setInterval(tick,1000);
}

function getClockedInEntry(key) { return clockEntries.find(e=>e.empKey===key&&e.date===today()&&e.timeIn&&!e.timeOut); }

function performClockAction(empKey) {
  const entry = getClockedInEntry(empKey);
  const area = document.getElementById("clock-action-area");
  const emp = employees.find(e=>e.key===empKey);
  if (!emp) return;
  const timeStamp = new Date().toLocaleTimeString("en-AU",{hour:"2-digit",minute:"2-digit",hour12:false});

  if (entry) {
    const idx = clockEntries.findIndex(e=>e.empKey===empKey&&e.date===today()&&e.timeIn&&!e.timeOut);
    clockEntries[idx].timeOut = timeStamp;
    saveLocal();
    const hrs = calcHours(clockEntries[idx].timeIn, timeStamp);
    area.innerHTML = `<div style="text-align:center;padding:.5rem">
      <div style="font-size:48px;margin-bottom:.5rem">👋</div>
      <div style="font-size:20px;font-weight:700;color:var(--igt-blue);margin-bottom:.25rem">See you, ${emp.name.split(" ")[0]}!</div>
      <div style="font-size:14px;color:var(--text2);margin-bottom:.25rem">Clocked out at <strong>${timeStamp}</strong></div>
      <div style="font-size:13px;color:var(--text2);margin-bottom:1.25rem">Total time today: <strong>${hrs!==null?hrs.toFixed(1)+"h":"—"}</strong></div>
      <button class="btn btn-primary" onclick="showScreen('screen-login')">← Back to home</button>
    </div>`;
    toast(`${emp.name} clocked out at ${timeStamp}`,"success");
  } else {
    clockEntries.push({ empKey, date:today(), timeIn:timeStamp, timeOut:null, name:emp.name, area:emp.area, empId:emp.empId, stdStart:emp.startTime, stdEnd:emp.endTime, stdHours:emp.hours, lunchMins:emp.lunchMins||settings.defaultLunch||30, status:emp.status||"Permanent" });
    saveLocal();
    const h = new Date().getHours();
    const greet = h<12?"morning":h<17?"afternoon":"evening";
    area.innerHTML = `<div style="text-align:center;padding:.5rem">
      <div style="font-size:48px;margin-bottom:.5rem">✅</div>
      <div style="font-size:20px;font-weight:700;color:var(--igt-blue);margin-bottom:.25rem">Good ${greet}, ${emp.name.split(" ")[0]}!</div>
      <div style="font-size:14px;color:var(--text2);margin-bottom:1.25rem">Clocked in at <strong>${timeStamp}</strong></div>
      <button class="btn btn-primary" onclick="showScreen('screen-login')">← Back to home</button>
    </div>`;
    toast(`${emp.name} clocked in at ${timeStamp}`,"success");
  }
  renderTodayTable(); renderActiveBanner();
  spAutoSync(); // auto push to SharePoint after every clock action
  setTimeout(()=>showScreen("screen-login"), 4000);
}

function forceClockIn(empKey) {
  const emp=employees.find(e=>e.key===empKey);
  const t=new Date().toLocaleTimeString("en-AU",{hour:"2-digit",minute:"2-digit",hour12:false});
  clockEntries.push({empKey,date:today(),timeIn:t,timeOut:null,name:emp.name,area:emp.area,empId:emp.empId,stdStart:emp.startTime,stdEnd:emp.endTime,stdHours:emp.hours,lunchMins:emp.lunchMins||settings.defaultLunch||30,status:emp.status||"Permanent"});
  saveLocal(); performClockAction(empKey);
}

function renderAll() {
  renderTodayTable();
  renderActiveBanner();
  renderEmpList();
  loadSettingsForm();
  renderReportRecipient();
  genReport();
  renderLatestVersion();
}

function renderActiveBanner() {
  const active=clockEntries.filter(e=>e.date===today()&&e.timeIn&&!e.timeOut);
  const el=document.getElementById("active-banner"); if(!el) return;
  if(!active.length){el.innerHTML="";return;}
  el.innerHTML=`<div class="active-banner"><div style="font-size:20px">🟢</div><div><div style="font-weight:700;color:#1a5c38;font-size:13px">${active.length} employee${active.length>1?"s":""} currently clocked in</div><div style="font-size:12px;color:#2d7a50;margin-top:2px">${active.map(e=>`${e.name} (since ${e.timeIn})`).join(" · ")}</div></div></div>`;
}

function renderTodayTable() {
  const entries=clockEntries.filter(e=>e.date===today());
  const wrap=document.getElementById("today-table-wrap"); if(!wrap) return;
  if(!entries.length){wrap.innerHTML='<div class="empty">⏰<br><br>No clock-ins recorded today</div>';return;}
  wrap.innerHTML=`<div style="overflow-x:auto"><table><thead><tr><th>Employee</th><th>Area</th><th>Clock in</th><th>Clock out</th><th>Break</th><th>Net hours</th><th>Status</th></tr></thead><tbody>${entries.map(e=>{
    const hrs=calcHours(e.timeIn,e.timeOut,e.lunchMins);
    return `<tr><td><div class="emp-row"><div class="emp-avatar" style="${avatarStyle(employees.findIndex(x=>x.key===e.empKey))};width:30px;height:30px;font-size:11px">${initials(e.name)}</div><div><div style="font-weight:600">${e.name}</div><div style="font-size:11px;color:var(--text2)">${e.empId}</div></div></div></td><td><span class="tag">${e.area}</span></td><td><strong>${e.timeIn||"—"}</strong></td><td><strong>${e.timeOut||"—"}</strong></td><td style="font-size:12px;color:var(--text2)">${e.lunchMins||0}m</td><td>${hrs!==null?hrs.toFixed(1)+"h":"—"}</td><td>${e.timeOut?'<span class="badge badge-green">✓ Done</span>':'<span class="badge badge-amber">● Active</span>'}</td></tr>`;
  }).join("")}</tbody></table></div>`;
}

function calcHours(tin, tout, lunchMins) {
  if (!tin || !tout) return null;
  const [h1,m1] = tin.split(":").map(Number);
  const [h2,m2] = tout.split(":").map(Number);
  const raw = (h2*60+m2 - h1*60-m1) / 60;
  const lb = (lunchMins !== undefined ? lunchMins : (settings.defaultLunch || 30)) / 60;
  return Math.max(0, raw - lb);
}

function timeDiffStr(t1,t2) {
  if(!t1||!t2) return null;
  const [h1,m1]=t1.split(":").map(Number),[h2,m2]=t2.split(":").map(Number);
  const diff=(h2*60+m2)-(h1*60+m1),sign=diff<0?"-":"+",abs=Math.abs(diff);
  return `${sign}${Math.floor(abs/60)}h ${abs%60}m`;
}

function renderReportRecipient() {
  const el = document.getElementById("report-recipient-info");
  if (!el) return;
  const schedules = getSchedules();
  const active = schedules.filter(s => s.active).length;
  el.innerHTML = active
    ? `📧 <strong>${active} active report schedule${active>1?"s":""}</strong> — go to Admin → Scheduled Reports to manage`
    : `📧 No active report schedules — go to Admin → Scheduled Reports to set up`;
}

function genReport() {
  const dateVal=document.getElementById("report-date")?.value||today();
  const entries=clockEntries.filter(e=>e.date===dateVal);
  const wrap=document.getElementById("report-content"); if(!wrap) return;
  if(!entries.length){wrap.innerHTML=`<div class="card"><div class="empty">📅<br><br>No records for ${dateVal}</div></div>`;return;}
  const rows=entries.map(e=>{
    const actual=calcHours(e.timeIn,e.timeOut,e.lunchMins),diff=actual!==null?actual-e.stdHours:null;
    const inVar=timeDiffStr(e.stdStart,e.timeIn),outVar=e.timeOut?timeDiffStr(e.stdEnd,e.timeOut):null;
    const status=e.timeOut?(diff!==null&&diff>=0?"On time":"Short"):e.timeIn?"In progress":"Absent";
    return {...e,actual,diff,inVar,outVar,status};
  });
  const totalStd=entries.reduce((s,e)=>s+e.stdHours,0),totalActual=rows.reduce((s,r)=>s+(r.actual||0),0),onTime=rows.filter(r=>r.status==="On time").length;
  wrap.innerHTML=`<div class="grid3" style="margin-bottom:1rem">
    <div class="stat-card"><div class="stat-label">Employees</div><div class="stat-value">${entries.length}</div></div>
    <div class="stat-card"><div class="stat-label">Std hours</div><div class="stat-value">${totalStd}h</div></div>
    <div class="stat-card"><div class="stat-label">Actual hours</div><div class="stat-value">${totalActual.toFixed(1)}h</div></div>
  </div>
  <div class="card"><div style="font-weight:700;font-size:15px;margin-bottom:.75rem">Timesheet — ${dateVal} <span class="tag" style="margin-left:6px">${settings.company||""}</span> <span class="badge badge-green" style="margin-left:6px">${onTime}/${entries.length} on time</span></div>
  <div style="overflow-x:auto"><table><thead><tr><th>Employee</th><th>ID</th><th>Area</th><th>Std start</th><th>Actual in</th><th>Variance</th><th>Std end</th><th>Actual out</th><th>Variance</th><th>Std hrs</th><th>Actual hrs</th><th>Diff</th><th>Status</th></tr></thead>
  <tbody>${rows.map(r=>`<tr>
    <td><div class="emp-row"><div class="emp-avatar" style="${avatarStyle(employees.findIndex(x=>x.key===r.empKey))};width:28px;height:28px;font-size:11px">${initials(r.name)}</div><div style="font-weight:600">${r.name}</div></div></td>
    <td style="color:var(--text2)">${r.empId}</td><td><span class="tag">${r.area}</span></td>
    <td>${r.stdStart}</td><td><strong>${r.timeIn||"—"}</strong></td>
    <td class="${r.inVar?(r.inVar.startsWith("+")?"time-diff-neg":"time-diff-pos"):""}">${r.inVar||"—"}</td>
    <td>${r.stdEnd}</td><td><strong>${r.timeOut||"—"}</strong></td>
    <td class="${r.outVar?(r.outVar.startsWith("-")?"time-diff-neg":"time-diff-pos"):""}">${r.outVar||"—"}</td>
    <td>${r.stdHours}h</td><td>${r.actual!==null?r.actual.toFixed(1)+"h":"—"}</td>
    <td class="${r.diff===null?"":r.diff>=0?"time-diff-pos":"time-diff-neg"}">${r.diff===null?"—":(r.diff>=0?"+":"")+r.diff.toFixed(1)+"h"}</td>
    <td>${r.status==="On time"?'<span class="badge badge-green">✓ On time</span>':r.status==="Short"?'<span class="badge badge-red">⚠ Short</span>':r.status==="In progress"?'<span class="badge badge-amber">● Active</span>':'<span class="badge badge-gray">Absent</span>'}</td>
  </tr>`).join("")}</tbody></table></div></div>`;
}

function exportExcel() {
  const dateVal=document.getElementById("report-date")?.value||today();
  const entries=clockEntries.filter(e=>e.date===dateVal);
  const wsData=[
    [`${settings.company||"IGT"} — Daily Timesheet Report`,"","","","","","","","","","",""],
    [`Date: ${dateVal}`,"","Report for:",`${settings.recipientName||""} <${settings.recipientEmail||""}>`, "","","","","","","",""],
    [],
    ["Employee","Employee ID","Work Area","Std Start","Actual Clock In","Start Variance","Std End","Actual Clock Out","End Variance","Std Hours","Actual Hours","Difference","Status"],
    ...entries.map(e=>{
      const actual=calcHours(e.timeIn,e.timeOut,e.lunchMins),diff=actual!==null?+(actual-e.stdHours).toFixed(2):null;
      const inVar=timeDiffStr(e.stdStart,e.timeIn),outVar=e.timeOut?timeDiffStr(e.stdEnd,e.timeOut):null;
      const status=e.timeOut?(diff!==null&&diff>=0?"On time":"Short hours"):e.timeIn?"In progress":"Absent";
      return [e.name,e.empId,e.area,e.stdStart,e.timeIn||"",inVar||"",e.stdEnd,e.timeOut||"",outVar||"",e.stdHours,actual!==null?+actual.toFixed(2):"",diff!==null?diff:"",status];
    }),
    [],
    ["","","","","","","","","",`=SUM(J5:J${entries.length+4})`,`=SUM(K5:K${entries.length+4})`,`=SUM(L5:L${entries.length+4})`,""],
    ["","","","","","","","","","Total Std Hrs","Total Actual Hrs","Total Diff",""],
  ];
  const wb=XLSX.utils.book_new(),ws=XLSX.utils.aoa_to_sheet(wsData);
  ws["!cols"]=[{wch:22},{wch:12},{wch:16},{wch:11},{wch:16},{wch:14},{wch:11},{wch:16},{wch:14},{wch:11},{wch:13},{wch:11},{wch:14}];
  ws["!merges"]=[{s:{r:0,c:0},e:{r:0,c:12}}];
  XLSX.utils.book_append_sheet(wb,ws,"Daily Timesheet");
  const allData=[["Employee","Employee ID","Area","Status","Date","Time In","Time Out","Lunch Break","Std Hours","Net Hours","Difference"],...clockEntries.map(e=>{const a=calcHours(e.timeIn,e.timeOut,e.lunchMins);return[e.name,e.empId,e.area,e.status||"",e.date,e.timeIn||"",e.timeOut||"",`${e.lunchMins||0}m`,e.stdHours,a!==null?+a.toFixed(2):"",a!==null?+(a-e.stdHours).toFixed(2):""];})];
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(allData),"Full History");
  XLSX.writeFile(wb,`IGT_Timesheet_${dateVal}.xlsx`);
  toast("Excel report downloaded!");
}

function renderEmpList() {
  const el=document.getElementById("emp-list"); if(!el) return;
  if(!employees.length){el.innerHTML='<div class="card"><div class="empty">No employees added yet</div></div>';return;}
  el.innerHTML=employees.map((e,i)=>`<div class="card" style="margin-bottom:8px;padding:1rem"><div style="display:flex;align-items:center;gap:10px">
    <div class="emp-avatar" style="${avatarStyle(i)}">${initials(e.name)}</div>
    <div style="flex:1"><div style="font-weight:700;font-size:14px">${e.name} <span class="tag">${e.empId}</span></div>
    <div style="font-size:12px;color:var(--text2)">${e.area} · ${e.startTime}–${e.endTime} · ${e.hours}h/day · PIN: ${"●".repeat(e.pin?.length||4)}${e.faceDescriptor?' · <span style="color:#1a7a4a;font-weight:600">✅ Face</span>':''}</div></div>
    <div style="display:flex;gap:6px">
      <button class="btn" onclick="openEmpModal('${e.key}')" style="padding:6px 10px">✏</button>
      <button class="btn btn-danger" onclick="deleteEmp('${e.key}')" style="padding:6px 10px">🗑</button>
    </div>
  </div></div>`).join("");
}

function openEmpModal(key) {
  editingEmpKey=key||null;
  const areas=(settings.areas||"Gaming Assembly,Fintech Assembly,Repair Centre,Warehouse,Operations Support").split(",").map(a=>a.trim());
  document.getElementById("emp-area").innerHTML=areas.map(a=>`<option>${a}</option>`).join("");
  window._pendingFaceDescriptor = null;
  window._clearFaceDescriptor = false;
  stopFaceEnroll();
  const captureBtn = document.getElementById("enroll-capture-btn");
  if (captureBtn) captureBtn.style.display = "none";
  document.getElementById("enroll-msg").textContent = "";

  if(key){
    const e=employees.find(x=>x.key===key);
    document.getElementById("modal-title").textContent="Edit Employee";
    document.getElementById("emp-name").value=e.name;
    document.getElementById("emp-id-field").value=e.empId;
    document.getElementById("emp-area").value=e.area;
    document.getElementById("emp-start").value=e.startTime;
    document.getElementById("emp-end").value=e.endTime;
    document.getElementById("emp-hours").value=e.hours;
    document.getElementById("emp-pin").value=e.pin||"";
    document.getElementById("emp-lunch").value=e.lunchMins||30;
    // Show face enroll status
    const hasFace = !!e.faceDescriptor;
    document.getElementById("enroll-status-badge").innerHTML = hasFace
      ? '<span class="face-enrolled-badge">✅ Face enrolled</span>'
      : '<span style="font-size:12px;color:var(--text2)">No face enrolled</span>';
    document.getElementById("enroll-clear-btn").style.display = hasFace ? "" : "none";
  } else {
    document.getElementById("modal-title").textContent="Add Employee";
    ["emp-name","emp-id-field","emp-pin"].forEach(id=>document.getElementById(id).value="");
    document.getElementById("emp-start").value="09:00";
    document.getElementById("emp-end").value="17:00";
    document.getElementById("emp-hours").value="8";
    document.getElementById("emp-lunch").value=settings.defaultLunch||30;
    document.getElementById("enroll-status-badge").innerHTML = '<span style="font-size:12px;color:var(--text2)">No face enrolled</span>';
    document.getElementById("enroll-clear-btn").style.display = "none";
  }
  document.getElementById("emp-modal").classList.add("open");
}

function closeEmpModal() { document.getElementById("emp-modal").classList.remove("open"); }

function saveEmployee() {
  const name=document.getElementById("emp-name").value.trim(),empId=document.getElementById("emp-id-field").value.trim();
  const area=document.getElementById("emp-area").value,startTime=document.getElementById("emp-start").value;
  const endTime=document.getElementById("emp-end").value,hours=parseFloat(document.getElementById("emp-hours").value);
  const pin=document.getElementById("emp-pin").value.trim();
  const lunchMins=parseInt(document.getElementById("emp-lunch").value)||0;
  const status=document.getElementById("emp-status-field")?.value||"Permanent";
  if(!name||!empId){toast("Name and ID required","error");return;}
  if(!/^\d{4}$/.test(pin)){toast("PIN must be exactly 4 digits","error");return;}
  const emp = { key: editingEmpKey ? employees.find(e=>e.key===editingEmpKey).key : "e"+Date.now(), name, empId, area, startTime, endTime, hours, pin, lunchMins, status };
  // Handle face descriptor
  if (window._pendingFaceDescriptor) emp.faceDescriptor = window._pendingFaceDescriptor;
  else if (!window._clearFaceDescriptor && editingEmpKey) {
    const existing = employees.find(e=>e.key===editingEmpKey);
    if (existing?.faceDescriptor) emp.faceDescriptor = existing.faceDescriptor;
  }
  window._pendingFaceDescriptor = null;
  window._clearFaceDescriptor = false;

  if(editingEmpKey){ const idx=employees.findIndex(e=>e.key===editingEmpKey); employees[idx]=emp; }
  else employees.push(emp);
  saveLocal();closeEmpModal();renderEmpList();renderEmpGrid();
  toast(editingEmpKey?"Employee updated":"Employee added","success");
}

function deleteEmp(key) {
  if(!confirm("Remove this employee?")) return;
  employees=employees.filter(e=>e.key!==key);
  saveLocal();renderEmpList();renderEmpGrid();toast("Employee removed");
}

function loadSettingsForm() {
  document.getElementById("cfg-admin-pin").value = settings.adminPin || "0000";
  document.getElementById("cfg-site").value = settings.siteUrl || "https://igtplc.sharepoint.com/sites/APACManufacturingOperationsTeam";
  document.getElementById("cfg-clientid").value = settings.spClientId || "";
  document.getElementById("cfg-areas").value = settings.areas || "";
  document.getElementById("cfg-company").value = settings.company || "";
  document.getElementById("cfg-lunch").value = settings.defaultLunch || 30;
  // Backup settings
  const bpEl = document.getElementById("cfg-backup-prefix");
  const aiEl = document.getElementById("cfg-autosave-interval");
  if (bpEl) bpEl.value = settings.backupPrefix || "IGT_Attendance";
  if (aiEl) aiEl.value = settings.autoSaveInterval || 60;
  const crEl = document.getElementById("cfg-csv-range");
  if (crEl) crEl.value = settings.csvRange || "all";
  updateBackupFilenamePreview();
  updateLastBackupTime();
  renderSchedulesList();
  renderLatestVersion();
}

function saveSettings() {
  const adminPin = document.getElementById("cfg-admin-pin").value.trim();
  if (!/^\d{4}$/.test(adminPin)) { toast("Admin PIN must be 4 digits", "error"); return; }
  settings = { ...settings,
    adminPin,
    siteUrl: document.getElementById("cfg-site").value.trim(),
    spClientId: document.getElementById("cfg-clientid").value.trim(),
    areas: document.getElementById("cfg-areas").value,
    company: document.getElementById("cfg-company").value.trim(),
    defaultLunch: parseInt(document.getElementById("cfg-lunch").value) || 30,
  };
  saveLocal();
  toast("Settings saved", "success");
  renderReportRecipient();
}

function showSection(id,btn) {
  document.querySelectorAll(".section").forEach(s=>s.classList.remove("active"));
  document.querySelectorAll(".nav-btn").forEach(b=>b.classList.remove("active"));
  document.getElementById("sec-"+id).classList.add("active");
  if(btn) btn.classList.add("active");
  if(id==="report") genReport();
  if(id==="clock") renderTodayTable();
}

function toast(msg,type="") {
  const t=document.getElementById("toast");
  t.textContent=msg;t.className="toast"+(type?" "+type:"");t.classList.add("show");
  clearTimeout(t._timer);t._timer=setTimeout(()=>t.classList.remove("show"),3000);
}

// ── Excel Import ──────────────────────────────────────────────
let importQueue = [];

function handleExcelImport(event) {
  const file = event.target.files[0];
  if (!file) return;
  event.target.value = "";
  const reader = new FileReader();
  reader.onload = (e) => parseImportFile(e.target.result);
  reader.readAsArrayBuffer(file);
}

function parseTime(str) {
  if (!str) return null;
  const s = str.toString().trim().toLowerCase().replace(" to ", " - ");
  const parts = s.split(" - ");
  if (parts.length !== 2) return null;
  function fmt(t) {
    t = t.trim();
    const m = t.match(/^(\d+)(?::(\d+))?\s*(am|pm)$/);
    if (!m) return null;
    let hr = parseInt(m[1]), mn = parseInt(m[2] || 0), ap = m[3];
    if (ap === "pm" && hr !== 12) hr += 12;
    if (ap === "am" && hr === 12) hr = 0;
    return `${String(hr).padStart(2,"0")}:${String(mn).padStart(2,"0")}`;
  }
  return [fmt(parts[0]), fmt(parts[1])];
}

function calcStdHours(start, end) {
  if (!start || !end) return 8;
  const [h1,m1] = start.split(":").map(Number);
  const [h2,m2] = end.split(":").map(Number);
  return +((h2*60+m2-h1*60-m1)/60).toFixed(1);
}

function parseImportFile(buffer) {
  let wb;
  try { wb = XLSX.read(buffer, { type:"array" }); }
  catch(e) { toast("Could not read file: " + e.message, "error"); return; }

  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { defval:"" });

  const errors = [], parsed = [];
  const seenNames = new Set(), seenIds = new Set();

  rows.forEach((row, idx) => {
    const rowNum = idx + 2;
    const name   = (row["Employee Name"] || row["Name"] || "").toString().trim();
    const area   = (row["Area"] || row["Work Area"] || "").toString().trim();
    const status = (row["Employment Status"] || row["Status"] || "Permanent").toString().trim();
    const monThu = (row["Monday to Thursday"] || row["Mon-Thu"] || row["Shift"] || "").toString().trim();
    const friday = (row["Friday"] || monThu).toString().trim();

    if (!name)   { errors.push(`Row ${rowNum}: Missing employee name`); return; }
    if (!area)   { errors.push(`Row ${rowNum}: Missing area for "${name}"`); return; }
    if (!monThu) { errors.push(`Row ${rowNum}: Missing shift time for "${name}"`); return; }

    const nameKey = name.toLowerCase();
    if (seenNames.has(nameKey)) { errors.push(`Row ${rowNum}: Duplicate in file — "${name}"`); return; }
    seenNames.add(nameKey);

    const monThuTimes = parseTime(monThu);
    if (!monThuTimes || !monThuTimes[0] || !monThuTimes[1]) {
      errors.push(`Row ${rowNum}: Cannot parse shift time "${monThu}" for "${name}"`); return;
    }
    const friTimes = parseTime(friday);

    // Auto-generate employee ID
    const parts = name.replace(/,/g,"").split(/\s+/);
    let empId = (parts[parts.length-1].slice(0,3) + parts[0].slice(0,2)).toUpperCase();
    const baseId = empId; let sfx = 1;
    while (seenIds.has(empId)) empId = baseId + sfx++;
    seenIds.add(empId);

    const existing = employees.find(e => e.name.toLowerCase() === nameKey);
    parsed.push({
      key: existing ? existing.key : "e" + Date.now() + Math.random().toString(36).slice(2,6),
      name, area, status,
      empId: existing ? existing.empId : empId,
      startTime: monThuTimes[0], endTime: monThuTimes[1],
      friStart: friTimes ? friTimes[0] : monThuTimes[0],
      friEnd:   friTimes ? friTimes[1] : monThuTimes[1],
      hours: calcStdHours(monThuTimes[0], monThuTimes[1]),
      pin: existing ? existing.pin : "0000",
      isExisting: !!existing, isNew: !existing,
    });
  });

  showImportModal(parsed, errors);
}

function showImportModal(parsed, errors) {
  importQueue = parsed;
  const newCount = parsed.filter(r=>r.isNew).length;
  const overwriteCount = parsed.filter(r=>r.isExisting).length;
  const hasErrors = errors.length > 0;

  // Summary badges
  let sumHtml = `<div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:.5rem">`;
  sumHtml += `<span class="badge badge-blue" style="font-size:13px">📋 ${parsed.length} employees found</span>`;
  if (newCount)       sumHtml += `<span class="badge badge-green" style="font-size:13px">✚ ${newCount} new</span>`;
  if (overwriteCount) sumHtml += `<span class="badge badge-amber" style="font-size:13px">⚠ ${overwriteCount} will overwrite</span>`;
  if (errors.length)  sumHtml += `<span class="badge badge-red" style="font-size:13px">✗ ${errors.length} error${errors.length>1?"s":""}</span>`;
  sumHtml += `</div>`;
  document.getElementById("import-summary").innerHTML = sumHtml;

  // Errors block
  document.getElementById("import-errors").innerHTML = hasErrors ? `
    <div style="background:#fce8e8;border:1px solid #f5c1c1;border-radius:var(--radius);padding:1rem;margin-bottom:.75rem">
      <div style="font-weight:700;color:#a32d2d;margin-bottom:.5rem">⛔ Errors found — fix these in your Excel file and re-upload:</div>
      ${errors.map(e=>`<div style="font-size:13px;color:#a32d2d;padding:2px 0">• ${e}</div>`).join("")}
    </div>` : "";

  // Preview table
  document.getElementById("import-table-wrap").innerHTML = parsed.length ? `
    <div style="font-size:12px;font-weight:700;color:var(--text2);text-transform:uppercase;letter-spacing:.4px;margin-bottom:.5rem">Import preview</div>
    <div style="overflow-x:auto"><table>
      <thead><tr><th>Status</th><th>Name</th><th>Area</th><th>Employment</th><th>Mon–Thu</th><th>Friday</th><th>Hrs</th><th>PIN</th></tr></thead>
      <tbody>${parsed.map(r=>`<tr style="${r.isExisting?"background:#fff8e1":""}">
        <td>${r.isExisting
          ? '<span class="badge badge-amber" style="font-size:11px">⚠ Overwrite</span>'
          : '<span class="badge badge-green" style="font-size:11px">✚ New</span>'}</td>
        <td style="font-weight:600;white-space:nowrap">${r.name}</td>
        <td><span class="tag">${r.area}</span></td>
        <td style="font-size:12px;color:var(--text2)">${r.status}</td>
        <td style="font-size:12px;white-space:nowrap">${r.startTime}–${r.endTime}</td>
        <td style="font-size:12px;white-space:nowrap">${r.friStart}–${r.friEnd}</td>
        <td style="font-size:12px">${r.hours}h</td>
        <td style="font-size:12px;color:var(--text3)">${r.isExisting?"unchanged":"0000"}</td>
      </tr>`).join("")}</tbody>
    </table></div>
    ${overwriteCount ? `<div style="font-size:12px;color:#7a5500;margin-top:.5rem;padding:.6rem .75rem;background:#fff8e1;border-radius:var(--radius-sm)">⚠ <strong>${overwriteCount} existing employee${overwriteCount>1?"s":""}</strong> will be overwritten. Their PINs and attendance history will be preserved.</div>` : ""}
    <div style="font-size:12px;color:var(--text2);margin-top:.5rem">💡 New employees get default PIN <strong>0000</strong> — update each PIN in Admin after import.</div>` : "";

  document.getElementById("import-confirm-btn").style.display = (!hasErrors && parsed.length) ? "" : "none";
  document.getElementById("import-modal").classList.add("open");
}

function confirmImport() {
  let added = 0, updated = 0;
  importQueue.forEach(r => {
    const idx = employees.findIndex(e => e.key === r.key);
    const emp = { key:r.key, name:r.name, empId:r.empId, area:r.area, startTime:r.startTime, endTime:r.endTime, hours:r.hours, pin:r.pin };
    if (idx >= 0) { employees[idx] = emp; updated++; }
    else          { employees.push(emp); added++; }
  });
  saveLocal();
  closeImportModal();
  renderEmpList();
  renderEmpGrid();
  toast(`✓ Imported: ${added} new, ${updated} updated`, "success");
  importQueue = [];
}

function closeImportModal() {
  document.getElementById("import-modal").classList.remove("open");
  importQueue = [];
}

// ── Scheduled Reports ─────────────────────────────────────────
let editingScheduleKey = null;

function getSchedules() {
  return JSON.parse(localStorage.getItem("tt_schedules") || "[]");
}
function saveSchedules(arr) {
  localStorage.setItem("tt_schedules", JSON.stringify(arr));
}

const DAY_NAMES = ["Sun","Mon","Tue","Wed","Thu","Fri","Sat"];
const DAY_IDS   = ["sun","mon","tue","wed","thu","fri","sat"];

function renderSchedulesList() {
  const el = document.getElementById("report-schedules-list");
  if (!el) return;
  const schedules = getSchedules();
  if (!schedules.length) {
    el.innerHTML = '<div style="font-size:13px;color:var(--text2);padding:.5rem 0;margin-bottom:.5rem">No schedules yet. Add one below.</div>';
    return;
  }
  el.innerHTML = schedules.map((s,i) => {
    const days = DAY_IDS.map((d,idx) => s.days.includes(d) ? `<span style="font-weight:600">${DAY_NAMES[idx]}</span>` : `<span style="color:var(--text3)">${DAY_NAMES[idx]}</span>`).join(" ");
    const areas = s.areas?.length ? s.areas.join(", ") : "All areas";
    const empType = s.empType === "permanent" ? "Permanent only" : s.empType === "contractor" ? "Contractors only" : "All staff";
    return `<div class="card" style="margin-bottom:8px;padding:.9rem 1rem">
      <div style="display:flex;align-items:flex-start;gap:10px">
        <div style="flex:1">
          <div style="font-weight:700;font-size:14px">${s.name} <span style="font-weight:400;color:var(--text2);font-size:12px">&lt;${s.email}&gt;</span>
            ${s.active ? '<span class="badge badge-green" style="font-size:11px;margin-left:6px">● Active</span>' : '<span class="badge badge-gray" style="font-size:11px;margin-left:6px">Paused</span>'}
          </div>
          ${s.subject ? `<div style="font-size:12px;font-weight:600;color:var(--igt-blue);margin-top:2px">📧 ${s.subject}</div>` : ""}
          <div style="font-size:12px;color:var(--text2);margin-top:3px">⏰ ${s.time} · ${days}</div>
          <div style="font-size:12px;color:var(--text2);margin-top:2px">📂 ${areas} · 👤 ${empType}</div>
        </div>
        <div style="display:flex;gap:6px;flex-shrink:0">
          <button class="btn btn-success" onclick="testSendNow(${i})" style="padding:5px 9px;font-size:12px">▶ Send now</button>
          <button class="btn" onclick="openScheduleModal(${i})" style="padding:5px 9px">✏</button>
          <button class="btn btn-danger" onclick="deleteSchedule(${i})" style="padding:5px 9px">🗑</button>
        </div>
      </div>
    </div>`;
  }).join("");
}

function openScheduleModal(idx) {
  editingScheduleKey = (idx !== undefined) ? idx : null;
  const areas = (settings.areas || "").split(",").map(a => a.trim()).filter(Boolean);
  const s = (idx !== undefined) ? getSchedules()[idx] : null;

  // Populate area checkboxes
  document.getElementById("sch-areas-wrap").innerHTML = areas.length
    ? areas.map(a => `<label style="text-transform:none;font-size:13px;display:flex;align-items:center;gap:4px;font-weight:400"><input type="checkbox" class="sch-area-cb" value="${a}" ${s?.areas?.includes(a)?"checked":""}> ${a}</label>`).join("")
    : '<div style="font-size:12px;color:var(--text2)">Save work areas in settings first</div>';

  document.getElementById("sch-name").value    = s?.name    || "";
  document.getElementById("sch-email").value   = s?.email   || "";
  document.getElementById("sch-subject").value = s?.subject || "";
  document.getElementById("sch-time").value    = s?.time    || "17:30";
  document.getElementById("sch-emptype").value = s?.empType || "all";
  document.getElementById("sch-active").value  = s?.active !== false ? "1" : "0";
  DAY_IDS.forEach(d => { document.getElementById("sch-"+d).checked = s ? s.days.includes(d) : ["mon","tue","wed","thu","fri"].includes(d); });
  document.getElementById("schedule-modal-title").textContent = s ? "Edit Report Schedule" : "Add Report Schedule";
  document.getElementById("schedule-modal").classList.add("open");
}

function closeScheduleModal() {
  document.getElementById("schedule-modal").classList.remove("open");
  editingScheduleKey = null;
}

function saveSchedule() {
  const name    = document.getElementById("sch-name").value.trim();
  const email   = document.getElementById("sch-email").value.trim();
  const subject = document.getElementById("sch-subject").value.trim();
  const time    = document.getElementById("sch-time").value;
  if (!name)  { toast("Recipient name required","error"); return; }
  if (!email) { toast("Recipient email required","error"); return; }
  const days = DAY_IDS.filter(d => document.getElementById("sch-"+d).checked);
  if (!days.length) { toast("Select at least one day","error"); return; }
  const areas   = [...document.querySelectorAll(".sch-area-cb:checked")].map(cb => cb.value);
  const empType = document.getElementById("sch-emptype").value;
  const active  = document.getElementById("sch-active").value === "1";

  const schedules = getSchedules();
  const entry = { key:"s"+Date.now(), name, email, subject, time, days, areas, empType, active };
  if (editingScheduleKey !== null) schedules[editingScheduleKey] = { ...schedules[editingScheduleKey], ...entry };
  else schedules.push(entry);
  saveSchedules(schedules);
  closeScheduleModal();
  renderSchedulesList();
  toast(editingScheduleKey !== null ? "Schedule updated" : "Schedule added", "success");
}

function deleteSchedule(idx) {
  if (!confirm("Remove this report schedule?")) return;
  const schedules = getSchedules();
  schedules.splice(idx, 1);
  saveSchedules(schedules);
  renderSchedulesList();
  toast("Schedule removed");
}

// Check schedules every minute and auto-export if time matches
function checkSchedules() {
  const now = new Date();
  const dayKey = DAY_IDS[now.getDay()];
  const timeStr = `${String(now.getHours()).padStart(2,"0")}:${String(now.getMinutes()).padStart(2,"0")}`;
  getSchedules().forEach(s => {
    if (!s.active) return;
    if (!s.days.includes(dayKey)) return;
    if (s.time !== timeStr) return;
    runScheduledExport(s);
  });
}

// ── Barcode Scanner ───────────────────────────────────────────
let codeReader = null;
let scannerActive = false;

async function startScanner() {
  const wrap = document.getElementById("scanner-wrap");
  const stopBtn = document.getElementById("stop-scan-btn");
  const video = document.getElementById("scanner-video");

  if (scannerActive) return;

  // Check ZXing loaded
  if (typeof ZXing === "undefined") {
    toast("Scanner library not loaded. Try refreshing the page.", "error");
    return;
  }

  try {
    await navigator.mediaDevices.getUserMedia({ video: { facingMode: "environment" } });
    wrap.style.display = "block";
    stopBtn.style.display = "block";
    scannerActive = true;

    codeReader = new ZXing.BrowserQRCodeReader();
    const devices = await ZXing.BrowserCodeReader.listVideoInputDevices();
    const backCamera = devices.find(d =>
      d.label.toLowerCase().includes("back") ||
      d.label.toLowerCase().includes("rear") ||
      d.label.toLowerCase().includes("environment")
    );
    const deviceId = backCamera ? backCamera.deviceId : (devices[0]?.deviceId);

    codeReader.decodeFromVideoDevice(deviceId, video, (result, err) => {
      if (result) {
        const text = result.getText();
        handleBarcodeScan(text);
      }
    });
  } catch(e) {
    toast("Camera not available: " + e.message, "error");
    stopScanner();
  }
}

function stopScanner() {
  if (codeReader) {
    codeReader.reset();
    codeReader = null;
  }
  scannerActive = false;
  document.getElementById("scanner-wrap").style.display = "none";
  document.getElementById("stop-scan-btn").style.display = "none";
}

function handleBarcodeScan(text) {
  // QR code contains employee key prefixed with "IGT-EMP:"
  stopScanner();
  if (!text.startsWith("IGT-EMP:")) {
    toast("Invalid barcode — not an IGT employee code", "error");
    return;
  }
  const empKey = text.replace("IGT-EMP:", "");
  const emp = employees.find(e => e.key === empKey);
  if (!emp) {
    toast("Employee not found — barcode may be outdated", "error");
    return;
  }
  // Success — beep and proceed
  playBeep();
  selectEmployee(empKey);
}

function playBeep() {
  try {
    const ctx = new (window.AudioContext || window.webkitAudioContext)();
    const osc = ctx.createOscillator();
    const gain = ctx.createGain();
    osc.connect(gain);
    gain.connect(ctx.destination);
    osc.frequency.value = 880;
    osc.type = "sine";
    gain.gain.setValueAtTime(0.3, ctx.currentTime);
    gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + 0.15);
    osc.start(ctx.currentTime);
    osc.stop(ctx.currentTime + 0.15);
  } catch(e) {}
}

// ── QR Code / Barcode Generation ─────────────────────────────
async function showBarcodes() {
  const grid = document.getElementById("barcode-grid");
  grid.innerHTML = "";
  document.getElementById("barcode-modal").classList.add("open");

  // Small delay to let modal render
  await new Promise(r => setTimeout(r, 100));

  if (!employees.length) {
    grid.innerHTML = '<div style="font-size:13px;color:var(--text2)">No employees found. Add employees first.</div>';
    return;
  }

  employees.forEach((emp) => {
    const card = document.createElement("div");
    card.className = "qr-card";

    const qrWrap = document.createElement("div");
    qrWrap.style.cssText = "width:120px;height:120px;margin:0 auto";
    card.appendChild(qrWrap);

    const nameEl = document.createElement("div");
    nameEl.className = "qr-card-name";
    nameEl.textContent = emp.name;
    card.appendChild(nameEl);

    const areaEl = document.createElement("div");
    areaEl.className = "qr-card-area";
    areaEl.textContent = emp.area;
    card.appendChild(areaEl);

    const idEl = document.createElement("div");
    idEl.className = "qr-card-id";
    idEl.textContent = emp.empId;
    card.appendChild(idEl);

    const brandEl = document.createElement("div");
    brandEl.style.cssText = "font-size:9px;color:var(--text3);margin-top:4px;font-family:monospace";
    brandEl.textContent = "IGT TimeTrack";
    card.appendChild(brandEl);

    grid.appendChild(card);

    // Generate QR code into the div
    try {
      new QRCode(qrWrap, {
        text: `IGT-EMP:${emp.key}`,
        width: 120,
        height: 120,
        colorDark: "#003087",
        colorLight: "#ffffff",
        correctLevel: QRCode.CorrectLevel.M,
      });
    } catch(e) {
      qrWrap.innerHTML = `<div style="font-size:11px;color:red;padding:10px">QR error</div>`;
      console.error("QR error for", emp.name, e);
    }
  });
}

function closeBarcodesModal() {
  document.getElementById("barcode-modal").classList.remove("open");
}

// Stop scanner when leaving login screen
function showScreen(id) {
  if (id !== "screen-login") { stopScanner(); stopFaceId(); }
  document.querySelectorAll(".screen").forEach(s => s.classList.remove("active"));
  document.getElementById(id).classList.add("active");
  if (id === "screen-login") {
    renderEmpGrid();
    pinBuffer = "";
    selectedEmpKey = null;
    const s = document.getElementById("emp-search");
    if (s) { s.value = ""; }
  }
  if (id === "screen-app") renderAll();
}

// ── CSV & Backup Engine ───────────────────────────────────────
let autoSaveTimer = null;

function getBackupPrefix() {
  return settings.backupPrefix || "IGT_Attendance";
}

function getCSVFilename(range) {
  const prefix = getBackupPrefix();
  if (range === "today") return `${prefix}_${today()}.csv`;
  if (range === "week") return `${prefix}_Week_${today()}.csv`;
  if (range === "month") return `${prefix}_Month_${today().slice(0,7)}.csv`;
  return `${prefix}_All_${today()}.csv`;
}

function filterByRange(entries, range) {
  const d = new Date();
  if (range === "today") return entries.filter(e => e.date === today());
  if (range === "week") {
    const startOfWeek = new Date(d);
    startOfWeek.setDate(d.getDate() - d.getDay() + 1);
    const weekStart = startOfWeek.toISOString().slice(0,10);
    return entries.filter(e => e.date >= weekStart);
  }
  if (range === "month") {
    const monthStart = today().slice(0,7);
    return entries.filter(e => e.date.startsWith(monthStart));
  }
  return entries; // all
}

function buildCSV(entries) {
  const headers = [
    "Employee Name", "Employee ID", "Work Area", "Employment Status",
    "Date", "Std Start", "Actual Clock In", "Start Variance",
    "Std End", "Actual Clock Out", "End Variance",
    "Std Hours", "Lunch Break (min)", "Net Hours", "Difference", "Status"
  ];

  const rows = entries.map(e => {
    const actual = calcHours(e.timeIn, e.timeOut, e.lunchMins);
    const diff = actual !== null ? +(actual - e.stdHours).toFixed(2) : "";
    const inVar = timeDiffStr(e.stdStart, e.timeIn) || "";
    const outVar = e.timeOut ? (timeDiffStr(e.stdEnd, e.timeOut) || "") : "";
    const status = e.timeOut
      ? (diff !== "" && diff >= 0 ? "On time" : "Short hours")
      : e.timeIn ? "In progress" : "Absent";
    return [
      e.name, e.empId, e.area, e.status || "Permanent",
      e.date, e.stdStart, e.timeIn || "", inVar,
      e.stdEnd, e.timeOut || "", outVar,
      e.stdHours, e.lunchMins || 0,
      actual !== null ? +actual.toFixed(2) : "",
      diff !== "" ? diff : "", status
    ].map(v => `"${String(v).replace(/"/g, '""')}"`).join(",");
  });

  return [headers.map(h => `"${h}"`).join(","), ...rows].join("\r\n");
}

function downloadFile(content, filename, mimeType) {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function downloadCSV(rangeOverride) {
  const range = rangeOverride || document.getElementById("cfg-csv-range")?.value || "all";
  const entries = filterByRange(clockEntries, range);
  if (!entries.length) { toast("No attendance records for selected range", "error"); return; }
  const csv = buildCSV(entries);
  const filename = getCSVFilename(range);
  downloadFile(csv, filename, "text/csv;charset=utf-8;");
  const now = new Date().toLocaleTimeString("en-AU", { hour:"2-digit", minute:"2-digit" });
  backupLog(`✓ CSV downloaded — ${entries.length} records (${range}) at ${now}`);
  toast(`✓ CSV saved — ${entries.length} records`, "success");
}

function autoBackup() {
  // Auto-save: download CSV for all records
  const entries = clockEntries;
  if (!entries.length) return;
  const csv = buildCSV(entries);
  const filename = getCSVFilename("all");
  downloadFile(csv, filename, "text/csv;charset=utf-8;");
  localStorage.setItem("tt_last_backup", new Date().toISOString());
  const now = new Date().toLocaleTimeString("en-AU", { hour:"2-digit", minute:"2-digit" });
  backupLog(`✓ Auto-saved CSV at ${now} — ${entries.length} records`);
  updateLastBackupTime();
}

function downloadFullBackup() {
  const data = {
    version: APP_VERSION,
    exportedAt: new Date().toISOString(),
    employees, clockEntries,
    settings: { ...settings, adminPin: "****" },
    schedules: getSchedules(),
  };
  const prefix = getBackupPrefix();
  downloadFile(JSON.stringify(data, null, 2), `${prefix}_FullBackup_${today()}.json`, "application/json");
  backupLog(`✓ Full JSON backup downloaded at ${new Date().toLocaleTimeString("en-AU",{hour:"2-digit",minute:"2-digit"})}`);
  toast("✓ Full backup downloaded", "success");
}

function startAutoSave() {
  if (autoSaveTimer) clearInterval(autoSaveTimer);
  const mins = parseInt(settings.autoSaveInterval || 60);
  if (!mins) { backupLog("Auto-save disabled"); return; }
  autoSaveTimer = setInterval(() => autoBackup(), mins * 60 * 1000);
  backupLog(`Auto-save active — CSV every ${mins} min`);
}

function restoreFromBackup(event) {
  const file = event.target.files[0];
  if (!file) return;
  event.target.value = "";
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = JSON.parse(e.target.result);
      if (!data.employees || !data.clockEntries) { toast("Invalid backup file", "error"); return; }
      const msg = `Restore backup from ${data.exportedAt?.slice(0,10) || "unknown"}?\n\n• ${data.employees.length} employees\n• ${data.clockEntries.length} attendance records\n\nThis will overwrite current data.`;
      if (!confirm(msg)) return;
      const currentPin = settings.adminPin;
      employees = data.employees || [];
      clockEntries = data.clockEntries || [];
      settings = { ...settings, ...data.settings, adminPin: currentPin };
      if (data.schedules) localStorage.setItem("tt_schedules", JSON.stringify(data.schedules));
      saveLocal();
      renderAll();
      loadSettingsForm();
      backupLog(`✓ Restored — ${employees.length} employees, ${clockEntries.length} records`);
      toast(`✓ Restored ${employees.length} employees and ${clockEntries.length} records`, "success");
    } catch(e) {
      toast("Failed to read backup: " + e.message, "error");
      backupLog("✗ Restore failed: " + e.message);
    }
  };
  reader.readAsText(file);
}

function saveBackupSettings() {
  settings.backupPrefix = document.getElementById("cfg-backup-prefix").value.trim() || "IGT_Attendance";
  settings.autoSaveInterval = parseInt(document.getElementById("cfg-autosave-interval").value) || 0;
  settings.csvRange = document.getElementById("cfg-csv-range").value || "all";
  saveLocal();
  startAutoSave();
  updateBackupFilenamePreview();
  toast("Backup settings saved", "success");
}

function updateBackupFilenamePreview() {
  const prefix = document.getElementById("cfg-backup-prefix")?.value.trim() || "IGT_Attendance";
  const el = document.getElementById("backup-filename-preview");
  if (el) el.textContent = `${prefix}_${today()}.csv`;
}

function updateLastBackupTime() {
  const el = document.getElementById("last-backup-time");
  const last = localStorage.getItem("tt_last_backup");
  if (el) el.textContent = last ? new Date(last).toLocaleString("en-AU") : "Never";
}

function backupLog(msg) {
  const el = document.getElementById("backup-log");
  if (!el) return;
  const line = document.createElement("div");
  line.textContent = new Date().toLocaleTimeString("en-AU", { hour:"2-digit", minute:"2-digit" }) + " — " + msg;
  el.insertBefore(line, el.firstChild);
  while (el.children.length > 5) el.removeChild(el.lastChild);
}

// ── Version History ───────────────────────────────────────────
const APP_VERSION = "DV21";
const VERSION_HISTORY = [
  {
    version: "DV21",
    date: "2026-06-26",
    status: "current",
    changes: [
      "Moved to IGT's official GitHub organisation — IGTOpsAPAC",
      "New app URL: https://IGTOpsAPAC.github.io/timetrack",
      "App now owned by IGT organisation — stays with company",
      "Strengthens Azure App Registration approval request to IT",
      "Redirect URI updated to IGTOpsAPAC GitHub Pages domain",
    ]
  },
  {
    version: "DV20",
    date: "2026-06-26",
    changes: [
      "CSV attendance backup — auto-saves to Downloads folder as CSV",
      "Download options: Today only, This week, This month, All records",
      "Configurable auto-save interval and filename prefix",
      "Full JSON backup/restore for data migration between PCs",
    ]
  },
  {
    version: "DV19",
    date: "2026-06-26",
    changes: [
      "Local file backup — auto-saves JSON backup to Downloads folder",
      "Configurable auto-save interval and filename prefix",
      "Manual backup download and restore from backup file",
    ]
  },
  {
    version: "DV18",
    date: "2026-06-26",
    changes: [
      "SharePoint List sync — real-time read/write to SharePoint Lists",
      "Sign in with Microsoft button in Admin → Settings",
      "Auto-sync to SharePoint on every clock in/out",
      "Manual push/pull and connection test buttons",
    ]
  },
  {
    version: "DV17",
    date: "2026-06-25",
    status: "current",
    changes: [
      "Fixed face detection — added camera warm-up delay and video ready check",
      "Improved face matching with live confidence percentage display",
      "Added 30-second timeout for face scanning with helpful fallback message",
      "Fixed face descriptor loading from localStorage (array/object format)",
      "Added version number to home screen and version history page",
    ]
  },
  {
    version: "DV16",
    date: "2026-06-24",
    changes: [
      "Face ID: loosened match threshold to 0.6 for better recognition",
      "Face ID: lowered face detection score threshold to 0.3",
      "Added live match percentage display during face scanning",
      "Fixed ZXing barcode scanner loading issue",
      "IGT logo now displays correctly in header (removed filter)",
    ]
  },
  {
    version: "DV15",
    date: "2026-06-24",
    changes: [
      "Fixed IGT logo display in header — removed CSS filter that made it invisible",
      "Logo now shows full colour in nav bar",
    ]
  },
  {
    version: "DV14",
    date: "2026-06-24",
    changes: [
      "Face API: switched to more reliable CDN with automatic fallback",
      "Face API: added script load wait logic before attempting to use models",
    ]
  },
  {
    version: "DV13",
    date: "2026-06-24",
    changes: [
      "Added Face ID login — enroll face in Admin, scan to clock in/out",
      "Face enrollment with camera capture in employee edit modal",
      "Face ID shows ✅ badge on enrolled employees in admin list",
      "Fixed admin navigation button selector causing access failure",
      "Face ID stops automatically when leaving login screen",
    ]
  },
  {
    version: "DV12",
    date: "2026-06-24",
    changes: [
      "Fixed QR code generation — switched to qrcodejs library",
      "QR codes now generate correctly in Print Barcodes modal",
    ]
  },
  {
    version: "DV11",
    date: "2026-06-24",
    changes: [
      "Fixed admin access broken by duplicate showScreen function",
      "ZXing barcode library loads asynchronously to prevent blocking app",
      "Barcode scanner checks if ZXing loaded before attempting to scan",
    ]
  },
  {
    version: "DV10",
    date: "2026-06-24",
    changes: [
      "Fixed admin access — nav button found by onclick attribute not position",
      "ZXing library now loads safely without blocking app if unavailable",
    ]
  },
  {
    version: "DV9",
    date: "2026-06-24",
    changes: [
      "Added QR code barcode scanning on home screen",
      "Print Barcodes feature in Admin — generates QR card for each employee",
      "Scan beep sound on successful recognition",
      "Back camera preferred for barcode scanning",
      "Print layout optimised for ID card printing",
    ]
  },
  {
    version: "DV8",
    date: "2026-06-23",
    changes: [
      "Integrated Power Automate webhook for email delivery",
      "Scheduled reports now email Excel file as attachment",
      "Fallback to local download if email fails",
      "Added Send Now button on each schedule for immediate testing",
    ]
  },
  {
    version: "DV7",
    date: "2026-06-23",
    changes: [
      "Added email subject field to report schedules",
      "Subject used as Excel filename and report header",
      "Subject displayed on schedule cards in admin panel",
    ]
  },
  {
    version: "DV6",
    date: "2026-06-23",
    changes: [
      "Updated default work areas: Gaming Assembly, Fintech Assembly, Repair Centre, Warehouse, Operations Support",
      "Work areas updated across employee modal, settings, and import",
    ]
  },
  {
    version: "DV5",
    date: "2026-06-23",
    changes: [
      "Home screen shows search only — no employee names visible until typed",
      "Lunch break deducted from total hours (per employee or default setting)",
      "Scheduled timesheet reports — multiple recipients, area/staff type filters",
      "Report schedule: choose days, time, areas, permanent/contractor/all",
    ]
  },
  {
    version: "DV4",
    date: "2026-06-23",
    changes: [
      "Excel mass import of employees from spreadsheet",
      "Import validates missing fields, duplicates, and unparseable shift times",
      "Preview screen shows new vs overwrite before confirming",
      "Auto-generates Employee IDs from names if not provided",
    ]
  },
  {
    version: "DV3",
    date: "2026-06-23",
    changes: [
      "Added popup warning if employee already clocked in",
      "Added popup warning if employee already completed a shift today",
      "Both warnings appear before PIN pad — option to proceed or go back",
    ]
  },
  {
    version: "DV2",
    date: "2026-06-23",
    changes: [
      "Replaced employee tile grid with search-by-name list",
      "Search highlights matching text in results",
      "Employee list shows clock-in status badges",
    ]
  },
  {
    version: "DV1",
    date: "2026-06-22",
    changes: [
      "Initial release — PIN-based employee clock in/out",
      "IGT branding with full colour logo",
      "Admin panel: employee management, settings, report schedules",
      "Daily timesheet report with Excel export",
      "Clock in/out with time variance vs standard hours",
      "Today's attendance table with live status",
    ]
  },
];

function renderLatestVersion() {
  const el = document.getElementById("version-latest");
  if (!el) return;
  const latest = VERSION_HISTORY[0];
  el.innerHTML = `<div style="font-size:12px;color:var(--text2);margin-bottom:4px">${latest.date} — What's new in ${latest.version}:</div>
    <ul style="padding-left:1.25rem;margin:0">
      ${latest.changes.slice(0,3).map(c=>`<li style="font-size:12px;color:var(--text);margin-bottom:2px">${c}</li>`).join("")}
      ${latest.changes.length > 3 ? `<li style="font-size:12px;color:var(--text2)">+${latest.changes.length-3} more…</li>` : ""}
    </ul>`;
}

function showVersionHistory() {
  const list = document.getElementById("version-list");
  list.innerHTML = VERSION_HISTORY.map(v => `
    <div style="margin-bottom:1rem;padding:1rem;border-radius:var(--radius);border:1px solid var(--border);${v.status==="current"?"border-color:var(--igt-blue);background:var(--igt-blue-light)":"background:#fff"}">
      <div style="display:flex;align-items:center;gap:8px;margin-bottom:.5rem">
        <span style="font-weight:700;font-size:15px;color:var(--igt-blue)">${v.version}</span>
        ${v.status==="current"?'<span class="badge badge-blue" style="font-size:11px">● Current</span>':""}
        <span style="font-size:12px;color:var(--text2);margin-left:auto">${v.date}</span>
      </div>
      <ul style="padding-left:1.25rem;margin:0">
        ${v.changes.map(c=>`<li style="font-size:13px;color:var(--text);margin-bottom:3px">${c}</li>`).join("")}
      </ul>
    </div>`).join("");
  document.getElementById("version-modal").classList.add("open");
}

function closeVersionModal() {
  document.getElementById("version-modal").classList.remove("open");
}

// ── Face ID ───────────────────────────────────────────────────
let faceApiLoaded = false;
let faceStream = null;
let faceDetecting = false;
let faceDescriptors = [];
let enrollStream = null;

const FACE_MODELS_URL = "https://cdn.jsdelivr.net/gh/justadudewhohacks/face-api.js@0.22.2/weights";
const FACE_MODELS_URL_FALLBACK = "https://unpkg.com/face-api.js@0.22.2/weights";
const FACE_MATCH_THRESHOLD = 0.6; // 0.6 is more forgiving, lower = stricter

async function loadFaceApi() {
  if (faceApiLoaded) return true;
  // Wait for face-api.js script to load (up to 10s)
  for (let i = 0; i < 20; i++) {
    if (typeof faceapi !== "undefined") break;
    await new Promise(r => setTimeout(r, 500));
  }
  if (typeof faceapi === "undefined") {
    console.error("face-api.js not loaded");
    return false;
  }
  try {
    setFaceStatus("Loading Face ID models…", "scanning");
    // Try primary CDN first, fallback if needed
    try {
      await Promise.all([
        faceapi.nets.tinyFaceDetector.loadFromUri(FACE_MODELS_URL),
        faceapi.nets.faceLandmark68TinyNet.loadFromUri(FACE_MODELS_URL),
        faceapi.nets.faceRecognitionNet.loadFromUri(FACE_MODELS_URL),
      ]);
    } catch(e) {
      await Promise.all([
        faceapi.nets.tinyFaceDetector.loadFromUri(FACE_MODELS_URL_FALLBACK),
        faceapi.nets.faceLandmark68TinyNet.loadFromUri(FACE_MODELS_URL_FALLBACK),
        faceapi.nets.faceRecognitionNet.loadFromUri(FACE_MODELS_URL_FALLBACK),
      ]);
    }
    faceApiLoaded = true;
    return true;
  } catch(e) {
    console.error("Face API model load error:", e);
    return false;
  }
}

function setFaceStatus(msg, type) {
  const el = document.getElementById("face-status");
  if (!el) return;
  el.textContent = msg;
  el.className = `face-status ${type}`;
  el.style.display = msg ? "block" : "none";
}

function buildFaceDescriptors() {
  faceDescriptors = [];
  employees.forEach(emp => {
    if (emp.faceDescriptor) {
      try {
        // Handle both array and object formats from localStorage
        const raw = emp.faceDescriptor;
        let arr;
        if (Array.isArray(raw)) {
          arr = new Float32Array(raw);
        } else if (typeof raw === "object") {
          arr = new Float32Array(Object.values(raw));
        } else {
          return;
        }
        if (arr.length === 128) { // valid face descriptor is 128 floats
          faceDescriptors.push({ empKey: emp.key, descriptor: arr });
          console.log(`Face descriptor loaded for ${emp.name}: ${arr.length} values`);
        }
      } catch(e) {
        console.error(`Failed to load face descriptor for ${emp.name}:`, e);
      }
    }
  });
  console.log(`Total face descriptors loaded: ${faceDescriptors.length}`);
}

async function startFaceId() {
  const btn = document.getElementById("faceid-btn");
  const wrap = document.getElementById("face-scanner-wrap");
  const stopBtn = document.getElementById("stop-face-btn");
  const enrolled = employees.filter(e => e.faceDescriptor);
  if (!enrolled.length) {
    setFaceStatus("No faces enrolled yet. Ask admin to enroll employee faces first.", "error");
    setTimeout(() => setFaceStatus("", ""), 3000);
    return;
  }
  btn.disabled = true;
  setFaceStatus("Loading Face ID models…", "scanning");
  const loaded = await loadFaceApi();
  if (!loaded) {
    setFaceStatus("Face ID unavailable — use barcode or search instead.", "error");
    btn.disabled = false;
    setTimeout(() => setFaceStatus("", ""), 3000);
    return;
  }
  try {
    faceStream = await navigator.mediaDevices.getUserMedia({
      video: { facingMode: "user", width: { ideal: 640 }, height: { ideal: 480 } }
    });
    const video = document.getElementById("face-video");
    video.srcObject = faceStream;

    // Wait for video to be truly ready
    await new Promise((resolve) => {
      video.onloadedmetadata = () => { video.play(); resolve(); };
      setTimeout(resolve, 3000); // fallback
    });
    // Extra delay for camera to warm up
    await new Promise(r => setTimeout(r, 1500));

    wrap.style.display = "block";
    stopBtn.style.display = "block";
    setFaceStatus("🔍 Looking for your face…", "scanning");
    faceDetecting = true;
    buildFaceDescriptors();
    console.log("Face descriptors built:", faceDescriptors.length);

    // Auto-stop after 30 seconds
    setTimeout(() => {
      if (faceDetecting) {
        faceDetecting = false;
        setFaceStatus("⏱ Not recognised — try re-enrolling in better lighting or use PIN/search", "error");
        stopFaceId();
      }
    }, 30000);

    detectFaceLoop(video);
  } catch(e) {
    setFaceStatus("Camera unavailable: " + e.message, "error");
    btn.disabled = false;
    setTimeout(() => setFaceStatus("", ""), 3000);
  }
}

async function detectFaceLoop(video) {
  if (!faceDetecting) return;
  try {
    // Ensure video is playing and has dimensions
    if (video.videoWidth === 0 || video.paused) {
      setFaceStatus("⏳ Waiting for camera…", "scanning");
      if (faceDetecting) setTimeout(() => detectFaceLoop(video), 500);
      return;
    }

    const options = new faceapi.TinyFaceDetectorOptions({ inputSize: 320, scoreThreshold: 0.3 });
    const detection = await faceapi
      .detectSingleFace(video, options)
      .withFaceLandmarks(true)
      .withFaceDescriptor();

    if (detection) {
      if (faceDescriptors.length) {
        let bestMatch = null, bestDist = Infinity;
        faceDescriptors.forEach(fd => {
          const dist = faceapi.euclideanDistance(detection.descriptor, fd.descriptor);
          console.log(`Distance to ${employees.find(e=>e.key===fd.empKey)?.name}: ${dist.toFixed(3)}`);
          if (dist < bestDist) { bestDist = dist; bestMatch = fd; }
        });

        const pct = Math.round((1 - Math.min(bestDist, 1)) * 100);
        const matchEmp = employees.find(e => e.key === bestMatch?.empKey);
        setFaceStatus(`🔍 Matching… ${pct}% confidence — ${matchEmp?.name}`, "scanning");

        if (bestDist < FACE_MATCH_THRESHOLD && bestMatch) {
          setFaceStatus(`✅ Recognised: ${matchEmp?.name}!`, "success");
          faceDetecting = false;
          playBeep();
          setTimeout(() => {
            stopFaceId();
            selectEmployee(bestMatch.empKey);
          }, 1000);
          return;
        }
      } else {
        setFaceStatus("⚠ Face detected but no enrolled faces found", "error");
      }
    } else {
      setFaceStatus("🔍 No face detected — look straight at the camera", "scanning");
    }
  } catch(e) {
    console.error("Face detection error:", e);
    setFaceStatus("⚠ Error: " + e.message, "error");
  }
  if (faceDetecting) setTimeout(() => detectFaceLoop(video), 500);
}

function stopFaceId() {
  faceDetecting = false;
  if (faceStream) { faceStream.getTracks().forEach(t => t.stop()); faceStream = null; }
  const wrap = document.getElementById("face-scanner-wrap");
  const stopBtn = document.getElementById("stop-face-btn");
  const btn = document.getElementById("faceid-btn");
  if (wrap) wrap.style.display = "none";
  if (stopBtn) stopBtn.style.display = "none";
  if (btn) btn.disabled = false;
  const video = document.getElementById("face-video");
  if (video) video.srcObject = null;
  setTimeout(() => setFaceStatus("", ""), 2000);
}

// ── Face Enroll ───────────────────────────────────────────────
async function startFaceEnroll() {
  const wrap = document.getElementById("enroll-video-wrap");
  const msg = document.getElementById("enroll-msg");
  const btn = document.getElementById("enroll-start-btn");
  msg.textContent = "Loading Face ID models…";
  btn.disabled = true;
  const loaded = await loadFaceApi();
  if (!loaded) {
    msg.textContent = "Face API failed to load. Check your internet connection.";
    btn.disabled = false;
    return;
  }
  try {
    enrollStream = await navigator.mediaDevices.getUserMedia({ video: { facingMode:"user", width:320, height:240 } });
    const video = document.getElementById("enroll-video");
    video.srcObject = enrollStream;
    wrap.style.display = "block";
    msg.textContent = "Position face in camera then click Capture.";
    let captureBtn = document.getElementById("enroll-capture-btn");
    if (!captureBtn) {
      captureBtn = document.createElement("button");
      captureBtn.id = "enroll-capture-btn";
      captureBtn.type = "button";
      captureBtn.className = "btn btn-primary";
      captureBtn.style.marginTop = "8px";
      captureBtn.textContent = "📸 Capture Face";
      captureBtn.onclick = captureFaceEnroll;
      wrap.after(captureBtn);
    }
    captureBtn.style.display = "";
    btn.disabled = false;
  } catch(e) {
    msg.textContent = "Camera unavailable: " + e.message;
    btn.disabled = false;
  }
}

async function captureFaceEnroll() {
  const video = document.getElementById("enroll-video");
  const msg = document.getElementById("enroll-msg");
  const captureBtn = document.getElementById("enroll-capture-btn");
  msg.textContent = "Detecting face…";
  try {
    const detection = await faceapi
      .detectSingleFace(video, new faceapi.TinyFaceDetectorOptions({ scoreThreshold:0.5 }))
      .withFaceLandmarks(true)
      .withFaceDescriptor();
    if (!detection) {
      msg.textContent = "No face detected. Make sure your face is clearly visible and try again.";
      return;
    }
    window._pendingFaceDescriptor = Array.from(detection.descriptor);
    window._clearFaceDescriptor = false;
    msg.textContent = "✅ Face captured! Click Save Employee to confirm.";
    document.getElementById("enroll-status-badge").innerHTML = '<span class="face-enrolled-badge">✅ Face captured — save to confirm</span>';
    document.getElementById("enroll-clear-btn").style.display = "";
    stopFaceEnroll();
    if (captureBtn) captureBtn.style.display = "none";
  } catch(e) {
    msg.textContent = "Error capturing face: " + e.message;
  }
}

function stopFaceEnroll() {
  if (enrollStream) { enrollStream.getTracks().forEach(t => t.stop()); enrollStream = null; }
  const wrap = document.getElementById("enroll-video-wrap");
  if (wrap) wrap.style.display = "none";
}

function clearFaceEnroll() {
  window._pendingFaceDescriptor = null;
  window._clearFaceDescriptor = true;
  document.getElementById("enroll-status-badge").innerHTML = '<span style="font-size:12px;color:var(--text2)">Face removed — save to confirm</span>';
  document.getElementById("enroll-msg").textContent = "";
  document.getElementById("enroll-clear-btn").style.display = "none";
  stopFaceEnroll();
}

// ── SharePoint List Sync Engine ───────────────────────────────
const SP_SCOPES = ["Sites.ReadWrite.All", "Files.ReadWrite", "User.Read"];
const SP_DEFAULT_CLIENT_ID = "d3590ed6-52b3-4102-aeff-aad2292ab01c";
const APP_GITHUB_URL = "https://igtopsapac.github.io/timetrack";
let spMsal = null;
let spToken = null;
let spSiteId = null;
let spSyncing = false;

function getSiteUrl() {
  return settings.siteUrl || "https://igtplc.sharepoint.com/sites/APACManufacturingOperationsTeam";
}

function getClientId() {
  return settings.spClientId || SP_DEFAULT_CLIENT_ID;
}

async function initMsal() {
  if (spMsal) return spMsal;
  spMsal = new msal.PublicClientApplication({
    auth: {
      clientId: getClientId(),
      authority: "https://login.microsoftonline.com/common",
      redirectUri: window.location.origin + window.location.pathname,
      navigateToLoginRequestUrl: false,
    },
    cache: { cacheLocation: "localStorage", storeAuthStateInCookie: true },
  });
  await spMsal.initialize();
  // Handle redirect
  try {
    const resp = await spMsal.handleRedirectPromise();
    if (resp?.accessToken) {
      spToken = resp.accessToken;
      updateSpAuthStatus(true, resp.account?.name);
    }
  } catch(e) { console.error("MSAL redirect error:", e); }
  return spMsal;
}

async function spSignIn() {
  const msal = await initMsal();
  const accounts = msal.getAllAccounts();
  if (accounts.length > 0) {
    try {
      const resp = await msal.acquireTokenSilent({ scopes: SP_SCOPES, account: accounts[0] });
      spToken = resp.accessToken;
      updateSpAuthStatus(true, accounts[0].name);
      spLog("✓ Already signed in as " + accounts[0].name);
      return true;
    } catch(e) {}
  }
  try {
    await msal.loginRedirect({ scopes: SP_SCOPES });
  } catch(e) {
    spLog("✗ Sign-in failed: " + e.message);
    return false;
  }
}

async function getSpToken() {
  const msalInst = await initMsal();
  const accounts = msalInst.getAllAccounts();
  if (!accounts.length) return null;
  try {
    const resp = await msalInst.acquireTokenSilent({ scopes: SP_SCOPES, account: accounts[0] });
    spToken = resp.accessToken;
    return spToken;
  } catch(e) {
    try {
      await msalInst.acquireTokenRedirect({ scopes: SP_SCOPES, account: accounts[0] });
    } catch(e2) {}
    return null;
  }
}

function spSignOut() {
  if (spMsal) spMsal.logoutRedirect();
}

function updateSpAuthStatus(signedIn, name) {
  const el = document.getElementById("sp-auth-status");
  const btn = document.getElementById("sp-signout-btn");
  if (!el) return;
  if (signedIn) {
    el.innerHTML = `<span class="badge badge-green">✓ Signed in${name ? " as " + name.split(" ")[0] : ""}</span>`;
    if (btn) btn.style.display = "";
  } else {
    el.innerHTML = `<span class="badge badge-gray">Not signed in</span>`;
    if (btn) btn.style.display = "none";
  }
}

async function getSiteId() {
  if (spSiteId) return spSiteId;
  const token = await getSpToken();
  if (!token) return null;
  const url = getSiteUrl();
  // Extract hostname and path from site URL
  const match = url.match(/https:\/\/([^/]+)(\/sites\/[^/]+)/);
  if (!match) { spLog("✗ Invalid site URL"); return null; }
  const hostname = match[1];
  const sitePath = match[2];
  const resp = await fetch(`https://graph.microsoft.com/v1.0/sites/${hostname}:${sitePath}`, {
    headers: { Authorization: `Bearer ${token}` }
  });
  if (!resp.ok) { spLog("✗ Could not find site: " + resp.status); return null; }
  const data = await resp.json();
  spSiteId = data.id;
  return spSiteId;
}

async function spRequest(method, path, body) {
  const token = await getSpToken();
  if (!token) throw new Error("Not signed in to SharePoint");
  const opts = {
    method,
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
  };
  if (body) opts.body = JSON.stringify(body);
  const resp = await fetch(`https://graph.microsoft.com/v1.0${path}`, opts);
  if (!resp.ok) {
    const err = await resp.text();
    throw new Error(`Graph API ${resp.status}: ${err}`);
  }
  if (resp.status === 204) return null;
  return resp.json();
}

async function getListItems(siteId, listName) {
  let items = [], url = `/sites/${siteId}/lists/${listName}/items?expand=fields&$top=500`;
  while (url) {
    const data = await spRequest("GET", url);
    items = items.concat(data.value || []);
    url = data["@odata.nextLink"] ? data["@odata.nextLink"].replace("https://graph.microsoft.com/v1.0", "") : null;
  }
  return items;
}

async function addListItem(siteId, listName, fields) {
  return spRequest("POST", `/sites/${siteId}/lists/${listName}/items`, { fields });
}

async function updateListItem(siteId, listName, itemId, fields) {
  return spRequest("PATCH", `/sites/${siteId}/lists/${listName}/items/${itemId}/fields`, fields);
}

async function deleteListItem(siteId, listName, itemId) {
  return spRequest("DELETE", `/sites/${siteId}/lists/${listName}/items/${itemId}`);
}

// ── Test Connection ───────────────────────────────────────────
async function spTestConnection() {
  spLog("Testing connection…");
  try {
    const siteId = await getSiteId();
    if (!siteId) { spLog("✗ Could not get site ID — check Site URL and sign in"); return; }
    // Test each list
    for (const listName of ["TimeTrack_Employees","TimeTrack_Attendance","TimeTrack_Settings"]) {
      const items = await getListItems(siteId, listName);
      spLog(`✓ ${listName} — ${items.length} items`);
    }
    spLog("✓ All lists accessible! Connection successful.");
    toast("✓ SharePoint connection successful!", "success");
  } catch(e) {
    spLog("✗ Connection failed: " + e.message);
    toast("Connection failed: " + e.message, "error");
  }
}

// ── Push all data to SharePoint ───────────────────────────────
async function spPushAll() {
  if (spSyncing) { toast("Sync already in progress", "error"); return; }
  spSyncing = true;
  spLog("Pushing to SharePoint…");
  setSyncStatus("Syncing…");
  try {
    const siteId = await getSiteId();
    if (!siteId) throw new Error("Not connected — sign in first");

    // ── Push Employees ──
    spLog("Syncing employees…");
    const spEmps = await getListItems(siteId, "TimeTrack_Employees");
    const spEmpMap = {};
    spEmps.forEach(i => { if (i.fields.EmpKey) spEmpMap[i.fields.EmpKey] = i.id; });

    for (const e of employees) {
      const fields = {
        Title: e.name,
        EmpKey: e.key,
        EmployeeID: e.empId || "",
        Area: e.area || "",
        StartTime: e.startTime || "",
        EndTime: e.endTime || "",
        HoursPerDay: e.hours || 8,
        LunchMins: e.lunchMins || 0,
        EmpStatus: e.status || "Permanent",
        PIN: e.pin || "",
        FaceData: e.faceDescriptor ? JSON.stringify(e.faceDescriptor) : "",
      };
      if (spEmpMap[e.key]) {
        await updateListItem(siteId, "TimeTrack_Employees", spEmpMap[e.key], fields);
      } else {
        await addListItem(siteId, "TimeTrack_Employees", fields);
      }
    }
    spLog(`✓ ${employees.length} employees synced`);

    // ── Push Attendance ──
    spLog("Syncing attendance records…");
    const spAtt = await getListItems(siteId, "TimeTrack_Attendance");
    const spAttMap = {};
    spAtt.forEach(i => { if (i.fields.EmpKey && i.fields.AttendanceDate) spAttMap[`${i.fields.EmpKey}_${i.fields.AttendanceDate}`] = i.id; });

    for (const e of clockEntries) {
      const key = `${e.empKey}_${e.date}`;
      const fields = {
        Title: `${e.name} — ${e.date}`,
        EmpKey: e.empKey || "",
        EmployeeID: e.empId || "",
        Area: e.area || "",
        AttendanceDate: e.date || "",
        TimeIn: e.timeIn || "",
        TimeOut: e.timeOut || "",
        StdStart: e.stdStart || "",
        StdEnd: e.stdEnd || "",
        StdHours: e.stdHours || 8,
        LunchMins: e.lunchMins || 0,
        EmpStatus: e.status || "",
      };
      if (spAttMap[key]) {
        await updateListItem(siteId, "TimeTrack_Attendance", spAttMap[key], fields);
      } else {
        await addListItem(siteId, "TimeTrack_Attendance", fields);
      }
    }
    spLog(`✓ ${clockEntries.length} attendance records synced`);

    // ── Push Settings ──
    spLog("Syncing settings…");
    const spSets = await getListItems(siteId, "TimeTrack_Settings");
    const spSetMap = {};
    spSets.forEach(i => { if (i.fields.SettingKey) spSetMap[i.fields.SettingKey] = i.id; });
    for (const [k, v] of Object.entries(settings)) {
      if (typeof v === "string" || typeof v === "number") {
        const fields = { Title: k, SettingKey: k, SettingValue: String(v) };
        if (spSetMap[k]) await updateListItem(siteId, "TimeTrack_Settings", spSetMap[k], fields);
        else await addListItem(siteId, "TimeTrack_Settings", fields);
      }
    }
    spLog(`✓ Settings synced`);

    const now = new Date().toLocaleTimeString("en-AU", { hour:"2-digit", minute:"2-digit" });
    setSyncStatus("Synced " + now);
    spLog(`✓ Push complete at ${now}`);
    toast("✓ Pushed to SharePoint successfully!", "success");
  } catch(e) {
    spLog("✗ Push failed: " + e.message);
    toast("Push failed: " + e.message, "error");
    setSyncStatus("Sync failed");
  }
  spSyncing = false;
}

// ── Pull all data from SharePoint ─────────────────────────────
async function spPullAll(silent = false) {
  if (spSyncing) return;
  const token = await getSpToken();
  if (!token) {
    if (!silent) spLog("Not signed in — click 'Sign in to SharePoint' first");
    return;
  }
  spSyncing = true;
  if (!silent) spLog("Pulling from SharePoint…");
  setSyncStatus("Syncing…");
  try {
    const siteId = await getSiteId();
    if (!siteId) throw new Error("Could not connect to site");

    // ── Pull Employees ──
    const spEmps = await getListItems(siteId, "TimeTrack_Employees");
    if (spEmps.length) {
      employees = spEmps.map(i => ({
        key: i.fields.EmpKey || "e" + i.id,
        name: i.fields.Title || "",
        empId: i.fields.EmployeeID || "",
        area: i.fields.Area || "",
        startTime: i.fields.StartTime || "09:00",
        endTime: i.fields.EndTime || "17:00",
        hours: i.fields.HoursPerDay || 8,
        lunchMins: i.fields.LunchMins || 0,
        status: i.fields.EmpStatus || "Permanent",
        pin: i.fields.PIN || "0000",
        faceDescriptor: i.fields.FaceData ? JSON.parse(i.fields.FaceData) : null,
      })).filter(e => e.name);
      if (!silent) spLog(`✓ ${employees.length} employees loaded`);
    }

    // ── Pull Attendance ──
    const spAtt = await getListItems(siteId, "TimeTrack_Attendance");
    if (spAtt.length) {
      const spEntries = spAtt.map(i => ({
        empKey: i.fields.EmpKey || "",
        empId: i.fields.EmployeeID || "",
        name: i.fields.Title?.replace(/ — .*/, "") || "",
        area: i.fields.Area || "",
        date: i.fields.AttendanceDate || "",
        timeIn: i.fields.TimeIn || null,
        timeOut: i.fields.TimeOut || null,
        stdStart: i.fields.StdStart || "",
        stdEnd: i.fields.StdEnd || "",
        stdHours: i.fields.StdHours || 8,
        lunchMins: i.fields.LunchMins || 0,
        status: i.fields.EmpStatus || "",
      })).filter(e => e.date);
      // Merge — keep local entries not yet on SharePoint
      const spKeys = new Set(spEntries.map(e => `${e.empKey}_${e.date}`));
      const localOnly = clockEntries.filter(e => !spKeys.has(`${e.empKey}_${e.date}`));
      clockEntries = [...spEntries, ...localOnly];
      if (!silent) spLog(`✓ ${clockEntries.length} attendance records loaded`);
    }

    // ── Pull Settings ──
    const spSets = await getListItems(siteId, "TimeTrack_Settings");
    if (spSets.length) {
      const pulled = {};
      spSets.forEach(i => { if (i.fields.SettingKey) pulled[i.fields.SettingKey] = i.fields.SettingValue; });
      settings = { ...settings, ...pulled };
    }

    saveLocal();
    renderAll();
    const now = new Date().toLocaleTimeString("en-AU", { hour:"2-digit", minute:"2-digit" });
    setSyncStatus("Synced " + now);
    if (!silent) { spLog(`✓ Pull complete at ${now}`); toast("✓ Pulled from SharePoint!", "success"); }
  } catch(e) {
    if (!silent) { spLog("✗ Pull failed: " + e.message); toast("Pull failed: " + e.message, "error"); }
    setSyncStatus("Sync failed");
  }
  spSyncing = false;
}

// ── Auto-sync on clock in/out ─────────────────────────────────
async function spAutoSync() {
  const token = await getSpToken();
  if (token) spPushAll();
}

function spLog(msg) {
  const el = document.getElementById("sp-sync-log");
  if (el) {
    const line = document.createElement("div");
    line.textContent = new Date().toLocaleTimeString("en-AU", {hour:"2-digit",minute:"2-digit"}) + " — " + msg;
    el.insertBefore(line, el.firstChild);
    // Keep last 10 lines
    while (el.children.length > 10) el.removeChild(el.lastChild);
  }
  console.log("[SP Sync]", msg);
}

function setSyncStatus(msg) {
  const el = document.getElementById("sync-status");
  if (el) el.textContent = msg;
}

// Init MSAL on load and check existing session
async function initSpSync() {
  try {
    const msalInst = await initMsal();
    const accounts = msalInst.getAllAccounts();
    if (accounts.length > 0) {
      updateSpAuthStatus(true, accounts[0].name);
      // Auto-pull on load
      await spPullAll(true);
    } else {
      updateSpAuthStatus(false);
    }
  } catch(e) {
    console.warn("SP init error:", e);
  }
}

// ── Boot ──────────────────────────────────────────────────────

async function runScheduledExport(s) {
  const dateVal = today();
  let entries = clockEntries.filter(e => e.date === dateVal);
  if (s.areas && s.areas.length) entries = entries.filter(e => s.areas.includes(e.area));
  if (s.empType === "permanent")  entries = entries.filter(e => (e.status||"").toLowerCase() !== "contractor");
  if (s.empType === "contractor") entries = entries.filter(e => (e.status||"").toLowerCase() === "contractor");

  if (!entries.length) {
    toast(`No data to send for ${s.name}`, "");
    return;
  }

  const reportTitle = s.subject || `${settings.company||"IGT"} Daily Timesheet Report`;
  const safeName    = (s.subject || s.name).replace(/[^a-zA-Z0-9 _-]/g,"").replace(/\s+/g,"_");
  const filename    = `${safeName}_${dateVal}.xlsx`;

  // Build Excel workbook
  const wsData = [
    [reportTitle,"","","","","","","","","","","",""],
    [`Date: ${dateVal}`,"","Report for:",`${s.name} <${s.email}>`,"","Areas:",s.areas?.length?s.areas.join(", "):"All","Staff:",s.empType==="permanent"?"Permanent only":s.empType==="contractor"?"Contractors only":"All","","",""],
    [],
    ["Employee","Employee ID","Area","Status","Std Start","Actual In","Start Var","Std End","Actual Out","End Var","Std Hrs","Net Hrs","Diff","Result"],
    ...entries.map(e => {
      const actual = calcHours(e.timeIn,e.timeOut,e.lunchMins);
      const diff   = actual!==null?+(actual-e.stdHours).toFixed(2):null;
      const inVar  = timeDiffStr(e.stdStart,e.timeIn);
      const outVar = e.timeOut?timeDiffStr(e.stdEnd,e.timeOut):null;
      const result = e.timeOut?(diff!==null&&diff>=0?"On time":"Short"):e.timeIn?"In progress":"Absent";
      return [e.name,e.empId,e.area,e.status||"",e.stdStart,e.timeIn||"",inVar||"",e.stdEnd,e.timeOut||"",outVar||"",e.stdHours,actual!==null?+actual.toFixed(2):"",diff!==null?diff:"",result];
    }),
    [],
    ["","","","","","","","","","",`=SUM(K5:K${entries.length+4})`,`=SUM(L5:L${entries.length+4})`,`=SUM(M5:M${entries.length+4})`,""],
    ["","","","","","","","","","","Total Std","Total Net","Total Diff",""],
  ];
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(wsData);
  ws["!cols"] = [{wch:22},{wch:12},{wch:16},{wch:12},{wch:11},{wch:11},{wch:11},{wch:11},{wch:11},{wch:11},{wch:10},{wch:10},{wch:10},{wch:12}];
  XLSX.utils.book_append_sheet(wb, ws, "Timesheet");

  // Convert workbook to base64
  const wbOut    = XLSX.write(wb, { bookType:"xlsx", type:"base64" });
  const areasTxt = s.areas?.length ? s.areas.join(", ") : "All areas";
  const staffTxt = s.empType==="permanent"?"Permanent only":s.empType==="contractor"?"Contractors only":"All staff";

  // Show sending indicator
  toast(`📤 Sending report to ${s.name}…`);

  try {
    const resp = await fetch(POWER_AUTOMATE_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        subject:        s.subject || reportTitle,
        recipientName:  s.name,
        recipientEmail: s.email,
        date:           dateVal,
        fileBase64:     wbOut,
        fileName:       filename,
        areas:          areasTxt,
        staffType:      staffTxt,
      }),
    });

    if (resp.ok || resp.status === 202) {
      toast(`✅ Report emailed to ${s.name}`, "success");
    } else {
      const errText = await resp.text();
      console.error("Power Automate error:", resp.status, errText);
      toast(`⚠ Email failed for ${s.name} (${resp.status}) — downloading instead`, "error");
      XLSX.writeFile(wb, filename); // fallback download
    }
  } catch(err) {
    console.error("Fetch error:", err);
    toast(`⚠ Could not reach email server — downloading instead`, "error");
    XLSX.writeFile(wb, filename); // fallback download
  }
}

function testSendNow(idx) {
  const s = getSchedules()[idx];
  if (!s) return;
  if (confirm(`Send report now to ${s.name} <${s.email}>?`)) {
    runScheduledExport(s);
  }
}

setInterval(checkSchedules, 60000);

init();
