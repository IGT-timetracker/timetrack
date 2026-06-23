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
    settings = { adminPin:"0000", areas:"Production,Warehouse,Office,Maintenance", company:"IGT APAC Manufacturing", recipientName:"Operations Manager", recipientEmail:"manager@igt.com", siteName:"APACManufacturingOperationsTeam", filePath:"General/ATTENDANCE/Attendance.xlsx" };
    saveLocal();
  }
}

function saveLocal() {
  localStorage.setItem("tt_employees", JSON.stringify(employees));
  localStorage.setItem("tt_entries", JSON.stringify(clockEntries));
  localStorage.setItem("tt_settings", JSON.stringify(settings));
}

function showScreen(id) {
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

const AVATAR_COLORS = [["#e6eef9","#0047BB"],["#e6f4ed","#1a7a4a"],["#fff0e8","#c0390b"],["#fff3e0","#e65100"],["#f3e8ff","#6b21a8"],["#e0f2fe","#0369a1"]];
function initials(n) { return (n||"?").split(" ").map(x=>x[0]).join("").toUpperCase().slice(0,2); }
function avatarStyle(i) { const c=AVATAR_COLORS[i%AVATAR_COLORS.length]; return `background:${c[0]};color:${c[1]}`; }

function renderEmpGrid(filter = "") {
  const g = document.getElementById("emp-grid");
  if (!g) return;
  const q = filter.toLowerCase().trim();
  const filtered = employees.filter(e =>
    !q || e.name.toLowerCase().includes(q) || e.empId.toLowerCase().includes(q) || e.area.toLowerCase().includes(q)
  );
  if (!employees.length) {
    g.innerHTML = '<div class="emp-empty">No employees yet.<br>Use Admin access to add employees.</div>';
    return;
  }
  if (!filtered.length) {
    g.innerHTML = '<div class="emp-empty">No employees match your search.</div>';
    return;
  }
  g.innerHTML = filtered.map((e) => {
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
    document.querySelectorAll(".admin-only").forEach(el=>el.style.display="");
    document.getElementById("admin-signout-btn").style.display="";
    showScreen("screen-app");
    showSection("admin", document.querySelectorAll(".nav-btn")[2]);
    toast("Admin access granted","success");
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
    clockEntries.push({ empKey, date:today(), timeIn:timeStamp, timeOut:null, name:emp.name, area:emp.area, empId:emp.empId, stdStart:emp.startTime, stdEnd:emp.endTime, stdHours:emp.hours });
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
  setTimeout(()=>showScreen("screen-login"), 4000);
}

function forceClockIn(empKey) {
  const emp=employees.find(e=>e.key===empKey);
  const t=new Date().toLocaleTimeString("en-AU",{hour:"2-digit",minute:"2-digit",hour12:false});
  clockEntries.push({empKey,date:today(),timeIn:t,timeOut:null,name:emp.name,area:emp.area,empId:emp.empId,stdStart:emp.startTime,stdEnd:emp.endTime,stdHours:emp.hours});
  saveLocal(); performClockAction(empKey);
}

function renderAll() { renderTodayTable(); renderActiveBanner(); renderEmpList(); loadSettingsForm(); renderReportRecipient(); genReport(); }

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
  wrap.innerHTML=`<div style="overflow-x:auto"><table><thead><tr><th>Employee</th><th>Area</th><th>Clock in</th><th>Clock out</th><th>Hours</th><th>Status</th></tr></thead><tbody>${entries.map(e=>{
    const hrs=calcHours(e.timeIn,e.timeOut);
    return `<tr><td><div class="emp-row"><div class="emp-avatar" style="${avatarStyle(employees.findIndex(x=>x.key===e.empKey))};width:30px;height:30px;font-size:11px">${initials(e.name)}</div><div><div style="font-weight:600">${e.name}</div><div style="font-size:11px;color:var(--text2)">${e.empId}</div></div></div></td><td><span class="tag">${e.area}</span></td><td><strong>${e.timeIn||"—"}</strong></td><td><strong>${e.timeOut||"—"}</strong></td><td>${hrs!==null?hrs.toFixed(1)+"h":"—"}</td><td>${e.timeOut?'<span class="badge badge-green">✓ Done</span>':'<span class="badge badge-amber">● Active</span>'}</td></tr>`;
  }).join("")}</tbody></table></div>`;
}

function calcHours(tin,tout) {
  if(!tin||!tout) return null;
  const [h1,m1]=tin.split(":").map(Number),[h2,m2]=tout.split(":").map(Number);
  return ((h2*60+m2)-(h1*60+m1))/60;
}

function timeDiffStr(t1,t2) {
  if(!t1||!t2) return null;
  const [h1,m1]=t1.split(":").map(Number),[h2,m2]=t2.split(":").map(Number);
  const diff=(h2*60+m2)-(h1*60+m1),sign=diff<0?"-":"+",abs=Math.abs(diff);
  return `${sign}${Math.floor(abs/60)}h ${abs%60}m`;
}

function renderReportRecipient() {
  const el=document.getElementById("report-recipient-info"); if(!el) return;
  el.innerHTML=`📧 Report recipient: <strong>${settings.recipientName||"—"}</strong> &lt;${settings.recipientEmail||"—"}&gt;`;
}

function genReport() {
  const dateVal=document.getElementById("report-date")?.value||today();
  const entries=clockEntries.filter(e=>e.date===dateVal);
  const wrap=document.getElementById("report-content"); if(!wrap) return;
  if(!entries.length){wrap.innerHTML=`<div class="card"><div class="empty">📅<br><br>No records for ${dateVal}</div></div>`;return;}
  const rows=entries.map(e=>{
    const actual=calcHours(e.timeIn,e.timeOut),diff=actual!==null?actual-e.stdHours:null;
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
      const actual=calcHours(e.timeIn,e.timeOut),diff=actual!==null?+(actual-e.stdHours).toFixed(2):null;
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
  const allData=[["Employee","Employee ID","Area","Date","Time In","Time Out","Std Hours","Actual Hours","Difference"],...clockEntries.map(e=>{const a=calcHours(e.timeIn,e.timeOut);return[e.name,e.empId,e.area,e.date,e.timeIn||"",e.timeOut||"",e.stdHours,a!==null?+a.toFixed(2):"",a!==null?+(a-e.stdHours).toFixed(2):""];})];
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
    <div style="font-size:12px;color:var(--text2)">${e.area} · ${e.startTime}–${e.endTime} · ${e.hours}h/day · PIN: ${"●".repeat(e.pin?.length||4)}</div></div>
    <div style="display:flex;gap:6px">
      <button class="btn" onclick="openEmpModal('${e.key}')" style="padding:6px 10px">✏</button>
      <button class="btn btn-danger" onclick="deleteEmp('${e.key}')" style="padding:6px 10px">🗑</button>
    </div>
  </div></div>`).join("");
}

function openEmpModal(key) {
  editingEmpKey=key||null;
  const areas=(settings.areas||"Production,Warehouse,Office").split(",").map(a=>a.trim());
  document.getElementById("emp-area").innerHTML=areas.map(a=>`<option>${a}</option>`).join("");
  if(key){const e=employees.find(x=>x.key===key);document.getElementById("modal-title").textContent="Edit Employee";document.getElementById("emp-name").value=e.name;document.getElementById("emp-id-field").value=e.empId;document.getElementById("emp-area").value=e.area;document.getElementById("emp-start").value=e.startTime;document.getElementById("emp-end").value=e.endTime;document.getElementById("emp-hours").value=e.hours;document.getElementById("emp-pin").value=e.pin||"";}
  else{document.getElementById("modal-title").textContent="Add Employee";["emp-name","emp-id-field","emp-pin"].forEach(id=>document.getElementById(id).value="");document.getElementById("emp-start").value="09:00";document.getElementById("emp-end").value="17:00";document.getElementById("emp-hours").value="8";}
  document.getElementById("emp-modal").classList.add("open");
}

function closeEmpModal() { document.getElementById("emp-modal").classList.remove("open"); }

function saveEmployee() {
  const name=document.getElementById("emp-name").value.trim(),empId=document.getElementById("emp-id-field").value.trim();
  const area=document.getElementById("emp-area").value,startTime=document.getElementById("emp-start").value;
  const endTime=document.getElementById("emp-end").value,hours=parseFloat(document.getElementById("emp-hours").value);
  const pin=document.getElementById("emp-pin").value.trim();
  if(!name||!empId){toast("Name and ID required","error");return;}
  if(!/^\d{4}$/.test(pin)){toast("PIN must be exactly 4 digits","error");return;}
  if(editingEmpKey){const idx=employees.findIndex(e=>e.key===editingEmpKey);employees[idx]={...employees[idx],name,empId,area,startTime,endTime,hours,pin};}
  else employees.push({key:"e"+Date.now(),name,empId,area,startTime,endTime,hours,pin});
  saveLocal();closeEmpModal();renderEmpList();renderEmpGrid();
  toast(editingEmpKey?"Employee updated":"Employee added","success");
}

function deleteEmp(key) {
  if(!confirm("Remove this employee?")) return;
  employees=employees.filter(e=>e.key!==key);
  saveLocal();renderEmpList();renderEmpGrid();toast("Employee removed");
}

function loadSettingsForm() {
  document.getElementById("cfg-admin-pin").value=settings.adminPin||"0000";
  document.getElementById("cfg-site").value=settings.siteName||"APACManufacturingOperationsTeam";
  document.getElementById("cfg-path").value=settings.filePath||"General/ATTENDANCE/Attendance.xlsx";
  document.getElementById("cfg-areas").value=settings.areas||"";
  document.getElementById("cfg-company").value=settings.company||"";
  document.getElementById("cfg-recipient-name").value=settings.recipientName||"";
  document.getElementById("cfg-recipient-email").value=settings.recipientEmail||"";
}

function saveSettings() {
  const adminPin=document.getElementById("cfg-admin-pin").value.trim();
  if(!/^\d{4}$/.test(adminPin)){toast("Admin PIN must be 4 digits","error");return;}
  settings={...settings,adminPin,siteName:document.getElementById("cfg-site").value.trim(),filePath:document.getElementById("cfg-path").value.trim(),areas:document.getElementById("cfg-areas").value,company:document.getElementById("cfg-company").value.trim(),recipientName:document.getElementById("cfg-recipient-name").value.trim(),recipientEmail:document.getElementById("cfg-recipient-email").value.trim()};
  saveLocal();toast("Settings saved","success");renderReportRecipient();
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

init();
