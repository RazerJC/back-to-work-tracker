/* ── Employee Attendance Analytics — Frontend JS ────────────────────── */

const PAGE_SIZE = 25;
let currentPage = 1;
let allRecords = [];
let debounceTimer;
let deptChartInstance, monthChartInstance, top10ChartInstance;

// ── Init ─────────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  document.getElementById('fileInput').addEventListener('change', handleFileUpload);
  const search = document.getElementById('searchInput');
  search.addEventListener('input', () => { clearTimeout(debounceTimer); debounceTimer = setTimeout(refresh, 350); });
  ['filterCompany','filterMonth','filterYear','filterDept'].forEach(id => {
    document.getElementById(id).addEventListener('change', refresh);
  });
  document.addEventListener('click', e => {
    const dd = document.getElementById('dropdownMenu');
    if (!document.getElementById('exportDropdown').contains(e.target)) dd.classList.add('hidden');
  });
});

// ── Scroll helper ────────────────────────────────────────────────────
function scrollTo(sel) { document.querySelector(sel)?.scrollIntoView({behavior:'smooth',block:'start'}); }

// ── Loading / Notification ───────────────────────────────────────────
function showLoading(msg) {
  document.getElementById('loadingText').textContent = msg || 'Processing data…';
  document.getElementById('loadingOverlay').classList.remove('hidden');
}
function hideLoading() { document.getElementById('loadingOverlay').classList.add('hidden'); }

function notify(text, type='success') {
  const el = document.getElementById('notification');
  const notifText = document.getElementById('notifText');
  notifText.textContent = text;
  el.className = 'fixed top-4 right-4 z-40 px-5 py-3 rounded-xl text-sm font-medium shadow-2xl transition-all duration-500 translate-x-0';
  if (type === 'error') el.classList.add('bg-rose-500/90', 'text-white');
  else el.classList.add('bg-emerald-500/90', 'text-white');
  setTimeout(() => el.classList.add('translate-x-[120%]'), 4000);
}

// ── Data Loading ─────────────────────────────────────────────────────
async function autoLoad() {
  showLoading('Loading live data from Google Sheets…');
  try {
    // Try Google Sheets first
    let r = await fetch('/api/load_sheet', {method:'POST'});
    let d = await r.json();
    if (!r.ok) {
      // Fallback to local file
      showLoading('Falling back to local file…');
      r = await fetch('/api/auto_load', {method:'POST'});
      d = await r.json();
    }
    if (r.ok) {
      document.getElementById('dataStatus').textContent = `Loaded: ${d.filename} (${d.count} records)`;
      document.getElementById('welcomeState').classList.add('hidden');
      document.getElementById('dashboard').classList.remove('hidden');
      notify(d.message);
      refresh();
    } else {
      notify(d.error, 'error');
    }
  } catch(e) { notify('Connection error', 'error'); }
  hideLoading();
}

async function handleFileUpload(e) {
  const file = e.target.files[0];
  if (!file) return;
  showLoading('Uploading & processing…');
  const fd = new FormData();
  fd.append('file', file);
  try {
    const r = await fetch('/api/upload', {method:'POST', body: fd});
    const d = await r.json();
    if (r.ok) {
      document.getElementById('dataStatus').textContent = `Loaded: ${file.name} (${d.count} records)`;
      document.getElementById('welcomeState').classList.add('hidden');
      document.getElementById('dashboard').classList.remove('hidden');
      notify(d.message);
      refresh();
    } else {
      notify(d.error, 'error');
    }
  } catch(e) { notify('Upload failed', 'error'); }
  hideLoading();
  e.target.value = '';
}

// ── Filter Params ────────────────────────────────────────────────────
function getFilterParams() {
  const p = new URLSearchParams();
  const s = document.getElementById('searchInput').value.trim();
  const c = document.getElementById('filterCompany').value;
  const m = document.getElementById('filterMonth').value;
  const y = document.getElementById('filterYear').value;
  const d = document.getElementById('filterDept').value;
  if (s) p.set('search', s);
  if (c) p.set('company', c);
  if (m) p.set('month', m);
  if (y) p.set('year', y);
  if (d) p.set('department', d);
  return p.toString();
}

function clearFilters() {
  document.getElementById('searchInput').value = '';
  document.getElementById('filterCompany').value = '';
  document.getElementById('filterMonth').value = '';
  document.getElementById('filterYear').value = '';
  document.getElementById('filterDept').value = '';
  refresh();
}

// ── Main Refresh ─────────────────────────────────────────────────────
async function refresh() {
  const q = getFilterParams();
  const [statsRes, dataRes, alertsRes] = await Promise.all([
    fetch(`/api/stats?${q}`), fetch(`/api/data?${q}`), fetch(`/api/alerts`)
  ]);
  const stats = await statsRes.json();
  const data = await dataRes.json();
  const alerts = await alertsRes.json();

  updateStats(stats);
  updateFilters(stats.filters);
  updateCompanyCards(stats.companies || []);
  updateCharts(stats);
  updateAlerts(alerts.alerts);
  allRecords = data.records || [];
  currentPage = 1;
  renderTable();
}

// ── Stats Update ─────────────────────────────────────────────────────
function updateStats(s) {
  animateNumber('statEmployees', s.unique_employees || 0);
  animateNumber('statDays', s.total_days || 0);
  animateNumber('statWarnings', s.warnings || 0);
  animateNumber('statCriticals', s.criticals || 0);
}

function animateNumber(id, target) {
  const el = document.getElementById(id);
  const start = parseInt(el.textContent) || 0;
  const diff = target - start;
  const steps = 30;
  let step = 0;
  const timer = setInterval(() => {
    step++;
    el.textContent = Math.round(start + (diff * (step/steps)));
    if (step >= steps) { el.textContent = target; clearInterval(timer); }
  }, 16);
}

// ── Filter Population ────────────────────────────────────────────────
function updateFilters(f) {
  if (!f) return;
  const monthNames = ['','January','February','March','April','May','June','July','August','September','October','November','December'];
  fillSelect('filterMonth', (f.months||[]).map(m => ({v:m, l:monthNames[m]||m})));
  fillSelect('filterYear', (f.years||[]).map(y => ({v:y, l:y})));
  fillSelect('filterDept', (f.departments||[]).map(d => ({v:d, l:d})));
  fillSelect('filterCompany', (f.companies||[]).map(c => ({v:c, l:c})));
}

function fillSelect(id, opts) {
  const sel = document.getElementById(id);
  const cur = sel.value;
  const first = sel.options[0].textContent;
  sel.innerHTML = `<option value="">${first}</option>`;
  opts.forEach(o => { const op = document.createElement('option'); op.value = o.v; op.textContent = o.l; sel.appendChild(op); });
  sel.value = cur;
}

// ── Company Cards ────────────────────────────────────────────────────
function updateCompanyCards(companies) {
  const container = document.getElementById('companyCards');
  document.getElementById('companyCount').textContent = `${companies.length} companies`;
  if (!companies.length) {
    container.innerHTML = '<p class="text-slate-500 text-sm col-span-full">No company data available</p>';
    return;
  }

  const colors = [
    {bg:'from-accent-500/20 to-accent-600/5', border:'border-accent-500/20', text:'text-accent-400', glow:'glow-blue'},
    {bg:'from-emerald-500/20 to-emerald-600/5', border:'border-emerald-500/20', text:'text-emerald-400', glow:'glow-emerald'},
    {bg:'from-amber-500/20 to-amber-600/5', border:'border-amber-500/20', text:'text-amber-400', glow:'glow-amber'},
    {bg:'from-rose-500/20 to-rose-600/5', border:'border-rose-500/20', text:'text-rose-400', glow:'glow-rose'},
    {bg:'from-violet-500/20 to-violet-600/5', border:'border-violet-500/20', text:'text-violet-400', glow:''},
  ];

  container.innerHTML = companies.map((c, i) => {
    const col = colors[i % colors.length];
    return `
      <div class="glass rounded-xl p-5 card-hover bg-gradient-to-br ${col.bg} border ${col.border} ${col.glow} fade-in"
           onclick="openCompanyDetail('${c.name.replace(/'/g,"\\'")}')">
        <div class="flex items-center gap-3 mb-4">
          <div class="w-10 h-10 rounded-lg bg-slate-800/50 flex items-center justify-center text-xl">🏢</div>
          <div class="flex-1 min-w-0">
            <h4 class="text-white font-semibold text-sm truncate">${c.name}</h4>
            <p class="text-xs text-slate-400">${c.department_count} department${c.department_count!==1?'s':''}</p>
          </div>
        </div>
        <div class="grid grid-cols-2 gap-3">
          <div>
            <div class="${col.text} text-xl font-bold">${c.employee_count}</div>
            <div class="text-xs text-slate-500">Employees</div>
          </div>
          <div>
            <div class="${col.text} text-xl font-bold">${c.total_days}</div>
            <div class="text-xs text-slate-500">Absent Days</div>
          </div>
        </div>
        <div class="mt-4 flex flex-wrap gap-1.5">
          ${c.departments.slice(0,4).map(d => `<span class="text-[10px] px-2 py-0.5 rounded-full bg-slate-800/50 text-slate-400">${d}</span>`).join('')}
          ${c.departments.length > 4 ? `<span class="text-[10px] px-2 py-0.5 rounded-full bg-slate-800/50 text-slate-400">+${c.departments.length-4} more</span>` : ''}
        </div>
        <div class="mt-3 text-xs ${col.text} font-medium flex items-center gap-1">
          Click to view details
          <svg class="w-3 h-3" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M9 18l6-6-6-6"/></svg>
        </div>
      </div>`;
  }).join('');
}

// ── Company Detail Panel ─────────────────────────────────────────────
async function openCompanyDetail(companyName) {
  const panel = document.getElementById('companyDetail');
  const section = document.getElementById('companySection');
  document.getElementById('breadcrumbCompany').textContent = companyName;
  document.getElementById('breadcrumbDept').classList.add('hidden');

  try {
    const r = await fetch(`/api/company/${encodeURIComponent(companyName)}`);
    const d = await r.json();
    if (!r.ok) { notify(d.error, 'error'); return; }

    // Stats
    document.getElementById('companyDetailStats').innerHTML = `
      <div class="text-center">
        <div class="text-2xl font-bold text-white">${d.unique_employees}</div>
        <div class="text-xs text-slate-400">Employees</div>
      </div>
      <div class="text-center">
        <div class="text-2xl font-bold text-white">${d.total_days}</div>
        <div class="text-xs text-slate-400">Total Absent Days</div>
      </div>
      <div class="text-center">
        <div class="text-2xl font-bold text-white">${d.departments.length}</div>
        <div class="text-xs text-slate-400">Departments</div>
      </div>`;

    // Department list
    document.getElementById('companyDetailBody').innerHTML = `
      <h4 class="text-sm font-semibold text-white mb-3">Departments</h4>
      <div class="space-y-2">
        ${d.departments.map(dept => {
          const pct = d.total_days > 0 ? Math.round((dept.total_days / d.total_days) * 100) : 0;
          const statusColor = dept.total_days >= 50 ? 'bg-rose-500' : dept.total_days >= 20 ? 'bg-amber-500' : 'bg-emerald-500';
          return `
            <div class="dept-row flex items-center gap-4 p-3 rounded-lg cursor-pointer transition-colors border border-transparent hover:border-accent-500/20"
                 onclick="openDeptDetail('${companyName.replace(/'/g,"\\'")}', '${dept.name.replace(/'/g,"\\'")}')">
              <div class="w-8 h-8 rounded-lg bg-slate-800/50 flex items-center justify-center text-sm">📁</div>
              <div class="flex-1 min-w-0">
                <div class="flex items-center gap-2">
                  <span class="text-white font-medium text-sm">${dept.name}</span>
                  <span class="text-[10px] px-2 py-0.5 rounded-full bg-slate-800 text-slate-400">${dept.employee_count} employees</span>
                </div>
                <div class="mt-1.5 flex items-center gap-2">
                  <div class="flex-1 h-1.5 bg-slate-800 rounded-full overflow-hidden">
                    <div class="${statusColor} h-full rounded-full transition-all duration-700" style="width:${pct}%"></div>
                  </div>
                  <span class="text-xs text-slate-400 font-mono w-16 text-right">${dept.total_days} days</span>
                </div>
              </div>
              <svg class="w-4 h-4 text-slate-500" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M9 18l6-6-6-6"/></svg>
            </div>`;
        }).join('')}
      </div>`;

    section.classList.add('hidden');
    panel.classList.remove('hidden');
    panel.scrollIntoView({behavior:'smooth', block:'start'});
  } catch(e) { notify('Error loading company details', 'error'); }
}

async function openDeptDetail(companyName, deptName) {
  document.getElementById('breadcrumbDept').classList.remove('hidden');
  document.getElementById('breadcrumbDeptName').textContent = deptName;
  document.getElementById('breadcrumbCompany').onclick = () => openCompanyDetail(companyName);

  try {
    const r = await fetch(`/api/department/${encodeURIComponent(companyName)}/${encodeURIComponent(deptName)}`);
    const d = await r.json();
    if (!r.ok) { notify(d.error, 'error'); return; }

    document.getElementById('companyDetailStats').innerHTML = `
      <div class="text-center">
        <div class="text-2xl font-bold text-white">${d.unique_employees}</div>
        <div class="text-xs text-slate-400">Employees</div>
      </div>
      <div class="text-center">
        <div class="text-2xl font-bold text-white">${d.total_days}</div>
        <div class="text-xs text-slate-400">Total Absent Days</div>
      </div>
      <div class="text-center">
        <div class="text-2xl font-bold text-white">${deptName}</div>
        <div class="text-xs text-slate-400">Department</div>
      </div>`;

    document.getElementById('companyDetailBody').innerHTML = `
      <h4 class="text-sm font-semibold text-white mb-3">Employees <span class="text-slate-400 font-normal">(${d.employees.length})</span></h4>
      <div class="space-y-1">
        ${d.employees.map((emp, i) => {
          const statusColors = { Critical:'bg-rose-500/10 text-rose-400 border-rose-500/20', Warning:'bg-amber-500/10 text-amber-400 border-amber-500/20', Normal:'bg-emerald-500/10 text-emerald-400 border-emerald-500/20' };
          const sc = statusColors[emp.status] || statusColors.Normal;
          return `
            <div class="emp-row flex items-center gap-4 p-3 rounded-lg cursor-pointer transition-colors border border-transparent hover:border-accent-500/20"
                 onclick="showProfile('${emp.name.replace(/'/g,"\\'")}')">
              <div class="w-7 h-7 rounded-full bg-gradient-to-br from-accent-500/30 to-accent-600/10 flex items-center justify-center text-xs text-accent-400 font-bold">${i+1}</div>
              <div class="flex-1 min-w-0">
                <div class="text-white text-sm font-medium">${emp.name}</div>
                <div class="text-xs text-slate-500">${emp.position} · ${emp.record_count} record${emp.record_count!==1?'s':''}</div>
                ${emp.reason && emp.reason !== 'N/A' ? `<div class="text-xs text-amber-400/70 mt-0.5">📋 ${emp.reason}</div>` : ''}
              </div>
              <div class="text-right">
                <div class="text-sm font-bold text-white">${emp.total_days} days</div>
                <span class="text-[10px] px-2 py-0.5 rounded-full border ${sc}">${emp.status}</span>
              </div>
              <svg class="w-4 h-4 text-slate-500" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M9 18l6-6-6-6"/></svg>
            </div>`;
        }).join('')}
      </div>`;
  } catch(e) { notify('Error loading department details', 'error'); }
}

function closeCompanyDetail() {
  document.getElementById('companyDetail').classList.add('hidden');
  document.getElementById('companySection').classList.remove('hidden');
}

// ── Charts ───────────────────────────────────────────────────────────
const chartDefaults = {
  responsive: true, maintainAspectRatio: false,
  plugins: { legend: { display: false } },
  scales: {
    x: { grid: { color: 'rgba(51,65,85,0.3)' }, ticks: { color: '#94a3b8', font: { size: 10 } } },
    y: { grid: { color: 'rgba(51,65,85,0.3)' }, ticks: { color: '#94a3b8', font: { size: 10 } } }
  }
};

function updateCharts(s) {
  // Department chart
  if (deptChartInstance) deptChartInstance.destroy();
  const deptCtx = document.getElementById('deptChart').getContext('2d');
  const deptColors = (s.dept_chart?.labels||[]).map((_,i) => {
    const hues = [199, 160, 38, 350, 270, 120, 30, 210, 300, 60];
    return `hsl(${hues[i%hues.length]}, 70%, 55%)`;
  });
  deptChartInstance = new Chart(deptCtx, {
    type: 'bar',
    data: { labels: s.dept_chart?.labels||[], datasets: [{ data: s.dept_chart?.values||[], backgroundColor: deptColors, borderRadius: 6, barThickness: 28 }] },
    options: { ...chartDefaults, onClick: (e, el) => { if (el.length) { const label = s.dept_chart.labels[el[0].index]; document.getElementById('filterDept').value = label; refresh(); } } }
  });

  // Monthly chart
  if (monthChartInstance) monthChartInstance.destroy();
  const monthCtx = document.getElementById('monthChart').getContext('2d');
  monthChartInstance = new Chart(monthCtx, {
    type: 'line',
    data: { labels: s.month_chart?.labels||[], datasets: [{ data: s.month_chart?.values||[], borderColor: '#38bdf8', backgroundColor: 'rgba(56,189,248,0.1)', fill: true, tension: 0.4, pointRadius: 4, pointHoverRadius: 7, pointBackgroundColor: '#38bdf8' }] },
    options: chartDefaults
  });

  // Top 10 chart
  if (top10ChartInstance) top10ChartInstance.destroy();
  const top10Ctx = document.getElementById('top10Chart').getContext('2d');
  const barColors = (s.top10?.values||[]).map(v => v >= 5 ? '#f43f5e' : v >= 3 ? '#f59e0b' : '#10b981');
  top10ChartInstance = new Chart(top10Ctx, {
    type: 'bar',
    data: { labels: s.top10?.labels||[], datasets: [{ data: s.top10?.values||[], backgroundColor: barColors, borderRadius: 6, barThickness: 20 }] },
    options: {
      ...chartDefaults, indexAxis: 'y',
      onClick: (e, el) => { if (el.length) showProfile(s.top10.labels[el[0].index]); }
    }
  });
}

// ── Alerts ────────────────────────────────────────────────────────────
function updateAlerts(alerts) {
  document.getElementById('alertCount').textContent = alerts.length;
  const list = document.getElementById('alertList');
  if (!alerts.length) { list.innerHTML = '<p class="text-slate-500 text-sm">No alerts</p>'; return; }
  list.innerHTML = alerts.slice(0, 50).map(a => {
    const isC = a.status === 'Critical';
    const dot = isC ? 'bg-rose-500' : 'bg-amber-500';
    const badge = isC ? 'bg-rose-500/10 text-rose-400 border-rose-500/20' : 'bg-amber-500/10 text-amber-400 border-amber-500/20';
    return `
      <div class="flex items-center gap-3 p-3 rounded-lg hover:bg-slate-800/30 cursor-pointer transition-colors" onclick="showProfile('${a.name.replace(/'/g,"\\'")}')">
        <div class="pulse-dot ${dot}"></div>
        <div class="flex-1 min-w-0">
          <div class="text-white text-sm font-medium">${a.name}</div>
          <div class="text-xs text-slate-500">${a.company} · ${a.department} · ${a.month} · ${a.total_absences} day(s)</div>
        </div>
        <span class="text-[10px] px-2 py-0.5 rounded-full border ${badge} whitespace-nowrap">${a.status}</span>
      </div>`;
  }).join('');
}

// ── Table ─────────────────────────────────────────────────────────────
function renderTable() {
  const total = allRecords.length;
  const pages = Math.ceil(total / PAGE_SIZE);
  currentPage = Math.min(currentPage, pages || 1);
  const start = (currentPage - 1) * PAGE_SIZE;
  const page = allRecords.slice(start, start + PAGE_SIZE);
  document.getElementById('tableCount').textContent = `(${total} records)`;

  const tbody = document.getElementById('tableBody');
  tbody.innerHTML = page.map(r => {
    const statusColors = { Critical:'bg-rose-500/10 text-rose-400 border-rose-500/20', Warning:'bg-amber-500/10 text-amber-400 border-amber-500/20', Normal:'bg-emerald-500/10 text-emerald-400 border-emerald-500/20' };
    const sc = statusColors[r.status] || statusColors.Normal;
    const reasonText = r.reason && r.reason !== 'N/A' ? r.reason : '<span class="text-slate-600">—</span>';
    return `
      <tr class="hover:bg-slate-800/30 cursor-pointer transition-colors" onclick="showProfile('${r.name.replace(/'/g,"\\'")}')">
        <td class="px-5 py-3 text-white font-medium">${r.name}</td>
        <td class="px-5 py-3 text-slate-400 text-xs">${r.company||''}</td>
        <td class="px-5 py-3"><span class="text-xs px-2 py-0.5 rounded-full bg-slate-800 text-slate-300">${r.department}</span></td>
        <td class="px-5 py-3 text-slate-400">${r.month}</td>
        <td class="px-5 py-3 text-slate-400">${r.year||''}</td>
        <td class="px-5 py-3 text-white font-bold">${r.total_absences}</td>
        <td class="px-5 py-3 text-slate-400 text-xs max-w-[200px] truncate" title="${(r.reason||'').replace(/"/g,'&quot;')}">${reasonText}</td>
        <td class="px-5 py-3"><span class="text-[10px] px-2 py-0.5 rounded-full border ${sc}">${r.status}</span></td>
      </tr>`;
  }).join('');

  // Pagination
  const pg = document.getElementById('pagination');
  if (pages <= 1) { pg.innerHTML = ''; return; }
  let html = '';
  const btn = (p, label, active) => `<button onclick="goPage(${p})" class="px-3 py-1.5 text-xs rounded-lg transition-colors ${active ? 'bg-accent-500 text-white' : 'bg-slate-800 text-slate-400 hover:bg-slate-700'}">${label}</button>`;
  if (currentPage > 1) html += btn(currentPage-1, '← Prev', false);
  for (let p = Math.max(1,currentPage-2); p <= Math.min(pages,currentPage+2); p++) html += btn(p, p, p===currentPage);
  if (currentPage < pages) html += btn(currentPage+1, 'Next →', false);
  pg.innerHTML = html;
}
function goPage(p) { currentPage = p; renderTable(); document.querySelector('.glass.rounded-xl.overflow-hidden.mb-6')?.scrollIntoView({behavior:'smooth'}); }

// ── Employee Profile Modal ───────────────────────────────────────────
async function showProfile(name) {
  const modal = document.getElementById('profileModal');
  modal.classList.remove('hidden');
  modal.classList.add('flex');
  document.getElementById('profileName').textContent = name;
  document.getElementById('profileBody').innerHTML = '<div class="flex justify-center py-8"><div class="spinner"></div></div>';

  try {
    const r = await fetch(`/api/employee/${encodeURIComponent(name)}`);
    const d = await r.json();
    if (!r.ok) { document.getElementById('profileBody').innerHTML = `<p class="text-rose-400">${d.error}</p>`; return; }

    const statusColors = { Critical:'bg-rose-500/10 text-rose-400 border-rose-500/20', Warning:'bg-amber-500/10 text-amber-400 border-amber-500/20', Normal:'bg-emerald-500/10 text-emerald-400 border-emerald-500/20' };
    const sc = statusColors[d.overall_status] || statusColors.Normal;

    document.getElementById('profileBody').innerHTML = `
      <div class="grid grid-cols-2 gap-4 mb-6">
        <div class="bg-slate-800/50 rounded-lg p-4">
          <div class="text-xs text-slate-400 mb-1">Company</div>
          <div class="text-white font-medium text-sm">${d.company || 'N/A'}</div>
        </div>
        <div class="bg-slate-800/50 rounded-lg p-4">
          <div class="text-xs text-slate-400 mb-1">Department</div>
          <div class="text-white font-medium text-sm">${d.department}</div>
        </div>
        <div class="bg-slate-800/50 rounded-lg p-4">
          <div class="text-xs text-slate-400 mb-1">Position</div>
          <div class="text-white font-medium text-sm">${d.position}</div>
        </div>
        <div class="bg-slate-800/50 rounded-lg p-4">
          <div class="text-xs text-slate-400 mb-1">Overall Status</div>
          <span class="text-xs px-2 py-0.5 rounded-full border ${sc}">${d.overall_status}</span>
        </div>
      </div>

      <div class="grid grid-cols-2 gap-4 mb-6">
        <div class="bg-accent-500/10 border border-accent-500/20 rounded-lg p-4 text-center">
          <div class="text-2xl font-bold text-accent-400">${d.total_days}</div>
          <div class="text-xs text-slate-400">Total Absent Days</div>
        </div>
        <div class="bg-slate-800/50 rounded-lg p-4 text-center">
          <div class="text-2xl font-bold text-white">${d.total_records}</div>
          <div class="text-xs text-slate-400">Records</div>
        </div>
      </div>

      <h4 class="text-sm font-semibold text-white mb-3">📅 Monthly Breakdown</h4>
      <div class="space-y-3 mb-6">
        ${d.monthly_breakdown.map(m => {
          const msc = statusColors[m.status] || statusColors.Normal;
          const barW = Math.min(Math.round((m.days / 23) * 100), 100);
          const barColor = m.status === 'Critical' ? 'bg-rose-500' : m.status === 'Warning' ? 'bg-amber-500' : 'bg-emerald-500';
          const datesList = (m.dates || []).map(dt => {
            const returnDate = new Date(dt.date + 'T00:00:00');
            const days = dt.days_absent;
            const reasonText = dt.reason && dt.reason !== 'N/A' ? dt.reason : '';
            let dateLabel;
            if (days > 1) {
              // Date in the form is the RETURN date, so absence = (date - days) to (date - 1)
              let absenceStart = new Date(returnDate);
              absenceStart.setDate(absenceStart.getDate() - days);
              let absenceEnd = new Date(returnDate);
              absenceEnd.setDate(absenceEnd.getDate() - 1);
              const startFmt = absenceStart.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
              const endFmt = absenceEnd.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
              dateLabel = `${startFmt} to ${endFmt}`;
            } else {
              // Single day: absent the day before the return date
              let absenceDay = new Date(returnDate);
              absenceDay.setDate(absenceDay.getDate() - 1);
              dateLabel = absenceDay.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
            }
            return `
              <div class="flex items-start gap-2 py-1.5 border-b border-slate-800/30 last:border-0">
                <span class="text-[11px] text-accent-400 font-mono flex-shrink-0" style="min-width:120px">${dateLabel}</span>
                <span class="text-[10px] text-slate-500 w-14 flex-shrink-0 text-center">${days} day${days > 1 ? 's' : ''}</span>
                ${reasonText ? `<span class="text-[10px] text-amber-400/80 flex-1">📋 ${reasonText}</span>` : `<span class="text-[10px] text-slate-600 flex-1">—</span>`}
              </div>`;
          }).join('');
          return `
            <div class="p-3 rounded-xl bg-slate-800/30 border border-slate-700/30">
              <div class="flex items-center gap-3 mb-2">
                <span class="text-xs text-slate-300 font-semibold w-28">${m.period}</span>
                <div class="flex-1 h-2 bg-slate-800 rounded-full overflow-hidden">
                  <div class="${barColor} h-full rounded-full transition-all duration-700" style="width:${barW}%"></div>
                </div>
                <span class="text-xs text-white font-bold w-12 text-right">${m.days} d</span>
                <span class="text-[10px] px-2 py-0.5 rounded-full border ${msc}">${m.status}</span>
              </div>
              ${datesList ? `<div class="ml-1 mt-2 pl-3 border-l-2 border-slate-700/50">${datesList}</div>` : ''}
            </div>`;
        }).join('')}
      </div>

      ${d.reasons && d.reasons.length ? `
        <h4 class="text-sm font-semibold text-white mb-3">📋 Reasons for Absence</h4>
        <div class="flex flex-wrap gap-1.5 mb-6">
          ${d.reasons.map(r => `<span class="text-xs px-3 py-1 rounded-full bg-amber-500/10 text-amber-400 border border-amber-500/20">${r}</span>`).join('')}
        </div>` : ''}
    `;
  } catch(e) { document.getElementById('profileBody').innerHTML = '<p class="text-rose-400">Failed to load profile</p>'; }
}

function closeModal() {
  const modal = document.getElementById('profileModal');
  modal.classList.add('hidden');
  modal.classList.remove('flex');
}

// ── Export / Dropdown ────────────────────────────────────────────────
function toggleDropdown() { document.getElementById('dropdownMenu').classList.toggle('hidden'); }
function exportReport(type) {
  window.location.href = `/api/export/${type}`;
  toggleDropdown();
  notify(`Downloading ${type} report…`);
}

// ── Employee List Modal ──────────────────────────────────────────────
let allEmployeeList = [];

async function showEmployeeList() {
  const modal = document.getElementById('employeeListModal');
  modal.classList.remove('hidden');
  modal.classList.add('flex');
  document.getElementById('empListBody').innerHTML = '<div class="text-center text-slate-400 py-8">Loading employees…</div>';
  document.getElementById('empListSearch').value = '';

  // Build query from current filters
  const params = new URLSearchParams();
  const search = document.getElementById('searchInput').value.trim();
  const company = document.getElementById('filterCompany').value;
  const month = document.getElementById('filterMonth').value;
  const year = document.getElementById('filterYear').value;
  const dept = document.getElementById('filterDept').value;
  if (search) params.set('search', search);
  if (company) params.set('company', company);
  if (month) params.set('month', month);
  if (year) params.set('year', year);
  if (dept) params.set('department', dept);

  try {
    const res = await fetch(`/api/employees?${params}`);
    const data = await res.json();
    allEmployeeList = data.employees || [];
    document.getElementById('empListCount').textContent = `${allEmployeeList.length} employees found`;
    renderEmpList(allEmployeeList);
  } catch (e) {
    document.getElementById('empListBody').innerHTML = '<div class="text-center text-rose-400 py-8">Failed to load employees</div>';
  }
}

function closeEmployeeList() {
  const modal = document.getElementById('employeeListModal');
  modal.classList.add('hidden');
  modal.classList.remove('flex');
}

function filterEmployeeList() {
  const q = document.getElementById('empListSearch').value.trim().toLowerCase();
  const filtered = q ? allEmployeeList.filter(e => e.name.toLowerCase().includes(q) || (e.department||'').toLowerCase().includes(q) || (e.company||'').toLowerCase().includes(q)) : allEmployeeList;
  document.getElementById('empListCount').textContent = `${filtered.length} of ${allEmployeeList.length} employees`;
  renderEmpList(filtered);
}

function renderEmpList(employees) {
  const body = document.getElementById('empListBody');
  if (!employees.length) {
    body.innerHTML = '<div class="text-center text-slate-400 py-8">No employees found</div>';
    return;
  }
  const statusColors = { Critical:'bg-rose-500/10 text-rose-400 border-rose-500/20', Warning:'bg-amber-500/10 text-amber-400 border-amber-500/20', Normal:'bg-emerald-500/10 text-emerald-400 border-emerald-500/20' };
  body.innerHTML = employees.map((emp, i) => {
    const sc = statusColors[emp.status] || statusColors.Normal;
    const reasonText = emp.reason && emp.reason !== 'N/A' ? `<div class="text-[10px] text-amber-400/70 mt-0.5 truncate max-w-[300px]" title="${(emp.reason||'').replace(/"/g,'&quot;')}">📋 ${emp.reason}</div>` : '';
    return `
      <div class="flex items-center gap-4 p-3 rounded-xl hover:bg-slate-800/40 cursor-pointer transition-colors border-b border-slate-800/30" onclick="closeEmployeeList(); showProfile('${emp.name.replace(/'/g,"\\'")}')">
        <div class="w-9 h-9 rounded-full bg-gradient-to-br from-accent-500/30 to-cyan-500/30 flex items-center justify-center text-white text-sm font-bold flex-shrink-0">${i + 1}</div>
        <div class="flex-1 min-w-0">
          <div class="text-white text-sm font-medium">${emp.name}</div>
          <div class="text-[10px] text-slate-500">${emp.company || 'Unknown'} · ${emp.department} · ${emp.position || 'N/A'}</div>
          ${reasonText}
        </div>
        <div class="text-right flex-shrink-0">
          <div class="text-sm font-bold text-white">${emp.total_days} days</div>
          <span class="text-[10px] px-2 py-0.5 rounded-full border ${sc}">${emp.status}</span>
        </div>
        <span class="text-slate-600 text-lg">›</span>
      </div>`;
  }).join('');
}
