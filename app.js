// ===================================
// APP.JS — Main Application Logic
// ===================================
/* global Chart, XLSX */

let currentPage = 'dashboard';
let currentPeriod = 'month';
let catePieMode  = 'day';  // 'day' | 'week'
let dashboardDate = null; // null = hôm nay
let importPage = 1;
let exportPage = 1;
let inventoryPage = 1;
let bookingPage = 1;
let fcMonth = 0;
let reportPeriod = 'month';
let activeCharts = [];

const importFilter  = { search: '', statuses: [], partner: 'all', category: 'all', dateFrom: '', dateTo: '' };
const exportFilter  = { search: '', statuses: [], partner: 'all', chiNhanh: 'all', category: 'all', dateFrom: '', dateTo: '' };
const bookingFilter = { search: '', supplier: 'all', store: 'all', tempCategory: 'all', dateFrom: '', dateTo: '' };


const PAGE_META = {
  dashboard: ['Dashboard',   'Tổng quan hoạt động kho hôm nay'],
  imports:   ['Nhập hàng',   'Quản lý phiếu nhập kho'],
  exports:   ['Xuất hàng',   'Quản lý phiếu xuất kho'],
  inventory: ['Tồn kho',     'Danh sách hàng hóa trong kho'],
  booking:   ['Booking',     'Danh sách đặt hàng từ cửa hàng'],
  fc:        ['FC',          'Forecast Control — Dự báo nhập hàng'],
  reports:    ['Báo cáo',    'Phân tích dữ liệu kho hàng'],
  pricelist:  ['Bảng giá',  'Danh sách giá bán sản phẩm'],
  masterdata: ['Mã hàng',   'Master data thịt cá — mã hàng, quy cách, trọng lượng'],
};

// ============================================================
// INIT
// ============================================================
document.addEventListener('DOMContentLoaded', async () => {
  // Load persisted data (IndexedDB — async)
  await MockData.init();

  document.getElementById('header-date').textContent =
    new Date().toLocaleDateString('vi-VN', { weekday: 'long', day: '2-digit', month: '2-digit', year: 'numeric' });

  refreshLowStockBadge();

  document.querySelectorAll('.nav-item[data-page]').forEach(item => {
    item.addEventListener('click', () => navigateTo(item.dataset.page));
  });

  // Import Excel nav button
  document.getElementById('nav-import-excel').addEventListener('click', showImportExcelModal);

  document.getElementById('modal-close').addEventListener('click', closeModal);
  document.getElementById('modal-overlay').addEventListener('click', e => {
    if (e.target.id === 'modal-overlay') closeModal();
  });

  document.getElementById('btn-add').addEventListener('click', () => showAddModal());

  document.getElementById('global-search').addEventListener('input',
    Utils.debounce(e => {
      const q = e.target.value;
      if (currentPage === 'inventory') { inventoryPage = 1; renderInventory(q); }
    }, 300)
  );

  // Show visible error if localStorage save fails (e.g. quota exceeded)
  document.addEventListener('kfm-save-error', e => {
    showToast('⚠️ Không thể lưu dữ liệu vào localStorage. Dùng nút Backup JSON để tránh mất dữ liệu!', 'red');
    console.warn('[KFM] kfm-save-error:', e.detail);
  });


  console.log('[KFM] Diagnostic: run kfmWeightDiag() in console');

  // Weight diagnostic: traces exactly how import weight is calculated
  window.kfmWeightDiag = (dateStr) => {
    const now = new Date();
    const today = dateStr || `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}-${String(now.getDate()).padStart(2,'0')}`;
    const imports = MockData.getTodayImports().filter(t => t.status !== 'cancelled');
    console.log(`[Weight Diag] Ngày: ${today}`);
    console.log(`[Weight Diag] Tổng dòng nhập hôm nay (non-cancelled): ${imports.length}`);
    console.log(`[Weight Diag] Master Data rows: ${MockData.masterData.length}`);

    let totalWeight = 0;
    let missingCodes = new Set();
    let foundCount = 0;

    const rows = imports.map(t => {
      const code = String(t.barcode || '').trim();
      const mdEntry = MockData.masterData.find(r => String(r.code || '').trim().toUpperCase() === code.toUpperCase());
      const spec = mdEntry ? (parseFloat(mdEntry.spec) || 0) : null;
      const qty  = t.quantity || 0;
      const w    = spec != null ? qty * spec : 0;
      if (spec != null) foundCount++;
      else if (code) missingCodes.add(code);
      totalWeight += w;
      return { maPO: t.maPO, barcode: code, qty, spec: spec ?? 'MISSING', weight: w, date: t.date, status: t.status };
    });

    console.table(rows);
    console.log(`[Weight Diag] Mã có trong MD: ${foundCount}/${imports.length}`);
    console.log(`[Weight Diag] Mã THIẾU trong MD (${missingCodes.size}):`, [...missingCodes].join(', ') || 'không có');
    console.log(`[Weight Diag] TỔNG TRỌNG LƯỢNG = ${Math.round(totalWeight).toLocaleString('vi-VN')} kg`);

    // Show sample of master data codes for comparison
    console.log('[Weight Diag] 5 dòng đầu Master Data:',
      MockData.masterData.slice(0, 5).map(r => `${r.code} | spec=${r.spec}`)
    );
  };

  // Safety net: ensure data is flushed before the page unloads
  window.addEventListener('beforeunload', () => MockData.save());

  navigateTo('dashboard');
});

// ============================================================
// NAVIGATION
// ============================================================
function navigateTo(page) {
  currentPage = page;

  document.querySelectorAll('.nav-item[data-page]').forEach(el =>
    el.classList.toggle('active', el.dataset.page === page)
  );

  const [title, subtitle] = PAGE_META[page] || ['', ''];
  document.getElementById('page-title').textContent = title;
  document.getElementById('page-subtitle').textContent = subtitle;
  document.getElementById('global-search').value = '';

  destroyCharts();

  const content = document.getElementById('page-content');
  content.innerHTML = '';
  content.className = 'page-content anim-fade-in';

  switch (page) {
    case 'dashboard': renderDashboard(); break;
    case 'imports':
      importPage = 1;
      importFilter.search = ''; importFilter.statuses = []; importFilter.partner = 'all';
      importFilter.category = 'all'; importFilter.dateFrom = ''; importFilter.dateTo = '';
      renderImports(); break;
    case 'exports':
      exportPage = 1;
      exportFilter.search = ''; exportFilter.statuses = []; exportFilter.partner = 'all';
      exportFilter.chiNhanh = 'all'; exportFilter.category = 'all'; exportFilter.dateFrom = ''; exportFilter.dateTo = '';
      renderExports(); break;
    case 'inventory': inventoryPage = 1; renderInventory(); break;
    case 'booking':
      bookingPage = 1;
      bookingFilter.search = ''; bookingFilter.supplier = 'all'; bookingFilter.store = 'all';
      bookingFilter.tempCategory = 'all'; bookingFilter.dateFrom = ''; bookingFilter.dateTo = '';
      renderBookings(); break;
    case 'fc':         fcMonth = 0; renderFC(); break;
    case 'reports':    renderReports(); break;
    case 'pricelist':  renderPriceList(); break;
    case 'masterdata': renderMasterData(); break;
  }
}

function destroyCharts() {
  activeCharts.forEach(c => { try { c.destroy(); } catch (_) {} });
  activeCharts = [];
}

function makeChart(id, config) {
  const canvas = document.getElementById(id);
  if (!canvas) return null;
  // Destroy any existing Chart on this canvas (safe for partial rebuilds)
  const existing = Chart.getChart(canvas);
  if (existing) { existing.destroy(); activeCharts = activeCharts.filter(c => c !== existing); }
  // Disable datalabels globally unless the caller explicitly configures them
  const p = (config.options = config.options || {}).plugins = (config.options.plugins || {});
  if (!p.datalabels) p.datalabels = { display: false };
  const chart = new Chart(canvas, config);
  activeCharts.push(chart);
  return chart;
}

// ============================================================
// DASHBOARD
// ============================================================
function renderDashboard() {
  const hasAnyData = MockData.transactions.length || MockData.products.length ||
                     MockData.bookings.length || MockData.getFC();
  if (!hasAnyData) {
    document.getElementById('page-content').innerHTML = `
      <div style="display:flex;align-items:center;justify-content:center;min-height:60vh;">
        <div style="text-align:center;max-width:420px;">
          <div style="font-size:5rem;margin-bottom:var(--space-lg);">📂</div>
          <div style="font-size:var(--fs-2xl);font-weight:800;color:var(--text-heading);margin-bottom:var(--space-sm);">
            Chưa có dữ liệu
          </div>
          <div style="color:var(--text-muted);font-size:var(--fs-base);margin-bottom:var(--space-xl);line-height:1.8;">
            Import file Excel nhập hàng, xuất hàng, booking và FC<br>để bắt đầu phân tích.
          </div>
          <button class="btn btn-primary" style="padding:12px 32px;font-size:var(--fs-md);" onclick="showImportExcelModal()">
            📂 Import dữ liệu từ Excel
          </button>
        </div>
      </div>`;
    return;
  }

  const _now  = new Date();
  const today = `${_now.getFullYear()}-${String(_now.getMonth()+1).padStart(2,'0')}-${String(_now.getDate()).padStart(2,'0')}`;
  const selDate = dashboardDate || today;
  const isToday = selDate === today;

  const shiftDate = (d, n) => {
    const dt = new Date(d + 'T00:00:00');
    dt.setDate(dt.getDate() + n);
    return `${dt.getFullYear()}-${String(dt.getMonth()+1).padStart(2,'0')}-${String(dt.getDate()).padStart(2,'0')}`;
  };
  const fmtLabel = d => new Date(d + 'T00:00:00').toLocaleDateString('vi-VN',
    { weekday: 'short', day: '2-digit', month: '2-digit', year: 'numeric' });

  const lowStock      = MockData.getLowStockProducts();
  const activeImports = MockData.getImports().filter(t => t.date && t.date.startsWith(selDate) && t.status !== 'cancelled');
  const activeExports = MockData.getExports().filter(t => t.date && t.date.startsWith(selDate) && t.status !== 'cancelled');

  const importNCC  = new Set(activeImports.map(t => t.supplier).filter(Boolean)).size;
  const importSKUs = new Set(activeImports.filter(t => (t.quantity || 0) > 0).map(t => t.barcode).filter(Boolean)).size;
  const importPO   = new Set(activeImports.map(t => t.maPO).filter(Boolean)).size;
  const exportCH   = new Set(activeExports.map(t => t.customer).filter(Boolean)).size;
  const exportSKUs = new Set(activeExports.map(t => t.barcode).filter(Boolean)).size;
  const exportPT   = new Set(activeExports.map(t => t.maPO).filter(Boolean)).size;

  // FC — theo ngày được chọn
  const fc           = MockData.getFC();
  const selIdx       = fc ? fc.dates.indexOf(selDate) : -1;
  const fcToday      = selIdx >= 0 ? Math.round(fc.totalDailyValues?.[selIdx] ?? 0) : null;


  // Booking — theo ngày được chọn
  const todayBk       = MockData.bookings.filter(b => b.deliveryDate && b.deliveryDate.startsWith(selDate));
  const totalBkStores = new Set(todayBk.map(b => b.storeId || b.storeName).filter(Boolean)).size;
  const bkSKUs        = new Set(todayBk.map(b => b.codeSp).filter(Boolean)).size;
  const bkPO          = new Set(todayBk.map(b => b.po).filter(Boolean)).size;

  // Trọng lượng = số lượng × quy cách (từ Master Data)
  const hasMd = MockData.masterData.length > 0;
  const wCalc = (items, getQty, getCode) => Math.round(
    items.reduce((s, item) => {
      const { spec } = mdLookup(getCode(item));
      return s + (spec != null ? getQty(item) * spec : 0);
    }, 0)
  );
  const importWeight = wCalc(activeImports, t => t.quantity || 0,                                           t => t.barcode || '');
  const exportWeight = wCalc(activeExports, t => (t.transferQty != null ? t.transferQty : t.quantity) || 0, t => t.barcode || '');
  const bkWeight     = wCalc(todayBk,       b => b.quantity || 0,                                           b => b.codeSp  || '');
  const mdNote = hasMd ? '' : ' (chưa có MD)';

  document.getElementById('page-content').innerHTML = `
    <div style="display:grid;gap:var(--space-lg);">

      <!-- Date picker bar -->
      <div style="display:flex;align-items:center;gap:10px;justify-content:flex-end;">
        <button class="btn btn-ghost btn-sm" onclick="setDashboardDate('${shiftDate(selDate,-1)}')">← Hôm qua</button>
        <div style="position:relative;display:inline-flex;align-items:center;">
          <div class="filter-input" style="width:148px;text-align:center;cursor:pointer;user-select:none;
            display:flex;align-items:center;justify-content:center;gap:6px;padding:6px 12px;"
            onclick="this.nextElementSibling.showPicker?.();this.nextElementSibling.click();">
            <span>${selDate.slice(8,10)}/${selDate.slice(5,7)}/${selDate.slice(0,4)}</span>
            <span style="color:var(--text-muted);font-size:12px;">📅</span>
          </div>
          <input type="date" value="${selDate}" onchange="setDashboardDate(this.value)"
            style="position:absolute;inset:0;opacity:0;width:100%;cursor:pointer;z-index:1;">
        </div>
        <button class="btn btn-ghost btn-sm" onclick="setDashboardDate('${shiftDate(selDate,1)}')"
          ${isToday ? 'disabled style="opacity:0.4;cursor:not-allowed;"' : ''}>Ngày mai →</button>
        ${!isToday ? `<button class="btn btn-sm" style="border-color:var(--accent-blue);color:var(--accent-blue);"
          onclick="setDashboardDate(null)">⟳ Hôm nay</button>` : ''}
        <span style="color:var(--text-muted);font-size:var(--fs-xs);">${fmtLabel(selDate)}</span>
      </div>

      <div id="db-kpi-charts" style="display:grid;gap:var(--space-lg);">
      <!-- Green banner title -->
      <div class="anim-fade-in-up" style="background:linear-gradient(135deg,#16a34a,#22c55e);border-radius:var(--radius-lg);
        padding:14px 32px;text-align:center;box-shadow:0 4px 20px rgba(34,197,94,0.25);">
        <h2 style="margin:0;font-size:var(--fs-lg);font-weight:900;color:#fff;letter-spacing:0.08em;text-transform:uppercase;">
          BÁO CÁO SẢN LƯỢNG THỊT CÁ NGÀY ${selDate.slice(8,10)}/${selDate.slice(5,7)}/${selDate.slice(0,4)}
        </h2>
      </div>
      <!-- KPI summary cards -->
      <div class="grid-4">
        ${kpiCard('blue',   '📥', 'Nhập',      importWeight, 'num', `${importSKUs} SKUs · ${importPO} PO · ${importNCC} NCC`,               null, '', 'kg')}
        ${kpiCard('green',  '📤', 'Xuất',      exportWeight, 'num', `${exportSKUs} SKUs · ${exportPT} PT · ${exportCH} chi nhánh`,           null, '', 'kg')}
        ${kpiCard('teal',   '📋', 'Booking',   bkWeight,     'num', `${bkSKUs} SKUs · ${bkPO} PO · ${totalBkStores} chi nhánh`,              null, '', 'kg')}
        ${kpiCard('purple', '📊', 'FC hôm nay', fcToday !== null ? fcToday : 0, 'num', fcToday !== null ? '' : 'Chưa có dữ liệu FC',         null, '', 'kg')}
      </div>

      <!-- Row 2: Chart (2/3) + Pie chart (1/3) -->
      <div class="grid-2-1 anim-fade-in-up anim-delay-1">
        <div class="panel" style="display:flex;flex-direction:column;">
          <div class="panel-header">
            <div class="panel-title">Báo cáo Nhập/Xuất kho Thịt Cá</div>
            <div class="toggle-group">
              <button class="toggle-btn ${currentPeriod === 'day'   ? 'active' : ''}" onclick="changePeriod('day')">Ngày</button>
              <button class="toggle-btn ${currentPeriod === 'week'  ? 'active' : ''}" onclick="changePeriod('week')">Tuần</button>
              <button class="toggle-btn ${currentPeriod === 'month' ? 'active' : ''}" onclick="changePeriod('month')">Tháng</button>
            </div>
          </div>
          <div class="chart-container" style="flex:1;min-height:280px;">
            <canvas id="chart-trend"></canvas>
          </div>
        </div>

        <!-- Single pie chart with Ngày/Tuần/Tháng toggle -->
        <div class="panel" style="display:flex;flex-direction:column;">
          <div class="panel-header" style="flex-direction:column;align-items:center;gap:8px;padding:12px 16px 8px;">
            <div style="font-size:var(--fs-md);font-weight:900;color:#f59e0b;letter-spacing:0.1em;text-transform:uppercase;text-align:center;">
              % Đóng Góp Sản Lượng
            </div>
            <div class="toggle-group">
              <button class="toggle-btn cate-pie-btn ${catePieMode === 'day'   ? 'active' : ''}" data-mode="day"   onclick="changeCatePeriod('day')">Ngày</button>
              <button class="toggle-btn cate-pie-btn ${catePieMode === 'week'  ? 'active' : ''}" data-mode="week"  onclick="changeCatePeriod('week')">Tuần</button>
              <button class="toggle-btn cate-pie-btn ${catePieMode === 'month' ? 'active' : ''}" data-mode="month" onclick="changeCatePeriod('month')">Tháng</button>
            </div>
          </div>
          <div class="chart-container" style="flex:1;min-height:260px;">
            <canvas id="chart-cate-pie"></canvas>
          </div>
        </div>
      </div>
      </div> <!-- end #db-kpi-charts -->

      <!-- Row 3: Ops form full width -->
      <div id="db-ops-row" class="panel anim-fade-in-up anim-delay-2">
        <div class="panel-header">
          <div class="panel-title">📝 Báo cáo vận hành ngày</div>
        </div>
        <div>${renderDailyOpsForm(selDate)}</div>
      </div>

      <div class="grid-2-1 anim-fade-in-up anim-delay-3">
        <div class="panel">
          <div class="panel-header">
            <div class="panel-title">Giao dịch gần đây</div>
            <button class="btn btn-ghost btn-sm" onclick="navigateTo('imports')">Xem tất cả →</button>
          </div>
          <div class="panel-body" style="padding:0;">
            ${transactionTable(MockData.transactions.slice(0, 8))}
          </div>
        </div>
        <div class="panel">
          <div class="panel-header">
            <div class="panel-title">⚠️ Sắp hết hàng</div>
          </div>
          <div class="panel-body" style="padding:0;">
            ${lowStockList(lowStock)}
          </div>
        </div>
      </div>

    </div>`;

  animateKPIs();
  initTrendChart();
  initCateChart();
}

function setDashboardDate(d) {
  dashboardDate = d || null;
  renderDashboard();
}


function changeCatePeriod(mode) {
  catePieMode = mode;
  document.querySelectorAll('.cate-pie-btn').forEach(b => b.classList.toggle('active', b.dataset.mode === mode));
  initCateChart(); // makeChart auto-destroys previous instance on same canvas
}

function saveDailyNote(date, field, el) {
  MockData.setDailyNote(date, field, el.value);
}

function opsAutoResize(el) {
  el.style.height = 'auto';
  el.style.height = el.scrollHeight + 'px';
}

const OPS_DEFAULTS = {
  tinh_nhap: 'Không có vấn đề phát sinh',
  tinh_chia: 'Không có vấn đề phát sinh',
  tinh_xuat: 'Không có vấn đề phát sinh',
  giao_van:  'Giao hàng trước 6h',
};

function renderDailyOpsForm(date) {
  const n = MockData.getDailyNote(date);
  const esc = s => String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  // Use saved value if exists (even if empty string); otherwise fall back to default
  const val = field => n[field] !== undefined ? n[field] : (OPS_DEFAULTS[field] || '');
  const ta = field => `<textarea class="ops-ta"
    onblur="saveDailyNote('${date}','${field}',this)"
    oninput="opsAutoResize(this)"
    placeholder="${esc(OPS_DEFAULTS[field] || 'Nhập ghi chú...')}"
  >${esc(val(field))}</textarea>`;

  // Auto-resize all textareas after they render
  setTimeout(() => {
    document.querySelectorAll('.ops-ta').forEach(el => opsAutoResize(el));
  }, 0);

  return `
    <table class="ops-table">
      <tbody>
        <tr>
          <td class="ops-cat" rowspan="3">Tình hình<br>vận hành</td>
          <td class="ops-sub">Nhập</td>
          <td class="ops-cell">${ta('tinh_nhap')}</td>
        </tr>
        <tr>
          <td class="ops-sub">Chia</td>
          <td class="ops-cell">${ta('tinh_chia')}</td>
        </tr>
        <tr>
          <td class="ops-sub">Xuất</td>
          <td class="ops-cell">${ta('tinh_xuat')}</td>
        </tr>
        <tr>
          <td class="ops-cat">Giao hàng KF</td>
          <td class="ops-sub">Vận chuyển</td>
          <td class="ops-cell">${ta('giao_van')}</td>
        </tr>
      </tbody>
    </table>`;
}

function renderTodayBookingSummary(bookings) {
  if (!bookings.length) {
    return `<div style="text-align:center;padding:28px 0;color:var(--text-muted);">
      <div style="font-size:2rem;margin-bottom:8px;">📋</div>
      <div style="font-size:var(--fs-sm);">Không có booking giao hôm nay</div>
    </div>`;
  }

  const byCat    = { CHILL: 0, FROZEN: 0, DRY: 0 };
  const byStore  = {};
  bookings.forEach(b => {
    const cat = b.tempCategory;
    if (cat in byCat) byCat[cat] += b.quantity || 0;
    const store = b.storeName || b.storeId || '—';
    byStore[store] = (byStore[store] || 0) + (b.quantity || 0);
  });

  const catColors = { CHILL: '#06b6d4', FROZEN: '#8b5cf6', DRY: '#f59e0b' };
  const topStores = Object.entries(byStore).sort((a, b) => b[1] - a[1]).slice(0, 6);

  return `
    <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:8px;margin-bottom:12px;">
      ${Object.entries(byCat).map(([cat, qty]) => `
        <div style="text-align:center;background:var(--bg-primary);border-radius:var(--radius-sm);
          padding:10px 6px;border-top:2px solid ${catColors[cat]};">
          <div style="font-size:var(--fs-lg);font-weight:700;color:${catColors[cat]};">
            ${Utils.formatNumber(Math.round(qty))}
          </div>
          <div style="font-size:var(--fs-xs);color:var(--text-muted);">${cat}</div>
        </div>`).join('')}
    </div>
    <div style="font-size:var(--fs-xs);color:var(--text-muted);font-weight:600;
      padding:0 0 6px;border-bottom:1px solid var(--border-light);margin-bottom:6px;letter-spacing:0.05em;">
      TOP CỬA HÀNG
    </div>
    ${topStores.map(([store, qty]) => `
      <div style="display:flex;justify-content:space-between;align-items:center;
        padding:5px 0;border-bottom:1px solid rgba(255,255,255,0.04);">
        <span style="font-size:var(--fs-xs);color:var(--text-secondary);
          overflow:hidden;text-overflow:ellipsis;white-space:nowrap;max-width:60%;">${store}</span>
        <span style="font-size:var(--fs-xs);font-weight:600;color:#c084fc;">${Utils.formatNumber(Math.round(qty))}</span>
      </div>`).join('')}`;
}

function kpiCard(color, icon, label, value, fmt, sub, trend, trendDir, unit = '') {
  return `
    <div class="kpi-card ${color} anim-fade-in-up">
      <div class="kpi-card-header">
        <div class="kpi-card-icon">${icon}</div>
        ${trend ? `<div class="kpi-card-trend ${trendDir}">${trend}</div>` : ''}
      </div>
      <div style="display:flex;align-items:baseline;gap:5px;">
        <div class="kpi-card-value" data-target="${value}" data-format="${fmt}">
          ${fmt === 'vnd' ? Utils.formatVND(0) : Utils.formatNumber(0)}
        </div>
        ${unit ? `<span style="font-size:var(--fs-md);font-weight:600;color:var(--text-muted);">${unit}</span>` : ''}
      </div>
      <div class="kpi-card-label">${label}</div>
      <div class="text-xs text-muted mt-sm">${sub}</div>
    </div>`;
}

function animateKPIs() {
  document.querySelectorAll('.kpi-card-value').forEach(el => {
    Utils.animateCounter(el, parseInt(el.dataset.target) || 0);
  });
}

// Period toggle for trend chart
function changePeriod(period) {
  currentPeriod = period;
  const labels = { day: 'Ngày', week: 'Tuần', month: 'Tháng' };
  document.querySelectorAll('.toggle-btn:not(.cate-pie-btn)').forEach(b =>
    b.classList.toggle('active', b.textContent === labels[period])
  );
  destroyCharts();
  initTrendChart();
  initCateChart();
}

function buildSpecMap() {
  return new Map(MockData.masterData.map(r =>
    [String(r.code || '').trim().toUpperCase(), parseFloat(r.spec) || 0]
  ));
}

function txnWeight(t, specMap) {
  const s = specMap.get(String(t.barcode || '').trim().toUpperCase());
  return s != null ? (t.quantity || 0) * s : 0;
}

function initTrendChart() {
  const specMap = buildSpecMap();
  const txns   = MockData.transactions.filter(t => t.status !== 'cancelled');
  const groups = Utils.groupByPeriod(txns, currentPeriod);
  const keys   = Object.keys(groups).sort().slice(-14);
  const labels = keys.map(k => Utils.formatPeriodLabel(k, currentPeriod));

  const impData = keys.map(k => Math.round(groups[k].filter(t => t.type === 'import').reduce((s, t) => s + txnWeight(t, specMap), 0)));
  // Xuất dùng transferQty (SL chuyển) — đồng nhất với KPI card
  const expData = keys.map(k => Math.round(groups[k].filter(t => t.type === 'export').reduce((s, t) => {
    const spec = specMap.get(String(t.barcode || '').trim().toUpperCase());
    const qty  = (t.transferQty != null ? t.transferQty : t.quantity) || 0;
    return s + (spec != null ? qty * spec : 0);
  }, 0)));
  // Tỷ lệ Thực Xuất / FC (%)
  const fc = MockData.getFC();
  const getFCForKey = (k, groupTxns) => {
    if (!fc) return null;
    if (currentPeriod === 'day') {
      const idx = fc.dates.findIndex(d => d.startsWith(k));
      return idx >= 0 ? (fc.totalDailyValues?.[idx] ?? null) : null;
    } else if (currentPeriod === 'month') {
      const mi = parseInt(k.split('-')[1]) - 1;
      return fc.totalMonthlyAvg?.[mi] ?? null;
    } else {
      const datesInGroup = new Set((groupTxns || []).map(t => t.date?.slice(0,10)).filter(Boolean));
      const sum = fc.dates.reduce((s, d, i) => datesInGroup.has(d.slice(0,10)) ? s + (fc.totalDailyValues?.[i] ?? 0) : s, 0);
      return sum > 0 ? sum : null;
    }
  };
  const ratioData = keys.map((k, i) => {
    const fcVal = getFCForKey(k, groups[k]);
    return (fcVal != null && fcVal > 0) ? parseFloat((expData[i] / fcVal * 100).toFixed(2)) : null;
  });

  // Inline plugin — vẽ nhãn kg trực tiếp lên canvas, không cần CDN
  const barLabelPlugin = {
    id: 'barLabels',
    afterDatasetsDraw(chart) {
      const ctx = chart.ctx;
      chart.data.datasets.forEach((ds, di) => {
        const meta = chart.getDatasetMeta(di);
        if (meta.hidden) return;
        meta.data.forEach((el, i) => {
          const v = ds.data[i];
          if (v == null || v <= 0) return;
          ctx.save();
          ctx.textAlign = 'center';
          ctx.shadowColor = 'rgba(0,0,0,0.75)';
          ctx.shadowBlur = 4;
          if (di < 2) {
            // Cột kg — hiện số bên trong, sát đỉnh cột
            ctx.fillStyle = '#ffffff';
            ctx.font = 'bold 11px Inter, sans-serif';
            ctx.textBaseline = 'top';
            ctx.fillText(Utils.formatNumber(v), el.x, el.y + 6);
          } else {
            // Đường % — hiện phía trên điểm
            ctx.fillStyle = '#f59e0b';
            ctx.font = 'bold 10px Inter, sans-serif';
            ctx.textBaseline = 'bottom';
            ctx.fillText(v.toFixed(2) + '%', el.x, el.y - 5);
          }
          ctx.restore();
        });
      });
    },
  };

  makeChart('chart-trend', {
    type: 'bar',
    plugins: [barLabelPlugin],
    data: {
      labels,
      datasets: [
        {
          type: 'bar', label: 'Nhập (kg)',
          data: impData,
          backgroundColor: 'rgba(59,130,246,0.75)', borderColor: '#3b82f6', borderWidth: 1, borderRadius: 3,
          yAxisID: 'y',
        },
        {
          type: 'bar', label: 'Xuất (kg)',
          data: expData,
          backgroundColor: 'rgba(34,197,94,0.75)', borderColor: '#22c55e', borderWidth: 1, borderRadius: 3,
          yAxisID: 'y',
        },
        {
          type: 'line', label: 'Thực Xuất/FC (%)',
          data: ratioData,
          borderColor: '#f59e0b', backgroundColor: 'rgba(245,158,11,0.1)',
          borderWidth: 2, pointRadius: 4, pointHoverRadius: 6,
          tension: 0.3, fill: false,
          yAxisID: 'y2',
        },
      ],
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: { labels: { color: '#94a3b8', font: { size: 12 }, boxWidth: 12 } },
        tooltip: {
          callbacks: {
            label: ctx => ctx.datasetIndex < 2
              ? ctx.dataset.label + ': ' + Utils.formatNumber(ctx.raw) + ' kg'
              : ctx.dataset.label + ': ' + (ctx.raw != null ? ctx.raw.toFixed(2) : '—') + '%',
          },
        },
      },
      scales: {
        x: { ticks: { color: '#64748b' }, grid: { color: 'rgba(51,65,85,0.4)' } },
        y: {
          position: 'left',
          ticks: { color: '#64748b', callback: v => v >= 1000 ? (v / 1000).toFixed(0) + 'k' : v },
          grid: { color: 'rgba(51,65,85,0.4)' },
        },
        y2: {
          position: 'right',
          ticks: { color: '#f59e0b', callback: v => v + '%' },
          grid: { drawOnChartArea: false },
        },
      },
    },
  });
}

function initCateChart() {
  const specMap = buildSpecMap();
  const selDate = dashboardDate || (() => { const n = new Date(); return `${n.getFullYear()}-${String(n.getMonth()+1).padStart(2,'0')}-${String(n.getDate()).padStart(2,'0')}`; })();

  const cate3Map = new Map();
  MockData.products.forEach(p => {
    if (!p.cate3) return;
    [p.barcode, p.id].filter(Boolean).forEach(k => { if (!cate3Map.has(k)) cate3Map.set(k, p.cate3); });
  });

  const dt  = new Date(selDate + 'T00:00:00');
  const dow = dt.getDay() === 0 ? 6 : dt.getDay() - 1;
  const mon = new Date(dt); mon.setDate(dt.getDate() - dow);
  const sun = new Date(mon); sun.setDate(mon.getDate() + 6);
  const fmt = d => `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
  const [wStart, wEnd] = [fmt(mon), fmt(sun)];
  const weekNum  = Utils.getWeekNumber(dt);
  const monthPfx = selDate.slice(0, 7); // "YYYY-MM"

  const allImports = MockData.getImports().filter(t => t.status !== 'cancelled');
  const txns = catePieMode === 'week'
    ? allImports.filter(t => t.date && t.date.slice(0,10) >= wStart && t.date.slice(0,10) <= wEnd)
    : catePieMode === 'month'
      ? allImports.filter(t => t.date && t.date.startsWith(monthPfx))
      : allImports.filter(t => t.date && t.date.startsWith(selDate));

  const subtitle = catePieMode === 'week' ? 'W' + weekNum
    : catePieMode === 'month' ? monthPfx.replace('-', '/') : '';

  const PIE_COLORS = ['#8b5cf6','#22c55e','#ef4444','#3b82f6','#f97316','#06b6d4','#f59e0b','#ec4899','#14b8a6','#a855f7','#84cc16','#64748b'];

  const g = {};
  txns.forEach(t => {
    const cat = cate3Map.get(t.barcode) || cate3Map.get(t.productId) || 'Khác';
    g[cat] = (g[cat] || 0) + txnWeight(t, specMap);
  });
  const entries = Object.entries(g).sort((a, b) => b[1] - a[1]);
  if (!entries.length) return;

  const labels   = entries.map(([k]) => k);
  const vals     = entries.map(([, v]) => Math.round(v));
  const total    = vals.reduce((s, v) => s + v, 0);
  const totalTan = (total / 1000).toFixed(1);

  const centerPlugin = {
    id: 'center_cate',
    beforeDraw(chart) {
      const { ctx, chartArea: { left, top, width, height } } = chart;
      const cx = left + width / 2, cy = top + height / 2;
      ctx.save();
      ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
      ctx.fillStyle = '#ffffff';
      ctx.font = 'bold 38px Inter, sans-serif';
      ctx.fillText(totalTan, cx, cy - 10);
      ctx.font = 'bold 14px Inter, sans-serif';
      ctx.fillStyle = '#94a3b8';
      ctx.fillText('TẤN', cx, cy + 22);
      if (subtitle) {
        ctx.font = 'bold 12px Inter, sans-serif';
        ctx.fillStyle = '#f59e0b';
        ctx.fillText(subtitle, cx, cy + 42);
      }
      ctx.restore();
    },
  };

  makeChart('chart-cate-pie', {
    type: 'doughnut',
    plugins: [centerPlugin],
    data: {
      labels,
      datasets: [{
        data: vals,
        backgroundColor: PIE_COLORS.slice(0, labels.length),
        borderColor: '#1e293b', borderWidth: 2, hoverOffset: 10,
      }],
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      cutout: '62%',
      layout: { padding: { right: 4 } },
      plugins: {
        legend: {
          position: 'right',
          align: 'center',
          labels: {
            color: '#ffffff',
            font: { size: 12, weight: '700' },
            boxWidth: 14, boxHeight: 14,
            padding: 10,
            generateLabels: chart => {
              const data = chart.data.datasets[0].data;
              const tot  = data.reduce((s, v) => s + v, 0);
              return chart.data.labels.map((label, i) => ({
                text: label + '  ' + (tot > 0 ? ((data[i] / tot) * 100).toFixed(1) : 0) + '%',
                fillStyle: PIE_COLORS[i] || '#64748b',
                strokeStyle: 'transparent', lineWidth: 0, index: i,
                fontColor: '#ffffff', color: '#ffffff',
              }));
            },
          },
        },
        tooltip: {
          callbacks: {
            label: ctx => {
              const pct = total > 0 ? ((ctx.raw / total) * 100).toFixed(1) : 0;
              return ctx.label + ': ' + (ctx.raw / 1000).toFixed(2) + ' tấn (' + pct + '%)';
            },
          },
        },
        datalabels: { display: false },
      },
    },
  });
}

// ============================================================
// IMPORTS
// ============================================================
function countPhieu(transactions) {
  return new Set(transactions.map(t => t.maPO || t.id)).size;
}

// Returns sorted list of codes present in items but missing from masterData
function getMissingCodes(items, getCode) {
  if (!MockData.masterData.length) return [];
  const missing = new Set();
  items.forEach(item => {
    const code = String(getCode(item) || '').trim();
    if (code && mdLookup(code).found === false) missing.add(code);
  });
  return [...missing].sort();
}

// Red alert banner listing missing codes with link to Mã hàng page
function missingCodesBanner(codes) {
  if (!codes.length) return '';
  const chips = codes.map(c =>
    `<span style="background:rgba(239,68,68,0.15);color:#f87171;border:1px solid rgba(239,68,68,0.35);
      border-radius:4px;padding:2px 8px;font-size:11px;font-family:monospace;font-weight:600;
      white-space:nowrap;">${c}</span>`
  ).join('');
  return `
    <div style="background:rgba(239,68,68,0.08);border:1px solid rgba(239,68,68,0.4);
      border-radius:var(--radius-md);padding:12px 16px;margin-bottom:var(--space-md);
      display:flex;gap:12px;align-items:flex-start;">
      <span style="font-size:1.3rem;flex-shrink:0;line-height:1.4;">⚠️</span>
      <div style="flex:1;min-width:0;">
        <div style="font-weight:700;color:#ef4444;font-size:var(--fs-sm);margin-bottom:8px;">
          ${codes.length} mã hàng chưa có trong Master Data — trọng lượng chưa tính được
        </div>
        <div style="display:flex;flex-wrap:wrap;gap:6px;margin-bottom:10px;">${chips}</div>
        <button class="btn btn-sm" style="border-color:#ef4444;color:#ef4444;"
          onclick="navigateTo('masterdata')">
          🗂️ Mở trang Mã hàng để thêm →
        </button>
      </div>
    </div>`;
}

function renderImports() {
  const allImports = MockData.getImports();
  let data = applyFilter(allImports, importFilter, 'supplier');
  const paged = Utils.paginate(data, importPage, 20);
  const partners   = [...new Set(allImports.map(t => t.supplier).filter(Boolean))].sort();
  const cate3List  = buildCate3List(allImports);

  const totalPhieu    = countPhieu(allImports);
  const filteredPhieu = countPhieu(data);
  const missingImport = getMissingCodes(allImports, t => t.barcode || '');

  document.getElementById('page-content').innerHTML = `
    <div class="section anim-fade-in-up">
      <div class="section-header">
        <div>
          <div class="section-title">Danh sách phiếu nhập</div>
          <div class="section-subtitle">${Utils.formatNumber(filteredPhieu)} / ${Utils.formatNumber(totalPhieu)} phiếu nhập</div>
        </div>
        <button class="btn btn-primary" onclick="showAddModal('import')">+ Tạo phiếu nhập</button>
      </div>
      ${missingCodesBanner(missingImport)}
      ${filterBar('import', importFilter, partners, 'Nhà cung cấp', null, cate3List)}
      <div class="panel">
        <div class="panel-body" style="padding:0;">${detailedTable(paged.data, 'import')}</div>
        ${pagination(paged, 'import')}
      </div>
    </div>`;
}

// ============================================================
// EXPORTS
// ============================================================
function renderExports() {
  const allExports = MockData.getExports();
  let data = applyFilter(allExports, exportFilter, 'customer');
  const paged = Utils.paginate(data, exportPage, 20);
  const partners   = [...new Set(allExports.map(t => t.customer).filter(Boolean))].sort();
  const chiNhanhs  = [...new Set(allExports.map(t => t.chiNhanh).filter(Boolean))].sort();
  const cate3List  = buildCate3List(allExports);

  const totalPhieu    = countPhieu(allExports);
  const filteredPhieu = countPhieu(data);
  const missingExport = getMissingCodes(allExports, t => t.barcode || '');

  document.getElementById('page-content').innerHTML = `
    <div class="section anim-fade-in-up">
      <div class="section-header">
        <div>
          <div class="section-title">Danh sách phiếu chuyển hàng</div>
          <div class="section-subtitle">${Utils.formatNumber(filteredPhieu)} / ${Utils.formatNumber(totalPhieu)} phiếu xuất</div>
        </div>
        <button class="btn btn-primary" onclick="showAddModal('export')">+ Tạo phiếu xuất</button>
      </div>
      ${missingCodesBanner(missingExport)}
      ${filterBar('export', exportFilter, partners, 'Chi nhánh nhận', chiNhanhs, cate3List)}
      <div class="panel">
        <div class="panel-body" style="padding:0;">${detailedTable(paged.data, 'export')}</div>
        ${pagination(paged, 'export')}
      </div>
    </div>`;
}

// ============================================================
// BOOKING
// ============================================================
function renderBookings() {
  const allBookings = MockData.getBookings();

  // Apply filters
  let data = allBookings;
  const f = bookingFilter;
  if (f.search)       { const q = f.search.toLowerCase(); data = data.filter(b => [b.itemName, b.po, b.supplierName, b.itemId].some(v => v && v.toLowerCase().includes(q))); }
  if (f.supplier !== 'all') data = data.filter(b => b.supplierName === f.supplier);
  if (f.store    !== 'all') data = data.filter(b => b.storeName    === f.store);
  if (f.tempCategory !== 'all') data = data.filter(b => b.tempCategory === f.tempCategory);
  if (f.dateFrom)     data = data.filter(b => b.deliveryDate >= f.dateFrom);
  if (f.dateTo)       data = data.filter(b => b.deliveryDate <= f.dateTo + 'T23:59:59');

  const paged     = Utils.paginate(data, bookingPage, 20);
  const suppliers = [...new Set(allBookings.map(b => b.supplierName).filter(Boolean))].sort();
  const stores    = [...new Set(allBookings.map(b => b.storeName).filter(Boolean))].sort();
  const tempCats  = [...new Set(allBookings.map(b => b.tempCategory).filter(Boolean))].sort();
  const uniquePOs = new Set(allBookings.map(b => b.po).filter(Boolean)).size;
  const dates     = [...new Set(allBookings.map(b => b.deliveryDate ? b.deliveryDate.split('T')[0] : '').filter(Boolean))].sort();
  const fmtDate   = d => d ? new Date(d).toLocaleDateString('vi-VN', { day: '2-digit', month: '2-digit', year: 'numeric' }) : '—';
  const missingBk = getMissingCodes(allBookings, b => b.codeSp || '');

  document.getElementById('page-content').innerHTML = `
    <div class="section anim-fade-in-up">
      <div class="section-header">
        <div>
          <div class="section-title">Danh sách Booking</div>
          <div class="section-subtitle">${Utils.formatNumber(data.length)} / ${Utils.formatNumber(allBookings.length)} dòng đặt hàng</div>
        </div>
      </div>

      <!-- Stats -->
      <div class="grid-4" style="margin-bottom:var(--space-md);">
        ${bkStat('📦', 'Tổng số PO',    Utils.formatNumber(uniquePOs), Utils.formatNumber(allBookings.length) + ' dòng đặt')}
        ${bkStat('🏭', 'Nhà cung cấp',  Utils.formatNumber(new Set(allBookings.map(b => b.supplierName).filter(Boolean)).size), 'nhà cung cấp')}
        ${bkStat('🏪', 'Cửa hàng',      Utils.formatNumber(new Set(allBookings.map(b => b.storeName).filter(Boolean)).size), 'điểm nhận hàng')}
        ${bkStat('📅', 'Ngày giao', dates.length ? fmtDate(dates[0]) : '—', dates.length > 1 ? '→ ' + fmtDate(dates[dates.length - 1]) : '')}
      </div>

      ${missingCodesBanner(missingBk)}

      <!-- Filter bar -->
      <div class="filter-bar" style="flex-direction:column;align-items:stretch;gap:8px;padding:12px 16px;margin-bottom:var(--space-md);">
        <div style="display:flex;flex-wrap:wrap;gap:8px;align-items:center;">
          <div class="filter-search-wrap" style="flex:1;min-width:220px;">
            <span style="position:absolute;left:12px;top:50%;transform:translateY(-50%);color:var(--text-muted);pointer-events:none;">🔍</span>
            <input id="bk-search" class="filter-input" type="text"
              placeholder="Tên hàng, PO, nhà cung cấp... (Enter để tìm)"
              value="${f.search}"
              onkeydown="if(event.key==='Enter') setBookingFilter('search',this.value)"
              style="padding-left:36px;width:100%;">
          </div>
          ${searchableSelect('bk-supplier', suppliers, f.supplier, "setBookingFilter('supplier',v)", 'Tất cả NCC', 200)}
          ${stores.length > 1 ? searchableSelect('bk-store', stores, f.store, "setBookingFilter('store',v)", 'Tất cả cửa hàng', 220) : ''}
          ${tempCats.length  ? searchableSelect('bk-cat', tempCats, f.tempCategory, "setBookingFilter('tempCategory',v)", 'Tất cả loại', 150) : ''}
          <input class="filter-input" type="date" title="Từ ngày" style="width:140px;"
            value="${f.dateFrom}" onchange="setBookingFilter('dateFrom',this.value)">
          <span style="color:var(--text-muted);font-size:var(--fs-xs);">→</span>
          <input class="filter-input" type="date" title="Đến ngày" style="width:140px;"
            value="${f.dateTo}" onchange="setBookingFilter('dateTo',this.value)">
          <button class="btn btn-sm" onclick="resetBookingFilter()"
            style="${(f.search || f.supplier !== 'all' || f.store !== 'all' || f.tempCategory !== 'all' || f.dateFrom || f.dateTo) ? 'color:var(--accent-amber);border-color:var(--accent-amber);' : ''}">
            ↺ Xóa lọc
          </button>
        </div>
      </div>

      <!-- Table -->
      <div class="panel">
        <div class="panel-body" style="padding:0;">
          ${allBookings.length === 0
            ? `<div style="text-align:center;padding:60px;color:var(--text-muted);">
                <div style="font-size:3rem;margin-bottom:12px;">📋</div>
                <div style="font-size:var(--fs-md);margin-bottom:8px;">Chưa có dữ liệu Booking</div>
                <div style="font-size:var(--fs-sm);">Import file <b>MD 1_KFM_BOOKING_*.xlsx</b> qua menu Import Excel</div>
               </div>`
            : bookingTable(paged.data)}
        </div>
        ${pagination(paged, 'booking')}
      </div>
    </div>`;

}

function bkStat(icon, label, value, sub) {
  return `
    <div class="panel" style="padding:16px;text-align:center;">
      <div style="font-size:1.4rem;margin-bottom:4px;">${icon}</div>
      <div style="font-size:var(--fs-xs);color:var(--text-muted);margin-bottom:2px;">${label}</div>
      <div style="font-size:var(--fs-xl);font-weight:800;color:var(--text-heading);">${value}</div>
      <div style="font-size:var(--fs-xs);color:var(--text-muted);margin-top:2px;">${sub}</div>
    </div>`;
}

function bookingTable(rows) {
  if (!rows.length) return '<div style="text-align:center;padding:40px;color:var(--text-muted);">Không có kết quả phù hợp</div>';

  const tempColor = { CHILL: '#3b82f6', FROZEN: '#06b6d4', DRY: '#f59e0b', AMBIENT: '#64748b' };

  const trs = rows.map(b => {
    const date     = b.deliveryDate ? new Date(b.deliveryDate).toLocaleDateString('vi-VN', { day: '2-digit', month: '2-digit', year: 'numeric' }) : '—';
    const tc       = b.tempCategory || '';
    const tcColor  = tempColor[tc.toUpperCase()] || 'var(--text-muted)';
    const poShort  = b.po.length > 20 ? b.po.slice(-12) + '…' : b.po;
    const nameShort = b.itemName.length > 40 ? b.itemName.slice(0, 38) + '…' : b.itemName;
    const storeShort = (b.storeName || '').length > 30 ? (b.storeName || '').slice(0, 28) + '…' : (b.storeName || '—');
    return `<tr style="border-bottom:1px solid var(--border-light);">
      <td style="padding:8px 12px;font-size:var(--fs-xs);color:var(--text-muted);white-space:nowrap;">${date}</td>
      <td style="padding:8px 12px;font-size:var(--fs-xs);color:var(--text-secondary);" title="${b.po}">${poShort}</td>
      <td style="padding:8px 12px;font-size:var(--fs-xs);color:var(--text-primary);">${b.supplierName || '—'}</td>
      <td style="padding:8px 12px;font-size:var(--fs-xs);color:var(--text-muted);">${b.itemId || '—'}</td>
      <td style="padding:8px 12px;font-size:var(--fs-sm);color:var(--text-primary);max-width:260px;" title="${b.itemName}">${nameShort}</td>
      <td style="padding:8px 12px;font-size:var(--fs-sm);font-weight:600;color:var(--text-heading);text-align:right;">${Utils.formatNumber(b.quantity)}</td>
      <td style="padding:8px 12px;font-size:var(--fs-xs);color:var(--text-muted);">${b.uom || '—'}</td>
      <td style="padding:8px 12px;font-size:var(--fs-xs);color:var(--text-secondary);" title="${b.storeName}">${storeShort}</td>
      <td style="padding:8px 12px;">
        ${tc ? `<span style="padding:2px 8px;border-radius:20px;font-size:var(--fs-xs);font-weight:600;background:${tcColor}22;color:${tcColor};">${tc}</span>` : '—'}
      </td>
      <td style="padding:8px 12px;font-size:var(--fs-xs);color:var(--text-muted);max-width:180px;" title="${b.notes}">${(b.notes || '').length > 30 ? b.notes.slice(0, 28) + '…' : (b.notes || '—')}</td>
    </tr>`;
  }).join('');

  return `
    <div style="overflow-x:auto;">
      <table style="width:100%;border-collapse:collapse;min-width:900px;">
        <thead>
          <tr style="border-bottom:2px solid var(--border);">
            <th style="padding:10px 12px;text-align:left;font-size:var(--fs-xs);color:var(--text-muted);font-weight:500;white-space:nowrap;">Ngày giao</th>
            <th style="padding:10px 12px;text-align:left;font-size:var(--fs-xs);color:var(--text-muted);font-weight:500;">PO</th>
            <th style="padding:10px 12px;text-align:left;font-size:var(--fs-xs);color:var(--text-muted);font-weight:500;">Nhà cung cấp</th>
            <th style="padding:10px 12px;text-align:left;font-size:var(--fs-xs);color:var(--text-muted);font-weight:500;">Mã hàng</th>
            <th style="padding:10px 12px;text-align:left;font-size:var(--fs-xs);color:var(--text-muted);font-weight:500;">Tên hàng</th>
            <th style="padding:10px 12px;text-align:right;font-size:var(--fs-xs);color:var(--text-muted);font-weight:500;">SL</th>
            <th style="padding:10px 12px;text-align:left;font-size:var(--fs-xs);color:var(--text-muted);font-weight:500;">ĐVT</th>
            <th style="padding:10px 12px;text-align:left;font-size:var(--fs-xs);color:var(--text-muted);font-weight:500;">Cửa hàng</th>
            <th style="padding:10px 12px;text-align:left;font-size:var(--fs-xs);color:var(--text-muted);font-weight:500;">Loại</th>
            <th style="padding:10px 12px;text-align:left;font-size:var(--fs-xs);color:var(--text-muted);font-weight:500;">Ghi chú</th>
          </tr>
        </thead>
        <tbody>${trs}</tbody>
      </table>
    </div>`;
}

// ============================================================
// FC — Forecast Control
// ============================================================
function renderFC() {
  const fc = MockData.getFC();

  if (!fc) {
    document.getElementById('page-content').innerHTML = `
      <div class="empty-state">
        <div class="empty-icon">📊</div>
        <div class="empty-text">Chưa có dữ liệu FC</div>
        <div class="empty-sub">Import file FC Excel qua mục Import Excel → tab FC (Forecast)</div>
        <button class="btn btn-primary" onclick="showImportExcelModal();setTimeout(()=>switchImportTab('fc'),120)">
          + Import FC
        </button>
      </div>`;
    return;
  }

  const MONTHS    = ['T1','T2','T3','T4','T5','T6','T7','T8','T9','T10','T11','T12'];
  const CAT_COLORS = ['#ef4444','#f97316','#f59e0b','#3b82f6','#06b6d4','#8b5cf6','#22c55e','#a855f7'];
  const _fcNow    = new Date();
  const today     = `${_fcNow.getFullYear()}-${String(_fcNow.getMonth()+1).padStart(2,'0')}-${String(_fcNow.getDate()).padStart(2,'0')}`;
  const shortName = cat => cat.replace(/^3\./, '').replace(/ AND /g, '/').slice(0, 20);

  // Compute monthly totals from daily values
  const monthlyTotals = fc.categories.map(() => new Array(12).fill(0));
  fc.dates.forEach((date, di) => {
    const m = parseInt(date.slice(5, 7)) - 1;
    fc.categories.forEach((_, ci) => {
      monthlyTotals[ci][m] += fc.values[ci][di];
    });
  });

  // Today's total — imported FINAL row daily value
  const todayIdx   = fc.dates.indexOf(today);
  const todayTotal = todayIdx >= 0 ? Math.round(fc.totalDailyValues?.[todayIdx] ?? 0) : null;

  // Current month total — imported FINAL row
  const curMonth      = new Date().getMonth();
  const curMonthTotal = Math.round(fc.totalMonthlyAvg?.[curMonth] ?? 0);

  // Year range display
  const yearFrom = fc.dates[0] || '';
  const yearTo   = fc.dates[fc.dates.length - 1] || '';

  // ── Monthly summary table ─────────────────────────────────
  const summaryTable = `
    <div class="panel">
      <div class="panel-header">
        <div class="panel-title">Tổng dự báo theo tháng (kg)</div>
        <div style="font-size:var(--fs-xs);color:var(--text-muted);">${fc.fileName || ''}</div>
      </div>
      <div class="panel-body" style="overflow-x:auto;">
        <table class="data-table">
          <thead>
            <tr>
              <th>Ngành hàng</th>
              ${MONTHS.map(m => `<th style="text-align:right;">${m}</th>`).join('')}
              <th style="text-align:right;background:rgba(255,255,255,0.03);">Cả năm</th>
            </tr>
          </thead>
          <tbody>
            ${fc.categories.map((cat, ci) => {
              const totals   = monthlyTotals[ci];
              const yearTot  = totals.reduce((s, v) => s + v, 0);
              return `<tr>
                <td>
                  <span style="display:inline-block;width:9px;height:9px;border-radius:50%;
                    background:${CAT_COLORS[ci % CAT_COLORS.length]};margin-right:6px;vertical-align:middle;"></span>
                  ${shortName(cat)}
                </td>
                ${totals.map((v, mi) => {
                  const isActive = fcMonth === mi + 1;
                  return `<td style="text-align:right;${isActive ? 'color:var(--accent-blue);font-weight:600;' : ''}">
                    ${Utils.formatNumber(Math.round(v))}
                  </td>`;
                }).join('')}
                <td style="text-align:right;font-weight:600;background:rgba(255,255,255,0.03);">
                  ${Utils.formatNumber(Math.round(yearTot))}
                </td>
              </tr>`;
            }).join('')}
            <tr style="border-top:2px solid var(--border);font-weight:700;">
              <td>TỔNG</td>
              ${MONTHS.map((_, mi) => {
                const t = Math.round(fc.totalMonthlyAvg?.[mi] ?? 0);
                const isActive = fcMonth === mi + 1;
                return `<td style="text-align:right;${isActive ? 'color:var(--accent-blue);' : ''}">
                  ${Utils.formatNumber(t)}
                </td>`;
              }).join('')}
              <td style="text-align:right;background:rgba(255,255,255,0.03);">
                ${Utils.formatNumber(Math.round((fc.totalMonthlyAvg || []).reduce((s, v) => s + v, 0)))}
              </td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>`;

  // ── Daily detail table (when a month is selected) ─────────
  let dailyTable = '';
  if (fcMonth > 0) {
    const monthDates = fc.dates
      .map((d, i) => ({ d, i }))
      .filter(({ d }) => parseInt(d.slice(5, 7)) === fcMonth);

    const dayNames = ['CN','T2','T3','T4','T5','T6','T7'];
    dailyTable = `
      <div class="panel">
        <div class="panel-header">
          <div class="panel-title">Chi tiết ngày — Tháng ${fcMonth} (${monthDates.length} ngày)</div>
        </div>
        <div class="panel-body" style="overflow-x:auto;">
          <table class="data-table">
            <thead>
              <tr>
                <th>Ngày</th>
                <th>Thứ</th>
                ${fc.categories.map((cat, ci) =>
                  `<th style="text-align:right;color:${CAT_COLORS[ci % CAT_COLORS.length]};">${shortName(cat)}</th>`
                ).join('')}
                <th style="text-align:right;">Tổng</th>
              </tr>
            </thead>
            <tbody>
              ${monthDates.map(({ d, i }) => {
                const dn      = dayNames[new Date(d + 'T00:00:00').getDay()];
                const vals    = fc.categories.map((_, ci) => fc.values[ci][i]);
                const total   = Math.round(fc.totalDailyValues?.[i] ?? 0);
                const isToday = d === today;
                return `<tr${isToday ? ' style="background:rgba(99,102,241,0.1);"' : ''}>
                  <td style="${isToday ? 'font-weight:700;color:var(--accent-blue);' : ''}">
                    ${d.slice(8, 10)}/${d.slice(5, 7)}${isToday ? ' ◀ Hôm nay' : ''}
                  </td>
                  <td style="color:var(--text-muted);">${dn}</td>
                  ${vals.map(v => `<td style="text-align:right;">${Utils.formatNumber(Math.round(v))}</td>`).join('')}
                  <td style="text-align:right;font-weight:600;">${Utils.formatNumber(total)}</td>
                </tr>`;
              }).join('')}
            </tbody>
          </table>
        </div>
      </div>`;
  }

  // ── Chart datasets ────────────────────────────────────────
  const datasets = fc.categories.map((cat, ci) => ({
    label:           shortName(cat),
    data:            monthlyTotals[ci].map(v => Math.round(v)),
    backgroundColor: CAT_COLORS[ci % CAT_COLORS.length] + 'bb',
    borderColor:     CAT_COLORS[ci % CAT_COLORS.length],
    borderWidth:     1,
  }));

  // ── Render ────────────────────────────────────────────────
  document.getElementById('page-content').innerHTML = `
    <div style="display:grid;gap:var(--space-lg);">
      <div class="stats-grid">
        ${fcStat('📅', 'Kỳ dự báo', yearFrom.slice(0,7) + ' → ' + yearTo.slice(0,7), fc.dates.length + ' ngày')}
        ${fcStat('📦', 'Ngành hàng', fc.categories.length, 'danh mục')}
        ${fcStat('📊', 'Tháng ' + MONTHS[curMonth], Utils.formatNumber(Math.round(curMonthTotal)), 'kg dự báo')}
        ${fcStat('🗓️', 'Hôm nay (' + (today.slice(8,10) + '/' + today.slice(5,7)) + ')',
          todayTotal !== null ? Utils.formatNumber(Math.round(todayTotal)) : '—',
          todayTotal !== null ? 'kg dự báo' : 'ngoài kỳ FC')}
      </div>

      <div class="filter-bar">
        <span style="color:var(--text-muted);font-size:var(--fs-sm);">Xem chi tiết:</span>
        ${[0,...Array.from({length:12},(_,i)=>i+1)].map(m => {
          const label  = m === 0 ? 'Cả năm' : MONTHS[m - 1];
          const active = fcMonth === m;
          return `<button class="btn btn-sm${active ? ' btn-primary' : ''}"
            onclick="fcMonth=${m};renderFC();">${label}</button>`;
        }).join('')}
      </div>

      <div class="panel">
        <div class="panel-header">
          <div class="panel-title">Biểu đồ dự báo hàng tháng (kg)</div>
        </div>
        <div class="panel-body">
          <canvas id="fc-chart" height="110"></canvas>
        </div>
      </div>

      ${summaryTable}
      ${dailyTable}
    </div>`;

  // Render stacked bar chart
  const ctx = document.getElementById('fc-chart');
  if (ctx) {
    activeCharts.forEach(c => { try { c.destroy(); } catch (_) {} });
    activeCharts = [];
    activeCharts.push(new Chart(ctx, {
      type: 'bar',
      data: { labels: MONTHS, datasets },
      options: {
        responsive: true,
        plugins: {
          legend: { position: 'bottom', labels: { color: '#94a3b8', font: { size: 11 }, boxWidth: 12 } },
          tooltip: {
            callbacks: {
              label: ctx => ` ${ctx.dataset.label}: ${Utils.formatNumber(ctx.parsed.y)} kg`,
            },
          },
        },
        scales: {
          x: { stacked: true, grid: { color: 'rgba(255,255,255,0.05)' }, ticks: { color: '#94a3b8' } },
          y: { stacked: true, grid: { color: 'rgba(255,255,255,0.05)' }, ticks: { color: '#94a3b8',
            callback: v => Utils.formatNumber(v) } },
        },
      },
    }));
  }
}

function fcStat(icon, label, value, sub) {
  return `
    <div class="stat-card">
      <div class="stat-header">
        <span class="stat-icon">${icon}</span>
        <span class="stat-label">${label}</span>
      </div>
      <div class="stat-value">${value}</div>
      <div class="stat-sub">${sub}</div>
    </div>`;
}

// ── Filter helpers ────────────────────────────────────────
function buildCate3List(transactions) {
  // Only map products that have cate3; first-write wins (price list products take priority)
  const cate3Map = new Map();
  MockData.products.forEach(p => {
    if (!p.cate3) return;
    [p.barcode, p.id].filter(Boolean).forEach(k => {
      if (!cate3Map.has(k)) cate3Map.set(k, p.cate3);
    });
  });
  return [...new Set(
    transactions
      .map(t => cate3Map.get(t.barcode) || cate3Map.get(t.productId) || '')
      .filter(Boolean)
  )].sort();
}

function applyFilter(data, f, partnerField) {
  let result = data;

  if (f.search) {
    const q = f.search.toLowerCase();
    result = result.filter(t =>
      (t.maPO          || '').toLowerCase().includes(q) ||
      (t.barcode       || '').toLowerCase().includes(q) ||
      (t.productName   || '').toLowerCase().includes(q) ||
      (t[partnerField] || '').toLowerCase().includes(q) ||
      (t.chiNhanh      || '').toLowerCase().includes(q)
    );
  }
  if (f.statuses && f.statuses.length) {
    result = result.filter(t => f.statuses.includes(t.status));
  }
  if (f.partner !== 'all') {
    result = result.filter(t => (t[partnerField] || '') === f.partner);
  }
  if (f.chiNhanh && f.chiNhanh !== 'all') {
    result = result.filter(t => (t.chiNhanh || '') === f.chiNhanh);
  }
  if (f.category && f.category !== 'all') {
    const keySet = new Set();
    MockData.products.filter(p => p.cate3 === f.category).forEach(p => {
      if (p.barcode) keySet.add(p.barcode);
      if (p.id)     keySet.add(p.id);
    });
    result = result.filter(t => keySet.has(t.barcode) || keySet.has(t.productId));
  }
  if (f.dateFrom) {
    const from = new Date(f.dateFrom);
    result = result.filter(t => new Date(t.date) >= from);
  }
  if (f.dateTo) {
    const to = new Date(f.dateTo);
    to.setHours(23, 59, 59);
    result = result.filter(t => new Date(t.date) <= to);
  }
  return result;
}

// ── Searchable dropdown ───────────────────────────────────
function searchableSelect(id, options, current, onSelectExpr, allLabel, minWidth = 180) {
  const currentLabel = current === 'all' || !current
    ? allLabel
    : (options.find(o => o === current) || allLabel);
  const isActive = current && current !== 'all';

  const items = ['<li onclick="closeSS(\'' + id + '\');' + onSelectExpr.replace('v', "'all'") + '" style="padding:8px 12px;cursor:pointer;color:var(--text-muted);">' + allLabel + '</li>',
    ...options.map(o => {
      const sel = o === current ? 'font-weight:600;color:var(--accent-blue);' : '';
      return `<li onclick="closeSS('${id}');${onSelectExpr.replace('v', "'" + o.replace(/'/g, "\\'") + "'")}"
        style="padding:8px 12px;cursor:pointer;${sel}"
        data-ss-item="${o.toLowerCase()}">${o}</li>`;
    })
  ].join('');

  return `
  <div id="${id}" style="position:relative;display:inline-block;min-width:${minWidth}px;">
    <button onclick="toggleSS('${id}')"
      style="width:100%;padding:6px 10px;border-radius:6px;cursor:pointer;text-align:left;
             display:flex;align-items:center;justify-content:space-between;gap:8px;
             border:1px solid ${isActive ? 'var(--accent-blue)' : 'var(--border)'};
             background:var(--bg-card);color:${isActive ? 'var(--accent-blue)' : 'var(--text-secondary)'};
             font-size:var(--fs-sm);white-space:nowrap;overflow:hidden;">
      <span style="overflow:hidden;text-overflow:ellipsis;">${currentLabel}</span>
      <span style="flex-shrink:0;font-size:10px;">▾</span>
    </button>
    <div id="${id}-drop"
      style="display:none;position:absolute;top:calc(100% + 4px);left:0;z-index:999;min-width:${minWidth}px;max-width:360px;
             background:var(--bg-card);border:1px solid var(--border);border-radius:8px;
             box-shadow:0 8px 24px rgba(0,0,0,.4);overflow:hidden;">
      <div style="padding:8px;">
        <input type="text" placeholder="Tìm kiếm..."
          oninput="filterSS('${id}',this.value)"
          onclick="event.stopPropagation()"
          style="width:100%;padding:6px 10px;border-radius:6px;border:1px solid var(--border);
                 background:var(--bg-sidebar);color:var(--text-primary);font-size:var(--fs-sm);outline:none;box-sizing:border-box;">
      </div>
      <ul id="${id}-list" style="max-height:240px;overflow-y:auto;margin:0;padding:0 0 4px;list-style:none;">${items}</ul>
    </div>
  </div>`;
}

function toggleSS(id) {
  document.querySelectorAll('[id$="-drop"]').forEach(d => {
    if (d.id !== id + '-drop') d.style.display = 'none';
  });
  const drop = document.getElementById(id + '-drop');
  if (!drop) return;
  const isOpen = drop.style.display !== 'none';
  drop.style.display = isOpen ? 'none' : 'block';
  if (!isOpen) {
    const inp = drop.querySelector('input');
    if (inp) { inp.value = ''; filterSS(id, ''); requestAnimationFrame(() => inp.focus()); }
  }
}

function closeSS(id) {
  const drop = document.getElementById(id + '-drop');
  if (drop) drop.style.display = 'none';
}

function filterSS(id, query) {
  const list = document.getElementById(id + '-list');
  if (!list) return;
  const q = query.toLowerCase();
  list.querySelectorAll('li[data-ss-item]').forEach(li => {
    li.style.display = li.dataset.ssItem.includes(q) ? '' : 'none';
  });
}

// Close searchable dropdowns when clicking outside
document.addEventListener('click', e => {
  if (!e.target.closest('[id$="-drop"]') && !e.target.closest('button[onclick*="toggleSS"]')) {
    document.querySelectorAll('[id$="-drop"]').forEach(d => { d.style.display = 'none'; });
  }
});

function filterBar(type, f, partners, partnerLabel, extraPartners, cate3List = []) {
  const isExport = type === 'export';

  const STATUS_LIST = [
    { val: 'completed', label: 'Hoàn thành', color: 'var(--accent-green)'  },
    { val: 'pending',   label: isExport ? 'Đang chuyển' : 'Chờ xử lý', color: 'var(--accent-amber)' },
    { val: 'cancelled', label: 'Đã hủy',     color: 'var(--accent-red)'    },
  ];

  const statusChips = STATUS_LIST.map(s => {
    const active = f.statuses.includes(s.val);
    return `<button
      onclick="toggleStatus('${type}','${s.val}')"
      style="
        padding:4px 12px;border-radius:20px;font-size:var(--fs-xs);font-weight:600;cursor:pointer;
        border:1px solid ${s.color};
        background:${active ? s.color : 'transparent'};
        color:${active ? '#fff' : s.color};
        transition:all .15s;
      ">${s.label}</button>`;
  }).join('');



  const activeCount = (f.statuses.length ? 1 : 0)
    + (f.partner !== 'all' ? 1 : 0)
    + ((f.chiNhanh && f.chiNhanh !== 'all') ? 1 : 0)
    + ((f.category && f.category !== 'all') ? 1 : 0)
    + (f.dateFrom ? 1 : 0)
    + (f.dateTo   ? 1 : 0)
    + (f.search   ? 1 : 0);

  return `
    <div class="filter-bar" style="flex-direction:column;align-items:stretch;gap:8px;padding:12px 16px;">

      <!-- Row 1: search + reset -->
      <div style="display:flex;gap:8px;align-items:center;">
        <div class="filter-search-wrap" style="flex:1;">
          <span style="position:absolute;left:12px;top:50%;transform:translateY(-50%);color:var(--text-muted);pointer-events:none;">🔍</span>
          <input id="filter-search-${type}" class="filter-input" type="text"
            placeholder="Tìm mã PO, mã hàng, tên hàng, đối tác... (Enter để tìm)"
            value="${f.search}"
            onkeydown="if(event.key==='Enter') setFilter('${type}','search',this.value)"
            style="padding-left:36px;width:100%;">
        </div>
        <button class="btn btn-sm" onclick="resetFilter('${type}')"
          style="${activeCount ? 'color:var(--accent-amber);border-color:var(--accent-amber);' : ''}">
          ↺ Xóa lọc${activeCount ? ' (' + activeCount + ')' : ''}
        </button>
      </div>

      <!-- Row 2: status chips + category + partner(s) + date -->
      <div style="display:flex;flex-wrap:wrap;gap:8px;align-items:center;">
        <div style="display:flex;gap:6px;align-items:center;">
          <span style="font-size:var(--fs-xs);color:var(--text-muted);white-space:nowrap;">Trạng thái:</span>
          ${statusChips}
        </div>
        <div style="width:1px;height:24px;background:var(--border);margin:0 2px;"></div>
        ${searchableSelect('ss-'+type+'-category', cate3List, f.category, "setFilter('"+type+"','category',v)", 'Tất cả danh mục', 180)}
        ${searchableSelect('ss-'+type+'-partner', partners, f.partner, "setFilter('"+type+"','partner',v)", 'Tất cả ' + partnerLabel, 200)}
        ${isExport ? searchableSelect('ss-'+type+'-chiNhanh', extraPartners||[], f.chiNhanh||'all', "setFilter('"+type+"','chiNhanh',v)", 'Tất cả chi nhánh chuyển', 200) : ''}
        <input class="filter-input" type="date" title="Từ ngày" style="width:140px;"
          value="${f.dateFrom}" onchange="setFilter('${type}','dateFrom',this.value)">
        <span style="color:var(--text-muted);font-size:var(--fs-xs);">→</span>
        <input class="filter-input" type="date" title="Đến ngày" style="width:140px;"
          value="${f.dateTo}" onchange="setFilter('${type}','dateTo',this.value)">
      </div>
    </div>`;
}

function toggleStatus(type, status) {
  const f = type === 'import' ? importFilter : exportFilter;
  const idx = f.statuses.indexOf(status);
  if (idx >= 0) f.statuses.splice(idx, 1);
  else f.statuses.push(status);
  if (type === 'import') { importPage = 1; renderImports(); }
  else                   { exportPage = 1; renderExports(); }
}

function setFilter(type, key, value) {
  const f = type === 'import' ? importFilter : exportFilter;
  f[key] = value;
  if (type === 'import') { importPage = 1; renderImports(); }
  else                   { exportPage = 1; renderExports(); }
}

function resetFilter(type) {
  const f = type === 'import' ? importFilter : exportFilter;
  f.search = ''; f.statuses = []; f.partner = 'all';
  f.chiNhanh = 'all'; f.category = 'all'; f.dateFrom = ''; f.dateTo = '';
  if (type === 'import') { importPage = 1; renderImports(); }
  else                   { exportPage = 1; renderExports(); }
}

function setBookingFilter(key, value) {
  bookingFilter[key] = value;
  bookingPage = 1;
  renderBookings();
}

function resetBookingFilter() {
  bookingFilter.search = ''; bookingFilter.supplier = 'all';
  bookingFilter.store = 'all'; bookingFilter.tempCategory = 'all';
  bookingFilter.dateFrom = ''; bookingFilter.dateTo = '';
  bookingPage = 1;
  renderBookings();
}

// ── Master data lookup: returns spec (quy cách) for a product code ──
function mdLookup(code) {
  if (!code || !MockData.masterData.length) return { spec: null, found: null };
  const c = String(code).trim().toUpperCase();
  const entry = MockData.masterData.find(r => String(r.code || '').trim().toUpperCase() === c);
  if (!entry) return { spec: null, found: false };
  return { spec: parseFloat(entry.spec) || null, found: true };
}

// ── Detailed table (import / export) ─────────────────────
function detailedTable(transactions, type) {
  if (!transactions.length) return `
    <div class="empty-state">
      <div class="empty-icon">📋</div>
      <div class="empty-text">Không có dữ liệu phù hợp</div>
    </div>`;

  const isImport = type === 'import';
  const hasMd = MockData.masterData.length > 0;

  return `
    <table class="data-table">
      <thead>
        <tr>
          <th>Mã PO</th>
          <th>Mã hàng</th>
          <th>Tên hàng</th>
          ${isImport
            ? '<th style="text-align:right;">Số lượng</th>'
            : '<th style="text-align:right;">SL chuyển</th><th style="text-align:right;">SL nhận</th><th style="text-align:right;">Chênh lệch</th>'
          }
          <th>Đơn vị</th>
          <th style="text-align:right;">Trọng lượng</th>
          ${isImport ? '<th style="text-align:right;">Đơn giá</th>' : ''}
          ${isImport ? '<th style="text-align:right;">Thành tiền</th>' : ''}
          <th>${isImport ? 'Nhà cung cấp' : 'Chi nhánh nhận'}</th>
          <th>Trạng thái</th>
          <th>Ngày</th>
        </tr>
      </thead>
      <tbody>
        ${transactions.map(t => {
          const slChuyen  = t.transferQty != null ? t.transferQty : t.quantity;
          const slNhan    = t.actualQty   != null ? t.actualQty   : (t.status === 'completed' ? t.quantity : null);
          const chenhlech = (slNhan != null) ? (slChuyen - slNhan) : null;
          const code      = String(t.barcode || '').trim();
          const { spec, found } = mdLookup(code);
          const qty       = isImport ? (t.quantity || 0) : (slChuyen || 0);
          const weight    = (spec != null) ? qty * spec : null;
          // found===false means code exists but not in MD; found===null means no MD loaded yet
          const mdMissing = hasMd && code && found === false;

          return `
          <tr${mdMissing ? ' class="md-missing-row"' : ''}>
            <td>
              <span style="color:var(--accent-blue);font-weight:600;font-size:var(--fs-xs);">
                ${t.maPO || '—'}
              </span>
            </td>
            <td style="font-size:var(--fs-xs);font-family:monospace;${mdMissing ? 'color:var(--accent-red);' : 'color:var(--text-muted);'}">
              ${t.barcode || t.productId || '—'}
              ${mdMissing ? '<span style="margin-left:3px;" title="Chưa có trong Master Data">⚠</span>' : ''}
            </td>
            <td style="color:var(--text-primary);font-weight:500;max-width:200px;">
              <div style="white-space:nowrap;overflow:hidden;text-overflow:ellipsis;" title="${t.productName}">
                ${t.productName}
              </div>
            </td>
            ${isImport
              ? `<td style="text-align:right;font-weight:600;color:var(--accent-blue-light);">${Utils.formatNumber(t.quantity)}</td>`
              : `<td style="text-align:right;font-weight:600;color:var(--accent-green-light);">${Utils.formatNumber(slChuyen)}</td>
                 <td style="text-align:right;font-weight:600;color:var(--accent-green);">${slNhan != null ? Utils.formatNumber(slNhan) : '—'}</td>
                 <td style="text-align:right;font-weight:600;color:${chenhlech > 0 ? 'var(--accent-amber)' : chenhlech < 0 ? 'var(--accent-red)' : 'var(--text-muted)'};">
                   ${chenhlech != null ? (chenhlech !== 0 ? (chenhlech > 0 ? '+' : '') + Utils.formatNumber(chenhlech) : '—') : '—'}
                 </td>`
            }
            <td style="color:var(--text-muted);">${t.unit || '—'}</td>
            <td style="text-align:right;font-weight:600;${mdMissing ? 'color:var(--accent-red);' : 'color:var(--text-secondary);'}">
              ${weight != null ? Utils.formatNumber(Math.round(weight * 100) / 100) : (hasMd && code ? '<span style="color:var(--accent-red);">—</span>' : '—')}
            </td>
            ${isImport ? `<td style="text-align:right;color:var(--text-secondary);">${t.price ? Utils.formatVND(t.price) : '—'}</td>` : ''}
            ${isImport ? `<td style="text-align:right;color:var(--accent-blue-light);font-weight:600;">${t.total ? Utils.formatVND(t.total) : '—'}</td>` : ''}
            <td style="color:var(--text-muted);font-size:var(--fs-xs);max-width:160px;">
              <div style="white-space:nowrap;overflow:hidden;text-overflow:ellipsis;" title="${t.supplier || t.customer || ''}">
                ${t.supplier || t.customer || '—'}
              </div>
            </td>
            <td>${statusBadge(t.status, t.type)}</td>
            <td style="color:var(--text-muted);font-size:var(--fs-xs);white-space:nowrap;">
              ${Utils.formatDate(t.date)}
            </td>
          </tr>`;
        }).join('')}
      </tbody>
    </table>`;
}

// ============================================================
// INVENTORY
// ============================================================
function renderInventory(search = '') {
  let products = MockData.products;
  if (search) products = Utils.searchFilter(products, search, ['id', 'name']);
  const paged = Utils.paginate(products, inventoryPage, 15);

  document.getElementById('page-content').innerHTML = `
    <div class="section anim-fade-in-up">
      <div class="section-header">
        <div>
          <div class="section-title">Tồn kho hàng hóa</div>
          <div class="section-subtitle">${Utils.formatNumber(products.length)} sản phẩm</div>
        </div>
        <button class="btn btn-sm" onclick="recalcStock()" title="Tính lại tồn kho theo logic mới nhất">
          🔄 Tính lại tồn kho
        </button>
      </div>
      <div class="panel">
        <div class="panel-body" style="padding:0;">
          <table class="data-table">
            <thead>
              <tr>
                <th>Mã SP</th><th>Tên sản phẩm</th><th>Danh mục</th><th>Đơn vị</th>
                <th>Tồn kho</th><th>Tối thiểu</th><th>Đơn giá</th><th>Giá trị kho</th><th>Trạng thái</th>
              </tr>
            </thead>
            <tbody>
              ${paged.data.map(p => {
                const cat = MockData.categories.find(c => c.id === p.category);
                const catLabel = p.cate3
                  ? p.cate3.replace(/^\d+\.\s*/, '').trim()
                  : (cat?.name || p.category);
                const isCritical = p.stock <= p.minStock;
                const isLow = !isCritical && p.stock <= p.minStock * 1.5;
                const stockColor = isCritical ? 'var(--accent-red)' : isLow ? 'var(--accent-amber)' : 'var(--accent-green)';
                return `
                  <tr>
                    <td><span style="color:var(--accent-cyan);font-weight:600;">${p.id}</span></td>
                    <td style="color:var(--text-primary);font-weight:500;">${p.name}</td>
                    <td>
                      <span class="badge" style="background:${cat?.color}22;color:${cat?.color};">${catLabel}</span>
                    </td>
                    <td>${p.unit}</td>
                    <td style="font-weight:700;color:${stockColor};">${Utils.formatNumber(p.stock)}</td>
                    <td style="color:var(--text-muted);">${Utils.formatNumber(p.minStock)}</td>
                    <td>${Utils.formatVND(p.price)}</td>
                    <td>${Utils.formatVND(p.stock * p.price)}</td>
                    <td>${isCritical
                      ? '<span class="badge badge-red">Cạn hàng</span>'
                      : isLow
                      ? '<span class="badge badge-amber">Sắp hết</span>'
                      : '<span class="badge badge-green">Đủ hàng</span>'
                    }</td>
                  </tr>`;
              }).join('')}
            </tbody>
          </table>
        </div>
        ${pagination(paged, 'inventory')}
      </div>
    </div>`;
}

// ============================================================
// MASTER DATA (Mã hàng thịt cá)
// ============================================================
// MASTER DATA — editable grid

function renderMasterData() {
  const data = MockData.getMasterData();
  const el   = document.getElementById('page-content');

  if (!data.length) {
    el.innerHTML = `
      <div class="empty-state">
        <div style="font-size:3rem;margin-bottom:var(--space-md);">🗂️</div>
        <div class="empty-title">Chưa có Master Data</div>
        <div class="empty-sub">Import file "Master data thịt cá.xlsx" để bắt đầu</div>
        <button class="btn btn-primary" onclick="showImportExcelModal();switchImportTab('masterdata');">
          📂 Import Master Data
        </button>
      </div>`;
    return;
  }

  const esc = s => String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');

  const rows = data.map((r, i) => `
    <tr class="md-row">
      <td class="md-rn">${i + 1}</td>
      <td class="md-cell"><input class="md-input md-code" value="${esc(r.code)}"
        data-idx="${i}" data-field="code"
        onchange="mdSaveCell(this)" onkeydown="mdKeyNav(event,${i},0)"></td>
      <td class="md-cell md-cell-name"><input class="md-input" value="${esc(r.name)}"
        data-idx="${i}" data-field="name"
        onchange="mdSaveCell(this)" onkeydown="mdKeyNav(event,${i},1)"></td>
      <td class="md-cell"><input class="md-input" value="${esc(r.unit)}"
        data-idx="${i}" data-field="unit" style="max-width:80px;"
        onchange="mdSaveCell(this)" onkeydown="mdKeyNav(event,${i},2)"></td>
      <td class="md-cell"><input class="md-input" value="${esc(r.spec)}"
        data-idx="${i}" data-field="spec" style="max-width:90px;"
        onchange="mdSaveCell(this)" onkeydown="mdKeyNav(event,${i},3)"></td>
      <td class="md-del"><button class="md-del-btn" onclick="mdDeleteRow(${i})" title="Xóa dòng">×</button></td>
    </tr>`).join('');

  el.innerHTML = `
    <div style="display:flex;flex-direction:column;gap:var(--space-sm);height:calc(100vh - 130px);">
      <div class="md-toolbar">
        <input id="md-search" class="md-search-input" type="text" placeholder="🔍 Tìm mã hàng hoặc tên sản phẩm..."
          oninput="mdSearch(this.value)">
        <button class="btn btn-primary btn-sm" onclick="mdAddRow()">+ Thêm dòng</button>
        <button class="btn btn-ghost btn-sm" onclick="showImportExcelModal();switchImportTab('masterdata');">📂 Import Excel</button>
        <span class="md-count" id="md-count">${Utils.formatNumber(data.length)} mã hàng</span>
      </div>
      <div class="md-wrap" style="flex:1;">
        <table class="md-grid" id="md-table">
          <thead>
            <tr>
              <th class="md-th-rn">#</th>
              <th style="min-width:100px;">Mã hàng</th>
              <th style="min-width:340px;">Tên sản phẩm</th>
              <th style="min-width:70px;">ĐVT</th>
              <th style="min-width:90px;">Quy cách</th>
              <th class="md-th-del"></th>
            </tr>
          </thead>
          <tbody id="md-tbody">
            ${rows}
            <tr class="md-add-row">
              <td class="md-rn"></td>
              <td colspan="5">
                <button class="md-add-btn" onclick="mdAddRow()">+ Thêm dòng mới</button>
              </td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>`;
}

function mdSaveCell(input) {
  const idx   = parseInt(input.dataset.idx);
  const field = input.dataset.field;
  if (!isNaN(idx) && MockData.masterData[idx]) {
    MockData.masterData[idx][field] = input.value;
    MockData.save();
  }
}

function mdDeleteRow(idx) {
  if (!confirm('Xóa dòng này?')) return;
  MockData.masterData.splice(idx, 1);
  MockData.save();
  renderMasterData();
}

function mdAddRow() {
  MockData.masterData.push({ id: 'md-' + Date.now(), code: '', name: '', unit: '', spec: '' });
  MockData.save();
  renderMasterData();
  // Focus mã hàng của dòng mới
  setTimeout(() => {
    const inputs = document.querySelectorAll('#md-tbody .md-input');
    if (inputs.length >= 4) inputs[inputs.length - 4].focus();
  }, 30);
}

function mdSearch(q) {
  const lq   = q.toLowerCase();
  const rows = document.querySelectorAll('#md-tbody .md-row');
  let vis = 0;
  rows.forEach(row => {
    const text = row.textContent.toLowerCase();
    const show = !lq || text.includes(lq);
    row.style.display = show ? '' : 'none';
    if (show) vis++;
  });
  const cnt = document.getElementById('md-count');
  if (cnt) cnt.textContent = Utils.formatNumber(vis) + ' mã hàng';
}

function mdKeyNav(e, rowIdx, colIdx) {
  if (e.key === 'Tab') return; // browser handles Tab naturally
  if (e.key === 'Enter') {
    e.preventDefault();
    // Move to same column, next row
    const inputs = document.querySelectorAll(`#md-tbody .md-input[data-field="${e.target.dataset.field}"]`);
    if (inputs[rowIdx + 1]) inputs[rowIdx + 1].focus();
    else mdAddRow();
  }
}

// ============================================================
// PRICE LIST
// ============================================================
function renderPriceList() {
  const esc = s => String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');

  const pricedProducts = MockData.products.filter(p => p.price > 0);

  const searchVal = (document.getElementById('pl-search') || {}).value || '';
  const catFilter = (document.getElementById('pl-cat') || {}).value || 'all';

  let filtered = pricedProducts;
  if (searchVal.trim()) {
    const q = searchVal.trim().toLowerCase();
    filtered = filtered.filter(p =>
      (p.name   || '').toLowerCase().includes(q) ||
      (p.id     || '').toLowerCase().includes(q) ||
      (p.barcode|| '').toLowerCase().includes(q) ||
      (p.cate1  || '').toLowerCase().includes(q) ||
      (p.cate2  || '').toLowerCase().includes(q) ||
      (p.cate3  || '').toLowerCase().includes(q)
    );
  }
  if (catFilter !== 'all') {
    filtered = filtered.filter(p => p.category === catFilter);
  }

  const PAGE_SIZE = 50;
  const page = window.__plPage || 1;
  const total = filtered.length;
  const totalPages = Math.max(1, Math.ceil(total / PAGE_SIZE));
  const safePage = Math.min(page, totalPages);
  const slice = filtered.slice((safePage - 1) * PAGE_SIZE, safePage * PAGE_SIZE);

  const catOptions = MockData.categories.map(c =>
    `<option value="${c.id}" ${catFilter === c.id ? 'selected' : ''}>${c.name}</option>`
  ).join('');

  const paginationHtml = totalPages <= 1 ? '' : `
    <div class="pagination">
      <button class="btn btn-sm" onclick="window.__plPage=${safePage-1};renderPriceList()" ${safePage<=1?'disabled':''}>‹</button>
      <span class="page-info">Trang ${safePage}/${totalPages}</span>
      <button class="btn btn-sm" onclick="window.__plPage=${safePage+1};renderPriceList()" ${safePage>=totalPages?'disabled':''}>›</button>
    </div>`;

  const isFiltered = searchVal || catFilter !== 'all';
  const emptyState = total === 0 ? `
    <div class="empty-state">
      <div class="empty-icon">💰</div>
      <div class="empty-title">${isFiltered ? 'Không tìm thấy sản phẩm' : 'Chưa có dữ liệu bảng giá'}</div>
      <div class="empty-sub">${isFiltered ? 'Thử thay đổi bộ lọc' : 'Import file bảng giá để xem dữ liệu giá'}</div>
      ${!isFiltered ? '<button class="btn btn-primary" onclick="showImportExcelModal()">Import bảng giá</button>' : ''}
    </div>` : '';

  const clearBtn = pricedProducts.length > 0
    ? `<button class="btn btn-sm" style="color:var(--accent-red);border-color:var(--accent-red);"
         onclick="if(confirm('Xoá toàn bộ ${pricedProducts.length} sản phẩm trong bảng giá?')){MockData.clearProducts();window.__plPage=1;renderPriceList();}">
        Xoá toàn bộ</button>`
    : '';

  document.getElementById('page-content').innerHTML = `
    <div class="section anim-fade-in-up">
      <div class="panel">
        <div class="panel-header">
          <div class="panel-title">Bảng giá (${pricedProducts.length} sản phẩm)</div>
          <div style="display:flex;gap:8px;align-items:center;">
            ${clearBtn}
            <button class="btn btn-primary btn-sm" onclick="showImportExcelModal()">+ Import bảng giá</button>
          </div>
        </div>
        <div class="filter-bar">
          <div class="filter-search-wrap">
            <span class="search-icon">🔍</span>
            <input class="filter-input" id="pl-search" placeholder="Tìm mã hàng, tên hàng, cate..." value="${esc(searchVal)}"
              oninput="window.__plPage=1;renderPriceList()">
          </div>
          <select class="filter-select" id="pl-cat" onchange="window.__plPage=1;renderPriceList()">
            <option value="all" ${catFilter==='all'?'selected':''}>Tất cả danh mục</option>
            ${catOptions}
          </select>
          <button class="btn btn-sm" onclick="window.__plPage=1;document.getElementById('pl-search').value='';document.getElementById('pl-cat').value='all';renderPriceList()">Đặt lại</button>
        </div>
        ${emptyState}
        ${total > 0 ? `
        <div class="table-wrap">
          <table class="data-table">
            <thead>
              <tr>
                <th>Mã hàng</th>
                <th>Tên hàng</th>
                <th>Cate level 1</th>
                <th>Cate level 2</th>
                <th>Cate level 3</th>
                <th style="text-align:right">Bảng giá chung</th>
              </tr>
            </thead>
            <tbody>
              ${slice.map(p => `<tr>
                <td><code style="font-size:0.8rem">${esc(p.barcode || p.id)}</code></td>
                <td>${esc(p.name || '—')}</td>
                <td><span class="text-muted" style="font-size:0.82rem">${esc(p.cate1 || '—')}</span></td>
                <td><span class="text-muted" style="font-size:0.82rem">${esc(p.cate2 || '—')}</span></td>
                <td><span class="text-muted" style="font-size:0.82rem">${esc(p.cate3 || '—')}</span></td>
                <td style="text-align:right;font-weight:600;color:var(--accent-green)">${p.price ? Utils.formatVND(p.price) : '—'}</td>
              </tr>`).join('')}
            </tbody>
          </table>
        </div>
        ${paginationHtml}
        ` : ''}
      </div>
    </div>`;
}

// ============================================================
// REPORTS
// ============================================================
function renderReports() {
  // ── Time filter ──────────────────────────────────────────
  const cutoff = (() => {
    const d = new Date();
    if (reportPeriod === 'week')     { d.setDate(d.getDate() - 7);    return d.toISOString().slice(0, 10); }
    if (reportPeriod === 'month')    { d.setMonth(d.getMonth() - 1);  return d.toISOString().slice(0, 10); }
    if (reportPeriod === '3months')  { d.setMonth(d.getMonth() - 3);  return d.toISOString().slice(0, 10); }
    if (reportPeriod === '6months')  { d.setMonth(d.getMonth() - 6);  return d.toISOString().slice(0, 10); }
    return null;
  })();

  const txFiltered = MockData.transactions.filter(t => !cutoff || t.date >= cutoff);
  const txDone     = txFiltered.filter(t => t.status === 'completed');
  const imports    = txDone.filter(t => t.type === 'import');

  // Booking filtered by deliveryDate
  const bookings = MockData.bookings.filter(b => !cutoff || (b.deliveryDate && b.deliveryDate >= cutoff));

  // FC data
  const fc = MockData.getFC();

  // ── Top 8 sản phẩm nhập ─────────────────────────────────
  const topMap = imports.reduce((acc, t) => {
    acc[t.productName] = (acc[t.productName] || 0) + t.quantity; return acc;
  }, {});
  const topProducts = Object.entries(topMap).sort((a, b) => b[1] - a[1]).slice(0, 8);

  // ── FC vs Actual (6 tháng gần nhất) ─────────────────────
  const last6M = Array.from({ length: 6 }, (_, i) => {
    const d = new Date(); d.setDate(1); d.setMonth(d.getMonth() - (5 - i));
    return d.toISOString().slice(0, 7);
  });
  const fcByMonth = {};
  if (fc) {
    fc.dates.forEach((date, di) => {
      const key = date.slice(0, 7);
      if (!fcByMonth[key]) fcByMonth[key] = 0;
      fc.categories.forEach((_, ci) => { fcByMonth[key] += fc.values[ci][di]; });
    });
  }
  const importByMonth = {};
  imports.forEach(t => {
    const key = t.date.slice(0, 7);
    importByMonth[key] = (importByMonth[key] || 0) + t.quantity;
  });

  // ── Xu hướng (filtered period) ───────────────────────────
  const periodGranularity = { week: 'day', month: 'day', '3months': 'week', '6months': 'week', all: 'month' };
  const granularity = periodGranularity[reportPeriod] || 'month';
  const trendGroups = Utils.groupByPeriod(txDone, granularity);
  const trendKeys   = Object.keys(trendGroups).sort().slice(-14);
  const trendLabels = trendKeys.map(k => Utils.formatPeriodLabel(k, granularity));

  // ── Booking theo tempCategory ────────────────────────────
  const bkByCat = {};
  bookings.forEach(b => {
    const cat = b.tempCategory || 'Khác';
    bkByCat[cat] = (bkByCat[cat] || 0) + (b.quantity || 0);
  });
  const bkCats = Object.keys(bkByCat);

  // ── Period filter buttons ────────────────────────────────
  const periodBtns = [
    ['week', 'Tuần này'], ['month', 'Tháng này'],
    ['3months', '3 tháng'], ['6months', '6 tháng'], ['all', 'Tất cả'],
  ].map(([v, lbl]) =>
    `<button class="btn btn-sm${reportPeriod === v ? ' btn-primary' : ''}"
      onclick="reportPeriod='${v}';renderReports();">${lbl}</button>`
  ).join('');

  // ── Render ───────────────────────────────────────────────
  document.getElementById('page-content').innerHTML = `
    <div style="display:grid;gap:var(--space-lg);">

      <div class="filter-bar anim-fade-in-up">
        <span style="color:var(--text-muted);font-size:var(--fs-sm);">Kỳ báo cáo:</span>
        ${periodBtns}
      </div>

      <div class="grid-2 anim-fade-in-up anim-delay-1">
        <div class="panel">
          <div class="panel-header">
            <div class="panel-title">📊 FC Dự báo vs Thực nhập (6 tháng)</div>
            ${!fc ? '<span style="font-size:var(--fs-xs);color:var(--accent-amber);">⚠️ Chưa có dữ liệu FC</span>' : ''}
          </div>
          <div class="chart-container chart-wrapper-lg"><canvas id="rpt-fc-vs-actual"></canvas></div>
        </div>
        <div class="panel">
          <div class="panel-header"><div class="panel-title">Top sản phẩm nhập nhiều nhất</div></div>
          <div class="chart-container chart-wrapper-lg"><canvas id="rpt-top"></canvas></div>
        </div>
      </div>

      <div class="grid-2 anim-fade-in-up anim-delay-2">
        <div class="panel">
          <div class="panel-header"><div class="panel-title">Xu hướng nhập / xuất</div></div>
          <div class="chart-container chart-wrapper-lg"><canvas id="rpt-trend"></canvas></div>
        </div>
        <div class="panel">
          <div class="panel-header"><div class="panel-title">📋 Booking theo loại hàng</div></div>
          ${bkCats.length ? `
            <div style="display:grid;grid-template-columns:1fr 1fr;height:280px;">
              <div class="chart-container" style="height:280px;">
                <canvas id="rpt-bk-cat"></canvas>
              </div>
              <div style="display:flex;flex-direction:column;justify-content:center;gap:12px;padding:var(--space-lg);">
                ${bkCats.map(cat => {
                  const colors = { CHILL: '#06b6d4', FROZEN: '#8b5cf6', DRY: '#f59e0b' };
                  const c = colors[cat] || '#64748b';
                  return `<div>
                    <div style="font-size:var(--fs-xs);color:var(--text-muted);margin-bottom:2px;">${cat}</div>
                    <div style="font-size:var(--fs-xl);font-weight:700;color:${c};">
                      ${Utils.formatNumber(Math.round(bkByCat[cat]))}
                    </div>
                  </div>`;
                }).join('')}
                <div style="font-size:var(--fs-xs);color:var(--text-muted);border-top:1px solid var(--border-light);padding-top:8px;">
                  Tổng: ${Utils.formatNumber(Math.round(bkCats.reduce((s, c) => s + bkByCat[c], 0)))} SL
                </div>
              </div>
            </div>` : `
            <div style="display:flex;align-items:center;justify-content:center;height:280px;color:var(--text-muted);">
              <div style="text-align:center;">
                <div style="font-size:2rem;margin-bottom:8px;">📋</div>
                <div style="font-size:var(--fs-sm);">Chưa có dữ liệu booking trong kỳ này</div>
              </div>
            </div>`}
        </div>
      </div>

    </div>`;

  // Chart: FC vs Actual
  const monthLabels = last6M.map(m => { const [y, mo] = m.split('-'); return `T${parseInt(mo)}/${y.slice(2)}`; });
  makeChart('rpt-fc-vs-actual', {
    type: 'bar',
    data: {
      labels: monthLabels,
      datasets: [
        {
          label: 'FC Dự báo (kg)',
          data: last6M.map(m => Math.round(fcByMonth[m] || 0)),
          backgroundColor: '#10b98180', borderColor: '#10b981', borderWidth: 1, borderRadius: 4,
        },
        {
          label: 'Thực nhập (đvt)',
          data: last6M.map(m => importByMonth[m] || 0),
          backgroundColor: '#3b82f680', borderColor: '#3b82f6', borderWidth: 1, borderRadius: 4,
        },
      ],
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { position: 'bottom', labels: { color: '#94a3b8', font: { size: 11 }, boxWidth: 12 } },
        tooltip: { callbacks: { label: ctx => ` ${ctx.dataset.label}: ${Utils.formatNumber(Math.round(ctx.parsed.y))}` } },
      },
      scales: {
        x: { ticks: { color: '#64748b' }, grid: { color: 'rgba(51,65,85,0.3)' } },
        y: { ticks: { color: '#64748b', callback: v => Utils.formatNumber(Math.round(v / 1000)) + 'k' },
             grid: { color: 'rgba(51,65,85,0.3)' } },
      },
    },
  });

  // Chart: Top products
  if (topProducts.length) {
    makeChart('rpt-top', {
      type: 'bar',
      data: {
        labels: topProducts.map(([n]) => n.length > 22 ? n.slice(0, 22) + '…' : n),
        datasets: [{ label: 'Số lượng nhập', data: topProducts.map(([, q]) => q),
          backgroundColor: '#3b82f6bb', borderRadius: 4 }],
      },
      options: {
        indexAxis: 'y', responsive: true, maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: { callbacks: { label: ctx => ` ${Utils.formatNumber(ctx.raw)} đvt` } },
        },
        scales: {
          x: { ticks: { color: '#64748b' }, grid: { color: 'rgba(51,65,85,0.3)' } },
          y: { ticks: { color: '#94a3b8', font: { size: 11 } }, grid: { display: false } },
        },
      },
    });
  }

  // Chart: Trend
  if (trendKeys.length) {
    makeChart('rpt-trend', {
      type: 'line',
      data: {
        labels: trendLabels,
        datasets: [
          {
            label: 'Nhập',
            data: trendKeys.map(k => trendGroups[k].filter(t => t.type === 'import').reduce((s, t) => s + t.total, 0)),
            borderColor: '#3b82f6', backgroundColor: 'rgba(59,130,246,0.08)',
            tension: 0.4, fill: true, pointRadius: 3,
          },
          {
            label: 'Xuất',
            data: trendKeys.map(k => trendGroups[k].filter(t => t.type === 'export').reduce((s, t) => s + t.total, 0)),
            borderColor: '#22c55e', backgroundColor: 'rgba(34,197,94,0.08)',
            tension: 0.4, fill: true, pointRadius: 3,
          },
        ],
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: {
          legend: { labels: { color: '#94a3b8', font: { size: 12 }, boxWidth: 12 } },
          tooltip: { callbacks: { label: ctx => ` ${ctx.dataset.label}: ${Utils.formatVND(ctx.raw)}` } },
        },
        scales: {
          x: { ticks: { color: '#64748b' }, grid: { color: 'rgba(51,65,85,0.3)' } },
          y: { ticks: { color: '#64748b', callback: v => Utils.formatNumber(v / 1_000_000) + 'M' },
               grid: { color: 'rgba(51,65,85,0.3)' } },
        },
      },
    });
  }

  // Chart: Booking by tempCategory
  if (bkCats.length) {
    const catColors = { CHILL: '#06b6d4', FROZEN: '#8b5cf6', DRY: '#f59e0b' };
    makeChart('rpt-bk-cat', {
      type: 'doughnut',
      data: {
        labels: bkCats,
        datasets: [{
          data: bkCats.map(c => Math.round(bkByCat[c])),
          backgroundColor: bkCats.map(c => catColors[c] || '#64748b'),
          borderWidth: 0, hoverOffset: 6,
        }],
      },
      options: {
        responsive: true, maintainAspectRatio: false, cutout: '58%',
        plugins: {
          legend: { display: false },
          tooltip: { callbacks: { label: ctx => ` ${ctx.label}: ${Utils.formatNumber(ctx.raw)}` } },
        },
      },
    });
  }
}

// ============================================================
// SHARED RENDER HELPERS
// ============================================================
function transactionTable(transactions) {
  if (!transactions.length) return `
    <div class="empty-state">
      <div class="empty-icon">📋</div>
      <div class="empty-text">Không có giao dịch nào</div>
    </div>`;

  return `
    <table class="data-table">
      <thead>
        <tr>
          <th>Mã GD</th><th>Loại</th><th>Sản phẩm</th><th>Số lượng</th>
          <th>Giá trị</th><th>Đối tác</th><th>Trạng thái</th><th>Thời gian</th>
        </tr>
      </thead>
      <tbody>
        ${transactions.map(t => `
          <tr>
            <td><span style="color:var(--accent-blue);font-weight:600;">${t.id}</span></td>
            <td>${typeBadge(t.type)}</td>
            <td style="color:var(--text-primary);font-weight:500;">${t.productName}</td>
            <td>${Utils.formatNumber(t.quantity)} ${t.unit}</td>
            <td style="color:${t.type === 'import' ? 'var(--accent-blue-light)' : 'var(--accent-green-light)'};">
              ${Utils.formatVND(t.total)}
            </td>
            <td style="color:var(--text-muted);font-size:var(--fs-xs);">${t.supplier || t.customer || '—'}</td>
            <td>${statusBadge(t.status, t.type)}</td>
            <td style="color:var(--text-muted);font-size:var(--fs-xs);">${Utils.timeAgo(t.date)}</td>
          </tr>`).join('')}
      </tbody>
    </table>`;
}

function lowStockList(products) {
  if (!products.length) return `
    <div class="empty-state" style="padding:var(--space-lg);">
      <div class="empty-icon">✅</div>
      <div class="empty-text">Kho hàng đầy đủ</div>
    </div>`;

  return `<div style="padding:var(--space-xs) 0;">
    ${products.slice(0, 10).map(p => {
      const pct = Math.min(Math.round((p.stock / (p.minStock * 1.5)) * 100), 100);
      const color = p.stock <= p.minStock ? 'var(--accent-red)' : 'var(--accent-amber)';
      return `
        <div style="padding:10px var(--space-md);border-bottom:1px solid var(--border-light);">
          <div style="display:flex;justify-content:space-between;margin-bottom:4px;">
            <span style="font-size:var(--fs-sm);font-weight:500;color:var(--text-primary);">${p.name}</span>
            <span style="font-size:var(--fs-xs);color:${color};font-weight:700;">${Utils.formatNumber(p.stock)} ${p.unit}</span>
          </div>
          <div style="height:4px;background:var(--border-light);border-radius:2px;overflow:hidden;">
            <div style="height:100%;width:${pct}%;background:${color};border-radius:2px;transition:width 0.6s ease;"></div>
          </div>
          <div style="font-size:var(--fs-xs);color:var(--text-muted);margin-top:3px;">
            Tối thiểu: ${Utils.formatNumber(p.minStock)} ${p.unit}
          </div>
        </div>`;
    }).join('')}
  </div>`;
}

function typeBadge(type) {
  return type === 'import'
    ? '<span class="badge badge-blue">📥 Nhập</span>'
    : '<span class="badge badge-green">📤 Xuất</span>';
}

function statusBadge(status, type) {
  if (status === 'completed') return '<span class="badge badge-green">Hoàn thành</span>';
  if (status === 'cancelled') return '<span class="badge badge-red">Đã hủy</span>';
  if (status === 'pending') {
    const label = type === 'export' ? 'Đang chuyển' : 'Chờ xử lý';
    return `<span class="badge badge-amber">${label}</span>`;
  }
  return `<span class="badge">${status}</span>`;
}

// ============================================================
// PAGINATION — uses named global handler __pg()
// ============================================================
function pagination(paged, context) {
  if (paged.totalPages <= 1) return '';
  const { currentPage: cp, totalPages, total } = paged;

  let pages = '';
  for (let i = 1; i <= totalPages; i++) {
    if (i === 1 || i === totalPages || Math.abs(i - cp) <= 1) {
      pages += `<button class="page-btn ${i === cp ? 'active' : ''}" onclick="__pg('${context}',${i})">${i}</button>`;
    } else if (Math.abs(i - cp) === 2) {
      pages += `<span style="color:var(--text-muted);padding:0 4px;">…</span>`;
    }
  }

  return `
    <div class="pagination">
      <button class="page-btn" onclick="__pg('${context}',${Math.max(1, cp - 1)})" ${cp === 1 ? 'disabled' : ''}>‹</button>
      ${pages}
      <button class="page-btn" onclick="__pg('${context}',${Math.min(totalPages, cp + 1)})" ${cp === totalPages ? 'disabled' : ''}>›</button>
      <span style="color:var(--text-muted);font-size:var(--fs-xs);margin-left:8px;">Tổng: ${Utils.formatNumber(total)}</span>
    </div>`;
}

function __pg(context, page) {
  switch (context) {
    case 'import':    importPage    = page; renderImports();   break;
    case 'export':    exportPage    = page; renderExports();   break;
    case 'inventory': inventoryPage = page; renderInventory(document.getElementById('global-search').value || ''); break;
    case 'booking':   bookingPage   = page; renderBookings();  break;
  }
}

// ============================================================
// MODAL — Add Transaction
// ============================================================
function showAddModal(type = 'import') {
  const title = document.getElementById('modal-title');
  const body  = document.getElementById('modal-body');

  title.textContent = type === 'import' ? '📥 Tạo phiếu nhập hàng' : '📤 Tạo phiếu xuất hàng';

  const productOpts = MockData.products.map(p =>
    `<option value="${p.id}">${p.id} — ${p.name} (Tồn: ${Utils.formatNumber(p.stock)} ${p.unit})</option>`
  ).join('');

  const partnerOpts = type === 'import'
    ? MockData.suppliers.map(s => `<option value="${s.name}">${s.name}</option>`).join('')
    : MockData.customers.map(c => `<option value="${c.name}">${c.name}</option>`).join('');

  body.innerHTML = `
    <div class="form-group">
      <label class="form-label">Sản phẩm *</label>
      <select id="f-product" class="form-select">
        <option value="">— Chọn sản phẩm —</option>
        ${productOpts}
      </select>
    </div>
    <div class="form-group">
      <label class="form-label">Số lượng *</label>
      <input id="f-qty" type="number" min="1" class="form-input" placeholder="Nhập số lượng">
    </div>
    <div class="form-group">
      <label class="form-label">${type === 'import' ? 'Nhà cung cấp' : 'Khách hàng'} *</label>
      <select id="f-partner" class="form-select">
        <option value="">— Chọn —</option>
        ${partnerOpts}
      </select>
    </div>
    <div class="form-group">
      <label class="form-label">Ghi chú</label>
      <input id="f-note" type="text" class="form-input" placeholder="Ghi chú (tùy chọn)">
    </div>
    <div style="display:flex;gap:var(--space-md);justify-content:flex-end;margin-top:var(--space-xl);">
      <button class="btn btn-secondary" onclick="closeModal()">Hủy</button>
      <button class="btn btn-primary" onclick="submitTransaction('${type}')">
        ${type === 'import' ? '📥 Tạo phiếu nhập' : '📤 Tạo phiếu xuất'}
      </button>
    </div>`;

  document.getElementById('modal-overlay').classList.remove('hidden');
}

function closeModal() {
  document.getElementById('modal-overlay').classList.add('hidden');
  document.querySelector('.modal-box').classList.remove('modal-fullpage');
}

function submitTransaction(type) {
  const productId = document.getElementById('f-product').value;
  const qty       = parseInt(document.getElementById('f-qty').value);
  const partner   = document.getElementById('f-partner').value;
  const note      = document.getElementById('f-note').value;

  if (!productId) { showToast('⚠️ Vui lòng chọn sản phẩm!', 'amber'); return; }
  if (!qty || qty < 1) { showToast('⚠️ Số lượng không hợp lệ!', 'amber'); return; }
  if (!partner) { showToast('⚠️ Vui lòng chọn ' + (type === 'import' ? 'nhà cung cấp' : 'khách hàng') + '!', 'amber'); return; }

  if (type === 'export') {
    const product = MockData.products.find(p => p.id === productId);
    if (product && qty > product.stock) {
      showToast('⚠️ Số lượng xuất vượt tồn kho (' + Utils.formatNumber(product.stock) + ' ' + product.unit + ')!', 'amber');
      return;
    }
  }

  MockData.addTransaction({
    type, productId, quantity: qty,
    supplier: type === 'import' ? partner : null,
    customer: type === 'export' ? partner : null,
    note,
  });

  closeModal();
  refreshLowStockBadge();
  navigateTo(currentPage);
  showToast(type === 'import' ? '✅ Phiếu nhập đã được tạo!' : '✅ Phiếu xuất đã được tạo!', 'green');
}

// ============================================================
// IMPORT EXCEL MODAL
// ============================================================

let _pendingRows = null;
let _importTab   = 'import';

// ── Import tab helpers ───────────────────────────────────
const IMPORT_CFG = {
  import: {
    icon: '📥', color: 'var(--accent-blue)', label: 'Nhập hàng',
    hint: 'DanhSachChiTietNhapHang_*.xlsx',
    cols: ['Tên hàng, ĐVT, Giá PO', 'Số lượng PR (thực nhận)', 'Ngày nhận hàng', 'Mã PO, Mã nội bộ'],
  },
  export: {
    icon: '📤', color: 'var(--accent-green)', label: 'Xuất hàng',
    hint: 'transfer_*.xlsx',
    cols: ['Tên hàng, Đơn vị tính', 'Số lượng chuyển', 'Ngày chuyển hàng', 'Mã chuyển hàng'],
  },
  pricelist: {
    icon: '💰', color: 'var(--accent-amber)', label: 'Bảng giá',
    hint: 'bang-gia_*.xlsx',
    cols: ['Mã hàng', 'Bảng giá chung / Giá bán', 'Tên hàng (khuyến nghị)', 'Cate level 1/2/3 (tùy chọn)'],
  },
  booking: {
    icon: '📋', color: '#8b5cf6', label: 'Booking',
    hint: 'MD 1_KFM_BOOKING_*.xlsx',
    cols: ['Item Name / Tên hàng', 'Quantity / Số lượng', 'Delivery Date (ngày giao)', 'PO, Supplier Name'],
  },
  fc: {
    icon: '📊', color: '#10b981', label: 'FC (Forecast)',
    hint: 'FC ABAMD THÁNG *.xlsx',
    cols: ['Dòng 3: Ngày (date serial, cột N→)', 'Dòng 7: FINAL total (cột B-M)', 'Dòng 13+: Ngành hàng bắt đầu bằng "3."', 'Import sẽ ghi đè FC hiện tại'],
  },
  masterdata: {
    icon: '🗂️', color: '#0ea5e9', label: 'Mã hàng',
    hint: 'Master data thịt cá.xlsx',
    cols: ['Sheet "Trọng lượng trừ bì": BUYER, CATE LEVEL 3', 'Code sản phẩm, Tên sản phẩm, ĐVT', 'Trọng lượng, Quy cách, Code Spec', 'Import sẽ ghi đè Master Data hiện tại'],
  },
};

function renderImportHistory(type) {
  const history = MockData.importHistory.filter(h => h.type === type);
  if (!history.length) {
    return `<div style="text-align:center;padding:20px;color:var(--text-muted);font-size:var(--fs-xs);">
      Chưa có file nào được import
    </div>`;
  }

  const rows = history.slice(0, 50).map(h => {
    const dt = new Date(h.importedAt);
    const dateStr = dt.toLocaleDateString('vi-VN', { day:'2-digit', month:'2-digit', year:'numeric' })
                  + ' ' + dt.toLocaleTimeString('vi-VN', { hour:'2-digit', minute:'2-digit' });
    let result;
    if (h.type === 'pricelist') {
      result = `${Utils.formatNumber(h.updated || 0)} cập nhật, ${Utils.formatNumber(h.added || 0)} thêm mới`;
    } else if (h.type === 'booking') {
      result = `${Utils.formatNumber(h.added || 0)} booking${h.skipped ? ` (bỏ qua ${Utils.formatNumber(h.skipped)} trùng)` : ''}`;
    } else if (h.type === 'fc') {
      result = `${Utils.formatNumber(h.dates || 0)} ngày, ${Utils.formatNumber(h.categories || 0)} ngành hàng`;
    } else if (h.type === 'masterdata') {
      result = `${Utils.formatNumber(h.count || 0)} mã hàng`;
    } else {
      result = `${Utils.formatNumber(h.added || 0)} giao dịch${h.skipped ? ` (bỏ qua ${Utils.formatNumber(h.skipped)} trùng)` : ''}`;
    }

    const canRollback = (h.transactionIds?.length > 0) || (h.productIds?.length > 0) || (h.bookingIds?.length > 0);
    let rollbackNote;
    if (h.type === 'fc' || h.type === 'masterdata') {
      rollbackNote = h.type === 'fc' ? 'Xoá FC (dữ liệu mới nhất sẽ được dùng)' : 'Xoá Master Data';
    } else if (!canRollback) {
      rollbackNote = 'Chỉ xoá khỏi lịch sử';
    } else if (h.type === 'pricelist') {
      rollbackNote = `Xoá ${h.productIds?.length || 0} sản phẩm đã thêm`;
    } else if (h.type === 'booking') {
      rollbackNote = `Hoàn tác ${h.bookingIds?.length || 0} booking`;
    } else {
      rollbackNote = `Hoàn tác ${h.transactionIds?.length || 0} giao dịch`;
    }
    const name = h.fileName.length > 38 ? h.fileName.slice(0, 36) + '…' : h.fileName;

    return `<tr style="border-bottom:1px solid rgba(255,255,255,0.04);">
      <td style="padding:7px 8px;color:var(--text-muted);font-size:var(--fs-xs);white-space:nowrap;">${dateStr}</td>
      <td style="padding:7px 8px;font-size:var(--fs-xs);color:var(--text-secondary);" title="${h.fileName}">${name}</td>
      <td style="padding:7px 8px;font-size:var(--fs-xs);color:var(--text-primary);">${result}</td>
      <td style="padding:7px 8px;text-align:right;">
        <button title="${rollbackNote}"
          style="background:none;border:1px solid var(--accent-red);color:var(--accent-red);border-radius:4px;padding:2px 8px;font-size:var(--fs-xs);cursor:pointer;"
          onclick="deleteImportRecord('${h.id}')">Xoá</button>
      </td>
    </tr>`;
  }).join('');

  return `
    <div style="background:var(--bg-primary);border-radius:var(--radius-md);padding:var(--space-md);">
      <div style="font-size:var(--fs-xs);color:var(--text-muted);font-weight:600;margin-bottom:10px;">
        LỊCH SỬ IMPORT (${history.length} file)
      </div>
      <div style="max-height:200px;overflow-y:auto;">
        <table style="width:100%;border-collapse:collapse;">
          <thead>
            <tr style="border-bottom:1px solid var(--border);">
              <th style="text-align:left;padding:4px 8px;font-size:var(--fs-xs);color:var(--text-muted);font-weight:500;">Thời gian</th>
              <th style="text-align:left;padding:4px 8px;font-size:var(--fs-xs);color:var(--text-muted);font-weight:500;">Tên file</th>
              <th style="text-align:left;padding:4px 8px;font-size:var(--fs-xs);color:var(--text-muted);font-weight:500;">Kết quả</th>
              <th></th>
            </tr>
          </thead>
          <tbody>${rows}</tbody>
        </table>
      </div>
    </div>`;
}

function renderImportTabContent(type) {
  const cfg     = IMPORT_CFG[type];
  const pending = _pendingRows?.[type];
  return `
    <div style="display:grid;gap:var(--space-md);">
      <div class="import-drop-zone${type === 'pricelist' ? ' import-drop-zone--price' : ''}"
           id="zone-${type}"
           onclick="document.getElementById('file-input-${type}').click()"
           style="border-color:${pending ? 'var(--accent-green)' : ''};">
        <div style="font-size:2rem;margin-bottom:6px;">${cfg.icon}</div>
        <div style="font-weight:700;color:var(--text-primary);font-size:var(--fs-sm);margin-bottom:2px;">${cfg.label}</div>
        <div style="font-size:var(--fs-xs);color:var(--text-muted);margin-bottom:4px;">${cfg.hint}</div>
        <div id="preview-${type}" style="margin-top:8px;">
          ${pending
            ? `<div style="color:var(--accent-green);font-size:var(--fs-xs);font-weight:600;">✅ ${pending.fileName} — sẵn sàng import</div>`
            : '<div style="font-size:var(--fs-xs);color:var(--text-muted);">Kéo thả hoặc click để chọn file</div>'}
        </div>
      </div>
      <div style="background:var(--bg-primary);border-radius:var(--radius-md);padding:12px 16px;">
        <div style="font-size:var(--fs-xs);color:var(--text-muted);font-weight:600;margin-bottom:8px;">CỘT BẮT BUỘC</div>
        ${cfg.cols.map(c => `<div style="font-size:var(--fs-xs);color:var(--text-secondary);margin-bottom:3px;">• ${c}</div>`).join('')}
      </div>
      ${renderImportHistory(type)}
    </div>`;
}

function switchImportTab(type) {
  _importTab = type;
  const content = document.getElementById('import-tab-content');
  if (!content) return;
  content.innerHTML = renderImportTabContent(type);

  Object.keys(IMPORT_CFG).forEach(t => {
    const btn = document.getElementById('tab-btn-' + t);
    if (!btn) return;
    const active = t === type;
    btn.style.color            = active ? IMPORT_CFG[t].color : 'var(--text-muted)';
    btn.style.borderBottomColor = active ? IMPORT_CFG[t].color : 'transparent';
    btn.style.fontWeight       = active ? '700' : '500';
  });

  setupDropZone('zone-' + type, type);
}

function deleteImportRecord(historyId) {
  const entry = MockData.importHistory.find(h => h.id === historyId);
  if (!entry) return;

  const canRollback = (entry.transactionIds?.length > 0) || (entry.productIds?.length > 0);
  let msg = `Xoá file "${entry.fileName}" khỏi lịch sử?`;
  if (canRollback) {
    msg += entry.type === 'pricelist'
      ? `\n\nSẽ xoá ${entry.productIds.length} sản phẩm đã thêm từ file này.`
      : `\n\nSẽ hoàn tác ${entry.transactionIds.length} giao dịch đã import từ file này.`;
  }
  if (!confirm(msg)) return;

  MockData.removeImportRecord(historyId);
  refreshLowStockBadge();
  if (entry.type === 'booking' && currentPage === 'booking') renderBookings();
  else if (entry.type === 'fc' && currentPage === 'fc') renderFC();
  else if (entry.type === 'masterdata' && currentPage === 'masterdata') renderMasterData();
  else navigateTo(currentPage);
  switchImportTab(entry.type); // refresh history on the same tab
  showToast(`🗑️ Đã xoá: ${entry.fileName}`, 'amber');
}

function exportData() {
  try {
    const blob = new Blob([MockData.exportJSON()], { type: 'application/json' });
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    a.href     = url;
    a.download = `kfm-backup-${Utils.today()}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    showToast('✅ Đã tải xuống backup dữ liệu', 'green');
  } catch (e) {
    showToast('❌ Lỗi xuất dữ liệu: ' + e.message, 'red');
  }
}

function restoreFromJSON(file) {
  const reader = new FileReader();
  reader.onload = e => {
    try {
      MockData.restoreJSON(e.target.result);
      refreshLowStockBadge();
      navigateTo(currentPage);
      showImportExcelModal();
      showToast('✅ Đã khôi phục dữ liệu từ backup', 'green');
    } catch (err) {
      showToast('❌ File backup không hợp lệ: ' + err.message, 'red');
    }
  };
  reader.readAsText(file);
}

function showImportExcelModal() {
  const title = document.getElementById('modal-title');
  const body  = document.getElementById('modal-body');
  title.textContent = '📂 Import dữ liệu từ Excel';

  const tabs = ['import', 'export', 'pricelist', 'booking', 'fc', 'masterdata'].map(t => {
    const cfg    = IMPORT_CFG[t];
    const active = _importTab === t;
    const hasPending = !!_pendingRows?.[t];
    return `<button id="tab-btn-${t}" onclick="switchImportTab('${t}')"
      style="position:relative;padding:10px 22px;background:none;border:none;cursor:pointer;
             font-size:var(--fs-sm);font-weight:${active ? '700' : '500'};
             color:${active ? cfg.color : 'var(--text-muted)'};
             border-bottom:2px solid ${active ? cfg.color : 'transparent'};
             margin-bottom:-2px;white-space:nowrap;">
      ${cfg.icon} ${cfg.label}
      ${hasPending ? '<span style="position:absolute;top:8px;right:6px;width:7px;height:7px;border-radius:50%;background:var(--accent-green);display:block;"></span>' : ''}
    </button>`;
  }).join('');

  body.innerHTML = `
    <div style="display:grid;gap:var(--space-md);">
      <div style="display:flex;border-bottom:2px solid var(--border);margin:0 -4px;">${tabs}</div>
      <div id="import-tab-content"></div>
      <div style="display:flex;gap:8px;justify-content:space-between;align-items:center;padding-top:4px;border-top:1px solid var(--border);">
        <div style="display:flex;gap:8px;">
          <button class="btn btn-sm" onclick="exportData()">💾 Backup JSON</button>
          <label class="btn btn-sm" style="cursor:pointer;" title="Khôi phục từ file backup JSON">
            📂 Khôi phục
            <input type="file" accept=".json" style="display:none;"
              onchange="restoreFromJSON(this.files[0]);this.value=''">
          </label>
          <button class="btn btn-danger btn-sm" onclick="confirmClearAll()">🗑️ Xóa toàn bộ</button>
        </div>
        <div style="display:flex;gap:var(--space-md);">
          <button class="btn btn-secondary" onclick="closeModal()">Đóng</button>
          <button class="btn btn-primary" id="btn-confirm-import" disabled onclick="confirmImport()">
            ✅ Xác nhận Import
          </button>
        </div>
      </div>
    </div>`;

  document.querySelector('.modal-box').classList.add('modal-fullpage');
  document.getElementById('modal-overlay').classList.remove('hidden');

  document.getElementById('file-input-import').onchange    = e => handleFileSelect(e, 'import');
  document.getElementById('file-input-export').onchange    = e => handleFileSelect(e, 'export');
  document.getElementById('file-input-pricelist').onchange = e => handleFileSelect(e, 'pricelist');
  document.getElementById('file-input-booking').onchange   = e => handleFileSelect(e, 'booking');
  document.getElementById('file-input-fc').onchange         = e => handleFileSelect(e, 'fc');
  document.getElementById('file-input-masterdata').onchange  = e => handleFileSelect(e, 'masterdata');

  switchImportTab(_importTab);

  // Re-enable confirm button if there are already pending rows from a previous file load
  const hasAnyPending = _pendingRows && Object.values(_pendingRows).some(Boolean);
  const confirmBtn = document.getElementById('btn-confirm-import');
  if (confirmBtn && hasAnyPending) confirmBtn.disabled = false;
}

function setupDropZone(zoneId, type) {
  const zone = document.getElementById(zoneId);
  if (!zone) return;
  zone.addEventListener('dragover', e => { e.preventDefault(); zone.style.borderColor = 'var(--accent-blue)'; });
  zone.addEventListener('dragleave', () => { zone.style.borderColor = ''; });
  zone.addEventListener('drop', e => {
    e.preventDefault();
    zone.style.borderColor = '';
    const file = e.dataTransfer.files[0];
    if (file) processExcelFile(file, type);
  });
}

async function handleFileSelect(e, type) {
  const file = e.target.files[0];
  if (!file) return;
  e.target.value = ''; // reset so same file can be re-selected
  await processExcelFile(file, type);
}

async function processExcelFile(file, type) {
  // Auto-switch to the relevant tab so preview elements exist
  if (_importTab !== type) switchImportTab(type);

  const previewEl = document.getElementById('preview-' + type);
  const zoneEl    = document.getElementById('zone-' + type);
  if (!previewEl || !zoneEl) return;

  previewEl.innerHTML = `<div style="color:var(--text-muted);font-size:var(--fs-xs);">⏳ Đang đọc file <b>${file.name}</b>...</div>`;

  try {
    // Yield to browser so "đang đọc" message renders before heavy parse
    await new Promise(r => setTimeout(r, 30));

    const wb = await ExcelImport.readFile(file);

    previewEl.innerHTML = `<div style="color:var(--text-muted);font-size:var(--fs-xs);">⏳ Đang phân tích dữ liệu...</div>`;
    await new Promise(r => setTimeout(r, 30));

    let rows, previewHtml;

    if (type === 'fc') {
      const fcData = ExcelImport.parseFC(wb);
      const info   = ExcelImport.previewFC(fcData);
      previewHtml = `
        <div style="background:var(--bg-primary);border-radius:var(--radius-sm);padding:10px;text-align:left;">
          <div style="color:#10b981;font-weight:700;font-size:var(--fs-sm);margin-bottom:6px;">
            ✅ ${file.name}
          </div>
          <div style="display:grid;grid-template-columns:1fr 1fr;gap:4px;font-size:var(--fs-xs);color:var(--text-muted);">
            <span>Kỳ dự báo:</span>    <span style="color:var(--text-primary);font-weight:600;">${info.yearFrom} → ${info.yearTo}</span>
            <span>Số ngày:</span>       <span style="color:#10b981;font-weight:600;">${Utils.formatNumber(info.totalDays)} ngày</span>
            <span>Ngành hàng:</span>    <span style="color:var(--text-primary);font-weight:600;">${info.categories} danh mục</span>
            <span>TB/ngày (FINAL):</span><span style="color:var(--accent-amber);font-weight:600;">${Utils.formatNumber(Math.round(info.avgTotal))} kg</span>
          </div>
          <div style="margin-top:6px;font-size:var(--fs-xs);color:var(--accent-amber);">⚠️ Import sẽ ghi đè dữ liệu FC hiện tại</div>
        </div>`;

      previewEl.innerHTML = previewHtml;
      zoneEl.style.borderColor = 'var(--accent-green)';
      if (!_pendingRows) _pendingRows = {};
      _pendingRows.fc = { fcData, fileName: file.name };
      const btn = document.getElementById('btn-confirm-import');
      if (btn) btn.disabled = false;
      return;
    } else if (type === 'masterdata') {
      const mdRows = ExcelImport.parseMasterData(wb);
      const cats   = [...new Set(mdRows.map(r => r.category).filter(Boolean))].sort();
      const supps  = new Set(mdRows.map(r => r.supplier).filter(Boolean)).size;
      const catLines = cats.map(c => {
        const cnt = mdRows.filter(r => r.category === c).length;
        return `<span style="color:var(--text-primary);font-weight:600;">${c.replace('3.','')}</span>: ${cnt} sp`;
      }).join(' &nbsp;|&nbsp; ');
      const previewHtml2 = `
        <div style="background:var(--bg-primary);border-radius:var(--radius-sm);padding:10px;text-align:left;">
          <div style="color:#0ea5e9;font-weight:700;font-size:var(--fs-sm);margin-bottom:6px;">✅ ${file.name}</div>
          <div style="display:grid;grid-template-columns:1fr 1fr;gap:4px;font-size:var(--fs-xs);color:var(--text-muted);">
            <span>Tổng sản phẩm:</span> <span style="color:#0ea5e9;font-weight:700;">${Utils.formatNumber(mdRows.length)} mã hàng</span>
            <span>Nhà cung cấp:</span>  <span style="color:var(--text-primary);font-weight:600;">${supps} NCC</span>
          </div>
          <div style="margin-top:6px;font-size:var(--fs-xs);color:var(--text-muted);">${catLines}</div>
          <div style="margin-top:6px;font-size:var(--fs-xs);color:var(--accent-amber);">⚠️ Import sẽ ghi đè Master Data hiện tại</div>
        </div>`;
      previewEl.innerHTML = previewHtml2;
      zoneEl.style.borderColor = 'var(--accent-green)';
      if (!_pendingRows) _pendingRows = {};
      _pendingRows.masterdata = { mdRows, fileName: file.name };
      const btnMd = document.getElementById('btn-confirm-import');
      if (btnMd) btnMd.disabled = false;
      return;
    } else if (type === 'booking') {
      rows = ExcelImport.parseBooking(wb);
      const info = ExcelImport.previewBooking(rows);
      const fmtDate = d => d ? new Date(d).toLocaleDateString('vi-VN', { day: '2-digit', month: '2-digit', year: 'numeric' }) : '—';
      previewHtml = `
        <div style="background:var(--bg-primary);border-radius:var(--radius-sm);padding:10px;text-align:left;">
          <div style="color:#8b5cf6;font-weight:700;font-size:var(--fs-sm);margin-bottom:6px;">
            ✅ ${file.name}
          </div>
          <div style="display:grid;grid-template-columns:1fr 1fr;gap:4px;font-size:var(--fs-xs);color:var(--text-muted);">
            <span>Tổng dòng:</span>      <span style="color:var(--text-primary);font-weight:600;">${Utils.formatNumber(info.total)}</span>
            <span>Số PO:</span>          <span style="color:#8b5cf6;font-weight:600;">${Utils.formatNumber(info.pos)}</span>
            <span>Nhà cung cấp:</span>   <span style="color:var(--text-primary);font-weight:600;">${Utils.formatNumber(info.suppliers)}</span>
            <span>Cửa hàng:</span>       <span style="color:var(--text-primary);font-weight:600;">${Utils.formatNumber(info.stores)}</span>
            <span>Tổng SL đặt:</span>    <span style="color:var(--accent-green);font-weight:600;">${Utils.formatNumber(info.totalQty)}</span>
            <span>Ngày giao:</span>      <span style="color:var(--text-primary);">${fmtDate(info.dateFrom)}${info.dateTo !== info.dateFrom ? ' → ' + fmtDate(info.dateTo) : ''}</span>
          </div>
        </div>`;
    } else if (type === 'pricelist') {
      rows = ExcelImport.parsePriceList(wb);
      const info = ExcelImport.previewPriceList(rows, MockData.products);
      previewHtml = `
        <div style="background:var(--bg-primary);border-radius:var(--radius-sm);padding:10px;text-align:left;">
          <div style="color:var(--accent-amber);font-weight:700;font-size:var(--fs-sm);margin-bottom:6px;">
            ✅ ${file.name}
          </div>
          <div style="display:grid;grid-template-columns:1fr 1fr;gap:4px;font-size:var(--fs-xs);color:var(--text-muted);">
            <span>Tổng sản phẩm:</span> <span style="color:var(--text-primary);font-weight:600;">${Utils.formatNumber(info.total)}</span>
            <span>Cập nhật giá:</span>  <span style="color:var(--accent-green);font-weight:600;">${Utils.formatNumber(info.updated)}</span>
            <span>Thêm mới:</span>      <span style="color:var(--accent-blue);font-weight:600;">${Utils.formatNumber(info.added)}</span>
            <span>Giá thấp nhất:</span> <span style="color:var(--text-primary);">${Utils.formatVND(info.minPrice)}</span>
            <span>Giá cao nhất:</span>  <span style="color:var(--text-primary);">${Utils.formatVND(info.maxPrice)}</span>
          </div>
        </div>`;
    } else {
      rows = type === 'import'
        ? ExcelImport.parseImport(wb)
        : ExcelImport.parseExport(wb);
      const info = ExcelImport.preview(rows);
      previewHtml = `
        <div style="background:var(--bg-primary);border-radius:var(--radius-sm);padding:10px;text-align:left;">
          <div style="color:var(--accent-green);font-weight:700;font-size:var(--fs-sm);margin-bottom:6px;">
            ✅ ${file.name}
          </div>
          <div style="display:grid;grid-template-columns:1fr 1fr;gap:4px;font-size:var(--fs-xs);color:var(--text-muted);">
            <span>Tổng dòng:</span>   <span style="color:var(--text-primary);font-weight:600;">${Utils.formatNumber(info.total)}</span>
            <span>Sản phẩm:</span>    <span style="color:var(--text-primary);font-weight:600;">${Utils.formatNumber(info.products)}</span>
            <span>Đối tác:</span>     <span style="color:var(--text-primary);font-weight:600;">${Utils.formatNumber(info.partners)}</span>
            <span>Hoàn thành:</span>  <span style="color:var(--accent-green);font-weight:600;">${Utils.formatNumber(info.byStatus.completed)}</span>
            <span>Đang chuyển:</span> <span style="color:var(--accent-amber);font-weight:600;">${Utils.formatNumber(info.byStatus.pending)}</span>
            <span>Đã hủy:</span>      <span style="color:var(--accent-red);font-weight:600;">${Utils.formatNumber(info.byStatus.cancelled || 0)}</span>
          </div>
        </div>`;
    }

    previewEl.innerHTML = previewHtml;

    zoneEl.style.borderColor = 'var(--accent-green)';

    if (!_pendingRows) _pendingRows = {};
    _pendingRows[type] = { rows, fileName: file.name };

    const btn = document.getElementById('btn-confirm-import');
    if (btn) btn.disabled = false;

  } catch (err) {
    previewEl.innerHTML = `<div style="color:var(--accent-red);font-size:var(--fs-xs);">❌ ${err.message}</div>`;
    zoneEl.style.borderColor = 'var(--accent-red)';
  }
}

function confirmImport() {
  if (!_pendingRows) return;

  const messages = [];
  const errors   = [];

  // Transaction imports (nhập hàng + chuyển hàng)
  let txAdded = 0, txSkipped = 0;
  ['import', 'export'].forEach(type => {
    const pending = _pendingRows[type];
    if (!pending) return;
    try {
      console.log(`[KFM] confirmImport: processing ${type} — ${pending.rows.length} rows`);
      const r = MockData.importRows(pending.rows);
      txAdded   += r.added;
      txSkipped += r.skipped;
      MockData.addImportRecord({
        type,
        fileName:       pending.fileName,
        added:          r.added,
        skipped:        r.skipped,
        transactionIds: r.addedIds,
      });
      console.log(`[KFM] confirmImport: ${type} OK — added=${r.added} skipped=${r.skipped}`);
    } catch (err) {
      console.error(`[KFM] confirmImport: ${type} FAILED —`, err);
      errors.push(`${type}: ${err.message}`);
    }
  });
  if (txAdded > 0 || txSkipped > 0) {
    messages.push(`${Utils.formatNumber(txAdded)} giao dịch mới` +
      (txSkipped ? ` (bỏ qua ${txSkipped} trùng)` : ''));
  }

  // Price list import
  if (_pendingRows.pricelist) {
    const pending = _pendingRows.pricelist;
    try {
      console.log(`[KFM] confirmImport: processing pricelist — ${pending.rows.length} rows`);
      const r = MockData.importPriceList(pending.rows);
      MockData.addImportRecord({
        type:       'pricelist',
        fileName:   pending.fileName,
        updated:    r.updated,
        added:      r.added,
        productIds: r.addedIds,
      });
      messages.push(`${Utils.formatNumber(r.updated)} giá cập nhật, ${Utils.formatNumber(r.added)} sản phẩm mới`);
      console.log(`[KFM] confirmImport: pricelist OK — updated=${r.updated} added=${r.added}`);
    } catch (err) {
      console.error('[KFM] confirmImport: pricelist FAILED —', err);
      errors.push(`pricelist: ${err.message}`);
    }
  }

  // Master Data import
  if (_pendingRows.masterdata) {
    const pending = _pendingRows.masterdata;
    try {
      MockData.setMasterData(pending.mdRows);
      MockData.addImportRecord({
        type:     'masterdata',
        fileName: pending.fileName,
        count:    pending.mdRows.length,
      });
      messages.push(`Master Data: ${Utils.formatNumber(pending.mdRows.length)} mã hàng`);
      if (currentPage === 'masterdata') renderMasterData();
    } catch (err) {
      errors.push(`Master Data: ${err.message}`);
    }
  }

  // FC import
  if (_pendingRows.fc) {
    const pending = _pendingRows.fc;
    try {
      console.log(`[KFM] confirmImport: processing FC — ${pending.fcData.dates.length} dates`);
      MockData.setFC({
        ...pending.fcData,
        fileName:   pending.fileName,
        importedAt: new Date().toISOString(),
      });
      MockData.addImportRecord({
        type:       'fc',
        fileName:   pending.fileName,
        dates:      pending.fcData.dates.length,
        categories: pending.fcData.categories.length,
      });
      messages.push(`FC: ${pending.fcData.dates.length} ngày dự báo, ${pending.fcData.categories.length} ngành hàng`);
      console.log('[KFM] confirmImport: FC OK');
    } catch (err) {
      console.error('[KFM] confirmImport: FC FAILED —', err);
      errors.push(`FC: ${err.message}`);
    }
  }

  // Booking import
  if (_pendingRows.booking) {
    const pending = _pendingRows.booking;
    try {
      console.log(`[KFM] confirmImport: processing booking — ${pending.rows.length} rows`);
      const r = MockData.importBookings(pending.rows);
      MockData.addImportRecord({
        type:       'booking',
        fileName:   pending.fileName,
        added:      r.added,
        skipped:    r.skipped,
        bookingIds: r.addedIds,
      });
      messages.push(`${Utils.formatNumber(r.added)} booking mới`);
      console.log(`[KFM] confirmImport: booking OK — added=${r.added} skipped=${r.skipped}`);
    } catch (err) {
      console.error('[KFM] confirmImport: booking FAILED —', err);
      errors.push(`booking: ${err.message}`);
    }
  }

  // Belt-and-suspenders: final save after all operations
  MockData.save();

  if (errors.length) {
    showToast('⚠️ Một số file lỗi: ' + errors.join(' | '), 'amber');
  }

  closeModal();
  refreshLowStockBadge();
  navigateTo(currentPage);

  if (messages.length) {
    showToast('✅ ' + messages.join(' | '), 'green');
  } else if (!errors.length) {
    showToast('✅ Import hoàn tất', 'green');
  }
  _pendingRows = null;
}

function confirmClearAll() {
  if (!confirm('Xóa toàn bộ dữ liệu giao dịch, sản phẩm, nhà cung cấp, khách hàng?\nHành động này không thể hoàn tác!')) return;
  MockData.clearAll();
  closeModal();
  refreshLowStockBadge();
  navigateTo('dashboard');
  showToast('🗑️ Đã xóa toàn bộ dữ liệu', 'amber');
}

// ============================================================
// UTILITIES
// ============================================================
function recalcStock() {
  MockData._recalcStock();
  MockData.save();
  renderInventory(document.getElementById('global-search')?.value || '');
  refreshLowStockBadge();
  showToast('✅ Đã tính lại tồn kho', 'green');
}

function refreshLowStockBadge() {
  const badge = document.getElementById('badge-lowstock');
  if (!badge) return;
  const count = MockData.getLowStockProducts().length;
  if (count > 0) {
    badge.textContent = count;
    badge.style.display = '';
  } else {
    badge.style.display = 'none';
  }
}

function showToast(msg, type = 'green') {
  const toast = document.getElementById('toast');
  const colors = { green: 'var(--accent-green)', amber: 'var(--accent-amber)', red: 'var(--accent-red)' };
  toast.textContent = msg;
  toast.style.borderColor = colors[type] || colors.green;
  toast.classList.remove('hidden');
  setTimeout(() => toast.classList.add('hidden'), 3000);
}
