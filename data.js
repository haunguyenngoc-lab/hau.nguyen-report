// ===================================
// DATA.JS — Data Store (localStorage backed)
// ===================================

const CATEGORIES = [
  { id: 'thit-heo',  name: 'Thịt heo',       color: '#ef4444' },
  { id: 'thit-bo',   name: 'Thịt bò',         color: '#f97316' },
  { id: 'thit-ga',   name: 'Thịt gà/vịt',     color: '#f59e0b' },
  { id: 'ca',        name: 'Cá',               color: '#3b82f6' },
  { id: 'tom-cua',   name: 'Tôm/Cua/Mực',     color: '#06b6d4' },
  { id: 'hai-san',   name: 'Hải sản khác',     color: '#8b5cf6' },
  { id: 'rau-cu',    name: 'Rau củ quả',       color: '#22c55e' },
  { id: 'trung-sua', name: 'Trứng/Sữa',        color: '#a855f7' },
  { id: 'khac',      name: 'Khác',             color: '#64748b' },
];

const STORAGE_KEY    = 'kfm_data_v2'; // v2 = LZ-String compressed
const STORAGE_KEY_V1 = 'kfm_data_v1'; // legacy uncompressed

const MockData = {
  categories:    CATEGORIES,
  suppliers:     [],
  customers:     [],
  products:      [],
  transactions:  [],
  bookings:      [],
  fc:            null,
  masterData:    [],
  importHistory: [],
  dailyNotes:    {},

  // ── Persistence (IndexedDB primary, localStorage fallback) ──
  _db:           null,
  _dbName:       'kfm_idb_v1',
  _storeName:    'data',
  _saveInFlight: false,
  _saveDirty:    false,

  // ── IndexedDB helpers ─────────────────────────────────────
  _openDB() {
    if (this._db) return Promise.resolve(this._db);
    return new Promise((resolve, reject) => {
      const req = indexedDB.open(this._dbName, 1);
      req.onupgradeneeded = e => {
        const db = e.target.result;
        if (!db.objectStoreNames.contains(this._storeName)) {
          db.createObjectStore(this._storeName);
        }
      };
      req.onsuccess = e => { this._db = e.target.result; resolve(this._db); };
      req.onerror   = e => reject(e.target.error);
    });
  },

  async _idbLoad() {
    const db    = await this._openDB();
    const keys  = ['transactions','bookings','products','suppliers','customers','fc','masterData','importHistory','dailyNotes'];
    return new Promise((resolve, reject) => {
      const tx    = db.transaction(this._storeName, 'readonly');
      const store = tx.objectStore(this._storeName);
      const out   = {};
      let pending = keys.length;
      keys.forEach(k => {
        const req    = store.get(k);
        req.onsuccess = () => { out[k] = req.result ?? null; if (--pending === 0) resolve(out); };
        req.onerror   = e => reject(e.target.error);
      });
    });
  },

  async _idbSave(data) {
    const db = await this._openDB();
    return new Promise((resolve, reject) => {
      const tx    = db.transaction(this._storeName, 'readwrite');
      const store = tx.objectStore(this._storeName);
      Object.entries(data).forEach(([k, v]) => store.put(v, k));
      tx.oncomplete = resolve;
      tx.onerror    = e => reject(e.target.error);
    });
  },

  async _idbClear() {
    const db = await this._openDB();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(this._storeName, 'readwrite');
      tx.objectStore(this._storeName).clear();
      tx.oncomplete = resolve;
      tx.onerror    = e => reject(e.target.error);
    });
  },

  // ── Legacy localStorage helpers (migration + fallback) ────
  _compress(str) {
    try { return (typeof LZString !== 'undefined') ? LZString.compressToBase64(str) : str; } catch (_) { return str; }
  },
  _decompress(str) {
    if (!str || typeof LZString === 'undefined') return str;
    try { const d = LZString.decompressFromBase64(str); return (d && d.length > 0) ? d : str; } catch (_) { return str; }
  },

  _loadFromLocalStorage() {
    const keys = [STORAGE_KEY, STORAGE_KEY_V1];
    for (const key of keys) {
      try {
        const raw = localStorage.getItem(key) || sessionStorage.getItem(key);
        if (!raw) continue;
        const saved = JSON.parse(this._decompress(raw));
        this.transactions  = saved.transactions  || [];
        this.products      = saved.products      || [];
        this.suppliers     = saved.suppliers     || [];
        this.customers     = saved.customers     || [];
        this.bookings      = saved.bookings      || [];
        this.fc            = saved.fc            || null;
        this.importHistory = saved.importHistory || [];
        console.log('[KFM] Loaded from localStorage:', this.transactions.length, 'txns');
        return true;
      } catch (_) {}
    }
    return false;
  },

  // ── Init (async) ──────────────────────────────────────────
  async init() {
    try {
      const saved = await this._idbLoad();
      const hasData = saved.transactions?.length || saved.bookings?.length || saved.products?.length;

      if (hasData) {
        this.transactions  = saved.transactions  || [];
        this.products      = saved.products      || [];
        this.suppliers     = saved.suppliers     || [];
        this.customers     = saved.customers     || [];
        this.bookings      = saved.bookings      || [];
        this.fc            = saved.fc            || null;
        this.masterData    = saved.masterData    || [];
        this.importHistory = saved.importHistory || [];
        this.dailyNotes    = saved.dailyNotes    || {};
        console.log('[KFM] Loaded from IndexedDB:', this.transactions.length, 'txns,', this.bookings.length, 'bookings');
      } else {
        // First run — try migrating from localStorage
        const migrated = this._loadFromLocalStorage();
        if (migrated) {
          this.masterData = this.masterData || [];
          await this._idbSave(this._payload());
          try { localStorage.removeItem(STORAGE_KEY); localStorage.removeItem(STORAGE_KEY_V1); } catch (_) {}
          console.log('[KFM] Migrated localStorage → IndexedDB');
        }
      }
    } catch (e) {
      console.warn('[KFM] IndexedDB unavailable, using localStorage:', e.message);
      this._loadFromLocalStorage();
    }

    this._fixTransferStatuses();
    this._recalcStock();
    console.log('[KFM] Init complete —', this.transactions.length, 'txns,', this.bookings.length, 'bookings');
  },

  // Fix bug: "Đang chuyển" was mapped to 'cancelled' because "chuyển" contains "huy"
  _fixTransferStatuses() {
    let changed = false;
    this.transactions.forEach(t => {
      if (t.type === 'export' && t.status === 'cancelled' && t.maPO) {
        // Export transactions with a valid maPO that aren't truly cancelled
        // should be 'pending' (Đang chuyển). True cancellations have no maPO or
        // are explicitly marked — we re-derive from the raw note field isn't available,
        // so conservatively: if customer (Chi nhánh nhận) exists → was "Đang chuyển"
        if (t.customer) {
          t.status = 'pending';
          changed = true;
        }
      }
    });
    if (changed) this.save();
  },

  _payload() {
    return {
      transactions:  this.transactions,
      products:      this.products,
      suppliers:     this.suppliers,
      customers:     this.customers,
      bookings:      this.bookings,
      fc:            this.fc,
      masterData:    this.masterData,
      importHistory: this.importHistory,
      dailyNotes:    this.dailyNotes,
    };
  },

  getDailyNote(date) { return this.dailyNotes[date] || {}; },
  setDailyNote(date, field, value) {
    if (!this.dailyNotes[date]) this.dailyNotes[date] = {};
    this.dailyNotes[date][field] = value;
    this.save();
  },

  save() {
    if (this._saveInFlight) { this._saveDirty = true; return true; }
    this._saveInFlight = true;
    this._idbSave(this._payload())
      .then(() => {
        console.log('[KFM] IndexedDB save ✓ —', this.transactions.length, 'txns,', this.bookings.length, 'bookings');
      })
      .catch(e => {
        console.warn('[KFM] IndexedDB save failed, trying localStorage:', e.message);
        try {
          const compressed = this._compress(JSON.stringify(this._payload()));
          localStorage.setItem(STORAGE_KEY, compressed);
          console.log('[KFM] localStorage fallback save ✓ —', (compressed.length/1024).toFixed(1), 'KB');
        } catch (lsErr) {
          console.error('[KFM] All storage failed:', lsErr.message);
          try { document.dispatchEvent(new CustomEvent('kfm-save-error', { detail: lsErr.message })); } catch (_) {}
        }
      })
      .finally(() => {
        this._saveInFlight = false;
        if (this._saveDirty) { this._saveDirty = false; this.save(); }
      });
    return true;
  },

  exportJSON() {
    return JSON.stringify({
      exportedAt:    new Date().toISOString(),
      transactions:  this.transactions,
      products:      this.products,
      suppliers:     this.suppliers,
      customers:     this.customers,
      bookings:      this.bookings,
      fc:            this.fc,
      importHistory: this.importHistory,
    }, null, 2);
  },

  restoreJSON(jsonStr) {
    const d = JSON.parse(jsonStr);
    this.transactions  = d.transactions  || [];
    this.products      = d.products      || [];
    this.suppliers     = d.suppliers     || [];
    this.customers     = d.customers     || [];
    this.bookings      = d.bookings      || [];
    this.fc            = d.fc            || null;
    this.importHistory = d.importHistory || [];
    this._recalcStock();
    this.save();
  },

  clearAll() {
    this.transactions  = [];
    this.products      = [];
    this.suppliers     = [];
    this.customers     = [];
    this.bookings      = [];
    this.fc            = null;
    this.masterData    = [];
    this.importHistory = [];
    this.dailyNotes    = {};
    this._idbClear().catch(console.error);
    try { localStorage.removeItem(STORAGE_KEY); } catch (_) {}
    try { localStorage.removeItem(STORAGE_KEY_V1); } catch (_) {}
  },

  clearProducts() {
    this.products = [];
    this.save();
  },

  addImportRecord(entry) {
    this.importHistory.unshift({
      id:         'IMP-' + Date.now(),
      importedAt: new Date().toISOString(),
      ...entry,
    });
    if (this.importHistory.length > 100) this.importHistory.splice(100);
    this.save();
  },

  removeImportRecord(historyId) {
    const idx = this.importHistory.findIndex(h => h.id === historyId);
    if (idx < 0) return;
    const entry = this.importHistory[idx];

    if (entry.type === 'import' || entry.type === 'export') {
      const removeSet = new Set(entry.transactionIds || []);
      if (removeSet.size) {
        this.transactions = this.transactions.filter(t => !removeSet.has(t.id));
        this._recalcStock();
      }
    } else if (entry.type === 'booking') {
      const removeSet = new Set(entry.bookingIds || []);
      if (removeSet.size) {
        this.bookings = this.bookings.filter(b => !removeSet.has(b.id));
      }
    } else if (entry.type === 'pricelist') {
      // Remove newly added products; products that were only updated keep their data
      const removeSet = new Set(entry.productIds || []);
      if (removeSet.size) {
        this.products = this.products.filter(p => !removeSet.has(p.id));
      }
    }

    this.importHistory.splice(idx, 1);
    this.save();
  },

  // ── Import from parsed Excel rows ────────────────────────
  importRows(rows) {
    const existingIds = new Set(this.transactions.map(t => t.id));
    let added = 0;
    let skipped = 0;
    const addedIds = [];

    rows.forEach(row => {
      if (existingIds.has(row.id)) { skipped++; return; }

      this.transactions.push(row);
      existingIds.add(row.id);
      addedIds.push(row.id);
      added++;

      // Auto-register product
      if (row.productId && !this.products.find(p => p.id === row.productId)) {
        this.products.push({
          id:       row.productId,
          name:     row.productName,
          barcode:  row.barcode || '',
          category: row.category || 'khac',
          unit:     row.unit,
          price:    row.price || 0,
          stock:    0,
          minStock: 0,
        });
      }

      // Auto-register supplier
      if (row.supplier && !this.suppliers.find(s => s.name === row.supplier)) {
        this.suppliers.push({
          id:      'ncc-' + this.suppliers.length,
          name:    row.supplier,
          code:    row.supplierCode || '',
          contact: '',
        });
      }

      // Auto-register customer
      if (row.customer && !this.customers.find(c => c.name === row.customer)) {
        this.customers.push({ id: 'kh-' + this.customers.length, name: row.customer });
      }
    });

    // Recompute stock from all completed transactions
    this._recalcStock();
    this.transactions.sort((a, b) => new Date(b.date) - new Date(a.date));
    this.save();

    return { added, skipped, addedIds };
  },

  // Tính tồn kho cho một ngày cụ thể: bắt đầu từ 0, cộng nhập, trừ xuất trong ngày đó.
  _calcStockForDate(date) {
    const d = date || new Date().toISOString().slice(0, 10);
    const stock = {};

    // Pass 1: imports — seed barcodes và cộng số lượng hoàn thành
    this.transactions
      .filter(t => t.type === 'import' && t.barcode && t.date && t.date.startsWith(d))
      .forEach(t => {
        if (!(t.barcode in stock)) stock[t.barcode] = 0;
        if (t.status === 'completed') stock[t.barcode] += t.quantity;
      });

    // Pass 2: exports — chỉ trừ những mã hàng đã có nhập trong ngày
    this.transactions
      .filter(t => t.type === 'export' && t.barcode && t.date && t.date.startsWith(d) && (t.barcode in stock))
      .forEach(t => {
        const slChuyen = t.transferQty != null ? t.transferQty : t.quantity;
        const slNhan   = t.actualQty   != null ? t.actualQty   : null;
        stock[t.barcode] = Math.max(0, stock[t.barcode] - slChuyen);
        if (t.status === 'cancelled') {
          stock[t.barcode] += slChuyen;
        } else if (t.status === 'completed' && slNhan != null) {
          const diff = slChuyen - slNhan;
          if (diff > 0) stock[t.barcode] += diff;
        }
      });

    return stock;
  },

  // Cập nhật p.stock theo ngày (mặc định = hôm nay)
  _recalcStock(date) {
    const stock = this._calcStockForDate(date);
    this.products.forEach(p => {
      const bc = p.barcode || p.id;
      p.stock = stock[bc] !== undefined ? stock[bc] : 0;
    });
  },

  // Lấy map barcode → tồn kho cho một ngày bất kỳ (dùng cho báo cáo)
  getStockForDate(date) {
    return this._calcStockForDate(date);
  },

  // ── Getters ──────────────────────────────────────────────
  getImports() {
    return this.transactions.filter(t => t.type === 'import');
  },

  getExports() {
    return this.transactions.filter(t => t.type === 'export');
  },

  getTodayImports() {
    const now = new Date();
    const today = `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}-${String(now.getDate()).padStart(2,'0')}`;
    return this.getImports().filter(t => t.date && t.date.startsWith(today));
  },

  getTodayExports() {
    const now = new Date();
    const today = `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}-${String(now.getDate()).padStart(2,'0')}`;
    return this.getExports().filter(t => t.date && t.date.startsWith(today));
  },

  getTotalStock() {
    return this.products.reduce((s, p) => s + (p.stock || 0), 0);
  },

  getTotalProducts() {
    return this.products.length;
  },

  getStockByCategory() {
    const result = {};
    CATEGORIES.forEach(c => { result[c.id] = { name: c.name, color: c.color, total: 0, value: 0 }; });
    this.products.forEach(p => {
      const cat = p.category in result ? p.category : 'khac';
      result[cat].total += p.stock || 0;
      result[cat].value += (p.stock || 0) * (p.price || 0);
    });
    return Object.values(result).filter(d => d.total > 0 || d.value > 0);
  },

  getLowStockProducts() {
    return this.products.filter(p => p.minStock > 0 && p.stock <= p.minStock * 1.5);
  },

  getDiscrepancy() {
    const exports = this.getExports().filter(t => t.status === 'completed');
    if (!exports.length) return { planned: 0, actual: 0, diff: 0, percent: '0.0' };
    const totalPlanned = exports.reduce((s, t) => s + t.quantity, 0);
    const totalActual  = exports.reduce((s, t) => s + (t.actualQty || t.quantity), 0);
    const ratio = totalPlanned > 0 ? totalActual / totalPlanned : 1;
    return {
      planned: totalPlanned,
      actual:  totalActual,
      diff:    totalPlanned - totalActual,
      percent: ((1 - ratio) * 100).toFixed(1),
    };
  },

  // ── Import bảng giá ──────────────────────────────────────
  importPriceList(rows) {
    let updated = 0;
    let added   = 0;
    const addedIds   = [];
    const updatedIds = [];

    rows.forEach(row => {
      const existing = this.products.find(p =>
        p.id === row.maHang || p.barcode === row.maHang
      );

      if (existing) {
        existing.price    = row.price;
        existing.cate1    = row.cate1 || existing.cate1 || '';
        existing.cate2    = row.cate2 || existing.cate2 || '';
        existing.cate3    = row.cate3 || existing.cate3 || '';
        if (!existing.name || existing.name === existing.id) existing.name = row.tenHang;
        if (!existing.category || existing.category === 'khac') existing.category = row.category;
        updatedIds.push(existing.id);
        updated++;
      } else {
        this.products.push({
          id:       row.maHang,
          name:     row.tenHang,
          barcode:  row.maHang,
          category: row.category || 'khac',
          cate1:    row.cate1 || '',
          cate2:    row.cate2 || '',
          cate3:    row.cate3 || '',
          unit:     '',
          price:    row.price,
          stock:    0,
          minStock: 0,
        });
        addedIds.push(row.maHang);
        added++;
      }
    });

    this._recalcStock();
    this.save();
    return { updated, added, addedIds, updatedIds };
  },

  // ── Booking ──────────────────────────────────────────────
  importBookings(rows) {
    const existingIds = new Set(this.bookings.map(b => b.id));
    let added = 0, skipped = 0;
    const addedIds = [];

    rows.forEach(row => {
      if (existingIds.has(row.id)) { skipped++; return; }
      this.bookings.push(row);
      existingIds.add(row.id);
      addedIds.push(row.id);
      added++;
    });

    this.bookings.sort((a, b) => new Date(a.deliveryDate) - new Date(b.deliveryDate));
    this.save();
    return { added, skipped, addedIds };
  },

  getBookings() {
    return this.bookings;
  },

  getBookingsByDate(date) {
    return this.bookings.filter(b => b.deliveryDate && b.deliveryDate.startsWith(date));
  },

  // ── FC (Forecast Control) ─────────────────────────────────
  setFC(data) {
    this.fc = data;
    this.save();
  },

  getFC() {
    return this.fc;
  },

  // ── Master Data (Mã hàng thịt cá) ────────────────────────
  setMasterData(rows) {
    this.masterData = rows;
    this.save();
  },

  getMasterData() {
    return this.masterData || [];
  },

  addTransaction(data) {
    const id = `GD${String(this.transactions.length + 1).padStart(4, '0')}`;
    const product = this.products.find(p => p.id === data.productId);
    const t = {
      id,
      type:        data.type,
      productId:   data.productId,
      productName: product ? product.name : (data.productName || ''),
      barcode:     product ? product.barcode : '',
      category:    product ? product.category : 'khac',
      quantity:    parseInt(data.quantity),
      unit:        product ? product.unit : (data.unit || ''),
      price:       product ? product.price : 0,
      total:       parseInt(data.quantity) * (product ? product.price : 0),
      date:        new Date().toISOString(),
      status:      'completed',
      supplier:    data.supplier || null,
      customer:    data.customer || null,
      note:        data.note || '',
    };
    this.transactions.unshift(t);
    this._recalcStock();
    this.save();
    return t;
  },
};
