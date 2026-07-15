// ===================================
// EXCEL-IMPORT.JS — Parse Excel files using SheetJS
// ===================================

const ExcelImport = {

  // ── Read file → workbook ─────────────────────────────────
  readFile(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = e => {
        try {
          const wb = XLSX.read(e.target.result, { type: 'array', cellDates: false });
          resolve(wb);
        } catch (err) {
          reject(new Error('Không đọc được file Excel: ' + err.message));
        }
      };
      reader.onerror = () => reject(new Error('Lỗi đọc file'));
      reader.readAsArrayBuffer(file);
    });
  },

  // ── Sheet → raw rows (array of arrays) ──────────────────
  sheetToRows(wb, sheetIndex = 0) {
    const sheetName = wb.SheetNames[sheetIndex];
    const ws = wb.Sheets[sheetName];
    return XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  },

  // ── Find column index by header name (flexible matching) ─
  findCol(headers, ...candidates) {
    for (const cand of candidates) {
      const idx = headers.findIndex(h =>
        String(h).trim().toLowerCase() === cand.toLowerCase()
      );
      if (idx >= 0) return idx;
    }
    // Partial match fallback
    for (const cand of candidates) {
      const idx = headers.findIndex(h =>
        String(h).trim().toLowerCase().includes(cand.toLowerCase())
      );
      if (idx >= 0) return idx;
    }
    return -1;
  },

  // ── Add N days to an ISO date string (UTC-safe) ──────────
  _addDays(isoStr, n) {
    const d = new Date(isoStr);
    return new Date(Date.UTC(
      d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate() + n
    )).toISOString();
  },

  // ── Parse date cell (Excel serial or string) ─────────────
  parseDate(val) {
    if (!val) return new Date().toISOString();
    // Already a string like "18/05/2026 23:00:00"
    if (typeof val === 'string') {
      // dd/mm/yyyy hh:mm:ss or dd/mm/yyyy — use UTC to avoid timezone-induced day shift
      const m = val.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
      if (m) {
        const [, d, mo, y] = m;
        return new Date(Date.UTC(parseInt(y), parseInt(mo) - 1, parseInt(d))).toISOString();
      }
      const parsed = new Date(val);
      if (!isNaN(parsed)) return parsed.toISOString();
    }
    // Excel date serial — use UTC to avoid timezone-induced day shift
    if (typeof val === 'number') {
      const date = XLSX.SSF.parse_date_code(val);
      if (date) {
        return new Date(Date.UTC(date.y, date.m - 1, date.d)).toISOString();
      }
    }
    return new Date().toISOString();
  },

  // ── Map Trạng thái → status ──────────────────────────────
  // NOTE: do NOT use s.includes('huy') — "chuyển" contains "huy" as substring
  mapStatus(raw) {
    const s = String(raw).trim().toLowerCase();
    if (s.includes('hủy') || s.includes('cancel')) return 'cancelled';
    if (
      s.includes('đã nhận') || s.includes('hoàn thành') ||
      s.includes('xong') || s.includes('done') || s.includes('complete')
    ) return 'completed';
    // "Đang chuyển", "Chờ xử lý", "Chờ nhận hàng" → pending
    return 'pending';
  },

  // ── Auto-detect product category from name ────────────────
  detectCategory(name) {
    const n = (name || '').toLowerCase();
    if (n.includes('heo') || n.includes('lợn') || n.includes('pork')) return 'thit-heo';
    if (n.includes('bò') || n.includes('beef')) return 'thit-bo';
    if (n.includes('gà') || n.includes('vịt') || n.includes('chicken') || n.includes('duck')) return 'thit-ga';
    if (n.includes('tôm') || n.includes('cua') || n.includes('mực') || n.includes('shrimp')) return 'tom-cua';
    if (n.includes('cá') || n.includes('fish') || n.includes('salmon') || n.includes('basa')) return 'ca';
    if (n.includes('nghêu') || n.includes('ốc') || n.includes('sò') || n.includes('hải sản')) return 'hai-san';
    if (n.includes('rau') || n.includes('củ') || n.includes('quả') || n.includes('khoai') || n.includes('cải') || n.includes('hành')) return 'rau-cu';
    if (n.includes('trứng') || n.includes('sữa') || n.includes('egg') || n.includes('milk')) return 'trung-sua';
    return 'khac';
  },

  // ════════════════════════════════════════════════════════
  // NHẬP HÀNG — DanhSachChiTietNhapHang format
  // Columns: Mã PO, Trạng thái, Ngày nhận hàng, Tên nhà cung cấp,
  //          Mã nội bộ, Mã hàng, Tên hàng, ĐVT, Giá PO, Số lượng PR (thực nhận)
  // ════════════════════════════════════════════════════════
  parseImport(wb) {
    const rows = this.sheetToRows(wb);
    if (rows.length < 2) throw new Error('File rỗng hoặc không có dữ liệu');

    const headers = rows[0].map(h => String(h).trim());

    // Locate columns
    const COL = {
      maPO:       this.findCol(headers, 'Mã PO', 'MaPO', 'Mã po'),
      maPR:       this.findCol(headers, 'Mã PR', 'MaPR'),
      trangThai:  this.findCol(headers, 'Trạng thái', 'Trang thai', 'Status'),
      chiNhanh:   this.findCol(headers, 'Chi nhánh', 'Chi nhanh', 'Branch'),
      ngayNhanNCC: this.findCol(headers, 'Ngày giao hàng NCC xác nhận', 'Ngày NCC xác nhận', 'Ngày xác nhận NCC', 'Ngay giao hang NCC xac nhan', 'Ngay NCC xac nhan', 'Ngày xác nhận'),
      maNCC:      this.findCol(headers, 'Mã nhà cung cấp', 'Mã NCC'),
      tenNCC:     this.findCol(headers, 'Tên nhà cung cấp', 'Nha cung cap', 'Supplier'),
      ghiChu:     this.findCol(headers, 'Ghi chú', 'Ghi chu', 'Note'),
      maNB:       this.findCol(headers, 'Mã nội bộ', 'Ma noi bo', 'Internal code'),
      maHang:     this.findCol(headers, 'Mã hàng', 'Ma hang', 'Barcode', 'SKU'),
      tenHang:    this.findCol(headers, 'Tên hàng', 'Ten hang', 'Product', 'Sản phẩm'),
      dvt:        this.findCol(headers, 'ĐVT', 'Đơn vị', 'Unit', 'DVT'),
      giaPO:      this.findCol(headers, 'Giá PO', 'Gia PO', 'Đơn giá', 'Don gia', 'Price'),
      slThucNhan: this.findCol(headers, 'Số lượng PR (thực nhận)', 'Số lượng PR', 'SL thực nhận', 'Thực nhận', 'Qty received', 'Quantity'),
      slPO:       this.findCol(headers, 'Số lượng PO', 'SL PO', 'Qty PO'),
    };

    // Validate required columns
    const missing = [];
    if (COL.tenHang < 0) missing.push('Tên hàng');
    if (COL.ngayNhanNCC < 0) missing.push('Ngày NCC xác nhận');
    if (missing.length) throw new Error('File thiếu cột: ' + missing.join(', '));

    const result = [];
    const seen = new Set();

    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r || r.every(c => c === '')) continue;

      const tenHang = String(r[COL.tenHang] || '').trim();
      if (!tenHang) continue;

      const maPO   = COL.maPO >= 0   ? String(r[COL.maPO]   || '').trim() : '';
      const maNB   = COL.maNB >= 0   ? String(r[COL.maNB]   || '').trim() : '';
      const maHang = COL.maHang >= 0 ? String(r[COL.maHang] || '').trim() : '';

      // Build unique ID: PO + internal code + row index fallback
      const productId = maNB || maHang || `SP-${i}`;
      const rawId     = (maPO || `ROW${i}`) + '-' + productId;
      // Deduplicate
      let id = rawId;
      let suffix = 1;
      while (seen.has(id)) { id = rawId + '-' + (++suffix); }
      seen.add(id);

      const qty    = parseFloat(r[COL.slThucNhan] || 0) || 0;
      const price  = parseFloat(r[COL.giaPO] || 0) || 0;

      result.push({
        id,
        type:         'import',
        productId,
        productName:  tenHang,
        barcode:      maHang,
        category:     this.detectCategory(tenHang),
        quantity:     qty,
        unit:         COL.dvt >= 0 ? String(r[COL.dvt] || '').trim() : '',
        price,
        total:        qty * price,
        date:         this._addDays(this.parseDate(r[COL.ngayNhanNCC]), 1),
        status:       this.mapStatus(COL.trangThai >= 0 ? r[COL.trangThai] : ''),
        supplier:     COL.tenNCC >= 0 ? String(r[COL.tenNCC] || '').trim() : '',
        supplierCode: COL.maNCC >= 0 ? String(r[COL.maNCC] || '').trim() : '',
        customer:     null,
        chiNhanh:     COL.chiNhanh >= 0 ? String(r[COL.chiNhanh] || '').trim() : '',
        maPO,
        note:         COL.ghiChu >= 0 ? String(r[COL.ghiChu] || '').trim() : '',
      });
    }

    if (!result.length) throw new Error('Không tìm thấy dữ liệu hợp lệ trong file');
    return result;
  },

  // ════════════════════════════════════════════════════════
  // XUẤT HÀNG — Phiếu chuyển hàng (transfer_*.xlsx)
  // Columns:
  //   Ngày chuyển hàng, Chi nhánh chuyển, Chi nhánh nhận
  //   Mã hàng, Tên hàng, Đơn vị tính
  //   Số lượng chuyển, Số lượng nhận
  //   Mã chuyển hàng, Đã nhận hàng, Trạng thái
  //   Ghi chú chuyển (phiếu)
  // ════════════════════════════════════════════════════════
  parseExport(wb) {
    const rows = this.sheetToRows(wb);
    if (rows.length < 2) throw new Error('File rỗng hoặc không có dữ liệu');

    const headers = rows[0].map(h => String(h).trim());

    const COL = {
      ngayXuat:    this.findCol(headers, 'Ngày chuyển hàng', 'Ngày xuất', 'Ngày giao'),
      chiNhanhXuat: this.findCol(headers, 'Chi nhánh chuyển', 'Kho chuyển', 'Kho xuất'),
      chiNhanhNhan: this.findCol(headers, 'Chi nhánh nhận', 'Kho nhận', 'Điểm nhận'),
      maHang:      this.findCol(headers, 'Mã hàng', 'Mã nội bộ', 'Barcode', 'SKU'),
      tenHang:     this.findCol(headers, 'Tên hàng', 'Sản phẩm'),
      dvt:         this.findCol(headers, 'Đơn vị tính', 'ĐVT', 'Đơn vị', 'Unit'),
      slChuyen:    this.findCol(headers, 'Số lượng chuyển', 'SL chuyển', 'Số lượng xuất', 'Số lượng'),
      slNhan:      this.findCol(headers, 'Số lượng nhận', 'SL nhận', 'SL thực nhận'),
      maPhieu:     this.findCol(headers, 'Mã chuyển hàng', 'Mã phiếu', 'Mã phiếu xuất', 'Mã SO'),
      daNhan:      this.findCol(headers, 'Đã nhận hàng'),
      trangThai:   this.findCol(headers, 'Trạng thái', 'Status'),
      ghiChu:      this.findCol(headers, 'Ghi chú chuyển (phiếu)', 'Ghi chú chuyển', 'Ghi chú', 'Note'),
    };

    const missing = [];
    if (COL.tenHang < 0) missing.push('Tên hàng');
    if (COL.slChuyen < 0) missing.push('Số lượng chuyển');
    if (missing.length) throw new Error('File thiếu cột bắt buộc: ' + missing.join(', '));

    const result = [];
    const seen = new Set();

    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r || r.every(c => c === '')) continue;

      const tenHang = String(r[COL.tenHang] || '').trim();
      if (!tenHang) continue;

      const maHang  = COL.maHang >= 0  ? String(r[COL.maHang]  || '').trim() : '';
      const maPhieu = COL.maPhieu >= 0 ? String(r[COL.maPhieu] || '').trim() : '';

      const productId = maHang || `SP-${i}`;
      const rawId     = (maPhieu || `TRANS${i}`) + '-' + productId;
      let id = rawId;
      let suffix = 1;
      while (seen.has(id)) { id = rawId + '-' + (++suffix); }
      seen.add(id);

      const slChuyen = parseFloat(r[COL.slChuyen] || 0) || 0;
      const slNhan   = COL.slNhan >= 0 ? parseFloat(r[COL.slNhan] || 0) : 0;

      // Quantity: use actual received if positive, else planned transfer qty
      const qty = slNhan > 0 ? slNhan : Math.abs(slChuyen);

      // Status: prioritise "Đã nhận hàng" column then "Trạng thái"
      let status;
      const daNhanVal = COL.daNhan >= 0 ? String(r[COL.daNhan] || '').trim().toLowerCase() : '';
      if (daNhanVal === 'có' || daNhanVal === 'co' || daNhanVal === 'yes') {
        status = 'completed';
      } else {
        status = this.mapStatus(COL.trangThai >= 0 ? r[COL.trangThai] : '');
      }

      result.push({
        id,
        type:         'export',
        productId,
        productName:  tenHang,
        barcode:      maHang,
        category:     this.detectCategory(tenHang),
        quantity:     qty,
        transferQty:  Math.abs(slChuyen),
        actualQty:    slNhan > 0 ? slNhan : null,
        unit:         COL.dvt >= 0 ? String(r[COL.dvt] || '').trim() : '',
        price:        0,
        total:        0,
        date:         this.parseDate(COL.ngayXuat >= 0 ? r[COL.ngayXuat] : null),
        status,
        supplier:     null,
        customer:     COL.chiNhanhNhan >= 0 ? String(r[COL.chiNhanhNhan] || '').trim() : '',
        chiNhanh:     COL.chiNhanhXuat >= 0 ? String(r[COL.chiNhanhXuat] || '').trim() : '',
        maPO:         maPhieu,
        note:         COL.ghiChu >= 0 ? String(r[COL.ghiChu] || '').trim() : '',
      });
    }

    if (!result.length) throw new Error('Không tìm thấy dữ liệu hợp lệ trong file');
    return result;
  },

  // ════════════════════════════════════════════════════════
  // BẢNG GIÁ — bang-gia_*.xlsx
  // Columns: Mã hàng, Tên hàng, Cate level 1, Cate level 2, Cate level 3, Bảng giá chung
  // ════════════════════════════════════════════════════════
  parsePriceList(wb) {
    const rows = this.sheetToRows(wb);
    if (rows.length < 2) throw new Error('File rỗng hoặc không có dữ liệu');

    const headers = rows[0].map(h => String(h).trim());

    const COL = {
      maHang:  this.findCol(headers, 'Mã hàng', 'Mã nội bộ', 'SKU', 'Barcode', 'Product code'),
      tenHang: this.findCol(headers, 'Tên hàng', 'Tên sản phẩm', 'Sản phẩm', 'Product name'),
      cate1:   this.findCol(headers, 'Cate level 1', 'Danh mục 1', 'Category 1', 'Category level 1'),
      cate2:   this.findCol(headers, 'Cate level 2', 'Danh mục 2', 'Category 2', 'Category level 2'),
      cate3:   this.findCol(headers, 'Cate level 3', 'Danh mục 3', 'Category 3', 'Category level 3'),
      bangGia: this.findCol(headers, 'Bảng giá chung', 'Giá chung', 'Giá bán', 'Đơn giá', 'Giá', 'Price'),
    };

    const missing = [];
    if (COL.maHang  < 0) missing.push('Mã hàng');
    if (COL.bangGia < 0) missing.push('Bảng giá chung / Giá bán');
    if (missing.length) throw new Error('File thiếu cột bắt buộc: ' + missing.join(', '));

    const result = [];
    const seen   = new Set();

    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r || r.every(c => c === '')) continue;

      const maHang = String(r[COL.maHang] || '').trim();
      if (!maHang) continue;
      if (seen.has(maHang)) continue; // keep first occurrence (lowest row = most specific)
      seen.add(maHang);

      const tenHang = COL.tenHang >= 0 ? String(r[COL.tenHang] || '').trim() : '';
      const cate1   = COL.cate1 >= 0   ? String(r[COL.cate1]   || '').trim() : '';
      const cate2   = COL.cate2 >= 0   ? String(r[COL.cate2]   || '').trim() : '';
      const cate3   = COL.cate3 >= 0   ? String(r[COL.cate3]   || '').trim() : '';
      const price   = parseFloat(r[COL.bangGia] || 0) || 0;

      result.push({
        maHang,
        tenHang,
        cate1,
        cate2,
        cate3,
        category: this._mapCateLevel(cate1, cate2, cate3, tenHang),
        price,
      });
    }

    if (!result.length) throw new Error('Không tìm thấy dữ liệu hợp lệ trong file');
    return result;
  },

  // ── Map Cate level → internal category ──────────────────
  _mapCateLevel(c1, c2, c3, name) {
    const all = [c1, c2, c3, name].join(' ').toLowerCase();
    // Seafood & fish
    if (all.includes('shrimp') || all.includes('tôm') || all.includes('tom ')
        || all.includes('crab') || all.includes('cua') || all.includes('squid')
        || all.includes('mực') || all.includes('muc ') || all.includes('lobster'))
      return 'tom-cua';
    if (all.includes('fish') || all.includes('seafood') || all.includes('cá ')
        || all.includes('ca ') || all.includes('hải sản') || all.includes('hai san')
        || all.includes('live fish') || all.includes('fresh fish'))
      return 'ca';
    if (all.includes('clam') || all.includes('nghêu') || all.includes('ngheu')
        || all.includes('oyster') || all.includes('sò') || all.includes('so ')
        || all.includes('snail') || all.includes('ốc') || all.includes('scallop'))
      return 'hai-san';
    // Meat
    if (all.includes('pork') || all.includes('heo') || all.includes('lợn') || all.includes('lon '))
      return 'thit-heo';
    if (all.includes('beef') || all.includes('bò') || all.includes('bo ') || all.includes('veal'))
      return 'thit-bo';
    if (all.includes('chicken') || all.includes('gà') || all.includes('ga ')
        || all.includes('duck') || all.includes('vịt') || all.includes('vit ')
        || all.includes('poultry'))
      return 'thit-ga';
    // Produce
    if (all.includes('vegetable') || all.includes('produce') || all.includes('rau')
        || all.includes('củ') || all.includes('cu ') || all.includes('quả')
        || all.includes('khoai') || all.includes('hành') || all.includes('cải'))
      return 'rau-cu';
    // Egg & dairy
    if (all.includes('egg') || all.includes('trứng') || all.includes('trung ')
        || all.includes('milk') || all.includes('sữa') || all.includes('sua ')
        || all.includes('dairy'))
      return 'trung-sua';
    // Fall back to name-based detection
    return this.detectCategory(name);
  },

  // ── Preview for price list ───────────────────────────────
  previewPriceList(rows, existingProducts) {
    const existingIds = new Set(existingProducts.map(p => p.id));
    const updated = rows.filter(r => existingIds.has(r.maHang)).length;
    const added   = rows.length - updated;
    const minPrice = rows.reduce((m, r) => Math.min(m, r.price), Infinity);
    const maxPrice = rows.reduce((m, r) => Math.max(m, r.price), 0);
    return { total: rows.length, updated, added, minPrice, maxPrice };
  },

  // ════════════════════════════════════════════════════════
  // BOOKING — MD 1_KFM_BOOKING_*.xlsx
  // Columns: Delivery Date, Region, Site, Customer Code, BillTo,
  //   SupplierID, Supplier Name, PO, PO Detail,
  //   ItemID (Code Spec), Item Name, Quantity, UOM,
  //   StoreID, Store Name, Type, Category,
  //   Code Sản phẩm, IsProvince, Zone, Notes, Weight, WeightOfPackage
  // ════════════════════════════════════════════════════════
  parseBooking(wb) {
    const rows = this.sheetToRows(wb);
    if (rows.length < 2) throw new Error('File rỗng hoặc không có dữ liệu');

    const headers = rows[0].map(h => String(h).trim());

    const COL = {
      deliveryDate:     this.findCol(headers, 'Delivery Date', 'Ngày giao', 'Delivery date'),
      region:           this.findCol(headers, 'Region', 'Khu vực'),
      site:             this.findCol(headers, 'Site', 'Chi nhánh'),
      customerCode:     this.findCol(headers, 'Customer Code', 'BillTo', 'Customer code'),
      supplierId:       this.findCol(headers, 'SupplierID', 'Supplier ID', 'Mã NCC'),
      supplierName:     this.findCol(headers, 'Supplier Name', 'Tên NCC', 'Tên nhà cung cấp'),
      po:               this.findCol(headers, 'PO'),
      poDetail:         this.findCol(headers, 'PO Detail', 'PO detail'),
      itemId:           this.findCol(headers, 'ItemID (Code Spec)', 'ItemID', 'Item ID', 'Mã hàng'),
      itemName:         this.findCol(headers, 'Item Name', 'Tên hàng', 'Sản phẩm'),
      quantity:         this.findCol(headers, 'Quantity', 'Số lượng', 'SL'),
      uom:              this.findCol(headers, 'UOM', 'ĐVT', 'Đơn vị'),
      storeId:          this.findCol(headers, 'StoreID', 'Store ID', 'Mã cửa hàng'),
      storeName:        this.findCol(headers, 'Store Name', 'Tên cửa hàng'),
      type:             this.findCol(headers, 'Type', 'Loại'),
      tempCategory:     this.findCol(headers, 'Category', 'Phân loại'),
      codeSp:           this.findCol(headers, 'Code Sản phẩm', 'Code SP'),
      isProvince:       this.findCol(headers, 'IsProvince'),
      zone:             this.findCol(headers, 'Zone'),
      notes:            this.findCol(headers, 'Notes', 'Ghi chú'),
      weight:           this.findCol(headers, 'Weight', 'Trọng lượng'),
      weightOfPackage:  this.findCol(headers, 'WeightOfPackage'),
    };

    const missing = [];
    if (COL.itemName < 0) missing.push('Item Name / Tên hàng');
    if (COL.quantity  < 0) missing.push('Quantity / Số lượng');
    if (missing.length) throw new Error('File thiếu cột bắt buộc: ' + missing.join(', '));

    const result = [];
    const seen   = new Set();

    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r || r.every(c => c === '')) continue;

      const itemName = String(r[COL.itemName] || '').trim();
      if (!itemName) continue;

      const poDetail = COL.poDetail >= 0 ? String(r[COL.poDetail] || '').trim() : '';
      const itemId   = COL.itemId   >= 0 ? String(r[COL.itemId]   || '').trim() : '';
      const po       = COL.po       >= 0 ? String(r[COL.po]       || '').trim() : '';

      const rawId = poDetail || (po + '-' + itemId + '-' + i);
      let id = 'BK-' + rawId;
      let suffix = 1;
      while (seen.has(id)) { id = 'BK-' + rawId + '-' + (++suffix); }
      seen.add(id);

      result.push({
        id,
        deliveryDate:    this.parseDate(COL.deliveryDate >= 0 ? r[COL.deliveryDate] : null),
        region:          COL.region       >= 0 ? String(r[COL.region]       || '').trim() : '',
        site:            COL.site         >= 0 ? String(r[COL.site]         || '').trim() : '',
        customerCode:    COL.customerCode >= 0 ? String(r[COL.customerCode] || '').trim() : '',
        supplierId:      COL.supplierId   >= 0 ? String(r[COL.supplierId]   || '').trim() : '',
        supplierName:    COL.supplierName >= 0 ? String(r[COL.supplierName] || '').trim() : '',
        po,
        poDetail,
        itemId,
        itemName,
        quantity:        parseFloat(r[COL.quantity] || 0) || 0,
        uom:             COL.uom          >= 0 ? String(r[COL.uom]          || '').trim() : '',
        storeId:         COL.storeId      >= 0 ? String(r[COL.storeId]      || '').trim() : '',
        storeName:       COL.storeName    >= 0 ? String(r[COL.storeName]    || '').trim() : '',
        type:            COL.type         >= 0 ? String(r[COL.type]         || '').trim() : '',
        tempCategory:    COL.tempCategory >= 0 ? String(r[COL.tempCategory] || '').trim() : '',
        codeSp:          COL.codeSp       >= 0 ? String(r[COL.codeSp]       || '').trim() : '',
        isProvince:      COL.isProvince   >= 0 ? (String(r[COL.isProvince] || '').trim() === '1') : false,
        zone:            COL.zone         >= 0 ? String(r[COL.zone]         || '').trim() : '',
        notes:           COL.notes        >= 0 ? String(r[COL.notes]        || '').trim() : '',
        weight:          COL.weight       >= 0 ? (parseFloat(r[COL.weight]          || 0) || 0) : 0,
        weightOfPackage: COL.weightOfPackage >= 0 ? (parseFloat(r[COL.weightOfPackage] || 0) || 0) : 0,
      });
    }

    if (!result.length) throw new Error('Không tìm thấy dữ liệu hợp lệ trong file');
    return result;
  },

  previewBooking(rows) {
    const pos       = new Set(rows.map(r => r.po).filter(Boolean));
    const suppliers = new Set(rows.map(r => r.supplierName).filter(Boolean));
    const stores    = new Set(rows.map(r => r.storeName).filter(Boolean));
    const dates     = rows.map(r => r.deliveryDate).filter(Boolean).sort();
    const totalQty  = rows.reduce((s, r) => s + (r.quantity || 0), 0);
    return {
      total:     rows.length,
      pos:       pos.size,
      suppliers: suppliers.size,
      stores:    stores.size,
      totalQty,
      dateFrom:  dates[0]               || null,
      dateTo:    dates[dates.length - 1] || null,
    };
  },

  // ── Preview summary (transactions) ──────────────────────
  preview(rows) {
    const byStatus = { completed: 0, pending: 0, cancelled: 0 };
    rows.forEach(r => { byStatus[r.status] = (byStatus[r.status] || 0) + 1; });
    const totalValue = rows.reduce((s, r) => s + (r.total || 0), 0);
    const products = [...new Set(rows.map(r => r.productId))].length;
    const partners = [...new Set(rows.map(r => r.supplier || r.customer).filter(Boolean))].length;
    return { total: rows.length, byStatus, totalValue, products, partners };
  },

  // ════════════════════════════════════════════════════════
  // FC — Forecast Control (In-outbound ABAMD)
  // Structure:
  //   Row 3 (idx 2):  date serials (cols 13+ = columns N→)
  //   Row 7 (idx 6):  FINAL total — monthly avg in cols B-M (idx 1-12)
  //   Rows 13-20 (idx 12-19): categories starting with "3."
  //     col A = category name, cols B-M = monthly avg, cols N→ = daily values
  // ════════════════════════════════════════════════════════
  parseFC(wb) {
    const rows = this.sheetToRows(wb);
    if (rows.length < 13) throw new Error('File không đúng định dạng FC (cần ít nhất 13 dòng)');

    const dateRow  = rows[2] || [];
    const finalRow = rows[6] || [];

    // Collect date columns (col index 13+, must be a valid Excel date serial)
    const dates          = [];
    const dateColIndices = [];
    for (let ci = 13; ci < dateRow.length; ci++) {
      const v = dateRow[ci];
      if (v && typeof v === 'number' && v > 40000) {
        const parsed = XLSX.SSF.parse_date_code(v);
        if (parsed && parsed.y >= 2020) {
          const iso = `${parsed.y}-${String(parsed.m).padStart(2,'0')}-${String(parsed.d).padStart(2,'0')}`;
          dates.push(iso);
          dateColIndices.push(ci);
        }
      }
    }
    if (!dates.length) throw new Error('Không tìm thấy cột ngày trong file FC. Kiểm tra dòng 3 (Ngày).');

    // Find all category rows (label starts with "3." in column A, first block only)
    const categories    = [];
    const catRowIndices = [];
    for (let ri = 5; ri < Math.min(rows.length, 25); ri++) {
      const label = String(rows[ri][0] || '').trim();
      if (label.startsWith('3.')) {
        categories.push(label);
        catRowIndices.push(ri);
        if (categories.length >= 10) break;
      }
    }
    if (!categories.length) throw new Error('Không tìm thấy dữ liệu ngành hàng (cột A bắt đầu bằng "3.")');

    // Daily values: values[catIdx][dateIdx]
    const values = categories.map((_, catIdx) => {
      const row = rows[catRowIndices[catIdx]] || [];
      return dateColIndices.map(ci => {
        const v = row[ci];
        return typeof v === 'number' ? Math.round(v * 100) / 100 : 0;
      });
    });

    // Monthly averages (cols B-M = indices 1-12, months 1-12)
    const monthlyAvg = categories.map((_, catIdx) => {
      const row = rows[catRowIndices[catIdx]] || [];
      return Array.from({ length: 12 }, (_, mi) => {
        const v = row[mi + 1];
        return typeof v === 'number' ? Math.round(v * 100) / 100 : 0;
      });
    });

    // FINAL total monthly avg (cols B-M)
    const totalMonthlyAvg = Array.from({ length: 12 }, (_, mi) => {
      const v = finalRow[mi + 1];
      return typeof v === 'number' ? Math.round(v * 100) / 100 : 0;
    });

    // FINAL total daily values (same date columns as category rows)
    const totalDailyValues = dateColIndices.map(ci => {
      const v = finalRow[ci];
      return typeof v === 'number' ? Math.round(v * 100) / 100 : 0;
    });

    return { categories, dates, values, monthlyAvg, totalMonthlyAvg, totalDailyValues };
  },

  previewFC(fc) {
    const totalDays = fc.dates.length;
    const yearFrom  = fc.dates[0]                    || '';
    const yearTo    = fc.dates[fc.dates.length - 1]  || '';
    const avgTotal  = Math.round(fc.totalMonthlyAvg.reduce((s, v) => s + v, 0) / 12);
    return { totalDays, categories: fc.categories.length, yearFrom, yearTo, avgTotal };
  },

  // ── Master Data — Sheet "Data mã hàng thịt cá" (Mã hàng, Tên, ĐVT, Quy cách) ──
  parseMasterData(wb) {
    const sheetNames = wb.SheetNames || [];
    let rows = null;

    for (const name of sheetNames) {
      const ws  = wb.Sheets[name];
      if (!ws) continue;
      const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      if (raw.length < 2) continue;
      const hdr = raw[0].map(c => String(c).trim());
      // Match sheet with "Mã hàng" + "Tên sản phẩm" headers (Sheet 2)
      if (hdr.includes('Mã hàng') && hdr.includes('Tên sản phẩm')) {
        rows = raw;
        break;
      }
    }

    if (!rows) throw new Error('Không tìm thấy sheet "Data mã hàng" (cần cột "Mã hàng" và "Tên sản phẩm")');

    const hdr     = rows[0].map(c => String(c).trim());
    const colCode = hdr.indexOf('Mã hàng');
    const colName = hdr.indexOf('Tên sản phẩm');
    const colUnit = hdr.indexOf('ĐVT');
    const colSpec = hdr.indexOf('Quy cách');

    const result = [];
    const ts     = Date.now();
    for (let ri = 1; ri < rows.length; ri++) {
      const r    = rows[ri];
      const code = String(r[colCode] ?? '').trim();
      if (!code) continue;
      result.push({
        id:   `md-${ts}-${ri}`,
        code,
        name: String(r[colName] ?? '').trim(),
        unit: colUnit >= 0 ? String(r[colUnit] ?? '').trim() : '',
        spec: colSpec >= 0 ? String(r[colSpec] ?? '').trim() : '',
      });
    }

    if (!result.length) throw new Error('Không tìm thấy dữ liệu hợp lệ trong file');
    return result;
  },
};
