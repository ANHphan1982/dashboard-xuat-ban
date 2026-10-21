/* =====================================================================
 * data-loader.js — ĐỌC & CHUẨN HÓA DỮ LIỆU TỪ FILE EXCEL
 * - Tìm sheet và dòng tiêu đề, nhận dạng định dạng file (EXPORT từ SAP hoặc PXK), ánh xạ cột.
 * - Chuẩn hóa ngày, tách số và đơn vị khối lượng, tạo khóa phân tích.
 * - Kiểm tra chất lượng dữ liệu (thiếu, sai định dạng, trùng lặp).
 * Không ghi đè dữ liệu gốc: giá trị gốc giữ trong record.raw.
 * ===================================================================== */
(function (root) {
  'use strict';
  const PXK = (root.PXK = root.PXK || {});
  const XLSX = root.XLSX || (typeof require === 'function' ? require('../lib/xlsx.full.min.js') : null);

  const EMPTY = '(Trống)';
  // Quy tắc gán Hình thức vận chuyển cho ô trống (áp dụng theo thứ tự, quy tắc đầu tiên khớp được dùng)
  const HT_RULES = [
    { id: 'road', value: 'Đường bộ', desc: 'Khối lượng < 50 M3/SM3', test: (r) => r.vol3 != null && r.vol3 < 50 },
    { id: 'pipe', value: 'Đường ống', desc: 'Phương tiện là "Pipeline" / "Pipe line"', test: (r) => fold(norm(r.raw.pt)).replace(/[\s\-_]/g, '') === 'pipeline' },
  ];

  /*
   * ĐỊNH DẠNG FILE ĐƯỢC HỖ TRỢ — tự nhận dạng theo dòng tiêu đề.
   * fields: trường bắt buộc (cột mặc định + tên tiêu đề có thể gặp)
   * volume: 'text' = 1 cột dạng "7,001 (M3)"; 'column' = 1 cột số, mọi sản phẩm cùng 1 đơn vị
   * optional: trường phụ (số chứng từ để kiểm tra trùng, trạng thái để lọc)
   */
  const PROFILES = [
    {
      id: 'EXPORT', name: 'Báo cáo xuất hàng SAP (EXPORT)',
      fields: {
        kho:  { label: 'Kho',                  col: 'AC', aliases: ['shipping point description'] },
        date: { label: 'Ngày xuất',            col: 'X',  aliases: ['loading date'] },
        kh:   { label: 'Khách hàng',           col: 'L',  aliases: ['tên khách hàng'] },
        ht:   { label: 'Hình thức vận chuyển', col: 'T',  aliases: ['hình thức vận chuyển'] },
        pt:   { label: 'Phương tiện',          col: 'C',  aliases: ['phương tiện'] },
        sp:   { label: 'Sản phẩm',             col: 'K',  aliases: ['mô tả sản phẩm'] },
      },
      // Mọi sản phẩm lấy cột AM (MT) để đồng nhất đơn vị
      volume: { type: 'column', col: { label: 'Khối lượng thực tế (MT)', col: 'AM', aliases: ['quantity (mt)'], unit: 'MT' } },
      optional: {
        so: { label: 'Số LXH trên SAP',      aliases: ['số lxh trên sap'] },
        tt: { label: 'Trạng thái giao hàng', aliases: ['shipment status desc.', 'shipment status desc'] },
        sm3: { label: 'Quantity (SM3)',      aliases: ['quantity (sm3)'] }, // chỉ dùng cho quy tắc HTVC "Đường bộ"
      },
    },
    {
      id: 'PXK', name: 'Phiếu xuất kho (PXK)',
      fields: {
        kho:  { label: 'Kho',                  col: 'G', aliases: ['kho'] },
        date: { label: 'Ngày xuất',            col: 'H', aliases: ['ngày phiếu xuất kho', 'ngày xuất', 'ngày xuất hàng'] },
        kh:   { label: 'Khách hàng',           col: 'J', aliases: ['đơn vị mua hàng', 'khách hàng'] },
        ht:   { label: 'Hình thức vận chuyển', col: 'K', aliases: ['hình thức vận chuyển'] },
        pt:   { label: 'Phương tiện',          col: 'L', aliases: ['phương tiện'] },
        sp:   { label: 'Sản phẩm',             col: 'M', aliases: ['sản phẩm'] },
      },
      volume: { type: 'text', col: { label: 'Khối lượng thực tế', col: 'N', aliases: ['khối lượng thực tế', 'khối lượng thực tế (m3/mt)'] } },
      optional: {
        so: { label: 'Số phiếu',         aliases: ['số phiếu'] },
        tt: { label: 'Tình trạng phiếu', aliases: ['tình trạng', 'trạng thái'] },
      },
    },
  ];
  const requiredOf = (p) => ({ ...p.fields, kl: p.volume.col });

  const norm = (v) => String(v == null ? '' : v).normalize('NFC').replace(/[\s ]+/g, ' ').trim();
  const normHeader = (v) => norm(v).toLowerCase();
  const colIndex = (letter) => XLSX.utils.decode_col(letter);

  /** Chuỗi tìm kiếm không dấu, không phân biệt hoa thường */
  function fold(s) {
    return String(s).normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'D').toLowerCase();
  }

  // ------------------------- Ngày tháng -------------------------
  function ymd(y, m, d) {
    if (!(y > 1900 && y < 2200) || m < 1 || m > 12 || d < 1 || d > 31) return null;
    const t = Date.UTC(y, m - 1, d);
    const dt = new Date(t);
    if (dt.getUTCDate() !== d) return null;
    return Math.round(t / 86400000); // số ngày kể từ 01/01/1970 (UTC)
  }

  /** Trả về số ngày (int) hoặc null. Hỗ trợ số serial Excel, Date, chuỗi DD/MM/YYYY, YYYY-MM-DD */
  function parseDate(v) {
    if (v == null || v === '') return null;
    if (typeof v === 'number') {
      const p = XLSX.SSF.parse_date_code(v);
      return p && p.y ? ymd(p.y, p.m, p.d) : null;
    }
    if (v instanceof Date) return isNaN(v) ? null : ymd(v.getFullYear(), v.getMonth() + 1, v.getDate());
    const s = String(v).trim();
    let m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})/);
    if (m) return ymd(+m[3], +m[2], +m[1]);
    m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m) return ymd(+m[1], +m[2], +m[3]);
    return null;
  }

  function dayToStr(day) {
    if (day == null) return '';
    const d = new Date(day * 86400000);
    const p = (n) => String(n).padStart(2, '0');
    return p(d.getUTCDate()) + '/' + p(d.getUTCMonth() + 1) + '/' + d.getUTCFullYear();
  }

  // ------------------------- Khối lượng -------------------------
  function normUnit(u) {
    const s = norm(u).toUpperCase().replace('³', '3');
    if (s === 'M3' || s === 'M 3') return 'M3';
    if (s === 'MT' || s === 'TẤN' || s === 'TAN' || s === 'T') return 'MT';
    return s || null;
  }

  /** Chuỗi số theo định dạng VN ("4.300,094") hoặc quốc tế ("4,300.094") -> Number */
  function parseNumber(s) {
    s = s.replace(/\s/g, '');
    if (!/^[-+]?[\d.,]*\d[\d.,]*$/.test(s)) return null;
    const lastDot = s.lastIndexOf('.'), lastComma = s.lastIndexOf(',');
    let t;
    if (lastDot >= 0 && lastComma >= 0) {
      t = lastComma > lastDot ? s.replace(/\./g, '').replace(',', '.') : s.replace(/,/g, '');
    } else if (lastComma >= 0) {
      t = (s.match(/,/g).length > 1) ? s.replace(/,/g, '') : s.replace(',', '.');
    } else if (lastDot >= 0) {
      // Chỉ có dấu chấm: theo quy ước VN, "1.234" hoặc "1.234.567" là phân cách hàng nghìn
      t = /^[-+]?\d{1,3}(\.\d{3})+$/.test(s) ? s.replace(/\./g, '') : s;
    } else t = s;
    const n = Number(t);
    return isFinite(n) ? n : null;
  }

  /** Tách khối lượng: { value, unit, error } */
  function parseVolume(v) {
    if (v == null || v === '') return { value: null, unit: null, error: 'Thiếu khối lượng' };
    if (typeof v === 'number') return { value: v, unit: null, error: null };
    let s = norm(v), unit = null;
    let m = s.match(/\(\s*([^)]*?)\s*\)\s*$/);
    if (m) { unit = normUnit(m[1]); s = s.slice(0, m.index).trim(); }
    else if ((m = s.match(/\s*(m3|m³|mt|tấn)\s*$/i))) { unit = normUnit(m[1]); s = s.slice(0, m.index).trim(); }
    if (s === '') return { value: null, unit, error: 'Thiếu khối lượng' };
    const n = parseNumber(s);
    if (n == null) return { value: null, unit, error: 'Khối lượng sai định dạng' };
    return { value: n, unit, error: null };
  }

  /** Quy tắc nghiệp vụ: LPG và Propylene tính theo MT, các sản phẩm còn lại tính theo thể tích (M3/SM3) */
  const MT_PRODUCTS = /\bLPG\b|propylene/i;
  function expectedUnit(product) {
    if (!product) return null;
    return MT_PRODUCTS.test(product) ? 'MT' : 'M3';
  }

  // ------------------------- Cấu trúc file -------------------------
  /** Tìm dòng tiêu đề và định dạng phù hợp nhất (khớp nhiều trường bắt buộc nhất, tối thiểu 3) */
  function findHeaderRow(rows) {
    const max = Math.min(rows.length, 40);
    let best = null;
    for (let i = 0; i < max; i++) {
      const cells = new Set((rows[i] || []).map(normHeader));
      for (const p of PROFILES) {
        const req = Object.values(requiredOf(p));
        const hits = req.filter((f) => f.aliases.some((a) => cells.has(a))).length;
        if (hits >= 3 && (!best || hits / req.length > best.score)) best = { idx: i, profile: p, score: hits / req.length };
      }
      if (best && best.score === 1) break;
    }
    return best;
  }

  function mapColumns(header, startCol, profile) {
    const FIELDS = requiredOf(profile), OPTIONAL = profile.optional;
    const map = {}, notes = [], missing = [];
    const find = (aliases) => {
      for (const a of aliases) {
        const idx = header.findIndex((h) => normHeader(h) === a);
        if (idx >= 0) return idx;
      }
      return -1;
    };
    for (const [k, f] of Object.entries(FIELDS)) {
      const fixed = colIndex(f.col) - startCol;
      if (fixed >= 0 && f.aliases.includes(normHeader(header[fixed]))) { map[k] = fixed; continue; }
      const idx = find(f.aliases);
      if (idx >= 0) {
        map[k] = idx;
        notes.push(`"${f.label}" không nằm ở cột ${f.col}, đã nhận theo tên tiêu đề tại cột ${XLSX.utils.encode_col(idx + startCol)}.`);
      } else missing.push(`${f.label} (cột ${f.col})`);
    }
    for (const [k, f] of Object.entries(OPTIONAL)) {
      const idx = find(f.aliases);
      map[k] = idx;
    }
    return { map, notes, missing };
  }

  class StructureError extends Error {}

  /**
   * Đọc ArrayBuffer của file Excel và trả về bộ dữ liệu đã chuẩn hóa.
   * @returns {{records:Array, meta:Object, issues:Object}}
   */
  function parseWorkbook(buffer, fileName) {
    if (!XLSX) throw new Error('Thiếu thư viện đọc Excel (lib/xlsx.full.min.js).');
    let wb;
    try {
      wb = XLSX.read(buffer, { type: 'array', cellDates: false, dense: true });
    } catch (e) {
      throw new StructureError('Không đọc được file. Hãy chắc chắn đây là file Excel (.xlsx/.xls) hợp lệ. Chi tiết: ' + e.message);
    }

    // Ưu tiên sheet "Data", sau đó các sheet khác có dòng tiêu đề phù hợp
    const names = [...wb.SheetNames].sort((a, b) => (b.toLowerCase() === 'data') - (a.toLowerCase() === 'data'));
    let chosen = null;
    for (const name of names) {
      const ws = wb.Sheets[name];
      if (!ws || !ws['!ref']) continue;
      const range = XLSX.utils.decode_range(ws['!ref']);
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true });
      const h = findHeaderRow(rows);
      if (h) { chosen = { name, rows, headerIdx: h.idx, profile: h.profile, range }; break; }
    }
    if (!chosen) {
      throw new StructureError('Không nhận dạng được dòng tiêu đề. File cần có các cột như "Loading date / Mô tả sản phẩm / Tên khách hàng / Quantity (MT)" (file EXPORT) hoặc "Ngày phiếu xuất kho / Sản phẩm / Khối lượng thực tế" (file PXK). Sheet có trong file: ' + wb.SheetNames.join(', '));
    }

    const { rows, headerIdx, range, profile } = chosen;
    const header = rows[headerIdx] || [];
    const { map, notes, missing } = mapColumns(header, range.s.c, profile);
    if (missing.length) {
      throw new StructureError('File (định dạng ' + profile.name + ') thiếu các cột bắt buộc: ' + missing.join(', ') + '. Dòng tiêu đề tìm thấy ở dòng ' + (range.s.r + headerIdx + 1) + ' của sheet "' + chosen.name + '".');
    }

    // Gom biến thể tên (khác khoảng trắng/hoa thường) về 1 khóa phân tích, hiển thị biến thể phổ biến nhất
    const variants = { kho: new Map(), kh: new Map(), ht: new Map(), pt: new Map(), sp: new Map(), tt: new Map() };
    const keyOf = (dim, v) => {
      const c = norm(v);
      if (!c) return EMPTY;
      const k = c.toLocaleUpperCase('vi');
      let m = variants[dim].get(k);
      if (!m) variants[dim].set(k, (m = new Map()));
      m.set(c, (m.get(c) || 0) + 1);
      return k;
    };

    const records = [];
    let blank = 0, summaryRows = 0;
    const isEmpty = (v) => v == null || norm(v) === '';
    for (let i = headerIdx + 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r || r.every(isEmpty)) { blank++; continue; }
      const get = (k) => (map[k] >= 0 ? r[map[k]] : null);
      // Dòng tổng cộng/ghi chú cuối bảng: không có ngày, khách hàng và sản phẩm → bỏ qua
      if (isEmpty(get('date')) && isEmpty(get('kh')) && isEmpty(get('sp'))) { summaryRows++; continue; }
      const raw = { kho: get('kho'), date: get('date'), kh: get('kh'), ht: get('ht'), pt: get('pt'), sp: get('sp'), so: get('so'), tt: get('tt') };
      raw.kl = get('kl');
      const volCol = profile.volume.type === 'column';
      if (volCol) raw.klUnit = profile.volume.col.unit;
      const rec = {
        id: records.length,
        row: range.s.r + i + 1, // số dòng trên Excel
        raw,
        day: parseDate(raw.date),
        kho: keyOf('kho', raw.kho), kh: keyOf('kh', raw.kh), ht: keyOf('ht', raw.ht),
        pt: keyOf('pt', raw.pt), sp: keyOf('sp', raw.sp), tt: keyOf('tt', raw.tt),
        so: norm(raw.so),
        flags: {},
      };
      const vol = parseVolume(raw.kl);
      if (volCol && !vol.unit) vol.unit = raw.klUnit; // file EXPORT: đơn vị xác định theo cột (MT)
      const exp = expectedUnit(rec.sp === EMPTY ? '' : rec.sp);
      rec.vol = vol.value;
      rec.unit = vol.unit;
      if (vol.error) rec.flags.volError = vol.error;
      if (!rec.unit) {
        if (exp) { rec.unit = exp; rec.flags.unitInferred = true; }
        else if (rec.vol != null) { rec.flags.volError = 'Không xác định được đơn vị'; }
      } else if (!volCol && exp && rec.unit !== exp) rec.flags.unitMismatch = true;
      rec.volOk = rec.vol != null && !!rec.unit && !rec.flags.volError;
      if (!rec.volOk) rec.vol = rec.flags.volError ? null : rec.vol;
      if (rec.day == null) rec.flags.dateError = (raw.date == null || raw.date === '') ? 'Thiếu ngày' : 'Ngày sai định dạng';
      if (rec.volOk && rec.vol < 0) rec.flags.negative = true;
      // Khối lượng thể tích (M3/SM3) của sản phẩm không tính theo MT — dùng cho quy tắc HTVC "Đường bộ"
      rec.vol3 = volCol
        ? (exp !== 'MT' ? parseVolume(get('sm3')).value : null)
        : (rec.volOk && rec.unit !== 'MT' ? rec.vol : null);
      // Quy tắc nghiệp vụ gán HTVC cho ô trống (giá trị gốc vẫn giữ trong raw.ht)
      if (rec.ht === EMPTY) {
        const rule = HT_RULES.find((x) => x.test(rec));
        if (rule) { rec.ht = keyOf('ht', rule.value); rec.flags.htDefault = rule; }
      }
      records.push(rec);
    }

    // Đổi khóa -> nhãn hiển thị
    const labels = {};
    for (const [dim, m] of Object.entries(variants)) {
      const L = (labels[dim] = new Map([[EMPTY, EMPTY]]));
      for (const [k, vs] of m) {
        let best = null, bc = -1;
        for (const [v, c] of vs) if (c > bc) { best = v; bc = c; }
        L.set(k, best);
      }
    }
    let whitespaceFixed = 0;
    for (const rec of records) {
      for (const dim of ['kho', 'kh', 'ht', 'pt', 'sp', 'tt']) rec[dim] = labels[dim].get(rec[dim]);
      for (const dim of ['kho', 'kh', 'pt', 'sp']) {
        const v = rec.raw[dim];
        if (v != null && String(v) !== norm(v)) { whitespaceFixed++; break; }
      }
      rec.dateStr = dayToStr(rec.day);
      rec.search = fold([rec.dateStr, rec.kho, rec.kh, rec.ht, rec.pt, rec.sp, rec.so, rec.tt, rec.raw.kl].join(' | '));
    }

    const issues = checkQuality(records, variants);
    issues.whitespaceFixed = whitespaceFixed;
    issues.blankRows = blank;
    issues.summaryRows = summaryRows;
    issues.zeroVol = records.filter((r) => r.volOk && r.vol === 0).map((r) => r.id);

    const days = records.map((r) => r.day).filter((d) => d != null);
    const meta = {
      fileName: fileName || '',
      sheet: chosen.name,
      sheetNames: wb.SheetNames,
      headerRow: range.s.r + headerIdx + 1,
      profileId: profile.id,
      profileName: profile.name,
      // Danh sách cột đã ánh xạ để hiển thị: [{label, col}]
      columnList: Object.entries(requiredOf(profile)).map(([k, f]) => ({ label: f.label, col: XLSX.utils.encode_col(map[k] + range.s.c) })),
      labels: { so: profile.optional.so.label, tt: profile.optional.tt.label },
      hasStatus: map.tt >= 0,
      hasVoucher: map.so >= 0,
      notes,
      minDay: days.length ? Math.min(...days) : null,
      maxDay: days.length ? Math.max(...days) : null,
      // Đơn vị thể tích trước, MT sau
      units: [...new Set(records.filter((r) => r.volOk).map((r) => r.unit))].sort((a, b) => (a === 'MT') - (b === 'MT') || a.localeCompare(b)),
      unitRule: profile.volume.type === 'column'
        ? `Tất cả sản phẩm = ${profile.volume.col.unit} (cột ${profile.volume.col.col})`
        : 'LPG, Propylene = MT; sản phẩm khác = M3',
      loadedAt: new Date(),
    };
    return { records, meta, issues };
  }

  function checkQuality(records, variants) {
    const by = (pred) => records.filter(pred).map((r) => r.id);
    const groupDup = (keyFn) => {
      const m = new Map();
      for (const r of records) {
        const k = keyFn(r);
        if (k == null) continue;
        (m.get(k) || m.set(k, []).get(k)).push(r.id);
      }
      const ids = [];
      let groups = 0;
      for (const arr of m.values()) if (arr.length > 1) { groups++; ids.push(...arr); }
      return { ids, groups };
    };
    const dupVoucher = groupDup((r) => (r.so ? r.so : null));
    // Có kèm số chứng từ: 1 xe nhận nhiều lần cùng ngày/sản phẩm/khối lượng nhưng khác số lệnh là các chuyến hợp lệ
    const dupRow = groupDup((r) => [r.day, r.kho, r.kh, r.ht, r.pt, r.sp, norm(r.raw.kl), r.so].join('|'));
    for (const id of dupVoucher.ids) records[id].flags.dupVoucher = true;
    for (const id of dupRow.ids) records[id].flags.dupRow = true;

    const merged = {};
    for (const [dim, m] of Object.entries(variants)) {
      merged[dim] = [...m.values()].filter((vs) => vs.size > 1).map((vs) => [...vs.keys()]);
    }

    return {
      dateMissing: by((r) => r.flags.dateError),
      volInvalid: by((r) => !r.volOk),
      unitInferred: by((r) => r.flags.unitInferred),
      unitMismatch: by((r) => r.flags.unitMismatch),
      negative: by((r) => r.flags.negative),
      missingHt: by((r) => r.ht === EMPTY),
      htDefault: by((r) => r.flags.htDefault),
      htRule: Object.fromEntries(HT_RULES.map((x) => [x.id, by((r) => r.flags.htDefault === x)])),
      missingPt: by((r) => r.pt === EMPTY),
      missingSp: by((r) => r.sp === EMPTY),
      missingKh: by((r) => r.kh === EMPTY),
      missingKho: by((r) => r.kho === EMPTY),
      dupVoucher: dupVoucher.ids, dupVoucherGroups: dupVoucher.groups,
      dupRow: dupRow.ids, dupRowGroups: dupRow.groups,
      merged,
    };
  }

  PXK.loader = { parseWorkbook, parseVolume, parseDate, parseNumber, dayToStr, fold, expectedUnit, StructureError, PROFILES, requiredOf, EMPTY, HT_RULES };
})(typeof window !== 'undefined' ? window : globalThis);
