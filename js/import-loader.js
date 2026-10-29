/* =====================================================================
 * import-loader.js — ĐỌC & CHUẨN HÓA FILE NHẬP DẦU THÔ / NGUYÊN LIỆU (IMPORT.xlsx, SAP)
 * - Tìm sheet và dòng tiêu đề, ánh xạ cột: Ngày nhập (N), Dầu thô/Nguyên liệu (D), Số lượng (I, MT).
 * - Loại dòng dầu thô nhập trùng (cùng tên cột D + số lượng cột L), giữ dòng đầu tiên.
 * - Kiểm tra chất lượng dữ liệu (thiếu/sai ngày, số lượng, đơn vị, trùng chứng từ).
 * Dùng chung tiện ích ngày/số của data-loader.js. Giá trị gốc giữ trong record.raw.
 * ===================================================================== */
(function (root) {
  'use strict';
  const PXK = (root.PXK = root.PXK || {});
  const XLSX = root.XLSX || (typeof require === 'function' ? require('../lib/xlsx.full.min.js') : null);
  const L = PXK.loader;

  const EMPTY = L.EMPTY;
  const UNIT = 'MT';

  /* Trường bắt buộc (cột mặc định + tên tiêu đề có thể gặp) và trường phụ */
  const FIELDS = {
    date: { label: 'Ngày nhập',           col: 'N', aliases: ['posting date', 'ngày nhập'] },
    sp:   { label: 'Dầu thô/Nguyên liệu', col: 'D', aliases: ['material description', 'dầu thô/nguyên liệu', 'nguyên liệu'] },
    kl:   { label: 'Số lượng (MT)',        col: 'I', aliases: ['quantity', 'số lượng'] },
  };
  const OPTIONAL = {
    doc:  { label: 'Số chứng từ (Material Document)', aliases: ['material document'] },
    item: { label: 'Dòng chứng từ',                   aliases: ['material doc.item'] },
    mat:  { label: 'Mã vật tư',                       aliases: ['material'] },
    kho:  { label: 'Kho (Storage location)',          aliases: ['storage location'] },
    unit: { label: 'Đơn vị gốc',                      aliases: ['base unit of measure'] },
    mvt:  { label: 'Loại giao dịch',                  aliases: ['movement type text'] },
    po:   { label: 'Đơn mua hàng',                    aliases: ['purchase order'] },
    ncc:  { label: 'Nhà cung cấp',                    aliases: ['supplier'] },
    qtyEntry: { label: 'Số lượng theo ĐV nhập (cột L)', aliases: ['qty in unit of entry'] },
  };

  /** Nhóm: tên bắt đầu bằng "Dầu thô" → Dầu thô; còn lại → Nguyên liệu khác */
  const GROUPS = ['Dầu thô', 'Nguyên liệu khác'];
  const groupOf = (sp) => (/^dau tho\b/.test(L.fold(sp)) ? GROUPS[0] : GROUPS[1]);

  const norm = (v) => String(v == null ? '' : v).normalize('NFC').replace(/[\s ]+/g, ' ').trim();
  const normHeader = (v) => norm(v).toLowerCase();

  class StructureError extends Error {}

  function findHeaderRow(rows) {
    const max = Math.min(rows.length, 40);
    for (let i = 0; i < max; i++) {
      const cells = new Set((rows[i] || []).map(normHeader));
      const hits = Object.values(FIELDS).filter((f) => f.aliases.some((a) => cells.has(a))).length;
      if (hits === 3) return i;
    }
    return -1;
  }

  function mapColumns(header, startCol) {
    const map = {}, notes = [], missing = [];
    const find = (aliases) => {
      for (const a of aliases) {
        const idx = header.findIndex((h) => normHeader(h) === a);
        if (idx >= 0) return idx;
      }
      return -1;
    };
    for (const [k, f] of Object.entries(FIELDS)) {
      const fixed = XLSX.utils.decode_col(f.col) - startCol;
      if (fixed >= 0 && f.aliases.includes(normHeader(header[fixed]))) { map[k] = fixed; continue; }
      const idx = find(f.aliases);
      if (idx >= 0) {
        map[k] = idx;
        notes.push(`"${f.label}" không nằm ở cột ${f.col}, đã nhận theo tên tiêu đề tại cột ${XLSX.utils.encode_col(idx + startCol)}.`);
      } else missing.push(`${f.label} (cột ${f.col})`);
    }
    for (const [k, f] of Object.entries(OPTIONAL)) map[k] = find(f.aliases);
    return { map, notes, missing };
  }

  /** Đọc ArrayBuffer file Excel nhập → { records, meta, issues } */
  function parseWorkbook(buffer, fileName) {
    if (!XLSX) throw new Error('Thiếu thư viện đọc Excel (lib/xlsx.full.min.js).');
    let wb;
    try {
      wb = XLSX.read(buffer, { type: 'array', cellDates: false, dense: true });
    } catch (e) {
      throw new StructureError('Không đọc được file. Hãy chắc chắn đây là file Excel (.xlsx/.xls) hợp lệ. Chi tiết: ' + e.message);
    }
    let chosen = null;
    for (const name of wb.SheetNames) {
      const ws = wb.Sheets[name];
      if (!ws || !ws['!ref']) continue;
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true });
      const idx = findHeaderRow(rows);
      if (idx >= 0) { chosen = { name, rows, headerIdx: idx, range: XLSX.utils.decode_range(ws['!ref']) }; break; }
    }
    if (!chosen) {
      throw new StructureError('Không nhận dạng được dòng tiêu đề. File nhập cần có các cột "Posting Date" (Ngày nhập), "Material Description" (Dầu thô/Nguyên liệu) và "Quantity" (Số lượng). Sheet có trong file: ' + wb.SheetNames.join(', '));
    }
    const { rows, headerIdx, range } = chosen;
    const { map, notes, missing } = mapColumns(rows[headerIdx] || [], range.s.c);
    if (missing.length) throw new StructureError('File thiếu các cột bắt buộc: ' + missing.join(', ') + '.');

    // Gom biến thể tên (khác khoảng trắng/hoa thường) về 1 khóa, hiển thị biến thể phổ biến nhất
    const variants = new Map();
    const keyOf = (v) => {
      const c = norm(v);
      if (!c) return EMPTY;
      const k = c.toLocaleUpperCase('vi');
      let m = variants.get(k);
      if (!m) variants.set(k, (m = new Map()));
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
      // Dòng tổng cộng cuối bảng: không có ngày và nguyên liệu → bỏ qua
      if (isEmpty(get('date')) && isEmpty(get('sp'))) { summaryRows++; continue; }
      const raw = {};
      for (const k of [...Object.keys(FIELDS), ...Object.keys(OPTIONAL)]) raw[k] = get(k);
      const rec = {
        id: records.length,
        row: range.s.r + i + 1,
        raw,
        day: L.parseDate(raw.date),
        sp: keyOf(raw.sp),
        doc: norm(raw.doc), item: norm(raw.item), mat: norm(raw.mat), kho: norm(raw.kho) || EMPTY,
        srcUnit: norm(raw.unit).toUpperCase(),
        flags: {},
      };
      const vol = L.parseVolume(raw.kl);
      rec.vol = vol.value;
      rec.unit = UNIT;
      if (vol.error) rec.flags.volError = vol.error;
      // Số lượng chỉ cộng khi đơn vị gốc là MT (hoặc file không có cột đơn vị)
      if (rec.srcUnit && rec.srcUnit !== UNIT) rec.flags.volError = `Đơn vị gốc ${rec.srcUnit}, không phải ${UNIT}`;
      rec.volOk = rec.vol != null && !rec.flags.volError;
      if (!rec.volOk) rec.vol = null;
      if (rec.day == null) rec.flags.dateError = isEmpty(raw.date) ? 'Thiếu ngày' : 'Ngày sai định dạng';
      if (rec.volOk && rec.vol < 0) rec.flags.negative = true;
      records.push(rec);
    }

    const labels = new Map([[EMPTY, EMPTY]]);
    for (const [k, vs] of variants) {
      let best = null, bc = -1;
      for (const [v, c] of vs) if (c > bc) { best = v; bc = c; }
      labels.set(k, best);
    }
    for (const rec of records) {
      rec.sp = labels.get(rec.sp);
      rec.nhom = rec.sp === EMPTY ? GROUPS[1] : groupOf(rec.sp);
      rec.dateStr = L.dayToStr(rec.day);
      rec.search = L.fold([rec.dateStr, rec.sp, rec.nhom, rec.doc, rec.mat, rec.kho, rec.raw.kl].join(' | '));
    }

    const crudeDup = dropCrudeDuplicates(records, map.qtyEntry >= 0);
    if (map.qtyEntry < 0) notes.push(`Không thấy cột "${OPTIONAL.qtyEntry.aliases[0]}" (cột L) — chưa loại dòng dầu thô nhập trùng.`);

    const issues = checkQuality(records, variants);
    issues.crudeDupRemoved = crudeDup;
    issues.blankRows = blank;
    issues.summaryRows = summaryRows;

    const days = records.map((r) => r.day).filter((d) => d != null);
    const meta = {
      fileName: fileName || '',
      sheet: chosen.name,
      headerRow: range.s.r + headerIdx + 1,
      columnList: Object.entries(FIELDS).map(([k, f]) => ({ label: f.label, col: XLSX.utils.encode_col(map[k] + range.s.c) })),
      hasDoc: map.doc >= 0,
      notes,
      minDay: days.length ? Math.min(...days) : null,
      maxDay: days.length ? Math.max(...days) : null,
      units: [UNIT],
      unitRule: `Số lượng theo ${UNIT} (cột ${FIELDS.kl.col})`,
      loadedAt: new Date(),
    };
    return { records, meta, issues };
  }

  /**
   * Dầu thô bị nhập trùng trên SAP: cùng tên (cột D) và cùng số lượng theo ĐV nhập (cột L)
   * → chỉ giữ dòng đầu tiên trong file. Nguyên liệu khác giữ nguyên.
   * Sửa mảng records tại chỗ (đánh lại id) và trả về danh sách dòng đã loại.
   */
  function dropCrudeDuplicates(records, enabled) {
    const removed = [];
    if (!enabled) return removed;
    const kept = new Map();
    const out = records.filter((r) => {
      const qty = norm(r.raw.qtyEntry);
      if (r.nhom !== GROUPS[0] || !qty) return true;
      const k = r.sp + '|' + qty;
      const first = kept.get(k);
      if (!first) { kept.set(k, r); return true; }
      removed.push({ row: r.row, sp: r.sp, qty, dateStr: r.dateStr, doc: r.doc, keptRow: first.row });
      return false;
    });
    records.length = 0;
    out.forEach((r, i) => { r.id = i; records.push(r); });
    return removed;
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
    // 1 chứng từ SAP có thể nhiều dòng (item) → trùng khi cùng số chứng từ VÀ cùng dòng chứng từ
    const dupDoc = groupDup((r) => (r.doc ? r.doc + '|' + r.item : null));
    const dupRow = groupDup((r) => [r.day, r.sp, norm(r.raw.kl), r.doc, r.item].join('|'));
    for (const id of dupDoc.ids) records[id].flags.dupDoc = true;
    for (const id of dupRow.ids) records[id].flags.dupRow = true;
    return {
      dateMissing: by((r) => r.flags.dateError),
      volInvalid: by((r) => !r.volOk),
      negative: by((r) => r.flags.negative),
      missingSp: by((r) => r.sp === EMPTY),
      dupDoc: dupDoc.ids, dupDocGroups: dupDoc.groups,
      dupRow: dupRow.ids, dupRowGroups: dupRow.groups,
      merged: [...variants.values()].filter((vs) => vs.size > 1).map((vs) => [...vs.keys()]),
    };
  }

  PXK.importLoader = { parseWorkbook, FIELDS, OPTIONAL, GROUPS, StructureError, UNIT };
})(typeof window !== 'undefined' ? window : globalThis);
