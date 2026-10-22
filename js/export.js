/* =====================================================================
 * export.js — TRÍCH XUẤT BÁO CÁO EXCEL
 * Mỗi file đều có sheet "Điều kiện lọc" ghi rõ nguồn dữ liệu và bộ lọc áp dụng.
 * Ngày xuất ghi dạng ngày thật của Excel (DD/MM/YYYY); khối lượng tách theo đơn vị.
 * ===================================================================== */
(function (root) {
  'use strict';
  const PXK = (root.PXK = root.PXK || {});
  const XLSX = root.XLSX;
  const A = () => PXK.analytics;

  const EXCEL_EPOCH = 25569; // 01/01/1970 trong hệ ngày Excel
  const FMT = { date: 'dd/mm/yyyy', num: '#,##0.000', int: '#,##0', pct: '0.00%' };

  function stamp() {
    const d = new Date(), p = (n) => String(n).padStart(2, '0');
    return `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}_${p(d.getHours())}${p(d.getMinutes())}${p(d.getSeconds())}`;
  }

  /** cols: [{h, type, w}] ; rows: mảng giá trị (date = số ngày int) */
  function sheet(cols, rows, footer) {
    const aoa = [cols.map((c) => c.h)];
    for (const r of rows) aoa.push(r.map((v, i) => (cols[i].type === 'date' && v != null ? v + EXCEL_EPOCH : v)));
    if (footer) aoa.push(footer);
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const range = XLSX.utils.decode_range(ws['!ref']);
    for (let R = 1; R <= range.e.r; R++) {
      cols.forEach((c, C) => {
        const cell = ws[XLSX.utils.encode_cell({ r: R, c: C })];
        if (cell && cell.t === 'n' && FMT[c.type]) cell.z = FMT[c.type];
      });
    }
    ws['!cols'] = cols.map((c) => ({ wch: c.w || 14 }));
    ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: Math.max(rows.length, 1), c: cols.length - 1 } }) };
    return ws;
  }

  function conditionsSheet(ctx, title, rowCount) {
    const aoa = [
      ['BÁO CÁO', title],
      ['Thời điểm xuất file', new Date().toLocaleString('vi-VN')],
      ['File dữ liệu nguồn', ctx.meta.fileName],
      ['Sheet nguồn', ctx.meta.sheet],
      ['Số bản ghi trong báo cáo', rowCount],
      [],
      ['ĐIỀU KIỆN LỌC ĐANG ÁP DỤNG', ''],
      ...ctx.conditions.map((c) => [c.label, c.value]),
      [],
      ['GHI CHÚ', ''],
      ['Đơn vị', 'Khối lượng tổng hợp riêng theo từng đơn vị (' + ctx.meta.units.join(', ') + '). ' + ctx.meta.unitRule + '. Không cộng lẫn các đơn vị.'],
      ['Bản ghi thiếu/sai khối lượng', 'Được đếm là giao dịch nhưng không cộng vào tổng khối lượng.'],
    ];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = [{ wch: 32 }, { wch: 90 }];
    return ws;
  }

  function save(wb, name) {
    XLSX.writeFile(wb, `${name}_${stamp()}.xlsx`, { compression: true });
  }

  // ---------------------- 1. Chi tiết giao dịch ----------------------
  function detailSheet(recs, labels) {
    const cols = [
      { h: 'STT', type: 'int', w: 7 }, { h: 'Dòng Excel gốc', type: 'int', w: 9 },
      { h: 'Ngày xuất', type: 'date', w: 12 }, { h: 'Kho', w: 34 }, { h: 'Khách hàng', w: 26 },
      { h: 'Hình thức vận chuyển', w: 16 }, { h: 'Phương tiện', w: 22 }, { h: 'Sản phẩm', w: 18 },
      { h: 'Khối lượng thực tế (gốc)', w: 20 }, { h: 'Khối lượng (số)', type: 'num', w: 15 }, { h: 'Đơn vị', w: 8 },
      { h: labels.so, w: 13 }, { h: labels.tt, w: 18 }, { h: 'Cảnh báo', w: 30 },
    ];
    const rows = recs.map((r, i) => [
      i + 1, r.row, r.day, raw(r.raw.kho), raw(r.raw.kh), r.flags.htDefault ? r.ht : raw(r.raw.ht), raw(r.raw.pt), raw(r.raw.sp),
      raw(r.raw.kl), r.volOk ? r.vol : null, r.unit || '', r.so, raw(r.raw.tt), PXK.flagText(r),
    ]);
    return sheet(cols, rows);
  }
  const raw = (v) => (v == null ? '' : v);

  // ------------------------ Bảng tổng hợp chung ------------------------
  function summarySheet(recs, dims, units, opts = {}) {
    dims = [].concat(dims);
    const groups = A().groupMulti(recs, dims);
    const totals = A().totalsByUnit(recs).units;
    const sortUnit = units[0];
    groups.sort(opts.sort || ((a, b) => ((b.byUnit[sortUnit] || {}).sum || 0) - ((a.byUnit[sortUnit] || {}).sum || 0) || b.count - a.count));
    const cols = [{ h: 'STT', type: 'int', w: 6 }];
    dims.forEach((d) => cols.push({ h: opts.dimHeaders ? opts.dimHeaders[d] : A().DIM_LABELS[d], type: opts.dimTypes ? opts.dimTypes[d] : undefined, w: opts.widths ? opts.widths[d] : 28 }));
    cols.push({ h: 'Số giao dịch', type: 'int', w: 12 });
    for (const u of units) {
      cols.push({ h: `Khối lượng (${u})`, type: 'num', w: 16 });
      cols.push({ h: `Số GD (${u})`, type: 'int', w: 11 });
      cols.push({ h: `Tỷ trọng ${u}`, type: 'pct', w: 11 });
    }
    cols.push({ h: 'GD thiếu/sai KL', type: 'int', w: 12 });
    const rows = groups.map((g, i) => {
      const row = [i + 1, ...(opts.keyValues ? opts.keyValues(g) : g.keys), g.count];
      for (const u of units) {
        const v = g.byUnit[u];
        row.push(v ? v.sum : 0, v ? v.count : 0, v && totals[u] && totals[u].sum ? v.sum / totals[u].sum : 0);
      }
      row.push(g.invalid);
      return row;
    });
    const footer = ['', 'TỔNG CỘNG', ...dims.slice(1).map(() => ''), recs.length];
    for (const u of units) footer.push(totals[u] ? totals[u].sum : 0, totals[u] ? totals[u].count : 0, totals[u] ? 1 : 0);
    footer.push(recs.filter((r) => !r.volOk).length);
    return sheet(cols, rows, footer);
  }

  function book(ctx, title, sheets, rowCount) {
    const wb = XLSX.utils.book_new();
    for (const [name, ws] of sheets) XLSX.utils.book_append_sheet(wb, ws, name);
    XLSX.utils.book_append_sheet(wb, conditionsSheet(ctx, title, rowCount), 'Điều kiện lọc');
    return wb;
  }

  function timeSheet(recs, units, gran) {
    const tmp = recs.filter((r) => r.day != null).map((r) => {
      const p = A().periodOf(r.day, gran);
      return Object.assign(Object.create(r), { _p: p.start, _pl: A().periodLabel(p, gran, false) });
    });
    const header = { day: 'Ngày xuất', week: 'Tuần', month: 'Tháng' }[gran];
    return summarySheet(tmp, gran === 'day' ? ['_p'] : ['_p', '_pl'], units, {
      sort: (a, b) => a.keys[0] - b.keys[0],
      dimHeaders: { _p: gran === 'day' ? 'Ngày xuất' : 'Từ ngày', _pl: header },
      dimTypes: { _p: 'date' }, widths: { _p: 12, _pl: 26 },
    });
  }

  // ------------------------------ API ------------------------------
  const exporter = {
    detail(ctx) {
      const recs = ctx.tableRecords;
      save(book({ ...ctx, conditions: ctx.tableConditions },'Chi tiết giao dịch xuất hàng', [['Chi tiết giao dịch', detailSheet(recs, ctx.meta.labels)]], recs.length), 'BaoCao_ChiTietXuatHang');
    },
    customer(ctx) {
      const r = ctx.records;
      save(book(ctx, 'Tổng hợp theo khách hàng', [
        ['Theo khách hàng', summarySheet(r, 'kh', ctx.units)],
        ['Khách hàng - Sản phẩm', summarySheet(r, ['kh', 'sp'], ctx.units)],
      ], r.length), 'BaoCao_TongHop_KhachHang');
    },
    product(ctx) {
      const r = ctx.records;
      save(book(ctx, 'Tổng hợp theo sản phẩm', [
        ['Theo sản phẩm', summarySheet(r, 'sp', ctx.units)],
        ['Sản phẩm - Kho', summarySheet(r, ['sp', 'kho'], ctx.units)],
      ], r.length), 'BaoCao_TongHop_SanPham');
    },
    time(ctx) {
      const r = ctx.records;
      save(book(ctx, 'Tổng hợp theo ngày / tuần / tháng', [
        ['Theo ngày', timeSheet(r, ctx.units, 'day')],
        ['Theo tuần', timeSheet(r, ctx.units, 'week')],
        ['Theo tháng', timeSheet(r, ctx.units, 'month')],
      ], r.length), 'BaoCao_TongHop_ThoiGian');
    },
    transport(ctx) {
      const r = ctx.records;
      save(book(ctx, 'Tổng hợp theo hình thức vận chuyển và phương tiện', [
        ['Theo HTVC', summarySheet(r, 'ht', ctx.units)],
        ['Theo phương tiện', summarySheet(r, 'pt', ctx.units)],
        ['HTVC - Phương tiện', summarySheet(r, ['ht', 'pt'], ctx.units)],
      ], r.length), 'BaoCao_TongHop_VanChuyen');
    },
    all(ctx) {
      const r = ctx.records;
      save(book(ctx, 'Báo cáo tổng hợp xuất hàng', [
        ['Theo khách hàng', summarySheet(r, 'kh', ctx.units)],
        ['Theo sản phẩm', summarySheet(r, 'sp', ctx.units)],
        ['Theo kho', summarySheet(r, 'kho', ctx.units)],
        ['Theo tháng', timeSheet(r, ctx.units, 'month')],
        ['Theo ngày', timeSheet(r, ctx.units, 'day')],
        ['Theo HTVC', summarySheet(r, 'ht', ctx.units)],
        ['Theo phương tiện', summarySheet(r, 'pt', ctx.units)],
        ['Chi tiết giao dịch', detailSheet(r, ctx.meta.labels)],
      ], r.length), 'BaoCao_TongHop_XuatHang');
    },
    // dùng cho kiểm thử
    _summarySheet: summarySheet, _detailSheet: detailSheet,
    // dùng chung cho trang nhập dầu thô/nguyên liệu
    _sheet: sheet, _save: save,
  };

  PXK.exporter = exporter;
})(window);
