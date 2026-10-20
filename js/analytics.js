/* =====================================================================
 * analytics.js — LỌC & TỔNG HỢP SỐ LIỆU
 * Mọi tổng khối lượng đều tách theo đơn vị (M3, MT...) — không cộng lẫn.
 * Bản ghi thiếu/sai khối lượng vẫn được đếm là giao dịch nhưng không cộng vào tổng.
 * ===================================================================== */
(function (root) {
  'use strict';
  const PXK = (root.PXK = root.PXK || {});

  const DIMS = ['kho', 'kh', 'ht', 'pt', 'sp', 'tt'];
  const DIM_LABELS = { kho: 'Kho', kh: 'Khách hàng', ht: 'Hình thức vận chuyển', pt: 'Phương tiện', sp: 'Sản phẩm', tt: 'Tình trạng' };

  /**
   * filters = { from: day|null, to: day|null, sets: { dim: Set } }  (Set rỗng = tất cả)
   * exceptDim: bỏ qua điều kiện của 1 chiều (dùng để đếm lựa chọn trong bộ lọc)
   */
  function applyFilters(records, filters, exceptDim) {
    const { from, to, sets, nonZero } = filters;
    const active = DIMS.filter((d) => d !== exceptDim && sets[d] && sets[d].size);
    const hasDate = from != null || to != null;
    const out = [];
    for (const r of records) {
      if (nonZero && r.volOk && r.vol === 0) continue;
      if (hasDate) {
        if (r.day == null) continue;
        if (from != null && r.day < from) continue;
        if (to != null && r.day > to) continue;
      }
      let ok = true;
      for (const d of active) if (!sets[d].has(r[d])) { ok = false; break; }
      if (ok) out.push(r);
    }
    return out;
  }

  /** { M3: {sum, count}, MT: {...} } + invalid */
  function totalsByUnit(recs) {
    const t = {};
    let invalid = 0;
    for (const r of recs) {
      if (!r.volOk) { invalid++; continue; }
      const u = t[r.unit] || (t[r.unit] = { sum: 0, count: 0 });
      u.sum += r.vol;
      u.count++;
    }
    for (const u of Object.values(t)) u.sum = round(u.sum);
    return { units: t, invalid };
  }

  function round(x) { return Math.round(x * 1e6) / 1e6; }

  function kpis(recs) {
    const distinct = (dim) => {
      const s = new Set();
      for (const r of recs) if (r[dim] !== PXK.loader.EMPTY) s.add(r[dim]);
      return s.size;
    };
    const tot = totalsByUnit(recs);
    return {
      count: recs.length,
      customers: distinct('kh'), products: distinct('sp'), vehicles: distinct('pt'), warehouses: distinct('kho'),
      totals: tot.units, invalid: tot.invalid,
      missingDate: recs.filter((r) => r.day == null).length,
    };
  }

  /** Tổng hợp theo 1 chiều cho 1 đơn vị: [{key, sum, count}] giảm dần */
  function groupBy(recs, dim, unit) {
    const m = new Map();
    for (const r of recs) {
      if (!r.volOk || r.unit !== unit) continue;
      const g = m.get(r[dim]) || m.set(r[dim], { key: r[dim], sum: 0, count: 0 }).get(r[dim]);
      g.sum += r.vol;
      g.count++;
    }
    return [...m.values()].map((g) => ((g.sum = round(g.sum)), g)).sort((a, b) => b.sum - a.sum);
  }

  /** Tổng hợp đa đơn vị theo 1 hoặc nhiều chiều: [{keys:[], key, count, invalid, byUnit:{u:{sum,count}}}] */
  function groupMulti(recs, dims) {
    dims = [].concat(dims);
    const m = new Map();
    for (const r of recs) {
      const keys = dims.map((d) => r[d]);
      const k = keys.join('\u0001');
      let g = m.get(k);
      if (!g) m.set(k, (g = { keys, key: keys.join(' / '), count: 0, invalid: 0, byUnit: {} }));
      g.count++;
      if (!r.volOk) { g.invalid++; continue; }
      const u = g.byUnit[r.unit] || (g.byUnit[r.unit] = { sum: 0, count: 0 });
      u.sum += r.vol;
      u.count++;
    }
    const arr = [...m.values()];
    for (const g of arr) for (const u of Object.values(g.byUnit)) u.sum = round(u.sum);
    return arr;
  }

  // ------------------------- Thời gian -------------------------
  const dayDate = (day) => new Date(day * 86400000);
  function periodOf(day, gran) {
    const d = dayDate(day);
    if (gran === 'day') return { key: day, start: day, end: day };
    if (gran === 'week') {
      const wd = (day + 3) % 7; // 0 = Thứ Hai (01/01/1970 là Thứ Năm)
      const start = day - wd;
      return { key: start, start, end: start + 6 };
    }
    const y = d.getUTCFullYear(), mo = d.getUTCMonth();
    const start = Math.round(Date.UTC(y, mo, 1) / 86400000);
    const end = Math.round(Date.UTC(y, mo + 1, 1) / 86400000) - 1;
    return { key: start, start, end };
  }

  function isoWeek(start) {
    const thu = dayDate(start + 3);
    const y = thu.getUTCFullYear();
    const jan1 = Math.round(Date.UTC(y, 0, 1) / 86400000);
    return { week: Math.floor((start + 3 - jan1) / 7) + 1, year: y };
  }

  function periodLabel(p, gran, short) {
    const s = PXK.loader.dayToStr;
    if (gran === 'day') return s(p.start);
    if (gran === 'week') {
      const w = isoWeek(p.start);
      return short ? `T${w.week}/${w.year}` : `Tuần ${w.week}/${w.year} (${s(p.start).slice(0, 5)}–${s(p.end).slice(0, 5)})`;
    }
    const d = dayDate(p.start);
    return `${String(d.getUTCMonth() + 1).padStart(2, '0')}/${d.getUTCFullYear()}`;
  }

  /** Chuỗi thời gian liên tục (điền 0 cho kỳ không có giao dịch) cho 1 đơn vị */
  function timeSeries(recs, gran, unit) {
    const m = new Map();
    let min = Infinity, max = -Infinity;
    for (const r of recs) {
      if (r.day == null) continue;
      if (r.day < min) min = r.day;
      if (r.day > max) max = r.day;
      if (!r.volOk || r.unit !== unit) continue;
      const p = periodOf(r.day, gran);
      const g = m.get(p.key) || m.set(p.key, { sum: 0, count: 0 }).get(p.key);
      g.sum += r.vol;
      g.count++;
    }
    if (min === Infinity) return [];
    const out = [];
    let p = periodOf(min, gran);
    const last = periodOf(max, gran).key;
    while (p.key <= last) {
      const g = m.get(p.key) || { sum: 0, count: 0 };
      out.push({ ...p, sum: round(g.sum), count: g.count, label: periodLabel(p, gran, true), longLabel: periodLabel(p, gran, false) });
      p = periodOf(p.end + 1, gran);
    }
    return out;
  }

  /** Top N + gộp phần còn lại thành "Khác" */
  function topN(groups, n, otherLabel = 'Khác') {
    if (groups.length <= n) return groups;
    const head = groups.slice(0, n);
    const rest = groups.slice(n);
    head.push({ key: otherLabel, sum: round(rest.reduce((a, g) => a + g.sum, 0)), count: rest.reduce((a, g) => a + g.count, 0), other: true, members: rest.map((g) => g.key) });
    return head;
  }

  PXK.analytics = { DIMS, DIM_LABELS, applyFilters, totalsByUnit, kpis, groupBy, groupMulti, timeSeries, periodOf, periodLabel, topN, round };
})(typeof window !== 'undefined' ? window : globalThis);
