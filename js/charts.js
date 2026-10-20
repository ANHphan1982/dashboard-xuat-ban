/* =====================================================================
 * charts.js — HIỂN THỊ BIỂU ĐỒ (Chart.js)
 * Mỗi biểu đồ chỉ hiển thị 1 đơn vị đo (chọn bằng nút M3/MT) để không cộng lẫn đơn vị.
 * Màu sắc gắn cố định với thực thể (sản phẩm, HTVC...) — không đổi theo thứ hạng sau lọc.
 * ===================================================================== */
(function (root) {
  'use strict';
  const PXK = (root.PXK = root.PXK || {});
  const A = () => PXK.analytics;

  // Bảng màu phân loại (thứ tự cố định, đã kiểm tra an toàn mù màu) + màu "Khác"
  const SERIES = ['#2a78d6', '#eb6834', '#1baf7a', '#eda100', '#e87ba4', '#008300', '#4a3aa7'];
  const OTHER = '#a7afb9';
  const PRIMARY = '#1f5fae';
  const PRIMARY_SOFT = 'rgba(31,95,174,0.10)';
  const INK = '#1d2733', INK2 = '#5b6675', GRID = '#e8ecf1';

  const fmt = (v, max = 3) => (v == null ? '—' : v.toLocaleString('vi-VN', { maximumFractionDigits: max }));
  const fmtInt = (v) => v.toLocaleString('vi-VN');

  let colorRank = {}; // dim -> Map(key -> slot)
  /** Gán slot màu theo thứ hạng số giao dịch trên TOÀN BỘ dữ liệu (ổn định khi lọc) */
  function setColorRanks(records) {
    colorRank = {};
    for (const dim of ['sp', 'ht', 'kho', 'kh']) {
      const m = new Map();
      for (const r of records) m.set(r[dim], (m.get(r[dim]) || 0) + 1);
      const order = [...m.entries()].sort((a, b) => b[1] - a[1]).map((e) => e[0]);
      colorRank[dim] = new Map(order.map((k, i) => [k, i]));
    }
  }
  const EXTRA = '#e34948'; // slot 8
  function colorFor(dim, key) {
    if (key === 'Khác') return OTHER;
    const i = colorRank[dim] ? colorRank[dim].get(key) : undefined;
    return i != null && i < SERIES.length ? SERIES[i] : null;
  }
  /** Màu cho một nhóm khóa cùng hiển thị: giữ màu cố định; mục ngoài 7 màu chính lấy màu còn trống */
  function colorsFor(dim, keys) {
    const used = new Set(keys.map((k) => colorFor(dim, k)).filter(Boolean));
    const spare = [...SERIES, EXTRA].filter((c) => !used.has(c));
    return keys.map((k) => colorFor(dim, k) || spare.shift() || OTHER);
  }

  function setupDefaults() {
    const C = root.Chart;
    C.defaults.font.family = '"Segoe UI", system-ui, -apple-system, Roboto, Arial, sans-serif';
    C.defaults.font.size = 12;
    C.defaults.color = INK2;
    C.defaults.animation.duration = 250;
    C.defaults.maintainAspectRatio = false;
    C.defaults.plugins.tooltip.backgroundColor = '#0f2740';
    C.defaults.plugins.tooltip.padding = 10;
    C.defaults.plugins.tooltip.titleFont = { weight: '600' };
    C.defaults.plugins.legend.labels.boxWidth = 12;
    C.defaults.plugins.legend.labels.boxHeight = 12;
    C.defaults.plugins.legend.labels.color = INK;
  }

  const charts = {};
  function upsert(id, config) {
    const canvas = document.getElementById(id);
    if (!canvas) return;
    if (charts[id] && charts[id].config.type === config.type) {
      charts[id].data = config.data;
      charts[id].options = config.options;
      charts[id].update();
    } else {
      if (charts[id]) charts[id].destroy();
      charts[id] = new root.Chart(canvas, config);
    }
  }
  function setEmpty(id, empty, msg) {
    const box = document.getElementById(id).closest('.chart-box');
    let ov = box.querySelector('.chart-empty');
    if (!ov) { ov = document.createElement('div'); ov.className = 'chart-empty'; box.appendChild(ov); }
    ov.textContent = msg || 'Không có dữ liệu phù hợp với bộ lọc';
    ov.hidden = !empty;
  }

  const axisTitle = (text) => ({ display: true, text, color: INK2, font: { weight: '600' } });
  const valueAxis = (unit) => ({
    beginAtZero: true,
    grid: { color: GRID },
    border: { display: false },
    ticks: { callback: (v) => fmt(v, 0) },
    title: axisTitle(`Khối lượng (${unit})`),
  });
  const catAxis = (title, extra = {}) => ({ grid: { display: false }, title: axisTitle(title), ticks: { color: INK, ...extra } });
  const truncate = (s, n = 28) => (String(s).length > n ? String(s).slice(0, n - 1) + '…' : s);

  function onPick(handler) {
    return (evt, els, chart) => {
      if (!els.length) return;
      handler(els[0], chart);
    };
  }
  const pointer = (evt, els) => { evt.native && (evt.native.target.style.cursor = els.length ? 'pointer' : 'default'); };

  // -------------------------- 1. Xu hướng --------------------------
  function trend(recs, unit, gran, onClick) {
    const s = A().timeSeries(recs, gran, unit);
    setEmpty('chTrend', !s.some((p) => p.count));
    const granName = { day: 'Ngày', week: 'Tuần', month: 'Tháng' }[gran];
    upsert('chTrend', {
      type: 'line',
      data: {
        labels: s.map((p) => p.label),
        datasets: [{
          label: `Khối lượng xuất (${unit})`, data: s.map((p) => p.sum),
          borderColor: PRIMARY, backgroundColor: PRIMARY_SOFT, fill: true, borderWidth: 2, tension: 0.25,
          pointRadius: s.length > 60 ? 0 : 3, pointHoverRadius: 6, pointBackgroundColor: PRIMARY, pointHitRadius: 10,
        }],
      },
      options: {
        interaction: { mode: 'index', intersect: false },
        onHover: pointer,
        onClick: onPick((el) => onClick(s[el.index])),
        plugins: {
          legend: { position: 'top', align: 'end' },
          tooltip: {
            callbacks: {
              title: (it) => s[it[0].dataIndex].longLabel,
              label: (it) => ` Khối lượng: ${fmt(it.raw)} ${unit}`,
              afterLabel: (it) => ` Số giao dịch: ${fmtInt(s[it.dataIndex].count)}`,
            },
          },
        },
        scales: { x: catAxis(granName, { maxRotation: 0, autoSkip: true, autoSkipPadding: 12 }), y: valueAxis(unit) },
      },
    });
  }

  // ----------------------- Biểu đồ cột chung -----------------------
  function bars(id, groups, { unit, horizontal, catTitle, color, onClick, legend }) {
    setEmpty(id, !groups.length);
    const valAx = valueAxis(unit), cat = catAxis(catTitle, horizontal ? { autoSkip: false } : { autoSkip: false, maxRotation: 55, minRotation: 0 });
    cat.ticks.callback = function (v) { return truncate(this.getLabelForValue(v), horizontal ? 30 : 18); };
    upsert(id, {
      type: 'bar',
      data: {
        labels: groups.map((g) => g.key),
        datasets: [{
          label: legend || `Khối lượng (${unit})`, data: groups.map((g) => g.sum),
          backgroundColor: groups.map((g) => (color ? color(g) : PRIMARY)),
          borderRadius: 4, borderSkipped: 'start', maxBarThickness: horizontal ? 22 : 36,
        }],
      },
      options: {
        indexAxis: horizontal ? 'y' : 'x',
        onHover: pointer,
        onClick: onPick((el) => onClick(groups[el.index])),
        plugins: {
          legend: { position: 'top', align: 'end' },
          tooltip: {
            callbacks: {
              title: (it) => groups[it[0].dataIndex].key,
              label: (it) => ` Khối lượng: ${fmt(it.raw)} ${unit}`,
              afterLabel: (it) => {
                const g = groups[it.dataIndex];
                return ` Số giao dịch: ${fmtInt(g.count)}` + (g.other ? `\n (gộp ${g.members.length} mục)` : '');
              },
            },
          },
        },
        scales: horizontal ? { x: valAx, y: cat } : { x: cat, y: valAx },
      },
    });
  }

  // ------------------ 4. HTVC xếp chồng theo sản phẩm ------------------
  function transportStacked(recs, unit, onClick) {
    const ht = A().groupBy(recs, 'ht', unit);
    const prods = A().topN(A().groupBy(recs, 'sp', unit), 5);
    const prodKeys = prods.map((p) => p.key);
    const prodColors = colorsFor('sp', prodKeys);
    const top = new Set(prodKeys.filter((k) => k !== 'Khác'));
    const cell = new Map();
    for (const r of recs) {
      if (!r.volOk || r.unit !== unit) continue;
      const p = top.has(r.sp) ? r.sp : 'Khác';
      const k = r.ht + '\u0001' + p;
      cell.set(k, (cell.get(k) || 0) + r.vol);
    }
    setEmpty('chTransport', !ht.length);
    const datasets = prodKeys.map((p) => ({
      label: p,
      data: ht.map((h) => A().round(cell.get(h.key + '\u0001' + p) || 0)),
      backgroundColor: prodColors[prodKeys.indexOf(p)],
      borderColor: '#ffffff', borderWidth: { right: 2, left: 0, top: 0, bottom: 0 },
      borderSkipped: false, maxBarThickness: 30,
    }));
    upsert('chTransport', {
      type: 'bar',
      data: { labels: ht.map((h) => h.key), datasets },
      options: {
        indexAxis: 'y',
        onHover: pointer,
        onClick: onPick((el) => {
          const p = prodKeys[el.datasetIndex];
          onClick(ht[el.index].key, p === 'Khác' ? null : p, p === 'Khác' ? [...top] : null);
        }),
        plugins: {
          legend: { position: 'top', align: 'start' },
          tooltip: {
            mode: 'index', intersect: false, filter: (it) => it.raw > 0,
            callbacks: {
              label: (it) => ` ${it.dataset.label}: ${fmt(it.raw)} ${unit}`,
              footer: (its) => its.length ? `Tổng: ${fmt(ht[its[0].dataIndex].sum)} ${unit} · ${fmtInt(ht[its[0].dataIndex].count)} giao dịch` : '',
            },
          },
        },
        scales: {
          x: { ...valueAxis(unit), stacked: true },
          y: { ...catAxis('Hình thức vận chuyển', { autoSkip: false }), stacked: true },
        },
      },
    });
  }

  // ----------------------------- 6. Cơ cấu -----------------------------
  function donut(recs, unit, dim, onClick) {
    const groups = A().topN(A().groupBy(recs, dim, unit), 6);
    const total = groups.reduce((a, g) => a + g.sum, 0);
    const colors = colorsFor(dim, groups.map((g) => g.key));
    setEmpty('chShare', !groups.length);
    upsert('chShare', {
      type: 'doughnut',
      data: {
        labels: groups.map((g) => g.key),
        datasets: [{ data: groups.map((g) => g.sum), backgroundColor: colors, borderColor: '#ffffff', borderWidth: 2, hoverOffset: 6 }],
      },
      options: {
        cutout: '62%',
        onHover: pointer,
        onClick: onPick((el) => onClick(groups[el.index])),
        plugins: {
          legend: {
            position: 'right',
            labels: {
              generateLabels: (chart) => groups.map((g, i) => ({
                text: `${truncate(g.key, 24)}  ${total ? ((g.sum / total) * 100).toLocaleString('vi-VN', { maximumFractionDigits: 1 }) : 0}%`,
                fillStyle: colors[i], strokeStyle: colors[i], lineWidth: 0, index: i,
                hidden: !chart.getDataVisibility(i), fontColor: INK,
              })),
            },
          },
          tooltip: {
            callbacks: {
              label: (it) => ` ${fmt(it.raw)} ${unit} (${((it.raw / total) * 100).toLocaleString('vi-VN', { maximumFractionDigits: 2 })}%)`,
              afterLabel: (it) => ` Số giao dịch: ${fmtInt(groups[it.dataIndex].count)}`,
            },
          },
        },
      },
    });
  }

  function resizeAll() { for (const c of Object.values(charts)) c.resize(); }

  PXK.charts = { setupDefaults, setColorRanks, colorFor, colorsFor, trend, bars, transportStacked, donut, resizeAll, fmt, fmtInt };
})(window);
