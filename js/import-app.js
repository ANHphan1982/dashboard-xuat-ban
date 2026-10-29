/* =====================================================================
 * import-app.js — GIAO DIỆN TRANG NHẬP DẦU THÔ / NGUYÊN LIỆU
 * Luồng: nạp IMPORT.xlsx → bộ lọc → KPI / biểu đồ / bảng → xuất báo cáo
 * Dữ liệu chỉ có 1 đơn vị (MT) nên không có nút chọn đơn vị như trang xuất bán.
 * ===================================================================== */
(function () {
  'use strict';
  const { loader: L, importLoader: IL, analytics: A, charts: C, exporter: X, storage: ST } = window.PXK;
  const $ = (s, el = document) => el.querySelector(s);
  const $$ = (s, el = document) => [...el.querySelectorAll(s)];
  const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const fmt = C.fmt, fmtInt = C.fmtInt;
  const fmt2 = (v) => fmt(v, 2);
  const U = IL.UNIT;
  const DIMS = ['nhom', 'sp'];
  const DIM_LABELS = { nhom: 'Nhóm', sp: 'Dầu thô/Nguyên liệu' };

  // ------------------------------ Trạng thái ------------------------------
  const S = {
    data: null,
    source: null,                    // { name, handle, file, server, cachedAt }
    filters: { from: null, to: null, sets: { nhom: new Set(), sp: new Set() } },
    filtered: [],
    gran: 'month',
    drill: null,                     // { label, pred, scope: 'filtered' | 'all' }
    search: '', sort: { key: 'day', dir: -1 }, page: 1, pageSize: 50,
    tableRecords: [],
    options: {},
  };

  // ------------------------------ Tiện ích UI ------------------------------
  let toastTimer;
  function toast(msg, isErr) {
    const t = $('#toast');
    t.textContent = msg;
    t.className = 'toast' + (isErr ? ' err' : '');
    t.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => (t.hidden = true), isErr ? 7000 : 3500);
  }
  function dialog(html) { $('#dlgBody').innerHTML = html; $('#dlg').showModal(); }
  function loading(on, text) { $('#loading').hidden = !on; if (text) $('#loadingText').textContent = text; }
  const dayToInput = (day) => (day == null ? '' : new Date(day * 86400000).toISOString().slice(0, 10));
  const inputToDay = (v) => {
    if (!v) return null;
    const [y, m, d] = v.split('-').map(Number);
    return Math.round(Date.UTC(y, m - 1, d) / 86400000);
  };

  function flagText(r) {
    const f = r.flags, out = [];
    if (f.dupDoc) out.push('Trùng số chứng từ + dòng chứng từ');
    if (f.dupRow) out.push('Trùng nội dung');
    if (f.volError) out.push(f.volError);
    if (f.dateError) out.push(f.dateError);
    if (f.negative) out.push('Số lượng âm');
    return out.join('; ');
  }

  // ============================ LỌC ============================
  function applyFilters(records, filters, exceptDim) {
    const { from, to, sets } = filters;
    const active = DIMS.filter((d) => d !== exceptDim && sets[d].size);
    const hasDate = from != null || to != null;
    return records.filter((r) => {
      if (hasDate) {
        if (r.day == null) return false;
        if (from != null && r.day < from) return false;
        if (to != null && r.day > to) return false;
      }
      return active.every((d) => sets[d].has(r[d]));
    });
  }

  // ============================ MULTI-SELECT ============================
  class MultiSelect {
    constructor(host, dim) {
      this.dim = dim;
      this.counts = new Map();
      this.btn = document.createElement('button');
      this.btn.type = 'button';
      this.btn.className = 'ms-btn';
      this.btn.innerHTML = '<span>Tất cả</span>';
      this.btn.setAttribute('aria-haspopup', 'listbox');
      host.appendChild(this.btn);
      this.panel = document.createElement('div');
      this.panel.className = 'ms-panel';
      this.panel.hidden = true;
      this.panel.innerHTML = `<input type="search" placeholder="Tìm ${esc(DIM_LABELS[dim].toLowerCase())}…">
        <div class="ms-actions"><button type="button" class="btn-link" data-a="all">Chọn tất cả (đang hiện)</button><button type="button" class="btn-link" data-a="none">Bỏ chọn</button></div>
        <div class="ms-list" role="listbox" aria-multiselectable="true"></div>`;
      host.appendChild(this.panel);
      this.q = $('input', this.panel);
      this.list = $('.ms-list', this.panel);
      this.btn.addEventListener('click', (e) => { e.stopPropagation(); this.toggle(); });
      this.panel.addEventListener('click', (e) => {
        e.stopPropagation();
        const a = e.target.dataset.a;
        if (!a) return;
        const set = S.filters.sets[dim];
        if (a === 'none') set.clear();
        else this.visible().forEach((v) => set.add(v));
        this.changed();
      });
      this.q.addEventListener('input', () => this.renderList());
      this.list.addEventListener('change', (e) => {
        const set = S.filters.sets[dim];
        e.target.checked ? set.add(e.target.value) : set.delete(e.target.value);
        this.changed();
      });
    }
    toggle() {
      const open = this.panel.hidden;
      Object.values(MS).forEach((m) => (m.panel.hidden = true));
      this.panel.hidden = !open;
      if (open) { this.q.value = ''; this.renderList(); this.q.focus(); }
    }
    visible() {
      const q = L.fold(this.q.value.trim());
      return (S.options[this.dim] || []).filter((v) => !q || L.fold(v).includes(q));
    }
    renderList() {
      const set = S.filters.sets[this.dim];
      const vals = this.visible();
      if (!vals.length) { this.list.innerHTML = '<div class="ms-empty">Không có lựa chọn phù hợp</div>'; return; }
      this.list.innerHTML = vals.map((v) => {
        const c = this.counts.get(v) || 0;
        return `<label class="ms-opt${c ? '' : ' zero'}"><input type="checkbox" value="${esc(v)}"${set.has(v) ? ' checked' : ''}><span class="t">${esc(v)}</span><span class="c">${fmtInt(c)}</span></label>`;
      }).join('');
    }
    changed() {
      this.renderButton();
      if (!this.panel.hidden) this.renderList();
      scheduleUpdate();
    }
    renderButton() {
      const set = S.filters.sets[this.dim];
      const total = (S.options[this.dim] || []).length;
      $('span', this.btn).textContent = set.size === 1 ? [...set][0] : set.size > 1 ? `Đã chọn ${set.size}/${total}` : 'Tất cả';
      this.btn.title = set.size ? [...set].join('\n') : 'Tất cả';
      this.btn.classList.toggle('active', set.size > 0);
    }
    setCounts(map) { this.counts = map; if (!this.panel.hidden) this.renderList(); }
  }
  const MS = {};
  $$('[data-ms]').forEach((host) => (MS[host.dataset.ms] = new MultiSelect(host, host.dataset.ms)));
  document.addEventListener('click', () => Object.values(MS).forEach((m) => (m.panel.hidden = true)));
  document.addEventListener('keydown', (e) => { if (e.key === 'Escape') Object.values(MS).forEach((m) => (m.panel.hidden = true)); });

  // ============================ NẠP DỮ LIỆU ============================
  const DEFAULT_FILE = 'IMPORT.xlsx';
  const CACHE_KEY = 'import-last-v1', HANDLE_KEY = 'import-handle';
  // Máy chủ cục bộ (Mo_Dashboard.bat): tự đọc IMPORT.xlsx trong thư mục; bản online: người dùng tự chọn file
  const viaServer = /^(127\.0\.0\.1|localhost|\[::1\])$/.test(location.hostname);

  function openPicker() { const inp = $('#fileInput'); inp.value = ''; inp.click(); }

  async function loadFile(file, { handle = null, keepFilters = false } = {}) {
    if (!/\.(xlsx|xlsm|xls)$/i.test(file.name)) {
      dialog(`<h2>File không đúng định dạng</h2><p class="err">"${esc(file.name)}" không phải file Excel (.xlsx, .xls).</p>`);
      return false;
    }
    let buf;
    try { buf = await file.arrayBuffer(); } catch (e) {
      toast('Không đọc lại được file (có thể file đã bị thay đổi/di chuyển). Hãy nhấn "Chọn file Excel" để chọn lại.', true);
      return false;
    }
    const ok = await loadBuffer(buf, file.name, { keepFilters, fileTime: file.lastModified });
    if (!ok) return false;
    S.source = { name: file.name, handle, file };
    ST.set(CACHE_KEY, { name: file.name, fileTime: file.lastModified, savedAt: Date.now(), buf });
    if (handle) ST.set(HANDLE_KEY, handle); else ST.del(HANDLE_KEY);
    renderSource();
    toast(`Đã nạp ${fmtInt(S.data.records.length)} dòng từ "${file.name}".`);
    return true;
  }

  async function loadBuffer(buf, name, { keepFilters = false, fileTime = null } = {}) {
    loading(true, `Đang đọc "${name}"…`);
    await new Promise((r) => setTimeout(r, 40));
    try {
      const data = IL.parseWorkbook(new Uint8Array(buf), name);
      if (!data.records.length) throw new IL.StructureError('File không có dòng dữ liệu nào bên dưới dòng tiêu đề.');
      data.meta.fileTime = fileTime;
      const sameFile = S.data && S.data.meta.fileName === name;
      S.data = data;
      initAfterLoad(keepFilters || sameFile);
      return true;
    } catch (e) {
      console.error(e);
      const isStruct = e instanceof IL.StructureError;
      dialog(`<h2>${isStruct ? 'File không đúng cấu trúc' : 'Lỗi khi đọc file'}</h2>
        <p class="err">${esc(e.message)}</p>
        <h4>Cấu trúc file nhập dầu thô/nguyên liệu (báo cáo SAP)</h4>
        <ul>${Object.values(IL.FIELDS).map((f) => `<li>Cột <b>${f.col}</b>: ${esc(f.label)} <span class="muted">(tiêu đề "${esc(f.aliases[0])}")</span></li>`).join('')}</ul>
        ${S.data ? '<p>Dữ liệu cũ vẫn được giữ nguyên.</p>' : ''}`);
      return false;
    } finally { loading(false); }
  }

  async function loadFromServer(keepFilters) {
    let res;
    try { res = await fetch(encodeURIComponent(DEFAULT_FILE) + '?t=' + Date.now(), { cache: 'no-store' }); } catch (e) { res = null; }
    if (!res || !res.ok) {
      toast(`Không tìm thấy file "${DEFAULT_FILE}" trong thư mục dashboard. Hãy nhấn "Chọn file Excel".`, true);
      return false;
    }
    const lm = res.headers.get('Last-Modified');
    const ok = await loadBuffer(await res.arrayBuffer(), DEFAULT_FILE, { keepFilters, fileTime: lm ? Date.parse(lm) : null });
    if (ok) {
      S.source = { name: DEFAULT_FILE, server: true };
      renderSource();
      toast(`Đã nạp ${fmtInt(S.data.records.length)} dòng từ "${DEFAULT_FILE}" (bản mới nhất trong thư mục).`);
    }
    return ok;
  }

  async function refresh() {
    const src = S.source;
    if (viaServer && (!src || src.server)) { await loadFromServer(true); return; }
    if (src && src.handle) {
      try {
        let p = await src.handle.queryPermission({ mode: 'read' });
        if (p !== 'granted') p = await src.handle.requestPermission({ mode: 'read' });
        if (p === 'granted') { await loadFile(await src.handle.getFile(), { handle: src.handle, keepFilters: true }); return; }
      } catch (e) { console.warn(e); src.handle = null; ST.del(HANDLE_KEY); }
    }
    if (src && src.file) { if (await loadFile(src.file, { keepFilters: true })) return; }
    toast('Trình duyệt không cho đọc lại file tự động. Hãy nhấn "Chọn file Excel" và chọn file mới nhất.', true);
  }

  async function clearSaved() {
    if (!confirm('Xóa dữ liệu Excel nhập dầu thô/nguyên liệu đã lưu trong trình duyệt? Trang sẽ trở về màn hình chọn file.')) return;
    await ST.del(CACHE_KEY);
    await ST.del(HANDLE_KEY);
    location.reload();
  }

  function renderSource() {
    const m = S.data && S.data.meta;
    if (!m) { $('#sourceInfo').textContent = 'Chưa nạp dữ liệu'; return; }
    const ft = m.fileTime ? new Date(m.fileTime).toLocaleString('vi-VN') : '—';
    const cached = S.source && S.source.server ? ' · đọc trực tiếp từ thư mục dashboard'
      : S.source && S.source.cachedAt ? ' · bản lưu trong trình duyệt (chọn lại file Excel để cập nhật số liệu)' : '';
    $('#sourceInfo').textContent = `Nguồn: ${m.fileName} · sheet "${m.sheet}" · file sửa lúc ${ft} · nạp lúc ${m.loadedAt.toLocaleString('vi-VN')}${cached}`;
    $('#foot').textContent = `Dữ liệu đọc từ ${m.fileName}, sheet "${m.sheet}", tiêu đề ở dòng ${m.headerRow}. Cột: ${m.columnList.map((c) => `${c.label}=${c.col}`).join(', ')}. Đơn vị: ${m.unitRule}.`;
  }

  function initAfterLoad(keepFilters) {
    const recs = S.data.records;
    C.setColorRanks(recs);
    for (const d of DIMS) {
      const m = new Map();
      for (const r of recs) m.set(r[d], (m.get(r[d]) || 0) + 1);
      S.options[d] = [...m.entries()].sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0], 'vi')).map((e) => e[0]);
    }
    const f = S.filters;
    if (keepFilters) {
      for (const d of DIMS) for (const v of [...f.sets[d]]) if (!S.options[d].includes(v)) f.sets[d].delete(v);
    } else {
      f.from = f.to = null;
      for (const d of DIMS) f.sets[d].clear();
      S.drill = null; S.search = ''; $('#search').value = '';
      const span = (S.data.meta.maxDay || 0) - (S.data.meta.minDay || 0);
      S.gran = span > 120 ? 'month' : span > 31 ? 'week' : 'day';
    }
    $$('.dp').forEach((dp) => { dp.min = dayToInput(S.data.meta.minDay); dp.max = dayToInput(S.data.meta.maxDay); });
    $('#welcome').hidden = true;
    $('#dash').hidden = false;
    DIMS.forEach((d) => MS[d].renderButton());
    renderSource();
    renderQuality();
    update();
  }

  // ============================ CẬP NHẬT ============================
  let raf = 0;
  function scheduleUpdate() { cancelAnimationFrame(raf); raf = requestAnimationFrame(update); }

  function update() {
    if (!S.data) return;
    const recs = S.data.records;
    S.filtered = applyFilters(recs, S.filters);
    for (const d of DIMS) {
      const m = new Map();
      for (const r of applyFilters(recs, S.filters, d)) m.set(r[d], (m.get(r[d]) || 0) + 1);
      MS[d].setCounts(m);
    }
    for (const [id, v] of [['fFrom', S.filters.from], ['fTo', S.filters.to]]) {
      $('#' + id).value = L.dayToStr(v);
      $('#' + id).classList.remove('bad');
    }
    renderApplied();
    renderKpis();
    renderCharts();
    renderRank();
    S.page = 1;
    renderTable();
  }

  // ----------------------- Điều kiện lọc đang áp dụng -----------------------
  const periodText = () => {
    const f = S.filters, m = S.data.meta;
    return f.from == null && f.to == null ? `Toàn bộ (${L.dayToStr(m.minDay)} – ${L.dayToStr(m.maxDay)})` : `${f.from != null ? L.dayToStr(f.from) : 'đầu kỳ'} – ${f.to != null ? L.dayToStr(f.to) : 'cuối kỳ'}`;
  };
  function conditions() {
    const out = [{ label: 'Thời gian (ngày nhập)', value: periodText() }];
    for (const d of DIMS) {
      const set = S.filters.sets[d];
      out.push({ label: DIM_LABELS[d], value: set.size ? [...set].join(', ') : 'Tất cả' });
    }
    return out;
  }
  function tableConditions() {
    const c = conditions();
    if (S.drill) c.push({ label: 'Xem chi tiết', value: S.drill.label + (S.drill.scope === 'all' ? ' (trên toàn bộ dữ liệu, bỏ qua bộ lọc)' : '') });
    if (S.search) c.push({ label: 'Từ khóa tìm kiếm', value: S.search });
    const s = COLS.find((x) => x.k === S.sort.key);
    c.push({ label: 'Sắp xếp', value: `${s ? s.h : S.sort.key} ${S.sort.dir > 0 ? 'tăng dần' : 'giảm dần'}` });
    return c;
  }

  function renderApplied() {
    const f = S.filters;
    const chips = [];
    if (f.from == null && f.to == null) chips.push(`<span class="chip fixed"><span>Thời gian: ${esc(periodText().toLowerCase())}</span></span>`);
    else chips.push(`<span class="chip"><span>Thời gian: ${esc(periodText())}</span><button data-clear="date" title="Bỏ lọc thời gian">×</button></span>`);
    for (const d of DIMS) {
      const vals = [...f.sets[d]];
      if (!vals.length) continue;
      const txt = vals.slice(0, 3).join(', ') + (vals.length > 3 ? ` (+${vals.length - 3})` : '');
      chips.push(`<span class="chip" title="${esc(vals.join('\n'))}"><span>${esc(DIM_LABELS[d])}: ${esc(txt)}</span><button data-clear="${d}" title="Bỏ lọc ${esc(DIM_LABELS[d])}">×</button></span>`);
    }
    $('#applied').innerHTML = `<span class="lbl">Đang áp dụng:</span>${chips.join('')}<span class="lbl" style="margin-left:auto">${fmtInt(S.filtered.length)} / ${fmtInt(S.data.records.length)} lần nhập</span>`;
  }

  // ----------------------------- Chất lượng -----------------------------
  const ISSUE_LABEL = {
    dateMissing: 'Thiếu/sai ngày nhập', volInvalid: 'Thiếu/sai số lượng', negative: 'Số lượng âm',
    missingSp: 'Thiếu tên dầu thô/nguyên liệu', dupDoc: 'Trùng số chứng từ + dòng chứng từ', dupRow: 'Dòng trùng toàn bộ nội dung',
  };
  function renderQuality() {
    const { issues: q, meta: m, records } = S.data;
    const items = [];
    let warnCount = -1;
    const link = (key) => ` <button class="btn-link" data-issue="${key}">Xem</button>`;
    const li = (cls, html) => { if (!cls) warnCount++; items.push(`<li${cls ? ` class="${cls}"` : ''}>${html}</li>`); };
    li('', `Đã đọc <b>${fmtInt(records.length)}</b> dòng từ sheet "${esc(m.sheet)}" (tiêu đề dòng ${m.headerRow}).`);
    if (q.summaryRows) li('', `${fmtInt(q.summaryRows)} dòng tổng cộng/ghi chú cuối bảng (không có ngày và nguyên liệu) — đã bỏ qua.`);
    if (q.crudeDupRemoved.length) {
      const rows = q.crudeDupRemoved.map((d) => `<li>Dòng ${d.row}: ${esc(d.sp)} — ${fmt2(+d.qty)} ${U}, ngày ${esc(d.dateStr)}${d.doc ? `, CT ${esc(d.doc)}` : ''} (trùng dòng ${d.keptRow})</li>`).join('');
      li('', `Đã loại <b>${fmtInt(q.crudeDupRemoved.length)}</b> dòng dầu thô nhập trùng (cùng tên cột D và số lượng cột L, giữ dòng đầu tiên) — không tính vào tổng. <details><summary>Danh sách</summary><ul>${rows}</ul></details>`);
    }
    if (q.dupDoc.length) li('', `<b>${fmtInt(q.dupDocGroups)}</b> cặp số chứng từ + dòng chứng từ bị trùng (${fmtInt(q.dupDoc.length)} dòng) — chưa loại bỏ, cần kiểm tra.${link('dupDoc')}`);
    if (q.dupRow.length) li('', `<b>${fmtInt(q.dupRowGroups)}</b> nhóm dòng trùng toàn bộ ngày/nguyên liệu/số lượng/chứng từ (${fmtInt(q.dupRow.length)} dòng) — vẫn tính đủ, cần kiểm tra có nhập trùng không.${link('dupRow')}`);
    if (q.missingSp.length) li('', `<b>${fmtInt(q.missingSp.length)}</b> dòng thiếu tên dầu thô/nguyên liệu.${link('missingSp')}`);
    li(q.dateMissing.length ? '' : 'ok', q.dateMissing.length ? `<b>${fmtInt(q.dateMissing.length)}</b> dòng thiếu/sai ngày nhập — bị loại khi lọc theo thời gian.${link('dateMissing')}` : '✓ Ngày nhập: tất cả hợp lệ.');
    li(q.volInvalid.length ? '' : 'ok', q.volInvalid.length ? `<b>${fmtInt(q.volInvalid.length)}</b> dòng thiếu/sai số lượng hoặc đơn vị gốc không phải ${U} — vẫn đếm là lần nhập nhưng không cộng vào tổng.${link('volInvalid')}` : `✓ Số lượng: tất cả đọc được, đơn vị ${U}.`);
    if (q.negative.length) li('', `<b>${fmtInt(q.negative.length)}</b> dòng số lượng âm (có thể là bút toán hủy/điều chỉnh) — giữ nguyên, được trừ vào tổng.${link('negative')}`);
    if (q.merged.length) li('', 'Gộp biến thể tên khác hoa/thường: ' + q.merged.map((g) => g.map(esc).join(' = ')).join('; '));
    const notes = m.notes.length ? `<p class="q-notes">${m.notes.map(esc).join('<br>')}</p>` : '';
    $('#quality').innerHTML = `<details class="q-box"><summary>Kiểm tra dữ liệu: ${warnCount > 0 ? `${warnCount} điểm cần lưu ý` : 'không phát hiện vấn đề'} — nhấn để xem</summary><ul>${items.join('')}</ul>${notes}</details>`;
  }

  // -------------------------------- KPI --------------------------------
  function renderKpis() {
    const recs = S.filtered;
    const tot = A.totalsByUnit(recs);
    const t = tot.units[U] || { sum: 0, count: 0 };
    const byGroup = Object.fromEntries(A.groupBy(recs, 'nhom', U).map((g) => [g.key, g]));
    const grp = (k) => byGroup[k] || { sum: 0, count: 0 };
    const pct = (v) => (t.sum ? ((v / t.sum) * 100).toLocaleString('vi-VN', { maximumFractionDigits: 1 }) + '%' : '—');
    const kinds = new Set(recs.filter((r) => r.sp !== L.EMPTY).map((r) => r.sp)).size;
    const days = new Set(recs.filter((r) => r.day != null).map((r) => r.day)).size;
    const [g1, g2] = IL.GROUPS;
    $('#kpis').innerHTML = [
      `<div class="kpi main"><div class="k-t">Tổng khối lượng nhập</div><div class="k-v num">${fmt2(t.sum)}<small>${U}</small></div>${tot.invalid ? `<div class="k-n warn">${fmtInt(tot.invalid)} dòng thiếu/sai số lượng không cộng</div>` : '<div class="k-n">Dầu thô + nguyên liệu khác</div>'}</div>`,
      `<div class="kpi"><div class="k-t">${g1}</div><div class="k-v num">${fmt2(grp(g1).sum)}<small>${U}</small></div><div class="k-n">${pct(grp(g1).sum)} tổng · ${fmtInt(grp(g1).count)} lần nhập</div></div>`,
      `<div class="kpi"><div class="k-t">${g2}</div><div class="k-v num">${fmt2(grp(g2).sum)}<small>${U}</small></div><div class="k-n">${pct(grp(g2).sum)} tổng · ${fmtInt(grp(g2).count)} lần nhập</div></div>`,
      `<div class="kpi"><div class="k-t">Số lần nhập</div><div class="k-v num">${fmtInt(recs.length)}</div><div class="k-n">Trong ${fmtInt(days)} ngày có nhập</div></div>`,
      `<div class="kpi"><div class="k-t">Số loại dầu thô/NL</div><div class="k-v num">${fmtInt(kinds)}</div><div class="k-n">Loại duy nhất</div></div>`,
      `<div class="kpi"><div class="k-t">Bình quân / lần nhập</div><div class="k-v num">${t.count ? fmt2(t.sum / t.count) : '—'}<small>${U}</small></div><div class="k-n">Tổng khối lượng / số lần</div></div>`,
    ].join('');
  }

  // ------------------------------ Biểu đồ ------------------------------
  const volPred = (r) => r.volOk;
  function drillSp(g) {
    if (g.other) {
      const set = new Set(g.members);
      setDrill(`Dầu thô/Nguyên liệu: nhóm "Khác" (${g.members.length} loại)`, (r) => set.has(r.sp) && volPred(r));
    } else setDrill(`Dầu thô/Nguyên liệu = ${g.key}`, (r) => r.sp === g.key && volPred(r));
  }

  function renderCharts() {
    const recs = S.filtered;
    const n = A.totalsByUnit(recs).units[U];
    const sub = `Đơn vị: ${U} · ${fmtInt(n ? n.count : 0)} lần nhập`;
    $$('.unit-sub').forEach((el) => (el.textContent = sub));
    $('#trendSub').textContent = `${sub} · theo ${{ day: 'ngày', week: 'tuần', month: 'tháng' }[S.gran]}`;
    $$('#segGran button').forEach((b) => b.classList.toggle('on', b.dataset.v === S.gran));

    C.trend(recs, U, S.gran, (p) => {
      if (!p) return;
      const lbl = S.gran === 'day' ? `Ngày ${L.dayToStr(p.start)}` : S.gran === 'week' ? p.longLabel : `Tháng ${p.label}`;
      setDrill(lbl, (r) => r.day != null && r.day >= p.start && r.day <= p.end && volPred(r));
    }, `Khối lượng nhập (${U})`);
    const groups = A.groupBy(recs, 'sp', U);
    const palette = C.colorsFor('sp', groups.map((g) => g.key));
    const colors = new Map(groups.map((g, i) => [g.key, palette[i]]));
    C.bars('chProduct', groups, { unit: U, horizontal: true, catTitle: 'Dầu thô/Nguyên liệu', color: (g) => colors.get(g.key), legend: `Khối lượng nhập (${U})`, onClick: drillSp });
    C.donut(recs, U, 'sp', drillSp);
    monthlyStacked(recs);
  }

  /** Khối lượng theo tháng, xếp chồng theo 5 loại lớn nhất + "Khác" */
  function monthlyStacked(recs) {
    const months = A.timeSeries(recs, 'month', U);
    const prods = A.topN(A.groupBy(recs, 'sp', U), 5);
    const keys = prods.map((p) => p.key);
    const colors = C.colorsFor('sp', keys);
    const top = new Set(keys.filter((k) => k !== 'Khác'));
    const cell = new Map();
    for (const r of recs) {
      if (!r.volOk || r.day == null) continue;
      const k = A.periodOf(r.day, 'month').key + '\u0001' + (top.has(r.sp) ? r.sp : 'Khác');
      cell.set(k, (cell.get(k) || 0) + r.vol);
    }
    C.setEmpty('chMonthly', !months.some((p) => p.count));
    C.upsert('chMonthly', {
      type: 'bar',
      data: {
        labels: months.map((p) => p.label),
        datasets: keys.map((k, i) => ({
          label: k, data: months.map((p) => A.round(cell.get(p.key + '\u0001' + k) || 0)),
          backgroundColor: colors[i], borderColor: '#ffffff', borderWidth: { top: 1, bottom: 0, left: 0, right: 0 }, borderSkipped: false, maxBarThickness: 44,
        })),
      },
      options: {
        onHover: C.pointer,
        onClick: C.onPick((el) => {
          const p = months[el.index], k = keys[el.datasetIndex];
          const inMonth = (r) => r.day != null && r.day >= p.start && r.day <= p.end && volPred(r);
          if (k === 'Khác') setDrill(`Tháng ${p.label} · Dầu thô/Nguyên liệu khác`, (r) => inMonth(r) && !top.has(r.sp));
          else setDrill(`Tháng ${p.label} · ${k}`, (r) => inMonth(r) && r.sp === k);
        }),
        plugins: {
          legend: { position: 'top', align: 'start' },
          tooltip: {
            mode: 'index', intersect: false, filter: (it) => it.raw > 0,
            callbacks: {
              // filter bỏ hết mục = 0 → its rỗng ở tháng không có nhập; lỗi ở đây làm hỏng mọi lần cập nhật biểu đồ sau đó
              title: (its) => (its.length ? `Tháng ${months[its[0].dataIndex].label}` : ''),
              label: (it) => ` ${it.dataset.label}: ${fmt(it.raw)} ${U}`,
              footer: (its) => its.length ? `Tổng: ${fmt(months[its[0].dataIndex].sum)} ${U} · ${fmtInt(months[its[0].dataIndex].count)} lần nhập` : '',
            },
          },
        },
        scales: { x: { ...C.catAxis('Tháng'), stacked: true }, y: { ...C.valueAxis(U), stacked: true } },
      },
    });
  }

  // --------------------------- Bảng xếp hạng ---------------------------
  function renderRank() {
    const groups = A.groupBy(S.filtered, 'sp', U);
    const total = groups.reduce((a, g) => a + g.sum, 0);
    const max = groups.length ? groups[0].sum : 0;
    const nhomOf = new Map(S.data.records.map((r) => [r.sp, r.nhom]));
    const head = `<thead><tr><th>#</th><th>Dầu thô/Nguyên liệu</th><th>Nhóm</th><th class="r">Số lần</th><th class="r">KL (${U})</th><th class="r">Tỷ trọng</th><th class="r">BQ/lần (${U})</th></tr></thead>`;
    if (!groups.length) { $('#rankTable').innerHTML = head + '<tbody><tr class="empty-row"><td colspan="7">Không có dữ liệu phù hợp với bộ lọc</td></tr></tbody>'; return; }
    $('#rankTable').innerHTML = head + '<tbody>' + groups.map((g, i) => `<tr data-key="${esc(g.key)}"><td class="muted">${i + 1}</td>
      <td class="name">${esc(g.key)}${max > 0 && g.sum > 0 ? `<div class="bar" style="width:${Math.max(1, (g.sum / max) * 100).toFixed(1)}%"></div>` : ''}</td>
      <td class="muted">${esc(nhomOf.get(g.key))}</td><td class="r">${fmtInt(g.count)}</td><td class="r">${fmt(g.sum)}</td>
      <td class="r">${total ? ((g.sum / total) * 100).toLocaleString('vi-VN', { maximumFractionDigits: 2 }) + '%' : '—'}</td>
      <td class="r">${g.count ? fmt2(g.sum / g.count) : '—'}</td></tr>`).join('') + '</tbody>';
  }

  // ---------------------------- Bảng chi tiết ----------------------------
  function setDrill(label, pred, scope = 'filtered') {
    S.drill = { label, pred, scope };
    S.page = 1;
    renderTable();
    $('#detailCard').scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  const COLS = [
    { k: 'day', h: 'Ngày nhập', sort: true },
    { k: 'sp', h: 'Dầu thô/Nguyên liệu', sort: true },
    { k: 'nhom', h: 'Nhóm', sort: true },
    { k: 'vol', h: `Số lượng (${U})`, sort: true, r: true },
    { k: 'doc', h: 'Số chứng từ', sort: true },
    { k: 'item', h: 'Dòng CT', sort: true },
    { k: 'mat', h: 'Mã vật tư', sort: true },
    { k: 'kho', h: 'Kho', sort: true },
    { k: 'flags', h: 'Ghi chú' },
  ];

  function tableData() {
    let base = S.drill && S.drill.scope === 'all' ? S.data.records : S.filtered;
    if (S.drill) base = base.filter(S.drill.pred);
    const terms = L.fold(S.search.trim()).split(/\s+/).filter(Boolean);
    if (terms.length) base = base.filter((r) => terms.every((t) => r.search.includes(t)));
    const { key, dir } = S.sort;
    const coll = new Intl.Collator('vi', { numeric: true });
    return base.slice().sort((a, b) => {
      const va = a[key], vb = b[key];
      const na = va == null || va === '', nb = vb == null || vb === '';
      if (na !== nb) return na ? 1 : -1; // giá trị trống luôn xếp cuối
      let c = 0;
      if (!na) c = typeof va === 'number' ? va - vb : coll.compare(va, vb);
      if (c === 0) c = (a.day || 0) - (b.day || 0) || a.row - b.row;
      return c * dir;
    });
  }

  function renderTable() {
    if (!S.data) return;
    const rows = (S.tableRecords = tableData());
    const pages = Math.max(1, Math.ceil(rows.length / S.pageSize));
    S.page = Math.min(Math.max(1, S.page), pages);
    const start = (S.page - 1) * S.pageSize;
    const page = rows.slice(start, start + S.pageSize);

    const d = S.drill;
    $('#drillBar').hidden = !d;
    if (d) $('#drillBar').innerHTML = `<span>Đang xem chi tiết: <b>${esc(d.label)}</b>${d.scope === 'all' ? ' <i>(trên toàn bộ dữ liệu, bỏ qua bộ lọc)</i>' : ''}</span><button class="btn btn-outline" id="btnDrillOff">× Bỏ chọn, quay lại toàn bộ</button>`;

    const head = '<thead><tr>' + COLS.map((c) => {
      const on = S.sort.key === c.k;
      return `<th class="${c.sort ? 'sortable' : ''}${on ? ' sorted' : ''}${c.r ? ' r' : ''}" ${c.sort ? `data-sort="${c.k}" title="Nhấn để sắp xếp"` : ''}>${c.h}${c.sort ? `<span class="arr">${on ? (S.sort.dir > 0 ? '▲' : '▼') : '↕'}</span>` : ''}</th>`;
    }).join('') + '</tr></thead>';
    let body;
    if (!page.length) {
      body = `<tr class="empty-row"><td colspan="${COLS.length}">Không có lần nhập phù hợp với điều kiện lọc${S.search ? ` và từ khóa "${esc(S.search)}"` : ''}.</td></tr>`;
    } else {
      const cell = (v) => (v === L.EMPTY || v === '' ? `<span class="muted">${v || '—'}</span>` : esc(v));
      body = page.map((r) => {
        const f = r.flags, badges = [];
        if (f.dupDoc) badges.push('<span class="badge">Trùng chứng từ</span>');
        if (f.dupRow) badges.push('<span class="badge">Trùng nội dung</span>');
        if (f.volError) badges.push(`<span class="badge del">${esc(f.volError)}</span>`);
        if (f.dateError) badges.push(`<span class="badge del">${esc(f.dateError)}</span>`);
        if (f.negative) badges.push('<span class="badge">SL âm</span>');
        return `<tr><td class="num">${r.dateStr || `<span class="muted">${esc(r.raw.date)}</span>`}</td><td>${cell(r.sp)}</td><td>${esc(r.nhom)}</td>
          <td class="r" title="Giá trị gốc: ${esc(r.raw.kl)}">${r.volOk ? fmt(r.vol) : `<span class="muted">${esc(r.raw.kl) || '—'}</span>`}</td>
          <td class="num">${cell(r.doc)}</td><td class="num">${cell(r.item)}</td><td class="num">${cell(r.mat)}</td><td>${cell(r.kho)}</td><td>${badges.join('')}</td></tr>`;
      }).join('');
    }
    $('#detailTable').innerHTML = head + `<tbody>${body}</tbody>`;

    const tot = A.totalsByUnit(rows);
    $('#tableTotals').innerHTML = `Tổng số lần nhập: <b>${fmtInt(rows.length)}</b> · Tổng khối lượng: <b>${fmt(tot.units[U] ? tot.units[U].sum : 0)}</b> ${U}${tot.invalid ? ` <span class="muted">(${fmtInt(tot.invalid)} dòng thiếu/sai số lượng không cộng)</span>` : ''}`;
    $('#detailSub').textContent = rows.length ? `Hiển thị ${fmtInt(start + 1)}–${fmtInt(start + page.length)} / ${fmtInt(rows.length)} bản ghi` : 'Không có bản ghi';

    const btn = (p, label, dis, on) => `<button data-page="${p}" ${dis ? 'disabled' : ''} class="${on ? 'on' : ''}">${label}</button>`;
    const nums = [];
    for (let p = Math.max(1, S.page - 2); p <= Math.min(pages, S.page + 2); p++) nums.push(btn(p, p, false, p === S.page));
    $('#pager').innerHTML = btn(1, '«', S.page === 1) + btn(S.page - 1, '‹', S.page === 1) + nums.join('') + btn(S.page + 1, '›', S.page === pages) + btn(pages, '»', S.page === pages) + `<span class="pinfo">Trang ${fmtInt(S.page)}/${fmtInt(pages)}</span>`;
  }

  // ============================ XUẤT BÁO CÁO ============================
  const raw = (v) => (v == null ? '' : v);
  function detailSheet(recs) {
    const cols = [
      { h: 'STT', type: 'int', w: 7 }, { h: 'Dòng Excel gốc', type: 'int', w: 9 }, { h: 'Ngày nhập', type: 'date', w: 12 },
      { h: 'Dầu thô/Nguyên liệu', w: 36 }, { h: 'Nhóm', w: 16 }, { h: 'Số lượng (gốc)', w: 14 }, { h: `Số lượng (${U})`, type: 'num', w: 15 },
      { h: 'Số chứng từ', w: 13 }, { h: 'Dòng CT', w: 8 }, { h: 'Mã vật tư', w: 10 }, { h: 'Kho', w: 8 }, { h: 'Loại giao dịch', w: 22 }, { h: 'Cảnh báo', w: 30 },
    ];
    return X._sheet(cols, recs.map((r, i) => [i + 1, r.row, r.day, raw(r.raw.sp), r.nhom, raw(r.raw.kl), r.volOk ? r.vol : null, r.doc, r.item, r.mat, raw(r.raw.kho), raw(r.raw.mvt), flagText(r)]));
  }
  /** Tổng hợp theo 1 chiều: số lần, khối lượng, tỷ trọng, bình quân, ngày nhập đầu/cuối */
  function groupSheet(recs, dim, header) {
    const groups = A.groupMulti(recs, dim);
    const first = new Map(), last = new Map();
    for (const r of recs) {
      if (r.day == null) continue;
      if (!first.has(r[dim]) || r.day < first.get(r[dim])) first.set(r[dim], r.day);
      if (!last.has(r[dim]) || r.day > last.get(r[dim])) last.set(r[dim], r.day);
    }
    const sum = (g) => (g.byUnit[U] ? g.byUnit[U].sum : 0);
    const total = A.totalsByUnit(recs).units[U] || { sum: 0, count: 0 };
    groups.sort((a, b) => sum(b) - sum(a) || b.count - a.count);
    const cols = [
      { h: 'STT', type: 'int', w: 6 }, { h: header, w: 36 }, { h: 'Số lần nhập', type: 'int', w: 12 }, { h: `Khối lượng (${U})`, type: 'num', w: 16 },
      { h: 'Tỷ trọng', type: 'pct', w: 10 }, { h: `Bình quân/lần (${U})`, type: 'num', w: 16 }, { h: 'Nhập lần đầu', type: 'date', w: 12 }, { h: 'Nhập lần cuối', type: 'date', w: 12 },
    ];
    const rows = groups.map((g, i) => [i + 1, g.key, g.count, sum(g), total.sum ? sum(g) / total.sum : 0, g.byUnit[U] ? sum(g) / g.byUnit[U].count : 0, first.get(g.key), last.get(g.key)]);
    return X._sheet(cols, rows, ['', 'TỔNG CỘNG', recs.length, total.sum, total.sum ? 1 : 0, total.count ? total.sum / total.count : 0]);
  }
  function timeSheet(recs, gran) {
    const s = A.timeSeries(recs, gran, U).filter((p) => p.count);
    const cols = gran === 'day'
      ? [{ h: 'Ngày nhập', type: 'date', w: 12 }, { h: 'Số lần nhập', type: 'int', w: 12 }, { h: `Khối lượng (${U})`, type: 'num', w: 16 }]
      : [{ h: 'Tháng', w: 10 }, { h: 'Số lần nhập', type: 'int', w: 12 }, { h: `Khối lượng (${U})`, type: 'num', w: 16 }];
    const total = A.totalsByUnit(recs).units[U] || { sum: 0, count: 0 };
    return X._sheet(cols, s.map((p) => [gran === 'day' ? p.start : p.label, p.count, p.sum]), ['TỔNG CỘNG', total.count, total.sum]);
  }
  /** Bảng chéo tháng × dầu thô/nguyên liệu */
  function pivotSheet(recs) {
    const months = A.timeSeries(recs, 'month', U).filter((p) => p.count);
    const prods = A.groupBy(recs, 'sp', U).map((g) => g.key);
    const cell = new Map();
    for (const r of recs) {
      if (!r.volOk || r.day == null) continue;
      const k = A.periodOf(r.day, 'month').key + '\u0001' + r.sp;
      cell.set(k, (cell.get(k) || 0) + r.vol);
    }
    const cols = [{ h: 'Tháng', w: 10 }, ...prods.map((p) => ({ h: p, type: 'num', w: Math.min(30, Math.max(14, p.length)) })), { h: `Tổng (${U})`, type: 'num', w: 16 }];
    const rows = months.map((p) => [p.label, ...prods.map((sp) => A.round(cell.get(p.key + '\u0001' + sp) || 0)), p.sum]);
    const footer = ['TỔNG CỘNG', ...prods.map((_, i) => A.round(rows.reduce((a, r) => a + r[i + 1], 0))), A.round(months.reduce((a, p) => a + p.sum, 0))];
    return X._sheet(cols, rows, footer);
  }
  function conditionsSheet(title, rowCount, conds) {
    const m = S.data.meta;
    const ws = XLSX.utils.aoa_to_sheet([
      ['BÁO CÁO', title],
      ['Thời điểm xuất file', new Date().toLocaleString('vi-VN')],
      ['File dữ liệu nguồn', m.fileName],
      ['Sheet nguồn', m.sheet],
      ['Số bản ghi trong báo cáo', rowCount],
      [],
      ['ĐIỀU KIỆN LỌC ĐANG ÁP DỤNG', ''],
      ...conds.map((c) => [c.label, c.value]),
      [],
      ['GHI CHÚ', ''],
      ['Đơn vị', m.unitRule + '.'],
      ['Bản ghi thiếu/sai số lượng', `Được đếm là lần nhập nhưng không cộng vào tổng khối lượng (gồm dòng có đơn vị gốc khác ${U}).`],
    ]);
    ws['!cols'] = [{ wch: 32 }, { wch: 90 }];
    return ws;
  }
  function saveBook(title, fileName, sheets, rowCount, conds) {
    const wb = XLSX.utils.book_new();
    for (const [name, ws] of sheets) XLSX.utils.book_append_sheet(wb, ws, name);
    XLSX.utils.book_append_sheet(wb, conditionsSheet(title, rowCount, conds || conditions()), 'Điều kiện lọc');
    X._save(wb, fileName);
  }
  const EXPORTS = {
    detail: () => saveBook('Chi tiết các lần nhập dầu thô/nguyên liệu', 'BaoCao_ChiTietNhap', [['Chi tiết nhập', detailSheet(S.tableRecords)]], S.tableRecords.length, tableConditions()),
    product: (r) => saveBook('Tổng hợp nhập theo dầu thô/nguyên liệu', 'BaoCao_Nhap_TheoNguyenLieu', [
      ['Theo dầu thô-nguyên liệu', groupSheet(r, 'sp', 'Dầu thô/Nguyên liệu')], ['Theo nhóm', groupSheet(r, 'nhom', 'Nhóm')],
    ], r.length),
    time: (r) => saveBook('Tổng hợp nhập theo ngày / tháng', 'BaoCao_Nhap_TheoThoiGian', [
      ['Theo ngày', timeSheet(r, 'day')], ['Theo tháng', timeSheet(r, 'month')], ['Tháng x Nguyên liệu', pivotSheet(r)],
    ], r.length),
    all: (r) => saveBook('Báo cáo tổng hợp nhập dầu thô/nguyên liệu', 'BaoCao_TongHop_Nhap', [
      ['Theo dầu thô-nguyên liệu', groupSheet(r, 'sp', 'Dầu thô/Nguyên liệu')], ['Theo nhóm', groupSheet(r, 'nhom', 'Nhóm')],
      ['Theo tháng', timeSheet(r, 'month')], ['Tháng x Nguyên liệu', pivotSheet(r)], ['Theo ngày', timeSheet(r, 'day')], ['Chi tiết nhập', detailSheet(r)],
    ], r.length),
  };

  // ============================ SỰ KIỆN ============================
  $('#btnOpen').addEventListener('click', openPicker);
  $('#btnOpen2').addEventListener('click', openPicker);
  $('#btnRefresh').addEventListener('click', refresh);
  $('#btnReopen').addEventListener('click', refresh);
  $('#btnClearSaved').addEventListener('click', clearSaved);
  $('#fileInput').addEventListener('change', (e) => {
    const f = e.target.files[0];
    if (f) loadFile(f);
    e.target.value = '';
  });

  // Kéo thả file
  document.addEventListener('dragover', (e) => { e.preventDefault(); document.body.classList.add('dragging'); });
  document.addEventListener('dragleave', (e) => { if (!e.relatedTarget) document.body.classList.remove('dragging'); });
  document.addEventListener('drop', async (e) => {
    e.preventDefault();
    document.body.classList.remove('dragging');
    const item = e.dataTransfer.items && e.dataTransfer.items[0];
    const handleP = item && item.getAsFileSystemHandle ? item.getAsFileSystemHandle().catch(() => null) : null;
    const file = e.dataTransfer.files[0];
    if (!file) return;
    const handle = handleP ? await handleP : null;
    loadFile(file, { handle: handle && handle.kind === 'file' ? handle : null });
  });

  // Bộ lọc thời gian (ô ngày DD/MM/YYYY + bộ chọn ngày gốc)
  function setDate(which, day) {
    S.filters[which] = day;
    $('#fQuick').value = '';
    const f = S.filters;
    if (f.from != null && f.to != null && f.from > f.to) {
      if (which === 'from') f.to = f.from; else f.from = f.to;
      toast('"Từ ngày" không được sau "Đến ngày" — đã tự điều chỉnh.');
    }
    scheduleUpdate();
  }
  for (const [id, which] of [['fFrom', 'from'], ['fTo', 'to']]) {
    const inp = $('#' + id);
    inp.addEventListener('change', () => {
      const v = inp.value.trim();
      if (!v) { setDate(which, null); return; }
      const m = v.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/);
      const day = m ? L.parseDate(`${m[1]}/${m[2]}/${m[3]}`) : null;
      if (day == null) { inp.classList.add('bad'); toast('Ngày không hợp lệ — nhập theo dạng DD/MM/YYYY, ví dụ 01/10/2026.', true); return; }
      setDate(which, day);
    });
    const dp = $(`.dp[data-for="${id}"]`);
    dp.addEventListener('change', () => setDate(which, inputToDay(dp.value)));
    $(`.dp-btn[data-for="${id}"]`).addEventListener('click', () => {
      dp.value = dayToInput(S.filters[which] != null ? S.filters[which] : S.data && S.data.meta.maxDay);
      try { dp.showPicker(); } catch (e) { inp.focus(); }
    });
  }
  $('#fQuick').addEventListener('change', (e) => {
    const v = e.target.value, max = S.data && S.data.meta.maxDay;
    if (!v || max == null) return;
    const d = new Date(max * 86400000), y = d.getUTCFullYear(), m = d.getUTCMonth();
    const day = (yy, mm, dd) => Math.round(Date.UTC(yy, mm, dd) / 86400000);
    const ranges = {
      all: [null, null], today: [max, max], '7d': [max - 6, max], '30d': [max - 29, max],
      thisMonth: [day(y, m, 1), day(y, m + 1, 0)], prevMonth: [day(y, m - 1, 1), day(y, m, 0)],
      thisQuarter: [day(y, Math.floor(m / 3) * 3, 1), day(y, Math.floor(m / 3) * 3 + 3, 0)], thisYear: [day(y, 0, 1), day(y, 11, 31)],
    };
    [S.filters.from, S.filters.to] = ranges[v];
    scheduleUpdate();
  });

  $('#btnReset').addEventListener('click', () => {
    S.filters.from = S.filters.to = null;
    for (const d of DIMS) S.filters.sets[d].clear();
    $('#fQuick').value = '';
    S.drill = null; S.search = ''; $('#search').value = '';
    DIMS.forEach((d) => MS[d].renderButton());
    update();
    toast('Đã xóa bộ lọc.');
  });
  $('#applied').addEventListener('click', (e) => {
    const k = e.target.dataset.clear;
    if (!k) return;
    if (k === 'date') { S.filters.from = S.filters.to = null; $('#fQuick').value = ''; }
    else { S.filters.sets[k].clear(); MS[k].renderButton(); }
    scheduleUpdate();
  });

  // Điều khiển biểu đồ & xếp hạng
  $('#segGran').addEventListener('click', (e) => { const v = e.target.dataset.v; if (v) { S.gran = v; renderCharts(); } });
  $('#rankTable').addEventListener('click', (e) => {
    const tr = e.target.closest('tr[data-key]');
    if (tr) setDrill(`Dầu thô/Nguyên liệu = ${tr.dataset.key}`, (r) => r.sp === tr.dataset.key);
  });

  // Chất lượng dữ liệu → xem bản ghi liên quan (trên toàn bộ dữ liệu)
  $('#quality').addEventListener('click', (e) => {
    const k = e.target.dataset.issue;
    if (!k) return;
    const ids = new Set(S.data.issues[k]);
    if (k === 'dupDoc') S.sort = { key: 'doc', dir: 1 };
    setDrill(ISSUE_LABEL[k], (r) => ids.has(r.id), 'all');
  });

  // Bảng chi tiết
  let searchTimer;
  $('#search').addEventListener('input', (e) => { clearTimeout(searchTimer); searchTimer = setTimeout(() => { S.search = e.target.value; S.page = 1; renderTable(); }, 200); });
  $('#pageSize').addEventListener('change', (e) => { S.pageSize = +e.target.value; S.page = 1; renderTable(); });
  $('#btnClearTable').addEventListener('click', () => { S.drill = null; S.search = ''; $('#search').value = ''; S.page = 1; renderTable(); });
  $('#detailCard').addEventListener('click', (e) => {
    if (e.target.id === 'btnDrillOff') { S.drill = null; S.page = 1; renderTable(); return; }
    const th = e.target.closest('th[data-sort]');
    if (th) {
      const k = th.dataset.sort;
      S.sort = S.sort.key === k ? { key: k, dir: -S.sort.dir } : { key: k, dir: k === 'day' || k === 'vol' ? -1 : 1 };
      renderTable();
      return;
    }
    const pb = e.target.closest('button[data-page]');
    if (pb && !pb.disabled) { S.page = +pb.dataset.page; renderTable(); $('#detailCard').scrollIntoView({ block: 'start' }); }
  });

  // Xuất báo cáo
  document.querySelector('.export').addEventListener('click', (e) => {
    const b = e.target.closest('[data-exp]');
    if (!b || !S.data) return;
    const kind = b.dataset.exp;
    if (kind === 'print') { window.print(); return; }
    const n = kind === 'detail' ? S.tableRecords.length : S.filtered.length;
    if (!n) { toast('Không có dữ liệu phù hợp với bộ lọc để xuất.', true); return; }
    loading(true, 'Đang tạo file Excel…');
    setTimeout(() => {
      try { EXPORTS[kind](S.filtered); toast(`Đã xuất báo cáo (${fmtInt(n)} bản ghi). Kiểm tra thư mục Tải xuống.`); }
      catch (err) { console.error(err); toast('Lỗi khi xuất file: ' + err.message, true); }
      finally { loading(false); }
    }, 30);
  });
  $('#btnPrint').addEventListener('click', () => { if (S.data) window.print(); else toast('Chưa có dữ liệu để in.', true); });
  window.addEventListener('beforeprint', () => C.resizeAll());
  window.addEventListener('afterprint', () => C.resizeAll());

  // Hướng dẫn
  $('#btnHelp').addEventListener('click', () => dialog(`<h2>Hướng dẫn — Nhập dầu thô/nguyên liệu</h2>
    <h4>1. Nạp / cập nhật dữ liệu</h4>
    <ul><li>Mở bằng <b>Mo_Dashboard.bat</b>: trang tự đọc file <b>IMPORT.xlsx</b> trong thư mục dashboard. Cập nhật file rồi nhấn <b>Làm mới dữ liệu</b>.</li>
    <li>Hoặc nhấn <b>Chọn file Excel</b> (kéo thả cũng được) và chọn file nhập từ SAP.</li>
    <li>Cột sử dụng: <b>N</b> = Ngày nhập (Posting Date), <b>D</b> = Dầu thô/Nguyên liệu (Material Description), <b>I</b> = Số lượng (Quantity), đơn vị <b>MT</b>.</li></ul>
    <h4>2. Lọc & xem số liệu</h4>
    <ul><li>Lọc theo thời gian, <b>Nhóm</b> (Dầu thô / Nguyên liệu khác) và từng loại <b>Dầu thô/Nguyên liệu</b>.</li>
    <li>Nhấn vào cột, điểm, lát biểu đồ hoặc dòng xếp hạng để xem các lần nhập liên quan ở bảng chi tiết.</li>
    <li>Nhóm "Dầu thô" gồm các loại có tên bắt đầu bằng "Dầu thô"; các loại còn lại thuộc "Nguyên liệu khác".</li></ul>
    <h4>3. Chuyển trang</h4>
    <ul><li>Nhấn vào tiêu đề dashboard (góc trên bên trái) để chuyển giữa <b>Xuất bán sản phẩm</b> và <b>Nhập dầu thô/nguyên liệu</b>.</li></ul>
    <h4>4. Trích xuất báo cáo</h4>
    <ul><li>Các nút xuất Excel ở cuối trang theo đúng bộ lọc hiện tại, kèm sheet <b>Điều kiện lọc</b>. <b>In / PDF</b>: chọn khổ ngang.</li></ul>`));

  // ============================ KHỞI ĐỘNG ============================
  C.setupDefaults();
  (async function boot() {
    if (viaServer && (await loadFromServer(false))) return;
    const [last, handle] = await Promise.all([ST.get(CACHE_KEY), ST.get(HANDLE_KEY)]);
    if (last && last.buf) {
      const ok = await loadBuffer(last.buf, last.name, { fileTime: last.fileTime });
      if (ok) {
        S.source = { name: last.name, handle: handle || null, cachedAt: last.savedAt };
        renderSource();
        return;
      }
    }
    if (handle) { S.source = { name: handle.name, handle }; $('#btnReopen').hidden = false; $('#btnReopen').textContent = `Mở lại "${handle.name}"`; }
  })();
})();
