/* =====================================================================
 * app.js — GIAO DIỆN & ĐIỀU PHỐI
 * Luồng: nạp file → bộ lọc → (analytics) → KPI / biểu đồ / bảng → xuất báo cáo
 * ===================================================================== */
(function () {
  'use strict';
  const { loader: L, analytics: A, charts: C, exporter: X, storage: ST } = window.PXK;
  const $ = (s, el = document) => el.querySelector(s);
  const $$ = (s, el = document) => [...el.querySelectorAll(s)];
  const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const fmt = C.fmt, fmtInt = C.fmtInt;
  const fmt2 = (v) => fmt(v, 2);
  const DIMS = A.DIMS;

  // ------------------------------ Trạng thái ------------------------------
  const S = {
    data: null,
    source: null,                    // { name, handle, file, cachedAt }
    filters: { from: null, to: null, nonZero: true, sets: Object.fromEntries(DIMS.map((d) => [d, new Set()])) },
    filtered: [],
    unit: 'M3', gran: 'month', shareDim: 'sp', rankDim: 'kh', vehTop: 15,
    drill: null,                     // { label, pred, scope: 'filtered' | 'all' }
    search: '', sort: { key: 'day', dir: -1 }, page: 1, pageSize: 50,
    tableRecords: [],
    options: {},                     // dim -> [values] (theo thứ tự)
  };
  const isDeleted = (v) => L.fold(v) === 'da xoa';
  const soLabel = () => (S.data ? S.data.meta.labels.so.toLowerCase() : 'số phiếu');

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

  window.PXK.flagText = function (r) {
    const f = r.flags, out = [];
    if (f.htDefault) out.push(`HTVC trống → mặc định "${f.htDefault.value}" (${f.htDefault.desc})`);
    if (f.dupVoucher) out.push('Trùng ' + soLabel());
    if (f.dupRow) out.push('Trùng nội dung');
    if (f.volError) out.push(f.volError);
    if (f.dateError) out.push(f.dateError);
    if (f.unitInferred) out.push('Đơn vị suy theo sản phẩm');
    if (f.unitMismatch) out.push('Đơn vị khác quy tắc ' + S.data.meta.unitRule);
    if (f.negative) out.push('Khối lượng âm');
    return out.join('; ');
  };

  // ============================ MULTI-SELECT ============================
  class MultiSelect {
    constructor(host, dim) {
      this.dim = dim;
      this.host = host;
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
      this.panel.innerHTML = `<input type="search" placeholder="Tìm ${esc(A.DIM_LABELS[dim].toLowerCase())}…">
        <div class="ms-actions"><button type="button" class="btn-link" data-a="all">Chọn tất cả (đang hiện)</button><button type="button" class="btn-link" data-a="none">Bỏ chọn</button></div>
        <div class="ms-list" role="listbox" aria-multiselectable="true"></div>`;
      host.appendChild(this.panel);
      this.q = $('input', this.panel);
      this.list = $('.ms-list', this.panel);
      this.btn.addEventListener('click', (e) => { e.stopPropagation(); this.toggle(); });
      this.panel.addEventListener('click', (e) => e.stopPropagation());
      this.q.addEventListener('input', () => this.renderList());
      this.panel.addEventListener('click', (e) => {
        const a = e.target.dataset.a;
        if (!a) return;
        const set = S.filters.sets[dim];
        if (a === 'none') set.clear();
        else this.visible().forEach((v) => set.add(v));
        this.changed();
      });
      this.list.addEventListener('change', (e) => {
        const cb = e.target;
        const set = S.filters.sets[dim];
        cb.checked ? set.add(cb.value) : set.delete(cb.value);
        this.changed();
      });
    }
    toggle(force) {
      const open = force != null ? force : this.panel.hidden;
      MultiSelect.all.forEach((m) => m !== this && (m.panel.hidden = true));
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
      let txt = 'Tất cả';
      if (set.size === 1) txt = [...set][0];
      else if (set.size > 1) txt = `Đã chọn ${set.size}/${total}`;
      $('span', this.btn).textContent = txt;
      this.btn.title = set.size ? [...set].join('\n') : 'Tất cả';
      this.btn.classList.toggle('active', set.size > 0);
    }
    setCounts(map) { this.counts = map; if (!this.panel.hidden) this.renderList(); }
  }
  MultiSelect.all = [];
  document.addEventListener('click', () => MultiSelect.all.forEach((m) => (m.panel.hidden = true)));
  document.addEventListener('keydown', (e) => { if (e.key === 'Escape') MultiSelect.all.forEach((m) => (m.panel.hidden = true)); });

  const MS = {};
  $$('[data-ms]').forEach((host) => {
    const m = new MultiSelect(host, host.dataset.ms);
    MS[host.dataset.ms] = m;
    MultiSelect.all.push(m);
  });

  // ============================ NẠP DỮ LIỆU ============================

  function openPicker() {
    const inp = $('#fileInput');
    inp.value = '';
    inp.click();
  }

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
    S.source = { name: file.name, handle, file, fileTime: file.lastModified };
    ST.set(CACHE_KEY, { name: file.name, fileTime: file.lastModified, savedAt: Date.now(), buf });
    if (handle) ST.set('handle', handle); else ST.del('handle');
    renderSource();
    toast(`Đã nạp ${fmtInt(S.data.records.length)} dòng từ "${file.name}".`);
    return true;
  }

  async function loadBuffer(buf, name, { keepFilters = false, fileTime = null } = {}) {
    loading(true, `Đang đọc "${name}"…`);
    await new Promise((r) => setTimeout(r, 40));
    try {
      const data = L.parseWorkbook(new Uint8Array(buf), name);
      if (!data.records.length) throw new L.StructureError('File không có dòng dữ liệu nào bên dưới dòng tiêu đề.');
      data.meta.fileTime = fileTime;
      const sameFile = S.data && S.data.meta.fileName === name;
      S.data = data;
      initAfterLoad(keepFilters || sameFile);
      return true;
    } catch (e) {
      console.error(e);
      const isStruct = e instanceof L.StructureError;
      dialog(`<h2>${isStruct ? 'File không đúng cấu trúc' : 'Lỗi khi đọc file'}</h2>
        <p class="err">${esc(e.message)}</p>
        <h4>Các định dạng được hỗ trợ</h4>
        ${L.PROFILES.map((p) => `<p><b>${esc(p.name)}</b></p><ul>${Object.values(L.requiredOf(p)).map((f) => `<li>Cột <b>${f.col}</b>: ${esc(f.label)} <span class="muted">(tiêu đề "${esc(f.aliases[0])}")</span></li>`).join('')}</ul>`).join('')}
        ${S.data ? '<p>Dữ liệu cũ vẫn được giữ nguyên.</p>' : ''}`);
      return false;
    } finally { loading(false); }
  }

  // ------- Đọc thẳng file dữ liệu mặc định khi mở qua Mo_Dashboard.bat (máy chủ cục bộ 127.0.0.1) -------
  const DEFAULT_FILE = 'EXPORT.xlsx';
  const CACHE_KEY = 'last-v2';
  // Máy chủ cục bộ (Mo_Dashboard.bat): tự đọc EXPORT.xlsx trong thư mục.
  // Bản đưa lên mạng (GitHub Pages...): không chứa dữ liệu — người dùng tự chọn file, xử lý ngay trong trình duyệt.
  const viaServer = /^(127\.0\.0\.1|localhost|\[::1\])$/.test(location.hostname);

  async function loadFromServer(keepFilters) {
    let res;
    try { res = await fetch(encodeURIComponent(DEFAULT_FILE) + '?t=' + Date.now(), { cache: 'no-store' }); } catch (e) { res = null; }
    if (!res || !res.ok) {
      toast(`Không tìm thấy file "${DEFAULT_FILE}" trong thư mục dashboard. Hãy nhấn "Chọn file Excel".`, true);
      return false;
    }
    const lm = res.headers.get('Last-Modified');
    const fileTime = lm ? Date.parse(lm) : null;
    const ok = await loadBuffer(await res.arrayBuffer(), DEFAULT_FILE, { keepFilters, fileTime });
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
      } catch (e) { console.warn(e); src.handle = null; ST.del('handle'); }
    }
    if (src && src.file) { if (await loadFile(src.file, { keepFilters: true })) return; }
    toast('Trình duyệt không cho đọc lại file tự động. Hãy nhấn "Chọn file Excel" và chọn file mới nhất.', true);
  }

  async function clearSaved() {
    if (!confirm('Xóa dữ liệu Excel đã lưu trong trình duyệt? Dashboard sẽ trở về màn hình chọn file.')) return;
    await ST.del(CACHE_KEY);
    await ST.del('last');
    await ST.del('handle');
    location.reload();
  }

  function renderSource() {
    const m = S.data && S.data.meta;
    if (!m) { $('#sourceInfo').textContent = 'Chưa nạp dữ liệu'; return; }
    const ft = m.fileTime ? new Date(m.fileTime).toLocaleString('vi-VN') : '—';
    const cached = S.source && S.source.server ? ' · đọc trực tiếp từ thư mục dashboard'
      : S.source && S.source.cachedAt ? ` · bản lưu trong trình duyệt (chọn lại file Excel để cập nhật số liệu)` : '';
    $('#sourceInfo').textContent = `Nguồn: ${m.fileName} · sheet "${m.sheet}" · file sửa lúc ${ft} · nạp lúc ${m.loadedAt.toLocaleString('vi-VN')}${cached}`;
    $('#foot').textContent = `Dữ liệu đọc từ ${m.fileName} (${m.profileName}), sheet "${m.sheet}", tiêu đề ở dòng ${m.headerRow}. Cột: ${m.columnList.map((c) => `${c.label}=${c.col}`).join(', ')}. Quy tắc đơn vị: ${m.unitRule}.`;
  }

  function initAfterLoad(keepFilters) {
    const recs = S.data.records;
    C.setColorRanks(recs);
    // Danh sách lựa chọn bộ lọc: sắp xếp theo số giao dịch giảm dần
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
      setDefaultStatus();
      S.drill = null; S.search = ''; $('#search').value = '';
      const span = (S.data.meta.maxDay || 0) - (S.data.meta.minDay || 0);
      S.gran = span > 120 ? 'month' : span > 31 ? 'week' : 'day';
    }
    $$('.dp').forEach((dp) => { dp.min = dayToInput(S.data.meta.minDay); dp.max = dayToInput(S.data.meta.maxDay); });
    $('[data-ms="tt"]').hidden = !S.data.meta.hasStatus;
    // Nhãn theo định dạng file (PXK: Tình trạng phiếu / Số phiếu; EXPORT: Trạng thái giao hàng / Số LXH trên SAP)
    A.DIM_LABELS.tt = S.data.meta.labels.tt;
    $('#ttLabel').textContent = S.data.meta.labels.tt;
    COLS.find((c) => c.k === 'so').h = S.data.meta.labels.so;
    COLS.find((c) => c.k === 'tt').h = S.data.meta.labels.tt;
    $('#nonZeroWrap').hidden = !S.data.issues.zeroVol.length;
    if (!S.data.meta.units.includes(S.unit)) S.unit = S.data.meta.units[0] || 'M3';
    if (S.unitPref && !S.data.meta.units.includes(S.unitPref)) S.unitPref = null;
    $('#welcome').hidden = true;
    $('#dash').hidden = false;
    DIMS.forEach((d) => MS[d].renderButton());
    renderSource();
    renderQuality();
    update();
  }

  function setDefaultStatus() {
    const set = S.filters.sets.tt;
    set.clear();
    const opts = S.options.tt || [];
    if (opts.some(isDeleted)) opts.filter((v) => !isDeleted(v)).forEach((v) => set.add(v));
  }

  // ============================ CẬP NHẬT ============================
  let raf = 0;
  function scheduleUpdate() { cancelAnimationFrame(raf); raf = requestAnimationFrame(update); }

  function update() {
    if (!S.data) return;
    const recs = S.data.records;
    S.filtered = A.applyFilters(recs, S.filters);
    // Đếm số giao dịch cho mỗi lựa chọn (theo các bộ lọc khác)
    for (const d of DIMS) {
      const sub = A.applyFilters(recs, S.filters, d);
      const m = new Map();
      for (const r of sub) m.set(r[d], (m.get(r[d]) || 0) + 1);
      MS[d].setCounts(m);
    }
    for (const [id, v] of [['fFrom', S.filters.from], ['fTo', S.filters.to]]) {
      $('#' + id).value = L.dayToStr(v);
      $('#' + id).classList.remove('bad');
    }
    renderApplied();
    renderKpis();
    renderUnitSeg();
    renderCharts();
    renderRank();
    S.page = 1;
    renderTable();
  }

  // ----------------------- Điều kiện lọc đang áp dụng -----------------------
  function conditions() {
    const f = S.filters, m = S.data.meta, out = [];
    out.push({ label: 'Thời gian', value: f.from == null && f.to == null ? `Toàn bộ (${L.dayToStr(m.minDay)} – ${L.dayToStr(m.maxDay)})` : `${f.from != null ? L.dayToStr(f.from) : 'đầu kỳ'} – ${f.to != null ? L.dayToStr(f.to) : 'cuối kỳ'}` });
    if (S.data.issues.zeroVol.length) out.push({ label: 'Khối lượng', value: f.nonZero ? 'Chỉ tính giao dịch có khối lượng > 0' : 'Tính cả giao dịch có khối lượng = 0' });
    for (const d of DIMS) {
      if (d === 'tt' && !m.hasStatus) continue;
      const set = f.sets[d];
      let v = 'Tất cả';
      if (set.size) {
        const excluded = S.options[d].filter((x) => !set.has(x));
        v = excluded.length && excluded.length < set.size ? `Tất cả trừ: ${excluded.join(', ')}` : [...set].join(', ');
      }
      out.push({ label: A.DIM_LABELS[d], value: v });
    }
    return out;
  }
  function tableConditions() {
    const c = conditions();
    if (S.drill) c.push({ label: 'Xem chi tiết', value: S.drill.label + (S.drill.scope === 'all' ? ' (trên toàn bộ dữ liệu, bỏ qua bộ lọc)' : '') });
    if (S.search) c.push({ label: 'Từ khóa tìm kiếm', value: S.search });
    const s = { day: 'Ngày xuất', vol: 'Khối lượng' }[S.sort.key] || S.sort.key;
    c.push({ label: 'Sắp xếp', value: `${s} ${S.sort.dir > 0 ? 'tăng dần' : 'giảm dần'}` });
    return c;
  }

  function renderApplied() {
    const f = S.filters, m = S.data.meta;
    const chips = [];
    if (f.from == null && f.to == null) chips.push(`<span class="chip fixed"><span>Thời gian: toàn bộ (${L.dayToStr(m.minDay)} – ${L.dayToStr(m.maxDay)})</span></span>`);
    else chips.push(`<span class="chip"><span>Thời gian: ${f.from != null ? L.dayToStr(f.from) : 'đầu kỳ'} – ${f.to != null ? L.dayToStr(f.to) : 'cuối kỳ'}</span><button data-clear="date" title="Bỏ lọc thời gian">×</button></span>`);
    const zero = S.data.issues.zeroVol.length;
    if (zero && f.nonZero) chips.push(`<span class="chip warn" title="Lệnh đã lập nhưng khối lượng thực tế = 0"><span>Bỏ qua ${fmtInt(zero)} dòng KL = 0</span><button data-clear="nonZero" title="Tính cả dòng KL = 0">×</button></span>`);
    for (const d of DIMS) {
      const set = f.sets[d];
      if (!set.size) continue;
      const vals = [...set];
      let txt;
      const excluded = S.options[d].filter((x) => !set.has(x));
      if (excluded.length && excluded.length < vals.length) txt = `trừ ${excluded.join(', ')}`;
      else txt = vals.slice(0, 3).join(', ') + (vals.length > 3 ? ` (+${vals.length - 3})` : '');
      const warn = d === 'tt' && excluded.some(isDeleted);
      chips.push(`<span class="chip${warn ? ' warn' : ''}" title="${esc(vals.join('\n'))}"><span>${esc(A.DIM_LABELS[d])}: ${esc(txt)}</span><button data-clear="${d}" title="Bỏ lọc ${esc(A.DIM_LABELS[d])}">×</button></span>`);
    }
    $('#applied').innerHTML = `<span class="lbl">Đang áp dụng:</span>${chips.join('')}<span class="lbl" style="margin-left:auto">${fmtInt(S.filtered.length)} / ${fmtInt(S.data.records.length)} giao dịch</span>`;
  }

  // ----------------------------- Chất lượng -----------------------------
  function renderQuality() {
    const { issues: q, meta: m, records } = S.data;
    const delIds = records.filter((r) => isDeleted(r.tt)).map((r) => r.id);
    const items = [];
    const link = (key, n) => (n ? ` <button class="btn-link" data-issue="${key}">Xem</button>` : '');
    let warnCount = -1; // dòng đầu tiên là thông tin chung
    const li = (cls, html) => { if (!cls) warnCount++; items.push(`<li${cls ? ` class="${cls}"` : ''}>${html}</li>`); };
    li('', `Đã đọc <b>${fmtInt(records.length)}</b> dòng từ sheet "${esc(m.sheet)}" (tiêu đề dòng ${m.headerRow}).`);
    if (delIds.length) li('', `<b>${fmtInt(delIds.length)}</b> phiếu có tình trạng "Đã xóa": <b>mặc định không tính</b> vào số liệu — có thể chọn lại ở bộ lọc "Tình trạng phiếu".${link('deleted', delIds.length)}`);
    if (q.summaryRows) li('', `${fmtInt(q.summaryRows)} dòng tổng cộng/ghi chú cuối bảng (không có ngày, khách hàng, sản phẩm) — đã bỏ qua.`);
    if (q.zeroVol.length) li('', `<b>${fmtInt(q.zeroVol.length)}</b> dòng có khối lượng thực tế = 0 (lệnh đã lập, chưa xuất) — <b>mặc định không tính</b>; bỏ chọn "Chỉ tính GD có KL &gt; 0" để tính.${link('zeroVol', 1)}`);
    if (q.dupVoucher.length) li('', `<b>${fmtInt(q.dupVoucherGroups)}</b> ${esc(soLabel())} bị trùng (${fmtInt(q.dupVoucher.length)} dòng) — chưa loại bỏ, cần kiểm tra.${link('dupVoucher', 1)}`);
    if (q.dupRow.length) li('', `<b>${fmtInt(q.dupRowGroups)}</b> nhóm dòng trùng toàn bộ ngày/kho/khách hàng/HTVC/phương tiện/sản phẩm/khối lượng và cùng ${esc(soLabel())} (${fmtInt(q.dupRow.length)} dòng) — vẫn tính đủ vào số liệu, cần kiểm tra có nhập trùng không. Các dòng giống nhau nhưng khác ${esc(soLabel())} là các chuyến riêng, không báo trùng.${link('dupRow', 1)}`);
    for (const rule of L.HT_RULES) {
      const n = q.htRule[rule.id].length;
      if (n) li('', `<b>${fmtInt(n)}</b> dòng trống Hình thức vận chuyển, ${esc(rule.desc.charAt(0).toLowerCase() + rule.desc.slice(1))} — <b>mặc định "${esc(rule.value)}"</b> theo quy tắc nghiệp vụ.${link('htRule_' + rule.id, 1)}`);
    }
    if (q.missingHt.length) li('', `<b>${fmtInt(q.missingHt.length)}</b> dòng thiếu Hình thức vận chuyển không thuộc quy tắc mặc định — hiển thị là "${L.EMPTY}".${link('missingHt', 1)}`);
    if (q.missingPt.length) li('', `<b>${fmtInt(q.missingPt.length)}</b> dòng thiếu Phương tiện.${link('missingPt', 1)}`);
    if (q.missingSp.length) li('', `<b>${fmtInt(q.missingSp.length)}</b> dòng thiếu Sản phẩm (đơn vị lấy theo ghi chú trong ô khối lượng).${link('missingSp', 1)}`);
    if (q.missingKh.length) li('', `<b>${fmtInt(q.missingKh.length)}</b> dòng thiếu Khách hàng.${link('missingKh', 1)}`);
    li(q.dateMissing.length ? '' : 'ok', q.dateMissing.length ? `<b>${fmtInt(q.dateMissing.length)}</b> dòng thiếu/sai ngày — không thuộc kỳ lọc nào và bị loại khi lọc theo thời gian.${link('dateMissing', 1)}` : '✓ Ngày xuất: tất cả hợp lệ.');
    li(q.volInvalid.length ? '' : 'ok', q.volInvalid.length ? `<b>${fmtInt(q.volInvalid.length)}</b> dòng thiếu/sai khối lượng — vẫn đếm là giao dịch nhưng không cộng vào tổng.${link('volInvalid', 1)}` : '✓ Khối lượng: tất cả đọc được số và đơn vị.');
    li(q.unitMismatch.length ? '' : 'ok', q.unitMismatch.length ? `<b>${fmtInt(q.unitMismatch.length)}</b> dòng có đơn vị khác quy tắc (${esc(m.unitRule)}) — giữ theo đơn vị ghi trong file.${link('unitMismatch', 1)}` : `✓ Đơn vị theo quy tắc: ${esc(m.unitRule)}.`);
    if (q.unitInferred.length) li('', `<b>${fmtInt(q.unitInferred.length)}</b> dòng không ghi đơn vị — đã suy theo quy tắc sản phẩm.${link('unitInferred', 1)}`);
    if (q.negative.length) li('', `<b>${fmtInt(q.negative.length)}</b> dòng khối lượng âm — giữ nguyên (có thể là điều chỉnh nghiệp vụ).${link('negative', 1)}`);
    if (q.whitespaceFixed) li('', `${fmtInt(q.whitespaceFixed)} dòng có khoảng trắng thừa trong tên — đã chuẩn hóa khi phân tích (dữ liệu gốc giữ nguyên khi xuất).`);
    const merged = Object.entries(q.merged).filter(([, v]) => v.length);
    if (merged.length) li('', 'Gộp biến thể tên khác hoa/thường: ' + merged.map(([d, groups]) => `${A.DIM_LABELS[d]}: ${groups.map((g) => g.map(esc).join(' = ')).join('; ')}`).join(' · '));
    const notes = m.notes.length ? `<p class="q-notes">${m.notes.map(esc).join('<br>')}</p>` : '';
    $('#quality').innerHTML = `<details class="q-box"><summary>Kiểm tra dữ liệu: ${warnCount > 0 ? `${warnCount} điểm cần lưu ý` : 'không phát hiện vấn đề'} — nhấn để xem</summary><ul>${items.join('')}</ul>${notes}</details>`;
    S.issueSets = { ...q, deleted: delIds };
    for (const rule of L.HT_RULES) {
      S.issueSets['htRule_' + rule.id] = q.htRule[rule.id];
      ISSUE_LABEL['htRule_' + rule.id] = `HTVC trống → mặc định "${rule.value}" (${rule.desc})`;
    }
  }
  const ISSUE_LABEL = {
    deleted: 'Phiếu có tình trạng "Đã xóa"', dupVoucher: 'Trùng số chứng từ', zeroVol: 'Khối lượng thực tế = 0', dupRow: 'Dòng trùng toàn bộ nội dung', missingHt: 'Thiếu hình thức vận chuyển', htDefault: 'HTVC được gán mặc định',
    missingPt: 'Thiếu phương tiện', missingSp: 'Thiếu sản phẩm', missingKh: 'Thiếu khách hàng', dateMissing: 'Thiếu/sai ngày', volInvalid: 'Thiếu/sai khối lượng',
    unitMismatch: 'Đơn vị khác quy tắc', unitInferred: 'Đơn vị suy theo sản phẩm', negative: 'Khối lượng âm',
  };

  // -------------------------------- KPI --------------------------------
  function renderKpis() {
    const k = A.kpis(S.filtered);
    const units = S.data.meta.units;
    const unitLines = (fn) => units.map((u) => `<div class="k-v num">${fn(u)}<small>${u}</small></div>`).join('');
    const invalidNote = k.invalid ? `<div class="k-n warn">${fmtInt(k.invalid)} GD thiếu/sai KL không cộng</div>` : '<div class="k-n">Tách riêng theo đơn vị, không cộng lẫn</div>';
    const byUnitCount = units.map((u) => `${u}: ${fmtInt(k.totals[u] ? k.totals[u].count : 0)}`).join(' · ');
    $('#kpis').innerHTML = [
      `<div class="kpi main"><div class="k-t">Tổng khối lượng thực tế</div><div class="k-u">${unitLines((u) => fmt2(k.totals[u] ? k.totals[u].sum : 0))}</div>${invalidNote}</div>`,
      `<div class="kpi"><div class="k-t">Số giao dịch xuất hàng</div><div class="k-v num">${fmtInt(k.count)}</div><div class="k-n">${byUnitCount}</div></div>`,
      `<div class="kpi"><div class="k-t">Số khách hàng</div><div class="k-v num">${fmtInt(k.customers)}</div><div class="k-n">Khách hàng duy nhất</div></div>`,
      `<div class="kpi"><div class="k-t">Số sản phẩm</div><div class="k-v num">${fmtInt(k.products)}</div><div class="k-n">Sản phẩm duy nhất</div></div>`,
      `<div class="kpi"><div class="k-t">Số phương tiện</div><div class="k-v num">${fmtInt(k.vehicles)}</div><div class="k-n">Phương tiện duy nhất (không tính ô trống)</div></div>`,
      `<div class="kpi"><div class="k-t">Bình quân / giao dịch</div><div class="k-u">${unitLines((u) => (k.totals[u] && k.totals[u].count ? fmt2(k.totals[u].sum / k.totals[u].count) : '—'))}</div><div class="k-n">Theo từng đơn vị</div></div>`,
    ].join('');
  }

  // ------------------------------ Biểu đồ ------------------------------
  function renderUnitSeg() {
    const units = S.data.meta.units;
    const tot = A.totalsByUnit(S.filtered).units;
    // Giữ đơn vị người dùng chọn; chỉ tạm chuyển khi đơn vị đó không có dữ liệu theo bộ lọc
    const has = (u) => tot[u] && tot[u].count;
    const pref = S.unitPref && units.includes(S.unitPref) ? S.unitPref : units[0];
    S.unit = has(pref) ? pref : units.find(has) || pref;
    $('#segUnit').innerHTML = units.map((u) => `<button data-v="${esc(u)}" class="${u === S.unit ? 'on' : ''}" ${tot[u] && tot[u].count ? '' : 'disabled title="Không có dữ liệu đơn vị này theo bộ lọc"'}>Đơn vị ${esc(u)}</button>`).join('');
    const sub = `Đơn vị: ${S.unit} · ${fmtInt(tot[S.unit] ? tot[S.unit].count : 0)} giao dịch`;
    $$('.unit-sub').forEach((el) => (el.textContent = sub));
    $$('.unit-txt').forEach((el) => (el.textContent = `Đơn vị: ${S.unit}`));
    $('#trendSub').textContent = `${sub} · theo ${{ day: 'ngày', week: 'tuần', month: 'tháng' }[S.gran]}`;
    $$('#segGran button').forEach((b) => b.classList.toggle('on', b.dataset.v === S.gran));
    $$('#segShare button').forEach((b) => b.classList.toggle('on', b.dataset.v === S.shareDim));
    $$('#segRank button').forEach((b) => b.classList.toggle('on', b.dataset.v === S.rankDim));
  }

  const unitPred = (u) => (r) => r.volOk && r.unit === u;
  function drillDim(dim, group) {
    const u = S.unit;
    if (group.other) {
      const set = new Set(group.members);
      setDrill(`${A.DIM_LABELS[dim]}: nhóm "Khác" (${group.members.length} mục) · ${u}`, (r) => set.has(r[dim]) && unitPred(u)(r));
    } else setDrill(`${A.DIM_LABELS[dim]} = ${group.key} · ${u}`, (r) => r[dim] === group.key && unitPred(u)(r));
  }

  function renderCharts() {
    const recs = S.filtered, u = S.unit;
    C.trend(recs, u, S.gran, (p) => {
      if (!p) return;
      const lbl = S.gran === 'day' ? `Ngày ${L.dayToStr(p.start)}` : S.gran === 'week' ? p.longLabel : `Tháng ${p.label}`;
      setDrill(`${lbl} · ${u}`, (r) => r.day != null && r.day >= p.start && r.day <= p.end && unitPred(u)(r));
    });
    C.bars('chCustomer', A.groupBy(recs, 'kh', u).slice(0, 10), { unit: u, horizontal: true, catTitle: 'Khách hàng', onClick: (g) => drillDim('kh', g) });
    C.bars('chProduct', A.groupBy(recs, 'sp', u), { unit: u, catTitle: 'Sản phẩm', onClick: (g) => drillDim('sp', g) });
    C.transportStacked(recs, u, (ht, sp, topList) => {
      if (sp) setDrill(`HTVC = ${ht} · Sản phẩm = ${sp} · ${u}`, (r) => r.ht === ht && r.sp === sp && unitPred(u)(r));
      else { const top = new Set(topList); setDrill(`HTVC = ${ht} · Sản phẩm khác · ${u}`, (r) => r.ht === ht && !top.has(r.sp) && unitPred(u)(r)); }
    });
    C.bars('chVehicle', A.groupBy(recs, 'pt', u).slice(0, S.vehTop), { unit: u, catTitle: 'Phương tiện', onClick: (g) => drillDim('pt', g) });
    C.donut(recs, u, S.shareDim, (g) => drillDim(S.shareDim, g));
  }

  // --------------------------- Bảng xếp hạng ---------------------------
  function renderRank() {
    const dim = S.rankDim, u = S.unit;
    const units = S.data.meta.units;
    const others = units.filter((x) => x !== u);
    const groups = A.groupMulti(S.filtered, dim);
    const val = (g, x) => (g.byUnit[x] ? g.byUnit[x].sum : 0);
    groups.sort((a, b) => val(b, u) - val(a, u) || others.reduce((s, x) => s + val(b, x) - val(a, x), 0) || b.count - a.count);
    const tot = A.totalsByUnit(S.filtered).units;
    const max = groups.length ? val(groups[0], u) : 0;
    const head = `<thead><tr><th>#</th><th>${esc(A.DIM_LABELS[dim])}</th><th class="r">Số GD</th><th class="r">KL (${u})</th><th class="r">Tỷ trọng</th>${others.map((x) => `<th class="r">KL (${x})</th>`).join('')}</tr></thead>`;
    if (!groups.length) { $('#rankTable').innerHTML = head + `<tbody><tr class="empty-row"><td colspan="${5 + others.length}">Không có dữ liệu phù hợp với bộ lọc</td></tr></tbody>`; return; }
    const rows = groups.map((g, i) => {
      const v = val(g, u);
      const pct = tot[u] && tot[u].sum ? (v / tot[u].sum) * 100 : 0;
      return `<tr data-key="${esc(g.key)}"><td class="muted">${i + 1}</td><td class="name">${esc(g.key)}${max > 0 && v > 0 ? `<div class="bar" style="width:${Math.max(1, (v / max) * 100).toFixed(1)}%"></div>` : ''}</td>
        <td class="r">${fmtInt(g.count)}</td><td class="r">${v ? fmt(v) : '<span class="muted">—</span>'}</td><td class="r">${v ? pct.toLocaleString('vi-VN', { maximumFractionDigits: 2 }) + '%' : '<span class="muted">—</span>'}</td>
        ${others.map((x) => `<td class="r">${val(g, x) ? fmt(val(g, x)) : '<span class="muted">—</span>'}</td>`).join('')}</tr>`;
    }).join('');
    $('#rankTable').innerHTML = head + `<tbody>${rows}</tbody>`;
  }

  // ---------------------------- Bảng chi tiết ----------------------------
  function setDrill(label, pred, scope = 'filtered') {
    S.drill = { label, pred, scope };
    S.page = 1;
    renderTable();
    $('#detailCard').scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  const COLS = [
    { k: 'day', h: 'Ngày xuất', sort: true },
    { k: 'kho', h: 'Kho', sort: true },
    { k: 'kh', h: 'Khách hàng', sort: true },
    { k: 'ht', h: 'Hình thức vận chuyển', sort: true },
    { k: 'pt', h: 'Phương tiện', sort: true },
    { k: 'sp', h: 'Sản phẩm', sort: true },
    { k: 'vol', h: 'Khối lượng thực tế', sort: true, r: true },
    { k: 'unit', h: 'ĐV', sort: true },
    { k: 'so', h: 'Số phiếu', sort: true },
    { k: 'tt', h: 'Tình trạng', sort: true },
    { k: 'flags', h: 'Ghi chú' },
  ];

  function tableData() {
    let base = S.drill && S.drill.scope === 'all' ? S.data.records : S.filtered;
    if (S.drill) base = base.filter(S.drill.pred);
    const terms = L.fold(S.search.trim()).split(/\s+/).filter(Boolean);
    if (terms.length) base = base.filter((r) => terms.every((t) => r.search.includes(t)));
    const { key, dir } = S.sort;
    const arr = base.slice();
    const coll = new Intl.Collator('vi', { numeric: true });
    arr.sort((a, b) => {
      const va = a[key], vb = b[key];
      const na = va == null || (key === 'vol' && !a.volOk), nb = vb == null || (key === 'vol' && !b.volOk);
      if (na !== nb) return na ? 1 : -1; // giá trị trống luôn xếp cuối
      let c = 0;
      if (!na) c = typeof va === 'number' ? va - vb : coll.compare(va, vb);
      if (c === 0) c = (a.day || 0) - (b.day || 0) || a.row - b.row;
      return c * dir;
    });
    return arr;
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
      body = `<tr class="empty-row"><td colspan="${COLS.length}">Không có giao dịch phù hợp với điều kiện lọc${S.search ? ` và từ khóa "${esc(S.search)}"` : ''}.</td></tr>`;
    } else {
      body = page.map((r) => {
        const f = r.flags;
        const badges = [];
        if (isDeleted(r.tt)) badges.push('<span class="badge del">Đã xóa</span>');
        if (f.htDefault) badges.push(`<span class="badge" title="Ô gốc trống; ${esc(f.htDefault.desc)} → ${esc(f.htDefault.value)}">HTVC mặc định</span>`);
        if (r.volOk && r.vol === 0) badges.push('<span class="badge">KL = 0</span>');
        if (f.dupVoucher) badges.push(`<span class="badge">Trùng ${esc(soLabel())}</span>`);
        if (f.dupRow) badges.push('<span class="badge">Trùng nội dung</span>');
        if (f.volError) badges.push(`<span class="badge del">${esc(f.volError)}</span>`);
        if (f.dateError) badges.push(`<span class="badge del">${esc(f.dateError)}</span>`);
        if (f.unitMismatch) badges.push('<span class="badge">ĐV khác quy tắc</span>');
        if (f.unitInferred) badges.push('<span class="badge">ĐV suy luận</span>');
        if (f.negative) badges.push('<span class="badge">KL âm</span>');
        const cell = (v) => (v === L.EMPTY ? `<span class="muted">${v}</span>` : esc(v));
        return `<tr><td class="num">${r.dateStr || `<span class="muted">${esc(r.raw.date)}</span>`}</td><td>${cell(r.kho)}</td><td>${cell(r.kh)}</td><td>${cell(r.ht)}</td><td>${cell(r.pt)}</td><td>${cell(r.sp)}</td>
          <td class="r" title="Giá trị gốc: ${esc(r.raw.kl)}">${r.volOk ? fmt(r.vol) : `<span class="muted">${esc(r.raw.kl) || '—'}</span>`}</td><td class="unit">${esc(r.unit || '')}</td>
          <td class="num">${esc(r.so)}</td><td>${cell(r.tt)}</td><td>${badges.join('')}</td></tr>`;
      }).join('');
    }
    $('#detailTable').innerHTML = head + `<tbody>${body}</tbody>`;

    const tot = A.totalsByUnit(rows);
    const parts = S.data.meta.units.map((u) => `<b>${fmt(tot.units[u] ? tot.units[u].sum : 0)}</b> ${u}`);
    $('#tableTotals').innerHTML = `Tổng số bản ghi: <b>${fmtInt(rows.length)}</b> · Tổng khối lượng: ${parts.join(' · ')}${tot.invalid ? ` <span class="muted">(${fmtInt(tot.invalid)} bản ghi thiếu/sai KL không cộng)</span>` : ''}`;
    $('#detailSub').textContent = rows.length ? `Hiển thị ${fmtInt(start + 1)}–${fmtInt(start + page.length)} / ${fmtInt(rows.length)} bản ghi` : 'Không có bản ghi';

    // Phân trang
    const btn = (p, label, dis, on) => `<button data-page="${p}" ${dis ? 'disabled' : ''} class="${on ? 'on' : ''}">${label}</button>`;
    const nums = [];
    const from = Math.max(1, S.page - 2), to = Math.min(pages, S.page + 2);
    for (let p = from; p <= to; p++) nums.push(btn(p, p, false, p === S.page));
    $('#pager').innerHTML = btn(1, '«', S.page === 1) + btn(S.page - 1, '‹', S.page === 1) + nums.join('') + btn(S.page + 1, '›', S.page === pages) + btn(pages, '»', S.page === pages) + `<span class="pinfo">Trang ${fmtInt(S.page)}/${fmtInt(pages)}</span>`;
  }

  // ============================ SỰ KIỆN ============================
  $('#btnOpen').addEventListener('click', openPicker);
  $('#btnOpen2').addEventListener('click', openPicker);
  $('#btnRefresh').addEventListener('click', refresh);
  $('#fileInput').addEventListener('change', (e) => {
    const f = e.target.files[0];
    if (f) loadFile(f);
    e.target.value = '';
  });
  $('#btnReopen').addEventListener('click', refresh);
  $('#btnClearSaved').addEventListener('click', clearSaved);

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

  // Bộ lọc thời gian
  // Ô ngày nhập theo DD/MM/YYYY (không phụ thuộc ngôn ngữ trình duyệt); nút lịch mở bộ chọn ngày gốc
  function setDate(which, day) {
    S.filters[which] = day;
    $('#fQuick').value = '';
    fixRange(which);
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
  function fixRange(changed) {
    const f = S.filters;
    if (f.from != null && f.to != null && f.from > f.to) {
      if (changed === 'from') f.to = f.from; else f.from = f.to;
      toast('"Từ ngày" không được sau "Đến ngày" — đã tự điều chỉnh.');
    }
  }
  $('#fQuick').addEventListener('change', (e) => {
    const v = e.target.value, max = S.data && S.data.meta.maxDay;
    if (!v || max == null) return;
    const d = new Date(max * 86400000), y = d.getUTCFullYear(), m = d.getUTCMonth();
    const day = (yy, mm, dd) => Math.round(Date.UTC(yy, mm, dd) / 86400000);
    const f = S.filters;
    const ranges = {
      all: [null, null], today: [max, max], '7d': [max - 6, max], '30d': [max - 29, max],
      thisMonth: [day(y, m, 1), day(y, m + 1, 0)], prevMonth: [day(y, m - 1, 1), day(y, m, 0)],
      thisQuarter: [day(y, Math.floor(m / 3) * 3, 1), day(y, Math.floor(m / 3) * 3 + 3, 0)], thisYear: [day(y, 0, 1), day(y, 11, 31)],
    };
    [f.from, f.to] = ranges[v];
    scheduleUpdate();
  });

  function resetFilters() {
    const f = S.filters;
    f.from = f.to = null;
    f.nonZero = true; $('#fNonZero').checked = true;
    for (const d of DIMS) f.sets[d].clear();
    setDefaultStatus();
    $('#fQuick').value = '';
    S.drill = null; S.search = ''; $('#search').value = '';
    DIMS.forEach((d) => MS[d].renderButton());
    update();
    toast('Đã xóa bộ lọc (giữ các mặc định: bỏ phiếu "Đã xóa" và dòng khối lượng = 0 nếu có).');
  }
  $('#btnReset').addEventListener('click', resetFilters);
  $('#fNonZero').addEventListener('change', (e) => { S.filters.nonZero = e.target.checked; scheduleUpdate(); });
  $('#applied').addEventListener('click', (e) => {
    const k = e.target.dataset.clear;
    if (!k) return;
    if (k === 'date') { S.filters.from = S.filters.to = null; $('#fQuick').value = ''; }
    else if (k === 'nonZero') { S.filters.nonZero = false; $('#fNonZero').checked = false; }
    else { S.filters.sets[k].clear(); MS[k].renderButton(); }
    scheduleUpdate();
  });

  // Điều khiển biểu đồ
  $('#segUnit').addEventListener('click', (e) => { const v = e.target.dataset.v; if (v && !e.target.disabled) { S.unit = S.unitPref = v; renderUnitSeg(); renderCharts(); renderRank(); } });
  $('#segGran').addEventListener('click', (e) => { const v = e.target.dataset.v; if (v) { S.gran = v; renderUnitSeg(); renderCharts(); } });
  $('#segShare').addEventListener('click', (e) => { const v = e.target.dataset.v; if (v) { S.shareDim = v; renderUnitSeg(); renderCharts(); } });
  $('#segRank').addEventListener('click', (e) => { const v = e.target.dataset.v; if (v) { S.rankDim = v; renderUnitSeg(); renderRank(); } });
  $('#vehTop').addEventListener('change', (e) => { S.vehTop = +e.target.value; renderCharts(); });
  $('#rankTable').addEventListener('click', (e) => {
    const tr = e.target.closest('tr[data-key]');
    if (!tr) return;
    const key = tr.dataset.key, dim = S.rankDim;
    setDrill(`${A.DIM_LABELS[dim]} = ${key} (mọi đơn vị)`, (r) => r[dim] === key);
  });

  // Chất lượng dữ liệu → xem bản ghi liên quan (trên toàn bộ dữ liệu)
  $('#quality').addEventListener('click', (e) => {
    const k = e.target.dataset.issue;
    if (!k) return;
    const ids = new Set(S.issueSets[k]);
    if (k === 'dupVoucher' || k === 'dupRow') S.sort = k === 'dupVoucher' ? { key: 'so', dir: 1 } : { key: 'day', dir: -1 };
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
  function doPrint() { window.print(); }
  document.querySelector('.export').addEventListener('click', (e) => {
    const b = e.target.closest('[data-exp]');
    if (!b || !S.data) return;
    const kind = b.dataset.exp;
    if (kind === 'print') { doPrint(); return; }
    const ctx = { meta: S.data.meta, records: S.filtered, tableRecords: S.tableRecords, units: S.data.meta.units, conditions: conditions(), tableConditions: tableConditions() };
    const n = kind === 'detail' ? ctx.tableRecords.length : ctx.records.length;
    if (!n) { toast('Không có dữ liệu phù hợp với bộ lọc để xuất.', true); return; }
    try {
      loading(true, 'Đang tạo file Excel…');
      setTimeout(() => {
        try { X[kind](ctx); toast(`Đã xuất báo cáo (${fmtInt(n)} bản ghi). Kiểm tra thư mục Tải xuống.`); }
        catch (err) { console.error(err); toast('Lỗi khi xuất file: ' + err.message, true); }
        finally { loading(false); }
      }, 30);
    } catch (err) { loading(false); toast('Lỗi khi xuất file: ' + err.message, true); }
  });
  $('#btnPrint').addEventListener('click', () => { if (S.data) doPrint(); else toast('Chưa có dữ liệu để in.', true); });
  window.addEventListener('beforeprint', () => C.resizeAll());
  window.addEventListener('afterprint', () => C.resizeAll());

  // Hướng dẫn
  $('#btnHelp').addEventListener('click', () => dialog(HELP));
  const HELP = `<h2>Hướng dẫn sử dụng</h2>
    <h4>1. Nạp / cập nhật dữ liệu</h4>
    <ul><li>Nhấn <b>Chọn file Excel</b> (hoặc kéo thả file vào trang) và chọn file EXPORT.xlsx (hoặc file PXK).</li>
    <li>Khi file Excel được cập nhật thêm dòng mới: lưu file, rồi nhấn <b>Làm mới dữ liệu</b>. Trên Chrome/Edge, dashboard đọc lại đúng file đã chọn; trình duyệt khác sẽ yêu cầu chọn lại file.</li>
    <li>Lần sau mở dashboard, dữ liệu gần nhất được nạp tự động từ bộ nhớ trình duyệt (chỉ lưu trên máy này).</li></ul>
    <h4>2. Lọc dữ liệu</h4>
    <ul><li>Chọn <b>Từ ngày – Đến ngày</b> hoặc kỳ <b>Nhanh</b> (tính theo ngày mới nhất có trong dữ liệu).</li>
    <li>Các bộ lọc Kho, Khách hàng, HTVC, Phương tiện, Sản phẩm, Tình trạng cho phép chọn nhiều giá trị, có ô tìm kiếm; con số bên phải là số giao dịch theo các bộ lọc khác.</li>
    <li>Điều kiện đang áp dụng hiển thị dưới thanh lọc; nhấn × để bỏ từng điều kiện, hoặc <b>Xóa bộ lọc</b> để bỏ tất cả.</li>
    <li>Phiếu có tình trạng <b>"Đã xóa"</b> (file PXK) mặc định không được tính; muốn tính, tích chọn lại trong bộ lọc Tình trạng phiếu.</li>
    <li>Dòng có <b>khối lượng = 0</b> (lệnh chưa xuất) mặc định không được tính; bỏ chọn ô <b>Chỉ tính GD có KL &gt; 0</b> để tính.</li></ul>
    <h4>3. Đọc số liệu</h4>
    <ul><li>File EXPORT: sản lượng <b>mọi sản phẩm lấy theo cột AM, đơn vị MT</b>. File PXK: LPG và Propylene theo <b>MT</b>, sản phẩm khác theo <b>M3</b>; không cộng lẫn.</li>
    <li>Khi dữ liệu có nhiều đơn vị (file PXK), nút <b>Đơn vị M3 / MT</b> chọn đơn vị cho tất cả biểu đồ và bảng xếp hạng.</li>
    <li>Nhấn vào một cột, điểm, lát biểu đồ hoặc dòng xếp hạng để xem các giao dịch liên quan ở bảng chi tiết. Nhấn <b>Hiện toàn bộ</b> để quay lại.</li>
    <li>Mục <b>Kiểm tra dữ liệu</b> liệt kê dòng thiếu thông tin, trùng số phiếu, trùng nội dung… kèm nút <b>Xem</b>.</li></ul>
    <h4>4. Trích xuất báo cáo</h4>
    <ul><li>Cuối trang có các nút xuất Excel: chi tiết, theo khách hàng, sản phẩm, thời gian, HTVC & phương tiện, hoặc báo cáo đầy đủ.</li>
    <li>File xuất theo đúng bộ lọc hiện tại, có sheet <b>Điều kiện lọc</b>; tên file gồm loại báo cáo và ngày giờ xuất (ví dụ <i>BaoCao_TongHop_KhachHang_20261020_1030.xlsx</i>).</li>
    <li><b>In / PDF</b>: chọn máy in "Microsoft Print to PDF" hoặc "Lưu dưới dạng PDF", khổ ngang.</li></ul>`;

  // ============================ KHỞI ĐỘNG ============================
  C.setupDefaults();
  (async function boot() {
    ST.del('last'); // bản lưu định dạng cũ (PXK) — không dùng nữa
    if (viaServer && (await loadFromServer(false))) return;
    const [last, handle] = await Promise.all([ST.get(CACHE_KEY), ST.get('handle')]);
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
