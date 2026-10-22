/* =====================================================================
 * nav.js — CHUYỂN TRANG GIỮA CÁC DASHBOARD (dropdown ở tiêu đề)
 * ===================================================================== */
(function () {
  'use strict';
  for (const sel of document.querySelectorAll('.page-nav')) {
    const current = sel.value;
    sel.addEventListener('change', () => { location.href = sel.value; });
    // Quay lại bằng nút Back (trang lấy từ bộ nhớ đệm) → hiện đúng trang đang xem
    window.addEventListener('pageshow', () => { sel.value = current; });
  }
})();
