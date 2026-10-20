/* =====================================================================
 * storage.js — LƯU FILE GẦN NHẤT TRONG TRÌNH DUYỆT (IndexedDB, chỉ trên máy này)
 * - Lưu bản sao file Excel gần nhất để lần sau mở dashboard có dữ liệu ngay.
 * - Lưu "tay cầm" file (Chrome/Edge) để nút Làm mới đọc lại file mới nhất trên ổ đĩa.
 * Mọi thao tác đều bọc try/catch: nếu trình duyệt chặn lưu trữ, dashboard vẫn chạy bình thường.
 * ===================================================================== */
(function (root) {
  'use strict';
  const PXK = (root.PXK = root.PXK || {});
  const DB = 'pxk-dashboard', STORE = 'kv';

  function open() {
    return new Promise((resolve, reject) => {
      let req;
      try { req = indexedDB.open(DB, 1); } catch (e) { reject(e); return; }
      req.onupgradeneeded = () => req.result.createObjectStore(STORE);
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }
  async function tx(mode, fn) {
    const db = await open();
    return new Promise((resolve, reject) => {
      const t = db.transaction(STORE, mode);
      const r = fn(t.objectStore(STORE));
      t.oncomplete = () => { db.close(); resolve(r && r.result); };
      t.onerror = () => { db.close(); reject(t.error); };
    });
  }
  const get = (k) => tx('readonly', (s) => s.get(k)).catch(() => null);
  const set = (k, v) => tx('readwrite', (s) => s.put(v, k)).catch(() => null);
  const del = (k) => tx('readwrite', (s) => s.delete(k)).catch(() => null);

  PXK.storage = { get, set, del };
})(window);
