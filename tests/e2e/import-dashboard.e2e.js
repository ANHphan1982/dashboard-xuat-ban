/* =====================================================================
 * E2E trang Nhập dầu thô/Nguyên liệu (nhap.html) — Playwright + node:test
 * Dữ liệu giả lập tạo ngay khi chạy (không dùng file dữ liệu thật).
 * Chạy: npm run test:e2e
 * ===================================================================== */
'use strict';
const { test, before, after } = require('node:test');
const assert = require('node:assert/strict');
const http = require('node:http');
const fs = require('node:fs');
const path = require('node:path');
const { chromium } = require('playwright');
const XLSX = require('../../lib/xlsx.full.min.js');

const ROOT = path.resolve(__dirname, '../..');

/** File IMPORT.xlsx giả lập: tiêu đề ở cột D (nguyên liệu), I (số lượng), N (ngày nhập) như báo cáo SAP */
function buildWorkbook() {
  const header = Array(14).fill('');
  Object.assign(header, {
    0: 'Material', 1: 'Plant', 2: 'Storage location', 3: 'Material Description', 4: 'Movement Type Text',
    5: 'Material Document', 6: 'Material Doc.Item', 7: 'Purchase Order', 8: 'Quantity', 9: 'Base Unit of Measure',
    13: 'Posting Date',
  });
  const rows = [header];
  let doc = 4900431000;
  const add = (sp, mat, date, kl) => {
    const r = Array(14).fill(null);
    Object.assign(r, { 0: mat, 1: '1000', 2: '1101', 3: sp, 4: 'GR goods receipt', 5: String(++doc), 6: '1', 8: kl, 9: 'MT', 13: date });
    rows.push(r);
  };
  // MTBE chỉ có ở tháng 03, 07, 09 → các tháng 04/05/06/08 bằng 0 sau khi lọc
  add('MTBE', '130027', '15/03/2026', 2496.2);
  add('MTBE', '130027', '10/07/2026', 2977.4);
  add('MTBE', '130027', '22/09/2026', 2976.424);
  for (let m = 1; m <= 10; m++) {
    const mm = String(m).padStart(2, '0');
    add('Dầu thô Dung Quất', '110001', `05/${mm}/2026`, 80000 + m * 1000);
    add('Dầu thô Murban', '110002', `18/${mm}/2026`, 60000 + m * 500);
    if (m % 2) add('Condensate', '120010', `25/${mm}/2026`, 12000);
  }
  const ws = XLSX.utils.aoa_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
}

const TYPES = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript', '.css': 'text/css', '.xlsx': 'application/octet-stream' };
let server, browser, baseUrl;

before(async () => {
  const xlsx = buildWorkbook();
  server = http.createServer((req, res) => {
    const p = decodeURIComponent(new URL(req.url, 'http://x').pathname);
    if (p === '/IMPORT.xlsx') { res.writeHead(200, { 'Content-Type': TYPES['.xlsx'] }); res.end(xlsx); return; }
    const file = path.join(ROOT, p === '/' ? 'nhap.html' : p);
    if (!file.startsWith(ROOT) || !fs.existsSync(file) || fs.statSync(file).isDirectory()) { res.writeHead(404); res.end(); return; }
    res.writeHead(200, { 'Content-Type': TYPES[path.extname(file)] || 'application/octet-stream' });
    fs.createReadStream(file).pipe(res);
  });
  await new Promise((r) => server.listen(0, '127.0.0.1', r));
  baseUrl = `http://127.0.0.1:${server.address().port}`;
  // Ưu tiên Chromium của Playwright; nếu chưa cài thì dùng Edge/Chrome có sẵn trên máy
  for (const opts of [{}, { channel: 'msedge' }, { channel: 'chrome' }]) {
    try { browser = await chromium.launch(opts); break; } catch (e) { if (opts.channel === 'chrome') throw e; }
  }
});
after(async () => {
  if (browser) await browser.close();
  if (server) server.close();
});

async function openDashboard() {
  const page = await browser.newPage({ viewport: { width: 1600, height: 1000 } });
  const errors = [];
  page.on('pageerror', (e) => errors.push(e.message));
  await page.goto(`${baseUrl}/nhap.html`);
  await page.locator('#dash').waitFor({ state: 'visible' });
  return { page, errors };
}

/** Chọn 1 giá trị trong bộ lọc nhiều lựa chọn (Nhóm / Dầu thô-Nguyên liệu) */
async function pick(page, dim, value) {
  const host = page.locator(`[data-ms="${dim}"]`);
  await host.locator('.ms-btn').click();
  await host.locator(`.ms-opt input[value="${value}"]`).check();
  await page.locator('body').click({ position: { x: 5, y: 5 } }); // đóng panel
}

/** Rê chuột lên cột thứ `index` của biểu đồ (theo trục x) */
async function hoverBar(page, canvasId, index) {
  await page.locator(`#${canvasId}`).scrollIntoViewIfNeeded();
  const pt = await page.evaluate(([id, i]) => {
    const ch = Chart.getChart(document.getElementById(id));
    const rect = ch.canvas.getBoundingClientRect();
    return { x: rect.left + ch.scales.x.getPixelForValue(i), y: rect.top + (ch.chartArea.top + ch.chartArea.bottom) / 2 };
  }, [canvasId, index]);
  await page.mouse.move(pt.x, pt.y);
  await page.waitForTimeout(80);
}

const monthLabels = (page) => page.evaluate(() => Chart.getChart(document.getElementById('chMonthly')).data.labels);
const rankKeys = (page) => page.$$eval('#rankTable tr[data-key]', (trs) => trs.map((t) => t.dataset.key));

test('lọc 1 nguyên liệu rồi rê chuột lên tháng không có nhập: không lỗi, dashboard vẫn cập nhật khi đổi bộ lọc', async () => {
  const { page, errors } = await openDashboard();
  try {
    await pick(page, 'nhom', 'Nguyên liệu khác');
    await pick(page, 'sp', 'MTBE');
    assert.deepEqual(await rankKeys(page), ['MTBE']);

    const labels = await monthLabels(page);
    assert.deepEqual(labels, ['03/2026', '04/2026', '05/2026', '06/2026', '07/2026', '08/2026', '09/2026']);
    // Rê qua tất cả các tháng, gồm các tháng = 0 (04, 05, 06, 08)
    for (let i = 0; i < labels.length; i++) await hoverBar(page, 'chMonthly', i);
    assert.deepEqual(errors, [], 'Không được phát sinh lỗi JS khi rê chuột');

    // Tooltip ở tháng 0 phải ẩn, ở tháng có nhập phải có nội dung
    await hoverBar(page, 'chMonthly', labels.indexOf('08/2026'));
    const tipZero = await page.evaluate(() => Chart.getChart(document.getElementById('chMonthly')).tooltip.dataPoints.length);
    assert.equal(tipZero, 0);
    await hoverBar(page, 'chMonthly', labels.indexOf('07/2026'));
    const tipBody = await page.evaluate(() => Chart.getChart(document.getElementById('chMonthly')).tooltip.body.map((b) => b.lines.join('')));
    assert.equal(tipBody.length, 1);
    assert.match(tipBody[0], /MTBE: 2\.977,4 MT/);

    // Chuột vẫn đang ở trên biểu đồ → đổi bộ lọc phải cập nhật toàn bộ trang (trước đây bị "đơ")
    await hoverBar(page, 'chMonthly', labels.indexOf('08/2026'));
    await page.locator('#applied button[data-clear="sp"]').click();
    await page.waitForFunction(() => document.querySelectorAll('#rankTable tr[data-key]').length === 2);
    assert.deepEqual((await rankKeys(page)).sort(), ['Condensate', 'MTBE']);
    await page.locator('#applied button[data-clear="nhom"]').click();
    await page.waitForFunction(() => document.querySelectorAll('#rankTable tr[data-key]').length === 4);
    assert.match(await page.locator('#applied').innerText(), /\/ 28 lần nhập/);
    assert.deepEqual(errors, [], 'Không được phát sinh lỗi JS khi đổi bộ lọc');
  } finally { await page.close(); }
});

test('nhấn cột tháng sau khi lọc: bảng chi tiết hiện đúng lần nhập, bỏ chọn quay lại toàn bộ', async () => {
  const { page, errors } = await openDashboard();
  try {
    await pick(page, 'sp', 'MTBE');
    const labels = await monthLabels(page);
    await hoverBar(page, 'chMonthly', labels.indexOf('09/2026'));
    await page.mouse.down(); await page.mouse.up();
    await page.locator('#drillBar').waitFor({ state: 'visible' });
    assert.match(await page.locator('#drillBar').innerText(), /Tháng 09\/2026 · MTBE/);
    assert.equal(await page.locator('#detailTable tbody tr').count(), 1);
    assert.match(await page.locator('#detailTable tbody tr').innerText(), /22\/09\/2026/);

    await page.locator('#btnDrillOff').click();
    await page.locator('#drillBar').waitFor({ state: 'hidden' });
    assert.equal(await page.locator('#detailTable tbody tr').count(), 3);

    // Rê lên tháng 0 rồi nhấn: không có lần nhập → không mở chế độ xem chi tiết, không lỗi
    await hoverBar(page, 'chMonthly', labels.indexOf('05/2026'));
    await page.mouse.down(); await page.mouse.up();
    assert.equal(await page.locator('#drillBar').isHidden(), true);
    assert.deepEqual(errors, []);
  } finally { await page.close(); }
});

test('các bộ lọc khác vẫn hoạt động bình thường sau khi rê chuột trên mọi biểu đồ', async () => {
  const { page, errors } = await openDashboard();
  try {
    await pick(page, 'sp', 'Condensate');
    for (const id of ['chMonthly', 'chTrend']) {
      const n = await page.evaluate((cid) => Chart.getChart(document.getElementById(cid)).data.labels.length, id);
      for (let i = 0; i < n; i++) await hoverBar(page, id, i);
    }
    await page.locator('#fQuick').selectOption('thisMonth'); // tháng 10/2026: Condensate không có nhập
    await page.waitForFunction(() => /Không có dữ liệu/.test(document.querySelector('#rankTable').textContent));
    await page.locator('#btnReset').click();
    await page.waitForFunction(() => document.querySelectorAll('#rankTable tr[data-key]').length === 4);
    assert.deepEqual(errors, []);
  } finally { await page.close(); }
});
