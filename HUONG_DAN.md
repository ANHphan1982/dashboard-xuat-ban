# Dashboard xuất bán sản phẩm — Hướng dẫn

Dashboard dạng trang web chạy **hoàn toàn trên máy tính** (không cần cài đặt, không cần Internet, không gửi dữ liệu ra ngoài).

## 1. Cách mở (khuyến nghị)

1. Mở thư mục `E:\Dashboard xuất bán sản phẩm`.
2. Nháy đúp file **`Mo_Dashboard.bat`**.
   - Một cửa sổ thu nhỏ tên **"Dashboard server"** chạy nền (máy chủ cục bộ `127.0.0.1:8765`, chỉ trong máy này, không ra Internet).
   - Trình duyệt mở `http://127.0.0.1:8765/index.html` và **tự đọc file `EXPORT.xlsx` mới nhất** trong thư mục — không cần chọn file.
3. Dùng xong: đóng cửa sổ "Dashboard server" để tắt máy chủ (hoặc để nguyên, lần sau mở nhanh hơn).

> Cần Python (đã có trên máy). Nếu không có Python, `Mo_Dashboard.bat` sẽ báo; khi đó dùng cách dưới.

**Cách mở thay thế (không cần Python):** nháy đúp `index.html` → nhấn **Chọn file Excel** → chọn `EXPORT.xlsx`. Theo bảo mật trình duyệt, cách này không tự đọc được file trên ổ đĩa, nên mỗi khi dữ liệu thay đổi phải chọn lại file.

## 2. Cập nhật dữ liệu bằng file Excel mới

- **Cập nhật EXPORT.xlsx**: ghi đè/lưu file `EXPORT.xlsx` mới vào thư mục dashboard (giữ nguyên tên file), rồi nhấn **Làm mới dữ liệu** (hoặc F5). Dashboard đọc ngay bản mới nhất.
- Nút **Xóa dữ liệu đã lưu**: xóa bản dữ liệu lưu trong trình duyệt (dùng khi mở bằng `index.html` mà trang vẫn hiện dữ liệu cũ).
- **Dùng file khác** (tên khác, hoặc file PXK): nhấn **Chọn file Excel** và chọn file. Bộ lọc được đặt lại.
- **Định dạng file được hỗ trợ** (dashboard tự nhận dạng theo dòng tiêu đề):

**a) File EXPORT.xlsx (báo cáo xuất hàng từ SAP) — nguồn dữ liệu chính**

| Cột | Tiêu đề trong file | Trường trên dashboard |
|---|---|---|
| AC | Shipping point description | Kho |
| X | Loading date | Ngày xuất |
| L | Tên khách hàng | Khách hàng |
| T | Hình thức vận chuyển | Hình thức vận chuyển |
| C | Phương tiện | Phương tiện |
| K | Mô tả sản phẩm | Sản phẩm |
| AL | Quantity (SM3) | Khối lượng thực tế — **các sản phẩm khác** (đơn vị **SM3**) |
| AM | Quantity (MT) | Khối lượng thực tế — **LPG và Propylene** (đơn vị **MT**) |
| B | Số LXH trên SAP | Kiểm tra trùng (không bắt buộc) |
| F | Shipment status Desc. | Bộ lọc "Trạng thái giao hàng" (không bắt buộc) |

  Các dòng tổng cộng ở cuối file (không có ngày, khách hàng, sản phẩm) tự động bỏ qua.
  Dòng có **khối lượng = 0** (lệnh đã lập nhưng chưa xuất, thường là trạng thái *Scheduled*) mặc định không tính — bỏ chọn ô **Chỉ tính GD có KL > 0** để tính.

**b) File PXK.xlsx (phiếu xuất kho) — vẫn dùng được**

| Cột | Trường |
|---|---|
| G | Kho |
| H | Ngày phiếu xuất kho |
| J | Đơn vị mua hàng (Khách hàng) |
| K | Hình thức vận chuyển |
| L | Phương tiện |
| M | Sản phẩm |
| N | Khối lượng thực tế dạng `7,001 (M3)` / `144,725 (MT)` |

  Nếu cột bị dịch vị trí nhưng vẫn đúng tên tiêu đề, dashboard tự nhận theo tên và ghi chú lại. Nếu thiếu cột bắt buộc, dashboard báo lỗi rõ ràng và **giữ nguyên dữ liệu cũ**.
  Muốn đổi ánh xạ cột: sửa phần `PROFILES` ở đầu file `js/data-loader.js`.

## 3. Quy tắc xử lý số liệu

- **Đơn vị**: LPG tính theo **MT**, các sản phẩm khác tính theo **M3**. Đơn vị đọc từ ô khối lượng; nếu ô không ghi đơn vị thì suy theo quy tắc trên (và được cảnh báo). **Không bao giờ cộng M3 với MT.**
- **Phiếu "Đã xóa"**: mặc định **không tính**. Muốn tính, mở bộ lọc *Tình trạng phiếu* và tích chọn "Đã xóa".
- **Thiếu khối lượng**: vẫn đếm là giao dịch nhưng không cộng vào tổng (KPI ghi rõ số bản ghi bị loại).
- **Thiếu ngày**: bị loại khi lọc theo thời gian.
- **Hình thức vận chuyển trống** được gán theo thứ tự:
  1. Khối lượng **< 50 M3** → **Đường bộ**.
  2. Các dòng còn lại có Phương tiện là **"Pipeline" / "Pipe line"** (không phân biệt hoa thường, dấu cách) → **Đường ống**.
  3. Không thuộc quy tắc nào → giữ là `(Trống)`.

  Các dòng được gán có nhãn "HTVC mặc định" trong bảng và ghi lý do ở cột *Cảnh báo* khi xuất Excel. Quy tắc nằm ở `HT_RULES` đầu file `js/data-loader.js`.
- **Ô trống** khác hiển thị là `(Trống)`.
- **Trùng lặp**: chỉ cảnh báo, không tự xóa. Xem tại mục *Kiểm tra dữ liệu* → **Xem**.
- **Khối lượng âm**: giữ nguyên (có thể là điều chỉnh nghiệp vụ) và được đánh dấu.
- Khoảng trắng thừa trong tên chỉ được chuẩn hóa khi phân tích. File xuất chi tiết giữ nguyên giá trị gốc.

## 4. Sử dụng

- **Bộ lọc** (thanh trên cùng, luôn hiển thị khi cuộn trang):
  - *Từ ngày / Đến ngày*: gõ DD/MM/YYYY hoặc bấm biểu tượng lịch. *Nhanh*: 7/30 ngày, tháng, quý, năm gần nhất (tính theo ngày mới nhất có trong dữ liệu).
  - *Kho, Khách hàng, HTVC, Phương tiện, Sản phẩm, Tình trạng*: chọn nhiều giá trị, có ô tìm kiếm. Con số cạnh mỗi lựa chọn là số giao dịch theo các bộ lọc còn lại.
  - Dòng **Đang áp dụng** liệt kê điều kiện lọc; nhấn × để bỏ từng điều kiện, hoặc **Xóa bộ lọc** để bỏ tất cả.
- **KPI**: tổng khối lượng, bình quân/giao dịch tách theo M3 và MT; số giao dịch; số khách hàng, sản phẩm, phương tiện duy nhất.
- **Biểu đồ**: nút **Đơn vị M3 / Đơn vị MT** chọn đơn vị cho tất cả biểu đồ. Xu hướng có thể xem theo Ngày/Tuần/Tháng; cơ cấu xem theo Sản phẩm/HTVC/Kho/Khách hàng.
- **Xem chi tiết từ biểu đồ**: nhấn vào cột, điểm, lát tròn hoặc dòng của bảng xếp hạng → bảng chi tiết hiển thị các giao dịch liên quan. Nhấn **Hiện toàn bộ** để quay lại.
- **Bảng tra cứu**: tìm kiếm không cần gõ dấu (ví dụ `duong bo`, `43C-226`), nhấn tiêu đề cột để sắp xếp tăng/giảm, chọn 25/50/100/200 dòng/trang. Dưới bảng hiển thị tổng số bản ghi và tổng khối lượng theo từng đơn vị.

## 5. Trích xuất báo cáo

Ở cuối trang, phần **Trích xuất báo cáo**:

| Nút | Nội dung file |
|---|---|
| Chi tiết giao dịch | Đúng các dòng đang có trong bảng tra cứu (gồm bộ lọc, tìm kiếm, xem chi tiết, thứ tự sắp xếp) |
| Tổng hợp theo khách hàng | Theo khách hàng; khách hàng × sản phẩm |
| Tổng hợp theo sản phẩm | Theo sản phẩm; sản phẩm × kho |
| Tổng hợp theo ngày/tuần/tháng | 3 sheet |
| Tổng hợp HTVC & phương tiện | Theo HTVC; theo phương tiện; HTVC × phương tiện |
| Báo cáo tổng hợp đầy đủ | Tất cả bảng tổng hợp + chi tiết trong 1 file |
| In / Xuất PDF | Chọn máy in *Microsoft Print to PDF* / *Lưu dưới dạng PDF*, khổ ngang |

- Mọi file đều theo bộ lọc hiện tại và có sheet **Điều kiện lọc** (thời điểm xuất, file nguồn, điều kiện lọc).
- Ngày lưu dạng ngày Excel định dạng `dd/mm/yyyy`. Khối lượng tách cột theo M3 và MT, có dòng TỔNG CỘNG.
- Tên file có loại báo cáo và thời điểm xuất, ví dụ `BaoCao_TongHop_KhachHang_20261020_103015.xlsx`. File lưu vào thư mục **Tải xuống** của trình duyệt.

## 6. Cấu trúc mã nguồn

```
index.html            Bố cục trang
css/style.css         Giao diện (gồm định dạng khi in/PDF)
js/data-loader.js     Đọc file Excel, nhận cột, chuẩn hóa ngày/khối lượng/đơn vị, kiểm tra chất lượng
js/analytics.js       Lọc dữ liệu, tổng hợp theo chiều/đơn vị/thời gian
js/charts.js          Vẽ biểu đồ (Chart.js)
js/export.js          Xuất báo cáo Excel (SheetJS)
js/storage.js         Lưu file gần nhất trong trình duyệt (IndexedDB, chỉ trên máy này)
js/app.js             Giao diện: bộ lọc, KPI, bảng, xem chi tiết, sự kiện, tự đọc EXPORT.xlsx
Mo_Dashboard.bat      Khởi động máy chủ cục bộ và mở dashboard
lib/                  Thư viện xlsx 0.18.5 và Chart.js 4.4.1 (bản cục bộ)
```

Muốn đưa dashboard sang máy khác: chép **cả thư mục**, rồi chạy `Mo_Dashboard.bat` (hoặc mở `index.html`).

## 7. Xử lý sự cố

- **Không thấy dữ liệu mới sau khi sửa Excel**: lưu file Excel trước (tên `EXPORT.xlsx`, trong thư mục dashboard), rồi nhấn *Làm mới dữ liệu*. Nếu đang mở bằng `index.html` thì chọn lại file, hoặc mở bằng `Mo_Dashboard.bat`.
- **Trang vẫn hiện dữ liệu file cũ**: nhấn *Xóa dữ liệu đã lưu*, hoặc mở bằng `Mo_Dashboard.bat`.
- **Báo cổng 8765 đang bận**: đóng cửa sổ "Dashboard server" cũ rồi chạy lại `Mo_Dashboard.bat`.
- **Báo "File không đúng cấu trúc"**: kiểm tra tên tiêu đề các cột ở bảng mục 2.
- **Muốn xóa dữ liệu đã lưu trong trình duyệt**: xóa dữ liệu trang web (Cài đặt trình duyệt → Quyền riêng tư → Xóa dữ liệu duyệt web → Cookie và dữ liệu trang web).
