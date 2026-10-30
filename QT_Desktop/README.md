# QT Quyết Toán 2.0 – Tạo mẫu biểu quyết toán (bản desktop)

**Tác giả: 8702 - Nguyễn Quốc Đạt**

Ứng dụng desktop cho Windows, chuyển đổi và nâng cấp từ chương trình `QT.exe` (QT2024, viết bằng AutoIt).
Đọc file CSV xuất từ IPCAS và tạo các mẫu biểu quyết toán dạng Excel (.xlsx) đã định dạng sẵn để in.

- Một file `.exe` duy nhất, **không cần cài đặt, không cần Microsoft Excel**.
- Chạy được trên máy cấu hình thấp: xử lý theo luồng nền, giao diện không bị treo, có thanh tiến trình.
- Có tính năng **tải lên mẫu biểu mới** (tự thiết kế trong Excel) và **xuất ra Excel**.

## Tải về

Vào mục **Releases** của kho mã → bản `qt-desktop-v2.0.0`:

| File | Dùng cho |
|---|---|
| `QT_QuyetToan.exe` | Windows 64-bit (khuyên dùng) |
| `QT_QuyetToan_32bit.exe` | Windows 32-bit / máy cũ |

Chép file `.exe` vào một thư mục (có thể là USB) rồi chạy. Cấu hình (`QT.ini`) và mẫu tuỳ chỉnh (`MauBieu\`)
được lưu ngay cạnh file chương trình.

## Các chức năng

| Chức năng | Bản cũ QT2024 | Bản mới 2.0 |
|---|---|---|
| Mẫu 05 – Kiểm kê hồ sơ, TS thế chấp, cầm cố | ✔ | ✔ |
| Mẫu 06 – Tổng hợp (tạo kèm khi làm mẫu 05) | ✔ (hỏi có/không) | ✔ (tuỳ chọn) |
| Mẫu 13 – Sao kê tiền gửi | ✔ | ✔ |
| Mẫu 14 – KH có số dư tiền gửi từ 10 tỷ | ✔ | ✔ |
| Mẫu 15a, 15b – Sao kê cho vay | ✔ | ✔ |
| Mẫu 16 – KH có dư nợ từ 10 tỷ | ✔ | ✔ |
| Dữ liệu mẫu 02 (từ 15a) – xác nhận dư nợ, có đọc số bằng chữ | ✔ (hỏi có/không) | ✔ (tuỳ chọn) |
| Mẫu 18 – TK 92 bảo lãnh, LC | ✔ | ✔ |
| Mẫu 20a – Nợ xử lý bằng dự phòng | ✔ | ✔ |
| Mẫu 22 – Doanh thu, chi phí chờ phân bổ (ghép file _1, _2) | ✔ | ✔ |
| Mẫu 24a – Phải thu, phải trả | ✔ (tên file `rt24`) | ✔ (nhận cả `rt24` và `rt24a`) |
| DP – DS xác nhận số dư tiền gửi tổ chức (MSIT81), đọc số bằng chữ | ✔ | ✔ |
| HTT – Sao kê dự thu cuối kỳ | ✔ | ✔ |
| Đối chiếu cân đối GL (`*glcb41*`) → sheet CheckGL + Mẫu 30 | ✔ | ✔ (thêm cột Mã CN, tô đỏ dòng lệch) |
| Tự ghép dữ liệu chia nhiều file | ✔ | ✔ |
| Chế độ kiểm tra (50 dòng đầu) | ✔ | ✔ |
| Tô màu dòng cộng | ✔ | ✔ |
| Toàn tỉnh: xử lý cả thư mục, lưu vào `Result`, bỏ qua mẫu đã có | ✔ | ✔ (có bảng chọn từng mẫu, xem trạng thái, nút Dừng) |
| Lưu cấu hình, danh mục chi nhánh trong `QT.ini` | ✔ | ✔ (đọc được QT.ini cũ) |
| Hướng dẫn sử dụng, phím Esc thoát | ✔ | ✔ |
| Không cần cài Excel | ✘ (phải có Excel) | ✔ |
| Ngày quyết toán, tỷ giá, loại QT | cố định trong mã | ✔ sửa trong Cài đặt |
| Mã ký tự TCVN3 / Unicode / UTF-16, dấu phẩy / TAB | một phần | ✔ tự nhận dạng |
| Sửa số cột khi IPCAS đổi cấu trúc file | ✘ | ✔ mục "Cột dữ liệu" |
| Tải lên mẫu biểu mới tự thiết kế | ✘ | ✔ mục "Mẫu biểu tuỳ chỉnh" |
| Xem trước kết quả, nhật ký xử lý | ✘ | ✔ |
| In sẵn họ tên người ký | ✘ | ✔ |
| Công thức cộng trong Excel | `SUM(...)/2`, `/3` | `SUBTOTAL(9, ...)` – đúng khi sửa số liệu |

## Cách dùng nhanh

1. Xuất file từ IPCAS, đặt tên `<mã CN>_<mẫu>.csv` (vd `5501_rt13.csv`, `5501_rt22_1.csv`).
2. Mở chương trình → **Tạo mẫu biểu** → *Chọn file* → kiểm tra mã chi nhánh/mẫu → **TẠO MẪU BIỂU** (F5).
3. Nhiều chi nhánh: **Xử lý hàng loạt** → chọn thư mục → *Quét thư mục* → *Chạy các mẫu đã chọn*.
4. Vào **Cài đặt** để đặt ngày quyết toán, tỷ giá, khổ giấy, người ký, danh mục chi nhánh.

## Mẫu biểu tuỳ chỉnh (tải lên mẫu mới)

1. **Mẫu biểu tuỳ chỉnh** → *Tạo mẫu khởi đầu từ file CSV* → chương trình sinh file Excel có sẵn tiêu đề,
   dòng dữ liệu `{c1}`, `{c2}`… và dòng cộng `{TONG}`.
2. Mở file, sửa hình thức theo mẫu mới, lưu lại.
3. *Tải lên mẫu mới*, đặt mã (vd `rt25`). File dữ liệu `5501_rt25.csv` sẽ tự dùng mẫu này.
   Đặt trùng mã mẫu có sẵn (vd `rt13`) và chọn *thay thế* nếu Hội sở đổi hình thức mẫu cũ.

Bảng mã: `{c5}` cột 5 · `{Tên khách hàng}` theo tên cột · `{STT}` · `{c3|phai:9}` · `{c3|trai:4}` · `{c6|bo:1}` ·
`{c7|ngay}` · `{c10|so}` · `{c10|chu:c8}` (bằng chữ) · `{c4|hoa}` · `{TONG}` · `{TONG:H}` · `{DEM}` ·
`{MA_CN}` `{TEN_CN}` `{NGAY_QT}` `{NGAY_QT_CHU}` `{NGAY_LAP}` `{LOAI_QT}` `{TEN_MAU}` `{NAM}` `{DDMM}`.
Sheet `CAU_HINH`: `SAP_XEP`, `NHOM_THEO`, `CONG_COT`, `NHAN_NHOM`, `COT_NHAN`, `LOC`, `BO_DONG_DAU`.

## Dành cho người phát triển

```
pip install -r requirements.txt
python main.py                 # chạy giao diện
python tests/run_all.py        # sinh dữ liệu giả lập và chạy toàn bộ mẫu
build_exe.bat                  # đóng gói .exe trên Windows
```

Mã nguồn: `qt_app/forms.py` (logic các mẫu), `sheet.py` (ghi Excel), `reader.py` (đọc CSV, TCVN3),
`glcheck.py` (CheckGL, Mẫu 30), `custom.py` (mẫu tuỳ chỉnh), `gui.py` (giao diện).
File `.exe` được GitHub Actions tự đóng gói (`.github/workflows/build-qt-desktop.yml`).
