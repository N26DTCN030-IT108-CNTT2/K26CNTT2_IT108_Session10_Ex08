# K26CNTT2_IT108_Session10_Ex08
bài tập số 8
# Bài tập 8: Chữa Lỗi Bảng Đăng Ký Báo Cáo Viên Không Lọc Được

## Phần 1. Phân Tích Lỗi Cấu Trúc Bảng Dữ Liệu (Database Violation Analysis)

### 1. Chi tiết các lỗi vi phạm 5 Quy tắc Vàng Database

- **Lỗi 1 (Hàng tiêu đề phụ / Title Header):**
  - **Vị trí:** Hàng 1 (A1:E1) chứa nội dung `DANH SÁCH BÁO CÁO VIÊN — HỘI NGHỊ KHKH 2026`.
  - **Vi phạm:** Quy tắc số 1 (Hàng đầu tiên của bảng phải là tên cột Header, không được chèn tiêu đề trang trí bên trên).
  - **Hậu quả:** Khi bật AutoFilter (`Ctrl + Shift + L`), Excel sẽ tự động nhận diện Hàng 1 làm Header lọc thay vì Hàng 2 (`STT`, `Họ và Tên`,...).

- **Lỗi 2 (Hàng trống ngắt đoạn dữ liệu):**
  - **Vị trí:** Hàng 5 hoàn toàn để trống.
  - **Vi phạm:** Quy tắc số 4 (Không được phép có dòng hoặc cột trống nằm giữa vùng dữ liệu).
  - **Hậu quả:** AutoFilter bị đứt gãy vùng quét (Range). Khi lọc, hệ thống chỉ quét từ Hàng 2 đến Hàng 4 và bỏ qua toàn bộ dữ liệu từ Hàng 6 trở đi.

- **Lỗi 3 (Dòng tính tổng/thống kê nằm xen kẽ trong vùng dữ liệu):**
  - **Vị trí:** Hàng 6 chứa ô `TỔNG SỐ CNTT: 45` ở cột C.
  - **Vi phạm:** Quy tắc số 5 (Các dòng tính tổng, trung bình, thống kê phải đặt ngoài/dưới bảng dữ liệu thô).
  - **Hậu quả:** Dòng thống kê bị coi là một "bản ghi báo cáo viên", làm sai lệch kết quả lọc và tính toán.

- **Lỗi 4 (Khoảng trắng ẩn trong chuỗi ký tự - Trailing Spaces):**
  - **Vị trí:** Ô B4 chứa giá trị `'Trần Thị Bình  '` (dính khoảng trắng thừa ở cuối).
  - **Vi phạm:** Quy tắc số 3 & Chuẩn hóa dữ liệu thô.
  - **Hậu quả:** Làm hỏng chức năng so sánh chính xác của AutoFilter và hàm tìm kiếm.

---

### 2. Cơ chế so sánh chuỗi của AutoFilter và lỗi khoảng trắng (Ô B4)

- **Cơ chế so sánh:** AutoFilter của Excel/Google Sheets so sánh chuỗi theo nguyên tắc **Exact Match (Khớp từng ký tự theo mã ASCII/Unicode)**.
- **Giải thích lỗi:**
  - Ô B4 chứa chuỗi ký tự: `"Trần Thị Bình  "` (16 ký tự, bao gồm khoảng trắng).
  - Khi người dùng chọn lọc điều kiện: `"Trần Thị Bình"` (14 ký tự, không có khoảng trắng).
  - Do mã Unicode của khoảng trắng (`ASCII 32`) khác với ký tự rỗng, hệ thống xác định hai chuỗi này **không trùng khớp** và loại bỏ ô B4 khỏi kết quả lọc.

---

## Phần 2. Sửa Lỗi & Chuẩn Hóa Cấu Trúc Bảng

### 1. Bảng dữ liệu tái thiết kế chuẩn Database

Bảng dữ liệu sau khi xóa bỏ tiêu đề trang trí, loại bỏ hàng trống, chuyển hàng tổng ra ngoài và giữ Header duy nhất ở Dòng 1:

| STT | Họ và Tên | Khoa / Viện | Điểm Xét Duyệt | Trạng Thái |
| :--- | :--- | :--- | :--- | :--- |
| 1 | Nguyễn Văn An | CNTT | 8.5 | Đã duyệt |
| 2 | Trần Thị Bình | CNTT | 7.9 | Chờ duyệt |
| 3 | Lê Minh Châu | Điện tử | 8.2 | Đã duyệt |

*(Dòng thống kê `TỔNG SỐ CNTT: 45` được chuyển ra ngoài bảng dữ liệu thô hoặc đặt ở file/sheet báo cáo riêng)*.

---

### 2. Xử lý khoảng trắng thừa và thao tác lọc đa điều kiện

#### Bước 1: Làm sạch dữ liệu bằng hàm TRIM()
Tạo một cột phụ `Họ và Tên Chuẩn` bên cạnh cột B và áp dụng công thức:

```excel
=TRIM(B3)
