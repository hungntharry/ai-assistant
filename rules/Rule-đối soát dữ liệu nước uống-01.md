# Rule: Đối soát dữ liệu nước uống - 01

## Thông tin chung
- **Tên rule:** Rule-đối soát dữ liệu nước uống-01
- **Loại:** Đối soát dữ liệu nước uống
- **Trạng thái:** Hoạt động

## Chức năng
So sánh 2 file bất kỳ (PDF-PDF, PDF-Excel, Excel-Excel) để đối soát mã KH + số lượng.

## Cấu trúc
- **Thư mục dữ liệu:** `D:\File storage\AI\AI Agent\Nước uống\`
- **Key match:** Mã KH (C0003773, C0003774...)
- **Giá trị đối soát:** Số lượng (cột 8 / cột thứ 7 trong PDF)

## Quy tắc đối soát
1. Đọc dữ liệu từ 2 file đầu vào (PDF/Excel/Image — không giới hạn)
2. **Tham chiếu: Số PGH (phiếu giao hàng)** — Số PGH và Số phiếu là cùng một định nghĩa
3. Trích xuất Số PGH + Số lượng từ mỗi file
4. So sánh số lượng theo từng Số PGH:
   - **Đúng:** Cả 2 file có Số PGH + số lượng trùng khớp
   - **Sai:** Cả 2 file có Số PGH nhưng số lượng khác
   - **Không có kết quả:** Số PGH chỉ có trong 1 file, hoặc dòng không phải dòng phiếu
5. **File kết quả:** giữ nguyên form file Excel + bổ sung 1 cột cuối ghi kết quả
6. Bỏ qua ô merged

## PDF extraction (nước uống)
- Tìm mã KH (C\d+) trong table cột 0 → lấy số lượng cột 6 (cột thứ 7)
- Nếu mã跨越 trang → lấy từ dòng "Theo mặt hàng/By SKU:"

## App
- Menu: II.2. Đối soát dữ liệu nước uống
- File: `D:\File storage\AI\AI Agent\Điện\app.py`
