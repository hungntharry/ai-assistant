# Rule: Phân bổ chi phí nước uống - 01

## Thông tin chung
- **Tên rule:** Rule-phân bổ chi phí nước uống-01
- **Loại:** Phân bổ chi phí nước uống
- **Trạng thái:** Đang thiết lập

## File template
- **File gốc:** `File gốc phân bổ chi phí nước uống.xlsx` (tháng 7 = mốc tham chiếu)
- **Thư mục dữ liệu:** `D:\File storage\AI\AI Agent\Nước uống\`

## Cấu trúc Excel
- **Sheet chính:** Sheet "T7" (Bảng phân bổ chi phí nước uống, 125 dòng)
- **Dòng 3:** Tiêu đề `BẢNG PHÂN BỔ CHI PHÍ NƯỚC UỐNG`
- **Dòng 4:** Header
- **Dòng 5+:** Data
- **Cột 2:** Code đơn vị (key match) — C0003773, C0003774...
- **Cột 8:** Số lượng (điền từ input)
- **Các cột khác:** giữ nguyên công thức Excel

## Quy tắc xử lý
1. Đọc dữ liệu đầu vào (PDF/Excel/Image/API): trích xuất Code đơn vị + Số lượng
2. **Tham chiếu duy nhất: Code đơn vị** (input ↔ Excel cột 2)
3. Match → điền **cột 8** (Số lượng) = số lượng từ input
4. **Code trùng tại cột 2:** chia đều số lượng cho mỗi dòng
   - Code xuất hiện N lần → mỗi dòng nhận (số lượng / N)
   - VD: 2 lần → 50% mỗi dòng | 3 lần → 33.33% mỗi dòng
5. **Các cột khác:** giữ nguyên công thức Excel
6. **Tiêu đề + tên sheet:** đổi theo tháng input
   - File gốc = tháng 7 (mốc tham chiếu)
   - Shift = tháng input - 7
   - Tiêu đề: `BẢNG PHÂN BỔ CHI PHÍ NƯỚC UỐNG THÁNG XX/2026`
   - Tên sheet: `T6`, `T7`, `T9`... theo tháng
7. **File output** đặt tên theo tháng input

## Script
(Chưa tạo - chờ PDF data để test)
