# Hướng dẫn sử dụng Script Đồng bộ Điểm Excel

Script này giúp tự động hóa việc lấy điểm từ 3 file kết quả bài thi (Khối 10, 11, 12) và điền vào file Bảng điểm tổng hợp mà vẫn giữ nguyên định dạng (màu sắc, kẻ bảng).

## 📋 Yêu cầu hệ thống
- Đã cài đặt Python 3.x
- Cài đặt các thư viện cần thiết bằng lệnh:
  ```bash
  pip install -r requirements.txt
  ```

## 🚀 Cách sử dụng
1. Chuẩn bị các file Excel:
   - Đặt 3 file nguồn vào cùng thư mục với script (Ví dụ: `K10_KetQua.xlsx`, `K11_KetQua.xlsx`, `K12_KetQua.xlsx`).
   - Đặt file đích vào cùng thư mục (Ví dụ: `BangDiemTongHop.xlsx`).
2. Mở file `sync_grades.py` bằng trình soạn thảo văn bản (Notepad, VS Code...) và kiểm tra tên file ở phần cuối:
   ```python
   source_files = ["K10_KetQua.xlsx", "K11_KetQua.xlsx", "K12_KetQua.xlsx"]
   target_file = "BangDiemTongHop.xlsx"
   ```
3. Chạy script bằng lệnh:
   ```bash
   python sync_grades.py
   ```
4. Kiểm tra file kết quả: `BangDiem_KetQua_DongBo.xlsx`.

## 🛠 Logic xử lý
- **File nguồn:** Tìm sheet có tên "Bài thi môn Tin học", đọc cột C (Tên), D (Lớp), J (Điểm).
- **File đích:** Duyệt từng sheet, lấy tên lớp từ tên sheet, so khớp Tên + Lớp để điền điểm vào cột F.
- **Định dạng:** Sử dụng thư viện `openpyxl` để đảm bảo các ô không bị mất màu sắc hay kẻ bảng của giáo viên.

## ⚠️ Lưu ý
- Nếu cột điểm trong file đích thực tế là cột G (thay vì F), hãy đổi giá trị `COL_SCORE = 6` thành `7` trong file `sync_grades.py`.
- Đảm bảo tên học sinh và lớp không có các ký tự lạ hoặc khoảng trắng thừa (Script đã tự động xử lý `strip()` cơ bản).
