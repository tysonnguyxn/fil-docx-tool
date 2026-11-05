========================================================================
    HƯỚNG DẪN SỬ DỤNG CÔNG CỤ TỰ ĐỘNG ĐIỀN DỮ LIỆU WORD TỪ EXCEL
========================================================================

1. YÊU CẦU DUY NHẤT:
   - Máy tính cần cài đặt sẵn Python (từ bản 3.10 trở lên).
   - Link tải: https://www.python.org/downloads/
   * LƯU Ý QUAN TRỌNG: Lúc cài đặt, nhớ tích chọn vào ô:
     "Add python.exe to PATH" ở chân màn hình cài đặt!

2. CÁCH KHỞI ĐỘNG CÔNG CỤ:
   - Chỉ cần NHẤP ĐÚP CHUỘT vào file: "run.bat"
   - Lần đầu tiên chạy: Hệ thống sẽ tự động tạo môi trường ảo và tải thư viện (mất 1-2 phút).
   - Các lần tiếp theo: Mở lên chạy ngay lập tức trong 2 giây.
   - Giao diện sẽ tự động bật lên trên trình duyệt web.

3. THƯ MỤC "File_Mau_Tham_Khao":
   - Chứa sẵn các file mẫu để bạn tham khảo cách đặt key:
     + mau_danh_sach_buc_xa.docx: File Word mẫu chứa bảng định dạng sẵn.
     + excel_nhan_su_buc_xa.xlsx: File Excel danh sách các hàng cho bảng.
     + excel_thong_tin_chung.xlsx: File Excel thông tin chung ngoài văn bản.

4. NGUYÊN TẮC HOẠT ĐỘNG:
   - Cột/Hàng mẫu trong Word format như thế nào (Phông chữ, Cỡ chữ, In đậm, Màu đỏ...),
     thì giá trị thay thế từ Excel sẽ ăn theo y hệt 100%.
   - Trong bảng, các chữ cố định trong ô (như "Số chứng nhận:", "- Ngày cấp:")
     được giữ nguyên vẹn trên tất cả các dòng sinh ra.
   - Cột số thứ tự {{tt}} tự động đánh số tăng dần 1, 2, 3...
