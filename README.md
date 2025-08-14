# Ứng dụng Xuất Ảnh Thẻ từ Excel

Ứng dụng này giúp xuất ảnh thẻ nhân viên được nhúng trong file Excel thành các ảnh PNG chất lượng cao, với tên file dựa trên mã nhân viên.

## Tính năng
- 🖼️ Trích xuất ảnh nhúng từ file Excel
- 📁 Lưu ảnh dưới dạng PNG với tên file theo mã nhân viên
- ⚙️ Tùy chỉnh chất lượng ảnh
- 📊 Ghi nhật ký chi tiết và báo cáo tiến trình
- 🖥️ Giao diện đồ họa đơn giản
- 🚀 Hỗ trợ đóng gói thành file thực thi (.exe)

## Yêu cầu hệ thống
- Hệ điều hành Windows (cần cài đặt Microsoft Excel)
- Python 3.6+
- Các gói cần thiết:
  ```
  pywin32
  Pillow
  ```

## Cài đặt
1. Sao chép kho lưu trữ này
2. Cài đặt các phụ thuộc:
   ```bash
   pip install pywin32 Pillow
   ```

## Hướng dẫn sử dụng
### Chạy từ mã nguồn
```bash
python ui.py
```

### Đóng gói thành file thực thi
1. Cài đặt PyInstaller:
   ```bash
   pip install pyinstaller
   ```
2. Đóng gói ứng dụng:
   ```bash
   pyinstaller ui.spec
   ```
3. File thực thi sẽ nằm trong thư mục `dist`

### Quy trình sử dụng ứng dụng
1. Chọn file Excel chứa ảnh thẻ nhân viên
2. Chọn thư mục đầu ra (mặc định: "ANHTHE")
3. Điều chỉnh cài đặt nếu cần:
   - Hệ số phóng to: Tăng độ phân giải ảnh (mặc định: 3.0)
   - Thời gian chờ: Độ trễ giữa các thao tác (mặc định: 0.5 giây)
4. Nhấn "Bắt Đầu Xuất Ảnh" để bắt đầu xuất ảnh
5. Xem tiến trình trong tab nhật ký

## Cấu trúc thư mục
```
.
├── ui.py             # Ứng dụng giao diện chính
├── van.py            # Logic xuất ảnh cốt lõi
├── ui.spec           # Cấu hình PyInstaller
├── build/            # Các file build của PyInstaller
└── README.md         # Tài liệu này
```

## Chi tiết kỹ thuật
- Sử dụng COM automation của Excel để truy cập ảnh nhúng
- Tạm thời phóng to ảnh để lấy phiên bản độ phân giải cao
- Ánh xạ ảnh vào bản ghi nhân viên dựa trên vị trí ô
- Hoạt động đa luồng để duy trì giao diện phản hồi

## Lưu ý
- File Excel nên có:
  - Cột A: Mã nhân viên
  - Cột B: Họ tên nhân viên
  - Cột C: Ảnh nhúng
- File đầu ra được đặt tên theo định dạng `[Mã nhân viên]_.png`
- Nhật ký chứa thông tin hoạt động chi tiết

## Hỗ trợ
Đối với sự cố hoặc yêu cầu tính năng mới, vui lòng liên hệ nhóm phát triển.
