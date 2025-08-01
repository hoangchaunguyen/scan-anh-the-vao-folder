# Excel Image Extractor

Script để trích xuất ảnh từ file Excel và lưu dưới dạng file PNG, với tên file dựa trên mã nhân viên.

## 📝 Tính năng chính
- Tự động nhận diện vị trí ảnh trong ô Excel
- Ánh xạ ảnh vào đúng hàng dữ liệu
- Lưu ảnh dưới dạng PNG với tên file theo mã nhân viên
- Hỗ trợ file Excel chứa nhiều ảnh

## ⚙️ Yêu cầu hệ thống
- Python 3.6+
- Thư viện cần thiết:
  ```bash
  pip install xlwings pillow
  ```

## 📁 Cấu trúc thư mục
```
.
├── README.md          # Hướng dẫn sử dụng
├── van.py             # Script chính
├── du_lieu.xlsx       # File Excel đầu vào (ví dụ)
└── ANHTHE/            # Thư mục đầu ra chứa ảnh
    ├── 123456_.png
    └── 789012_.png
```

## 🚀 Cách sử dụng
1. Chuẩn bị file Excel:
   - Đặt ảnh trong cột C (cột số 3)
   - Cột A: Mã nhân viên
   - Cột B: Họ tên

2. Chạy script:
   ```bash
   python van.py
   ```

3. Kết quả:
   - Ảnh được lưu trong thư mục `ANHTHE`
   - Tên file: `[mã_nhân_viên]_.png`

## ⚠️ Lưu ý quan trọng
1. Đảm bảo file Excel không mở khi chạy script
2. Script sẽ tạo thư mục `ANHTHE` nếu chưa tồn tại
3. Tên file ảnh chỉ sử dụng mã nhân viên:
   - Ví dụ: `373555_.png`
   - Các ký tự đặc biệt bị loại bỏ tự động

## 🛠 Xử lý lỗi
Các lỗi thường gặp và giải pháp:
1. **Không tìm thấy ảnh trong Excel**  
   Kiểm tra lại cách chèn ảnh vào file Excel

2. **Ảnh không được ánh xạ đúng hàng**  
   Script sử dụng thuật toán 2 bước:
   - Bước 1: Ánh xạ theo trung tâm ô (dung sai 20px)
   - Bước 2: Ánh xạ theo ranh giới ô (dung sai 30px)

3. **Lỗi clipboard khi lưu ảnh**  
   Script tự động thử lại 3 lần nếu không lấy được ảnh từ clipboard

## 📄 Giấy phép
MIT License - Sử dụng tự do cho mục đích cá nhân và thương mại
