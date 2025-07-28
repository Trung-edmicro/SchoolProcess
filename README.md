# School Process Application

Ứng dụng xử lý dữ liệu trường học tối ưu - Chỉ tập trung vào 2 chức năng chính.

## 🏗️ Cấu trúc dự án

```
SchoolProcess/
├── app.py                    # Ứng dụng chính (đã được tối ưu)
├── requirements.txt          # Dependencies
├── .env                     # Cấu hình môi trường
├── README.md               # Tài liệu này
│
├── config/                 # Cấu hình
│   ├── __init__.py
│   ├── config.py          # Cấu hình cơ bản
│   ├── config_manager.py  # Quản lý cấu hình
│   ├── google_api.py      # Google APIs
│   ├── onluyen_api.py     # OnLuyen API client
│   └── service_account.json # Google service account
│
├── extractors/            # Trích xuất dữ liệu
│   ├── __init__.py
│   └── sheets_extractor.py # Google Sheets extractor
│
├── converters/            # Chuyển đổi dữ liệu
│   ├── __init__.py
│   └── json_to_excel_converter.py # JSON → Excel
│
├── processors/            # Xử lý dữ liệu
│   ├── __init__.py
│   ├── base_processor.py  # Base class
│   ├── local_processor.py # Xử lý file local
│   ├── google_processor.py # Xử lý Google Sheets
│   └── config_checker.py  # Kiểm tra cấu hình
│
├── utils/                 # Tiện ích
│   ├── __init__.py
│   ├── menu_utils.py      # Menu utilities
│   └── file_utils.py      # File utilities
│
├── data/                  # Dữ liệu
│   ├── input/            # File đầu vào
│   ├── output/           # File đầu ra
│   ├── temp/             # Template files
│   └── sync/             # Đồng bộ dữ liệu
│
├── logs/                  # Log files
└── backups/              # Sao lưu
```

## 🚀 Tính năng chính (Đã tối ưu)

### 1. Xử lý dữ liệu local (Excel files)

- Đọc file Excel giáo viên và học sinh
- Tạo file export theo template chuẩn
- Validation dữ liệu đầu vào

### 2. OnLuyen API Integration - Workflow hoàn chỉnh

- Trích xuất dữ liệu từ Google Sheets
- Login vào OnLuyen API với tài khoản từ Sheets
- Lấy danh sách giáo viên và học sinh
- Lưu dữ liệu JSON và chuyển đổi sang Excel
- Tự động hóa toàn bộ quy trình từ A đến Z

## 🔧 Cài đặt

1. **Clone repository:**

   ```bash
   git clone <repository-url>
   cd SchoolProcess
   ```

2. **Tạo virtual environment:**

   ```bash
   python -m venv .venv
   .venv\\Scripts\\activate  # Windows
   ```

3. **Cài đặt dependencies:**

   ```bash
   pip install -r requirements.txt
   ```

4. **Cấu hình môi trường:**
   - Thêm file `.env`
   - Cập nhật các thông tin cấu hình cần thiết
   - Đặt file `service_account.json` vào thư mục `config/`

## 🎯 Sử dụng

### Chạy ứng dụng:

```bash
python app.py
```

### Menu chính (Đã tối ưu - chỉ 2 lựa chọn):

1. **Xử lý dữ liệu local (Excel files)** - Xử lý file Excel local
2. **OnLuyen API Integration** - Workflow hoàn chỉnh

### Quy trình OnLuyen API Integration:

1. **Trích xuất dữ liệu từ Google Sheets** - Lấy thông tin trường học và admin
2. **Login vào OnLuyen API** - Sử dụng thông tin admin từ Sheets
3. **Lấy dữ liệu giáo viên và học sinh** - Download toàn bộ data từ API
4. **Lưu workflow JSON** - Backup dữ liệu dạng JSON
5. **Chuyển đổi sang Excel** - Tạo file Excel theo template chuẩn
6. **Tổng hợp báo cáo** - Hiển thị kết quả và thống kê

> **Lưu ý**: Ứng dụng đã được tối ưu để chỉ tập trung vào 2 chức năng chính, loại bỏ các tính năng phức tạp không cần thiết.

## 📄 License

MIT License - xem file LICENSE để biết thêm chi tiết.
