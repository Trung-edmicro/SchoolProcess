# School Process - UI Application

## Giao diện người dùng hiện đại cho ứng dụng School Process

### 🚀 Khởi động nhanh

#### Cách 1: Sử dụng file batch (Windows)
```batch
# Double-click vào file
start_ui.bat
```

#### Cách 2: Sử dụng Python trực tiếp
```bash
python main_ui.py
```

#### Cách 3: Sử dụng module UI
```bash
python app_ui.py
```

### 📋 Tính năng chính

#### 🔄 Workflow tự động
- **Case 1**: Xuất toàn bộ dữ liệu từ OnLuyen API
- **Case 2**: Xuất dữ liệu theo file import với tính năng so sánh

#### 📊 Chức năng đơn lẻ
- Lấy danh sách Giáo viên
- Lấy danh sách Học sinh  
- Chuyển đổi JSON → Excel
- Upload Google Drive

#### ⚙️ Cấu hình
- Cấu hình OnLuyen API
- Cấu hình Google Sheets
- Quản lý đường dẫn thư mục

#### 📈 Theo dõi tiến trình
- Log realtime với màu sắc
- Progress bar chi tiết
- Hiển thị trạng thái từng bước

### 🎨 Giao diện

#### Layout chính
```
┌─────────────────────────────────────────────────────────┐
│ 🏫 School Process Application               ● Sẵn sàng │
│ Ứng dụng xử lý dữ liệu trường học với OnLuyen API      │
├─────────────────┬───────────────────────────────────────┤
│ Chức năng       │ [📋 Log & Tiến trình] [⚙️ Cấu hình]  │
│                 │ [📊 Kết quả]                         │
│ 📊 Case 1       │                                       │
│ 🔍 Case 2       │ ┌─ Tiến trình ─────────────────────┐ │
│                 │ │ Đang xử lý...                    │ │
│ 👨‍🏫 Giáo viên     │ │ ████████████░░░░  75%           │ │
│ 👨‍🎓 Học sinh      │ └─────────────────────────────────┘ │
│ 📄 JSON→Excel   │                                       │
│                 │ ┌─ Log Output ─────────────────────┐ │
│ ⚙️ Cấu hình      │ │ [12:34:56] Bắt đầu workflow...   │ │
│ ℹ️ Về ứng dụng   │ │ [12:34:58] ✅ Login thành công    │ │
│                 │ │ [12:35:01] 📊 Đang lấy dữ liệu... │ │
└─────────────────┼───────────────────────────────────────┤
│ ● Sẵn sàng                              v1.0.0 │
└─────────────────────────────────────────────────────────┘
```

#### Tab hệ thống
1. **📋 Log & Tiến trình**: Theo dõi quá trình xử lý
2. **⚙️ Cấu hình**: Thiết lập kết nối và đường dẫn
3. **📊 Kết quả**: Xem files đã tạo và thống kê

### 🔧 Cấu hình

#### OnLuyen API
```
Base URL: https://onluyen.vnpost.vn/api
Username: [từ Google Sheets]
Password: [từ Google Sheets]
```

#### Google Sheets
```
Spreadsheet ID: [ID từ URL Google Sheets]
Sheet Name: [Tên sheet chứa dữ liệu]
```

#### Đường dẫn
```
Input Directory: data/input
Output Directory: data/output
```

### 📱 Cách sử dụng

#### Workflow Case 1: Toàn bộ dữ liệu
1. Click **📊 Case 1: Toàn bộ dữ liệu**
2. Hệ thống sẽ tự động:
   - Trích xuất dữ liệu từ Google Sheets
   - Login OnLuyen API
   - Lấy danh sách Giáo viên
   - Lấy danh sách Học sinh
   - Lưu JSON
   - Chuyển đổi Excel
   - Upload Google Drive

#### Workflow Case 2: Dữ liệu theo file import  
1. Click **🔍 Case 2: Dữ liệu theo file import**
2. Hệ thống sẽ thực hiện như Case 1 + thêm:
   - Tải file import từ Google Drive
   - So sánh và lọc dữ liệu
   - Chỉ xuất dữ liệu có trong file import

#### Chức năng đơn lẻ
- **👨‍🏫 Lấy danh sách Giáo viên**: Chỉ lấy dữ liệu giáo viên
- **👨‍🎓 Lấy danh sách Học sinh**: Chỉ lấy dữ liệu học sinh
- **📄 Chuyển đổi JSON → Excel**: Convert file JSON có sẵn

### 🎯 Log màu sắc

- 🔵 **Xanh**: Thông tin bình thường
- 🟢 **Xanh lá**: Thành công
- 🟡 **Vàng**: Cảnh báo
- 🔴 **Đỏ**: Lỗi
- 🟣 **Tím**: Header/Tiêu đề

### 📊 Kết quả

#### Files được tạo
- **JSON**: `data_[Tên trường]_[timestamp].json`
- **Excel**: `Export_[Tên trường].xlsx` 
- **HT/HP Info**: `ht_hp_info_[Tên trường]_[timestamp].json`
- **Login Info**: `onluyen_login_[timestamp].json`

#### Thao tác với files
- **Double-click**: Mở file
- **Right-click**: Context menu
  - Mở file
  - Mở thư mục chứa
  - Sao chép đường dẫn
  - Xóa file

### 🛠️ Khắc phục sự cố

#### Lỗi Dependencies
```
Thiếu module: pandas, requests, google-auth...
→ Chạy: pip install -r requirements.txt
```

#### Lỗi kết nối OnLuyen API
```
401 Unauthorized
→ Kiểm tra username/password trong Google Sheets
```

#### Lỗi Google Sheets
```
403 Forbidden  
→ Kiểm tra quyền truy cập Google Sheets
→ Kiểm tra service account credentials
```

#### Lỗi file không tìm thấy
```
FileNotFoundError
→ Kiểm tra đường dẫn trong cấu hình
→ Tạo thư mục data/input, data/output
```

### 🔑 Keyboard Shortcuts

- **Ctrl+Q**: Thoát ứng dụng
- **F5**: Refresh UI
- **Ctrl+L**: Focus vào log
- **Ctrl+C**: Copy log content

### 📞 Hỗ trợ

Nếu gặp vấn đề, vui lòng:
1. Kiểm tra log trong tab **📋 Log & Tiến trình**
2. Xem file log chi tiết trong `logs/`
3. Kiểm tra cấu hình trong tab **⚙️ Cấu hình**

---

**Phiên bản**: 1.0.0  
**Ngày cập nhật**: 2025-07-29  
**Phát triển bởi**: Assistant
