# 📊 TKSHEET INTEGRATION - GOOGLE SHEETS LIKE INTERFACE

## 🎯 Tổng quan

Dự án School Process đã được nâng cấp với module **Tksheet** để tạo ra giao diện bảng tính giống Google Sheets, cho phép edit cells trực tiếp và các tương tác nâng cao.

## 🚀 Cài đặt

### 1. Cài đặt Tksheet
```bash
pip install tksheet>=7.5.0
```

### 2. Hoặc sử dụng script tự động
```bash
# Windows
install_and_test_tksheet.bat

# Manual
pip install -r requirements.txt
```

## 🧪 Test Demo

Chạy demo để test Tksheet trước khi sử dụng ứng dụng chính:

```bash
python test_tksheet_demo.py
```

Demo này sẽ hiển thị:
- ✅ Bảng tính giống Google Sheets
- ✏️ Edit cells trực tiếp
- 🎨 Styling Material Design
- 🔍 Features như copy/paste, right-click menu

## 📋 Tính năng mới trong Sheets Viewer

### 🎨 **Giao diện Google Sheets**
- **Colors**: Sử dụng Google Material Design colors
- **Fonts**: Segoe UI với typography chuẩn
- **Grid**: Đường kẻ bảng như Google Sheets
- **Selection**: Highlight cells giống Google Sheets

### ✏️ **Edit trực tiếp**
- **Double-click**: Edit cell trực tiếp
- **Enter/Tab**: Navigate giữa các cells
- **Auto-save**: Tự động lưu thay đổi

### 🔧 **Tính năng nâng cao**
- **Right-click menu**: Insert/delete rows, copy/paste
- **Column resize**: Kéo thả để điều chỉnh độ rộng
- **Row selection**: Select cả row hoặc column
- **Search/Filter**: Tìm kiếm và lọc dữ liệu realtime

### 📤 **Export/Import**
- **CSV Export**: Xuất dữ liệu ra CSV
- **Copy/Paste**: Tương thích với Excel, Google Sheets
- **Bulk operations**: Thêm/xóa nhiều rows cùng lúc

## 🎛️ Cách sử dụng

### 1. **Mở Sheets Viewer**
```bash
python main_ui.py
```
→ Chọn tab "📊 Google Sheets"

### 2. **Edit dữ liệu**
- **Double-click** vào cell để edit
- **Enter** để confirm
- **Escape** để cancel

### 3. **Thêm rows**
- Click **"➕ Add Row"**
- Hoặc right-click → "Insert Row"

### 4. **Search dữ liệu**
- Gõ vào search box
- Filter realtime theo tất cả columns

### 5. **Export dữ liệu**
- Click **"📤 Export"**
- Chọn location và save CSV

## 🔧 Configuration

### Styling customization
File: `ui/sheets_viewer.py` - function `configure_sheet_styling()`

```python
self.sheet_widget.set_options(
    font=('Segoe UI', 11),
    table_bg='white',
    table_selected_cells_bg='#e8f0fe',
    header_bg='#f8f9fa',
    # ... more options
)
```

### Column configuration
```python
headers = [
    "STT", "ID Trường", "Admin", "Mật khẩu",
    "Link Driver Dữ liệu", "SL GV nạp", "SL HS nạp", "Notes"
]
column_widths = [60, 130, 150, 120, 250, 100, 100, 180]
```

## 🎯 Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| **Double-click** | Edit cell |
| **Enter** | Confirm edit |
| **Escape** | Cancel edit |
| **Tab** | Next cell |
| **Shift+Tab** | Previous cell |
| **Ctrl+C** | Copy |
| **Ctrl+V** | Paste |
| **Delete** | Delete content |
| **Right-click** | Context menu |

## 🔄 Integration với Google Sheets API

### Data flow:
```
Google Sheets API → GoogleSheetsExtractor → Tksheet Display
                                                    ↓
User Edits → Local Changes → Save to file/API (future)
```

### Auto-sync (future enhancement):
- Real-time sync với Google Sheets
- Conflict resolution
- Multi-user collaboration

## 🐛 Troubleshooting

### ❌ Tksheet không hiển thị
```bash
pip install --upgrade tksheet
```

### ❌ Import error
```bash
pip install --force-reinstall tksheet
```

### ❌ Styling không đúng
- Kiểm tra theme compatibility
- Reset options trong `configure_sheet_styling()`

### ❌ Performance issues
- Limit số rows hiển thị
- Implement pagination cho datasets lớn

## 📈 Performance

### Recommended limits:
- **Rows**: < 1000 rows để performance tốt
- **Columns**: < 20 columns
- **Cell content**: < 1000 characters

### Optimization tips:
- Sử dụng search/filter thay vì scroll
- Export CSV cho datasets lớn
- Batch operations cho bulk changes

## 🔮 Future Enhancements

### Planned features:
- [ ] **Real-time sync** với Google Sheets
- [ ] **Multi-user collaboration**
- [ ] **Data validation** rules
- [ ] **Conditional formatting**
- [ ] **Chart/graph integration**
- [ ] **Advanced filtering** (date, number ranges)
- [ ] **Undo/Redo** functionality
- [ ] **Cell comments** và notes

### API Integration:
- [ ] **Save changes** trực tiếp lên Google Sheets
- [ ] **Conflict resolution** cho multi-user
- [ ] **Version history** tracking
- [ ] **Permission management**

## 📞 Support

Nếu gặp vấn đề với Tksheet integration:

1. **Check requirements**: Đảm bảo tksheet>=7.5.0
2. **Run demo**: Test với `test_tksheet_demo.py`
3. **Check logs**: Xem console output khi chạy
4. **Reset config**: Xóa cache và config files

---

🎉 **Enjoy the new Google Sheets-like experience!** 🎉