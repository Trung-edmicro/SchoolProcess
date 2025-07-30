"""
Verify Tksheet Fix - Kiểm tra lỗi main_frame đã được sửa
"""

import sys
from pathlib import Path

# Thêm project root vào Python path
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

def test_import():
    """Test import modules"""
    print("🧪 KIỂM TRA TKSHEET FIX")
    print("=" * 50)
    
    # Test 1: Import tksheet
    try:
        import tksheet
        print(f"✅ Tksheet version {tksheet.__version__} - OK")
    except ImportError:
        print("❌ Tksheet chưa được cài đặt")
        return False
    
    # Test 2: Import GoogleSheetsViewer
    try:
        from ui.sheets_viewer import GoogleSheetsViewer
        print("✅ GoogleSheetsViewer import - OK")
    except Exception as e:
        print(f"❌ Lỗi import GoogleSheetsViewer: {e}")
        return False
    
    # Test 3: Test khởi tạo với mock frame
    try:
        import tkinter as tk
        root = tk.Tk()
        root.withdraw()  # Ẩn window
        
        test_frame = tk.Frame(root)
        viewer = GoogleSheetsViewer(test_frame)
        
        # Kiểm tra main_frame đã được tạo
        if hasattr(viewer, 'main_frame'):
            print("✅ main_frame đã được tạo - OK")
        else:
            print("❌ main_frame vẫn chưa được tạo")
            return False
            
        root.destroy()
        
    except Exception as e:
        print(f"❌ Lỗi khởi tạo GoogleSheetsViewer: {e}")
        return False
    
    print("\n🎉 TẤT CẢ TESTS PASS!")
    print("✅ Lỗi 'main_frame' đã được sửa hoàn toàn")
    print("✅ Bạn có thể chạy ứng dụng chính:")
    print("   python main_ui.py")
    
    return True

def main():
    """Main function"""
    success = test_import()
    
    if success:
        print("\n" + "=" * 50)
        print("🚀 HƯỚNG DẪN SỬ DỤNG:")
        print("1. Chạy: python main_ui.py")
        print("2. Chọn tab '📊 Google Sheets'")  
        print("3. Enjoy Google Sheets-like interface!")
        print("=" * 50)
    else:
        print("\n❌ Vẫn còn lỗi. Vui lòng kiểm tra lại.")

if __name__ == "__main__":
    main()