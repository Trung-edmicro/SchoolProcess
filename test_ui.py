"""
Test script để kiểm tra UI components
"""

import sys
from pathlib import Path

# Thêm project root vào Python path
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

def test_ui_import():
    """Test import các UI components"""
    try:
        print("Testing UI imports...")
        
        # Test basic tkinter
        import tkinter as tk
        print("✅ tkinter imported successfully")
        
        # Test UI components
        from ui.components import StatusIndicator, ProgressCard, LogViewer
        print("✅ UI components imported successfully")
        
        # Test main window
        from ui.main_window import SchoolProcessMainWindow
        print("✅ Main window imported successfully")
        
        # Test config
        from config.config_manager import get_config
        config = get_config()
        print("✅ Config manager imported successfully")
        
        print("\n🎉 All imports successful! UI should work properly.")
        return True
        
    except ImportError as e:
        print(f"❌ Import error: {e}")
        return False
    except Exception as e:
        print(f"❌ Error: {e}")
        return False

def test_ui_creation():
    """Test tạo UI window"""
    try:
        print("\nTesting UI window creation...")
        
        import tkinter as tk
        
        # Test basic window
        root = tk.Tk()
        root.title("Test Window")
        root.geometry("300x200")
        
        label = tk.Label(root, text="UI Test Successful! 🎉", font=('Arial', 12))
        label.pack(expand=True)
        
        # Close after 2 seconds
        root.after(2000, root.quit)
        
        print("✅ Test window created")
        
        # Start and immediately close
        root.mainloop()
        root.destroy()
        
        print("✅ UI window test completed successfully")
        return True
        
    except Exception as e:
        print(f"❌ UI creation error: {e}")
        return False

def main():
    """Main test function"""
    print("=" * 50)
    print("    School Process UI - Component Test")
    print("=" * 50)
    
    # Test imports
    import_success = test_ui_import()
    
    if import_success:
        # Test UI creation
        ui_success = test_ui_creation()
        
        if ui_success:
            print("\n🎉 ALL TESTS PASSED!")
            print("✅ UI is ready to use")
            print("\n🚀 You can now run:")
            print("   python main_ui.py")
            print("   or")
            print("   start_ui.bat")
        else:
            print("\n❌ UI creation test failed")
    else:
        print("\n❌ Import test failed")
        print("💡 Please install dependencies:")
        print("   pip install -r requirements.txt")

if __name__ == "__main__":
    main()
