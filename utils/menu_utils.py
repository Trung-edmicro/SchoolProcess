"""
Menu Utilities
Các hàm tiện ích cho menu và UI
Author: Assistant
Date: 2025-07-26
"""

from typing import List, Optional, Callable
from datetime import datetime


def print_header(title: str, subtitle: Optional[str] = None):
    """
    In header chương trình
    
    Args:
        title (str): Tiêu đề chính
        subtitle (Optional[str]): Tiêu đề phụ
    """
    print("=" * 80)
    print(f"🏫 {title}")
    print("=" * 80)
    print(f"📅 Thời gian: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    if subtitle:
        print(f"📋 {subtitle}")
    
    print("=" * 80)


def print_menu(title: str, options: List[str], exit_option: str = "Thoát"):
    """
    In menu lựa chọn
    
    Args:
        title (str): Tiêu đề menu
        options (List[str]): Danh sách các lựa chọn
        exit_option (str): Text cho option thoát
    """
    print(f"\n📋 {title}:")
    
    for i, option in enumerate(options, 1):
        print(f"   {i}️⃣  {option}")
    
    print(f"   0️⃣  {exit_option}")
    print("-" * 60)


def get_user_choice(max_choice: int, prompt: str = "👉 Nhập lựa chọn của bạn") -> int:
    """
    Lấy lựa chọn từ user với validation
    
    Args:
        max_choice (int): Số lựa chọn tối đa
        prompt (str): Text prompt
        
    Returns:
        int: Lựa chọn của user (0 để thoát)
    """
    while True:
        try:
            choice = input(f"{prompt} (0-{max_choice}): ").strip()
            
            if not choice:
                print("⚠️  Vui lòng nhập lựa chọn")
                continue
            
            choice_num = int(choice)
            
            if 0 <= choice_num <= max_choice:
                return choice_num
            else:
                print(f"❌ Lựa chọn không hợp lệ. Vui lòng chọn từ 0-{max_choice}")
                
        except ValueError:
            print("❌ Vui lòng nhập số")
        except KeyboardInterrupt:
            print("\n⏹️  Thoát bởi người dùng")
            return 0


def get_user_input(prompt: str, required: bool = False, default: Optional[str] = None) -> Optional[str]:
    """
    Lấy input text từ user
    
    Args:
        prompt (str): Text prompt
        required (bool): Có bắt buộc nhập không
        default (Optional[str]): Giá trị mặc định
        
    Returns:
        Optional[str]: Input từ user hoặc None
    """
    try:
        if default:
            full_prompt = f"{prompt} (mặc định: {default}): "
        else:
            full_prompt = f"{prompt}: "
        
        user_input = input(full_prompt).strip()
        
        # Sử dụng default nếu không có input
        if not user_input and default:
            return default
        
        # Kiểm tra required
        if required and not user_input:
            print("⚠️  Trường này là bắt buộc")
            return get_user_input(prompt, required, default)
        
        return user_input if user_input else None
        
    except KeyboardInterrupt:
        print("\n⏹️  Hủy bởi người dùng")
        return None


def confirm_action(message: str, default_yes: bool = False) -> bool:
    """
    Xác nhận hành động từ user
    
    Args:
        message (str): Message xác nhận
        default_yes (bool): Mặc định có phải Yes không
        
    Returns:
        bool: True nếu user xác nhận
    """
    try:
        default_text = "Y/n" if default_yes else "y/N"
        response = input(f"🤔 {message} ({default_text}): ").lower().strip()
        
        if not response:
            return default_yes
        
        return response in ['y', 'yes', 'có', 'c']
        
    except KeyboardInterrupt:
        print("\n⏹️  Hủy bởi người dùng")
        return False


def run_menu_loop(title: str, options: List[str], handlers: List[Callable], 
                  exit_option: str = "Thoát", auto_continue: bool = True) -> None:
    """
    Chạy menu loop với các handler
    
    Args:
        title (str): Tiêu đề menu
        options (List[str]): Danh sách options
        handlers (List[Callable]): Danh sách handler functions
        exit_option (str): Text cho option thoát
        auto_continue (bool): Tự động hỏi tiếp tục sau mỗi action
    """
    if len(options) != len(handlers):
        raise ValueError("Số lượng options và handlers phải bằng nhau")
    
    while True:
        try:
            print_menu(title, options, exit_option)
            choice = get_user_choice(len(options))
            
            if choice == 0:
                print(f"\n👋 {exit_option}...")
                break
            
            # Gọi handler tương ứng
            handler = handlers[choice - 1]
            if callable(handler):
                try:
                    handler()
                except Exception as e:
                    print(f"❌ Lỗi thực hiện: {e}")
            
            # Hỏi có muốn tiếp tục không
            if auto_continue and choice != 0:
                if not confirm_action("Tiếp tục sử dụng menu?", default_yes=True):
                    print("\n👋 Thoát menu...")
                    break
                    
        except KeyboardInterrupt:
            print("\n\n⏹️  Menu bị dừng bởi người dùng")
            break
        except Exception as e:
            print(f"\n💥 Lỗi menu: {e}")
            if not confirm_action("Tiếp tục sử dụng menu?"):
                break


def print_separator(title: str = "", length: int = 60, char: str = "-"):
    """
    In dòng phân cách
    
    Args:
        title (str): Tiêu đề (optional)
        length (int): Độ dài dòng
        char (str): Ký tự phân cách
    """
    if title:
        padding = (length - len(title) - 2) // 2
        line = char * padding + f" {title} " + char * padding
        # Đảm bảo độ dài chính xác
        if len(line) < length:
            line += char * (length - len(line))
        print(line)
    else:
        print(char * length)


def print_status(message: str, status: str = "info"):
    """
    In message với status
    
    Args:
        message (str): Nội dung message
        status (str): Loại status (success, error, warning, info)
    """
    icons = {
        'success': '✅',
        'error': '❌', 
        'warning': '⚠️',
        'info': 'ℹ️'
    }
    
    icon = icons.get(status, 'ℹ️')
    print(f"{icon} {message}")


def print_progress(current: int, total: int, message: str = ""):
    """
    In progress bar đơn giản
    
    Args:
        current (int): Tiến độ hiện tại
        total (int): Tổng số
        message (str): Message kèm theo
    """
    percentage = (current / total) * 100 if total > 0 else 0
    bar_length = 20
    filled_length = int(bar_length * current // total) if total > 0 else 0
    
    bar = '█' * filled_length + '░' * (bar_length - filled_length)
    
    progress_text = f"🔄 [{bar}] {percentage:.1f}% ({current}/{total})"
    if message:
        progress_text += f" - {message}"
    
    print(f"\r{progress_text}", end='', flush=True)
    
    if current >= total:
        print()  # New line khi hoàn thành


# Alias functions để dễ sử dụng
get_user_confirmation = confirm_action
