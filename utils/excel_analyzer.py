#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
File: excel_analyzer.py
Purpose: Phân tích cấu trúc file import Excel để hiểu format dữ liệu
Author: School Process Application
Date: 2025-07-28
Location: utils/excel_analyzer.py
"""

import pandas as pd
import os
import sys
from pathlib import Path

def analyze_excel_structure(file_path):
    """
    Phân tích cấu trúc file Excel import
    
    Args:
        file_path (str): Đường dẫn tới file Excel
        
    Returns:
        dict: Thông tin cấu trúc file
    """
    print(f"📊 PHÂN TÍCH CẤU TRÚC FILE: {Path(file_path).name}")
    print("=" * 60)
    
    try:
        # Kiểm tra file tồn tại
        if not os.path.exists(file_path):
            print(f"❌ File không tồn tại: {file_path}")
            return None
        
        # Đọc tất cả sheets
        excel_file = pd.ExcelFile(file_path)
        sheets_info = {}
        
        print(f"📄 File: {file_path}")
        print(f"📋 Số sheets: {len(excel_file.sheet_names)}")
        print(f"📝 Tên sheets: {excel_file.sheet_names}")
        print()
        
        # Phân tích từng sheet
        for sheet_name in excel_file.sheet_names:
            print(f"🔍 PHÂN TÍCH SHEET: '{sheet_name}'")
            print("-" * 40)
            
            try:
                # Đọc sheet
                df = pd.read_excel(file_path, sheet_name=sheet_name)
                
                # Thông tin cơ bản
                sheet_info = {
                    'name': sheet_name,
                    'rows': len(df),
                    'columns': len(df.columns),
                    'column_names': list(df.columns),
                    'dtypes': df.dtypes.to_dict(),
                    'sample_data': {},
                    'null_counts': df.isnull().sum().to_dict(),
                    'unique_counts': {}
                }
                
                print(f"   📏 Kích thước: {len(df)} rows x {len(df.columns)} columns")
                print(f"   📋 Tên cột:")
                
                # Hiển thị thông tin từng cột
                for i, col in enumerate(df.columns, 1):
                    col_type = str(df[col].dtype)
                    non_null_count = df[col].count()
                    unique_count = df[col].nunique()
                    
                    sheet_info['unique_counts'][col] = unique_count
                    
                    print(f"      {i:2d}. '{col}' ({col_type}) - {non_null_count} non-null, {unique_count} unique")
                    
                    # Lấy mẫu dữ liệu (5 giá trị đầu tiên không null)
                    sample_values = df[col].dropna().head(5).tolist()
                    sheet_info['sample_data'][col] = sample_values
                    
                    if sample_values:
                        print(f"          💡 Mẫu: {sample_values}")
                
                print()
                
                # Tìm cột có thể chứa email/username
                potential_email_columns = []
                for col in df.columns:
                    col_lower = col.lower()
                    if any(keyword in col_lower for keyword in ['email', 'mail', 'username', 'user', 'tài khoản', 'account']):
                        potential_email_columns.append(col)
                        
                        # Kiểm tra pattern email trong dữ liệu
                        sample_data = df[col].dropna().head(10)
                        email_like_count = 0
                        for value in sample_data:
                            if isinstance(value, str) and '@' in value:
                                email_like_count += 1
                        
                        print(f"   📧 Cột email tiềm năng: '{col}' ({email_like_count}/{len(sample_data)} có '@')")
                
                sheet_info['potential_email_columns'] = potential_email_columns
                
                # Hiển thị vài dòng đầu tiên
                print(f"   📄 PREVIEW DỮ LIỆU (5 dòng đầu):")
                print(df.head().to_string(index=True))
                print()
                
                sheets_info[sheet_name] = sheet_info
                
            except Exception as e:
                print(f"   ❌ Lỗi đọc sheet '{sheet_name}': {e}")
                print()
        
        return {
            'file_path': file_path,
            'file_name': Path(file_path).name,
            'total_sheets': len(excel_file.sheet_names),
            'sheet_names': excel_file.sheet_names,
            'sheets_info': sheets_info
        }
        
    except Exception as e:
        print(f"❌ Lỗi phân tích file: {e}")
        return None

def find_import_files():
    """Tìm tất cả file import trong thư mục temp"""
    # Lấy thư mục gốc của project (parent của utils)
    project_root = Path(__file__).parent.parent
    temp_dir = project_root / "data" / "temp"
    import_files = []
    
    if temp_dir.exists():
        for file in temp_dir.iterdir():
            if file.suffix.lower() in ['.xlsx', '.xls']:
                import_files.append(str(file))
    
    return import_files

def main():
    """Hàm main để chạy phân tích"""
    print("🔍 PHÂN TÍCH CẤU TRÚC FILE IMPORT EXCEL")
    print("=" * 60)
    
    # Lấy thư mục gốc của project
    project_root = Path(__file__).parent.parent
    template_file = project_root / "data" / "temp" / "Template_Import.xlsx"
    
    if template_file.exists():
        print(f"📄 Tìm thấy file template: {template_file}")
        result = analyze_excel_structure(str(template_file))
        
        if result:
            print(f"\n✅ Phân tích hoàn tất!")
            
            # Tóm tắt kết quả
            print(f"\n📊 TÓM TẮT:")
            print(f"   📄 File: {result['file_name']}")
            print(f"   📋 Số sheets: {result['total_sheets']}")
            
            for sheet_name, sheet_info in result['sheets_info'].items():
                print(f"   📝 Sheet '{sheet_name}': {sheet_info['rows']} rows x {sheet_info['columns']} cols")
                
                if sheet_info['potential_email_columns']:
                    print(f"      📧 Cột email tiềm năng: {sheet_info['potential_email_columns']}")
                else:
                    print(f"      ⚠️ Không tìm thấy cột email rõ ràng")
                    
    else:
        print(f"❌ Không tìm thấy file template: {template_file}")
        
        # Tìm các file import khác
        import_files = find_import_files()
        
        if import_files:
            print(f"\n🔍 Tìm thấy {len(import_files)} file Excel khác:")
            for i, file_path in enumerate(import_files, 1):
                print(f"   {i}. {Path(file_path).name}")
            
            if len(import_files) == 1:
                print(f"\n📊 Phân tích file duy nhất...")
                analyze_excel_structure(import_files[0])
            else:
                try:
                    choice = input(f"\nChọn file để phân tích (1-{len(import_files)}): ").strip()
                    choice_idx = int(choice) - 1
                    if 0 <= choice_idx < len(import_files):
                        analyze_excel_structure(import_files[choice_idx])
                    else:
                        print("❌ Lựa chọn không hợp lệ")
                except (ValueError, KeyboardInterrupt):
                    print("❌ Lựa chọn không hợp lệ hoặc bị hủy")
        else:
            print("❌ Không tìm thấy file Excel nào trong data/temp/")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n❌ Đã dừng bởi người dùng")
    except Exception as e:
        print(f"\n❌ Lỗi không mong đợi: {e}")
