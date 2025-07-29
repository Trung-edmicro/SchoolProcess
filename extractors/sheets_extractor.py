import sys
import pandas as pd
from pathlib import Path
from typing import Dict, List, Optional, Any

# Add project root to Python path
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

from config.config_manager import get_config
from processors.google_processor import GoogleDataProcessor


class GoogleSheetsExtractor:
    """Class chuyên biệt để trích xuất dữ liệu từ Google Sheets"""
    
    def __init__(self):
        """Khởi tạo extractor"""
        self.config = get_config()
        self.processor = None
        self._init_processor()
    
    def _init_processor(self) -> None:
        """Khởi tạo Google processor"""
        try:
            paths = self.config.get_paths_config()
            self.processor = GoogleDataProcessor(
                input_folder=paths['input_dir'],
                temp_folder=paths['temp_dir'],
                output_folder=paths['output_dir']
            )
        except Exception as e:
            print(f"❌ Lỗi khởi tạo processor: {e}")
            self.processor = None
    
    def extract_required_columns(self, sheet_id: str = None, sheet_name: str = 'ED-2025', 
                                required_columns: List[str] = None) -> Optional[Dict[str, Any]]:
        """
        Trích xuất các cột cần thiết từ Google Sheets
        
        Args:
            sheet_id (str): Sheet ID hoặc URL
            sheet_name (str): Tên sheet cụ thể
            required_columns (List[str]): Danh sách tên cột cần trích xuất
            
        Returns:
            Optional[Dict[str, Any]]: Dictionary chứa dữ liệu đã trích xuất
        """
        if not self.processor:
            print("❌ Processor không khả dụng")
            return None
        
        # Sử dụng test sheet nếu không có input
        if not sheet_id:
            google_config = self.config.get_google_config()
            sheet_id = google_config['test_sheet_id']
        
        # Mặc định các cột cần thiết
        if not required_columns:
            required_columns = ['Tên trường', 'Admin', 'Mật khẩu', 'Link driver dữ liệu']
        
        print(f"🔍 TRÍCH XUẤT DỮ LIỆU TỪ GOOGLE SHEETS")
        print("=" * 60)
        print(f"📋 Sheet ID: {sheet_id}")
        print(f"📄 Sheet Name: {sheet_name or 'ED-2025'}")
        print(f"📊 Cột cần trích xuất: {', '.join(required_columns)}")
        print()
        
        try:
            # Đọc dữ liệu từ sheet
            print("🔄 Đang đọc dữ liệu...")
            df = self.processor.google_client.read_shared_google_sheet(sheet_id, sheet_name)
            
            if df is None or df.empty:
                print("❌ Không đọc được dữ liệu hoặc sheet trống")
                return None
            
            print(f"✅ Đã đọc {len(df)} hàng, {len(df.columns)} cột")
            
            # Xử lý header (nếu columns là số thì cần mapping)
            if all(isinstance(col, (int, float)) for col in df.columns):
                # Nếu columns là số, lấy hàng đầu tiên làm header
                if len(df) > 0:
                    header_row = df.iloc[0].tolist()
                    df.columns = header_row
                    df = df.iloc[1:].reset_index(drop=True)
                    print(f"📋 Đã chuyển đổi header từ hàng đầu tiên")
            
            print(f"📋 Available columns: {list(df.columns)}")
            
            # Tìm các cột cần thiết
            found_columns = {}
            missing_columns = []
            
            for required_col in required_columns:
                # Tìm exact match trước
                if required_col in df.columns:
                    found_columns[required_col] = required_col
                else:
                    # Tìm partial match (case insensitive)
                    matched = False
                    for col in df.columns:
                        if str(col).lower().strip() in required_col.lower() or \
                           required_col.lower() in str(col).lower().strip():
                            found_columns[required_col] = col
                            matched = True
                            break
                    
                    if not matched:
                        missing_columns.append(required_col)
            
            print(f"\n🔍 KẾT QUẢ TÌM KIẾM CỘT:")
            for req_col, found_col in found_columns.items():
                print(f"   ✅ '{req_col}' → '{found_col}'")
            
            for missing_col in missing_columns:
                print(f"   ❌ '{missing_col}' → Không tìm thấy")
            
            if not found_columns:
                print("\n❌ Không tìm thấy cột nào cần thiết")
                return None
            
            # Trích xuất dữ liệu
            extracted_data = {
                'metadata': {
                    'sheet_id': sheet_id,
                    'sheet_name': sheet_name,
                    'total_rows': len(df),
                    'total_columns': len(df.columns),
                    'found_columns': found_columns,
                    'missing_columns': missing_columns
                },
                'data': []
            }
            
            # Lấy dữ liệu từ các cột tìm được
            for index, row in df.iterrows():
                row_data = {}
                for req_col, found_col in found_columns.items():
                    value = row[found_col]
                    # Xử lý giá trị null/NaN
                    if pd.isna(value) or value == '' or str(value).lower() == 'nan':
                        value = None
                    else:
                        value = str(value).strip()
                    
                    # Lấy trực tiếp text value cho tất cả các cột, bao gồm "Link driver dữ liệu"
                    if req_col == 'Link driver dữ liệu':
                        print(f"   🔍 Processing Link driver dữ liệu: '{value}'")
                        # Sử dụng trực tiếp text value mà không extract hyperlink
                        if value and value != 'None':
                            print(f"   � Using text value for row {index}: {value}")
                        else:
                            print(f"   ⚠️ No text value found for row {index}")
                    
                    row_data[req_col] = value
                
                # Chỉ thêm hàng có ít nhất 1 giá trị không null
                if any(v is not None for v in row_data.values()):
                    row_data['row_index'] = index
                    extracted_data['data'].append(row_data)
            
            print(f"\n📊 ĐÃ TRÍCH XUẤT:")
            print(f"   📋 Số hàng có dữ liệu: {len(extracted_data['data'])}")
            print(f"   📊 Số cột tìm được: {len(found_columns)}")
            
            # Hiển thị preview dữ liệu
            print(f"\n📋 PREVIEW DỮ LIỆU:")
            for i, row_data in enumerate(extracted_data['data'][:3]):  # Chỉ hiển thị 3 hàng đầu
                print(f"   Hàng {i+1}:")
                for req_col in required_columns:
                    if req_col in row_data:
                        value = row_data[req_col]
                        display_value = value[:50] + "..." if value and len(value) > 50 else value
                        print(f"      {req_col}: {display_value}")
                print()
            
            if len(extracted_data['data']) > 3:
                print(f"   ... và {len(extracted_data['data']) - 3} hàng khác")
            
            return extracted_data
            
        except Exception as e:
            print(f"❌ Lỗi trích xuất dữ liệu: {e}")
            return None
    
    def extract_school_data(self, sheet_id: str = None, sheet_name: str = "ED-2025") -> Optional[List[Dict[str, str]]]:
        """
        Trích xuất dữ liệu trường học cần thiết
        
        Args:
            sheet_id (str): Sheet ID hoặc URL
            sheet_name (str): Tên sheet
            
        Returns:
            Optional[List[Dict[str, str]]]: Danh sách dữ liệu trường học
        """
        required_columns = ['Tên trường', 'Admin', 'Mật khẩu', 'Link driver dữ liệu']
        
        extracted = self.extract_required_columns(sheet_id, sheet_name, required_columns)
        
        if extracted and extracted['data']:
            # Trả về chỉ dữ liệu, bỏ qua metadata
            return extracted['data']
        
        return None
    
    def get_school_info_summary(self, sheet_id: str = None, sheet_name: str = "ED-2025") -> Dict[str, Any]:
        """
        Lấy thông tin tóm tắt về dữ liệu trường học
        
        Args:
            sheet_id (str): Sheet ID hoặc URL
            sheet_name (str): Tên sheet
            
        Returns:
            Dict[str, Any]: Thông tin tóm tắt
        """
        school_data = self.extract_school_data(sheet_id, sheet_name)
        
        if not school_data:
            return {'status': 'error', 'message': 'Không lấy được dữ liệu'}
        
        summary = {
            'status': 'success',
            'total_schools': len(school_data),
            'schools_with_complete_info': 0,
            'schools_missing_info': 0,
            'missing_fields_stats': {
                'Tên trường': 0,
                'Admin': 0,
                'Mật khẩu': 0,
                'Link driver dữ liệu': 0
            }
        }
        
        for school in school_data:
            missing_count = 0
            for field in ['Tên trường', 'Admin', 'Mật khẩu', 'Link driver dữ liệu']:
                if not school.get(field):
                    summary['missing_fields_stats'][field] += 1
                    missing_count += 1
            
            if missing_count == 0:
                summary['schools_with_complete_info'] += 1
            else:
                summary['schools_missing_info'] += 1
        
        return summary
    
    def _extract_hyperlinks_from_column(self, sheet_id: str, sheet_name: str, column_name: str) -> dict:
        """
        [DEPRECATED] Trích xuất tất cả hyperlinks từ một cột trong Google Sheets
        NOTE: Method này đã được disable theo yêu cầu không trích xuất hyperlinks
        
        Args:
            sheet_id (str): ID của Google Sheet
            sheet_name (str): Tên sheet
            column_name (str): Tên cột cần extract hyperlinks
            
        Returns:
            dict: Empty dict (không còn extract hyperlinks)
        """
        print(f"      ℹ️ Hyperlink extraction disabled for column '{column_name}'")
        print(f"      📝 Using direct text values instead of hyperlinks")
        return {}
    
    def _extract_hyperlink_url(self, row_index: int, column_name: str, sheet_id: str, sheet_name: str = None) -> str:
        """
        Trích xuất URL thực tế từ hyperlink trong Google Sheets cell
        
        Args:
            row_index (int): Index của hàng (0-based)
            column_name (str): Tên cột 
            sheet_id (str): ID của Google Sheet
            sheet_name (str): Tên sheet
            
        Returns:
            str: URL thực tế hoặc None nếu không có hyperlink
        """
        try:
            print(f"      🔍 Extracting hyperlink for row {row_index}, column '{column_name}'")
            
            # Lấy service từ processor
            sheets_service = self.processor.google_client.sheets_service
            
            # Lấy metadata để tìm sheet ID
            sheet_metadata = sheets_service.spreadsheets().get(
                spreadsheetId=sheet_id
            ).execute()
            
            # Tìm sheet ID từ tên
            target_sheet_id = None
            available_sheets = []
            for sheet in sheet_metadata['sheets']:
                sheet_title = sheet['properties']['title']
                available_sheets.append(sheet_title)
                if sheet_name and sheet_title == sheet_name:
                    target_sheet_id = sheet['properties']['sheetId']
                    break
                elif not sheet_name and sheet['properties']['index'] == 0:
                    target_sheet_id = sheet['properties']['sheetId']
                    break
            
            print(f"      📋 Available sheets: {available_sheets}")
            print(f"      🎯 Target sheet: {sheet_name}, ID: {target_sheet_id}")
            
            if target_sheet_id is None:
                print(f"      ❌ Could not find sheet ID")
                return None
            
            # Tìm column index từ tên cột
            # Đọc header để map column name → index
            range_name = f"{sheet_name}!1:1" if sheet_name else "A1:ZZ1"
            header_result = sheets_service.spreadsheets().values().get(
                spreadsheetId=sheet_id,
                range=range_name
            ).execute()
            
            header_values = header_result.get('values', [[]])[0] if header_result.get('values') else []
            print(f"      📋 Header values: {header_values}")
            
            column_index = None
            for i, header in enumerate(header_values):
                if str(header).strip() == column_name:
                    column_index = i
                    break
            
            print(f"      📊 Column '{column_name}' found at index: {column_index}")
            
            if column_index is None:
                print(f"      ❌ Column not found in header")
                return None
            
            # Chuyển đổi column index thành A1 notation
            def col_index_to_letter(index):
                """Convert 0-based column index to letter (A, B, C, ...)"""
                result = ""
                while index >= 0:
                    result = chr(index % 26 + ord('A')) + result
                    index = index // 26 - 1
                return result
            
            column_letter = col_index_to_letter(column_index)
            cell_address = f"{column_letter}{row_index + 2}"  # +2 vì row_index 0-based và có header
            
            print(f"      📍 Cell address: {cell_address}")
            
            # Lấy cell data với hyperlink information
            range_name = f"{sheet_name}!{cell_address}" if sheet_name else cell_address
            
            print(f"      🔍 Getting cell data for range: {range_name}")
            
            result = sheets_service.spreadsheets().get(
                spreadsheetId=sheet_id,
                ranges=[range_name],
                includeGridData=True
            ).execute()
            
            print(f"      📄 Response keys: {result.keys()}")
            
            # Extract hyperlink URL từ response
            sheets = result.get('sheets', [])
            print(f"      📊 Number of sheets in response: {len(sheets)}")
            
            if sheets:
                data = sheets[0].get('data', [])
                print(f"      📊 Number of data sections: {len(data)}")
                
                if data:
                    row_data = data[0].get('rowData', [])
                    print(f"      📊 Number of rows: {len(row_data)}")
                    
                    if row_data:
                        values = row_data[0].get('values', [])
                        print(f"      📊 Number of values: {len(values)}")
                        
                        if values:
                            cell_data = values[0]
                            print(f"      📊 Cell data keys: {cell_data.keys()}")
                            
                            hyperlink = cell_data.get('hyperlink')
                            if hyperlink:
                                print(f"      ✅ Found hyperlink: {hyperlink}")
                                return hyperlink
                            else:
                                print(f"      ❌ No hyperlink in cell data")
                        else:
                            print(f"      ❌ No values in row data")
                    else:
                        print(f"      ❌ No row data")
                else:
                    print(f"      ❌ No data sections")
            else:
                print(f"      ❌ No sheets in response")
            
            return None
            
        except Exception as e:
            print(f"      ❌ Error extracting hyperlink: {e}")
            import traceback
            print(f"      🔍 Traceback: {traceback.format_exc()}")
            return None


def main():
    """Test function"""
    extractor = GoogleSheetsExtractor()
    
    print("🧪 TEST GOOGLE SHEETS EXTRACTOR")
    print("=" * 60)
    
    # Test 1: Trích xuất dữ liệu trường học
    school_data = extractor.extract_school_data()
    
    if school_data:
        print(f"\n✅ Đã trích xuất {len(school_data)} trường học")
    
    # Test 2: Lấy thông tin tóm tắt
    summary = extractor.get_school_info_summary()
    print(f"\n📊 THÔNG TIN TÓM TẮT:")
    print(f"   Status: {summary.get('status')}")
    if summary.get('status') == 'success':
        print(f"   Tổng số trường: {summary.get('total_schools')}")
        print(f"   Trường có đủ thông tin: {summary.get('schools_with_complete_info')}")
        print(f"   Trường thiếu thông tin: {summary.get('schools_missing_info')}")


if __name__ == "__main__":
    main()
