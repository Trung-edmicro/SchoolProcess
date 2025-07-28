"""
Processors package for School Process
Chứa các module xử lý dữ liệu theo từng chế độ
"""

# Import các class chính
from .base_processor import BaseDataProcessor
from .local_processor import LocalDataProcessor
from .google_processor import GoogleDataProcessor
from .config_checker import ConfigChecker

__all__ = [
    'BaseDataProcessor',
    'LocalDataProcessor', 
    'GoogleDataProcessor',
    'ConfigChecker'
]
