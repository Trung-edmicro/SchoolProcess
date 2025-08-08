"""
UI Module for School Process Application
"""

from .main_window import SchoolProcessMainWindow

# Alias để compatibility với app.py
MainWindow = SchoolProcessMainWindow
from .components import (
    StatusIndicator, 
    ProgressCard, 
    LogViewer, 
    FileList, 
    ConfigSection, 
    WorkflowCard
)

__all__ = [
    'SchoolProcessMainWindow',
    'MainWindow',  # Alias
    'StatusIndicator',
    'ProgressCard', 
    'LogViewer',
    'FileList',
    'ConfigSection',
    'WorkflowCard'
]
