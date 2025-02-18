# general_monitor.py
import win32gui
import win32process
import psutil
from datetime import datetime
import os
from .base_monitor import BaseWindowMonitor

class GeneralWindowMonitor(BaseWindowMonitor):
    def is_target_window(self, window):
        try:
            class_name = win32gui.GetClassName(window)
            if not class_name:
                return False
            return class_name not in ["CabinetWClass", "ExploreWClass"]
        except Exception as e:
            print(f"Error in General is_target_window: {str(e)}")
            return False

    def get_active_window_info(self):
        try:
            window = win32gui.GetForegroundWindow()
            window_title = win32gui.GetWindowText(window)
            
            # タイトルのみで重複チェック
            if window_title == self.last_title:
                return None

            pid = win32process.GetWindowThreadProcessId(window)[1]
            process = psutil.Process(pid)
            
            process_name = process.name()
            file_path = process.exe()
            
            # 現在のタイトルを保存
            self.last_title = window_title

            return {
                'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'process_name': process_name,
                'window_title': window_title,
                'process_id': pid,
                'file_name': os.path.basename(file_path),
                'file_path': file_path,
                'explorer_path': 'NULL'  # 空文字列から'NULL'に変更
            }
        except Exception as e:
            print(f"Error in General get_active_window_info: {str(e)}")
            return None