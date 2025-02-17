# general_monitor.py
from .base_monitor import BaseWindowMonitor
import win32gui
import win32process
import psutil
from datetime import datetime
import os

class GeneralWindowMonitor(BaseWindowMonitor):
    def is_target_window(self, window):
        try:
            # エクスプローラーウィンドウでない場合のみTrueを返す
            class_name = win32gui.GetClassName(window)
            if not class_name:  # クラス名が取得できない場合
                return False
            return class_name not in ["CabinetWClass", "ExploreWClass"]
        except Exception as e:
            print(f"Error in GeneralWindowMonitor.is_target_window: {str(e)}")
            return False

    def get_active_window_info(self):
        try:
            window = win32gui.GetForegroundWindow()
            if window == self.last_window:
                return None

            pid = win32process.GetWindowThreadProcessId(window)[1]
            process = psutil.Process(pid)
            
            window_title = win32gui.GetWindowText(window)
            process_name = process.name()
            file_path = process.exe()
            
            self.last_window = window

            return {
                'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'process_name': process_name,
                'window_title': window_title,
                'process_id': pid,
                'file_name': os.path.basename(file_path),
                'file_path': file_path
            }
        except:
            return None