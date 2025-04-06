# window_monitor.py
import win32gui
import win32process
import psutil
from datetime import datetime
import os

class WindowMonitor:
    def __init__(self):
        self.last_window = None

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
                'file_name': os.path.basename(file_path),  # application_name として使用
                'file_path': file_path,                    # application_path として使用
                'explorer_path': ''                        # directory_path として使用（一般アプリは空文字）
            }
        except Exception as e:
            return None