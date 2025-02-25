# general_monitor.py
import win32gui
import win32process
import psutil
import os
from .base_monitor import BaseWindowMonitor
from .window_info import WindowInfo

class GeneralWindowMonitor(BaseWindowMonitor):
    def is_target_window(self, window: int) -> bool:
        try:
            class_name = win32gui.GetClassName(window)
            if not class_name:
                return False
            return class_name not in ["CabinetWClass", "ExploreWClass"]
        except Exception as e:
            print(f"Error in General is_target_window: {e}")
            return False

    def get_active_window_info(self):
        try:
            window = win32gui.GetForegroundWindow()
            window_title = win32gui.GetWindowText(window)
            
            if window_title == self.last_title:
                return None

            pid = win32process.GetWindowThreadProcessId(window)[1]
            process = psutil.Process(pid)
            
            process_name = process.name()
            application_path = process.exe()
            
            self.last_title = window_title

            return WindowInfo.create(
                process_name=process_name,
                window_title=window_title,
                process_id=pid,
                application_name=os.path.basename(application_path),
                application_path=application_path,
                working_directory='',  # 一般アプリは作業ディレクトリなし
                monitor_type='general'
            )

        except Exception as e:
            print(f"Error in General get_active_window_info: {e}")
            return None