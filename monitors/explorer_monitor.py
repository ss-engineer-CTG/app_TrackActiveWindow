# explorer_monitor.py
from .base_monitor import BaseWindowMonitor
import win32com.client
import win32gui
import win32process
import pythoncom
from datetime import datetime
import os

class ExplorerWindowMonitor(BaseWindowMonitor):
    def __init__(self):
        super().__init__()
        self._initialize_com()

    def _initialize_com(self):
        try:
            pythoncom.CoInitialize()
            self.shell = win32com.client.Dispatch("Shell.Application")
            print("Explorer monitor COM initialized")
        except Exception as e:
            print(f"COM initialization failed: {str(e)}")
            self.shell = None

    def __del__(self):
        try:
            pythoncom.CoUninitialize()
        except:
            pass

    def is_target_window(self, window):
        try:
            class_name = win32gui.GetClassName(window)
            is_target = class_name in ["CabinetWClass", "ExploreWClass"]
            print(f"Explorer monitor checking window class: {class_name}, is target: {is_target}")
            return is_target
        except Exception as e:
            print(f"Error in Explorer is_target_window: {str(e)}")
            return False

    def get_active_window_info(self):
        try:
            active_window = win32gui.GetForegroundWindow()
            if not self.is_target_window(active_window):
                return None

            window_title = win32gui.GetWindowText(active_window)
            current_path = self._get_explorer_path(active_window)
            if current_path is None:
                return None

            # タイトルのみで重複チェック
            if window_title == self.last_title:
                return None

            pid = win32process.GetWindowThreadProcessId(active_window)[1]

            # 現在のタイトルを保存
            self.last_title = window_title

            return {
                'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'process_name': 'explorer.exe',
                'window_title': window_title,
                'process_id': pid,
                'file_name': 'explorer.exe',
                'file_path': os.path.join(os.environ['WINDIR'], 'explorer.exe'),
                'explorer_path': current_path
            }
        except Exception as e:
            print(f"Error in Explorer get_active_window_info: {str(e)}")
            return None

    def _get_explorer_path(self, hwnd):
        try:
            if self.shell is None:
                self._initialize_com()
                if self.shell is None:
                    return None

            windows = self.shell.Windows()
            hwnd_str = str(hwnd)
            
            for window in windows:
                try:
                    window_hwnd = str(window.HWND)
                    print(f"Comparing HWNDs: {window_hwnd} vs {hwnd_str}")
                    if window_hwnd == hwnd_str:
                        path = window.Document.Folder.Self.Path
                        print(f"Found matching window with path: {path}")
                        return path
                except Exception as e:
                    print(f"Error processing window in shell.Windows(): {str(e)}")
                    continue
            
            print("No matching Explorer window found")
            return None
        except Exception as e:
            print(f"Error in _get_explorer_path: {str(e)}")
            self.shell = None
            return None