# explorer_monitor.py
import win32com.client
import win32gui
import win32process
import pythoncom
import os
from typing import Optional
from .base_monitor import BaseWindowMonitor
from .window_info import WindowInfo

class ExplorerWindowMonitor(BaseWindowMonitor):
    def __init__(self):
        super().__init__()
        self._initialize_com()

    def _initialize_com(self):
        try:
            pythoncom.CoInitialize()
            self.shell = win32com.client.Dispatch("Shell.Application")
        except Exception as e:
            print(f"COM initialization failed: {str(e)}")
            self.shell = None

    def __del__(self):
        try:
            pythoncom.CoUninitialize()
        except:
            pass

    def is_target_window(self, window: int) -> bool:
        try:
            class_name = win32gui.GetClassName(window)
            return class_name in ["CabinetWClass", "ExploreWClass"]
        except Exception as e:
            print(f"Error in Explorer is_target_window: {e}")
            return False

    def get_active_window_info(self):
        try:
            window = win32gui.GetForegroundWindow()
            if not self.is_target_window(window):
                return None

            window_title = win32gui.GetWindowText(window)
            current_directory = self._get_explorer_path(window)
            if current_directory is None:
                return None

            if window_title == self.last_title:
                return None

            pid = win32process.GetWindowThreadProcessId(window)[1]
            explorer_path = os.path.join(os.environ['WINDIR'], 'explorer.exe')

            self.last_title = window_title

            return WindowInfo.create(
                process_name='explorer.exe',
                window_title=window_title,
                process_id=pid,
                application_name='explorer.exe',
                application_path=explorer_path,
                working_directory=current_directory,
                monitor_type='explorer'
            )

        except Exception as e:
            print(f"Error in Explorer get_active_window_info: {e}")
            return None

    def _get_explorer_path(self, hwnd: int) -> Optional[str]:
        try:
            if self.shell is None:
                self._initialize_com()
                if self.shell is None:
                    return None

            windows = self.shell.Windows()
            hwnd_str = str(hwnd)
            
            for window in windows:
                try:
                    if str(window.HWND) == hwnd_str:
                        return window.Document.Folder.Self.Path
                except Exception as e:
                    continue
            
            return None
        except Exception as e:
            print(f"Error in _get_explorer_path: {e}")
            self.shell = None
            return None