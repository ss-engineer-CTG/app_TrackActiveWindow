# office_monitor.py
from .base_monitor import BaseWindowMonitor
import win32com.client
import win32gui
import win32process
import pythoncom
import psutil
from datetime import datetime
import os
from typing import Optional, Dict, Any
import logging
import time

class OfficeWindowMonitor(BaseWindowMonitor):
    def __init__(self):
        super().__init__()
        self._initialize_com()
        self._app_instances: Dict[str, Any] = {}
        self._last_connection_attempt: Dict[str, float] = {}
        self._cache: Dict[int, Dict[str, Any]] = {}
        self._cache_timeout = 5  # キャッシュの有効期限（秒）
        
        # Officeアプリケーションの設定
        self.office_apps = {
            'winword.exe': {'name': 'Word', 'prog_id': 'Word.Application'},
            'excel.exe': {'name': 'Excel', 'prog_id': 'Excel.Application'},
            'powerpnt.exe': {'name': 'PowerPoint', 'prog_id': 'PowerPoint.Application'}
        }

    def _initialize_com(self) -> None:
        try:
            pythoncom.CoInitialize()
        except Exception as e:
            logging.error(f"COM initialization failed: {e}")

    def __del__(self):
        try:
            for app in self._app_instances.values():
                try:
                    app.Quit()
                except:
                    pass
            pythoncom.CoUninitialize()
        except:
            pass

    def is_target_window(self, window: int) -> bool:
        try:
            _, pid = win32process.GetWindowThreadProcessId(window)
            process = psutil.Process(pid)
            return process.name().lower() in self.office_apps
        except Exception as e:
            logging.error(f"Error in is_target_window: {e}")
            return False

    def _get_running_instance(self, prog_id: str) -> Optional[Any]:
        current_time = time.time()
        last_attempt = self._last_connection_attempt.get(prog_id, 0)
        
        if current_time - last_attempt < 5:
            return self._app_instances.get(prog_id)

        try:
            self._last_connection_attempt[prog_id] = current_time
            app = win32com.client.GetObject(None, prog_id)
            self._app_instances[prog_id] = app
            return app
        except Exception as e:
            logging.debug(f"Could not connect to existing {prog_id}: {e}")
            return None

    def _get_document_info(self, window: int) -> Optional[Dict[str, Any]]:
        try:
            if window in self._cache:
                cache_time, cache_data = self._cache.get(window, (0, None))
                if time.time() - cache_time < self._cache_timeout:
                    return cache_data

            _, pid = win32process.GetWindowThreadProcessId(window)
            process = psutil.Process(pid)
            process_name = process.name().lower()
            
            if process_name not in self.office_apps:
                return None

            app_info = self.office_apps[process_name]
            app = self._get_running_instance(app_info['prog_id'])
            
            if not app:
                return self._create_basic_info(window, process, app_info['name'])

            try:
                active_doc = app.ActiveDocument if app_info['name'] in ['Word', 'PowerPoint'] else app.ActiveWorkbook
                file_path = active_doc.FullName if hasattr(active_doc, 'FullName') else ''
                is_new = not bool(file_path)
                
                # アプリケーションパスの取得
                app_path = file_path if file_path else process.exe()
                
                info = {
                    'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    'process_name': process_name,
                    'window_title': win32gui.GetWindowText(window),
                    'process_id': pid,
                    'file_name': os.path.basename(file_path) if file_path else f'New {app_info["name"]} Document',
                    'file_path': app_path,  # application_path
                    'explorer_path': app_path,  # directory_path（application_pathと同じ値を設定）
                    'office_app_type': app_info['name'],
                    'is_new_document': is_new
                }
                
                self._cache[window] = (time.time(), info)
                return info

            except Exception as e:
                logging.error(f"Error getting document info: {e}")
                return self._create_basic_info(window, process, app_info['name'])

        except Exception as e:
            logging.error(f"Error in _get_document_info: {e}")
            return None

    def _create_basic_info(self, window: int, process: psutil.Process, app_type: str) -> Dict[str, Any]:
        process_path = process.exe()
        return {
            'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'process_name': process.name(),
            'window_title': win32gui.GetWindowText(window),
            'process_id': process.pid,
            'file_name': f'Unknown {app_type} Document',
            'file_path': process_path,  # application_path
            'explorer_path': process_path,  # directory_path（application_pathと同じ値を設定）
            'office_app_type': app_type,
            'is_new_document': True
        }

    def get_active_window_info(self) -> Optional[Dict[str, Any]]:
        try:
            window = win32gui.GetForegroundWindow()
            if not self.is_target_window(window):
                return None

            window_title = win32gui.GetWindowText(window)
            if window_title == self.last_title:
                return None

            info = self._get_document_info(window)
            if info:
                self.last_title = window_title
                return info

            return None

        except Exception as e:
            logging.error(f"Error in get_active_window_info: {e}")
            return None