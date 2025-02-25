# office_powerpoint_monitor.py
from .base_monitor import BaseWindowMonitor
import win32com.client
import win32gui
import win32process
import pythoncom
import psutil
import os
from typing import Optional
import logging
import time
from .window_info import WindowInfo

class OfficePowerPointMonitor(BaseWindowMonitor):
    def __init__(self):
        super().__init__()
        self._initialize_com()
        self._powerpoint_instance = None
        self._last_connection_attempt = 0
        self._cache = {}
        self._cache_timeout = 5  # キャッシュタイムアウト（秒）
        self._retry_interval = 5  # 再接続試行間隔（秒）

    def _initialize_com(self) -> None:
        """COMオブジェクトの初期化"""
        try:
            pythoncom.CoInitialize()
            self._powerpoint_instance = win32com.client.GetObject(None, 'PowerPoint.Application')
        except Exception as e:
            logging.error(f"PowerPoint COM initialization failed: {e}")
            self._powerpoint_instance = None

    def __del__(self):
        """リソースのクリーンアップ"""
        try:
            pythoncom.CoUninitialize()
        except:
            pass

    def is_target_window(self, window: int) -> bool:
        """PowerPointウィンドウかどうかを判定"""
        try:
            _, pid = win32process.GetWindowThreadProcessId(window)
            process = psutil.Process(pid)
            return process.name().lower() == 'powerpnt.exe'
        except:
            return False

    def get_active_window_info(self) -> Optional[WindowInfo]:
        """アクティブウィンドウの情報を取得"""
        try:
            window = win32gui.GetForegroundWindow()
            if not self.is_target_window(window):
                return None

            window_title = win32gui.GetWindowText(window)
            if window_title == self.last_title:
                return None

            # キャッシュチェック
            if window in self._cache:
                cache_time, cache_data = self._cache.get(window, (0, None))
                if time.time() - cache_time < self._cache_timeout:
                    return cache_data

            info = self._get_powerpoint_document_info(window)
            if info:
                self.last_title = window_title
                self._cache[window] = (time.time(), info)

            return info

        except Exception as e:
            logging.error(f"Error in PowerPoint get_active_window_info: {e}")
            return None

    def _get_powerpoint_document_info(self, window: int) -> Optional[WindowInfo]:
        """アクティブなPowerPointプレゼンテーションの情報を取得"""
        try:
            _, pid = win32process.GetWindowThreadProcessId(window)
            process = psutil.Process(pid)
            window_title = win32gui.GetWindowText(window)
            application_path = process.exe()

            # 前回の接続試行から一定時間経過していれば再接続を試みる
            current_time = time.time()
            if not self._powerpoint_instance and current_time - self._last_connection_attempt >= self._retry_interval:
                self._last_connection_attempt = current_time
                try:
                    self._powerpoint_instance = win32com.client.GetObject(None, 'PowerPoint.Application')
                except Exception as e:
                    logging.debug(f"Could not connect to PowerPoint: {e}")

            if not self._powerpoint_instance:
                return self._create_basic_info(window, process)

            # アクティブプレゼンテーションの情報を取得
            try:
                active_presentation = self._powerpoint_instance.ActivePresentation
                presentation_path = active_presentation.FullName if hasattr(active_presentation, 'FullName') else ''
                is_new = not bool(presentation_path)
                
                if is_new:
                    return WindowInfo.create(
                        process_name='powerpnt.exe',
                        window_title=window_title,
                        process_id=pid,
                        application_name='powerpnt.exe',
                        application_path=application_path,
                        working_directory='',
                        monitor_type='office',
                        is_new_document=True,
                        office_app_type='PowerPoint'
                    )

                return WindowInfo.create(
                    process_name='powerpnt.exe',
                    window_title=window_title,
                    process_id=pid,
                    application_name='powerpnt.exe',
                    application_path=application_path,
                    working_directory=presentation_path,  # ← ここを変更: フォルダパスではなくフルパスを使用
                    monitor_type='office',
                    is_new_document=False,
                    office_app_type='PowerPoint'
                )
            except Exception as e:
                logging.error(f"Error accessing PowerPoint presentation: {e}")
                return self._create_basic_info(window, process)

        except Exception as e:
            logging.error(f"Error in _get_powerpoint_document_info: {e}")
            return self._create_basic_info(window, psutil.Process(pid))

    def _create_basic_info(self, window: int, process: psutil.Process) -> WindowInfo:
        """基本的なウィンドウ情報を作成"""
        application_path = process.exe()
        return WindowInfo.create(
            process_name='powerpnt.exe',
            window_title=win32gui.GetWindowText(window),
            process_id=process.pid,
            application_name='powerpnt.exe',
            application_path=application_path,
            working_directory='',
            monitor_type='office',
            is_new_document=True,
            office_app_type='PowerPoint'
        )