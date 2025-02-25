# office_word_monitor.py
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

class OfficeWordMonitor(BaseWindowMonitor):
    def __init__(self):
        super().__init__()
        self._initialize_com()
        self._word_instance = None
        self._last_connection_attempt = 0
        self._cache = {}
        self._cache_timeout = 5  # キャッシュタイムアウト（秒）
        self._retry_interval = 5  # 再接続試行間隔（秒）

    def _initialize_com(self) -> None:
        """COMオブジェクトの初期化"""
        try:
            pythoncom.CoInitialize()
            self._word_instance = win32com.client.GetObject(None, 'Word.Application')
        except Exception as e:
            logging.error(f"Word COM initialization failed: {e}")
            self._word_instance = None

    def __del__(self):
        """リソースのクリーンアップ"""
        try:
            pythoncom.CoUninitialize()
        except:
            pass

    def is_target_window(self, window: int) -> bool:
        """Wordウィンドウかどうかを判定"""
        try:
            _, pid = win32process.GetWindowThreadProcessId(window)
            process = psutil.Process(pid)
            return process.name().lower() == 'winword.exe'
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

            # キャッシュチェックを追加
            if window in self._cache:
                cache_time, cache_data = self._cache.get(window, (0, None))
                if time.time() - cache_time < self._cache_timeout:
                    return cache_data

            info = self._get_word_document_info(window)
            if info:
                self.last_title = window_title
                # キャッシュ保存を追加
                self._cache[window] = (time.time(), info)

            return info

        except Exception as e:
            logging.error(f"Error in Word get_active_window_info: {e}")
            return None

    def _get_word_document_info(self, window: int) -> Optional[WindowInfo]:
        """アクティブなWordドキュメントの情報を取得"""
        try:
            _, pid = win32process.GetWindowThreadProcessId(window)
            process = psutil.Process(pid)
            window_title = win32gui.GetWindowText(window)
            application_path = process.exe()

            # 前回の接続試行から一定時間経過していれば再接続を試みる
            current_time = time.time()
            if not self._word_instance and current_time - self._last_connection_attempt >= self._retry_interval:
                self._last_connection_attempt = current_time
                try:
                    self._word_instance = win32com.client.GetObject(None, 'Word.Application')
                except Exception as e:
                    logging.debug(f"Could not connect to Word: {e}")

            if not self._word_instance:
                return self._create_basic_info(window, process)

            # アクティブドキュメントの情報を取得
            try:
                active_document = self._word_instance.ActiveDocument
                document_path = active_document.FullName if hasattr(active_document, 'FullName') else ''
                is_new = not bool(document_path)

                if is_new:
                    return WindowInfo.create(
                        process_name='winword.exe',
                        window_title=window_title,
                        process_id=pid,
                        application_name='winword.exe',
                        application_path=application_path,
                        working_directory='',
                        monitor_type='office',
                        is_new_document=True,
                        office_app_type='Word'
                    )

                return WindowInfo.create(
                    process_name='winword.exe',
                    window_title=window_title,
                    process_id=pid,
                    application_name='winword.exe',
                    application_path=application_path,
                    working_directory=document_path,  # ← ここを変更: ディレクトリではなくフルパスを使用
                    monitor_type='office',
                    is_new_document=False,
                    office_app_type='Word'
                )
            except Exception as e:
                logging.error(f"Error accessing Word document: {e}")
                return self._create_basic_info(window, process)

        except Exception as e:
            logging.error(f"Error in _get_word_document_info: {e}")
            return self._create_basic_info(window, psutil.Process(pid))

    def _create_basic_info(self, window: int, process: psutil.Process) -> WindowInfo:
        """基本的なウィンドウ情報を作成"""
        application_path = process.exe()
        return WindowInfo.create(
            process_name='winword.exe',
            window_title=win32gui.GetWindowText(window),
            process_id=process.pid,
            application_name='winword.exe',
            application_path=application_path,
            working_directory='',
            monitor_type='office',
            is_new_document=True,
            office_app_type='Word'
        )