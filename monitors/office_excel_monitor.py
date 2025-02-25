# office_excel_monitor.py
from .base_monitor import BaseWindowMonitor
import win32com.client
import win32gui
import win32process
import pythoncom
import psutil
import os
from typing import Optional, Dict, Any
import logging
from .window_info import WindowInfo

class OfficeExcelMonitor(BaseWindowMonitor):
    def __init__(self):
        super().__init__()
        self._initialize_com()
        self._excel_instance = None

    def _initialize_com(self) -> None:
        """COMオブジェクトの初期化"""
        try:
            pythoncom.CoInitialize()
            self._excel_instance = win32com.client.GetObject(None, 'Excel.Application')
        except Exception as e:
            logging.error(f"COM initialization failed: {e}")
            self._excel_instance = None

    def __del__(self):
        """リソースのクリーンアップ"""
        try:
            pythoncom.CoUninitialize()
        except:
            pass

    def is_target_window(self, window: int) -> bool:
        """Excelウィンドウかどうかを判定"""
        try:
            _, pid = win32process.GetWindowThreadProcessId(window)
            process = psutil.Process(pid)
            return process.name().lower() == 'excel.exe'
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

            info = self._get_excel_document_info(window)
            if info:
                self.last_title = window_title

            return info

        except Exception as e:
            logging.error(f"Error in get_active_window_info: {e}")
            return None

    def _get_excel_document_info(self, window: int) -> Optional[WindowInfo]:
        """アクティブなExcelドキュメントの情報を取得"""
        try:
            _, pid = win32process.GetWindowThreadProcessId(window)
            process = psutil.Process(pid)
            window_title = win32gui.GetWindowText(window)
            application_path = process.exe()

            if not self._excel_instance:
                self._excel_instance = win32com.client.GetObject(None, 'Excel.Application')

            if not self._excel_instance:
                return self._create_basic_info(window, process)

            active_workbook = self._excel_instance.ActiveWorkbook
            workbook_path = active_workbook.FullName if hasattr(active_workbook, 'FullName') else ''
            is_new = not bool(workbook_path)

            if is_new:
                return WindowInfo.create(
                    process_name='excel.exe',
                    window_title=window_title,
                    process_id=pid,
                    application_name='excel.exe',
                    application_path=application_path,
                    working_directory='',
                    monitor_type='office',
                    is_new_document=True,
                    office_app_type='Excel'
                )

            return WindowInfo.create(
                process_name='excel.exe',
                window_title=window_title,
                process_id=pid,
                application_name='excel.exe',
                application_path=application_path,
                working_directory=workbook_path,
                monitor_type='office',
                is_new_document=False,
                office_app_type='Excel'
            )

        except Exception as e:
            logging.error(f"Error in _get_excel_document_info: {e}")
            return None

    def _create_basic_info(self, window: int, process: psutil.Process) -> WindowInfo:
        """基本的なウィンドウ情報を作成"""
        application_path = process.exe()
        return WindowInfo.create(
            process_name='excel.exe',
            window_title=win32gui.GetWindowText(window),
            process_id=process.pid,
            application_name='excel.exe',
            application_path=application_path,
            working_directory='',
            monitor_type='office',
            is_new_document=True,
            office_app_type='Excel'
        )