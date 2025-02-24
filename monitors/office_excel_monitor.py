# office_excel_monitor.py
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

    def _get_excel_document_info(self, window: int) -> Optional[Dict[str, Any]]:
        """アクティブなExcelドキュメントの情報を取得"""
        try:
            _, pid = win32process.GetWindowThreadProcessId(window)
            process = psutil.Process(pid)
            window_title = win32gui.GetWindowText(window)
            exe_path = process.exe()  # Excelアプリケーションのパス

            if not self._excel_instance:
                self._excel_instance = win32com.client.GetObject(None, 'Excel.Application')

            if not self._excel_instance:
                return self._create_basic_info(window, process)

            active_workbook = self._excel_instance.ActiveWorkbook
            workbook_path = active_workbook.FullName if hasattr(active_workbook, 'FullName') else ''
            is_new = not bool(workbook_path)

            if is_new:
                return {
                    'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    'process_name': 'excel.exe',
                    'window_title': window_title,
                    'process_id': pid,
                    'file_name': 'New Excel Workbook',
                    'file_path': exe_path,         # Excelアプリケーションのパス
                    'explorer_path': '',           # 新規文書なのでパスなし
                    'office_app_type': 'Excel',
                    'is_new_document': True
                }

            return {
                'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'process_name': 'excel.exe',
                'window_title': window_title,
                'process_id': pid,
                'file_name': os.path.basename(workbook_path),
                'file_path': exe_path,              # Excelアプリケーションのパス
                'explorer_path': workbook_path,     # ブックの完全パス
                'office_app_type': 'Excel',
                'is_new_document': False
            }

        except Exception as e:
            logging.error(f"Error in _get_excel_document_info: {e}")
            return None

    def _create_basic_info(self, window: int, process: psutil.Process) -> Dict[str, Any]:
        """基本的なウィンドウ情報を作成"""
        exe_path = process.exe()
        return {
            'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'process_name': 'excel.exe',
            'window_title': win32gui.GetWindowText(window),
            'process_id': process.pid,
            'file_name': 'Unknown Excel Workbook',
            'file_path': exe_path,          # Excelアプリケーションのパス
            'explorer_path': '',            # 不明な場合は空文字列
            'office_app_type': 'Excel',
            'is_new_document': True
        }

    def get_active_window_info(self) -> Optional[Dict[str, Any]]:
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