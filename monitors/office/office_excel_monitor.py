# office_excel_monitor.py
import win32com.client
import win32gui
import win32process
import pythoncom
import psutil
import os
from typing import Optional
import logging
from ...models.window_info import WindowInfo
from ..base.office_base_monitor import OfficeBaseMonitor

class OfficeExcelMonitor(OfficeBaseMonitor):
    def __init__(self):
        super().__init__(
            app_name='Excel.Application',
            process_name='excel.exe',
            app_type='Excel'
        )

    def _get_application_document_info(self, window: int) -> Optional[WindowInfo]:
        """アクティブなExcelドキュメントの情報を取得"""
        try:
            _, pid = win32process.GetWindowThreadProcessId(window)
            process = psutil.Process(pid)
            window_title = win32gui.GetWindowText(window)
            application_path = process.exe()

            # COMオブジェクトを取得
            excel = self._get_com_object()
            if not excel:
                return self._create_basic_info(window, process)

            # アクティブブックの情報を取得
            try:
                active_workbook = excel.ActiveWorkbook
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
                logging.error(f"Error accessing Excel workbook: {e}")
                return self._create_basic_info(window, process)

        except Exception as e:
            logging.error(f"Error in _get_excel_document_info: {e}")
            return None