# office_word_monitor.py
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

class OfficeWordMonitor(OfficeBaseMonitor):
    def __init__(self):
        super().__init__(
            app_name='Word.Application',
            process_name='winword.exe',
            app_type='Word'
        )

    def _get_application_document_info(self, window: int) -> Optional[WindowInfo]:
        """アクティブなWordドキュメントの情報を取得"""
        try:
            _, pid = win32process.GetWindowThreadProcessId(window)
            process = psutil.Process(pid)
            window_title = win32gui.GetWindowText(window)
            application_path = process.exe()

            # COMオブジェクトを取得
            word = self._get_com_object()
            if not word:
                return self._create_basic_info(window, process)

            # アクティブドキュメントの情報を取得
            try:
                active_document = word.ActiveDocument
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
                    working_directory=document_path,
                    monitor_type='office',
                    is_new_document=False,
                    office_app_type='Word'
                )
            except Exception as e:
                logging.error(f"Error accessing Word document: {e}")
                return self._create_basic_info(window, process)

        except Exception as e:
            logging.error(f"Error in _get_word_document_info: {e}")
            return self._create_basic_info(window, process)