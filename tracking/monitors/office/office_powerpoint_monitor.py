# office_powerpoint_monitor.py
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

class OfficePowerPointMonitor(OfficeBaseMonitor):
    def __init__(self):
        super().__init__(
            app_name='PowerPoint.Application',
            process_name='powerpnt.exe',
            app_type='PowerPoint'
        )

    def _get_application_document_info(self, window: int) -> Optional[WindowInfo]:
        """アクティブなPowerPointプレゼンテーションの情報を取得（改善版）"""
        try:
            _, pid = win32process.GetWindowThreadProcessId(window)
            process = psutil.Process(pid)
            window_title = win32gui.GetWindowText(window)
            application_path = process.exe()
            
            # ドキュメントパス（初期値は空）
            document_path = ''
            is_new_document = True
            
            # 方法1: COMオブジェクト経由でパス取得
            powerpoint = self._get_com_object()
            if powerpoint:
                try:
                    active_presentation = powerpoint.ActivePresentation
                    if hasattr(active_presentation, 'FullName') and active_presentation.FullName:
                        document_path = active_presentation.FullName
                        is_new_document = False
                        logging.debug(f"PowerPoint path via COM: {document_path}")
                except Exception as e:
                    logging.warning(f"Error accessing PowerPoint presentation via COM: {e}")
            
            # 方法2: COMでの取得に失敗した場合は代替手段を試す
            if not document_path:
                alternative_path = self._get_document_path_alternative(window, process)
                if alternative_path:
                    document_path = alternative_path
                    is_new_document = False
                    logging.debug(f"PowerPoint path via alternative method: {document_path}")
            
            # 結果を返す
            return WindowInfo.create(
                process_name='powerpnt.exe',
                window_title=window_title,
                process_id=pid,
                application_name='powerpnt.exe',
                application_path=application_path,
                working_directory=document_path,
                monitor_type='office',
                is_new_document=is_new_document,
                office_app_type='PowerPoint'
            )

        except Exception as e:
            logging.error(f"Error in _get_powerpoint_document_info: {e}")
            return self._create_basic_info(window, process)