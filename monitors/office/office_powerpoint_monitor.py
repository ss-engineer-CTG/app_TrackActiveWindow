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
        """アクティブなPowerPointプレゼンテーションの情報を取得"""
        try:
            _, pid = win32process.GetWindowThreadProcessId(window)
            process = psutil.Process(pid)
            window_title = win32gui.GetWindowText(window)
            application_path = process.exe()

            # COMオブジェクトを取得
            powerpoint = self._get_com_object()
            if not powerpoint:
                return self._create_basic_info(window, process)

            # アクティブプレゼンテーションの情報を取得
            try:
                active_presentation = powerpoint.ActivePresentation
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
                    working_directory=presentation_path,
                    monitor_type='office',
                    is_new_document=False,
                    office_app_type='PowerPoint'
                )
            except Exception as e:
                logging.error(f"Error accessing PowerPoint presentation: {e}")
                return self._create_basic_info(window, process)

        except Exception as e:
            logging.error(f"Error in _get_powerpoint_document_info: {e}")
            return self._create_basic_info(window, process)