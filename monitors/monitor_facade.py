# monitor_facade.py
from typing import Optional
from .window_selector import WindowSelector
from .general_monitor import GeneralWindowMonitor
from .explorer_monitor import ExplorerWindowMonitor
from .office_excel_monitor import OfficeExcelMonitor
from .office_word_monitor import OfficeWordMonitor
from .office_powerpoint_monitor import OfficePowerPointMonitor
from .window_info import WindowInfo

class WindowMonitorFacade:
    def __init__(self):
        self._selector = WindowSelector()
        self._setup_monitors()
    
    def _setup_monitors(self) -> None:
        # モニターの優先順位を設定: Explorer -> Excel -> Word -> PowerPoint -> General
        self._selector.register_monitor('explorer', ExplorerWindowMonitor())
        self._selector.register_monitor('excel', OfficeExcelMonitor())
        self._selector.register_monitor('word', OfficeWordMonitor())
        self._selector.register_monitor('powerpoint', OfficePowerPointMonitor())
        self._selector.register_monitor('default', GeneralWindowMonitor())
    
    def get_active_window_info(self) -> Optional[WindowInfo]:
        info = self._selector.get_window_info()
        if info:
            print(f"Window Info: {info}")
        return info
    
    def __del__(self):
        for monitor in self._selector.monitors.values():
            if hasattr(monitor, '__del__'):
                monitor.__del__()