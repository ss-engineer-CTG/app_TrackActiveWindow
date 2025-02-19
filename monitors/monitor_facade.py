# monitor_facade.py
from .general_monitor import GeneralWindowMonitor
from .explorer_monitor import ExplorerWindowMonitor
from .office_monitor import OfficeWindowMonitor
from .window_selector import WindowSelector

class WindowMonitorFacade:
    def __init__(self):
        self._selector = WindowSelector()
        self._setup_monitors()
    
    def _setup_monitors(self):
        # モニターの優先順位を設定: Explorer -> Office -> General
        self._selector.register_monitor('explorer', ExplorerWindowMonitor())
        self._selector.register_monitor('office', OfficeWindowMonitor())
        self._selector.register_monitor('default', GeneralWindowMonitor())
    
    def get_active_window_info(self):
        info = self._selector.get_window_info()
        if info:
            print(f"Window Info: {info}")  # デバッグ用
        return info
    
    def __del__(self):
        for monitor in self._selector.monitors.values():
            if hasattr(monitor, '__del__'):
                monitor.__del__()