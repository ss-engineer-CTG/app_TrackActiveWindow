# monitor_facade.py
from typing import Optional
from .window_selector import WindowSelector
from .core.general_monitor import GeneralWindowMonitor
from .core.explorer_monitor import ExplorerWindowMonitor
from .core.pdf_monitor import PDFWindowMonitor
from .core.browser_monitor import BrowserWindowMonitor  # 新規追加：ブラウザモニターのインポート
from .office.office_excel_monitor import OfficeExcelMonitor
from .office.office_word_monitor import OfficeWordMonitor
from .office.office_powerpoint_monitor import OfficePowerPointMonitor
from ..models.window_info import WindowInfo
from ..utils.cache_manager import CacheManager

class WindowMonitorFacade:
    def __init__(self):
        self._selector = WindowSelector()
        self._setup_monitors()
        # 最後のアクティブウィンドウを記録
        self._last_window = None
    
    def _setup_monitors(self) -> None:
        # モニターの優先順位を設定: Explorer -> Excel -> Word -> PowerPoint -> Browser -> PDF -> General
        self._selector.register_monitor('explorer', ExplorerWindowMonitor())
        self._selector.register_monitor('excel', OfficeExcelMonitor())
        self._selector.register_monitor('word', OfficeWordMonitor())
        self._selector.register_monitor('powerpoint', OfficePowerPointMonitor())
        self._selector.register_monitor('browser', BrowserWindowMonitor())  # 新規追加：ブラウザモニターの登録
        self._selector.register_monitor('pdf', PDFWindowMonitor())
        self._selector.register_monitor('default', GeneralWindowMonitor())
    
    def get_active_window_info(self) -> Optional[WindowInfo]:
        info = self._selector.get_window_info()
        
        # 同じウィンドウが連続して検出された場合はスキップ
        if self._last_window == info:
            return None
            
        self._last_window = info
        
        if info:
            print(f"Window Info: {info}")
        return info
    
    def __del__(self):
        for monitor in self._selector.monitors.values():
            if hasattr(monitor, '__del__'):
                monitor.__del__()