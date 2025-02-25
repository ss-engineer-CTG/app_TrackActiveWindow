# window_selector.py
from typing import Dict, Optional, List
import win32gui
from .base_monitor import BaseWindowMonitor
from .window_info import WindowInfo

class WindowSelector:
    def __init__(self):
        self.monitors: Dict[str, BaseWindowMonitor] = {}
        self.monitor_order: List[str] = []

    def register_monitor(self, window_class: str, monitor: BaseWindowMonitor) -> None:
        self.monitors[window_class] = monitor
        
        # 優先順位に基づいてモニターを追加
        # Explorer -> Excel -> Word -> PowerPoint -> その他
        if window_class == 'explorer':
            self._insert_at_position(window_class, 0)
        elif window_class == 'excel': 
            self._insert_after(window_class, 'explorer')
        elif window_class == 'word':
            self._insert_after(window_class, 'excel')
        elif window_class == 'powerpoint':
            self._insert_after(window_class, 'word')
        else:
            self.monitor_order.append(window_class)

    def _insert_at_position(self, window_class: str, position: int) -> None:
        """特定の位置にモニタークラスを挿入"""
        if window_class in self.monitor_order:
            self.monitor_order.remove(window_class)
        
        # 位置が範囲外の場合は末尾に追加
        if position >= len(self.monitor_order):
            self.monitor_order.append(window_class)
        else:
            self.monitor_order.insert(position, window_class)

    def _insert_after(self, window_class: str, after_class: str) -> None:
        """特定のクラスの後にモニタークラスを挿入"""
        if window_class in self.monitor_order:
            self.monitor_order.remove(window_class)
            
        if after_class in self.monitor_order:
            position = self.monitor_order.index(after_class) + 1
            self._insert_at_position(window_class, position)
        else:
            # 指定されたクラスが存在しない場合は先頭に追加
            self._insert_at_position(window_class, 0)

    def get_appropriate_monitor(self, window_handle: int) -> Optional[BaseWindowMonitor]:
        try:
            for monitor_class in self.monitor_order:
                monitor = self.monitors[monitor_class]
                if monitor.is_target_window(window_handle):
                    return monitor
            return self.monitors.get('default')
        except Exception as e:
            print(f"Error in monitor selection: {str(e)}")
            return self.monitors.get('default')

    def get_window_info(self) -> Optional[WindowInfo]:
        try:
            active_window = win32gui.GetForegroundWindow()
            monitor = self.get_appropriate_monitor(active_window)
            if monitor:
                info = monitor.get_active_window_info()
                if info:
                    print(f"Selected monitor: {monitor.__class__.__name__}")
                return info
        except Exception as e:
            print(f"Error in window selection: {str(e)}")
        return None