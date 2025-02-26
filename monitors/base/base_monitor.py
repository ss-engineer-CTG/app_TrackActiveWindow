# base_monitor.py
from typing import Optional, Dict, Any

class BaseWindowMonitor:
    def __init__(self):
        self.last_title = None

    def get_active_window_info(self) -> Optional[Dict[str, Any]]:
        raise NotImplementedError("Subclasses must implement get_active_window_info()")

    def is_target_window(self, window: int) -> bool:
        raise NotImplementedError("Subclasses must implement is_target_window()")