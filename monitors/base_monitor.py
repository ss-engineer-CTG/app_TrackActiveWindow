# base_monitor.py
class BaseWindowMonitor:
    def __init__(self):
        self.last_title = None

    def get_active_window_info(self):
        raise NotImplementedError("Subclasses must implement get_active_window_info()")

    def is_target_window(self, window):
        raise NotImplementedError("Subclasses must implement is_target_window()")