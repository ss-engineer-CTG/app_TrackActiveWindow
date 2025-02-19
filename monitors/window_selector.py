# window_selector.py
import win32gui

class WindowSelector:
    def __init__(self):
        self.monitors = {}
        self.monitor_order = []

    def register_monitor(self, window_class, monitor):
        self.monitors[window_class] = monitor
        # 登録順序の管理
        if window_class == 'explorer':
            self.monitor_order.insert(0, window_class)
        elif window_class == 'office':
            # officeは2番目に優先順位を設定
            if 'explorer' in self.monitor_order:
                self.monitor_order.insert(1, window_class)
            else:
                self.monitor_order.insert(0, window_class)
        else:
            self.monitor_order.append(window_class)

    def get_appropriate_monitor(self, window_handle):
        try:
            # 定義された順序でモニターを評価
            for monitor_class in self.monitor_order:
                monitor = self.monitors[monitor_class]
                if monitor.is_target_window(window_handle):
                    return monitor
            return self.monitors.get('default')
        except Exception as e:
            print(f"Error in monitor selection: {str(e)}")
            return self.monitors.get('default')

    def get_window_info(self):
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