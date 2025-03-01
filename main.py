# main.py
from tracking.config import Config
from tracking.data_manager import DataManager
from tracking.gui import TrackerGUI
from tracking.monitors.monitor_facade import WindowMonitorFacade
import threading
import time

def main():
    # Initialize components
    config = Config()
    monitor = WindowMonitorFacade()
    data_manager = DataManager(
        buffer_size=int(config.get_value('General', 'buffer_size'))
    )

    # Initialize GUI
    gui = TrackerGUI(data_manager)

    # Monitoring thread function
    def monitor_windows():
        write_interval = int(config.get_value('General', 'write_interval'))
        last_write = time.time()

        while True:
            window_info = monitor.get_active_window_info()
            if window_info:
                data_manager.add_record(window_info)
                print(f"Window info recorded: {window_info}")  # デバッグ用

            # 強制的な保存間隔の確認
            if time.time() - last_write >= write_interval:
                data_manager.save_buffer()
                last_write = time.time()

            time.sleep(1)  # CPU負荷を下げるために1秒待機

    # Start monitoring thread
    monitor_thread = threading.Thread(target=monitor_windows, daemon=True)
    monitor_thread.start()

    # Start GUI
    gui.run()

if __name__ == "__main__":
    main()