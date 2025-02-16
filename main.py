# main.py
from tracking.config import Config
from tracking.window_monitor import WindowMonitor
from tracking.data_manager import DataManager
from tracking.gui import TrackerGUI
import threading
import time

def main():
    # Initialize components
    config = Config()
    monitor = WindowMonitor()
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

            # Check if it's time to write buffer
            if time.time() - last_write >= write_interval:
                data_manager.save_buffer()
                last_write = time.time()

            time.sleep(1)

    # Start monitoring thread
    monitor_thread = threading.Thread(target=monitor_windows, daemon=True)
    monitor_thread.start()

    # Start GUI
    gui.run()

if __name__ == "__main__":
    main()