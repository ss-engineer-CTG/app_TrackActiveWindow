# main.py
from .config import Config
from .data_manager import DataManager
from .gui import TrackerGUI
from .monitors.monitor_facade import WindowMonitorFacade
import threading
import time
import sys
from .version import __version__, __app_name__

def main():
    print(f"{__app_name__} v{__version__} を起動中...")
    
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
        try:
            write_interval = int(config.get_value('General', 'write_interval'))
            last_write = time.time()

            while True:
                try:
                    window_info = monitor.get_active_window_info()
                    if window_info:
                        data_manager.add_record(window_info)
                except Exception as e:
                    print(f"監視エラー: {e}")

                # 強制的な保存間隔の確認
                if time.time() - last_write >= write_interval:
                    try:
                        data_manager.save_buffer()
                        last_write = time.time()
                    except Exception as e:
                        print(f"保存エラー: {e}")

                time.sleep(1)  # CPU負荷を下げるために1秒待機
        except Exception as e:
            print(f"監視スレッド致命的エラー: {e}")
            sys.exit(1)

    # Start monitoring thread
    monitor_thread = threading.Thread(target=monitor_windows, daemon=True)
    monitor_thread.start()

    # Start GUI
    try:
        gui.run()
    except Exception as e:
        print(f"GUIエラー: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()