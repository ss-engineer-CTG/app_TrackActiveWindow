# window_monitor.py
import win32gui
import win32process
import psutil
from datetime import datetime
import os
import re

class WindowMonitor:
    def __init__(self):
        self.last_window = None
        self.excluded_processes = set()

    def set_excluded_processes(self, processes_str):
        """除外プロセスを設定"""
        self.excluded_processes = set(p.strip() for p in processes_str.split(','))

    def _sanitize_text(self, text):
        """Remove problematic Unicode characters while preserving standard text"""
        if not isinstance(text, str):
            return str(text)
        # Remove zero-width spaces and other invisible characters
        cleaned = re.sub(r'[\u200b\u200c\u200d\ufeff\u200e\u200f]', '', text)
        return cleaned

    def get_active_window_info(self):
        try:
            window = win32gui.GetForegroundWindow()
            if window == self.last_window:
                return None

            pid = win32process.GetWindowThreadProcessId(window)[1]
            process = psutil.Process(pid)
            
            # プロセス名を取得して除外チェック
            process_name = process.name()
            if process_name in self.excluded_processes:
                return None

            window_title = win32gui.GetWindowText(window)
            file_path = process.exe()
            file_name = os.path.basename(file_path)

            # 空のウィンドウタイトルはスキップ
            if not window_title.strip():
                return None

            self.last_window = window

            # 全てのテキストフィールドをサニタイズ
            return {
                'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'process_name': self._sanitize_text(process_name),
                'window_title': self._sanitize_text(window_title),
                'process_id': pid,
                'file_name': self._sanitize_text(file_name),
                'file_path': self._sanitize_text(file_path)
            }

        except Exception as e:
            print(f"Error in get_active_window_info: {str(e)}")  # デバッグ用
            return None

    def _is_valid_window(self, window):
        """ウィンドウが有効かどうかをチェック"""
        try:
            return (
                win32gui.IsWindowVisible(window) and
                win32gui.GetWindowText(window).strip() != ""
            )
        except Exception:
            return False