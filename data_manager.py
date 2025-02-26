# data_manager.py
import csv
import os
import re
from datetime import datetime
import shutil
import threading
import codecs
from typing import List, Dict, Any, Optional, Set
from .models.window_info import WindowInfo

class DataManager:
    def __init__(self, buffer_size: int = 500):
        self.buffer: List[WindowInfo] = []
        self.buffer_size = buffer_size
        self.buffer_lock = threading.Lock()
        # 重複検出用ハッシュセット
        self.window_hash_set: Set[str] = set()
        
        self.base_dir = os.path.dirname(os.path.abspath(__file__))
        self.logs_dir = os.path.join(self.base_dir, 'logs')
        self.temp_dir = os.path.join(self.base_dir, 'temp')
        
        self.setup_directories()

    def setup_directories(self):
        for directory in [self.logs_dir, self.temp_dir]:
            if not os.path.exists(directory):
                os.makedirs(directory)

    def _generate_window_hash(self, record: WindowInfo) -> str:
        """WindowInfoから一意のハッシュを生成"""
        return f"{record.process_id}_{record.process_name}_{record.window_title}"

    def add_record(self, record: Optional[WindowInfo]) -> None:
        if not record:
            return

        window_hash = self._generate_window_hash(record)

        with self.buffer_lock:
            # 重複チェック (同じウィンドウの場合はスキップ)
            if window_hash in self.window_hash_set:
                return
                
            # 新しいウィンドウとして追加
            self.buffer.append(record)
            self.window_hash_set.add(window_hash)
            print(f"Record added to buffer. Buffer size: {len(self.buffer)}")
            
            if len(self.buffer) >= self.buffer_size * 0.8:
                self.save_buffer()

    def save_buffer(self) -> None:
        if not self.buffer:
            return

        with self.buffer_lock:
            current_date = datetime.now().strftime('%Y%m%d')
            filename = f"{current_date}_activity_log.csv"
            filepath = os.path.join(self.logs_dir, filename)
            temp_filepath = os.path.join(self.temp_dir, f"temp_{filename}")

            try:
                self._write_to_csv(temp_filepath)
                print(f"Temporary file written: {temp_filepath}")
                
                shutil.copy2(temp_filepath, filepath)
                print(f"Log file updated: {filepath}")
                
                # バッファとハッシュセットを両方クリア
                self.buffer.clear()
                self.window_hash_set.clear()
            except Exception as e:
                self._log_error(f"Error saving buffer: {str(e)}")
                print(f"Error saving buffer: {str(e)}")

    def _sanitize_text(self, text: str) -> str:
        """文字列をサニタイズする"""
        if not text:
            return ""
            
        # 制御文字や特殊文字を削除
        sanitized = re.sub(r'[\u0000-\u001F\u007F-\u009F\u200B-\u200F\u2028-\u202F]', '', text)
        return sanitized

    def _write_to_csv(self, filepath: str) -> None:
        fieldnames = [
            'timestamp',
            'process_name',
            'window_title',
            'process_id',
            'application_name',
            'application_path',
            'working_directory',
            'monitor_type',
            'is_new_document',
            'office_app_type'
        ]

        mode = 'a' if os.path.exists(filepath) else 'w'
        
        # codecs.openの代わりに標準のopen関数を使用する
        with open(filepath, mode, encoding='utf-8-sig', newline='') as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            if mode == 'w':
                writer.writeheader()
            
            for record in self.buffer:
                writer.writerow({
                    'timestamp': record.timestamp,
                    'process_name': record.process_name,
                    'window_title': self._sanitize_text(record.window_title),
                    'process_id': record.process_id,
                    'application_name': record.application_name,
                    'application_path': self._sanitize_text(record.application_path),
                    'working_directory': self._sanitize_text(record.working_directory),
                    'monitor_type': record.monitor_type,
                    'is_new_document': record.is_new_document,
                    'office_app_type': record.office_app_type or ''
                })

    def _log_error(self, error_message: str) -> None:
        error_log_path = os.path.join(self.logs_dir, 'tracking_error.log')
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        with open(error_log_path, 'a', encoding='utf-8') as f:
            f.write(f"{timestamp}: {error_message}\n")