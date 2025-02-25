# data_manager.py
import csv
import os
from datetime import datetime
import shutil
import threading
from typing import List, Dict, Any, Optional
from .monitors.window_info import WindowInfo

class DataManager:
    def __init__(self, buffer_size: int = 500):
        self.buffer: List[WindowInfo] = []
        self.buffer_size = buffer_size
        self.buffer_lock = threading.Lock()
        
        self.base_dir = os.path.dirname(os.path.abspath(__file__))
        self.logs_dir = os.path.join(self.base_dir, 'logs')
        self.temp_dir = os.path.join(self.base_dir, 'temp')
        
        self.setup_directories()

    def setup_directories(self):
        for directory in [self.logs_dir, self.temp_dir]:
            if not os.path.exists(directory):
                os.makedirs(directory)

    def add_record(self, record: Optional[WindowInfo]) -> None:
        if not record:
            return

        with self.buffer_lock:
            self.buffer.append(record)
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
                
                self.buffer.clear()
            except Exception as e:
                self._log_error(f"Error saving buffer: {str(e)}")
                print(f"Error saving buffer: {str(e)}")

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
        with open(filepath, mode, encoding='shift-jis', newline='') as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            if mode == 'w':
                writer.writeheader()
            
            for record in self.buffer:
                writer.writerow({
                    'timestamp': record.timestamp,
                    'process_name': record.process_name,
                    'window_title': record.window_title,
                    'process_id': record.process_id,
                    'application_name': record.application_name,
                    'application_path': record.application_path,
                    'working_directory': record.working_directory,
                    'monitor_type': record.monitor_type,
                    'is_new_document': record.is_new_document,
                    'office_app_type': record.office_app_type or ''
                })

    def _log_error(self, error_message: str) -> None:
        error_log_path = os.path.join(self.logs_dir, 'tracking_error.log')
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        with open(error_log_path, 'a', encoding='utf-8') as f:
            f.write(f"{timestamp}: {error_message}\n")