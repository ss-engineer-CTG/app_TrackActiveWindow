# data_manager.py
import csv
import os
from datetime import datetime
import shutil
import re

class DataManager:
    def __init__(self, buffer_size=500):
        self.buffer = []
        self.buffer_size = buffer_size
        self.base_dir = os.path.dirname(os.path.abspath(__file__))
        self.logs_dir = os.path.join(self.base_dir, 'logs')
        self.temp_dir = os.path.join(self.base_dir, 'temp')
        self.setup_directories()

    def setup_directories(self):
        for directory in [self.logs_dir, self.temp_dir]:
            if not os.path.exists(directory):
                os.makedirs(directory)
                print(f"Created directory: {directory}")

    def _sanitize_text(self, text):
        """Remove problematic Unicode characters while preserving standard text"""
        if not isinstance(text, str):
            return str(text)
        # Remove zero-width spaces and other invisible characters
        cleaned = re.sub(r'[\u200b\u200c\u200d\ufeff\u200e\u200f]', '', text)
        return cleaned

    def _sanitize_record(self, record):
        """Sanitize all text fields in the record"""
        return {
            key: self._sanitize_text(value)
            for key, value in record.items()
        }

    def add_record(self, record):
        if record:
            sanitized_record = self._sanitize_record(record)
            self.buffer.append(sanitized_record)
            if len(self.buffer) >= self.buffer_size * 0.8:
                self.save_buffer()

    def save_buffer(self):
        if not self.buffer:
            return

        current_date = datetime.now().strftime('%Y%m%d')
        filename = f"{current_date}_activity_log.csv"
        filepath = os.path.join(self.logs_dir, filename)
        temp_filepath = os.path.join(self.temp_dir, f"temp_{filename}")
        utf8_filepath = os.path.join(self.logs_dir, f"{current_date}_activity_log_utf8.csv")

        try:
            # Save to temporary file first
            self._write_to_csv(temp_filepath)
            
            # If successful, copy to actual log files
            shutil.copy2(temp_filepath, filepath)
            
            # Create UTF-8 backup
            self._write_to_csv(utf8_filepath, encoding='utf-8')
            
            # Clear buffer after successful save
            self.buffer.clear()
        except Exception as e:
            self._log_error(str(e))

    def _write_to_csv(self, filepath, encoding='shift-jis'):
        mode = 'a' if os.path.exists(filepath) else 'w'
        with open(filepath, mode, encoding=encoding, newline='', errors='replace') as f:
            writer = csv.DictWriter(f, fieldnames=[
                'timestamp', 'process_name', 'window_title', 
                'process_id', 'file_name', 'file_path'
            ])
            if mode == 'w':
                writer.writeheader()
            writer.writerows(self.buffer)

    def _log_error(self, error_message):
        error_log_path = os.path.join(self.logs_dir, 'tracking_error.log')
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        with open(error_log_path, 'a', encoding='utf-8') as f:
            f.write(f"{timestamp}: {error_message}\n")