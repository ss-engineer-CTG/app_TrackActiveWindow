# data_manager.py
import csv
import os
from datetime import datetime
import shutil

class DataManager:
    def __init__(self, buffer_size=500):
        self.buffer = []
        self.buffer_size = buffer_size
        self.logs_dir = 'logs'
        self.temp_dir = 'temp'
        self.setup_directories()

    def setup_directories(self):
        for directory in [self.logs_dir, self.temp_dir]:
            if not os.path.exists(directory):
                os.makedirs(directory)

    def add_record(self, record):
        if record:
            self.buffer.append(record)
            if len(self.buffer) >= self.buffer_size * 0.8:
                self.save_buffer()

    def save_buffer(self):
        if not self.buffer:
            return

        current_date = datetime.now().strftime('%Y%m%d')
        filename = f"{current_date}_activity_log.csv"
        filepath = os.path.join(self.logs_dir, filename)
        temp_filepath = os.path.join(self.temp_dir, f"temp_{filename}")

        try:
            # Save to temporary file first
            self._write_to_csv(temp_filepath)
            
            # If successful, copy to actual log file
            shutil.copy2(temp_filepath, filepath)
            
            # Clear buffer after successful save
            self.buffer.clear()
        except Exception as e:
            # Log error
            self._log_error(str(e))

    def _write_to_csv(self, filepath):
        mode = 'a' if os.path.exists(filepath) else 'w'
        with open(filepath, mode, encoding='shift-jis', newline='') as f:
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