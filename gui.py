# gui.py
import tkinter as tk
from tkinter import ttk
import psutil
from datetime import datetime
import os

class TrackerGUI:
    def __init__(self, data_manager):
        self.data_manager = data_manager
        self.root = tk.Tk()
        self.root.title("Window Tracker")
        self.root.geometry("400x300")
        self.setup_gui()

    def setup_gui(self):
        # Status Frame
        status_frame = ttk.LabelFrame(self.root, text="Status", padding=10)
        status_frame.pack(fill=tk.X, padx=5, pady=5)

        # Status Labels
        self.status_label = ttk.Label(status_frame, text="Status: Running")
        self.status_label.pack(anchor=tk.W)

        self.record_count_label = ttk.Label(status_frame, text="Records Today: 0")
        self.record_count_label.pack(anchor=tk.W)

        self.last_record_label = ttk.Label(status_frame, text="Last Record: -")
        self.last_record_label.pack(anchor=tk.W)

        self.memory_label = ttk.Label(status_frame, text="Memory Usage: 0 MB")
        self.memory_label.pack(anchor=tk.W)

        # Control Buttons
        control_frame = ttk.Frame(self.root, padding=10)
        control_frame.pack(fill=tk.X)

        self.pause_button = ttk.Button(control_frame, text="Pause", command=self.toggle_pause)
        self.pause_button.pack(side=tk.LEFT, padx=5)

        self.export_button = ttk.Button(control_frame, text="Export CSV", command=self.export_csv)
        self.export_button.pack(side=tk.LEFT, padx=5)

    def update_status(self):
        # Update memory usage
        process = psutil.Process(os.getpid())
        memory_mb = process.memory_info().rss / 1024 / 1024

        # Update labels
        self.memory_label.config(text=f"Memory Usage: {memory_mb:.1f} MB")
        
        # Count today's records
        current_date = datetime.now().strftime('%Y%m%d')
        log_file = os.path.join(self.data_manager.logs_dir, f"{current_date}_activity_log.csv")
        record_count = 0
        if os.path.exists(log_file):
            with open(log_file, 'r', encoding='shift-jis') as f:
                record_count = sum(1 for line in f) - 1  # Subtract header row

        self.record_count_label.config(text=f"Records Today: {record_count}")
        
        # Schedule next update
        self.root.after(1000, self.update_status)

    def toggle_pause(self):
        # Implementation for pause/resume functionality
        pass

    def export_csv(self):
        # Implementation for CSV export functionality
        pass

    def run(self):
        self.update_status()
        self.root.mainloop()