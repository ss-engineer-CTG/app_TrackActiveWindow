# gui.py
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import psutil
from datetime import datetime
import os
from typing import Optional
from .version import __version__, __app_name__

class TrackerGUI:
    def __init__(self, data_manager):
        self.data_manager = data_manager
        self.root = tk.Tk()
        self.root.title(f"{__app_name__} v{__version__}")
        self.root.geometry("400x300")
        self.is_running = True
        self.setup_gui()

    def setup_gui(self):
        # Status Frame
        status_frame = ttk.LabelFrame(self.root, text="ステータス", padding=10)
        status_frame.pack(fill=tk.X, padx=5, pady=5)

        # Status Labels
        self.status_label = ttk.Label(status_frame, text="ステータス: 実行中")
        self.status_label.pack(anchor=tk.W)

        self.record_count_label = ttk.Label(status_frame, text="今日の記録数: 0")
        self.record_count_label.pack(anchor=tk.W)

        self.last_record_label = ttk.Label(status_frame, text="最新の記録: -")
        self.last_record_label.pack(anchor=tk.W)

        self.memory_label = ttk.Label(status_frame, text="メモリ使用量: 0 MB")
        self.memory_label.pack(anchor=tk.W)

        # Monitor Type Label
        self.monitor_type_label = ttk.Label(status_frame, text="モニタータイプ: -")
        self.monitor_type_label.pack(anchor=tk.W)

        # Control Buttons
        control_frame = ttk.Frame(self.root, padding=10)
        control_frame.pack(fill=tk.X)

        self.pause_button = ttk.Button(control_frame, text="一時停止", command=self.toggle_pause)
        self.pause_button.pack(side=tk.LEFT, padx=5)

        self.export_button = ttk.Button(control_frame, text="CSVエクスポート", command=self.export_csv)
        self.export_button.pack(side=tk.LEFT, padx=5)
        
        self.open_logs_button = ttk.Button(control_frame, text="ログフォルダを開く", command=self.open_logs_folder)
        self.open_logs_button.pack(side=tk.LEFT, padx=5)
        
        # メニューバーの追加
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="エクスポート", command=self.export_csv)
        file_menu.add_command(label="ログフォルダを開く", command=self.open_logs_folder)
        file_menu.add_separator()
        file_menu.add_command(label="終了", command=self.quit_app)
        menubar.add_cascade(label="ファイル", menu=file_menu)
        
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="バージョン情報", command=self.show_about)
        menubar.add_cascade(label="ヘルプ", menu=help_menu)

    def update_status(self):
        if not self.root.winfo_exists():
            return

        process = psutil.Process(os.getpid())
        memory_mb = process.memory_info().rss / 1024 / 1024
        self.memory_label.config(text=f"メモリ使用量: {memory_mb:.1f} MB")
        
        current_date = datetime.now().strftime('%Y%m%d')
        log_file = os.path.join(self.data_manager.logs_dir, f"{current_date}_activity_log.csv")
        record_count = 0
        if os.path.exists(log_file):
            try:
                with open(log_file, 'r', encoding='utf-8') as f:
                    record_count = sum(1 for line in f) - 1
            except UnicodeDecodeError:
                try:
                    with open(log_file, 'r', encoding='shift-jis', errors='ignore') as f:
                        record_count = sum(1 for line in f) - 1
                except Exception as e:
                    print(f"Error reading log file: {e}")

        self.record_count_label.config(text=f"今日の記録数: {record_count}")

        # 最新レコードの情報を更新
        if self.data_manager.buffer:
            latest_record = self.data_manager.buffer[-1]
            
            # モニタータイプの表示
            monitor_type = latest_record.monitor_type
            self.monitor_type_label.config(text=f"モニタータイプ: {monitor_type}")
            
            # 最後のレコードのタイトル表示
            window_title = latest_record.window_title
            display_title = f"{window_title[:30]}..." if len(window_title) > 30 else window_title
            self.last_record_label.config(text=f"最新の記録: {display_title}")
        
        if self.root.winfo_exists():
            self.root.after(1000, self.update_status)

    def toggle_pause(self):
        current_status = self.status_label.cget("text")
        if "実行中" in current_status:
            self.status_label.config(text="ステータス: 一時停止中")
            self.pause_button.config(text="再開")
            self.is_running = False
        else:
            self.status_label.config(text="ステータス: 実行中")
            self.pause_button.config(text="一時停止")
            self.is_running = True

    def export_csv(self):
        self.data_manager.save_buffer()
        messagebox.showinfo("エクスポート完了", f"データを {self.data_manager.logs_dir} に保存しました。")

    def open_logs_folder(self):
        os.startfile(self.data_manager.logs_dir)

    def show_about(self):
        about_text = f"{__app_name__}\nバージョン: {__version__}\n\nアクティブウィンドウを追跡して作業ログを記録するツール"
        messagebox.showinfo("バージョン情報", about_text)

    def quit_app(self):
        if messagebox.askyesno("終了確認", "アプリケーションを終了してもよろしいですか？"):
            self.data_manager.save_buffer()
            self.root.destroy()

    def run(self):
        self.update_status()
        
        # ウィンドウが閉じられたときの処理
        self.root.protocol("WM_DELETE_WINDOW", self.quit_app)
        
        self.root.mainloop()