# pdf_monitor.py
import win32gui
import win32process
import psutil
import os
import re
from typing import Optional, Dict, List, Set
from ..base.base_monitor import BaseWindowMonitor
from ...models.window_info import WindowInfo

class PDFWindowMonitor(BaseWindowMonitor):
    """PDFリーダーアプリケーション監視クラス"""
    
    def __init__(self):
        super().__init__()
        # 一般的なPDFリーダープロセス名のリスト
        self.pdf_processes: Set[str] = {
            'acrobat.exe',         # Adobe Acrobat
            'acrord32.exe',        # Adobe Reader
            'SumatraPDF.exe',      # Sumatra PDF
            'FoxitPDFReader.exe',  # Foxit Reader
            'PDFXEdit.exe',        # PDF-XChange Editor
            'chrome.exe',          # Chrome (PDF表示時)
            'msedge.exe',          # Edge (PDF表示時)
            'firefox.exe',         # Firefox (PDF表示時)
            'evince.exe',          # Evince
            'xpdf.exe',            # XPdf
            'PDFCreator.exe',      # PDF Creator
            'PdfFactory.exe'       # PDF Factory
        }
        # PDF拡張子
        self.pdf_extension = '.pdf'
        
    def is_target_window(self, window: int) -> bool:
        """PDFリーダーウィンドウかどうかを判定"""
        try:
            # ウィンドウのプロセスIDを取得
            _, pid = win32process.GetWindowThreadProcessId(window)
            process = psutil.Process(pid)
            process_name = process.name().lower()
            
            # 既知のPDFプロセスかチェック
            if process_name in self.pdf_processes:
                return True
                
            # ウィンドウタイトルでPDFを検出
            window_title = win32gui.GetWindowText(window)
            return window_title.lower().endswith(self.pdf_extension)
        except Exception as e:
            print(f"Error in PDF is_target_window: {e}")
            return False
    
    def get_active_window_info(self) -> Optional[WindowInfo]:
        """アクティブなPDFウィンドウの情報を取得"""
        try:
            window = win32gui.GetForegroundWindow()
            if not self.is_target_window(window):
                return None
                
            window_title = win32gui.GetWindowText(window)
            if window_title == self.last_title:
                return None
                
            # プロセス情報取得
            pid = win32process.GetWindowThreadProcessId(window)[1]
            process = psutil.Process(pid)
            process_name = process.name()
            application_path = process.exe()
            
            # PDF関連情報抽出
            pdf_path = self._extract_pdf_path(window, process, window_title)
            
            self.last_title = window_title
            
            return WindowInfo.create(
                process_name=process_name,
                window_title=window_title,
                process_id=pid,
                application_name=os.path.basename(application_path),
                application_path=application_path,
                working_directory=pdf_path or '',  # PDFファイルのパス（取得できない場合は空文字）
                monitor_type='pdf',  # PDFタイプとして記録
                is_new_document=False,
                office_app_type=None  # Officeではないのでなし
            )
            
        except Exception as e:
            print(f"Error in PDF get_active_window_info: {e}")
            return None
    
    def _extract_pdf_path(self, window: int, process: psutil.Process, window_title: str) -> Optional[str]:
        """PDFのファイルパスを抽出する試み
        
        複数の方法を試み、利用可能な情報からPDFパスを特定します:
        1. ウィンドウタイトルからの抽出
        2. プロセスのコマンドライン引数
        3. 開いているファイルの検査
        """
        # 方法1: タイトルから抽出
        # 一般的なパターン: "ファイル名.pdf - リーダー名" または "リーダー名 - ファイル名.pdf"
        title_patterns = [
            r'(.*\.pdf)[\s]*[-–—]',  # ファイル名.pdf - リーダー名
            r'[-–—][\s]*(.*\.pdf)',   # リーダー名 - ファイル名.pdf
            r'(.*\.pdf)'             # 単にファイル名.pdf
        ]
        
        for pattern in title_patterns:
            match = re.search(pattern, window_title, re.IGNORECASE)
            if match:
                pdf_name = match.group(1).strip()
                # ファイル名だけの場合は、現在のディレクトリから検索
                for dir_path in [os.getcwd(), os.path.expanduser("~/Documents")]:
                    potential_path = os.path.join(dir_path, pdf_name)
                    if os.path.exists(potential_path):
                        return potential_path
        
        # 方法2: コマンドライン引数から抽出
        try:
            cmd_line = process.cmdline()
            for arg in cmd_line:
                if arg.lower().endswith('.pdf') and os.path.exists(arg):
                    return arg
        except:
            pass
            
        # 方法3: 開いているファイルを調査
        try:
            for file in process.open_files():
                if file.path.lower().endswith('.pdf'):
                    return file.path
        except:
            pass
            
        return None