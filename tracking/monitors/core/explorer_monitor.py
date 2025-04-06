# explorer_monitor.py
import win32com.client
import win32gui
import win32process
import pythoncom
import psutil
import os
from typing import Optional
from ..base.base_monitor import BaseWindowMonitor
from ...models.window_info import WindowInfo

class ExplorerWindowMonitor(BaseWindowMonitor):
    def __init__(self):
        super().__init__()
        self._initialize_com()
        # Explorerの既知のクラス名リスト（拡張版）
        self._explorer_classes = [
            "CabinetWClass",  # 標準的なエクスプローラーウィンドウ
            "ExploreWClass",  # 古いスタイルのエクスプローラー
            "WorkerW",        # デスクトップ
            "Progman",        # プログラムマネージャ（デスクトップ関連）
            "ShellTabWindowClass"  # タブ付きエクスプローラー
        ]

    def _initialize_com(self):
        try:
            # まずCoInitializeを試す
            try:
                pythoncom.CoInitialize()
            except:
                # すでに初期化されている可能性があるため無視
                pass
            
            # 複数の方法を試みる
            try:
                self.shell = win32com.client.Dispatch("Shell.Application")
                print("COM object created via Dispatch successfully")
            except Exception as dispatch_e:
                print(f"Dispatch failed: {dispatch_e}")
                try:
                    # GetObjectも試してみる
                    self.shell = win32com.client.GetObject("Shell.Application")
                    print("COM object created via GetObject successfully")
                except Exception as getobj_e:
                    print(f"GetObject also failed: {getobj_e}")
                    self.shell = None
        except Exception as e:
            print(f"COM initialization failed: {str(e)}")
            self.shell = None

    def __del__(self):
        try:
            pythoncom.CoUninitialize()
        except:
            pass

    def is_target_window(self, window: int) -> bool:
        """対象のExplorerウィンドウかどうかを判定（改善版）"""
        try:
            # クラス名による判定
            class_name = win32gui.GetClassName(window)
            
            # 既知のExplorerクラス名リストとのマッチング
            if class_name in self._explorer_classes:
                return True
                
            # プロセス名による補助判定
            _, pid = win32process.GetWindowThreadProcessId(window)
            try:
                process = psutil.Process(pid)
                process_name = process.name().lower()
                
                # explorer.exeプロセスかつウィンドウタイトルがパスっぽい
                if process_name == "explorer.exe":
                    window_title = win32gui.GetWindowText(window)
                    # 「エクスプローラー」という文字が含まれるか、パスのような形式か
                    if " - エクスプローラー" in window_title or " - Explorer" in window_title or \
                       ":\\" in window_title or "/" in window_title:
                        return True
            except Exception as proc_e:
                print(f"Process check error: {proc_e}")
                pass
                
            return False
        except Exception as e:
            print(f"Error in Explorer is_target_window: {e}")
            return False

    def get_active_window_info(self):
        """エクスプローラーのウィンドウ情報を取得（改善版）"""
        try:
            window = win32gui.GetForegroundWindow()
            if not self.is_target_window(window):
                return None

            window_title = win32gui.GetWindowText(window)
            if window_title == self.last_title:
                return None
                
            # COMオブジェクト経由でパス取得
            current_directory = self._get_explorer_path(window)
            
            # COMでの取得に失敗した場合は代替手段を試す
            if current_directory is None:
                current_directory = self._get_explorer_path_alternative(window)
                print(f"Using alternative path detection: {current_directory}")

            pid = win32process.GetWindowThreadProcessId(window)[1]
            explorer_path = os.path.join(os.environ['WINDIR'], 'explorer.exe')

            self.last_title = window_title

            return WindowInfo.create(
                process_name='explorer.exe',
                window_title=window_title,
                process_id=pid,
                application_name='explorer.exe',
                application_path=explorer_path,
                working_directory=current_directory,  # 必ず何らかの値を設定
                monitor_type='explorer'
            )

        except Exception as e:
            print(f"Error in Explorer get_active_window_info: {e}")
            # 例外発生時も可能な限り情報を返す
            try:
                window = win32gui.GetForegroundWindow()
                window_title = win32gui.GetWindowText(window)
                pid = win32process.GetWindowThreadProcessId(window)[1]
                return WindowInfo.create(
                    process_name='explorer.exe',
                    window_title=window_title,
                    process_id=pid,
                    application_name='explorer.exe',
                    application_path=os.path.join(os.environ['WINDIR'], 'explorer.exe'),
                    working_directory="explorer://error-recovery",
                    monitor_type='explorer'
                )
            except Exception as recovery_e:
                print(f"Recovery attempt also failed: {recovery_e}")
                return None

    def _get_explorer_path(self, hwnd: int) -> Optional[str]:
        """COMオブジェクト経由でExplorerのパスを取得（既存メソッド）"""
        try:
            if self.shell is None:
                self._initialize_com()
                if self.shell is None:
                    print("COM object still null after re-initialization")
                    return None

            windows = self.shell.Windows()
            hwnd_str = str(hwnd)
            
            for window in windows:
                try:
                    if str(window.HWND) == hwnd_str:
                        return window.Document.Folder.Self.Path
                except Exception as e:
                    continue
            
            return None
        except Exception as e:
            print(f"Error in _get_explorer_path: {e}")
            self.shell = None
            return None
            
    def _get_explorer_path_alternative(self, window: int) -> Optional[str]:
        """COMオブジェクト経由でのパス取得に失敗した場合の代替手段"""
        try:
            # 方法1: ウィンドウタイトルからパスを推測
            window_title = win32gui.GetWindowText(window)
            
            # 「エクスプローラー」という文字を取り除く（日本語UI対応）
            for suffix in [" - エクスプローラー", " - Explorer", " - File Explorer"]:
                if window_title.endswith(suffix):
                    title = window_title[:-len(suffix)]
                    break
            else:
                title = window_title
            
            # パスっぽい形式かチェック
            if os.path.exists(title):
                return title
                
            # 特定のキーワードによる特殊フォルダ識別
            special_folders = {
                "デスクトップ": os.path.join(os.path.expanduser("~"), "Desktop"),
                "Desktop": os.path.join(os.path.expanduser("~"), "Desktop"),
                "ドキュメント": os.path.join(os.path.expanduser("~"), "Documents"),
                "Documents": os.path.join(os.path.expanduser("~"), "Documents"),
                "ピクチャ": os.path.join(os.path.expanduser("~"), "Pictures"),
                "Pictures": os.path.join(os.path.expanduser("~"), "Pictures"),
                "ダウンロード": os.path.join(os.path.expanduser("~"), "Downloads"),
                "Downloads": os.path.join(os.path.expanduser("~"), "Downloads"),
                "ミュージック": os.path.join(os.path.expanduser("~"), "Music"),
                "Music": os.path.join(os.path.expanduser("~"), "Music"),
                "ビデオ": os.path.join(os.path.expanduser("~"), "Videos"),
                "Videos": os.path.join(os.path.expanduser("~"), "Videos"),
                "PC": "C:\\"
            }
            
            for keyword, path in special_folders.items():
                if keyword == title or f"{keyword} " in title:
                    return path
                
            # 一般的なパスの場所をチェック
            common_locations = [
                os.path.expanduser("~"),  # ユーザーホーム
                os.path.join(os.path.expanduser("~"), "Desktop"),  # デスクトップ
                os.path.join(os.path.expanduser("~"), "Documents"),  # ドキュメント
                os.path.join(os.path.expanduser("~"), "Downloads"),  # ダウンロード
                "C:\\",  # Cドライブ
                "D:\\"   # Dドライブ
            ]
            
            # 特定のフォルダ名でのマッチング
            for location in common_locations:
                if location in window_title or os.path.basename(location) in window_title:
                    return location
            
            # 方法2: プロセスの作業ディレクトリを取得
            _, pid = win32process.GetWindowThreadProcessId(window)
            process = psutil.Process(pid)
            try:
                cwd = process.cwd()
                if cwd and cwd != os.environ.get('WINDIR'):
                    return cwd
            except:
                pass
                
            # 特別な表記を返す（パスが特定できないことを示す）
            return f"explorer://{window_title}"
        except Exception as e:
            print(f"Error in alternative path detection: {e}")
            return "explorer://unknown"  # 失敗しても何かしらの値を返す