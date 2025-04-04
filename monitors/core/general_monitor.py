# general_monitor.py
import win32gui
import win32process
import psutil
import os
import logging
from typing import List, Set, Optional, Dict
from ..base.base_monitor import BaseWindowMonitor
from ...models.window_info import WindowInfo
from ...config import Config

class GeneralWindowMonitor(BaseWindowMonitor):
    def __init__(self):
        super().__init__()
        self._config = Config()
        
        # 除外プロセスリスト（config.iniから取得）
        self._excluded_processes: Set[str] = set()
        self._load_excluded_processes()
        
        # 除外ウィンドウクラス（他のモニターで処理されるクラス）- 拡張版
        self._excluded_classes: Set[str] = {
            # Explorer関連クラス（拡張）
            "CabinetWClass", "ExploreWClass",  # 標準エクスプローラー
            "WorkerW", "Progman",  # デスクトップ関連
            "ShellTabWindowClass"  # タブ付きエクスプローラー
        }
        
        # 除外ブラウザプロセス（ブラウザモニターで処理するプロセス）
        self._excluded_browser_processes: Set[str] = {
            "chrome.exe", "firefox.exe", "msedge.exe", "opera.exe", 
            "brave.exe", "iexplore.exe", "safari.exe"
        }
        
        # プロセス情報のキャッシュ（軽量なLRUキャッシュ）
        self._process_cache: Dict[int, Dict] = {}
        self._cache_size = 10  # キャッシュするプロセス数
    
    def _load_excluded_processes(self):
        """設定から除外プロセスリストを読み込む"""
        try:
            excluded_str = self._config.get_value('General', 'excluded_processes')
            if excluded_str:
                self._excluded_processes = set(p.strip() for p in excluded_str.split(','))
        except Exception as e:
            logging.warning(f"Failed to load excluded processes: {e}")

    def is_target_window(self, window: int) -> bool:
        """
        このウィンドウが一般モニターの対象かどうかを判定
        他の特化型モニターの対象でないウィンドウなら対象とする
        """
        try:
            # ウィンドウクラスのチェック
            class_name = win32gui.GetClassName(window)
            if not class_name or class_name in self._excluded_classes:
                return False
            
            # Explorer検出の強化（プロセス名とウィンドウタイトルの組み合わせ）
            _, pid = win32process.GetWindowThreadProcessId(window)
            
            # キャッシュからプロセス名を取得
            if pid in self._process_cache:
                process_name = self._process_cache[pid].get('name')
                if process_name:
                    process_name_lower = process_name.lower()
                    # プロセス名がエクスプローラーで、かつタイトルにエクスプローラー特有の特徴がある場合は除外
                    if process_name_lower == "explorer.exe":
                        window_title = win32gui.GetWindowText(window)
                        # エクスプローラーっぽいタイトルパターンをチェック
                        if any(pattern in window_title for pattern in [
                            " - エクスプローラー", " - Explorer", " - File Explorer", ":\\"
                        ]):
                            return False
                    
                    # 除外プロセスまたはブラウザプロセスの場合は除外
                    if (process_name_lower in self._excluded_processes or 
                        process_name_lower in self._excluded_browser_processes):
                        return False
            else:
                # キャッシュにない場合は取得してチェック
                try:
                    process = psutil.Process(pid)
                    process_name = process.name().lower()
                    
                    # プロセス名がエクスプローラーで、かつタイトルにエクスプローラー特有の特徴がある場合は除外
                    if process_name == "explorer.exe":
                        window_title = win32gui.GetWindowText(window)
                        # エクスプローラーっぽいタイトルパターンをチェック
                        if any(pattern in window_title for pattern in [
                            " - エクスプローラー", " - Explorer", " - File Explorer", ":\\"
                        ]):
                            return False
                    
                    if (process_name in self._excluded_processes or 
                        process_name in self._excluded_browser_processes):
                        return False
                except:
                    pass
            
            return True
        except Exception as e:
            logging.error(f"Error in General is_target_window: {e}")
            return False

    def get_active_window_info(self) -> Optional[WindowInfo]:
        """アクティブウィンドウの情報を取得"""
        try:
            window = win32gui.GetForegroundWindow()
            window_title = win32gui.GetWindowText(window)
            
            # 同じウィンドウタイトルが連続して検出された場合はスキップ
            if window_title == self.last_title:
                return None

            pid = win32process.GetWindowThreadProcessId(window)[1]
            
            # プロセス情報を取得（キャッシュを活用）
            if pid in self._process_cache:
                cache = self._process_cache[pid]
                process_name = cache['name']
                application_path = cache['path']
            else:
                process = psutil.Process(pid)
                process_name = process.name()
                application_path = process.exe()
                
                # 作業ディレクトリを取得（可能な場合）
                working_dir = ''
                try:
                    working_dir = process.cwd()
                except:
                    pass
                
                # キャッシュを更新
                self._process_cache[pid] = {
                    'name': process_name,
                    'path': application_path,
                    'working_dir': working_dir
                }
                
                # キャッシュサイズを制限
                if len(self._process_cache) > self._cache_size:
                    # 最も古いエントリを削除 (簡易版LRU)
                    oldest_pid = next(iter(self._process_cache))
                    self._process_cache.pop(oldest_pid)
            
            # 作業ディレクトリを取得
            working_directory = self._process_cache[pid].get('working_dir', '')
            
            self.last_title = window_title

            return WindowInfo.create(
                process_name=process_name,
                window_title=window_title,
                process_id=pid,
                application_name=os.path.basename(application_path),
                application_path=application_path,
                working_directory=working_directory,  # 利用可能な場合は作業ディレクトリを設定
                monitor_type='general'
            )

        except Exception as e:
            logging.error(f"Error in General get_active_window_info: {e}", exc_info=True)
            return None
    
    def reset_cache(self):
        """キャッシュをリセット"""
        self._process_cache.clear()