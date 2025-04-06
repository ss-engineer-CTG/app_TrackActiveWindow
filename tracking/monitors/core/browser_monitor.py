# browser_monitor.py
import win32gui
import win32process
import psutil
import os
import re
from typing import Optional, Dict, List, Tuple
from ..base.base_monitor import BaseWindowMonitor
from ...models.window_info import WindowInfo

class BrowserWindowMonitor(BaseWindowMonitor):
    """Webブラウザウィンドウ監視クラス"""
    
    def __init__(self):
        super().__init__()
        # 対応するWebブラウザのプロセス名とブラウザ名のマッピング
        self.browser_processes = {
            'chrome.exe': 'Chrome',
            'firefox.exe': 'Firefox',
            'msedge.exe': 'Edge',
            'opera.exe': 'Opera',
            'brave.exe': 'Brave',
            'iexplore.exe': 'Internet Explorer',
            'safari.exe': 'Safari'
        }
        # ブラウザのタイトルパターン (正規表現)
        self.title_patterns = [
            r'^(.+) - (.+)$',  # 例: "Google - Chrome"
            r'^(.+) — (.+)$',  # 例: "Google — Firefox"
            r'^(.+) – (.+)$',  # 例: "Google – Microsoft Edge"
        ]
        # 「新しいタブ」を示すタイトル (各ブラウザで異なる)
        self.new_tab_titles = {
            'Chrome': ['New Tab', '新しいタブ'],
            'Firefox': ['New Tab', '新しいタブ'],
            'Edge': ['New tab', '新しいタブ'],
            'Opera': ['Speed Dial', 'スピードダイヤル'],
            'Brave': ['New Tab', '新しいタブ'],
            'Internet Explorer': ['New Tab', '新しいタブ'],
            'Safari': ['Start Page', 'スタートページ']
        }
        
    def is_target_window(self, window: int) -> bool:
        """Webブラウザウィンドウかどうかを判定"""
        try:
            # ウィンドウのプロセスIDを取得
            _, pid = win32process.GetWindowThreadProcessId(window)
            process = psutil.Process(pid)
            process_name = process.name().lower()
            
            # ブラウザプロセスリストに含まれるか確認
            return process_name in self.browser_processes
        except Exception as e:
            print(f"Error in Browser is_target_window: {e}")
            return False
    
    def get_active_window_info(self) -> Optional[WindowInfo]:
        """アクティブなブラウザウィンドウの情報を取得"""
        try:
            window = win32gui.GetForegroundWindow()
            if not self.is_target_window(window):
                return None
                
            window_title = win32gui.GetWindowText(window)
            if window_title == self.last_title:
                return None
                
            # プロセス情報取得
            _, pid = win32process.GetWindowThreadProcessId(window)
            process = psutil.Process(pid)
            process_name = process.name()
            application_path = process.exe()
            
            # ブラウザタイプを特定
            browser_type = self.browser_processes.get(process_name, 'Unknown')
            
            # ページタイトルとURLを抽出 (タイトルから推定)
            page_title, browser_name = self._parse_browser_title(window_title, browser_type)
            
            # 新しいタブかどうかを判定
            is_new_tab = self._is_new_tab(page_title, browser_type)
            
            self.last_title = window_title
            
            return WindowInfo.create(
                process_name=process_name,
                window_title=window_title,
                process_id=pid,
                application_name=process_name,
                application_path=application_path,
                working_directory=page_title,  # URLの代わりにページタイトルを使用
                monitor_type='browser',
                is_new_document=is_new_tab,
                office_app_type=browser_type  # ブラウザの種類
            )
            
        except Exception as e:
            print(f"Error in Browser get_active_window_info: {e}")
            return None
    
    def _parse_browser_title(self, window_title: str, browser_type: str) -> Tuple[str, str]:
        """ブラウザのウィンドウタイトルからページタイトルとブラウザ名を抽出"""
        # 各パターンを試す
        for pattern in self.title_patterns:
            match = re.match(pattern, window_title)
            if match:
                # グループ1がページタイトル、グループ2がブラウザ名と仮定
                return match.group(1), match.group(2)
        
        # パターンにマッチしない場合はそのままタイトルを返す
        return window_title, browser_type
    
    def _is_new_tab(self, page_title: str, browser_type: str) -> bool:
        """新しいタブかどうかを判定"""
        # ブラウザごとの「新しいタブ」タイトルをチェック
        new_tab_titles = self.new_tab_titles.get(browser_type, [])
        return page_title in new_tab_titles or not page_title

