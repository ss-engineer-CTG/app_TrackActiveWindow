# office_base_monitor.py
import win32com.client
import win32gui
import win32process
import pythoncom
import psutil
import os
import threading
import time
import logging
from typing import Optional, Dict, Any, Generic, TypeVar
from .base_monitor import BaseWindowMonitor
from ...models.window_info import WindowInfo
from ...utils.cache_manager import CacheManager

class OfficeBaseMonitor(BaseWindowMonitor):
    """すべてのOfficeモニターの基底クラス"""
    
    # 共有キャッシュマネージャー (すべてのOfficeモニターで共有)
    _shared_cache = CacheManager[WindowInfo](capacity=50, timeout=5)
    
    def __init__(self, app_name: str, process_name: str, app_type: str):
        """
        Parameters:
            app_name (str): COMオブジェクト名 ('Excel.Application' など)
            process_name (str): プロセス名 ('excel.exe' など)
            app_type (str): アプリケーションタイプ ('Excel', 'Word', 'PowerPoint')
        """
        super().__init__()
        self.app_name = app_name
        self.process_name = process_name
        self.app_type = app_type
        self._com_object = None
        self._com_lock = threading.Lock()
        self._last_access = 0
        self._timeout = 30  # 30秒間使用されなければCOMオブジェクトを解放
        
    def _get_com_object(self) -> Optional[Any]:
        """オンデマンドでCOMオブジェクトを取得"""
        current_time = time.time()
        
        with self._com_lock:
            # 既存のCOMオブジェクトが有効かチェック
            if self._com_object is not None:
                self._last_access = current_time
                return self._com_object
                
            # 新しいCOMオブジェクトを作成
            try:
                pythoncom.CoInitialize()
                self._com_object = win32com.client.GetObject(None, self.app_name)
                self._last_access = current_time
                return self._com_object
            except Exception as e:
                logging.error(f"COM initialization failed for {self.app_name}: {e}")
                return None
                
    def _release_com_object_if_idle(self) -> None:
        """一定時間アクセスがなければCOMオブジェクトを解放"""
        if self._com_object is None:
            return
            
        current_time = time.time()
        if current_time - self._last_access > self._timeout:
            with self._com_lock:
                if self._com_object is not None:
                    try:
                        self._com_object = None
                        pythoncom.CoUninitialize()
                        logging.debug(f"Released idle COM object for {self.app_name}")
                    except Exception as e:
                        logging.error(f"Error releasing COM object: {e}")
    
    def is_target_window(self, window: int) -> bool:
        """対象のOfficeウィンドウかどうかを判定"""
        try:
            _, pid = win32process.GetWindowThreadProcessId(window)
            process = psutil.Process(pid)
            return process.name().lower() == self.process_name.lower()
        except:
            return False
    
    def get_active_window_info(self) -> Optional[WindowInfo]:
        """アクティブウィンドウの情報を取得 (共通実装)"""
        try:
            window = win32gui.GetForegroundWindow()
            if not self.is_target_window(window):
                return None

            window_title = win32gui.GetWindowText(window)
            if window_title == self.last_title:
                return None
                
            # キャッシュをチェック
            cached_info = self._shared_cache.get(window)
            if cached_info:
                self.last_title = window_title
                return cached_info

            # キャッシュにない場合は情報を取得
            info = self._get_application_document_info(window)
            if info:
                self.last_title = window_title
                self._shared_cache.set(window, info)

            # アイドル状態のCOMオブジェクトを必要に応じて解放
            self._release_com_object_if_idle()
                
            return info

        except Exception as e:
            logging.error(f"Error in {self.app_type} get_active_window_info: {e}")
            return None
    
    def _get_application_document_info(self, window: int) -> Optional[WindowInfo]:
        """アクティブなドキュメントの情報を取得 (サブクラスで実装)"""
        raise NotImplementedError("Subclasses must implement _get_application_document_info()")
    
    def _create_basic_info(self, window: int, process: psutil.Process) -> WindowInfo:
        """基本的なウィンドウ情報を作成"""
        application_path = process.exe()
        return WindowInfo.create(
            process_name=self.process_name,
            window_title=win32gui.GetWindowText(window),
            process_id=process.pid,
            application_name=self.process_name,
            application_path=application_path,
            working_directory='',
            monitor_type='office',
            is_new_document=True,
            office_app_type=self.app_type
        )
        
    def __del__(self):
        """デストラクター - 解放されていないCOMオブジェクトをクリーンアップ"""
        if self._com_object is not None:
            try:
                self._com_object = None
                pythoncom.CoUninitialize()
            except:
                pass