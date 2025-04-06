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
        
        # Office特有のファイル拡張子
        self.file_extensions = {
            'Excel': ['.xlsx', '.xls', '.xlsm', '.xlsb', '.csv'],
            'Word': ['.docx', '.doc', '.docm', '.rtf'],
            'PowerPoint': ['.pptx', '.ppt', '.pptm']
        }
        
    def _get_com_object(self) -> Optional[Any]:
        """より堅牢なCOMオブジェクト取得"""
        current_time = time.time()
        
        with self._com_lock:
            # 既存のCOMオブジェクトが有効かチェック
            if self._com_object is not None:
                self._last_access = current_time
                return self._com_object
                
            # 新しいCOMオブジェクトを作成
            try:
                pythoncom.CoInitialize()
                
                # 方法1: Dispatchを試す（新しいインスタンスの作成）
                try:
                    self._com_object = win32com.client.Dispatch(self.app_name)
                    self._last_access = current_time
                    logging.debug(f"Successfully created COM object via Dispatch for {self.app_name}")
                    return self._com_object
                except Exception as dispatch_e:
                    logging.debug(f"Dispatch failed for {self.app_name}: {dispatch_e}")
                    
                    # 方法2: GetObjectを試す（既存インスタンスへの接続）
                    try:
                        self._com_object = win32com.client.GetObject(None, self.app_name)
                        self._last_access = current_time
                        logging.debug(f"Successfully got COM object via GetObject for {self.app_name}")
                        return self._com_object
                    except Exception as getobj_e:
                        logging.error(f"GetObject also failed for {self.app_name}: {getobj_e}")
                        return None
                        
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
        """対象のOfficeウィンドウかどうかを判定（改善版）"""
        try:
            # プロセス名による基本チェック
            _, pid = win32process.GetWindowThreadProcessId(window)
            process = psutil.Process(pid)
            process_name_lower = process.name().lower()
            
            if process_name_lower != self.process_name.lower():
                return False
                
            # ウィンドウタイトルのパターンチェック（追加）
            window_title = win32gui.GetWindowText(window)
            
            # Officeアプリ固有のパターン
            if self.app_type == 'Excel' and (' - Excel' in window_title or any(ext in window_title.lower() for ext in self.file_extensions['Excel'])):
                return True
            elif self.app_type == 'Word' and (' - Word' in window_title or any(ext in window_title.lower() for ext in self.file_extensions['Word'])):
                return True
            elif self.app_type == 'PowerPoint' and (' - PowerPoint' in window_title or any(ext in window_title.lower() for ext in self.file_extensions['PowerPoint'])):
                return True
                
            # 通常のプロセス名チェックで真となった場合
            return True
                
        except Exception as e:
            logging.debug(f"Error in Office is_target_window: {e}")
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
    
    def _get_document_path_alternative(self, window: int, process: psutil.Process) -> Optional[str]:
        """代替手段でドキュメントのパスを取得する"""
        window_title = win32gui.GetWindowText(window)
        
        # このアプリタイプの拡張子を取得
        extensions = self.file_extensions.get(self.app_type, [])
        
        # 方法1: ウィンドウタイトルからの抽出
        # 一般的なパターン: "ファイル名.xxx - アプリ名"
        if " - " in window_title:
            # 例: "Document.docx - Word" から "Document.docx" を抽出
            doc_name = window_title.split(" - ")[0].strip()
            
            # 拡張子があるかチェック、なければ適切な拡張子を追加
            has_extension = any(doc_name.lower().endswith(ext) for ext in extensions)
            
            # よく使われるディレクトリでパスを探す
            search_dirs = [
                os.path.expanduser("~/Documents"),
                os.path.expanduser("~/Desktop"),
                os.path.expanduser("~/Downloads"),
            ]
            
            # プロセスの作業ディレクトリも追加
            try:
                cwd = process.cwd()
                if cwd:
                    search_dirs.append(cwd)
            except:
                pass
                
            # 拡張子がない場合は各拡張子を試す
            if not has_extension:
                for ext in extensions:
                    doc_name_with_ext = f"{doc_name}{ext}"
                    for directory in search_dirs:
                        potential_path = os.path.join(directory, doc_name_with_ext)
                        if os.path.exists(potential_path):
                            return potential_path
            else:
                # 拡張子がある場合は直接検索
                for directory in search_dirs:
                    potential_path = os.path.join(directory, doc_name)
                    if os.path.exists(potential_path):
                        return potential_path
        
        # 方法2: プロセスの開いているファイルを確認
        try:
            for file in process.open_files():
                file_path = file.path
                if any(file_path.lower().endswith(ext) for ext in extensions):
                    return file_path
        except Exception as e:
            logging.debug(f"Error checking open files: {e}")
        
        # 方法3: プロセスの作業ディレクトリを取得
        try:
            working_dir = process.cwd()
            if working_dir and working_dir != os.environ.get('WINDIR'):
                # 作業ディレクトリに関連する Office ファイルがあるか確認
                for file in os.listdir(working_dir):
                    if any(file.lower().endswith(ext) for ext in extensions):
                        file_path = os.path.join(working_dir, file)
                        # 最終更新時間でソートし、最新のものを返す
                        try:
                            return file_path
                        except:
                            pass
                return working_dir  # ファイルが見つからない場合は作業ディレクトリを返す
        except:
            pass
        
        return None
    
    def _get_application_document_info(self, window: int) -> Optional[WindowInfo]:
        """アクティブなドキュメントの情報を取得 (サブクラスで実装)"""
        raise NotImplementedError("Subclasses must implement _get_application_document_info()")
    
    def _create_basic_info(self, window: int, process: psutil.Process, document_path: str = '', is_new_document: bool = True) -> WindowInfo:
        """基本的なウィンドウ情報を作成（改善版）"""
        application_path = process.exe()
        window_title = win32gui.GetWindowText(window)
        
        # 新規文書の場合で document_path が空だった場合は特別な表記を使用
        if is_new_document and not document_path:
            document_path = f"<新規文書> - {self.app_type}"
        
        return WindowInfo.create(
            process_name=self.process_name,
            window_title=window_title,
            process_id=process.pid,
            application_name=self.process_name,
            application_path=application_path,
            working_directory=document_path,
            monitor_type='office',
            is_new_document=is_new_document,
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