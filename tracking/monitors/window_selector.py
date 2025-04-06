# window_selector.py
from typing import Dict, Optional, List
import win32gui
import logging
from .base.base_monitor import BaseWindowMonitor
from ..models.window_info import WindowInfo

class WindowSelector:
    def __init__(self):
        self.monitors: Dict[str, BaseWindowMonitor] = {}
        self.monitor_order: List[str] = []
        # モニタータイプごとのエラー発生カウント
        self._error_count: Dict[str, int] = {}
        # エラーしきい値（この回数以上エラーが発生したモニターは一時的に無効に）
        self._error_threshold = 5
        # エラーリセット間隔（Windows起動時からのミリ秒）
        self._last_error_reset = 0
        self._error_reset_interval = 60000  # 1分ごとにリセット

    def register_monitor(self, window_class: str, monitor: BaseWindowMonitor) -> None:
        self.monitors[window_class] = monitor
        self._error_count[window_class] = 0
        
        # 優先順位に基づいてモニターを追加
        # Explorer -> Excel -> Word -> PowerPoint -> Browser -> PDF -> その他
        if window_class == 'explorer':
            self._insert_at_position(window_class, 0)
        elif window_class == 'excel': 
            self._insert_after(window_class, 'explorer')
        elif window_class == 'word':
            self._insert_after(window_class, 'excel')
        elif window_class == 'powerpoint':
            self._insert_after(window_class, 'word')
        elif window_class == 'browser':  # 新規追加：ブラウザモニターの優先順位設定
            self._insert_after(window_class, 'powerpoint')
        elif window_class == 'pdf':  # PDFの優先順位をブラウザの後に変更
            self._insert_after(window_class, 'browser')
        else:
            self.monitor_order.append(window_class)

    def _insert_at_position(self, window_class: str, position: int) -> None:
        """特定の位置にモニタークラスを挿入"""
        if window_class in self.monitor_order:
            self.monitor_order.remove(window_class)
        
        # 位置が範囲外の場合は末尾に追加
        if position >= len(self.monitor_order):
            self.monitor_order.append(window_class)
        else:
            self.monitor_order.insert(position, window_class)

    def _insert_after(self, window_class: str, after_class: str) -> None:
        """特定のクラスの後にモニタークラスを挿入"""
        if window_class in self.monitor_order:
            self.monitor_order.remove(window_class)
            
        if after_class in self.monitor_order:
            position = self.monitor_order.index(after_class) + 1
            self._insert_at_position(window_class, position)
        else:
            # 指定されたクラスが存在しない場合は先頭に追加
            self._insert_at_position(window_class, 0)

    def _should_skip_monitor(self, monitor_class: str) -> bool:
        """エラーが多発しているモニターをスキップすべきかどうか判定"""
        return self._error_count.get(monitor_class, 0) >= self._error_threshold

    def _reset_error_counts_if_needed(self) -> None:
        """一定時間ごとにエラーカウントをリセット"""
        import time
        current_time = int(time.time() * 1000)
        if current_time - self._last_error_reset > self._error_reset_interval:
            self._error_count = {key: 0 for key in self._error_count}
            self._last_error_reset = current_time

    def get_appropriate_monitor(self, window_handle: int) -> Optional[BaseWindowMonitor]:
        """適切なモニターを選択（改善版）"""
        try:
            self._reset_error_counts_if_needed()
            
            # Explorer関連の特別処理（高優先度）
            explorer_monitor = self.monitors.get('explorer')
            if explorer_monitor and not self._should_skip_monitor('explorer'):
                try:
                    if explorer_monitor.is_target_window(window_handle):
                        return explorer_monitor
                except Exception as e:
                    logging.error(f"Error in Explorer monitor check: {e}")
                    self._error_count['explorer'] = self._error_count.get('explorer', 0) + 1
            
            # 他のモニターを優先順位に従って試す
            for monitor_class in self.monitor_order:
                if monitor_class == 'explorer':
                    continue  # Explorerは上で既にチェック済み
                    
                if self._should_skip_monitor(monitor_class):
                    continue  # エラー多発モニターはスキップ
                
                monitor = self.monitors[monitor_class]
                try:
                    if monitor.is_target_window(window_handle):
                        return monitor
                except Exception as e:
                    logging.error(f"Error checking {monitor_class} monitor: {e}")
                    self._error_count[monitor_class] = self._error_count.get(monitor_class, 0) + 1
            
            # 最終手段として一般モニターを返す
            return self.monitors.get('default')
        except Exception as e:
            logging.error(f"Error in monitor selection: {str(e)}")
            return self.monitors.get('default')

    def get_window_info(self) -> Optional[WindowInfo]:
        """ウィンドウ情報を取得（改善版）"""
        try:
            active_window = win32gui.GetForegroundWindow()
            monitor = self.get_appropriate_monitor(active_window)
            if monitor:
                try:
                    info = monitor.get_active_window_info()
                    if info:
                        logging.debug(f"Selected monitor: {monitor.__class__.__name__}")
                        print(f"Selected monitor: {monitor.__class__.__name__}")
                    return info
                except Exception as e:
                    monitor_class = monitor.__class__.__name__
                    logging.error(f"Error getting info from {monitor_class}: {e}")
                    self._error_count[monitor_class] = self._error_count.get(monitor_class, 0) + 1
                    
                    # エラー発生時は一般モニターで代替
                    default_monitor = self.monitors.get('default')
                    if default_monitor and default_monitor != monitor:
                        try:
                            return default_monitor.get_active_window_info()
                        except:
                            pass
        except Exception as e:
            logging.error(f"Error in window selection: {str(e)}")
        return None