# cache_manager.py
import time
from collections import OrderedDict
from typing import Dict, Any, Tuple, Optional, TypeVar, Generic

T = TypeVar('T')  # キャッシュされる値の型

class CacheManager(Generic[T]):
    """複数のモニター間で共有されるLRUキャッシュマネージャー"""
    
    def __init__(self, capacity: int = 100, timeout: int = 5):
        """
        Parameters:
            capacity (int): キャッシュの最大エントリ数
            timeout (int): キャッシュエントリの有効期間（秒）
        """
        self.capacity = capacity
        self.timeout = timeout
        self.cache: OrderedDict[int, Tuple[float, T]] = OrderedDict()
        
    def get(self, key: int) -> Optional[T]:
        """キーに対応する値を取得（存在しないか期限切れならNone）"""
        if key not in self.cache:
            return None
            
        timestamp, value = self.cache[key]
        
        # 期限切れチェック
        if time.time() - timestamp > self.timeout:
            del self.cache[key]
            return None
            
        # LRU順序を更新
        self.cache.move_to_end(key)
        return value
        
    def set(self, key: int, value: T) -> None:
        """キーと値をキャッシュに設定"""
        # キャッシュ容量チェック
        if len(self.cache) >= self.capacity and key not in self.cache:
            # 最も古いエントリを削除
            self.cache.popitem(last=False)
            
        # 値を設定
        self.cache[key] = (time.time(), value)
        
    def clear(self) -> None:
        """キャッシュをクリア"""
        self.cache.clear()
        
    def cleanup(self) -> None:
        """期限切れエントリを削除"""
        expired_keys = []
        current_time = time.time()
        
        for key, (timestamp, _) in self.cache.items():
            if current_time - timestamp > self.timeout:
                expired_keys.append(key)
                
        for key in expired_keys:
            del self.cache[key]