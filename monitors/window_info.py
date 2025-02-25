# window_info.py
from dataclasses import dataclass
from datetime import datetime
from typing import Optional

@dataclass
class WindowInfo:
    timestamp: str
    process_name: str
    window_title: str
    process_id: int
    application_name: str  # 実行ファイル名 (explorer.exe, excel.exe など)
    application_path: str  # 実行ファイルの完全パス
    working_directory: str # 作業ディレクトリパス (Explorerの場合は現在のフォルダ、Officeの場合はドキュメントの保存場所)
    monitor_type: str     # モニタータイプ ('general', 'explorer', 'office')
    is_new_document: bool = False
    office_app_type: Optional[str] = None  # 'Word', 'Excel', 'PowerPoint'

    @classmethod
    def create(cls, **kwargs):
        # タイムスタンプが指定されていない場合は現在時刻を使用
        if 'timestamp' not in kwargs:
            kwargs['timestamp'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        return cls(**kwargs)