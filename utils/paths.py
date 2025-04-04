# tracking/utils/paths.py
"""アプリケーションパス管理ユーティリティ"""
import os
import sys

def is_frozen():
    """PyInstallerでパッケージ化されているかどうかを判定"""
    return getattr(sys, 'frozen', False)

def get_app_dir():
    """アプリケーション実行ディレクトリを取得"""
    if is_frozen():
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

def get_user_data_dir():
    """ユーザーデータディレクトリを取得"""
    user_docs = os.path.join(os.path.expanduser('~'), 'Documents', 'TrackActiveWindow')
    ensure_dir_exists(user_docs)
    return user_docs

def get_logs_dir():
    """ログディレクトリを取得"""
    logs_dir = os.path.join(get_user_data_dir(), 'logs')
    ensure_dir_exists(logs_dir)
    return logs_dir

def get_temp_dir():
    """一時ディレクトリを取得"""
    temp_dir = os.path.join(get_user_data_dir(), 'temp')
    ensure_dir_exists(temp_dir)
    return temp_dir

def get_config_path():
    """設定ファイルのパスを取得"""
    config_dir = os.path.join(get_user_data_dir(), 'config')
    ensure_dir_exists(config_dir)
    return os.path.join(config_dir, 'tracking.ini')

def ensure_dir_exists(directory):
    """ディレクトリが存在することを保証"""
    if not os.path.exists(directory):
        os.makedirs(directory)