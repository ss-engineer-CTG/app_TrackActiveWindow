# config.py
import configparser
import os
from .utils.paths import get_config_path, ensure_dir_exists

class Config:
    def __init__(self, config_path=None):
        self.config = configparser.ConfigParser()
        
        # 設定ファイルのパスを取得
        self.config_path = config_path if config_path else get_config_path()
        self.load_config()

    def load_config(self):
        if not os.path.exists(self.config_path):
            self.create_default_config()
        self.config.read(self.config_path, encoding='utf-8')

    def create_default_config(self):
        self.config['General'] = {
            'log_retention_days': '30',
            'buffer_size': '500',
            'write_interval': '3',
            'excluded_processes': 'explorer.exe,SystemSettings.exe',
            'office_cache_timeout': '5',
            'office_retry_interval': '5',
            'office_com_timeout': '30',  # COMオブジェクトのアイドルタイムアウト（秒）
            'cache_capacity': '50'       # キャッシュの最大容量
        }
        
        # ブラウザ設定セクションを追加
        self.config['Browser'] = {
            'capture_urls': 'false',      # URL取得機能（拡張機能連携時のみ有効）
            'enhanced_monitoring': 'false', # 拡張機能との連携
            'track_tab_changes': 'false',  # タブ切り替え監視機能
            'excluded_domains': 'example.com,internal.local'  # 監視対象外ドメイン
        }
        
        # 設定ファイルのディレクトリが存在することを確認
        config_dir = os.path.dirname(self.config_path)
        ensure_dir_exists(config_dir)
        
        with open(self.config_path, 'w', encoding='utf-8') as configfile:
            self.config.write(configfile)

    def get_value(self, section, key):
        return self.config.get(section, key)
        
    def set_value(self, section, key, value):
        """設定値を更新し、ファイルに保存する"""
        if not self.config.has_section(section):
            self.config.add_section(section)
            
        self.config.set(section, key, value)
        
        with open(self.config_path, 'w', encoding='utf-8') as configfile:
            self.config.write(configfile)