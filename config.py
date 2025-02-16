# config.py
import configparser
import os

class Config:
    def __init__(self, config_path='tracking.ini'):
        self.config = configparser.ConfigParser()
        self.config_path = config_path
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
            'excluded_processes': 'explorer.exe,SystemSettings.exe'
        }

        self.config['Advanced'] = {
            'debug_mode': 'false',
            'max_title_length': '256',
            'sanitize_strings': 'true'  # 文字列のサニタイズを有効化
        }

        with open(self.config_path, 'w', encoding='utf-8') as configfile:
            self.config.write(configfile)

    def get_value(self, section, key):
        return self.config.get(section, key)

    def set_value(self, section, key, value):
        """設定値を更新する"""
        if not self.config.has_section(section):
            self.config.add_section(section)
        self.config.set(section, key, str(value))
        with open(self.config_path, 'w', encoding='utf-8') as configfile:
            self.config.write(configfile)

    def get_boolean(self, section, key):
        """ブール値として設定を取得する"""
        return self.config.getboolean(section, key)