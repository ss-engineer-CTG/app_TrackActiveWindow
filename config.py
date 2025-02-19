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
            'excluded_processes': 'explorer.exe,SystemSettings.exe',
            'office_cache_timeout': '5',
            'office_retry_interval': '5'
        }
        with open(self.config_path, 'w', encoding='utf-8') as configfile:
            self.config.write(configfile)

    def get_value(self, section, key):
        return self.config.get(section, key)