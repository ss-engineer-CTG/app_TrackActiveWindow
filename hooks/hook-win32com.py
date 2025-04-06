# hooks/hook-win32com.py
from PyInstaller.utils.hooks import collect_submodules

# win32comの全サブモジュールを収集
hiddenimports = collect_submodules('win32com')

# Office関連の重要なモジュールを追加
hiddenimports += [
    'win32com.client.dynamic',
    'win32com.client.gencache',
    'win32com.client.makepy',
    'win32com.shell',
    'win32com.adsi',
    'win32com.axscript',
    'win32com.mapi',
    'win32com.ifilter',
]