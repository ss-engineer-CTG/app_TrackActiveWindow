# hooks/runtime_hook.py
import os
import sys
import tempfile

# アプリケーションの実行時ディレクトリを設定
if getattr(sys, 'frozen', False):
    # パッケージ化された環境での一時ディレクトリ設定
    app_dir = os.path.dirname(sys.executable)
    user_docs = os.path.join(os.path.expanduser('~'), 'Documents', 'TrackActiveWindow')
    
    # 一時ディレクトリが存在することを確認
    temp_dir = os.path.join(user_docs, 'temp')
    if not os.path.exists(temp_dir):
        os.makedirs(temp_dir)
    
    # 一時ファイル用ディレクトリを設定
    tempfile.tempdir = temp_dir
    
    # COMオブジェクト関連の設定
    os.environ['PYTHONCOM_SKIPREGISTRATION'] = '1'