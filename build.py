#!/usr/bin/env python
# build.py - ビルドとパッケージングの自動化スクリプト
import os
import sys
import subprocess
import shutil
import argparse
from tracking.version import __version__, __app_name__

def run_command(command, cwd=None):
    """コマンドを実行し結果を表示"""
    print(f"実行コマンド: {command}")
    result = subprocess.run(command, shell=True, cwd=cwd)
    if result.returncode != 0:
        print(f"エラー: コマンド {command} が失敗しました")
        return False
    return True

def clean_build():
    """ビルドディレクトリをクリーンアップ"""
    dirs_to_clean = ['build', 'dist', '__pycache__']
    for dir_name in dirs_to_clean:
        if os.path.exists(dir_name):
            print(f"ディレクトリ {dir_name} を削除します")
            shutil.rmtree(dir_name)

def build_exe():
    """PyInstallerを使用して実行可能ファイルをビルド"""
    return run_command("pyinstaller TrackActiveWindow.spec")

def build_installer():
    """Inno Setupを使用してインストーラーをビルド"""
    # Inno Setupの正確なパスを指定
    iscc_path = r"C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
    
    if os.path.exists(iscc_path):
        # パスにスペースが含まれるため、引用符で囲む
        return run_command(f'"{iscc_path}" installer.iss')
    else:
        print(f"Inno Setup実行ファイルが見つかりません: {iscc_path}")
        print("インストーラーは作成されません。")
        return False

def main():
    parser = argparse.ArgumentParser(description='TrackActiveWindow ビルドスクリプト')
    parser.add_argument('--clean', action='store_true', help='ビルド前にビルドディレクトリをクリーンアップ')
    parser.add_argument('--exe-only', action='store_true', help='実行可能ファイルのみビルド（インストーラーは作成しない）')
    args = parser.parse_args()

    print(f"{__app_name__} v{__version__} ビルドを開始します")

    if args.clean:
        clean_build()

    if build_exe():
        print(f"実行可能ファイルが dist\\{__app_name__} ディレクトリに作成されました")
        
        if not args.exe_only:
            if build_installer():
                print(f"インストーラーが dist\\installer ディレクトリに作成されました")
            else:
                print("インストーラーの作成に失敗しました。実行可能ファイルのみが利用可能です。")
    else:
        print("ビルドに失敗しました。")
        sys.exit(1)

if __name__ == "__main__":
    # アプリケーションのルートディレクトリをPYTHONPATHに追加
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    main()