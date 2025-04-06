#!/usr/bin/env python
# run.py - アプリケーションのエントリーポイント
import os
import sys
import argparse
from tracking.version import __version__

def parse_args():
    parser = argparse.ArgumentParser(description='Window Activity Tracker')
    parser.add_argument('--version', action='store_true', help='Show version and exit')
    return parser.parse_args()

def main():
    args = parse_args()
    if args.version:
        print(f"Window Activity Tracker v{__version__}")
        return
        
    # メインモジュールをインポートして実行
    from tracking import main as app_main
    app_main.main()

if __name__ == "__main__":
    # アプリケーションのルートディレクトリをPYTHONPATHに追加
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    main()