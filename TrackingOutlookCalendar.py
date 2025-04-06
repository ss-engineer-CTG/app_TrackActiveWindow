import win32com.client
import datetime
import os
from pathlib import Path
import sys

def export_outlook_calendar(export_path):
    """
    今日の予定をOutlookからエクスポートする関数
    
    Parameters:
    export_path (str): エクスポートしたファイルを保存するパス
    """
    try:
        # Outlookアプリケーションへの接続
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        # カレンダーフォルダーの取得
        calendar = namespace.GetDefaultFolder(9)  # 9はolFolderCalendarを表す
        print(f"カレンダーフォルダー: {calendar.Name}")  # デバッグ情報
        
        # 今日の日付の取得
        today = datetime.datetime.now()
        start_time = today.replace(hour=0, minute=0, second=0, microsecond=0)
        end_time = today.replace(hour=23, minute=59, second=59, microsecond=999999)
        
        print(f"検索期間: {start_time} から {end_time}")  # デバッグ情報
        
        # 指定した日付の予定を取得
        appointments = calendar.Items
        appointments.Sort("[Start]")
        appointments.IncludeRecurrences = True
        
        # 日付で予定をフィルタリング
        # フィルター文字列の修正
        filter_string = f"[Start] >= '{start_time.strftime('%m/%d/%Y 00:00 AM')}' AND [Start] <= '{end_time.strftime('%m/%d/%Y 11:59 PM')}'"
        print(f"フィルター文字列: {filter_string}")  # デバッグ情報
        
        filtered_items = appointments.Restrict(filter_string)
        
        # フィルタリング結果の確認
        item_count = 0
        print("\n本日の予定一覧:")  # デバッグ情報
        for item in filtered_items:
            print(f"予定: {item.Subject} ({item.Start} - {item.End})")  # デバッグ情報
            item_count += 1
        print(f"予定の総数: {item_count}")  # デバッグ情報
        
        # エクスポートファイル名の設定
        export_filename = f"calendar_export_{today.strftime('%Y%m%d')}.csv"
        full_export_path = Path(export_path) / export_filename
        
        # CSVファイルへの書き出し
        with open(full_export_path, 'w', encoding='utf-8-sig', newline='') as f:
            # ヘッダーの書き込み
            f.write("件名,開始時刻,終了時刻,場所,必須参加者,任意参加者,会議室,説明\n")
            
            # 予定の書き込み
            items_written = 0  # デバッグ用カウンター
            for item in filtered_items:
                try:
                    # 各フィールドのNull チェックと文字列処理
                    subject = cleanup_text(item.Subject) if hasattr(item, 'Subject') and item.Subject else ''
                    start = item.Start.strftime('%Y-%m-%d %H:%M') if hasattr(item, 'Start') and item.Start else ''
                    end = item.End.strftime('%Y-%m-%d %H:%M') if hasattr(item, 'End') and item.End else ''
                    location = cleanup_text(item.Location) if hasattr(item, 'Location') and item.Location else ''
                    
                    # 参加者情報の取得（エラーハンドリング付き）
                    required = get_attendees(item, 'Required')
                    optional = get_attendees(item, 'Optional')
                    
                    # 会議室情報
                    room = cleanup_text(item.Resources) if hasattr(item, 'Resources') and item.Resources else ''
                    
                    # 本文（説明）の処理
                    body = cleanup_text(item.Body) if hasattr(item, 'Body') and item.Body else ''
                    
                    # CSVの1行を作成
                    line = f"{subject},{start},{end},{location},{required},{optional},{room},{body}\n"
                    f.write(line)
                    items_written += 1
                    print(f"書き込み完了: {subject}")  # デバッグ情報
                    
                except Exception as e:
                    print(f"予定の処理中にエラーが発生しました: {str(e)}", file=sys.stderr)
                    continue
        
        print(f"予定表を {full_export_path} にエクスポートしました。")
        print(f"書き込まれた予定の数: {items_written}")  # デバッグ情報
        return True
        
    except Exception as e:
        print(f"エラーが発生しました: {str(e)}", file=sys.stderr)
        return False

def cleanup_text(text):
    """
    テキストのクリーンアップ処理
    """
    if not text:
        return ''
    return str(text).replace(',', '，').replace('\n', ' ').strip()

def get_attendees(item, attendee_type='Required'):
    """
    参加者リストの取得
    """
    try:
        if attendee_type == 'Required' and hasattr(item, 'RequiredAttendees'):
            attendees = item.RequiredAttendees
        elif attendee_type == 'Optional' and hasattr(item, 'OptionalAttendees'):
            attendees = item.OptionalAttendees
        else:
            return ''
        
        if not attendees:
            return ''
            
        names = [name.strip() for name in attendees.split(';') if name.strip()]
        return ';'.join(names)
        
    except Exception:
        return ''

if __name__ == "__main__":
    # エクスポート先のパスを指定（ユーザー環境に合わせて変更）
    export_folder = r"C:\Users\YourUsername\Documents\Calendar_Exports"
    
    # フォルダが存在しない場合は作成
    Path(export_folder).mkdir(parents=True, exist_ok=True)
    
    # エクスポート実行
    if export_outlook_calendar(export_folder):
        print("エクスポートが正常に完了しました。")
    else:
        print("エクスポート中にエラーが発生しました。", file=sys.stderr)