import win32com.client
import datetime
import os
def export_today_emails(target_folder):
    """
    今日付のメールをOutlookから指定フォルダーにエクスポートする
   
    Parameters:
    target_folder (str): エクスポート先のフォルダーパス
    """
    # Outlookアプリケーションへの接続
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
   
    # 受信トレイの取得
    inbox = namespace.GetDefaultFolder(6)  # 6は受信トレイを示す
   
    # 今日の日付を取得
    today = datetime.datetime.now().date()
   
    # エクスポート先フォルダーが存在しない場合は作成
    if not os.path.exists(target_folder):
        os.makedirs(target_folder)
   
    # メールの検索と保存
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)  # 受信日時でソート
   
    for message in messages:
        try:
            received_date = message.ReceivedTime.date()
           
            # 今日のメールのみを処理
            if received_date == today:
                # ファイル名の作成（重複を避けるため受信時刻を含める）
                safesubject = "".join(c for c in message.Subject if c.isalnum() or c in (' ', '-', ''))
                filename = f"{message.ReceivedTime.strftime('%Y%m%d%H%M%S')}{safe_subject[:50]}.msg"
                file_path = os.path.join(target_folder, filename)
               
                # メールの保存
                message.SaveAs(file_path)
                print(f"Saved: {filename}")
       
        except Exception as e:
            print(f"Error processing message: {e}")
   
    print("Export completed!")
# スクリプトの使用例
if name == "main":
    # エクスポート先フォルダーを指定
    export_folder = r"C:\Users\YourUsername\Documents\OutlookExports"
    export_today_emails(export_folder)