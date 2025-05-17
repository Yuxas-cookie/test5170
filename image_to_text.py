import os
from dotenv import load_dotenv
import google.generativeai as genai
from PIL import Image
from google.oauth2 import service_account
from googleapiclient.discovery import build
import datetime
import re

# .envファイルから環境変数を読み込む
load_dotenv()

# Gemini APIの設定
GOOGLE_API_KEY = os.getenv('GOOGLE_API_KEY')
genai.configure(api_key=GOOGLE_API_KEY)

# Google Sheets APIのスコープ
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

def format_text_to_table(text):
    """
    抽出したテキストを表形式に変換する
    
    Args:
        text (str): 抽出したテキスト
        
    Returns:
        list: 表形式のデータ（2次元リスト）
    """
    try:
        # 改行で分割
        lines = text.split('\n')
        
        # 空行を削除
        lines = [line.strip() for line in lines if line.strip()]
        
        # 各行をタブまたはスペースで分割
        table_data = []
        for line in lines:
            # タブまたは複数のスペースで分割
            columns = re.split(r'\t|\s{2,}', line)
            # 空の要素を削除
            columns = [col.strip() for col in columns if col.strip()]
            if columns:
                table_data.append(columns)
        
        return table_data
    
    except Exception as e:
        print(f"テキストの表形式変換中にエラーが発生しました: {str(e)}")
        return [[text]]  # 変換できない場合は元のテキストをそのまま返す

def get_google_sheets_service():
    """
    Google Sheets APIのサービスアカウント認証を行い、サービスオブジェクトを返す
    
    Returns:
        googleapiclient.discovery.Resource: Google Sheets APIのサービスオブジェクト
    """
    try:
        # サービスアカウントの認証情報を読み込む
        credentials = service_account.Credentials.from_service_account_file(
            'credentials.json', scopes=SCOPES)
        
        return build('sheets', 'v4', credentials=credentials)
    
    except Exception as e:
        print(f"認証エラーが発生しました: {str(e)}")
        return None

def save_to_spreadsheet(text, spreadsheet_id, range_name):
    """
    抽出したテキストをスプレッドシートに保存する
    
    Args:
        text (str): 保存するテキスト
        spreadsheet_id (str): スプレッドシートのID
        range_name (str): 書き込む範囲（例: 'シート1!A1'）
    """
    try:
        service = get_google_sheets_service()
        if not service:
            return
        
        # テキストを表形式に変換
        table_data = format_text_to_table(text)
        
        # 現在の日時を取得
        current_time = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        
        # ヘッダー行を追加
        header = ['日時'] + [f'列{i+1}' for i in range(len(table_data[0]) if table_data else 0)]
        values = [header]
        
        # データ行を追加
        for row in table_data:
            values.append([current_time] + row)
        
        body = {
            'values': values
        }
        
        # スプレッドシートにデータを追加
        result = service.spreadsheets().values().append(
            spreadsheetId=spreadsheet_id,
            range=range_name,
            valueInputOption='RAW',
            body=body
        ).execute()
        
        print(f"{result.get('updates').get('updatedCells')} セルが更新されました")
        
    except Exception as e:
        print(f"スプレッドシートへの保存中にエラーが発生しました: {str(e)}")

def extract_text_from_image(image_path):
    """
    画像から文字を抽出する関数
    
    Args:
        image_path (str): 画像ファイルのパス
        
    Returns:
        str: 抽出されたテキスト
    """
    try:
        # 画像を読み込む
        image = Image.open(image_path)
        
        # Geminiモデルの初期化（新しいモデルを使用）
        model = genai.GenerativeModel('gemini-1.5-flash')
        
        # 画像からテキストを抽出
        response = model.generate_content(["この画像から文字を抽出してください。抽出した内容のみ出力してください。", image])
        
        return response.text
    
    except Exception as e:
        return f"エラーが発生しました: {str(e)}"

if __name__ == "__main__":
    # テスト用の画像パス
    image_path = "test.jpg"
    
    # テキスト抽出の実行
    extracted_text = extract_text_from_image(image_path)
    print("抽出されたテキスト:")
    print(extracted_text)
    
    # スプレッドシートに保存
    SPREADSHEET_ID = os.getenv('SPREADSHEET_ID')  # .envファイルからスプレッドシートIDを取得
    RANGE_NAME = 'シート1!A1'  # 開始位置を指定
    
    if SPREADSHEET_ID:
        save_to_spreadsheet(extracted_text, SPREADSHEET_ID, RANGE_NAME)
    else:
        print("スプレッドシートIDが設定されていません。") 