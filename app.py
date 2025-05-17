from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, TextMessage, TextSendMessage, ImageMessage
import os
from dotenv import load_dotenv
import uuid
import google.generativeai as genai
from PIL import Image
from google.oauth2 import service_account
from googleapiclient.discovery import build
import datetime
import re

# 環境変数の読み込み
load_dotenv()

app = Flask(__name__)

# Line Botの設定
line_bot_api = LineBotApi(os.getenv('LINE_CHANNEL_ACCESS_TOKEN'))
handler = WebhookHandler(os.getenv('LINE_CHANNEL_SECRET'))

# Gemini APIの設定
GOOGLE_API_KEY = os.getenv('GOOGLE_API_KEY')
genai.configure(api_key=GOOGLE_API_KEY)

# Google Sheets APIのスコープ
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# 画像を保存するディレクトリ
SAVE_DIR = 'images'
if not os.path.exists(SAVE_DIR):
    os.makedirs(SAVE_DIR)

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
    """
    try:
        # 画像を読み込む
        image = Image.open(image_path)
        
        # Geminiモデルの初期化
        model = genai.GenerativeModel('gemini-1.5-flash')
        
        # 画像からテキストを抽出
        response = model.generate_content(["この画像から文字を抽出してください。抽出した内容のみ出力してください。", image])
        
        return response.text
    
    except Exception as e:
        return f"エラーが発生しました: {str(e)}"

@app.route("/callback", methods=['POST'])
def callback():
    # リクエストヘッダーから署名検証のための値を取得
    signature = request.headers['X-Line-Signature']

    # リクエストボディを取得
    body = request.get_data(as_text=True)
    app.logger.info("Request body: " + body)

    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)

    return 'OK'

@handler.add(MessageEvent, message=ImageMessage)
def handle_image_message(event):
    try:
        # 画像メッセージのIDを取得
        message_id = event.message.id
        
        # 画像のバイナリデータを取得
        message_content = line_bot_api.get_message_content(message_id)
        
        # 一意のファイル名を生成
        file_name = f"{uuid.uuid4()}.jpg"
        file_path = os.path.join(SAVE_DIR, file_name)
        
        # 画像を保存
        with open(file_path, 'wb') as f:
            for chunk in message_content.iter_content():
                f.write(chunk)
        
        # 画像からテキストを抽出
        extracted_text = extract_text_from_image(file_path)
        
        # スプレッドシートに保存
        SPREADSHEET_ID = os.getenv('SPREADSHEET_ID')
        RANGE_NAME = 'シート1!A1'
        
        if SPREADSHEET_ID:
            save_to_spreadsheet(extracted_text, SPREADSHEET_ID, RANGE_NAME)
            # ユーザーに応答メッセージを送信
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(text=f"画像から文字を抽出し、スプレッドシートに保存しました。\n\n抽出した内容:\n{extracted_text}")
            )
        else:
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(text="スプレッドシートIDが設定されていません。")
            )
        
    except Exception as e:
        # エラーが発生した場合の処理
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text=f"エラーが発生しました: {str(e)}")
        )

if __name__ == "__main__":
    app.run() 