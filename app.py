import streamlit as st
import os
import tempfile
import google.generativeai as genai
from PIL import Image
import gspread
from google.oauth2.service_account import Credentials
from pydrive2.auth import GoogleAuth
from pydrive2.drive import GoogleDrive
from gspread.exceptions import WorksheetNotFound, SpreadsheetNotFound
import json # jsonモジュールはここでインポート

# --- パスワード認証 ---
PASSWORD = "xpost00"  # ←ここを好きなパスワードに変更

if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

if not st.session_state["authenticated"]:
    pw = st.text_input("パスワードを入力してください", type="password")
    if pw == PASSWORD:
        st.session_state["authenticated"] = True
        st.success("認証成功！")
        # 認証成功後、アプリを再実行してUIを表示
        st.rerun()
    else:
        st.error("パスワードが異なります。")
        st.stop() # 認証失敗時は処理を停止

# 認証成功後の処理（ここからメインアプリのコードが実行される）

# --- 設定 ---
# Gemini APIキーをここに設定してください
GENAI_API_KEY = "AIzaSyCc2MQQ2ytt32gzMq53L_Z8SKhWWMRjJ1s"

# Google SheetsとGoogle Driveのスコープ
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

# Google Sheetsのヘッダー定義
# ヘッダーに「アカウントID」を追加
headers = ["画像", "投稿内容", "発信者名", "アカウントID", "投稿時間", "いいね数", "RT数", "コメント数", "インプレッション", "ブックマーク数"]

# --- Google Sheets認証 ---
@st.cache_resource
def authenticate_gspread():
    """gspreadを認証し、認証オブジェクトをキャッシュする"""
    try:
        google_credentials = st.secrets["GOOGLE_CREDENTIALS"]
        if isinstance(google_credentials, str):
            cred_dict = json.loads(google_credentials)
        else:
            cred_dict = google_credentials
        creds = Credentials.from_service_account_info(cred_dict, scopes=SCOPES)
        gc = gspread.authorize(creds)
        st.success("Google Sheets認証に成功しました。")
        return gc
    except Exception as e:
        st.error(f"Google Sheets認証に失敗しました。認証情報を確認してください: {e}")
        st.stop()

from pydrive2.auth import GoogleAuth
from pydrive2.drive import GoogleDrive

# --- Google Drive認証 ---
pydrive_settings = {
    "client_config_backend": "service",
    "service_config": {
        "client_json": st.secrets["GOOGLE_CREDENTIALS"]
    }
}
gauth = GoogleAuth(settings=pydrive_settings)
gauth.ServiceAuth()
drive = GoogleDrive(gauth)

# --- Gemini API ---
@st.cache_resource
def configure_gemini():
    """Gemini APIを設定し、モデルをキャッシュする"""
    try:
        genai.configure(api_key=GENAI_API_KEY)
        model = genai.GenerativeModel("gemini-2.5-flash")
        st.success("Gemini API設定に成功しました。")
        return model
    except Exception as e:
        st.error(f"Gemini APIキーの設定に失敗しました。APIキーを確認してください: {e}")
        st.stop() # 設定失敗時は処理を停止

model = configure_gemini()

def upload_image_to_drive(image_path, drive_service):
    """
    画像をGoogle Driveにアップロードし、公開URLを返す。
    """
    try:
        file_name = os.path.basename(image_path)
        # ファイルを作成
        file = drive_service.CreateFile({'title': file_name})
        file.SetContentFile(image_path)
        file.Upload()
        # 誰でも閲覧できるように権限を設定
        file.InsertPermission({'type': 'anyone', 'value': 'anyone', 'role': 'reader'})
        st.write(f"画像 '{file_name}' をGoogle Driveにアップロードしました。")
        return f"https://drive.google.com/uc?id={file['id']}"
    except Exception as e:
        st.error(f"Google Driveへの画像アップロード中にエラーが発生しました: {e}")
        return None

import json
import os

def get_or_create_user_spreadsheet(gc, email, title_prefix="Xポスト一覧_"):
    db_file = "user_sheets.json"
    if os.path.exists(db_file):
        with open(db_file, "r") as f:
            user_sheets = json.load(f)
    else:
        user_sheets = {}

    if email in user_sheets:
        spreadsheet_id = user_sheets[email]
        sh = gc.open_by_key(spreadsheet_id)
    else:
        sh = gc.create(f"{title_prefix}{email}")
        spreadsheet_id = sh.id
        user_sheets[email] = spreadsheet_id
        with open(db_file, "w") as f:
            json.dump(user_sheets, f)
        # メールアドレスに編集者権限を付与
        sh.share(email, perm_type='user', role='writer')
    return sh

def extract_post_info(image_path, gemini_model):
    """
    Gemini APIを使用して画像から投稿情報を抽出する。
    プロンプトを修正し、「アカウントID」も抽出するように変更。
    """
    try:
        image_data = Image.open(image_path)
        prompt = """
        この画像はX（旧Twitter）のポストです。
        投稿内容、発信者名、アカウントID、投稿時間、いいね数、RT数、コメント数、インプレッション、ブックマーク数を日本語で表にしてください。
        発信者名とアカウントIDは必ず分けて出力してください。
        投稿時間は「2025年7月3日　午後11:41」のように、日付→時刻の順で出力してください。
        例:
        | 投稿内容 | 発信者名 | アカウントID | 投稿時間 | いいね数 | RT数 | コメント数 | インプレッション | ブックマーク数 |
        | 例の投稿内容 | なまいきくん | 1namaiki | 2025年7月3日 午後11:41 | 100 | 10 | 5 | 1万 | 20 |
        """
        response = gemini_model.generate_content([prompt, image_data])
        return response.text
    except Exception as e:
        st.error(f"Gemini APIでの情報抽出中にエラーが発生しました: {e}")
        return None

def parse_table(text):
    """
    Geminiからのテキスト結果をパースして辞書形式で返す。
    """
    if not text:
        return None
    lines = [l for l in text.splitlines() if "|" in l]
    if len(lines) < 2:
        return None
    data_lines = []
    for l in lines:
        cells = [c.strip() for c in l.split("|")[1:-1]]
        # Markdownの区切り行をスキップ
        if all(cell.startswith(":") or set(cell) <= set("-:") for cell in cells):
            continue
        data_lines.append(l)
    if len(data_lines) < 2:
        return None
    headers_row = [h.strip() for h in data_lines[0].split("|")[1:-1]]
    values_row = [v.strip() for v in data_lines[1].split("|")[1:-1]]
    # ヘッダーと値の数が一致しない場合はNoneを返す
    if len(headers_row) != len(values_row):
        st.warning("Geminiの出力形式が予期せぬものでした。")
        return None
    return dict(zip(headers_row, values_row))

def get_or_create_spreadsheet(gspread_client, drive_service, user_name):
    """
    ユーザー名に対応するスプレッドシートを取得または新規作成する。
    """
    spreadsheet_title = f"Xポスト自動化_{user_name}"
    try:
        # 既存のスプレッドシートをタイトルで検索
        spreadsheet = gspread_client.open(spreadsheet_title)
        st.write(f"既存のスプレッドシート '{spreadsheet_title}' を使用します。")
        return spreadsheet
    except SpreadsheetNotFound:
        # スプレッドシートが存在しない場合、新規作成
        st.write(f"スプレッドシート '{spreadsheet_title}' を新規作成します。")
        spreadsheet = gspread_client.create(spreadsheet_title)
        # 作成したスプレッドシートを誰でも閲覧できるように共有設定
        spreadsheet.share('', perm_type='anyone', role='reader')
        st.success(f"新しいスプレッドシート '{spreadsheet_title}' を作成しました。")
        # デフォルトで作成される'Sheet1'を削除し、最初のワークシートを適切に管理
        # gspread 5.0以降では、create時にデフォルトで1つワークシートが作成される
        # 必要に応じて、既存の'Sheet1'を削除するロジックを追加することも可能だが、
        # 今回は発信者ごとのタブを作成するため、そのままにしておくか、
        # 後続のget_or_create_worksheetで最初のタブを適切に扱う。
        return spreadsheet
    except Exception as e:
        st.error(f"スプレッドシートの取得または作成中にエラーが発生しました: {e}")
        return None

def get_or_create_spreadsheet(gspread_client, drive_service, user_email):
    """
    ユーザーのメールアドレスに対応するスプレッドシートを取得または新規作成する。
    新規作成時には、指定されたメールアドレスに編集権限を付与する。
    """
    spreadsheet_title = f"Xポスト自動化_{user_email}"
    try:
        # 既存のスプレッドシートをタイトルで検索
        spreadsheet = gspread_client.open(spreadsheet_title)
        st.write(f"既存のスプレッドシート '{spreadsheet_title}' を使用します。")
        return spreadsheet
    except SpreadsheetNotFound:
        # スプレッドシートが存在しない場合、新規作成
        st.write(f"スプレッドシート '{spreadsheet_title}' を新規作成します。")
        spreadsheet = gspread_client.create(spreadsheet_title)
        
        # 作成したスプレッドシートに、指定されたメールアドレスに編集権限を付与
        try:
            spreadsheet.share(user_email, perm_type='user', role='writer')
            st.success(f"新しいスプレッドシート '{spreadsheet_title}' を作成し、{user_email} に編集権限を付与しました。")
        except Exception as share_e:
            st.warning(f"スプレッドシートの共有設定中にエラーが発生しました。手動で共有設定を行ってください: {share_e}")
            st.success(f"新しいスプレッドシート '{spreadsheet_title}' を作成しました。")

        # デフォルトで作成される'Sheet1'を削除し、最初のワークシートを適切に管理
        # gspread 5.0以降では、create時にデフォルトで1つワークシートが作成される
        # 今回は発信者ごとのタブを作成するため、そのままにしておくか、
        # 後続のget_or_create_worksheetで最初のタブを適切に扱う。
        return spreadsheet
    except Exception as e:
        st.error(f"スプレッドシートの取得または作成中にエラーが発生しました: {e}")
        return None

def get_or_create_spreadsheet(gspread_client, drive_service, user_email):
    """
    ユーザーのメールアドレスに対応するスプレッドシートを取得または新規作成する。
    新規作成時には、指定されたメールアドレスに編集権限を付与する。
    """
    spreadsheet_title = f"Xポスト自動化_{user_email}"
    spreadsheet = None # 初期化
    try:
        # 既存のスプレッドシートをタイトルで検索
        spreadsheet = gspread_client.open(spreadsheet_title)
        st.write(f"既存のスプレッドシート '{spreadsheet_title}' を使用します。")
    except SpreadsheetNotFound:
        # スプレッドシートが存在しない場合、新規作成
        st.write(f"スプレッドシート '{spreadsheet_title}' を新規作成します。")
        spreadsheet = gspread_client.create(spreadsheet_title)
        st.success(f"新しいスプレッドシート '{spreadsheet_title}' を作成しました。")
    except Exception as e:
        st.error(f"スプレッドシートの取得または作成中にエラーが発生しました: {e}")
        return None

    # スプレッドシートが取得または作成された場合のみ共有設定を試みる
    if spreadsheet:
        try:
            # 指定されたメールアドレスに編集権限を付与
            # 既に権限がある場合は更新される
            spreadsheet.share(user_email, perm_type='user', role='writer')
            st.success(f"スプレッドシート '{spreadsheet_title}' に {user_email} の編集権限を設定しました。")
        except Exception as share_e:
            st.warning(f"スプレッドシートの共有設定中にエラーが発生しました。手動で共有設定を行ってください: {share_e}")
    
    return spreadsheet

def get_or_create_worksheet(spreadsheet, sheet_title, headers_list):
    """
    指定されたスプレッドシート内で、指定されたタイトル（発信者名）のワークシートを取得または新規作成する。
    新規作成時にはヘッダーを書き込む。
    """
    try:
        # 既存のワークシートをタイトルで取得
        worksheet = spreadsheet.worksheet(sheet_title)
        st.write(f"既存のワークシート '{sheet_title}' を使用します。")
        return worksheet
    except WorksheetNotFound:
        # ワークシートが存在しない場合、新規作成
        st.write(f"ワークシート '{sheet_title}' を新規作成します。")
        # add_worksheetのrows/colsは目安。必要に応じて調整。
        worksheet = spreadsheet.add_worksheet(title=sheet_title, rows="1000", cols="20")
        # ヘッダーを書き込む
        worksheet.append_row(headers_list)
        return worksheet
    except Exception as e:
        st.error(f"ワークシートの取得または作成中にエラーが発生しました: {e}")
        return None

gc = authenticate_gspread()
drive = authenticate_pydrive()
model = configure_gemini()

# --- Streamlit UI ---
st.title("Xポスト画像→スプレッドシート自動化アプリ")
st.write("画像をアップロードすると、内容を自動で抽出してあなた専用のGoogleスプレッドシートの、投稿者ごとのタブに追記します。")

# ユーザーのGoogleメールアドレス入力
email = st.text_input("あなたのGoogleメールアドレスを入力してください")

uploaded_file = st.file_uploader("画像をアップロードしてください（PNG/JPG）", type=["png", "jpg", "jpeg"])

# メールアドレスとファイルが両方入力された場合のみ処理を開始
if email and uploaded_file is not None:
    # st.write("ファイルがアップロードされました") # 処理メッセージを削除
    
    # 一時ファイルの作成
    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_file:
            tmp_file.write(uploaded_file.read())
            tmp_path = tmp_file.name

        st.image(tmp_path, caption="アップロード画像", use_column_width=True)
        st.info("画像を解析中...") # ユーザーへの状態表示は残す

        # ユーザー専用のスプレッドシートを取得または作成
        # st.write(f"'{email}'さんのスプレッドシートを取得/作成します。") # 処理メッセージを削除
        user_spreadsheet = get_or_create_spreadsheet(gc, drive, email) # emailを渡す
        if user_spreadsheet is None:
            st.error("スプレッドシートの準備に失敗しました。")
            # エラー発生時は後続処理を行わない
            if tmp_path and os.path.exists(tmp_path):
                os.remove(tmp_path)
                # st.write("一時ファイル削除完了") # 処理メッセージを削除
            st.stop()

        # 画像をGoogleドライブにアップロード
        # st.write("Googleドライブにアップロード開始") # 処理メッセージを削除
        image_url = upload_image_to_drive(tmp_path, drive)
        if image_url is None:
            st.error("Google Driveへの画像アップロードに失敗しました。")
            if tmp_path and os.path.exists(tmp_path):
                os.remove(tmp_path)
                # st.write("一時ファイル削除完了") # 処理メッセージを削除
            st.stop()
        
        # st.write(f"image_url: {image_url}") # 処理メッセージを削除
        image_formula = f'=IMAGE("{image_url}", 2)'  # 元サイズで表示

        # Geminiで情報抽出
        # st.write("Geminiで情報抽出開始") # 処理メッセージを削除
        result_text = extract_post_info(tmp_path, model)
        if result_text is None:
            st.error("Geminiでの情報抽出に失敗しました。")
            if tmp_path and os.path.exists(tmp_path):
                os.remove(tmp_path)
                # st.write("一時ファイル削除完了") # 処理メッセージを削除
            st.stop()

        st.text_area("Gemini抽出結果", result_text, height=200) # ユーザーが結果を確認できるよう残す

        info = parse_table(result_text)
        # st.write(f"parse_tableの結果: {info}") # 処理メッセージを削除
        
        if info:
            # 発信者名とアカウントIDを取得
            author_name = info.get("発信者名")
            account_id = info.get("アカウントID")

            if not author_name:
                st.error("発信者名情報を抽出できませんでした。Geminiの出力形式を確認してください。")
                if tmp_path and os.path.exists(tmp_path):
                    os.remove(tmp_path)
                    # st.write("一時ファイル削除完了") # 処理メッセージを削除
                st.stop()

            # タブ名を「発信者名（@アカウントID）」の形式で生成
            tab_name = f"{author_name}（@{account_id}）" if account_id else author_name
            
            # 発信者ごとのワークシートを取得または作成
            # st.write(f"'{tab_name}'さんのタブを取得/作成します。") # 処理メッセージを削除
            target_worksheet = get_or_create_worksheet(user_spreadsheet, tab_name, headers)
            if target_worksheet is None:
                st.error("ワークシートの準備に失敗しました。")
                if tmp_path and os.path.exists(tmp_path):
                    os.remove(tmp_path)
                    # st.write("一時ファイル削除完了") # 処理メッセージを削除
                st.stop()

            # データを追記
            # headersの順番に合わせてデータを準備
            row_data = [image_formula, info.get("投稿内容", ""), info.get("発信者名", ""), info.get("アカウントID", ""), 
                        info.get("投稿時間", ""), info.get("いいね数", ""), info.get("RT数", ""), 
                        info.get("コメント数", ""), info.get("インプレッション", ""), info.get("ブックマーク数", "")]
            try:
                target_worksheet.append_row(row_data, value_input_option='USER_ENTERED')
                st.success(f"スプレッドシート '{user_spreadsheet.title}' の '{tab_name}' タブに追記しました！")
                st.markdown(f"[スプレッドシートを開く]({user_spreadsheet.url})")
            except Exception as e:
                st.error(f"スプレッドシートへの追記中にエラーが発生しました: {e}")
        else:
            st.error("情報の抽出に失敗しました。Geminiの出力形式を確認してください。")

    except Exception as e:
        st.error(f"予期せぬエラーが発生しました: {e}")
    finally:
        # 一時ファイル削除
        if tmp_path and os.path.exists(tmp_path):
            os.remove(tmp_path)
            # st.write("一時ファイル削除完了") # 処理メッセージを削除
        # else: # 処理メッセージを削除
            # st.write("一時ファイルは作成されなかったか、既に削除されています。")

elif uploaded_file is not None and not email:
    st.warning("画像をアップロードする前に、あなたのGoogleメールアドレスを入力してください。")
elif email and uploaded_file is None:
    st.info("画像をアップロードしてください。")

    st.info("画像をアップロードしてください。")
