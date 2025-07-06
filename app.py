import streamlit as st
import os

PASSWORD = "xpost00"  # ←ここを好きなパスワードに変更

if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

if not st.session_state["authenticated"]:
    pw = st.text_input("パスワードを入力してください", type="password")
    if pw == PASSWORD:
        st.session_state["authenticated"] = True
        st.success("認証成功！")
    else:
        st.stop()

# --- Secretsから認証情報を取得 ---
GENAI_API_KEY = st.secrets["GENAI_API_KEY"]
GOOGLE_CREDENTIALS = st.secrets["GOOGLE_CREDENTIALS"]

# credentials.jsonを一時ファイルとして保存
with open("credentials.json", "w") as f:
    f.write(GOOGLE_CREDENTIALS)
import tempfile
import google.generativeai as genai
from PIL import Image
import gspread
from google.oauth2.service_account import Credentials
from pydrive2.auth import GoogleAuth
from pydrive2.drive import GoogleDrive

# --- 設定 ---
GENAI_API_KEY = "AIzaSyCc2MQQ2ytt32gzMq53L_Z8SKhWWMRjJ1s"  # ←ここを自分のGemini APIキーに書き換えてください
SPREADSHEET_ID = "1ZXxNhnvjix50IDimcdtPvOX8SZymJcDbQF_P3C3OqLw"  # ←自分のスプレッドシートID

# --- Gemini API ---
genai.configure(api_key=GENAI_API_KEY)
model = genai.GenerativeModel("gemini-2.5-flash")

# --- Google Sheets認証 ---
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]
# credentials.jsonはコードと同じディレクトリに配置してください
creds = Credentials.from_service_account_file('credentials.json', scopes=SCOPES)
gc = gspread.authorize(creds)

from pydrive2.auth import GoogleAuth
from pydrive2.drive import GoogleDrive

# --- Google Drive認証 ---
pydrive_settings = {
    "client_config_backend": "service",
    "service_config": {
        "client_json": st.secrets["GOOGLE_CREDENTIALS"]  # ←ここはJSON文字列
    }
}
gauth = GoogleAuth(settings=pydrive_settings)
gauth.ServiceAuth()
drive = GoogleDrive(gauth)

# --- スプレッドシート ---
headers = ["画像", "投稿内容", "発信者", "投稿時間", "いいね数", "RT数", "コメント数", "インプレッション", "ブックマーク数"]

def upload_image_to_drive(image_path):
    file = drive.CreateFile({'title': os.path.basename(image_path)})
    file.SetContentFile(image_path)
    file.Upload()
    file.InsertPermission({'type': 'anyone', 'value': 'anyone', 'role': 'reader'})
    return f"https://drive.google.com/uc?id={file['id']}"

def extract_post_info(image_path):
    image_data = Image.open(image_path)
    prompt = """
    この画像はX（旧Twitter）のポストです。
    投稿内容、発信者、投稿時間、いいね数、RT数、コメント数、インプレッション、ブックマーク数を日本語で表にしてください。
    投稿時間は「2025年7月3日　午後11:41」のように、日付→時刻の順で出力してください。
    例:
    | 投稿内容 | 発信者 | 投稿時間 | いいね数 | RT数 | コメント数 | インプレッション | ブックマーク数 |
    | 例の投稿内容 | 例の発信者 | 2025年73日　午後11:41 | 100 | 10 | 5 | 1万 | 20 |
    """
    response = model.generate_content([prompt, image_data])
    return response.text

def parse_table(text):
    lines = [l for l in text.splitlines() if "|" in l]
    if len(lines) < 2:
        return None
    data_lines = []
    for l in lines:
        cells = [c.strip() for c in l.split("|")[1:-1]]
        if all(cell.startswith(":") or set(cell) <= set("-:") for cell in cells):
            continue
        data_lines.append(l)
    if len(data_lines) < 2:
        return None
    headers_row = [h.strip() for h in data_lines[0].split("|")[1:-1]]
    values_row = [v.strip() for v in data_lines[1].split("|")[1:-1]]
    return dict(zip(headers_row, values_row))

def get_or_create_worksheet(spreadsheet, sheet_title, headers):
    try:
        # 既存のワークシートをタイトルで取得
        worksheet = spreadsheet.worksheet(sheet_title)
        st.write(f"既存のワークシート '{sheet_title}' を使用します。")
        return worksheet
    except gspread.exceptions.WorksheetNotFound:
        # ワークシートが存在しない場合、新規作成
        st.write(f"ワークシート '{sheet_title}' を新規作成します。")
        worksheet = spreadsheet.add_worksheet(title=sheet_title, rows="1000", cols="20")
        # ヘッダーを書き込む
        worksheet.append_row(headers)
        return worksheet

import json
import os

def get_or_create_user_spreadsheet(gc, user, title_prefix="Xポスト一覧_"):
    db_file = "user_sheets.json"
    if os.path.exists(db_file):
        with open(db_file, "r") as f:
            user_sheets = json.load(f)
    else:
        user_sheets = {}

    if user in user_sheets:
        spreadsheet_id = user_sheets[user]
        sh = gc.open_by_key(spreadsheet_id)
    else:
        sh = gc.create(f"{title_prefix}{user}")
        spreadsheet_id = sh.id
        user_sheets[user] = spreadsheet_id
        with open(db_file, "w") as f:
            json.dump(user_sheets, f)
    return sh

# --- Streamlit UI ---
st.title("Xポスト画像→スプレッドシート自動化アプリ")
st.write("画像をアップロードすると、内容を自動で抽出してGoogleスプレッドシートの、投稿者ごとのタブに追記します。")

# ユーザー名入力
user_name = st.text_input("あなたの名前を入力してください（スプレッドシートの識別に使われます）")

uploaded_file = st.file_uploader("画像をアップロードしてください（PNG/JPG）", type=["png", "jpg", "jpeg"])

# ユーザー名とファイルが両方入力された場合のみ処理を開始
if user_name and uploaded_file is not None:
    st.write("ファイルがアップロードされました")
    
    # 一時ファイルの作成
    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_file:
            tmp_file.write(uploaded_file.read())
            tmp_path = tmp_file.name

        st.image(tmp_path, caption="アップロード画像", use_column_width=True)
        st.info("画像を解析中...")

        # ユーザー専用のスプレッドシートを取得または作成
        st.write(f"'{user_name}'さんのスプレッドシートを取得/作成します。")
        user_spreadsheet = get_or_create_spreadsheet(gc, drive, user_name)
        if user_spreadsheet is None:
            st.error("スプレッドシートの準備に失敗しました。")
            # エラー発生時は後続処理を行わない
            if tmp_path and os.path.exists(tmp_path):
                os.remove(tmp_path)
                st.write("一時ファイル削除完了")
            st.stop()

        # 画像をGoogleドライブにアップロード
        st.write("Googleドライブにアップロード開始")
        image_url = upload_image_to_drive(tmp_path, drive)
        if image_url is None:
            st.error("Google Driveへの画像アップロードに失敗しました。")
            if tmp_path and os.path.exists(tmp_path):
                os.remove(tmp_path)
                st.write("一時ファイル削除完了")
            st.stop()
        
        st.write(f"image_url: {image_url}")
        image_formula = f'=IMAGE("{image_url}", 2)'  # 元サイズで表示

        # Geminiで情報抽出
        st.write("Geminiで情報抽出開始")
        result_text = extract_post_info(tmp_path, model)
        if result_text is None:
            st.error("Geminiでの情報抽出に失敗しました。")
            if tmp_path and os.path.exists(tmp_path):
                os.remove(tmp_path)
                st.write("一時ファイル削除完了")
            st.stop()

        st.text_area("Gemini抽出結果", result_text, height=200)

        info = parse_table(result_text)
        st.write(f"parse_tableの結果: {info}")
        
        if info:
            # 発信者名を取得
            poster_name = info.get("発信者")
            if not poster_name:
                st.error("発信者情報を抽出できませんでした。Geminiの出力形式を確認してください。")
                # エラー発生時は後続処理を行わない
                if tmp_path and os.path.exists(tmp_path):
                    os.remove(tmp_path)
                    st.write("一時ファイル削除完了")
                st.stop()

            # 発信者ごとのワークシートを取得または作成
            st.write(f"'{poster_name}'さんのタブを取得/作成します。")
            target_worksheet = get_or_create_worksheet(user_spreadsheet, poster_name, headers)
            if target_worksheet is None:
                st.error("ワークシートの準備に失敗しました。")
                if tmp_path and os.path.exists(tmp_path):
                    os.remove(tmp_path)
                    st.write("一時ファイル削除完了")
                st.stop()

            # データを追記
            row_data = [image_formula] + [info.get(h, "") for h in headers[1:]]
            try:
                target_worksheet.append_row(row_data, value_input_option='USER_ENTERED')
                st.success(f"スプレッドシート '{user_spreadsheet.title}' の '{poster_name}' タブに追記しました！")
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
            st.write("一時ファイル削除完了")
        else:
            st.write("一時ファイルは作成されなかったか、既に削除されています。")

elif uploaded_file is not None and not user_name:
    st.warning("画像をアップロードする前に、あなたの名前を入力してください。")
elif user_name and uploaded_file is None:
    st.info("画像をアップロードしてください。")
