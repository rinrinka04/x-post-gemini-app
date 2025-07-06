import streamlit as st
import os
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

# --- Google Drive認証 ---
gauth = GoogleAuth()
gauth.ServiceAuth()
drive = GoogleDrive(gauth)

# --- スプレッドシート ---
sh = gc.open_by_key(SPREADSHEET_ID)

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

# --- Streamlit UI ---
st.title("Xポスト画像→スプレッドシート自動化アプリ")
st.write("画像をアップロードすると、内容を自動で抽出してGoogleスプレッドシートの、投稿者ごとのタブに追記します。")

uploaded_file = st.file_uploader("画像をアップロードしてください（PNG/JPG）", type=["png", "jpg", "jpeg"])

if uploaded_file is not None:
    st.write("ファイルがアップロードされました")
    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_file:
        tmp_file.write(uploaded_file.read())
        tmp_path = tmp_file.name

    st.image(tmp_path, caption="アップロード画像", use_column_width=True)
    st.info("画像を解析中...")

    # 画像をGoogleドライブにアップロード
    st.write("Googleドライブにアップロード開始")
    image_url = upload_image_to_drive(tmp_path)
    st.write(f"image_url: {image_url}")
    image_formula = f'=IMAGE("{image_url}", 2)'  # 元サイズで表示

    # Geminiで情報抽出
    st.write("Geminiで情報抽出開始")
    result_text = extract_post_info(tmp_path)
    st.text_area("Gemini抽出結果", result_text, height=200)

    info = parse_table(result_text)
    st.write(f"parse_tableの結果: {info}")
    if info:
        # 発信者名を取得
        poster_name = info.get("発信者")
        if not poster_name:
            st.error("発信者情報を抽出できませんでした。")
            os.remove(tmp_path)
            st.write("一時ファイル削除完了")
            st.stop()

        # 発信者ごとのワークシートを取得または作成
        try:
            worksheet = get_or_create_worksheet(sh, poster_name, headers)
        except Exception as e:
            st.error(f"ワークシートの取得または作成中にエラーが発生しました: {e}")
            os.remove(tmp_path)
            st.write("一時ファイル削除完了")
            st.stop()

        # データを追記
        row = [image_formula] + [info.get(h, "") for h in headers[1:]]
        try:
            worksheet.append_row(row, value_input_option='USER_ENTERED')
            st.success(f"スプレッドシートの'{poster_name}'タブに追記しました！")
            st.markdown(f"[スプレッドシートを開く]({sh.url})")
        except Exception as e:
            st.error(f"スプレッドシートへの追記中にエラーが発生しました: {e}")
    else:
        st.error("情報の抽出に失敗しました。")

    # 一時ファイル削除
    os.remove(tmp_path)
    st.write("一時ファイル削除完了")
