```python
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
import json

# --- パスワード認証 ---
PASSWORD = "xpost00"

if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

if not st.session_state["authenticated"]:
    pw = st.text_input("パスワードを入力してください", type="password")
    if pw == PASSWORD:
        st.session_state["authenticated"] = True
        st.success("認証成功！")
        st.rerun()
    else:
        st.error("パスワードが異なります。")
        st.stop()

# --- 設定 ---
GENAI_API_KEY = "AIzaSyCc2MQQ2ytt32gzMq53L_Z8SKhWWMRjJ1s"  # 必ずご自身のAPIキーに変更
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]
headers = ["画像", "投稿内容", "発信者名", "アカウントID", "投稿時間", "いいね数", "RT数", "コメント数", "インプレッション", "ブックマーク数"]

# --- gspread認証 ---
@st.cache_resource
def authenticate_gspread():
    try:
        google_credentials = st.secrets["GOOGLE_CREDENTIALS"]
        if isinstance(google_credentials, str):
            cred_dict = json.loads(google_credentials)
        else:
            cred_dict = google_credentials
        creds = Credentials.from_service_account_info(cred_dict, scopes=SCOPES)
        gc = gspread.authorize(creds)
        return gc
    except Exception as e:
        st.error(f"Google Sheets認証に失敗しました。認証情報を確認してください: {e}")
        st.stop()

# --- PyDrive2認証 ---
@st.cache_resource
def authenticate_pydrive():
    try:
        google_credentials = st.secrets["GOOGLE_CREDENTIALS"]
        if isinstance(google_credentials, str):
            cred_dict = json.loads(google_credentials)
        else:
            cred_dict = google_credentials

        temp_dir = tempfile.mkdtemp()
        client_secrets_path = os.path.join(temp_dir, "client_secrets.json")
        with open(client_secrets_path, "w") as f:
            json.dump(cred_dict, f)

        settings_yaml = f"""
client_config_backend: file
client_config_file: client_secrets.json
save_credentials: False
oauth_scope:
  - https://www.googleapis.com/auth/drive
  - https://www.googleapis.com/auth/drive.file
  - https://www.googleapis.com/auth/drive.appdata
  - https://www.googleapis.com/auth/drive.metadata
  - https://www.googleapis.com/auth/drive.scripts
service_config:
  client_user_email: {cred_dict['client_email']}
  client_json_file_path: {client_secrets_path}
"""
        settings_path = os.path.join(temp_dir, "settings.yaml")
        with open(settings_path, "w") as f:
            f.write(settings_yaml)

        old_cwd = os.getcwd()
        os.chdir(temp_dir)

        gauth = GoogleAuth(settings_file=settings_path)
        gauth.ServiceAuth()
        drive = GoogleDrive(gauth)

        os.chdir(old_cwd)
        return drive
    except Exception as e:
        st.error(f"Google Drive認証に失敗しました。認証設定を確認してください: {e}")
        st.stop()

# --- Gemini API ---
@st.cache_resource
def configure_gemini():
    try:
        genai.configure(api_key=GENAI_API_KEY)
        model = genai.GenerativeModel("gemini-2.5-flash")
        return model
    except Exception as e:
        st.error(f"Gemini APIキーの設定に失敗しました。APIキーを確認してください: {e}")
        st.stop()

gc = authenticate_gspread()
drive = authenticate_pydrive()
model = configure_gemini()

def upload_image_to_drive(image_path, drive_service):
    try:
        file_name = os.path.basename(image_path)
        file = drive_service.CreateFile({'title': file_name})
        file.SetContentFile(image_path)
        file.Upload()
        file.InsertPermission({'type': 'anyone', 'value': 'anyone', 'role': 'reader'})
        return f"https://drive.google.com/uc?id={file['id']}"
    except Exception as e:
        st.error(f"Google Driveへの画像アップロード中にエラーが発生しました: {e}")
        return None

def extract_post_info(image_path, gemini_model):
    try:
        image_data = Image.open(image_path)
        prompt = """
この画像はX（旧Twitter）のポストです。
下記の9項目を必ず「Markdownテーブル形式（1行目:ヘッダー, 2行目:値）」で出力してください。
- 投稿内容
- 発信者名
- アカウントID
- 投稿時間
- いいね数
- RT数
- コメント数
- インプレッション
- ブックマーク数

【出力例】
| 投稿内容 | 発信者名 | アカウントID | 投稿時間 | いいね数 | RT数 | コメント数 | インプレッション | ブックマーク数 |
| 例の投稿内容 | 田中太郎 | tanaka_taro | 2025年7月3日 午後11:41 | 100 | 10 | 5 | 1万 | 20 |

**絶対に上記のようなMarkdownテーブル形式、もしくはJSON形式でのみ出力してください。**
**それ以外の文章や説明は一切不要です。**
"""
        response = gemini_model.generate_content([prompt, image_data])
        cleaned_text = response.text.replace('<br>', '\n').replace('<BR>', '\n').replace('<br/>', '\n').replace('<BR/>', '\n')
        return cleaned_text
    except Exception as e:
        st.error(f"Gemini APIでの情報抽出中にエラーが発生しました: {e}")
        return None

def parse_table(text):
    if not text:
        return None
    # JSON形式で返ってきた場合に対応
    try:
        data = json.loads(text)
        if isinstance(data, dict):
            return data
        if isinstance(data, list) and len(data) == 1 and isinstance(data[0], dict):
            return data[0]
    except Exception:
        pass
    # テーブルパース
    lines = [l for l in text.splitlines() if "|" in l]
    if len(lines) < 2:
        return None
    # 区切り線（|---|---|）を除外
    data_lines = []
    for l in lines:
        cells = [c.strip() for c in l.split("|")[1:-1]]
        if all(set(cell) <= set("-:") for cell in cells):
            continue
        data_lines.append(l)
    if len(data_lines) < 2:
        return None
    headers_row = [h.strip() for h in data_lines[0].split("|")[1:-1]]
    # 2行目以降で値行を探す
    for values_line in data_lines[1:]:
        values_row = [v.strip() for v in values_line.split("|")[1:-1]]
        if len(headers_row) == len(values_row):
            return dict(zip(headers_row, values_row))
    return None

def get_or_create_spreadsheet(gspread_client, drive_service, user_email):
    spreadsheet_title = f"Xポスト自動化_{user_email}"
    spreadsheet = None
    try:
        spreadsheet = gspread_client.open(spreadsheet_title)
    except SpreadsheetNotFound:
        spreadsheet = gspread_client.create(spreadsheet_title)
        st.success(f"新しいスプレッドシート '{spreadsheet_title}' を作成しました。")
    except Exception as e:
        st.error(f"スプレッドシートの取得または作成中にエラーが発生しました: {e}")
        return None

    if spreadsheet:
        try:
            spreadsheet.share(user_email, perm_type='user', role='writer')
            st.success(f"スプレッドシート '{spreadsheet_title}' に {user_email} の編集権限を設定しました。")
        except Exception as share_e:
            st.warning(f"スプレッドシートの共有設定中にエラーが発生しました。手動で共有設定を行ってください: {share_e}")
    return spreadsheet

def set_worksheet_format(spreadsheet, worksheet):
    try:
        sheet_id = worksheet._properties['sheetId']
        requests = []

        # 1. 1行固定
        requests.append({
            "updateSheetProperties": {
                "properties": {
                    "sheetId": sheet_id,
                    "gridProperties": {
                        "frozenRowCount": 1
                    }
                },
                "fields": "gridProperties.frozenRowCount"
            }
        })

        # 2. 全て文字は中央揃え（垂直方向、水平方向両方とも）
        requests.append({
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id
                },
                "cell": {
                    "userEnteredFormat": {
                        "horizontalAlignment": "CENTER",
                        "verticalAlignment": "MIDDLE"
                    }
                },
                "fields": "userEnteredFormat.horizontalAlignment,userEnteredFormat.verticalAlignment"
            }
        })

        # 3. 行2以降 280ピクセル
        requests.append({
            "updateDimensionProperties": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "ROWS",
                    "startIndex": 1,
                    "endIndex": 1000
                },
                "properties": {
                    "pixelSize": 280
                },
                "fields": "pixelSize"
            }
        })

        # 4. 列A 280ピクセル
        requests.append({
            "updateDimensionProperties": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "startIndex": 0,
                    "endIndex": 1
                },
                "properties": {
                    "pixelSize": 280
                },
                "fields": "pixelSize"
            }
        })

        # 5. 列B 200ピクセル & テキストを折り返す
        requests.append({
            "updateDimensionProperties": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "startIndex": 1,
                    "endIndex": 2
                },
                "properties": {
                    "pixelSize": 200
                },
                "fields": "pixelSize"
            }
        })
        requests.append({
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startColumnIndex": 1,
                    "endColumnIndex": 2
                },
                "cell": {
                    "userEnteredFormat": {
                        "wrapStrategy": "WRAP"
                    }
                },
                "fields": "userEnteredFormat.wrapStrategy"
            }
        })

        # 6. 列K〜T削除 (0-indexed: K=10, T=19)
        requests.append({
            "deleteDimension": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "startIndex": 10,
                    "endIndex": 20
                }
            }
        })

        # 7. 列C〜Jの幅は100ピクセル (0-indexed: C=2, J=9)
        requests.append({
            "updateDimensionProperties": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "startIndex": 2,
                    "endIndex": 10
                },
                "properties": {
                    "pixelSize": 100
                },
                "fields": "pixelSize"
            }
        })

        spreadsheet.batch_update({"requests": requests})
        st.success(f"ワークシート '{worksheet.title}' の初期設定を適用しました。")
    except Exception as e:
        st.warning(f"ワークシート '{worksheet.title}' の初期設定適用中にエラーが発生しました: {e}")

def get_or_create_worksheet(spreadsheet, sheet_title, headers_list):
    try:
        worksheet = spreadsheet.worksheet(sheet_title)
        return worksheet
    except WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(title=sheet_title, rows="1000", cols="20")
        worksheet.append_row(headers_list)
        set_worksheet_format(spreadsheet, worksheet)
        return worksheet
    except Exception as e:
        st.error(f"ワークシートの取得または作成中にエラーが発生しました: {e}")
        return None

# --- Streamlit UI ---
st.title("Xポスト画像→スプレッドシート自動化アプリ")
st.write("画像をアップロードすると、内容を自動で抽出してあなた専用のGoogleスプレッドシートの、投稿者ごとのタブに追記します。")

email = st.text_input("あなたのGoogleメールアドレスを入力してください")
uploaded_files = st.file_uploader(
    "画像をアップロードしてください（PNG/JPG、最大30枚）",
    type=["png", "jpg", "jpeg"],
    accept_multiple_files=True
)

if email and uploaded_files:
    total_files = min(len(uploaded_files), 30)
    errors = []
    user_spreadsheet = None
    for i, uploaded_file in enumerate(uploaded_files[:30]):
        tmp_path = None
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_file:
                tmp_file.write(uploaded_file.read())
                tmp_path = tmp_file.name

            if user_spreadsheet is None:
                user_spreadsheet = get_or_create_spreadsheet(gc, drive, email)
            if user_spreadsheet is None:
                errors.append(f"{i+1}枚目: スプレッドシートの準備に失敗しました。")
                if tmp_path and os.path.exists(tmp_path):
                    os.remove(tmp_path)
                continue

            image_url = upload_image_to_drive(tmp_path, drive)
            if image_url is None:
                errors.append(f"{i+1}枚目: Google Driveへの画像アップロードに失敗しました。")
                if tmp_path and os.path.exists(tmp_path):
                    os.remove(tmp_path)
                continue

            image_formula = f'=IMAGE("{image_url}", 2)'

            result_text = extract_post_info(tmp_path, model)
            # --- Geminiの生出力を必ず表示（デバッグ用） ---
            st.text_area(f"Geminiの生出力 ({i+1}枚目)", result_text or "None", height=200)

            if result_text is None:
                errors.append(f"{i+1}枚目: Geminiでの情報抽出に失敗しました。")
                if tmp_path and os.path.exists(tmp_path):
                    os.remove(tmp_path)
                continue

            info = parse_table(result_text)

            if info:
                author_name = info.get("発信者名")
                account_id = info.get("アカウントID")

                if not author_name or not account_id:
                    errors.append(f"{i+1}枚目: 発信者名またはアカウントID情報を抽出できませんでした。Geminiの出力形式を確認してください。")
                    if tmp_path and os.path.exists(tmp_path):
                        os.remove(tmp_path)
                    continue

                tab_name = f"{author_name}"

                target_worksheet = get_or_create_worksheet(user_spreadsheet, tab_name, headers)
                if target_worksheet is None:
                    errors.append(f"{i+1}枚目: ワークシートの準備に失敗しました。")
                    if tmp_path and os.path.exists(tmp_path):
                        os.remove(tmp_path)
                    continue

                post_content = info.get("投稿内容", "").replace('<br>', '\n').replace('<BR>', '\n').replace('<br/>', '\n').replace('<BR/>', '\n')

                row_data = [image_formula, post_content, info.get("発信者名", ""), info.get("アカウントID", ""),
                            info.get("投稿時間", ""), info.get("いいね数", ""), info.get("RT数", ""),
                            info.get("コメント数", ""), info.get("インプレッション", ""), info.get("ブックマーク数", "")]
                try:
                    target_worksheet.append_row(row_data, value_input_option='USER_ENTERED')
                except Exception as e:
                    errors.append(f"{i+1}枚目: スプレッドシートへの追記中にエラーが発生しました: {e}")
            else:
                errors.append(f"{i+1}枚目: 情報の抽出に失敗しました。Geminiの出力形式を確認してください。")

        except Exception as e:
            errors.append(f"{i+1}枚目: 予期せぬエラーが発生しました: {e}")
        finally:
            if tmp_path and os.path.exists(tmp_path):
                os.remove(tmp_path)

    if user_spreadsheet:
        st.markdown(f"[スプレッドシートを開く]({user_spreadsheet.url})")
    if errors:
        st.warning("\\n".join(errors))

elif uploaded_files and not email:
    st.warning("画像をアップロードする前に、あなたのGoogleメールアドレスを入力してください。")
elif email and not uploaded_files:
    st.info("画像をアップロードしてください。")
```

---
