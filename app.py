import streamlit as st
import os
import tempfile
import google.generativeai as genai
from PIL import Image
import gspread
from google.oauth2.service_account import Credentials # gspreadとPyDrive2の認証で使用
from pydrive2.auth import GoogleAuth
from pydrive2.drive import GoogleDrive
from gspread.exceptions import WorksheetNotFound, SpreadsheetNotFound
import json # jsonモジュールはここでインポート
import streamlit as st

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
def authenticate_pydrive():
    """PyDrive2をサービスアカウントで認証し、認証オブジェクトをキャッシュする"""
    import tempfile
    import os
    import json
    try:
        with tempfile.TemporaryDirectory() as temp_dir:
            # Streamlit secretsから認証情報を取得
            google_credentials = st.secrets["GOOGLE_CREDENTIALS"]
            if isinstance(google_credentials, str):
                cred_dict = json.loads(google_credentials)
            else:
                cred_dict = google_credentials

            # client_secrets.jsonを一時ファイルに保存
            client_secrets_path = os.path.join(temp_dir, "client_secrets.json")
            with open(client_secrets_path, "w") as f:
                json.dump(cred_dict, f)

            # settings.yamlを一時ファイルに保存
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

            # カレントディレクトリを一時ディレクトリに移動
            old_cwd = os.getcwd()
            os.chdir(temp_dir)

            from pydrive2.auth import GoogleAuth
            from pydrive2.drive import GoogleDrive

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
    """Gemini APIを設定し、モデルをキャッシュする"""
    try:
        genai.configure(api_key=GENAI_API_KEY)
        model = genai.GenerativeModel("gemini-2.5-flash")
        # st.success("Gemini API設定に成功しました。") # この行を削除/コメントアウト
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
        # 誰でも閲覧できるように権限を設定 (修正点: コメントアウトを解除し、公開設定を有効化)
        file.InsertPermission({'type': 'anyone', 'value': 'anyone', 'role': 'reader'})
        # st.write(f"画像 '{file_name}' をGoogle Driveにアップロードしました。") # 処理メッセージを削除
        return f"https://drive.google.com/uc?id={file['id']}"
    except Exception as e:
        st.error(f"Google Driveへの画像アップロード中にエラーが発生しました: {e}")
        return None

def extract_post_info(image_path, gemini_model):
    """
    Gemini APIを使用して画像から投稿情報を抽出する。
    プロンプトを修正し、「アカウントID」も抽出するように変更。
    また、Geminiの出力に含まれる<br>タグを改行コードに置換する。
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
        # <br>タグを改行コードに置換
        cleaned_text = response.text.replace('<br>', '\n').replace('<BR>', '\n')
        return cleaned_text
    except Exception as e:
        st.error(f"Gemini APIでの情報抽出中にエラーが発生しました: {e}")
        return None

def parse_table(text):
    """
    Geminiからのテキスト結果をパースして辞書形式で返す。
    Geminiの出力形式のばらつきに対応するため、より堅牢なパース処理を行う。
    """
    if not text:
        return None

    lines = text.splitlines()
    
    # ヘッダー行とデータ行を特定
    header_line = None
    data_line = None
    
    for line in lines:
        if "|" in line:
            # Markdownの区切り行をスキップ
            if all(c.strip().startswith(":") or set(c.strip()) <= set("-:") for c in line.split("|")[1:-1]):
                continue
            
            if header_line is None:
                header_line = line
            elif data_line is None:
                data_line = line
                break # 最初のデータ行を見つけたら終了

    if header_line is None or data_line is None:
        st.warning("Geminiの出力から有効なヘッダー行とデータ行を特定できませんでした。")
        return None

    headers_raw = [h.strip() for h in header_line.split("|")[1:-1]]
    values_raw = [v.strip() for v in data_line.split("|")[1:-1]]

    # ヘッダーと値の数を比較し、調整
    info = {}
    num_headers = len(headers_raw)
    num_values = len(values_raw)

    for i in range(num_headers):
        header = headers_raw[i]
        value = values_raw[i] if i < num_values else "" # 値が足りない場合は空文字列で埋める
        info[header] = value
    
    # もし値の数がヘッダーより多い場合、余分な値は無視する
    if num_values > num_headers:
        st.warning(f"Geminiの出力に予期せぬ追加の列が含まれていました。ヘッダー数: {num_headers}, 値の数: {num_values}")

    return info

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
        # st.write(f"既存のスプレッドシート '{spreadsheet_title}' を使用します。") # 処理メッセージを削除
    except SpreadsheetNotFound:
        # スプレッドシートが存在しない場合、新規作成
        # st.write(f"スプレッドシート '{spreadsheet_title}' を新規作成します。") # 処理メッセージを削除
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
    新規作成時にはヘッダーを書き込み、初期設定を適用する。
    """
    try:
        # 既存のワークシートをタイトルで取得
        worksheet = spreadsheet.worksheet(sheet_title)
        # st.write(f"既存のワークシート '{sheet_title}' を使用します。") # 処理メッセージを削除
        return worksheet
    except WorksheetNotFound:
        # ワークシートが存在しない場合、新規作成
        # st.write(f"ワークシート '{sheet_title}' を新規作成します。") # 処理メッセージを削除
        # add_worksheetのrows/colsは目安。必要に応じて調整。
        worksheet = spreadsheet.add_worksheet(title=sheet_title, rows="1000", cols="20")
        # ヘッダーを書き込む
        worksheet.append_row(headers_list)

        # --- 新規作成されたワークシートへの初期設定適用 ---
        sheet_id = worksheet._properties['sheetId']
        requests = []

        # 1. 1行固定 (最初の行を固定)
        requests.append({
            "updateSheetProperties": {
                "properties": {
                    "sheetId": sheet_id,
                    "gridProperties": {
                        "frozenRowCount": 1 # frozenColumnCount を frozenRowCount に変更
                    }
                },
                "fields": "gridProperties.frozenRowCount" # fields も変更
            }
        })

        # 2. 全て文字は中央揃え（垂直方向、水平方向両方とも）
        requests.append({
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 0,
                    "endRowIndex": 1000, # 十分な範囲を指定
                    "startColumnIndex": 0,
                    "endColumnIndex": 20 # 十分な範囲を指定
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

        # 3. 行2 280ピクセル (0-indexed: startIndex=1, endIndex=2)
        requests.append({
            "updateDimensionProperties": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "ROWS",
                    "startIndex": 1,
                    "endIndex": 2
                },
                "properties": {
                    "pixelSize": 280
                },
                "fields": "pixelSize"
            }
        })

        # 4. 列A 280ピクセル (0-indexed: startIndex=0, endIndex=1)
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

        # 5. 列B テキストを折り返す (0-indexed: startIndex=1, endIndex=2)
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
                    "endIndex": 20 # T列の次まで
                }
            }
        })

        # 7. 列C〜Jの幅を100ピクセルに固定 (0-indexed: C=2, J=9)
        requests.append({
            "updateDimensionProperties": { # autoResizeDimensions から変更
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "startIndex": 2,
                    "endIndex": 10 # J列の次まで
                },
                "properties": {
                    "pixelSize": 100 # 幅を100ピクセルに設定
                },
                "fields": "pixelSize" # fields も変更
            }
        })
        
        try:
            spreadsheet.batch_update({"requests": requests})
            st.success(f"ワークシート '{sheet_title}' の初期設定を適用しました。")
        except Exception as update_e:
            st.warning(f"ワークsheet '{sheet_title}' の初期設定適用中にエラーが発生しました。手動で設定してください: {update_e}")

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

# 複数ファイルアップロードを許可し、最大30枚まで
uploaded_files = st.file_uploader("画像をアップロードしてください（PNG/JPG、最大30枚）", type=["png", "jpg", "jpeg"], accept_multiple_files=True)

# メールアドレスとファイルが両方入力された場合のみ処理を開始
if email and uploaded_files: # uploaded_filesが空リストでないことを確認
    
    # ここに全体の処理を囲むtryブロックを追加
    try:
        # ユーザー専用のスプレッドシートを取得または作成
        user_spreadsheet = get_or_create_spreadsheet(gc, drive, email) # emailを渡す
        if user_spreadsheet is None:
            st.error("スプレッドシートの準備に失敗しました。")
            st.stop()

        # プログレスバーの初期化
        progress_text = "画像を処理中..."
        progress_bar = st.progress(0, text=progress_text)
        total_files = len(uploaded_files)

        for i, uploaded_file in enumerate(uploaded_files):
            current_file_tmp_path = None # 各ファイルのテンポラリパスをここで初期化
            try:
                # プログレスバーの更新
                progress_percent = (i + 1) / total_files
                progress_bar.progress(progress_percent, text=f"画像を処理中: {i+1}/{total_files}枚目")

                with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_file:
                    tmp_file.write(uploaded_file.read())
                    current_file_tmp_path = tmp_file.name

                st.image(current_file_tmp_path, caption=f"アップロード画像 {i+1}", use_container_width=True)
                st.info(f"画像を解析中... ({i+1}/{total_files}枚目)") # ユーザーへの状態表示は残す

                # 画像をGoogleドライブにアップロード
                image_url = upload_image_to_drive(current_file_tmp_path, drive)
                if image_url is None:
                    st.error(f"Google Driveへの画像アップロードに失敗しました。({i+1}/{total_files}枚目)")
                    continue # 次のファイルへ

                image_formula = f'=IMAGE("{image_url}", 2)'  # 元サイズで表示

                # Geminiで情報抽出
                result_text = extract_post_info(current_file_tmp_path, model)
                if result_text is None:
                    st.error(f"Geminiでの情報抽出に失敗しました。({i+1}/{total_files}枚目)")
                    continue # 次のファイルへ

                # 修正点: Gemini抽出結果のデバッグ表示を追加
                st.text_area(f"Gemini抽出結果 ({i+1}/{total_files}枚目)", result_text, height=200) 
                st.json(parse_table(result_text)) # 抽出されたinfo辞書をJSON形式で表示

                info = parse_table(result_text)
                
                if info:
                    # 発信者名とアカウントIDを取得
                    author_name = info.get("発信者名")
                    account_id = info.get("アカウントID")

                    if not author_name:
                        st.error(f"発信者名情報を抽出できませんでした。Geminiの出力形式を確認してください。({i+1}/{total_files}枚目)")
                        continue # 次のファイルへ

                    # タブ名を「発信者名（@アカウントID）」の形式で生成
                    tab_name = f"{author_name}（@{account_id}）" if account_id else author_name
                    
                    # 発信者ごとのワークシートを取得または作成
                    target_worksheet = get_or_create_worksheet(user_spreadsheet, tab_name, headers)
                    if target_worksheet is None:
                        st.error(f"ワークシートの準備に失敗しました。({i+1}/{total_files}枚目)")
                        continue # 次のファイルへ

                    # データを追記
                    row_data = [image_formula, info.get("投稿内容", ""), info.get("発信者名", ""), info.get("アカウントID", ""), 
                                info.get("投稿時間", ""), info.get("いいね数", ""), info.get("RT数", ""), 
                                info.get("コメント数", ""), info.get("インプレッション", ""), info.get("ブックマーク数", "")]
                    try:
                        target_worksheet.append_row(row_data, value_input_option='USER_ENTERED')
                        st.success(f"スプレッドシート '{user_spreadsheet.title}' の '{tab_name}' タブに追記しました！ ({i+1}/{total_files}枚目)")
                        # URL表示をここから削除
                        # st.markdown(f"[スプレッドシートを開く]({user_spreadsheet.url})")
                    except Exception as e:
                        st.error(f"スプレッドシートへの追記中にエラーが発生しました: {e} ({i+1}/{total_files}枚目)")
                else:
                    st.error(f"情報の抽出に失敗しました。Geminiの出力形式を確認してください。({i+1}/{total_files}枚目)")

            except Exception as e:
                st.error(f"ファイル処理中に予期せぬエラーが発生しました: {e} ({i+1}/{total_files}枚目)")
            finally:
                # 一時ファイル削除
                if current_file_tmp_path and os.path.exists(current_file_tmp_path):
                    os.remove(current_file_tmp_path)
        
        progress_bar.empty() # プログレスバーを非表示にする
        st.success("すべての画像の処理が完了しました！")
        # 全ての処理が完了した後にスプレッドシートのURLを表示
        if user_spreadsheet: # user_spreadsheetがNoneでないことを確認
            st.markdown(f"[スプレッドシートを開く]({user_spreadsheet.url})")

    except Exception as outer_e: # 全体の処理で発生したエラーをキャッチ
        st.error(f"全体の処理中に予期せぬエラーが発生しました: {outer_e}")

elif uploaded_files and not email: # uploaded_filesが空リストでないことを確認
    st.warning("画像をアップロードする前に、あなたのGoogleメールアドレスを入力してください。")
elif email and not uploaded_files: # uploaded_filesが空リストであることを確認
    st.info("画像をアップロードしてください。")
