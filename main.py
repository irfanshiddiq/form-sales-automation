import os
import json
import datetime
from playwright.sync_api import sync_playwright
import gspread
from google.oauth2.service_account import Credentials

# ==========================================
# 1. 設定項目
# ==========================================
# 「営業お断り」を判定するためのNGワードリスト
# ページ内にこの単語が含まれていた場合は送信をスキップします
NG_WORDS = [
    "営業お断り", "セールスお断り", "営業目的", "ご遠慮", 
    "お断りいたします", "お断りします", "固くお断り", "営業のご連絡",
    "特定電子メール法", "人材紹介", "採用代行"
]

# スプレッドシートへアクセスするための権限設定
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

# ==========================================
# 2. 補助関数（便利ツール）
# ==========================================
def get_gspread_client():
    """Google Cloudのサービスアカウントキーを読み込み、認証する"""
    sa_json = os.environ.get("GCP_SA_KEY")
    if not sa_json:
        raise ValueError("エラー: 環境変数 'GCP_SA_KEY' が設定されていません。GitHubのSecretsを確認してください。")
    
    creds_dict = json.loads(sa_json)
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)

def check_ng_words(page_text):
    """ページテキスト内にNGワードが含まれていないかチェックする"""
    for word in NG_WORDS:
        if word in page_text:
            return True, word
    return False, None

def fill_input(page, selectors, value):
    """複数の候補(CSSセレクタ)から、最初に見つかった入力欄にテキストを入れる"""
    for selector in selectors:
        element = page.locator(selector).first
        if element.count() > 0 and element.is_visible():
            try:
                element.fill(value)
                return True
            except Exception:
                continue
    return False

# ==========================================
# 3. フォーム送信のメイン処理関数
# ==========================================
def process_form(page, url, text_data):
    """1件のフォームにアクセスして入力・送信を行う"""
    try:
        # フォームにアクセス（最大30秒待機）
        page.goto(url, timeout=30000, wait_until="domcontentloaded")
        page.wait_for_timeout(3000) # 画面が完全に描画されるまで少し待機

        # --- ① NGワード（営業お断り）チェック ---
        page_content = page.content()
        is_ng, detected_word = check_ng_words(page_content)
        if is_ng:
            return "SKIPPED_NG", f"NGワード検出: {detected_word}"

        # --- ② 各項目の自動入力 ---
        # 氏名の入力
        fill_input(page, [
            'input[name*="name" i]', 'input[name*="namae" i]', 
            'input[placeholder*="名" i]', 'input[placeholder*="氏" i]', 'input[id*="name" i]'
        ], text_data.get("氏名", ""))

        # 会社名の入力
        fill_input(page, [
            'input[name*="company" i]', 'input[name*="corp" i]', 
            'input[placeholder*="会社" i]', 'input[placeholder*="貴社" i]'
        ], text_data.get("貴社名", ""))

        # メールアドレスの入力
        fill_input(page, [
            'input[type="email"]', 'input[name*="email" i]', 
            'input[name*="mail" i]', 'input[placeholder*="メール" i]'
        ], text_data.get("メールアドレス", ""))

        # 電話番号の入力
        fill_input(page, [
            'input[type="tel"]', 'input[name*="tel" i]', 
            'input[name*="phone" i]', 'input[placeholder*="電話" i]'
        ], text_data.get("電話番号", ""))

        # 件名の入力
        fill_input(page, [
            'input[name*="subject" i]', 'input[name*="title" i]', 
            'input[placeholder*="件名" i]', 'input[placeholder*="タイトル" i]'
        ], text_data.get("件名", ""))

        # 本文の入力
        fill_input(page, [
            'textarea[name*="body" i]', 'textarea[name*="message" i]', 
            'textarea[name*="content" i]', 'textarea[placeholder*="内容" i]', 'textarea[placeholder*="本文" i]'
        ], text_data.get("本文", ""))

        page.wait_for_timeout(2000)

        # --- ③ 送信ボタンのクリック ---
        # 「送信」や「確認」ボタンを探す
        submit_button = page.locator('button[type="submit"], input[type="submit"], button:has-text("確認"), button:has-text("送信"), input:has-text("確認")').first
        
        if submit_button.count() > 0 and submit_button.is_visible():
            # ========================================================
            # 【重要】実際の送信処理（テスト段階ではコメントアウト推奨）
            # ========================================================
            submit_button.click()
            page.wait_for_timeout(5000) # 送信完了画面への遷移を待つ
            return "SUCCESS", "送信成功"
            
            # テスト時は上記をコメントアウトし、下記を有効にしてください。
            # return "SUCCESS_TEST", "テスト: 入力完了（送信ボタン未押下）"
        else:
            return "WARNING", "送信ボタンが見つかりませんでした"

    except Exception as e:
        return "ERROR", f"エラー発生: {str(e)}"

# ==========================================
# 4. 実行の司令塔（メインプログラム）
# ==========================================
def main():
    # スプレッドシートIDの取得
    spreadsheet_id = os.environ.get("SPREADSHEET_ID")
    if not spreadsheet_id:
        raise ValueError("エラー: 環境変数 'SPREADSHEET_ID' が設定されていません。")

    print("スプレッドシートに接続中...")
    gc = get_gspread_client()
    sh = gc.open_by_key(spreadsheet_id)

    # 各シートの取得
    target_sheet = sh.worksheet("ターゲットリスト")
    config_sheet = sh.worksheet("送信文面")
    history_sheet = sh.worksheet("送信履歴")

    # --- 送信文面の読み込み ---
    config_rows = config_sheet.get_all_values()[1:] # 1行目のヘッダーを除外
    text_data = {row[0]: row[1] for row in config_rows if len(row) >= 2}

    # --- 送信履歴の読み込み（重複チェック用） ---
    # 履歴シートのB列（URL）を取得
    history_urls = set(history_sheet.col_values(2)[1:]) 

    # --- ターゲットリストの読み込み ---
    targets = target_sheet.get_all_records()

    # --- Playwright（ブラウザ）の起動 ---
    with sync_playwright() as p:
        # headless=True だと画面を出さずに裏で動く（サーバー向け）
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()

        # リストを1件ずつ処理
        for target in targets:
            company_name = target.get("企業名", "")
            url = target.get("フォームURL", "")
            
            # C列の「送信対象」が TRUE かどうかチェック
            is_target = str(target.get("送信対象", "")).upper() == "TRUE"

            # 送信対象外ならスキップ
            if not is_target:
                continue

            # 重複チェック: すでに履歴にあるURLならスキップ
            if url in history_urls:
                print(f"スキップ（重複あり）: {company_name} - {url}")
                continue

            # 処理開始
            print(f"送信処理開始: {company_name} - {url}")
            status, note = process_form(page, url, text_data)

            # --- 結果をスプレッドシートに追記 ---
            now = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=9))).strftime("%Y-%m-%d %H:%M:%S")
            history_sheet.append_row([company_name, url, now, status, note])
            
            # 今回送ったURLを重複チェック用リストに追加
            history_urls.add(url)

            print(f"結果: {status} ({note})")
            print("-" * 30)

            # 相手先のサーバーに負荷をかけないよう、次の送信まで10秒待機
            page.wait_for_timeout(10000)

        # ブラウザを閉じて終了
        browser.close()
        print("すべての処理が完了しました！")

if __name__ == "__main__":
    main()
