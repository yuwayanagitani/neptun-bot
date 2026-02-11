import pyotp, time, json, os, platform, requests
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException
from webdriver_manager.chrome import ChromeDriverManager

# 月名の変換マップ
MONTH_MAP = {
    "január": 1, "február": 2, "március": 3, "április": 4, "május": 5, "június": 6,
    "július": 7, "augusztus": 8, "szeptember": 9, "október": 10, "november": 11, "december": 12,
    "January": 1, "February": 2, "March": 3, "April": 4, "May": 5, "June": 6,
    "July": 7, "August": 8, "September": 9, "October": 10, "November": 11, "December": 12
}

def load_config():
    config_path = os.path.join(os.path.dirname(__file__), 'config.json')
    with open(config_path, 'r', encoding='utf-8') as f:
        return json.load(f)

def speak_message(message):
    current_os = platform.system()
    try:
        if current_os == "Darwin": os.system(f"say '{message}'")
        elif current_os == "Windows": os.system(f'PowerShell -Command "Add-Type –AssemblyName System.Speech; (New-Object System.Speech.Synthesis.SpeechSynthesizer).Speak(\'{message}\')"')
        elif current_os == "Linux": os.system(f"espeak '{message}' &")
    except: pass

def send_telegram(conf, message):
    token, chat_id = conf['telegram']['bot_token'], conf['telegram']['chat_id']
    url = f"https://api.telegram.org/bot{token}/sendMessage"
    try: requests.post(url, data={"chat_id": chat_id, "text": message, "parse_mode": "Markdown"})
    except: pass

def get_2fa_code(raw_secret):
    secret = raw_secret.replace(" ", "").upper()
    padding = len(secret) % 8
    if padding != 0: secret += '=' * (8 - padding)
    return pyotp.TOTP(secret).now()

def parse_neptun_date(date_str):
    try:
        clean_str = date_str.replace('at', '').replace('.', '').replace(':', ' ')
        parts = clean_str.split()
        if parts[0].isdigit() and len(parts[0]) == 4:
            year, month, day = int(parts[0]), MONTH_MAP.get(parts[1].lower()), int(parts[2])
        else:
            day, month, year = int(parts[0]), MONTH_MAP.get(parts[1]), int(parts[2])
        return datetime(year, month, day)
    except: return None

def login_and_prepare(conf):
    options = webdriver.ChromeOptions()
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    wait = WebDriverWait(driver, 25)
    try:
        driver.get("https://neptunweb.semmelweis.hu/hallgato/login.aspx")
        try: wait.until(EC.element_to_be_clickable((By.ID, "notification-button-accept"))).click()
        except: pass
        wait.until(EC.presence_of_element_located((By.ID, "userName"))).send_keys(conf['neptun']['user_id'])
        driver.find_element(By.ID, "password-form-password").send_keys(conf['neptun']['password'])
        driver.find_element(By.ID, "login-button").click()
        wait.until(EC.visibility_of_element_located((By.ID, "two-factor-qr-code-input-form-input"))).send_keys(get_2fa_code(conf['neptun']['secret']), Keys.ENTER)
        wait.until(EC.element_to_be_clickable((By.ID, "menu-btn"))).click()
        time.sleep(1)
        driver.execute_script("arguments[0].click();", wait.until(EC.element_to_be_clickable((By.ID, "Exams"))))
        time.sleep(1)
        driver.execute_script("arguments[0].click();", wait.until(EC.element_to_be_clickable((By.ID, "ExamRegistration"))))
        wait.until(EC.presence_of_element_located((By.TAG_NAME, "neptun-secondary-title")))
        return driver, wait
    except:
        if driver: driver.quit()
        return None, None

def start_monitoring():
    conf = load_config()
    target_code = conf['neptun']['target_subject_code']
    priority_tutors = conf['neptun']['target_tutors_priority']
    min_date_str = conf['neptun']['earliest_date']
    min_date = datetime.strptime(min_date_str, "%Y-%m-%d")
    
    start_msg = (
        "🚀 *Neptun 監視ボットを起動しました*\n\n"
        f"🖥 実行環境: `{platform.system()}`\n"
        f"📚 対象科目: `{target_code}`\n"
        f"📅 対象開始日: `{min_date_str}`\n"
        f"👨‍🏫 優先教官: `{', '.join([t if t else '全教官対象' for t in priority_tutors])}`\n\n"
        "条件に合う試験が見つかり次第、登録を試みます。"
    )
    print(f"監視開始: {target_code}")
    send_telegram(conf, start_msg)

    while True:
        driver, wait = login_and_prepare(conf)
        if not driver:
            time.sleep(60); continue
        
        session_start = datetime.now()
        try:
            while datetime.now() < session_start + timedelta(minutes=conf['monitoring']['session_timeout_minutes']):
                try:
                    # すでに登録済みかチェック
                    page_src = driver.page_source
                    if target_code in page_src and ("Felvéve" in page_src or "Admitted" in page_src):
                        rows = driver.find_elements(By.CSS_SELECTOR, "tr.mat-mdc-row")
                        for r in rows:
                            if target_code in r.text and ("Felvéve" in r.text or "Admitted" in r.text):
                                send_telegram(conf, f"✅ `{target_code}` は既に登録されています。監視を終了します。")
                                return

                    for tutor in priority_tutors:
                        subject_blocks = driver.find_elements(By.TAG_NAME, "neptun-secondary-title")
                        for block in subject_blocks:
                            if target_code in block.text:
                                container = block.find_element(By.XPATH, "./following-sibling::div")
                                rows = container.find_elements(By.CSS_SELECTOR, "tr.mat-mdc-row")
                                for row in rows:
                                    row_text = row.text
                                    if any(x in row_text for x in ["Felvéve", "Admitted", "Betelt", "Full"]): continue
                                    
                                    if tutor == "" or tutor in row_text:
                                        date_cell = row.find_element(By.CSS_SELECTOR, ".cdk-column-fromDate").text
                                        exam_date = parse_neptun_date(date_cell)
                                        if exam_date and exam_date >= min_date:
                                            # 発見報告
                                            found_tutor = tutor if tutor else "指定なし"
                                            send_telegram(conf, f"🎯 *ターゲットを発見しました！*\n\n教官: `{found_tutor}`\n日時: `{date_cell}`\n\n登録ボタンを押します...")
                                            
                                            # ボタンクリック
                                            btn = row.find_element(By.XPATH, ".//button[contains(., 'Felvétel') or contains(., 'Take')]")
                                            driver.execute_script("arguments[0].click();", btn)
                                            try:
                                                confirm = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Jelentkezés') or contains(., 'Take')]")))
                                                confirm.click()
                                            except: pass
                                            
                                            # ポップアップ監視 (10秒)
                                            for _ in range(20):
                                                if "felvétele sikeres" in driver.page_source or "successful" in driver.page_source.lower():
                                                    success_msg = (
                                                        "🎉 *【ミッション完了】試験登録に成功しました！*\n\n"
                                                        f"📖 科目: `{target_code}`\n"
                                                        f"👨‍🏫 教官: `{found_tutor}`\n"
                                                        f"📅 日時: `{date_cell}`\n\n"
                                                        "ボットを安全に停止します。お疲れ様でした！"
                                                    )
                                                    send_telegram(conf, success_msg)
                                                    speak_message("Mission accomplished. Registration is successful.")
                                                    return 
                                                time.sleep(0.5)

                    time.sleep(conf['monitoring']['idle_interval_minutes'] * 60)
                    driver.refresh(); time.sleep(2)
                except StaleElementReferenceException:
                    driver.refresh(); time.sleep(2)
        except Exception as e:
            # 既に成功していないか最後に確認
            if target_code in driver.page_source and ("Felvéve" in driver.page_source or "Admitted" in driver.page_source):
                send_telegram(conf, "🎉 エラーが発生しましたが、登録は正常に完了しているようです。")
                return
            err_msg = f"⚠️ *一時的なエラーが発生しました*\n\n内容: `{str(e)[:100]}...`\n\n自動的に再起動して監視を続行します。"
            send_telegram(conf, err_msg)
            time.sleep(5)
        finally:
            if driver: driver.quit()

if __name__ == "__main__":
    start_monitoring()