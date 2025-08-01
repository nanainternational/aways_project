from flask import Flask, render_template, request
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
import time
import os

app = Flask(__name__)

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/download", methods=["POST"])
def download_excel():
    print("🟢 /download 진입 확인")

    username = request.form.get("username")
    password = request.form.get("password")

    download_path = os.path.join(os.path.expanduser("~"), "Downloads")

    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    prefs = {
        "download.default_directory": download_path,
        "download.prompt_for_download": False,
        "safebrowsing.enabled": True
    }
    chrome_options.add_experimental_option("prefs", prefs)
    driver = webdriver.Chrome(options=chrome_options)
    wait = WebDriverWait(driver, 20)

    try:
        # ✅ 로그인
        driver.get("https://alwayzseller.ilevit.com/login")
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='text']"))).send_keys(username)
        driver.find_element(By.CSS_SELECTOR, "input[type='password']").send_keys(password)
        login_button = driver.find_element(By.XPATH, "//button[contains(., '로그인')]")
        driver.execute_script("arguments[0].click();", login_button)
        time.sleep(3)

        # ✅ 배송관리 페이지 이동
        driver.get("https://alwayzseller.ilevit.com/shippings")
        time.sleep(3)

        # ✅ 택배사 'CJ대한통운' 선택
        select_element = wait.until(EC.presence_of_element_located((By.ID, "shipping_company")))
        select = Select(select_element)
        for idx, opt in enumerate(select.options):
            if "CJ대한통운" in opt.text:
                select.select_by_index(idx)
                break
        time.sleep(1)

        # ✅ preShippingPreExcelOrders 영역 안의 엑셀 버튼만 클릭
        print("🔍 영역 #preShippingPreExcelOrders 안에서 버튼 검색 중...")
        section = wait.until(EC.presence_of_element_located((By.ID, "preShippingPreExcelOrders")))
        buttons = section.find_elements(By.TAG_NAME, "button")
        print(f"🔎 발견된 버튼 수: {len(buttons)}개")

        found = False
        for idx, btn in enumerate(buttons):
            try:
                text = btn.text.strip()
                print(f"👉 버튼 {idx+1} 텍스트: '{text}'")
                if '엑셀 추출하기' in text:
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)
                    time.sleep(0.5)
                    driver.execute_script("arguments[0].click();", btn)
                    print("✅ 엑셀 추출 버튼 클릭 완료")
                    found = True
                    break
            except Exception as e:
                print(f"❌ 버튼 {idx+1} 예외 발생: {e}")

        if not found:
            raise Exception("❌ preShippingPreExcelOrders 영역 안에 '엑셀 추출하기' 버튼을 찾지 못했습니다.")

        return "✅ 완료: 엑셀 추출 버튼 클릭 성공"

    except Exception as e:
        import traceback
        return f"❌ 오류 발생:\n{traceback.format_exc()}"

    finally:
        time.sleep(5)
        driver.quit()

if __name__ == "__main__":
    app.run(debug=False, use_reloader=False)
