from flask import Flask, render_template, request, send_file
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
import time
import os
import pandas as pd
from datetime import datetime

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
        driver.get("https://alwayzseller.ilevit.com/login")
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='text']"))).send_keys(username)
        driver.find_element(By.CSS_SELECTOR, "input[type='password']").send_keys(password)
        login_button = driver.find_element(By.XPATH, "//button[contains(., '로그인')]")
        driver.execute_script("arguments[0].click();", login_button)
        time.sleep(3)

        driver.get("https://alwayzseller.ilevit.com/shippings")
        time.sleep(3)

        select_element = wait.until(EC.presence_of_element_located((By.ID, "shipping_company")))
        select = Select(select_element)
        for idx, opt in enumerate(select.options):
            if "CJ대한통운" in opt.text:
                select.select_by_index(idx)
                break
        time.sleep(1)

        section = wait.until(EC.presence_of_element_located((By.ID, "preShippingPreExcelOrders")))
        buttons = section.find_elements(By.TAG_NAME, "button")
        found = False
        for btn in buttons:
            try:
                if '엑셀 추출하기' in btn.text.strip():
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)
                    time.sleep(0.5)
                    driver.execute_script("arguments[0].click();", btn)
                    print("✅ 엑셀 추출 버튼 클릭 완료")
                    found = True
                    break
            except Exception as e:
                print(f"❌ 버튼 처리 예외: {e}")

        if not found:
            raise Exception("❌ '엑셀 추출하기' 버튼을 찾지 못했습니다.")

        return "✅ 엑셀 다운로드 완료. '/convert'에서 변환해주세요."

    except Exception as e:
        import traceback
        return f"❌ 오류 발생:\n{traceback.format_exc()}"

    finally:
        time.sleep(5)
        driver.quit()


@app.route("/convert", methods=["POST"])
def convert_excel():
    print("🟢 /convert 진입 확인")

    file = request.files.get("file")
    if not file:
        return "❌ 파일이 전송되지 않았습니다."

    try:
        df = pd.read_excel(file)

        def format_phone(num):
            num = str(num).strip().replace("-", "")
            if not num.startswith("0") and num.isdigit():
                return "0" + num
            return num

        def join_name_option(row):
            return f"{str(row[4]).strip()} {str(row[5]).strip()}".strip()  # 상품명 + 옵션

        cj_df = pd.DataFrame({
            "받는분성명": df.iloc[:, 18],            # S열
            "받는분전화번호": df.iloc[:, 19].apply(format_phone),  # T열
            "받는분기타연락처": "",                  # 비움
            "받는분우편번호": df.iloc[:, 15],        # P열
            "받는분주소(전체, 분할)": df.iloc[:, 14], # O열
            "품목명": df.apply(join_name_option, axis=1),  # E + F열
            "내품수량": df.iloc[:, 6],               # G열
            "사용안함": "",                          # 비움
            "배송메세지1": df.iloc[:, 17]            # R열
        })

        now = datetime.now().strftime("%Y%m%d_%H%M")
        output_path = os.path.join(os.path.expanduser("~"), "Desktop", f"cj_{now}.xlsx")
        cj_df.to_excel(output_path, index=False)
        return send_file(output_path, as_attachment=True)

    except Exception as e:
        import traceback
        return f"❌ 변환 중 오류:\n{traceback.format_exc()}"






if __name__ == "__main__":
    app.run(debug=False, use_reloader=False)
