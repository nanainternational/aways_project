from flask import Flask, render_template, request, send_file
import pandas as pd
import os
from datetime import datetime

app = Flask(__name__)

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/convert", methods=["POST"])
def convert_excel():
    print("🟢 /convert 진입 확인")

    file = request.files.get("file")
    if not file:
        return "❌ 파일이 전송되지 않았습니다."

    try:
        df = pd.read_excel(file)

        # 전화번호 포맷 정리
        def format_phone(num):
            num = str(num).strip().replace("-", "")
            if not num.startswith("0") and num.isdigit():
                return "0" + num
            return num

        # 품목명 조합: E열 + F열
        def join_name_option(row):
            return f"{str(row[4]).strip()} {str(row[5]).strip()}".strip()

        # CJ 양식으로 재구성
        cj_df = pd.DataFrame({
            "받는분성명": df.iloc[:, 18],       # S열 → CJ A열
            "받는분전화번호": df.iloc[:, 19].apply(format_phone),  # T열 → CJ B열
            "받는분기타연락처": "",            # CJ C열
            "받는분우편번호": df.iloc[:, 15],  # P열 → CJ D열
            "받는분주소(전체, 분할)": df.iloc[:, 14],  # O열 → CJ E열
            "품목명": df.apply(join_name_option, axis=1),  # E, F열 → CJ F열
            "내품수량": df.iloc[:, 6],         # G열 → CJ G열
            "사용안함": "",                   # CJ H열
            "배송메세지1": df.iloc[:, 17]      # R열 → CJ I열
        })

        # 저장
        now = datetime.now().strftime("%Y%m%d_%H%M")
        output_path = os.path.join(os.path.expanduser("~"), "Desktop", f"cj_{now}.xlsx")
        cj_df.to_excel(output_path, index=False)
        return send_file(output_path, as_attachment=True)

    except Exception as e:
        import traceback
        return f"❌ 변환 중 오류:\n{traceback.format_exc()}"


if __name__ == "__main__":
    app.run(debug=False, use_reloader=False)
