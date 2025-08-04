import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
from dateutil.parser import parse
from datetime import datetime
from tkinter import messagebox, Tk, Button
import re
import webbrowser
import os

base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
json_file_path = os.path.join(base_path, 'numeric-haven-455700-k8-ce44177240c2.json')

# 인증 정보 로딩
creds = ServiceAccountCredentials.from_json_keyfile_name(json_file_path, scope)

# GUI 생성
root = Tk()
root.title("Google Sheets 자동화")
root.geometry("300x150")

# 구글 시트 URL
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/17qWvyVONniRI758kESiYS680ChnF7RFHAX-iP-FbrVI/edit#gid=0"
LOG_SHEET_GID = "1370173144"  # Log 시트 gid 넣기

# 인증 파일
CREDENTIALS_FILE = 'numeric-haven-455700-k8-ce44177240c2.json'


# ▶️ 작업 실행 함수
def run_task():
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
        client = gspread.authorize(creds)

        sheet = client.open_by_url(SPREADSHEET_URL)
        ws1 = sheet.get_worksheet(0)
        ws2 = sheet.get_worksheet(1)
        ws_log = sheet.worksheet("Log")

        df1 = pd.DataFrame(ws1.get_all_records())
        df2 = pd.DataFrame(ws2.get_all_records())

        condition = (
            (df1["브랜드"] == "코웨이") &
            (df1["진행상황"].isin(["계약서", "해피콜", "동의서", "대기"])) &
            (df1["비가망유형"].astype(str).str.strip() != "")
        )

        def write_log(customer_name, v_value, content, note=""):
            now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            ws_log.append_row([now_str, customer_name, v_value, content, note])

        updated_count = 0
        for idx, row1 in df1[condition].iterrows():
            v_value = str(row1["비가망유형"]).strip()
            v_last4 = ''.join(re.findall(r'\d+', v_value))[-4:] if re.search(r'\d', v_value) else ''
            f_value = str(row1["고객명"]).strip()

            if not v_last4:
                write_log(f_value, v_value, "⛔ 비가망유형에 숫자 없음", "")
                continue

            for _, row2 in df2.iterrows():
                b_last4 = str(row2.get("주문번호", ""))[-4:]
                상태값 = str(row2.get("상태", "")).strip()
                고객명2 = str(row2.get("고객명", "")).strip()

                if v_last4 == b_last4 and f_value in 고객명2:
                    if 상태값 in ["신용조사(가완료)", "신용조사", "출고의뢰"]:
                        raw_date = str(row2.get("설치예정일", "")).strip()
                        try:
                            parsed_date = parse(raw_date)
                            l_val = parsed_date.strftime("%m-%d")
                        except:
                            l_val = raw_date

                        m_val = str(row2.get("배정시간", "")).strip()
                        existing_note = str(row1.get("특이사항", "")).strip()
                        new_note = f"{l_val} {m_val}"
                        combined_note = f"{new_note} | {existing_note}" if existing_note else new_note

                        ws1.update_cell(idx + 2, 16, combined_note)  # P열
                        ws1.update_cell(idx + 2, 3, "승인완료")       # C열
                        write_log(f_value, v_value, "진행상황 → 승인완료, 특이사항 업데이트", combined_note)
                        updated_count += 1
                        break
                    else:
                        write_log(f_value, v_value, f"⛔ 상태 불일치: {상태값}", "")
                        break

        messagebox.showinfo("완료", f"업데이트 완료!\n총 {updated_count}건 변경됨 ✅")

    except Exception as e:
        messagebox.showerror("에러", str(e))


# 🧾 로그 확인 버튼
def open_log_sheet():
    log_url = f"https://docs.google.com/spreadsheets/d/17qWvyVONniRI758kESiYS680ChnF7RFHAX-iP-FbrVI/edit#gid={LOG_SHEET_GID}"
    webbrowser.open(log_url)


# 🔘 버튼 UI 배치
btn1 = Button(root, text="실행", command=run_task, width=20, height=2)
btn1.pack(pady=10)

btn2 = Button(root, text="Log 시트 열기", command=open_log_sheet, width=20, height=2)
btn2.pack(pady=5)

# GUI 실행
root.mainloop()
