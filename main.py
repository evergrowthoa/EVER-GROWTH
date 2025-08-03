import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
from dateutil.parser import parse
from datetime import datetime
from tkinter import messagebox, Tk
import re

# ⛳ GUI 팝업용 Tk 설정
root = Tk()
root.withdraw()

def alert(msg, title="알림"):
    messagebox.showinfo(title, msg)

def error(msg, title="오류"):
    messagebox.showerror(title, msg)

# 🔒 구글 API 인증 정보
CREDENTIALS_FILE = 'numeric-haven-455700-k8-7e15ff3d6313.json'
SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/17qWvyVONniRI758kESiYS680ChnF7RFHAX-iP-FbrVI/edit'

try:
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
    client = gspread.authorize(creds)

    # 📄 시트 불러오기
    sheet = client.open_by_url(SPREADSHEET_URL)
    ws1 = sheet.get_worksheet(0)
    ws2 = sheet.get_worksheet(1)
    ws_log = sheet.worksheet("Log")  # 로그 시트

    # 📥 데이터프레임으로 읽기
    df1 = pd.DataFrame(ws1.get_all_records())
    df2 = pd.DataFrame(ws2.get_all_records())

    # ✅ 시트1 조건 필터
    condition = (
        (df1["브랜드"] == "코웨이") &
        (df1["진행상황"].isin(["계약서", "해피콜", "동의서", "대기"])) &
        (df1["비가망유형"].astype(str).str.strip() != "")
    )

    # 📝 로그 기록 함수
    def write_log(customer_name, v_value, content, note=""):
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ws_log.append_row([now_str, customer_name, v_value, content, note])

    # ▶️ 조건 만족하는 행 반복
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

    # ✅ 완료 팝업
    alert(f"업데이트 완료!\n총 {updated_count}건 변경됨 ✅")

except Exception as e:
    error(f"에러 발생:\n{str(e)}")
