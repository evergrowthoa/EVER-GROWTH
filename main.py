import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
from dateutil.parser import parse
from datetime import datetime
from tkinter import messagebox, Tk, Button
import webbrowser
import re
import os
import sys

if getattr(sys, 'frozen', False):
    # PyInstaller로 빌드된 실행파일일 경우
    BASE_DIR = sys._MEIPASS
else:
    # Python 스크립트로 실행할 경우
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

CREDENTIALS_FILE = os.path.join(BASE_DIR, 'numeric-haven-455700-k8-541f203927de.json')
SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/17qWvyVONniRI758kESiYS680ChnF7RFHAX-iP-FbrVI/edit'

# 🔁 기존 기능: 진행상황 자동 업데이트
def run_script():
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
                write_log(f_value, v_value, "⛔ 비가망유형에 숫자 없음")
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
                        write_log(f_value, v_value, f"⛔ 상태 불일치: {상태값}")
                        break

        messagebox.showinfo("완료", f"진행상황 업데이트 완료!\n총 {updated_count}건 변경됨 ✅")

    except Exception as e:
        messagebox.showerror("에러 발생", str(e))


# 🆕 두 번째 기능: 설치일 입력 (설치확정일 시트1에 기록)
def run_install_date_updater():
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
        client = gspread.authorize(creds)

        sheet = client.open_by_url(SPREADSHEET_URL)
        ws1 = sheet.get_worksheet(0)
        ws2 = sheet.get_worksheet(1)

        df1 = pd.DataFrame(ws1.get_all_records())
        df2 = pd.DataFrame(ws2.get_all_records())

        condition = (df1["진행상황"] == "승인완료") & (df1["브랜드"] == "코웨이")

        updated = 0
        for idx, row in df1[condition].iterrows():
            v_value = str(row["비가망유형"]).strip()
            v_last4 = ''.join(re.findall(r'\d+', v_value))[-4:] if re.search(r'\d', v_value) else ''
            customer1 = str(row["고객명"]).strip()

            if not v_last4:
                continue

            for row2 in df2.itertuples():
                b_last4 = str(getattr(row2, "주문번호"))[-4:]
                customer2 = str(getattr(row2, "고객명")).strip()
                status = str(getattr(row2, "상태")).strip()

                if (
                    v_last4 == b_last4 and
                    customer1 in customer2 and
                    status == "순주문확정"
                ):
                    raw_date = str(getattr(row2, "설치예정일")).strip()
                    try:
                        formatted_date = datetime.strptime(raw_date, "%Y.%m.%d").strftime("%y-%m-%d")
                        ws1.update_cell(idx + 2, 3, formatted_date)  # C열
                        updated += 1
                    except:
                        pass
                    break

        messagebox.showinfo("완료", f"설치일 입력 완료!\n총 {updated}건 변경됨 ✅")

    except Exception as e:
        messagebox.showerror("에러 발생", str(e))


# 🌐 로그 시트 열기
def open_log_sheet():
    webbrowser.open(SPREADSHEET_URL + "#gid=1347292722")


# 🖥️ GUI 구성
root = Tk()
root.title("EVER-GROWTH 자동화 도구")
root.geometry("300x220")

Button(root, text="코웨이 진행상황 업데이트", command=run_script, width=30, height=2, bg="lightgreen").pack(pady=10)
Button(root, text="코웨이 설치일 자동입력", command=run_install_date_updater, width=30, height=2, bg="lightyellow").pack(pady=10)
Button(root, text="수정 로그 확인", command=open_log_sheet, width=30, height=2, bg="lightblue").pack(pady=10)

root.mainloop()
