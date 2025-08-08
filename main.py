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
from gspread_formatting import CellFormat, Color, format_cell_range

# -------------------------------
# 공통: PyInstaller/로컬 경로 헬퍼
# -------------------------------
def resource_path(relative_path: str) -> str:
    try:
        base_path = sys._MEIPASS  # PyInstaller 임시 폴더
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# 🔑 JSON 키 (파일명 맞춰 수정)
CREDENTIALS_FILE = resource_path('numeric-haven-455700-k8-541f203927de.json')

# 🔗 스프레드시트 URL
SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/17qWvyVONniRI758kESiYS680ChnF7RFHAX-iP-FbrVI/edit'

# 포맷
yellow_fill = CellFormat(backgroundColor=Color(1, 1, 0))

# -------------------------------
# 유틸: 헤더 문제 우회용 로더(청호용에서 사용)
# -------------------------------
def make_unique_headers_from_row(row, width=None):
    if width is None:
        width = len(row or [])
    row = (row or []) + [""] * (width - len(row or []))
    seen, out = {}, []
    for i, h in enumerate(row):
        h = (h or "").strip() or f"col{i+1}"
        if h in seen:
            seen[h] += 1
            h = f"{h}_{seen[h]}"
        else:
            seen[h] = 0
        out.append(h)
    return out

def worksheet_to_dataframe(ws):
    values = ws.get_all_values()
    if not values:
        return pd.DataFrame()
    max_w = max(len(r) for r in values)
    padded = [r + [""] * (max_w - len(r)) for r in values]
    headers = make_unique_headers_from_row(padded[0], width=max_w)
    return pd.DataFrame(padded[1:], columns=headers)

# -------------------------------
# 1) 코웨이 진행상황 업데이트
# -------------------------------
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

        # 공백 제거해 비교
        df1["브랜드"] = df1["브랜드"].astype(str).str.strip()
        df1["진행상황"] = df1["진행상황"].astype(str).str.strip()
        df1["비가망유형"] = df1["비가망유형"].astype(str)

        df2["주문번호"] = df2["주문번호"].astype(str)
        df2["상태"] = df2["상태"].astype(str).str.strip()
        df2["고객명"] = df2["고객명"].astype(str).str.strip()

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
            f_value = str(row1.get("고객명", "")).strip()

            if not v_last4:
                write_log(f_value, v_value, "⛔ 비가망유형에 숫자 없음")
                continue

            for _, row2 in df2.iterrows():
                b_last4 = row2["주문번호"][-4:]
                상태값 = row2["상태"]
                고객명2 = row2["고객명"]

                if v_last4 == b_last4 and f_value and (f_value in 고객명2):
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

                        # 값 쓰기
                        ws1.update_cell(idx + 2, 16, combined_note)  # P
                        ws1.update_cell(idx + 2, 3, "승인완료")      # C

                        # 값 썼으면 무조건 칠하기
                        try:
                            format_cell_range(ws1, f'P{idx + 2}', yellow_fill)
                            format_cell_range(ws1, f'C{idx + 2}', yellow_fill)
                        except Exception:
                            pass

                        write_log(f_value, v_value, "진행상황 → 승인완료, 특이사항 업데이트", combined_note)
                        updated_count += 1
                        break
                    else:
                        write_log(f_value, v_value, f"⛔ 상태 불일치: {상태값}")
                        break

        messagebox.showinfo("완료", f"진행상황 업데이트 완료!\n총 {updated_count}건 변경됨 ✅")

    except Exception as e:
        messagebox.showerror("에러 발생", str(e))

# -------------------------------
# 2) 코웨이 설치일 자동입력 (C열)
# -------------------------------
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

        # 공백 제거/형변환
        for col in ["진행상황", "브랜드", "비가망유형", "고객명"]:
            if col in df1.columns:
                df1[col] = df1[col].astype(str).str.strip()

        for col in ["주문번호", "고객명", "상태", "설치예정일"]:
            if col in df2.columns:
                df2[col] = df2[col].astype(str).str.strip()

        condition = (df1["진행상황"] == "승인완료") & (df1["브랜드"] == "코웨이")

        updated = 0
        for idx, row in df1[condition].iterrows():
            v_value = str(row.get("비가망유형", "")).strip()
            v_last4 = ''.join(re.findall(r'\d+', v_value))[-4:] if v_value else ''
            customer1 = str(row.get("고객명", "")).strip()

            if not v_last4:
                continue

            for row2 in df2.itertuples():
                b_last4 = str(getattr(row2, "주문번호", ""))[-4:]
                customer2 = str(getattr(row2, "고객명", "")).strip()
                status = str(getattr(row2, "상태", "")).strip()

                if v_last4 == b_last4 and customer1 and (customer1 in customer2) and status == "순주문확정":
                    raw_date = str(getattr(row2, "설치예정일", "")).strip()
                    try:
                        formatted_date = datetime.strptime(raw_date, "%Y.%m.%d").strftime("%y-%m-%d")
                        ws1.update_cell(idx + 2, 3, formatted_date)  # C열
                        try:
                            format_cell_range(ws1, f'C{idx + 2}', yellow_fill)
                        except Exception:
                            pass
                        updated += 1
                    except:
                        pass
                    break

        messagebox.showinfo("완료", f"설치일 입력 완료!\n총 {updated}건 변경됨 ✅")

    except Exception as e:
        messagebox.showerror("에러 발생", str(e))

# -------------------------------
# 3) 청호 설치확정일·월 입력 (C열=YY-MM-DD, Y열=MM)
#    헤더 중복/빈칸 대응 + 공백 제거 비교 + 하이라이트 보장
# -------------------------------
def run_chungho_install_date_updater():
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
        client = gspread.authorize(creds)

        sheet = client.open_by_url(SPREADSHEET_URL)
        ws1 = sheet.get_worksheet(0)      # 시트1
        ws3 = sheet.get_worksheet(2)      # 시트3

        df1 = worksheet_to_dataframe(ws1) # 안전 로딩
        df3 = worksheet_to_dataframe(ws3)

        if df1.empty or df3.empty:
            messagebox.showinfo("알림", "시트 데이터가 비어있습니다.")
            return

        # ✅ 공백 제거한 조건 (26~206행 스킵 방지)
        condition = (
            df1.iloc[:, 7].astype(str).str.strip().eq("청호") &        # H열
            df1.iloc[:, 2].astype(str).str.strip().eq("승인완료") &    # C열
            df1.iloc[:, 21].astype(str).str.strip().ne("")             # V열
        )

        def pick_col(df, candidates):
            for name in candidates:
                if name in df.columns:
                    return name
            return None

        col_contract = pick_col(df3, ["계약번호", "주문번호", "B", "col2"])
        col_customer = pick_col(df3, ["고객명", "성명", "C", "col3"])
        col_status   = pick_col(df3, ["진행상태", "상태", "N", "col14"])
        col_m_date   = pick_col(df3, ["M열", "설치예정일", "매출일", "M", "col13"])

        if not all([col_contract, col_customer, col_status, col_m_date]):
            messagebox.showerror("에러", "시트3 열을 찾지 못했습니다.")
            return

        updated_count = 0

        for idx1, row1 in df1[condition].iterrows():
            v_value = str(row1.iloc[21]).strip()   # V열
            v_last4 = re.sub(r'\D', '', v_value)[-4:]
            f_value = str(row1.iloc[5]).strip()    # F열(고객명)

            if not v_last4:
                continue

            for _, row3 in df3.iterrows():
                b_last4 = str(row3[col_contract])[-4:]
                c_value_sheet3 = str(row3[col_customer]).strip()
                n_value = str(row3[col_status]).strip()
                m_value = str(row3[col_m_date]).strip()  # YYYY-MM-DD 기대

                if (v_last4 == b_last4) and (f_value and f_value in c_value_sheet3) and (n_value == "매출확정"):
                    try:
                        dt = datetime.strptime(m_value, "%Y-%m-%d")
                        formatted_c = dt.strftime("%y-%m-%d")  # C열(3)
                        month_only  = dt.strftime("%m")         # Y열(25)

                        ws1.update_cell(idx1 + 2, 3,  formatted_c)  # C
                        ws1.update_cell(idx1 + 2, 25, month_only)   # Y

                        # 값 썼으면 무조건 칠하기
                        try:
                            format_cell_range(ws1, f"C{idx1+2}", yellow_fill)
                            format_cell_range(ws1, f"Y{idx1+2}", yellow_fill)
                        except Exception:
                            pass

                        updated_count += 1
                    except Exception as e:
                        print(f"날짜 변환 오류: {m_value} -> {e}")
                    break

        messagebox.showinfo("완료", f"청호: 총 {updated_count}건 변경 완료 ✅")

    except Exception as e:
        messagebox.showerror("에러 발생", str(e))

# -------------------------------
# 4) 로그 시트 열기
# -------------------------------
def open_log_sheet():
    webbrowser.open(SPREADSHEET_URL + "#gid=1347292722")

# -------------------------------
# GUI
# -------------------------------
root = Tk()
root.title("EVER-GROWTH 자동화 도구")
root.geometry("360x340")

Button(root, text="코웨이 진행상황 업데이트",   command=run_script,                      width=34, height=2, bg="lightgreen").pack(pady=8)
Button(root, text="코웨이 설치일 자동입력 (C열)", command=run_install_date_updater,      width=34, height=2, bg="lightyellow").pack(pady=8)
Button(root, text="청호 설치확정일·월 입력 (C/Y)", command=run_chungho_install_date_updater, width=34, height=2, bg="khaki").pack(pady=8)
Button(root, text="수정 로그 확인",             command=open_log_sheet,                 width=34, height=2, bg="lightblue").pack(pady=8)

root.mainloop()
