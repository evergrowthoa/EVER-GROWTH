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
import time

<<<<<<< HEAD
# 🔐 인증 및 시트 주소
CREDENTIALS_FILE = 'numeric-haven-455700-k8-ce44177240c2.json'
SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/17qWvyVONniRI758kESiYS680ChnF7RFHAX-iP-FbrVI/edit'

=======
from gspread.exceptions import WorksheetNotFound, APIError
from gspread_formatting import CellFormat, Color, format_cell_range, format_cell_ranges
from gspread.utils import rowcol_to_a1

def _with_retry(fn, *args, **kwargs):
    tries = kwargs.pop("_tries", 5)
    for i in range(tries):
        try:
            return fn(*args, **kwargs)
        except APIError as e:
            status = getattr(getattr(e, "response", None), "status_code", None)
            if status in (429, 503) and i < tries - 1:
                time.sleep(1.5 * (2 ** i))
                continue
            raise

def batch_values_update(spreadsheet, data, value_input_option="RAW", chunk=400):
    for i in range(0, len(data), chunk):
        body = {"valueInputOption": value_input_option, "data": data[i:i+chunk]}
        _with_retry(spreadsheet.values_batch_update, body)

# -------------------------------
# 공통: PyInstaller/로컬 경로 헬퍼
# -------------------------------
def resource_path(relative_path: str) -> str:
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# 🔑 JSON 키
CREDENTIALS_FILE = resource_path('numeric-haven-455700-k8-541f203927de.json')

# 🔗 스프레드시트 URL
SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/17qWvyVONniRI758kESiYS680ChnF7RFHAX-iP-FbrVI/edit'

# 포맷
yellow_fill = CellFormat(backgroundColor=Color(1, 1, 0))

# -------------------------------
# 공용: 로그 유틸
# -------------------------------
def get_or_create_log(sheet):
    try:
        return sheet.worksheet("Log")
    except WorksheetNotFound:
        ws = sheet.add_worksheet(title="Log", rows=1000, cols=10)
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ws.append_row(["timestamp", "customer_name", "v_value", "content", "note"])
        ws.append_row([now_str, "-", "-", "Log sheet auto-created", "-"])
        return ws

def append_log(ws_log, customer_name, v_value, content, note=""):
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ws_log.append_row([now_str, str(customer_name or ""), str(v_value or ""), str(content or ""), str(note or "")])

# -------------------------------
# 유틸: 헤더 기반 열 번호 찾기
# -------------------------------
def get_col_index(ws, header_name):
    headers = ws.row_values(1)
    return headers.index(header_name) + 1

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
>>>>>>> f29392a170b3c55d41d840ed74949ffb129536d1
def run_script():
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
        client = gspread.authorize(creds)

        sheet = client.open_by_url(SPREADSHEET_URL)
        ws1 = sheet.get_worksheet(0)
        ws2 = sheet.get_worksheet(1)
<<<<<<< HEAD
        ws_log = sheet.worksheet("Log")
=======
        ws_log = get_or_create_log(sheet)
>>>>>>> f29392a170b3c55d41d840ed74949ffb129536d1

        df1 = pd.DataFrame(ws1.get_all_records())
        df2 = pd.DataFrame(ws2.get_all_records())

<<<<<<< HEAD
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
=======
        col_customer = get_col_index(ws1, "고객명")
        col_status   = get_col_index(ws1, "진행상황")
        col_note     = get_col_index(ws1, "특이사항")
        col_v        = get_col_index(ws1, "계약번호")

        df1["브랜드"] = df1["브랜드"].astype(str).str.strip()
        df1["진행상황"] = df1["진행상황"].astype(str).str.strip()
        df1["계약번호"] = df1["계약번호"].astype(str)

        df2["주문번호"] = df2["주문번호"].astype(str)
        df2["상태"] = df2["상태"].astype(str).str.strip()
        df2["고객명"] = df2["고객명"].astype(str).str.strip()

        condition = (
            (df1["브랜드"] == "코웨이") &
            (df1["진행상황"].isin(["계약서", "해피콜", "동의서", "대기"])) &
            (df1["계약번호"].astype(str).str.strip() != "")
        )

        updated_count = 0
        for idx, row1 in df1[condition].iterrows():
            v_value = str(row1["계약번호"]).strip()
            v_last4 = ''.join(re.findall(r'\d+', v_value))[-4:] if re.search(r'\d', v_value) else ''
            f_value = str(row1.get("고객명", "")).strip()

            if not v_last4:
                append_log(ws_log, f_value, v_value, "⛔ 계약번호에 숫자 없음")
                continue

            for _, row2 in df2.iterrows():
                b_last4 = row2["주문번호"][-4:]
                상태값 = row2["상태"]
                고객명2 = row2["고객명"]

                if v_last4 == b_last4 and f_value and (f_value in 고객명2):
>>>>>>> f29392a170b3c55d41d840ed74949ffb129536d1
                    if 상태값 in ["신용조사(가완료)", "신용조사", "출고의뢰"]:
                        raw_date = str(row2.get("설치예정일", "")).strip()
                        try:
                            parsed_date = parse(raw_date)
                            l_val = parsed_date.strftime("%m-%d")
<<<<<<< HEAD
                        except:
=======
                        except Exception:
>>>>>>> f29392a170b3c55d41d840ed74949ffb129536d1
                            l_val = raw_date

                        m_val = str(row2.get("배정시간", "")).strip()
                        existing_note = str(row1.get("특이사항", "")).strip()
<<<<<<< HEAD
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

        messagebox.showinfo("완료", f"업데이트 완료!\n총 {updated_count}건 변경됨 ✅")

    except Exception as e:
        messagebox.showerror("에러 발생", str(e))


def open_log_sheet():
    webbrowser.open(SPREADSHEET_URL + "#gid=1347292722")  # Log 시트 gid에 맞게 수정


# 🖥️ GUI 구성
root = Tk()
root.title("EVER-GROWTH 자동화 도구")
root.geometry("300x150")

Button(root, text="▶ 실행", command=run_script, width=20, height=2, bg="lightgreen").pack(pady=10)
Button(root, text="📜 로그 확인", command=open_log_sheet, width=20, height=2, bg="lightblue").pack()

root.mainloop()
=======
                        new_note = f"{l_val} {m_val}".strip()
                        combined_note = f"{new_note} | {existing_note}" if existing_note else new_note

                        ws1.update_cell(idx + 2, col_note, combined_note)
                        ws1.update_cell(idx + 2, col_status, "승인완료")

                        try:
                            format_cell_range(ws1, f'{chr(64+col_note)}{idx + 2}', yellow_fill)
                            format_cell_range(ws1, f'{chr(64+col_status)}{idx + 2}', yellow_fill)
                        except Exception:
                            pass

                        append_log(ws_log, f_value, v_value, "진행상황 → 승인완료, 특이사항 업데이트", combined_note)
                        updated_count += 1
                        break
                    else:
                        append_log(ws_log, f_value, v_value, f"⛔ 상태 불일치: {상태값}")
                        break

        messagebox.showinfo("완료", f"진행상황 업데이트 완료!\n총 {updated_count}건 변경됨 ✅")

    except Exception as e:
        messagebox.showerror("에러 발생", str(e))

# -------------------------------
# 2) 코웨이 설치일 자동입력 (C열)
# -------------------------------
def run_install_date_updater():
    try:
        scope = [
            'https://spreadsheets.google.com/feeds',
            'https://www.googleapis.com/auth/drive'
        ]
        creds = ServiceAccountCredentials.from_json_keyfile_name(
            CREDENTIALS_FILE, scope
        )
        client = gspread.authorize(creds)

        sheet = client.open_by_url(SPREADSHEET_URL)
        ws1 = sheet.get_worksheet(0)
        ws2 = sheet.get_worksheet(1)

        df1 = pd.DataFrame(ws1.get_all_records())
        df2 = pd.DataFrame(ws2.get_all_records())

        col_status   = get_col_index(ws1, "진행상황")
        col_contract = get_col_index(ws1, "계약번호")

        value_updates = []
        format_requests = []

        # ✅ 고객명 정규화 (한글 + 영어 대응)
        def normalize_name(name):
            if not name:
                return ""
            # 한글 + 영어만 남기기
            cleaned = re.sub(r"[^가-힣A-Za-z]", "", str(name))
            return cleaned.upper()

        condition = (
            (df1["브랜드"].astype(str).str.strip() == "코웨이") &
            (df1["진행상황"].astype(str).str.strip() == "승인완료")
        )

        for idx, row in df1[condition].iterrows():
            rownum = idx + 2

            계약번호 = str(row.get("계약번호", "")).strip()
            고객명1_raw = str(row.get("고객명", "")).strip()
            고객명1 = normalize_name(고객명1_raw)

            changed = False

            for row2 in df2.itertuples():
                상태 = str(getattr(row2, "상태", "")).replace(" ", "")

                if "순주문확정" in 상태 or "설치확정" in 상태:
                    주문번호 = str(getattr(row2, "주문번호", "")).strip()
                    고객명2_raw = str(getattr(row2, "고객명", "")).strip()
                    고객명2 = normalize_name(고객명2_raw)

                    # ✅ 고객명 포함 매칭 (양방향)
                    name_match = (
                        고객명1 and 고객명2 and
                        (고객명1 in 고객명2 or 고객명2 in 고객명1)
                    )

                    if (
                        계약번호[-4:] == 주문번호[-4:]
                        and name_match
                    ):
                        raw = str(getattr(row2, "설치예정일", "")).strip()
                        try:
                            parsed = datetime.strptime(raw, "%Y.%m.%d")
                            formatted = parsed.strftime("%y-%m-%d")
                        except:
                            formatted = raw

                        a1 = rowcol_to_a1(rownum, col_status)
                        value_updates.append({
                            "range": f"{ws1.title}!{a1}",
                            "values": [[formatted]]
                        })

                        # 🟨 진행상황 노란색
                        format_requests.append({
                            "repeatCell": {
                                "range": {
                                    "sheetId": ws1.id,
                                    "startRowIndex": rownum - 1,
                                    "endRowIndex": rownum,
                                    "startColumnIndex": col_status - 1,
                                    "endColumnIndex": col_status
                                },
                                "cell": {
                                    "userEnteredFormat": {
                                        "backgroundColor": {
                                            "red": 1,
                                            "green": 1,
                                            "blue": 0.6
                                        }
                                    }
                                },
                                "fields": "userEnteredFormat.backgroundColor"
                            }
                        })

                        changed = True
                        break

            # ❌ 매칭 실패 → 계약번호 빨간색
            if not changed:
                format_requests.append({
                    "repeatCell": {
                        "range": {
                            "sheetId": ws1.id,
                            "startRowIndex": rownum - 1,
                            "endRowIndex": rownum,
                            "startColumnIndex": col_contract - 1,
                            "endColumnIndex": col_contract
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "backgroundColor": {
                                    "red": 1,
                                    "green": 0.6,
                                    "blue": 0.6
                                }
                            }
                        },
                        "fields": "userEnteredFormat.backgroundColor"
                    }
                })

        # ✅ API 호출 최소화 (Quota-safe)
        if value_updates:
            batch_values_update(sheet, value_updates)

        if format_requests:
            sheet.batch_update({"requests": format_requests})

        messagebox.showinfo(
            "완료",
            "설치일 처리 완료\n"
            "✔ 매칭 성공 → 날짜 입력 (노란색)\n"
            "❌ 매칭 실패 → 계약번호 빨간색"
        )

    except Exception as e:
        messagebox.showerror("에러 발생", str(e))

# -------------------------------
# 3) 청호 설치확정일·월 입력
# -------------------------------
def run_chungho_install_date_updater():
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
        client = gspread.authorize(creds)

        sheet = client.open_by_url(SPREADSHEET_URL)
        ws1 = sheet.get_worksheet(0)
        ws3 = sheet.get_worksheet(2)
        ws_log = get_or_create_log(sheet)

        df1 = worksheet_to_dataframe(ws1)
        df3 = worksheet_to_dataframe(ws3)

        if df1.empty or df3.empty:
            messagebox.showinfo("알림", "시트 데이터가 비어있습니다.")
            return

        col_brand  = get_col_index(ws1, "브랜드")
        col_status = get_col_index(ws1, "진행상황")
        col_v      = get_col_index(ws1, "계약번호")
        col_customer = get_col_index(ws1, "고객명")
        col_c      = get_col_index(ws1, "진행상황")  # 설치확정일 들어갈 열
        col_y      = get_col_index(ws1, "정산") if "정산" in ws1.row_values(1) else get_col_index(ws1, "예상수수료")

        condition = (
            df1.iloc[:, col_brand-1].astype(str).str.strip().eq("청호") &
            df1.iloc[:, col_status-1].astype(str).str.strip().eq("승인완료") &
            df1.iloc[:, col_v-1].astype(str).str.strip().ne("")
        )

        updates, fmt_ranges, log_rows = [], [], []
        updated_count = 0

        for idx1, row1 in df1[condition].iterrows():
            v_value = str(row1.iloc[col_v-1]).strip()
            v_last4 = re.sub(r'\D', '', v_value)[-4:]
            f_value = str(row1.iloc[col_customer-1]).strip()
            rownum  = idx1 + 2

            if not v_last4:
                log_rows.append([datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                f_value, v_value, "⛔ 계약번호에 숫자 없음", ""])
                continue

            for _, row3 in df3.iterrows():
                b_last4 = str(row3.get("계약번호") or row3.get("주문번호") or "")[-4:]
                c_name  = str(row3.get("고객명") or "").strip()
                n_val   = str(row3.get("진행상태") or row3.get("상태") or "").strip()
                m_val   = str(row3.get("설치예정일") or row3.get("매출일") or "").strip()

                if (v_last4 == b_last4) and f_value and c_name and (c_name in f_value):
                    if n_val == "매출확정":
                        try:
                            dt = datetime.strptime(m_val, "%Y-%m-%d")
                            c_text = dt.strftime("%y-%m-%d")
                            y_val  = str(dt.month)

                            updates.append({"range": f"{ws1.title}!{chr(64+col_c)}{rownum}:{chr(64+col_c)}{rownum}", "values": [[c_text]]})
                            updates.append({"range": f"{ws1.title}!{chr(64+col_y)}{rownum}:{chr(64+col_y)}{rownum}", "values": [[y_val]]})

                            fmt_ranges.append((f"{chr(64+col_c)}{rownum}", yellow_fill))
                            fmt_ranges.append((f"{chr(64+col_y)}{rownum}", yellow_fill))

                            log_rows.append([datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                            f_value, v_value, "청호 설치확정일 및 정산월 입력", f"설치확정일={c_text}, 정산월={y_val}"])
                            updated_count += 1
                        except Exception:
                            log_rows.append([datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                            f_value, v_value, "⛔ 날짜 변환 오류", f"raw={m_val}"])
                        break
                    else:
                        log_rows.append([datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                        f_value, v_value, f"⛔ 상태 불일치: {n_val}", ""])
                        break

        if updates:
            batch_values_update(sheet, updates)
        if fmt_ranges:
            _with_retry(format_cell_ranges, ws1, fmt_ranges)
        if log_rows:
            try:
                _with_retry(ws_log.append_rows, log_rows, value_input_option="RAW")
            except AttributeError:
                _with_retry(sheet.values_append,
                            'Log!A1',
                            {'valueInputOption': 'RAW', 'insertDataOption': 'INSERT_ROWS'},
                            {'values': log_rows})

        messagebox.showinfo("완료", f"청호: 총 {updated_count}건 변경 완료 ✅")

    except Exception as e:
        messagebox.showerror("에러 발생", str(e))

# -------------------------------
# 청호 진행상황 업데이트
# -------------------------------
def worksheet_to_dataframe(ws):
    data = ws.get_all_values()
    df = pd.DataFrame(data[1:], columns=data[0])
    return df

def get_or_create_log(sheet):
    try:
        ws_log = sheet.worksheet("Log")
    except:
        ws_log = sheet.add_worksheet("Log", rows=1000, cols=10)
        ws_log.append_row(["시간", "고객명", "계약번호", "내용"])
    return ws_log


def run_script_cheongho():
    try:
        # ==========================
        # 시트 불러오기
        # ==========================
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
        client = gspread.authorize(creds)

        sheet = client.open_by_url(SPREADSHEET_URL)
        ws1 = sheet.get_worksheet(0)   # 메인 시트
        ws3 = sheet.get_worksheet(2)   # 청호 설치 시트
        ws_log = get_or_create_log(sheet)

        df1 = worksheet_to_dataframe(ws1)
        df3 = worksheet_to_dataframe(ws3)

        updated_rows = []

        # ==========================
        # 조건 ① 브랜드, 진행상황, 계약번호
        # ==========================
        condition = (
            (df1["브랜드"] == "청호") &
            (df1["진행상황"].isin(["계약서", "해피콜", "대기"])) &
            (df1["계약번호"].astype(str).str.strip() != "")
        )

        for idx, row in df1[condition].iterrows():
            계약번호 = str(row["계약번호"]).strip()
            고객명 = str(row["고객명"]).strip()
            if 계약번호 == "" or 고객명 == "":
                continue

            계약번호_뒤4 = 계약번호[-4:]

    # ✅ 루프 돌면서 df1 안에 df3 고객명이 포함되는지 확인
            for _, row3 in df3.iterrows():
                if str(row3["계약번호"]).strip()[-4:] == 계약번호_뒤4:
                    고객명3 = str(row3["고객명"]).strip()
                    if 고객명3 and (고객명.find(고객명3) != -1):   # df1 안에 df3이 포함되어 있으면 매칭
                        진행상태 = str(row3.get("진행상태", "")).strip()
                        설치예정일값 = str(row3.get("설치예정일", "")).strip()

                        if 진행상태 in ["출고의뢰", "출고확정"]:
                            try:
                                설치예정일 = pd.to_datetime(설치예정일값, errors="coerce")
                                설치예정일_str = 설치예정일.strftime("%m-%d 설치예정//") if pd.notna(설치예정일) else 설치예정일값
                            except Exception:
                                설치예정일_str = 설치예정일값

                            기존특이 = str(row.get("특이사항", "")).strip()
                            새로운특이 = f"{설치예정일_str} {기존특이}".strip()

                            df1.at[idx, "진행상황"] = "승인완료"
                            df1.at[idx, "특이사항"] = 새로운특이

                            rownum = idx + 2
                            highlight = CellFormat(backgroundColor=Color(1, 1, 0))
                            format_cell_range(ws1, f"A{rownum}:Z{rownum}", highlight)

                            updated_rows.append([
                                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                row["고객명"], 계약번호,
                                f"승인완료 / 특이사항: {새로운특이}"
                            ])
                        break


        # ==========================
        # 시트 반영
        # ==========================
        ws1.update([df1.columns.values.tolist()] + df1.values.tolist())

        if updated_rows:
            ws_log.append_rows(updated_rows)

        messagebox.showinfo("완료", f"청호 진행상황 업데이트 완료!\n총 {len(updated_rows)}건 변경됨")

    except Exception as e:
        messagebox.showerror("오류 발생", str(e))


# -------------------------------
# 4) 로그 시트 열기
# -------------------------------
def open_log_sheet():
    webbrowser.open(SPREADSHEET_URL + "#gid=1347292722")

# -------------------------------
# GUI
# -------------------------------
if __name__ == "__main__":
    root = Tk()
    root.title("EVER-GROWTH 자동화 도구")
    root.geometry("360x340")

    Button(root, text="코웨이 진행상황 업데이트",   command=run_script,                      width=34, height=2, bg="lightgreen").pack(pady=8)
    Button(root, text="코웨이 설치확정일 자동입력", command=run_install_date_updater,      width=34, height=2, bg="lightyellow").pack(pady=8)
    Button(root, text="청호 진행상황 업데이트", command=run_script_cheongho,      width=34, height=2, bg="orange").pack(pady=8)
    Button(root, text="청호 설치확정일·정산월 입력 ", command=run_chungho_install_date_updater, width=34, height=2, bg="khaki").pack(pady=8)
    Button(root, text="수정 로그 확인",             command=open_log_sheet,                 width=34, height=2, bg="lightblue").pack(pady=8)

    root.mainloop()

#확장프로그램 목록
#dracula-theme.theme-dracula
#github.copilot
#github.copilot-chat
#ms-ceintl.vscode-language-pack-ko
#ms-python.debugpy
#ms-python.python
#ms-python.vscode-pylance
#oderwat.indent-rainbow
#pkief.material-icon-theme
#usernamehw.errorlens


# 한 번만 설정
#git config --global user.name "본인 이름"
#git config --global user.email "가입한 GitHub 이메일"

# GitHub 저장소 연결 (최초 1번만)
#git remote add origin https://github.com/계정명/저장소명.git

# 이후 매번
#git add .
#git commit -m "변경내용 설명"
#git push origin main


# github 지금 거 날리고 최신본만 가져오는것
# git reset --hard HEAD
# git pull origin main

# EXE파일 만드는 bash
# pyinstaller --onefile --noconsole --add-data "numeric-haven-455700-k8-541f203927de.json;." main.py
>>>>>>> f29392a170b3c55d41d840ed74949ffb129536d1
