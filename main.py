import gspread
from oauth2client.service_account import ServiceAccountCredentials
from dateutil.parser import parse
from datetime import datetime
from tkinter import messagebox, Tk, Button
import webbrowser
import re
import os
import sys
import time

from tkinter import *
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

def resource_path(relative_path: str) -> str:
    base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

CREDENTIALS_FILE = resource_path('evergrowth-493504-abdee694f352.json')
SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/17qWvyVONniRI758kESiYS680ChnF7RFHAX-iP-FbrVI/edit'

yellow_fill = CellFormat(backgroundColor=Color(1, 1, 0))

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

def worksheet_to_table(ws):
    values = ws.get_all_values()
    if not values:
        return [], [], []
    max_w = max(len(r) for r in values)
    padded = [r + [""] * (max_w - len(r)) for r in values]
    headers = make_unique_headers_from_row(padded[0], width=max_w)
    rows = padded[1:]
    records = []
    for r in rows:
        rec = {}
        for i, h in enumerate(headers):
            rec[h] = r[i] if i < len(r) else ""
        records.append(rec)
    return headers, rows, records

def safe_strip(v):
    return str(v or "").strip()

def digits_last4_keep0(v):
    digits = re.sub(r"\D", "", str(v or ""))
    if not digits:
        return ""
    return digits[-4:].zfill(4)

def normalize_name_upper(v):
    if not v:
        return ""
    return re.sub(r"[^가-힣A-Za-z]", "", str(v)).upper()

def normalize_date_yy_mm_dd(val):
    if not val:
        return ""
    try:
        return parse(str(val)).strftime("%y-%m-%d")
    except:
        return str(val).strip()

def run_script():
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
        client = gspread.authorize(creds)

        sheet = client.open_by_url(SPREADSHEET_URL)
        ws1 = sheet.get_worksheet(0)
        ws2 = sheet.get_worksheet(1)
        ws_log = get_or_create_log(sheet)

        ws1_records = ws1.get_all_records()
        ws2_records = ws2.get_all_records()

        col_customer = get_col_index(ws1, "고객명")
        col_status   = get_col_index(ws1, "진행상황")
        col_note     = get_col_index(ws1, "특이사항")
        col_v        = get_col_index(ws1, "계약번호")

        updated_count = 0

        for idx, row1 in enumerate(ws1_records):
            brand = safe_strip(row1.get("브랜드"))
            status = safe_strip(row1.get("진행상황"))
            contract_no = str(row1.get("계약번호", ""))
            contract_no_stripped = contract_no.strip()

            if not (brand == "코웨이" and status in ["계약서", "해피콜", "동의서", "대기"] and contract_no_stripped != ""):
                continue

            v_value = contract_no_stripped
            digits = re.findall(r"\d+", v_value)
            v_last4 = "".join(digits)[-4:] if re.search(r"\d", v_value) else ""
            f_value = safe_strip(row1.get("고객명", ""))

            if not v_last4:
                append_log(ws_log, f_value, v_value, "⛔ 계약번호에 숫자 없음")
                continue

            matched_any = False

            for row2 in ws2_records:
                order_no = str(row2.get("주문번호", ""))
                b_last4 = order_no[-4:] if order_no else ""
                상태값 = safe_strip(row2.get("상태"))
                고객명2 = safe_strip(row2.get("고객명"))

                if v_last4 == b_last4 and f_value and 고객명2 and (f_value in 고객명2):
                    matched_any = True
                    if 상태값 in ["신용조사(가완료)", "신용조사", "출고의뢰"]:
                        raw_date = safe_strip(row2.get("설치예정일", ""))
                        try:
                            parsed_date = parse(raw_date)
                            l_val = parsed_date.strftime("%m-%d")
                        except Exception:
                            l_val = raw_date

                        m_val = safe_strip(row2.get("배정시간", ""))
                        existing_note = safe_strip(row1.get("특이사항", ""))
                        new_note = f"{l_val} {m_val}".strip()
                        combined_note = f"{new_note} | {existing_note}" if existing_note else new_note

                        rownum = idx + 2
                        ws1.update_cell(rownum, col_note, combined_note)
                        ws1.update_cell(rownum, col_status, "승인완료")

                        try:
                            format_cell_range(ws1, f'{chr(64+col_note)}{rownum}', yellow_fill)
                            format_cell_range(ws1, f'{chr(64+col_status)}{rownum}', yellow_fill)
                        except Exception:
                            pass

                        append_log(ws_log, f_value, v_value, "진행상황 → 승인완료, 특이사항 업데이트", combined_note)
                        updated_count += 1
                    else:
                        append_log(ws_log, f_value, v_value, f"⛔ 상태 불일치: {상태값}")
                    break

            if not matched_any:
                append_log(ws_log, f_value, v_value, "⛔ 주문번호 매칭 실패")

        messagebox.showinfo("완료", f"진행상황 업데이트 완료!\n총 {updated_count}건 변경됨 ✅")

    except Exception as e:
        messagebox.showerror("에러 발생", str(e))

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
        ws_log = get_or_create_log(sheet)

        ws1_records = ws1.get_all_records()
        ws2_records = ws2.get_all_records()

        col_status = get_col_index(ws1, "진행상황")
        col_expect = get_col_index(ws1, "예상설치일")
        total_cols = len(ws1.row_values(1))

        value_updates = []
        format_requests = []

        for idx, row in enumerate(ws1_records):
            brand = safe_strip(row.get("브랜드"))
            if brand != "코웨이":
                continue

            rownum = idx + 2
            계약번호_4 = digits_last4_keep0(row.get("계약번호", ""))
            고객명1 = normalize_name_upper(row.get("고객명", ""))

            matched = False

            for row2 in ws2_records:
                상태 = safe_strip(row2.get("상태", "")).replace(" ", "")
                if "순주문확정" in 상태 or "설치확정" in 상태:
                    주문번호_4 = digits_last4_keep0(row2.get("주문번호", ""))
                    고객명2 = normalize_name_upper(row2.get("고객명", ""))

                    if (
                        계약번호_4
                        and 계약번호_4 == 주문번호_4
                        and 고객명1 and 고객명2
                        and (고객명1 in 고객명2 or 고객명2 in 고객명1)
                    ):
                        sheet2_date = normalize_date_yy_mm_dd(row2.get("설치예정일", ""))

                        a1 = rowcol_to_a1(rownum, col_status)
                        value_updates.append({
                            "range": f"{ws1.title}!{a1}",
                            "values": [[sheet2_date]]
                        })

                        append_log(
                            ws_log,
                            row.get("고객명"),
                            row.get("계약번호"),
                            "설치일 입력",
                            f"last4={계약번호_4}, 진행상황→{sheet2_date}"
                        )

                        matched = True
                        break

            if not matched:
                append_log(
                    ws_log,
                    row.get("고객명"),
                    row.get("계약번호"),
                    "⛔ 설치일 매칭 실패",
                    f"last4={계약번호_4}"
                )

        if value_updates:
            batch_values_update(sheet, value_updates)

        ws1_records_after = ws1.get_all_records()

        for idx, row in enumerate(ws1_records_after):
            rownum = idx + 2

            브랜드 = safe_strip(row.get("브랜드", ""))
            진행상황 = normalize_date_yy_mm_dd(row.get("진행상황", ""))
            예상설치일 = normalize_date_yy_mm_dd(row.get("예상설치일", ""))
            상태원본 = safe_strip(row.get("진행상황", ""))

            if (
                브랜드 == "코웨이"
                and 상태원본 == "승인완료"
                and 진행상황 != 예상설치일
            ):
                format_requests.append({
                    "repeatCell": {
                        "range": {
                            "sheetId": ws1.id,
                            "startRowIndex": rownum - 1,
                            "endRowIndex": rownum,
                            "startColumnIndex": 0,
                            "endColumnIndex": total_cols
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "backgroundColor": {
                                    "red": 1,
                                    "green": 0.8,
                                    "blue": 0.8
                                }
                            }
                        },
                        "fields": "userEnteredFormat.backgroundColor"
                    }
                })

                append_log(
                    ws_log,
                    row.get("고객명"),
                    row.get("계약번호"),
                    "🌸 핑크 처리",
                    f"진행상황({진행상황}) ≠ 예상설치일({예상설치일})"
                )

        if format_requests:
            sheet.batch_update({"requests": format_requests})

        messagebox.showinfo(
            "완료",
            "설치일 업데이트 완료\n(로그 기록 완료)\n🌸핑크색상 개별확인"
        )

    except Exception as e:
        messagebox.showerror("에러 발생", str(e))

def run_chungho_install_date_updater():
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
        client = gspread.authorize(creds)

        sheet = client.open_by_url(SPREADSHEET_URL)
        ws1 = sheet.get_worksheet(0)
        ws3 = sheet.get_worksheet(2)
        ws_log = get_or_create_log(sheet)

        _, _, ws1_records = worksheet_to_table(ws1)
        _, _, ws3_records = worksheet_to_table(ws3)

        if not ws1_records or not ws3_records:
            messagebox.showinfo("알림", "시트 데이터가 비어있습니다.")
            return

        col_brand  = get_col_index(ws1, "브랜드")
        col_status = get_col_index(ws1, "진행상황")
        col_v      = get_col_index(ws1, "계약번호")
        col_customer = get_col_index(ws1, "고객명")
        col_c      = get_col_index(ws1, "진행상황")
        col_y      = get_col_index(ws1, "정산") if "정산" in ws1.row_values(1) else get_col_index(ws1, "예상수수료")

        updates, fmt_ranges, log_rows = [], [], []
        updated_count = 0

        for idx1, row1 in enumerate(ws1_records):
            brand = safe_strip(row1.get("브랜드", ""))
            status = safe_strip(row1.get("진행상황", ""))
            contract = safe_strip(row1.get("계약번호", ""))

            if not (brand == "청호" and status == "승인완료" and contract != ""):
                continue

            v_value = contract
            v_last4 = re.sub(r'\D', '', v_value)[-4:]
            f_value = safe_strip(row1.get("고객명", ""))
            rownum = idx1 + 2

            if not v_last4:
                log_rows.append([datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                f_value, v_value, "⛔ 계약번호에 숫자 없음", ""])
                continue

            for row3 in ws3_records:
                b_raw = str(row3.get("계약번호") or row3.get("주문번호") or "")
                b_last4 = b_raw[-4:]
                c_name  = safe_strip(row3.get("고객명") or "")
                n_val   = safe_strip(row3.get("진행상태") or row3.get("상태") or "")
                m_val   = safe_strip(row3.get("설치예정일") or row3.get("매출일") or "")

                if (v_last4 == b_last4) and f_value and c_name and (c_name in f_value):
                    if n_val == "매출확정":
                        try:
                            dt = datetime.strptime(m_val, "%Y-%m-%d")
                            c_text = dt.strftime("%y-%m-%d")
                            y_val  = str(dt.month)

                            updates.append({
                                "range": f"{ws1.title}!{chr(64+col_c)}{rownum}:{chr(64+col_c)}{rownum}",
                                "values": [[c_text]]
                            })
                            updates.append({
                                "range": f"{ws1.title}!{chr(64+col_y)}{rownum}:{chr(64+col_y)}{rownum}",
                                "values": [[y_val]]
                            })

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

def run_script_cheongho():
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
        client = gspread.authorize(creds)

        sheet = client.open_by_url(SPREADSHEET_URL)
        ws1 = sheet.get_worksheet(0)
        ws3 = sheet.get_worksheet(2)
        ws_log = get_or_create_log(sheet)

        ws1_headers, ws1_rows, ws1_records = worksheet_to_table(ws1)
        _, _, ws3_records = worksheet_to_table(ws3)

        if not ws1_headers:
            messagebox.showinfo("알림", "메인 시트가 비어있습니다.")
            return

        if not ws3_records:
            messagebox.showinfo("알림", "청호 설치 시트가 비어있습니다.")
            return

        def header_index(name):
            return ws1_headers.index(name)

        idx_brand = header_index("브랜드") if "브랜드" in ws1_headers else None
        idx_status = header_index("진행상황") if "진행상황" in ws1_headers else None
        idx_contract = header_index("계약번호") if "계약번호" in ws1_headers else None
        idx_customer = header_index("고객명") if "고객명" in ws1_headers else None
        idx_note = header_index("특이사항") if "특이사항" in ws1_headers else None

        if None in (idx_brand, idx_status, idx_contract, idx_customer, idx_note):
            messagebox.showerror("오류 발생", "필수 헤더(브랜드/진행상황/계약번호/고객명/특이사항)가 없습니다.")
            return

        updated_rows = []

        for idx, rec in enumerate(ws1_records):
            brand = safe_strip(rec.get("브랜드"))
            status = safe_strip(rec.get("진행상황"))
            contract_no = safe_strip(rec.get("계약번호"))
            customer_name = safe_strip(rec.get("고객명"))

            if not (brand == "청호" and status in ["계약서", "해피콜", "대기"] and contract_no != ""):
                continue

            if customer_name == "":
                continue

            계약번호_뒤4 = contract_no[-4:]

            for rec3 in ws3_records:
                rec3_contract = safe_strip(rec3.get("계약번호"))
                if rec3_contract == "":
                    rec3_contract = safe_strip(rec3.get("주문번호"))

                if rec3_contract and rec3_contract[-4:] == 계약번호_뒤4:
                    고객명3 = safe_strip(rec3.get("고객명"))
                    if 고객명3 and (customer_name.find(고객명3) != -1):
                        진행상태 = safe_strip(rec3.get("진행상태"))
                        설치예정일값 = safe_strip(rec3.get("설치예정일"))

                        if 진행상태 in ["출고의뢰", "출고확정"]:
                            try:
                                dt = parse(설치예정일값) if 설치예정일값 else None
                                설치예정일_str = dt.strftime("%m-%d 설치예정//") if dt else 설치예정일값
                            except Exception:
                                설치예정일_str = 설치예정일값

                            기존특이 = safe_strip(rec.get("특이사항"))
                            새로운특이 = f"{설치예정일_str} {기존특이}".strip()

                            ws1_rows[idx][idx_status] = "승인완료"
                            ws1_rows[idx][idx_note] = 새로운특이

                            rownum = idx + 2
                            highlight = CellFormat(backgroundColor=Color(1, 1, 0))
                            format_cell_range(ws1, f"A{rownum}:Z{rownum}", highlight)

                            updated_rows.append([
                                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                customer_name, contract_no,
                                f"승인완료 / 특이사항: {새로운특이}"
                            ])
                        break

        ws1.update([ws1_headers] + ws1_rows)

        if updated_rows:
            ws_log.append_rows(updated_rows)

        messagebox.showinfo("완료", f"청호 진행상황 업데이트 완료!\n총 {len(updated_rows)}건 변경됨")

    except Exception as e:
        messagebox.showerror("오류 발생", str(e))

def open_log_sheet():
    webbrowser.open(SPREADSHEET_URL + "#gid=1347292722")

def on_enter(btn, color):
    btn["background"] = color

def on_leave(btn, color):
    btn["background"] = color

if __name__ == "__main__":
    root = Tk()
    root.title("코웨이 설치일 입력 폼")

    icon_path = resource_path("icon.ico")
    icon_path = icon_path.replace("\\", "/")
    if os.path.exists(icon_path):
        try:
            root.iconbitmap(icon_path)
        except Exception:
            pass

    root.geometry("380x320")
    root.configure(bg="#f5f6f7")

    main_frame = Frame(root, bg="#f5f6f7")
    main_frame.pack(fill="both", expand=True, padx=16, pady=16)

    btn_run = Button(
        main_frame,
        text="코웨이 설치확정일 자동입력",
        command=run_install_date_updater,
        font=("맑은 고딕", 12, "bold"),
        bg="#FFF2CC",
        fg="#333333",
        relief="flat",
        activebackground="#FFE699",
        cursor="hand2"
    )
    btn_run.pack(fill="both", expand=True, pady=(0, 10))
    btn_run.bind("<Enter>", lambda e: on_enter(btn_run, "#FFE699"))
    btn_run.bind("<Leave>", lambda e: on_leave(btn_run, "#FFF2CC"))

    btn_log = Button(
        main_frame,
        text="수정 로그 확인",
        command=open_log_sheet,
        font=("맑은 고딕", 11, "bold"),
        bg="#D9EDF7",
        fg="#333333",
        relief="flat",
        activebackground="#BEE5EB",
        cursor="hand2"
    )
    btn_log.pack(fill="both", expand=True)
    btn_log.bind("<Enter>", lambda e: on_enter(btn_log, "#BEE5EB"))
    btn_log.bind("<Leave>", lambda e: on_leave(btn_log, "#D9EDF7"))

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
# pyinstaller main.py --onefile --noconsole --icon=icon.ico --add-data "numeric-haven-455700-k8-541f203927de.json;."