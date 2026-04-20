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

from gspread.exceptions import WorksheetNotFound, APIError
from gspread.utils import rowcol_to_a1


def _with_retry(fn, *args, **kwargs):
    tries = kwargs.pop("_tries", 6)
    for i in range(tries):
        try:
            return fn(*args, **kwargs)
        except APIError as e:
            status = getattr(getattr(e, "response", None), "status_code", None)
            msg = str(e)
            is_retryable = status in (429, 500, 502, 503, 504) or "Quota exceeded" in msg or "Write requests per minute" in msg
            if is_retryable and i < tries - 1:
                time.sleep(2 * (2 ** i))
                continue
            raise


def batch_values_update(spreadsheet, data, value_input_option="RAW", chunk=200):
    for i in range(0, len(data), chunk):
        body = {
            "valueInputOption": value_input_option,
            "data": data[i:i + chunk]
        }
        _with_retry(spreadsheet.values_batch_update, body)


def batch_format_update(spreadsheet, requests, chunk=100):
    for i in range(0, len(requests), chunk):
        body = {"requests": requests[i:i + chunk]}
        _with_retry(spreadsheet.batch_update, body)


def resource_path(relative_path: str) -> str:
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


CREDENTIALS_FILE = resource_path("numeric-haven-455700-k8-541f203927de.json")
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/17qWvyVONniRI758kESiYS680ChnF7RFHAX-iP-FbrVI/edit"


def get_client():
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
    return gspread.authorize(creds)


def get_sheet_bundle():
    client = get_client()
    spreadsheet = client.open_by_url(SPREADSHEET_URL)
    ws1 = spreadsheet.get_worksheet(0)
    ws2 = spreadsheet.get_worksheet(1)
    ws_log = get_or_create_log(spreadsheet)
    return spreadsheet, ws1, ws2, ws_log


def make_unique_headers_from_row(row, width=None):
    if width is None:
        width = len(row or [])
    row = (row or []) + [""] * (width - len(row or []))
    seen = {}
    out = []
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
    values = _with_retry(ws.get_all_values)
    if not values:
        return pd.DataFrame()
    max_w = max(len(r) for r in values)
    padded = [r + [""] * (max_w - len(r)) for r in values]
    headers = make_unique_headers_from_row(padded[0], width=max_w)
    return pd.DataFrame(padded[1:], columns=headers)


def get_headers(ws):
    return _with_retry(ws.row_values, 1)


def get_col_index_by_headers(headers, header_name):
    if header_name not in headers:
        raise KeyError(f"시트 헤더에 '{header_name}' 컬럼이 없습니다.")
    return headers.index(header_name) + 1


def get_or_create_log(spreadsheet):
    try:
        return spreadsheet.worksheet("Log")
    except WorksheetNotFound:
        ws = _with_retry(spreadsheet.add_worksheet, title="Log", rows=2000, cols=10)
        _with_retry(
            ws.append_rows,
            [
                ["timestamp", "customer_name", "contract_no", "content", "note"],
                [datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "-", "-", "Log sheet auto-created", "-"]
            ],
            value_input_option="RAW"
        )
        return ws


def append_logs_bulk(ws_log, log_rows):
    if not log_rows:
        return
    for i in range(0, len(log_rows), 200):
        chunk = log_rows[i:i + 200]
        _with_retry(ws_log.append_rows, chunk, value_input_option="RAW")


def normalize_name(name):
    if not name:
        return ""
    cleaned = re.sub(r"[^가-힣A-Za-z]", "", str(name))
    return cleaned.upper()


def digits_last4(value):
    digits = re.sub(r"\D", "", str(value or ""))
    if len(digits) < 4:
        return ""
    return digits[-4:]


def parse_date_flexible(raw):
    raw = str(raw or "").strip()
    if not raw:
        return None
    for fmt in ("%Y.%m.%d", "%Y-%m-%d", "%Y/%m/%d", "%y-%m-%d"):
        try:
            return datetime.strptime(raw, fmt)
        except Exception:
            pass
    try:
        return parse(raw)
    except Exception:
        return None


def color_request(sheet_id, rownum, col_index, red, green, blue):
    return {
        "repeatCell": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": rownum - 1,
                "endRowIndex": rownum,
                "startColumnIndex": col_index - 1,
                "endColumnIndex": col_index
            },
            "cell": {
                "userEnteredFormat": {
                    "backgroundColor": {
                        "red": red,
                        "green": green,
                        "blue": blue
                    }
                }
            },
            "fields": "userEnteredFormat.backgroundColor"
        }
    }


def prepare_dataframes(ws1, ws2):
    df1 = worksheet_to_dataframe(ws1)
    df2 = worksheet_to_dataframe(ws2)

    required1 = ["브랜드", "진행상황", "계약번호", "고객명"]
    required2 = ["주문번호", "고객명", "상태", "설치예정일"]

    for col in required1:
        if col not in df1.columns:
            raise KeyError(f"시트1에 '{col}' 컬럼이 없습니다.")
    for col in required2:
        if col not in df2.columns:
            raise KeyError(f"시트2에 '{col}' 컬럼이 없습니다.")

    df1["브랜드"] = df1["브랜드"].astype(str).str.strip()
    df1["진행상황"] = df1["진행상황"].astype(str).str.strip()
    df1["계약번호"] = df1["계약번호"].astype(str).str.strip()
    df1["고객명"] = df1["고객명"].astype(str).str.strip()

    if "특이사항" in df1.columns:
        df1["특이사항"] = df1["특이사항"].astype(str).str.strip()
    else:
        df1["특이사항"] = ""

    df2["주문번호"] = df2["주문번호"].astype(str).str.strip()
    df2["고객명"] = df2["고객명"].astype(str).str.strip()
    df2["상태"] = df2["상태"].astype(str).str.strip()
    df2["설치예정일"] = df2["설치예정일"].astype(str).str.strip()

    if "배정시간" not in df2.columns:
        df2["배정시간"] = ""
    else:
        df2["배정시간"] = df2["배정시간"].astype(str).str.strip()

    return df1, df2


def build_coway_lookup(df2):
    lookup = {}
    for _, row2 in df2.iterrows():
        order_last4 = digits_last4(row2.get("주문번호", ""))
        customer2 = normalize_name(row2.get("고객명", ""))
        if not order_last4 or not customer2:
            continue
        lookup.setdefault(order_last4, []).append(row2)
    return lookup


def find_matching_row(row1, lookup):
    contract_last4 = digits_last4(row1.get("계약번호", ""))
    customer1 = normalize_name(row1.get("고객명", ""))

    if not contract_last4 or not customer1:
        return None

    candidates = lookup.get(contract_last4, [])
    for row2 in candidates:
        customer2 = normalize_name(row2.get("고객명", ""))
        if customer1 and customer2 and (customer1 in customer2 or customer2 in customer1):
            return row2
    return None



def run_install_date_updater():
    try:
        spreadsheet, ws1, ws2, ws_log = get_sheet_bundle()
        headers1 = get_headers(ws1)

        col_status = get_col_index_by_headers(headers1, "진행상황")
        col_expected = get_col_index_by_headers(headers1, "예상설치일")

        df1, df2 = prepare_dataframes(ws1, ws2)
        lookup = build_coway_lookup(df2)

        value_updates = []
        format_requests = []
        log_rows = []

        success_count = 0
        mismatch_count = 0
        fail_count = 0

        total_cols = len(headers1)

        def row_color_request(sheet_id, rownum, total_columns, red, green, blue):
            return {
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": rownum - 1,
                        "endRowIndex": rownum,
                        "startColumnIndex": 0,
                        "endColumnIndex": total_columns
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": {
                                "red": red,
                                "green": green,
                                "blue": blue
                            }
                        }
                    },
                    "fields": "userEnteredFormat.backgroundColor"
                }
            }

        condition = (
            (df1["브랜드"] == "코웨이") &
            (df1["진행상황"] == "승인완료")
        )

        for idx, row1 in df1[condition].iterrows():
            rownum = idx + 2
            contract_no = str(row1.get("계약번호", "")).strip()
            customer_name = str(row1.get("고객명", "")).strip()
            expected_raw = str(row1.get("예상설치일", "")).strip()

            match_row = find_matching_row(row1, lookup)
            if match_row is None:
                format_requests.append(row_color_request(ws1.id, rownum, total_cols, 1, 1, 0.6))
                log_rows.append([
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    customer_name,
                    contract_no,
                    "설치확정 미매칭",
                    "시트2에서 주문번호/고객명 매칭 실패"
                ])
                fail_count += 1
                continue

            status2 = str(match_row.get("상태", "")).replace(" ", "")
            if "순주문확정" not in status2 and "설치확정" not in status2:
                log_rows.append([
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    customer_name,
                    contract_no,
                    "설치확정 상태 제외",
                    f"시트2 상태={status2}"
                ])
                continue

            raw_date = str(match_row.get("설치예정일", "")).strip()
            parsed_sheet2 = parse_date_flexible(raw_date)
            parsed_sheet1 = parse_date_flexible(expected_raw)

            if parsed_sheet2 is None or parsed_sheet1 is None:
                format_requests.append(row_color_request(ws1.id, rownum, total_cols, 1, 0.7, 0.7))
                log_rows.append([
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    customer_name,
                    contract_no,
                    "날짜 비교 실패",
                    f"시트1 예상설치일={expected_raw} / 시트2 설치예정일={raw_date}"
                ])
                mismatch_count += 1
                continue

            formatted_sheet2 = parsed_sheet2.strftime("%y-%m-%d")
            formatted_sheet1 = parsed_sheet1.strftime("%y-%m-%d")

            if formatted_sheet1 != formatted_sheet2:
                format_requests.append(row_color_request(ws1.id, rownum, total_cols, 1, 0.7, 0.7))
                log_rows.append([
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    customer_name,
                    contract_no,
                    "날짜 불일치",
                    f"시트1 예상설치일={formatted_sheet1} / 시트2 설치예정일={formatted_sheet2}"
                ])
                mismatch_count += 1
                continue

            a1_status = rowcol_to_a1(rownum, col_status)
            value_updates.append({
                "range": f"{ws1.title}!{a1_status}",
                "values": [[formatted_sheet2]]
            })

            log_rows.append([
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                customer_name,
                contract_no,
                "설치확정일 입력 완료",
                formatted_sheet2
            ])

            success_count += 1

        if value_updates:
            batch_values_update(spreadsheet, value_updates)

        if format_requests:
            batch_format_update(spreadsheet, format_requests)

        append_logs_bulk(ws_log, log_rows)

        messagebox.showinfo(
            "완료",
            "설치일 처리 완료\n"
            f"✔ 입력 완료: {success_count}건\n"
            f"⚠ 날짜 불일치(행 전체 핑크): {mismatch_count}건\n"
            f"❌ 미매칭(행 전체 노란색): {fail_count}건"
        )

    except Exception as e:
        messagebox.showerror("에러 발생", str(e))


def open_log_sheet():
    webbrowser.open(SPREADSHEET_URL + "#gid=1347292722")


def main():
    root = Tk()
    root.title("코웨이 설치일 입력 폼")
    root.geometry("360x280")

    Button(root, text="코웨이 설치확정일 자동입력", command=run_install_date_updater, width=34, height=2, bg="lightyellow").pack(pady=20)
    Button(root, text="수정 로그 확인", command=open_log_sheet, width=34, height=2, bg="lightblue").pack(pady=10)

    root.mainloop()


if __name__ == "__main__":
    main()


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