import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
from tkinter import messagebox, Tk
import re
import os
import sys
from gspread_formatting import CellFormat, Color, format_cell_range

# -------------------------------
# PyInstaller/로컬 공통 경로 헬퍼
# -------------------------------
def resource_path(relative_path: str) -> str:
    try:
        base_path = sys._MEIPASS  # PyInstaller 임시 폴더
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# 🔑 JSON 키 경로 (파일명 본인 것에 맞추세요)
CREDENTIALS_FILE = resource_path('evergrowth-493504-abdee694f352.json')

# 🔗 구글 스프레드시트 URL
SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/17qWvyVONniRI758kESiYS680ChnF7RFHAX-iP-FbrVI/edit'

# ✅ 노란색 셀 포맷
yellow_fill = CellFormat(backgroundColor=Color(1, 1, 0))

# -------------------------------
# 헤더 생성 유틸: 빈칸/중복 안전 처리
# -------------------------------
def make_unique_headers_from_row(row, width=None):
    """row: 1행 값 리스트. width 지정 시 그 길이에 맞게 패딩."""
    if width is None:
        width = len(row)
    # 길이 보정 (행이 짧으면 빈칸으로 채움)
    row = (row or []) + [""] * (width - len(row or []))
    seen = {}
    out = []
    for i, h in enumerate(row):
        h = (h or "").strip()
        if not h:
            h = f"col{i+1}"
        if h in seen:
            seen[h] += 1
            h = f"{h}_{seen[h]}"
        else:
            seen[h] = 0
        out.append(h)
    return out

def worksheet_to_dataframe(ws):
    """
    get_all_values()로 전부 읽고,
    1행을 가져와 유니크 헤더로 변환해 DataFrame 구성.
    헤더 검사를 피하므로 get_all_records()에서 나던 오류를 원천 차단.
    """
    values = ws.get_all_values()  # 전체 값
    if not values:
        return pd.DataFrame()

    # 행들 중 가장 긴 길이로 맞춰서 패딩
    max_w = max(len(r) for r in values)
    padded = [r + [""] * (max_w - len(r)) for r in values]

    headers = make_unique_headers_from_row(padded[0], width=max_w)
    data = padded[1:]
    df = pd.DataFrame(data, columns=headers)
    return df

# -------------------------------
# 설치확정일/월 자동 입력
# -------------------------------
def run_install_date_updater():
    try:
        scope = [
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive'
        ]
        creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
        client = gspread.authorize(creds)

        # 시트 열기
        sheet = client.open_by_url(SPREADSHEET_URL)
        ws1 = sheet.get_worksheet(0)  # 시트1
        ws3 = sheet.get_worksheet(2)  # 시트3 (index 2)

        # 🛡️ 헤더 검증 우회 로딩
        df1 = worksheet_to_dataframe(ws1)
        df3 = worksheet_to_dataframe(ws3)

        # ---- 시트1 조건 필터 ----
        # col 인덱스 기반 (H=7, C=2, V=21) — df1가 비어있지 않은지 체크
        if df1.empty:
            messagebox.showinfo("알림", "시트1에 데이터가 없습니다.")
            return

        # 문자열 비교 안전하게 처리
        def col_eq(df, idx, value):
            return (df.iloc[:, idx].astype(str) == str(value))

        condition = (
            col_eq(df1, 7, "청호") &                   # H열
            col_eq(df1, 2, "승인완료") &              # C열
            (df1.iloc[:, 21].astype(str).str.strip() != "")  # V열
        )

        updated_count = 0

        # 시트3 컬럼 추정 함수
        def pick_col(df, candidates):
            for name in candidates:
                if name in df.columns:
                    return name
            return None

        if df3.empty:
            messagebox.showinfo("알림", "시트3에 데이터가 없습니다.")
            return

        col_contract = pick_col(df3, ["계약번호", "주문번호", "B", "col2"])
        col_customer = pick_col(df3, ["고객명", "성명", "C", "col3"])
        col_status   = pick_col(df3, ["진행상태", "상태", "N", "col14"])
        col_m_date   = pick_col(df3, ["M열", "설치예정일", "매출일", "M", "col13"])

        if not all([col_contract, col_customer, col_status, col_m_date]):
            messagebox.showerror("에러", "시트3에서 필요한 열을 찾지 못했습니다.\n(계약번호/고객명/진행상태/M열)")
            return

        for idx1, row1 in df1[condition].iterrows():
            v_value = str(row1.iloc[21]).strip()              # V열
            v_last4 = re.sub(r'\D', '', v_value)[-4:]
            f_value = str(row1.iloc[5]).strip()               # F열(고객명)

            if not v_last4:
                continue

            for _, row3 in df3.iterrows():
                b_last4 = str(row3[col_contract])[-4:]
                c_value_sheet3 = str(row3[col_customer]).strip()
                n_value = str(row3[col_status]).strip()
                m_value = str(row3[col_m_date]).strip()       # 기대: YYYY-MM-DD

                if (v_last4 == b_last4) and (f_value in c_value_sheet3) and (n_value == "매출확정"):
                    try:
                        dt = datetime.strptime(m_value, "%Y-%m-%d")
                        formatted_b = dt.strftime("%y-%m-%d")  # B열(YY-MM-DD)
                        month_only  = dt.strftime("%m")         # Y열(MM)

                        ws1.update_cell(idx1 + 2, 2, formatted_b)  # B열
                        ws1.update_cell(idx1 + 2, 25, month_only)  # Y열

                        format_cell_range(ws1, f"B{idx1+2}", yellow_fill)
                        format_cell_range(ws1, f"Y{idx1+2}", yellow_fill)
                        updated_count += 1
                    except Exception as e:
                        print(f"날짜 변환 오류: {m_value} -> {e}")
                    break

        messagebox.showinfo("완료", f"총 {updated_count}건 변경 완료 ✅")

    except Exception as e:
        messagebox.showerror("에러 발생", str(e))

# -------------------------------
# 엔트리 포인트
# -------------------------------
if __name__ == "__main__":
    root = Tk()
    root.withdraw()
    run_install_date_updater()
