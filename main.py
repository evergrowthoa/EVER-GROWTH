import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
from dateutil.parser import parse
import re

#진행상황 업데이트하는 파일



# 🔸 인증 JSON 파일명 - 실제 경로와 이름으로 수정
CREDENTIALS_FILE = 'numeric-haven-455700-k8-7e15ff3d6313.json'

# 🔹 스프레드시트 URL
SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/17qWvyVONniRI758kESiYS680ChnF7RFHAX-iP-FbrVI//edit'

# 🔹 인증 및 시트 접근
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
client = gspread.authorize(creds)

sheet = client.open_by_url(SPREADSHEET_URL)
ws1 = sheet.get_worksheet(0)  # 시트1
ws2 = sheet.get_worksheet(1)  # 시트2

# 📥 데이터프레임으로 변환
df1 = pd.DataFrame(ws1.get_all_records())
df2 = pd.DataFrame(ws2.get_all_records())

# 🔍 필터 조건 (시트1에서)
condition = (
    (df1["브랜드"] == "코웨이") &
    (df1["진행상황"].isin(["계약서", "해피콜", "동의서", "대기"])) &
    (df1["비가망유형"].astype(str).str.strip() != "")
)

# ▶️ 조건에 맞는 행 반복
for idx, row in df1[condition].iterrows():
    v_value = str(row["비가망유형"]).strip()
    v_last4 = ''.join(re.findall(r'\d+', v_value))[-4:] if re.search(r'\d', v_value) else ''
    f_value = str(row["고객명"]).strip()

    print(f"▶️ 검사중 - 시트1 {idx+2}행: 비가망유형={v_value}, 마지막4자리={v_last4}, 고객명={f_value}")

    if not v_last4:
        print("  ⛔ 숫자 4자리가 없어 비교 생략")
        continue

    for jdx, row2 in df2.iterrows():
        b_last4 = str(row2.get("주문번호", ""))[-4:]
        상태값 = str(row2.get("상태", "")).strip()
        고객명2 = str(row2.get("고객명", ""))

        if v_last4 == b_last4 and f_value in 고객명2:
            if 상태값 in ["신용조사(가완료)","신용조사" ,"출고의뢰"]:
                raw_date = str(row2.get("설치예정일", "")).strip()
                try:
                    parsed_date = parse(raw_date)
                    l_val = parsed_date.strftime("%m-%d")  # 날짜형식: MM-DD
                except:
                    l_val = raw_date

                m_val = str(row2.get("배정시간", "")).strip()
                existing_note = str(row.get("특이사항", "")).strip()
                new_note = f"{l_val} {m_val}"
                combined_note = f"{new_note} | {existing_note}" if existing_note else new_note

                ws1.update_cell(idx + 2, 16, combined_note)  # P열 (특이사항)
                ws1.update_cell(idx + 2, 3, "승인완료")       # C열 (진행상황)
                print(f"✅ 업데이트 완료 - {idx+2}행 → 특이사항: '{combined_note}', 진행상황: '승인완료'")
                break
            else:
                print(f"  ⛔ 상태값 불일치 (필요: '신용조사(가완료)' 또는 '출고의뢰') → 현재: '{상태값}'")
