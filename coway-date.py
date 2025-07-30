import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
import re
from datetime import datetime

#설치일 입력하는파일



# 인증
CREDENTIALS_FILE = 'numeric-haven-455700-k8-7e15ff3d6313.json'
SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/17qWvyVONniRI758kESiYS680ChnF7RFHAX-iP-FbrVI//edit'

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
client = gspread.authorize(creds)

# 워크시트 가져오기
sheet = client.open_by_url(SPREADSHEET_URL)
ws1 = sheet.get_worksheet(0)  # 시트1
ws2 = sheet.get_worksheet(1)  # 시트2

# DataFrame 변환
df1 = pd.DataFrame(ws1.get_all_records())
df2 = pd.DataFrame(ws2.get_all_records())

# 조건: 시트1 - 진행상황이 '승인완료', 브랜드가 '코웨이'
condition = (df1["진행상황"] == "승인완료") & (df1["브랜드"] == "코웨이")

for idx, row in df1[condition].iterrows():
    v_value = str(row["비가망유형"]).strip()
    v_last4 = ''.join(re.findall(r'\d+', v_value))[-4:] if re.search(r'\d', v_value) else ''
    customer1 = str(row["고객명"]).strip()

    print(f"▶️ 검사중 - 시트1 {idx+2}행: 비가망유형={v_value}, 마지막4자리={v_last4}, 고객명={customer1}")

    if not v_last4:
        print(f"❌ {idx+2}행 - 비가망유형에서 숫자 4자리 없음, 건너뜀")
        continue

    updated = False

    for jdx, row2 in df2.iterrows():
        b_last4 = str(row2["주문번호"])[-4:]
        customer2 = str(row2["고객명"]).strip()
        status = str(row2["상태"]).strip()

        # ✅ 조건 추가:
        # 1) 마지막 4자리 일치
        # 2) 시트1 고객명이 시트2 고객명에 포함
        # 3) 상태가 '신용조사(가완료)' 또는 '출고의뢰'
        if (
            v_last4 == b_last4 and
            customer1 in customer2 and
            status in ["순주문확정"]
        ):
            raw_date = str(row2["설치예정일"]).strip()
            try:
                formatted_date = datetime.strptime(raw_date, "%Y.%m.%d").strftime("%y-%m-%d")
                ws1.update_cell(idx + 2, 3, formatted_date)  # C열(3번째)에 날짜 업데이트
                print(f"✅ {idx+2}행 - C열에 날짜 '{formatted_date}' 입력됨")
                updated = True
                break
            except Exception as e:
                print(f"❌ {idx+2}행 - 날짜 형식 오류: {raw_date}")
                break

    if not updated:
        print(f"❌ {idx+2}행 - 조건 만족하는 시트2 데이터 없음")

