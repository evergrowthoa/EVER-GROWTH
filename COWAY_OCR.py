from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import time

# ===========================
# 로그인 정보
# ===========================
USER_ID = "evergrowth"
USER_PW = "wkrlfjqm1!"

# ===========================
# 크롬 옵션
# ===========================
chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_experimental_option("detach", True)

driver = webdriver.Chrome(options=chrome_options)
wait = WebDriverWait(driver, 20)

# ===========================
# 1. 사이트 접속
# ===========================
driver.get("https://allnup.com")

wait.until(lambda d: len(d.find_elements(By.TAG_NAME, "input")) >= 2)
inputs = driver.find_elements(By.TAG_NAME, "input")

inputs[0].send_keys(USER_ID)
inputs[1].send_keys(USER_PW)

buttons = driver.find_elements(By.TAG_NAME, "button")
if buttons:
    buttons[0].click()
else:
    driver.find_element(By.CSS_SELECTOR, "input[type='submit']").click()

time.sleep(3)

# ===========================
# 2. 접수리스트 이동
# ===========================
driver.get("https://allnup.com/layout.php?page=receipt.php")

wait.until(EC.presence_of_element_located((By.TAG_NAME, "iframe")))
iframe = driver.find_element(By.TAG_NAME, "iframe")
driver.switch_to.frame(iframe)

wait.until(EC.presence_of_element_located((By.TAG_NAME, "table")))

print("\n📌 고객을 수동 클릭해서 모달을 띄우세요.\n")

# ===========================
# 3. 중복 방지용 저장소
# ===========================
processed_phones = set()

# ===========================
# 4. 모달 감시 루프
# ===========================
while True:
    try:
        inputs = driver.find_elements(By.TAG_NAME, "input")

        if len(inputs) > 15:

            name = inputs[10].get_attribute("value")
            birth = inputs[11].get_attribute("value")
            phone = inputs[12].get_attribute("value")
            account = inputs[13].get_attribute("value")
            zipcode = inputs[14].get_attribute("value")

            # 📌 이미 처리한 번호면 스킵
            if phone in processed_phones:
                time.sleep(1)
                continue

            # 새 번호면 저장
            processed_phones.add(phone)

            print("\n✅ 신규 고객 감지")
            print("이름:", name)
            print("생년월일:", birth)
            print("전화번호:", phone)
            print("계좌:", account)
            print("우편번호:", zipcode)
            print("-" * 40)

            time.sleep(2)

    except:
        pass