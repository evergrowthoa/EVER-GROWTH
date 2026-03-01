from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import time

# ===========================
# 1. 드라이버 실행
# ===========================
driver = webdriver.Chrome()
wait = WebDriverWait(driver, 20)

driver.get("https://allnup.com")

# ===========================
# 2. 로그인
# ===========================
wait.until(EC.presence_of_element_located((By.NAME, "mb_id"))).send_keys("아이디입력")
driver.find_element(By.NAME, "mb_password").send_keys("비밀번호입력")

driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()
print("로그인 버튼 클릭")

time.sleep(3)
print("로그인 성공")

# ===========================
# 3. 접수리스트 페이지로 바로 이동
# ===========================
driver.get("https://allnup.com/layout.php?page=receipt.php")
print("접수리스트 이동")

# ===========================
# 4. iframe 진입 (중요)
# ===========================
wait.until(EC.presence_of_element_located((By.TAG_NAME, "iframe")))
iframe = driver.find_element(By.TAG_NAME, "iframe")
driver.switch_to.frame(iframe)
print("iframe 진입 완료")

# ===========================
# 5. 리스트 로딩 대기
# ===========================
wait.until(EC.presence_of_element_located((By.TAG_NAME, "table")))
print("리스트 로딩 완료")

print("\n📌 이제 원하는 고객을 수동 클릭해서 모달을 띄우세요")

# ===========================
# 6. 모달 감지 루프 (핵심)
# ===========================
while True:
    try:
        # 모달 안에 있는 고객명 input 감지
        customer_name = wait.until(
            EC.presence_of_element_located((By.NAME, "customer_name"))
        )

        print("\n✅ 모달 감지 성공!")

        # 값 읽기
        name = customer_name.get_attribute("value")
        phone = driver.find_element(By.NAME, "customer_tel").get_attribute("value")
        birth = driver.find_element(By.NAME, "customer_birth").get_attribute("value")

        print("고객명:", name)
        print("연락처:", phone)
        print("생년월일:", birth)

        print("\n📌 다음 고객을 클릭하세요 (계속 감시 중)\n")

        time.sleep(2)

    except TimeoutException:
        # 모달이 아직 안뜬 상태
        pass