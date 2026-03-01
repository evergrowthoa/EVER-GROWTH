from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# 🔹 아이디 / 비번
USER_ID = "evergrowth"
USER_PW = "wkrlfjqm1!"

driver = webdriver.Chrome()
driver.maximize_window()
wait = WebDriverWait(driver, 15)

# 1️⃣ 로그인 페이지 직접 접속
driver.get("https://allnup.com/login.php")

# 2️⃣ 아이디 입력
wait.until(EC.presence_of_element_located((By.ID, "userid")))
driver.find_element(By.ID, "userid").send_keys(USER_ID)

# 3️⃣ 비밀번호 입력
driver.find_element(By.ID, "password").send_keys(USER_PW)

# 4️⃣ 로그인 버튼 클릭
driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()

print("로그인 버튼 클릭 완료")

# 5️⃣ 로그인 완료 대기 (메인 페이지 로딩)
wait.until(EC.url_contains("allnup.com"))
print("로그인 성공")

# 6️⃣ 접수 리스트 페이지 이동
driver.get("https://allnup.com/layout.php?page=receipt.php")

# 7️⃣ iframe 기다렸다가 진입
wait.until(EC.presence_of_element_located((By.TAG_NAME, "iframe")))
iframe = driver.find_element(By.TAG_NAME, "iframe")
driver.switch_to.frame(iframe)

print("iframe 진입 완료")

# 8️⃣ 리스트 행 로딩 대기
wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table tbody tr")))
rows = driver.find_elements(By.CSS_SELECTOR, "table tbody tr")

print("행 개수:", len(rows))

# 9️⃣ 첫 번째 행 클릭 테스트
if len(rows) > 0:
    rows[0].click()

    wait.until(EC.visibility_of_element_located((By.ID, "customer_name")))
    name = driver.find_element(By.ID, "customer_name").get_attribute("value")
    print("고객명:", name)

else:
    print("행이 없습니다.")

time.sleep(5)
driver.quit()