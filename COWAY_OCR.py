from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

USER_ID = "evergrowth"
USER_PW = "wkrlfjqm1!"

driver = webdriver.Chrome()
driver.maximize_window()
wait = WebDriverWait(driver, 15)

# 1️⃣ 로그인
driver.get("https://allnup.com/login.php")

wait.until(EC.presence_of_element_located((By.ID, "userid")))
driver.find_element(By.ID, "userid").send_keys(USER_ID)
driver.find_element(By.ID, "password").send_keys(USER_PW)
driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()

print("로그인 완료")

# 2️⃣ 접수요청 페이지 이동
driver.get("https://allnup.com/layout.php?page=receipt.php")

print("📌 이제 수동으로 리스트에서 원하는 고객을 클릭해서 모달을 띄우세요.")
print("📌 모달이 뜨면 자동으로 고객정보를 읽습니다.")

# 3️⃣ 무한 대기 → 모달 뜨는 순간 감지
while True:
    try:
        wait.until(EC.visibility_of_element_located((By.ID, "customer_name")))

        name = driver.find_element(By.ID, "customer_name").get_attribute("value")
        phone = driver.find_element(By.ID, "customer_phone").get_attribute("value")

        print("고객명:", name)
        print("전화번호:", phone)
        print("------------")

        time.sleep(3)  # 중복 감지 방지용 대기

    except:
        pass