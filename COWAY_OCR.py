from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import time
import re
import threading
import queue

import uiautomator2 as u2

# ===========================
# 로그인 정보
# ===========================
USER_ID = "evergrowth"
USER_PW = "wkrlfjqm1!"

# ===========================
# 에뮬레이터 설정
# ===========================
ADB_SERIAL = "emulator-5554"   # adb devices에서 보이는 값으로 바꾸세요
DO_EMULATOR = True            # 에뮬레이터 자동화를 잠깐 끄고 싶으면 False

# ===========================
# 크롬 옵션
# ===========================
chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_experimental_option("detach", True)

driver = webdriver.Chrome(options=chrome_options)
wait = WebDriverWait(driver, 20)

# ===========================
# 작업 큐 (웹 → 에뮬레이터)
# ===========================
job_q = queue.Queue()

def normalize_phone(s: str) -> str:
    if s is None:
        return ""
    digits = re.sub(r"\D", "", str(s))
    return digits

def connect_emulator():
    d = u2.connect(ADB_SERIAL)
    d.implicitly_wait(5.0)
    return d

def ensure_on_mobile_order_home(d):
    # 1번 홈 화면(모바일 주문 탭) 기준으로 맞추는 용도
    # 화면이 어디에 있든, 뒤로가기를 몇 번 누르고 "일반 주문하기"가 보이면 성공으로 처리
    for _ in range(6):
        if d(text="일반 주문하기").exists:
            return True
        d.press("back")
        time.sleep(0.6)

    # 그래도 못 찾으면, 하단 탭의 "모바일 주문"을 눌러서 진입 시도
    if d(text="모바일 주문").exists:
        d(text="모바일 주문").click()
        time.sleep(0.8)
        if d(text="일반 주문하기").exists:
            return True

    return d(text="일반 주문하기").exists

def send_auth_request(d, name: str, phone11: str) -> bool:
    # 3번 화면(주문접수 1/6)까지 진입 후
    # 개인 클릭 → 고객명 입력 → 휴대폰 입력 → 본인인증 요청 클릭

    if not ensure_on_mobile_order_home(d):
        print("❌ 에뮬레이터: 홈(모바일 주문) 화면을 찾지 못했습니다.")
        return False

    if not d(text="일반 주문하기").exists:
        print("❌ 에뮬레이터: '일반 주문하기' 버튼이 없습니다.")
        return False

    d(text="일반 주문하기").click()
    time.sleep(0.8)

    # 주문접수 화면 확인(상단 텍스트로 가볍게 체크)
    if not d(text="주문접수").exists:
        # 기기/앱에 따라 상단 텍스트가 다를 수 있어서, 다음 요소로도 체크
        if not d(text="본인인증 요청").exists:
            print("❌ 에뮬레이터: 주문접수(고객정보 입력) 화면이 아닙니다.")
            return False

    # 고객 구분: 개인
    if d(text="개인").exists:
        d(text="개인").click()
        time.sleep(0.3)

    # 입력칸 2개: 보통 첫번째=고객명, 두번째=휴대폰
    edits = d(className="android.widget.EditText")
    if edits.count < 2:
        print("❌ 에뮬레이터: 입력칸(EditText)이 2개가 아닙니다.")
        return False

    # 고객명 입력
    edits[0].click()
    time.sleep(0.1)
    edits[0].set_text(name)

    # 휴대폰 입력(11자리)
    edits[1].click()
    time.sleep(0.1)
    edits[1].set_text(phone11)

    # 본인인증 요청 버튼이 활성화될 때까지 잠깐 대기 후 클릭
    btn = d(text="본인인증 요청")
    for _ in range(20):
        if btn.exists and btn.info.get("enabled", False):
            btn.click()
            time.sleep(0.8)
            print("✅ 에뮬레이터: 본인인증 요청 클릭 완료")
            return True
        time.sleep(0.2)

    # enabled 확인이 기기별로 안 잡히는 경우가 있어서, 마지막으로 그냥 클릭 1회 시도
    if btn.exists:
        btn.click()
        time.sleep(0.8)
        print("✅ 에뮬레이터: 본인인증 요청 클릭(강제) 시도 완료")
        return True

    print("❌ 에뮬레이터: '본인인증 요청' 버튼을 찾지 못했습니다.")
    return False

def emulator_worker():
    if not DO_EMULATOR:
        print("ℹ️ 에뮬레이터 자동화가 꺼져 있습니다(DO_EMULATOR=False).")
        while True:
            job_q.get()
            job_q.task_done()

    d = connect_emulator()
    print("✅ 에뮬레이터 연결 완료:", ADB_SERIAL)

    while True:
        job = job_q.get()
        try:
            name = job.get("name", "")
            phone11 = job.get("phone11", "")
            print(f"🚀 에뮬레이터 작업 시작: {name} / {phone11}")

            ok = send_auth_request(d, name, phone11)
            if ok:
                print(f"✅ 인증발송 성공: {name} / {phone11}")
            else:
                print(f"❌ 인증발송 실패: {name} / {phone11}")

        except Exception as e:
            print("❌ 에뮬레이터 워커 에러:", e)
        finally:
            job_q.task_done()

# 워커 스레드 시작
t = threading.Thread(target=emulator_worker, daemon=True)
t.start()

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
        # 화면에 보이는 input만 대상으로 하면 index가 덜 흔들립니다.
        all_inputs = driver.find_elements(By.TAG_NAME, "input")
        inputs = [x for x in all_inputs if x.is_displayed()]

        # 모달이 열렸을 때 (input이 많아짐)
        if len(inputs) > 15:
            name = (inputs[10].get_attribute("value") or "").strip()
            birth = (inputs[11].get_attribute("value") or "").strip()
            phone_raw = (inputs[12].get_attribute("value") or "").strip()
            account = (inputs[13].get_attribute("value") or "").strip()
            zipcode = (inputs[14].get_attribute("value") or "").strip()

            phone_digits = normalize_phone(phone_raw)

            # 전화번호 없으면 무시 (빈 모달 방지)
            if not phone_digits:
                time.sleep(0.5)
                continue

            # 11자리만 사용(앞자리 0 포함)
            if len(phone_digits) >= 11:
                phone11 = phone_digits[-11:]
            else:
                print(f"❌ 전화번호가 11자리가 아닙니다: {phone_raw}")
                time.sleep(0.8)
                continue

            # 이미 처리한 번호면 스킵
            if phone11 in processed_phones:
                time.sleep(1)
                continue

            processed_phones.add(phone11)

            print("\n✅ 신규 고객 감지")
            print("이름:", name)
            print("생년월일:", birth)
            print("전화번호(11):", phone11)
            print("계좌:", account)
            print("우편번호:", zipcode)
            print("-" * 40)

            # 에뮬레이터 작업 큐에 전달
            job_q.put({
                "name": name,
                "phone11": phone11,
                "birth": birth,
                "account": account,
                "zipcode": zipcode,
            })

            time.sleep(1.2)

        else:
            time.sleep(0.5)

    except Exception as e:
        print("에러 발생:", e)
        time.sleep(1)