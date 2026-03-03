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
ADB_SERIAL = "emulator-5554"
DO_EMULATOR = True

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


# ===========================
# 유틸: 문자열/속성
# ===========================
def _attr(el, key: str) -> str:
    return (el.get_attribute(key) or "").strip()


def _meta(el) -> str:
    return " ".join([
        _attr(el, "id"),
        _attr(el, "name"),
        _attr(el, "placeholder"),
        _attr(el, "aria-label"),
        _attr(el, "class"),
        _attr(el, "type"),
        _attr(el, "inputmode"),
        _attr(el, "autocomplete"),
    ]).lower()


def normalize_digits(s: str) -> str:
    return re.sub(r"\D", "", str(s or ""))


def normalize_phone11_only_010(s: str) -> str:
    digits = normalize_digits(s)

    # +82 10xxxxxxxx 형태 보정(혹시 몰라서)
    if digits.startswith("82") and len(digits) >= 12 and digits[2:4] == "10":
        digits = "0" + digits[2:]

    # 우리 플로우는 010 11자리만
    if digits.startswith("010") and len(digits) == 11:
        return digits
    return ""


def is_accountish_value(v: str) -> bool:
    v = str(v or "")
    if any(k in v for k in ["계좌", "은행", "뱅크", "카카오", "신한", "국민", "우리", "하나", "/"]):
        return True
    return False


def has_phone_hint(meta: str) -> bool:
    return any(k in meta for k in ["phone", "tel", "mobile", "휴대", "연락", "핸드폰", "휴대폰", "11자리"])


# ===========================
# 모달 값 추출(인덱스 의존 제거)
# ===========================
def pick_name_from_inputs(inputs) -> str:
    for el in inputs:
        v = (_attr(el, "value") or "").strip()
        if re.fullmatch(r"[가-힣]{2,6}", v):
            return v
    return ""


def pick_birth_from_inputs(inputs) -> str:
    for el in inputs:
        v = (_attr(el, "value") or "").strip()
        if re.fullmatch(r"\d{6}-\d{1}", v):
            return v
        v2 = normalize_digits(v)
        if re.fullmatch(r"\d{7}", v2):
            return v
    return ""


def pick_zip_from_inputs(inputs) -> str:
    for el in inputs:
        v = (_attr(el, "value") or "").strip()
        if re.fullmatch(r"\d{5}", v):
            return v
    return ""


def pick_account_from_inputs(inputs) -> str:
    for el in inputs:
        v = (_attr(el, "value") or "").strip()
        meta = _meta(el)
        if any(k in meta for k in ["account", "acct", "bank", "계좌", "은행", "뱅크"]) or is_accountish_value(v):
            if v:
                return v
    return ""


def pick_phone11_from_inputs(inputs):
    """
    핵심: 숫자 패턴으로만 찾지 않고,
    '휴대폰 필드 힌트(placeholder/id/name/type/...)'가 있는 input에서만 phone을 추출한다.
    """
    best_score = -10**9
    best_phone11 = ""
    best_raw = ""

    for el in inputs:
        raw = (_attr(el, "value") or "").strip()
        if not raw:
            continue

        meta = _meta(el)

        # 계좌/은행 쪽 힌트면 애초에 제외
        if any(k in meta for k in ["account", "acct", "bank", "계좌", "은행", "뱅크"]) or is_accountish_value(raw):
            continue

        phone11 = normalize_phone11_only_010(raw)
        if not phone11:
            continue

        # ✅ 휴대폰 필드 힌트가 없는 값은 후보에서 제외 (010 계좌/숫자 오인 방지)
        if not has_phone_hint(meta):
            continue

        score = 0
        score += 100  # 힌트 기반 필드라는 자체가 큰 가점

        if "tel" in meta:
            score += 30
        if "010" in phone11:
            score += 10
        if "-" in raw:
            score += 3

        if score > best_score:
            best_score = score
            best_phone11 = phone11
            best_raw = raw

    return best_phone11, best_raw


# ===========================
# 에뮬레이터 제어
# ===========================
def connect_emulator():
    d = u2.connect(ADB_SERIAL)
    d.implicitly_wait(5.0)
    return d


def ensure_on_mobile_order_home(d):
    # ✅ 화면을 "절대 건드리지 않고", '일반 주문하기'가 보일 때만 True
    return d(text="일반 주문하기").exists


def send_auth_request(d, name: str, phone11: str) -> bool:
    if not ensure_on_mobile_order_home(d):
        print("❌ 에뮬레이터: '일반 주문하기'가 보이는 화면이 아닙니다. (화면 전환 안 함)")
        return False

    d(text="일반 주문하기").click()
    time.sleep(0.8)

    if not d(text="주문접수").exists and not d(text="본인인증 요청").exists:
        print("❌ 에뮬레이터: 주문접수(고객정보 입력) 화면이 아닙니다.")
        return False

    if d(text="개인").exists:
        d(text="개인").click()
        time.sleep(0.2)

    edits = d(className="android.widget.EditText")
    if edits.count < 2:
        print("❌ 에뮬레이터: 입력칸(EditText)이 2개가 아닙니다.")
        return False

    edits[0].click()
    time.sleep(0.1)
    edits[0].set_text(name)

    edits[1].click()
    time.sleep(0.1)
    edits[1].set_text(phone11)

    btn = d(text="본인인증 요청")
    for _ in range(20):
        if btn.exists and btn.info.get("enabled", False):
            btn.click()
            time.sleep(0.8)
            print("✅ 에뮬레이터: 본인인증 요청 클릭 완료")
            return True
        time.sleep(0.2)

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
        all_inputs = driver.find_elements(By.TAG_NAME, "input")
        inputs = [x for x in all_inputs if x.is_displayed()]

        if len(inputs) > 15:
            name = pick_name_from_inputs(inputs)
            birth = pick_birth_from_inputs(inputs)
            account = pick_account_from_inputs(inputs)
            zipcode = pick_zip_from_inputs(inputs)

            phone11, phone_raw = pick_phone11_from_inputs(inputs)

            # ✅ 휴대폰을 못 찾으면 "오인 방지"를 위해 스킵 (중요)
            if not phone11:
                print("❌ 전화번호(휴대폰 필드) 추출 실패 → 오인 방지로 건너뜀")
                time.sleep(0.8)
                continue

            if phone11 in processed_phones:
                time.sleep(0.8)
                continue

            processed_phones.add(phone11)

            print("\n✅ 신규 고객 감지")
            print("이름:", name)
            print("생년월일:", birth)
            print("전화번호(11):", phone11)
            print("전화번호 원문:", phone_raw)
            print("계좌:", account)
            print("우편번호:", zipcode)
            print("-" * 40)

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