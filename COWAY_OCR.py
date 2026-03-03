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
ADB_SERIAL = "emulator-5554"   # adb devices에서 보이는 값
DO_EMULATOR = True            # 테스트로 에뮬레이터 끄려면 False

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
# 공통 유틸
# ===========================
def normalize_digits(s: str) -> str:
    return re.sub(r"\D", "", str(s or ""))

def normalize_phone11_only_010(s: str) -> str:
    digits = normalize_digits(s)

    # +82 10xxxxxxxx → 010xxxxxxxx 보정(혹시)
    if digits.startswith("82") and len(digits) >= 12 and digits[2:4] == "10":
        digits = "0" + digits[2:]

    # 여기서는 010 11자리만 유효
    if len(digits) == 11 and digits.startswith("010"):
        return digits
    return ""

# ===========================
# 모달(라벨 기반) 값 추출
# ===========================
DEBUG_MODAL_ON_FAIL = True
_last_debug_ts = 0.0

def find_open_modal():
    # 모달이 부트스트랩 계열이라면 아래 셀렉터 중 하나로 잡힘
    selectors = [
        ".modal.show",
        ".modal.in",
        ".modal[style*='display: block']",
        ".modal[style*='display:block']",
    ]
    for sel in selectors:
        ms = driver.find_elements(By.CSS_SELECTOR, sel)
        for m in ms:
            try:
                if m.is_displayed():
                    return m
            except Exception:
                continue
    return None

def find_input_near_label(modal, label_keywords):
    """
    label_keywords(예: ["연락처","휴대폰"]) 중 하나를 포함하는 텍스트를 찾고,
    그 주변(같은 영역)에서 가장 가까운 input을 찾아 value를 반환.
    """
    # 1) (가장 흔함) 라벨 텍스트 뒤에 바로 input이 오는 케이스
    for kw in label_keywords:
        xpaths = [
            f".//label[contains(normalize-space(.), '{kw}')]/following::input[not(@type='hidden')][1]",
            f".//*[contains(normalize-space(.), '{kw}')]/following::input[not(@type='hidden')][1]",
            f".//tr[.//*[contains(normalize-space(.), '{kw}')]]//input[not(@type='hidden')][1]",
        ]
        for xp in xpaths:
            try:
                el = modal.find_element(By.XPATH, xp)
                if el and el.is_displayed():
                    v = (el.get_attribute("value") or "").strip()
                    if v:
                        return v
            except Exception:
                pass

    # 2) 같은 “행/그룹” 안에서 input 찾기(조금 더 안전)
    for kw in label_keywords:
        try:
            label_el = modal.find_element(By.XPATH, f".//*[contains(normalize-space(.), '{kw}')]")
        except Exception:
            continue

        try:
            container = label_el.find_element(
                By.XPATH,
                "./ancestor::*[contains(@class,'form-group') or contains(@class,'row') or contains(@class,'col') or self::tr][1]"
            )
            cand = container.find_elements(By.XPATH, ".//input[not(@type='hidden')]")
            for el in cand:
                try:
                    if el.is_displayed():
                        v = (el.get_attribute("value") or "").strip()
                        if v:
                            return v
                except Exception:
                    continue
        except Exception:
            pass

    return ""

def debug_modal_inputs(modal):
    global _last_debug_ts
    now = time.time()
    if not DEBUG_MODAL_ON_FAIL:
        return
    if now - _last_debug_ts < 2.0:
        return
    _last_debug_ts = now

    print("----- [DEBUG] 모달 input 목록(앞 40개) -----")
    try:
        els = modal.find_elements(By.CSS_SELECTOR, "input")
        for i, el in enumerate(els[:40]):
            try:
                if not el.is_displayed():
                    continue
                v = (el.get_attribute("value") or "").strip()
                if not v:
                    continue
                _id = (el.get_attribute("id") or "").strip()
                _name = (el.get_attribute("name") or "").strip()
                _ph = (el.get_attribute("placeholder") or "").strip()
                _type = (el.get_attribute("type") or "").strip()
                print(f"[{i}] type={_type} id={_id} name={_name} placeholder={_ph} value={v}")
            except Exception:
                continue
    except Exception as e:
        print("DEBUG 실패:", e)
    print("----- [DEBUG] 끝 -----")

def extract_fields_from_modal(modal):
    # 라벨 기반으로 최대한 정확히 찾기
    name = find_input_near_label(modal, ["고객명", "이름", "성명"])
    birth = find_input_near_label(modal, ["생년월일", "주민", "생년"])
    phone_raw = find_input_near_label(modal, ["연락처", "휴대폰번호", "휴대폰", "전화번호", "전화"])
    account = find_input_near_label(modal, ["계좌", "은행", "계좌번호", "결제정보", "결제"])
    zipcode = find_input_near_label(modal, ["우편번호", "우편", "ZIP", "postcode"])

    phone11 = normalize_phone11_only_010(phone_raw)

    return {
        "name": name,
        "birth": birth,
        "phone_raw": phone_raw,
        "phone11": phone11,
        "account": account,
        "zipcode": zipcode,
    }

# ===========================
# 에뮬레이터 제어
# ===========================
def connect_emulator():
    d = u2.connect(ADB_SERIAL)
    d.implicitly_wait(5.0)
    return d

def ensure_on_mobile_order_home(d):
    # ✅ 화면을 건드리지 않고 확인만
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

    clicked_auth = False
    for _ in range(20):
        if btn.exists and btn.info.get("enabled", False):
            btn.click()
            time.sleep(0.6)
            print("✅ 에뮬레이터: 본인인증 요청 클릭 완료")
            clicked_auth = True
            break
        time.sleep(0.2)

    if not clicked_auth:
        if btn.exists:
            btn.click()
            time.sleep(0.6)
            print("✅ 에뮬레이터: 본인인증 요청 클릭(강제) 시도 완료")
            clicked_auth = True

    if not clicked_auth:
        print("❌ 에뮬레이터: '본인인증 요청' 버튼을 찾지 못했습니다.")
        return False

    # ✅ 확인 팝업이 뜨면 [발송]까지 자동 클릭
    # 팝업이 안 뜨는 환경도 있으니: "있으면 누른다" 방식으로 안전하게 처리
    send_clicked = False
    for _ in range(40):
        title_ok = d(textContains="메시지를 발송").exists or d(textContains="고객인증").exists
        if title_ok and d(text="발송").exists:
            d(text="발송").click()
            time.sleep(0.6)
            print("✅ 에뮬레이터: 확인 팝업 [발송] 클릭 완료")
            send_clicked = True
            break
        time.sleep(0.2)

    # 팝업 타이틀을 못 잡는 기기 대비(타이틀 없이 버튼만 잡히는 경우)
    if not send_clicked:
        for _ in range(10):
            if d(text="발송").exists:
                d(text="발송").click()
                time.sleep(0.6)
                print("✅ 에뮬레이터: [발송] 클릭(타이틀 미확인) 완료")
                send_clicked = True
                break
            time.sleep(0.2)

    # 팝업이 있었으면 닫힐 때까지 짧게 대기(중복 클릭 방지)
    if send_clicked:
        for _ in range(20):
            if not d(text="발송").exists:
                break
            time.sleep(0.2)

    return True

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
        modal = find_open_modal()
        if not modal:
            time.sleep(0.5)
            continue

        data = extract_fields_from_modal(modal)

        name = (data.get("name") or "").strip()
        birth = (data.get("birth") or "").strip()
        phone_raw = (data.get("phone_raw") or "").strip()
        phone11 = (data.get("phone11") or "").strip()
        account = (data.get("account") or "").strip()
        zipcode = (data.get("zipcode") or "").strip()

        # ✅ 전화번호(라벨 기반) 추출 실패 시: 오인 방지로 스킵 + 디버그 출력
        if not phone11:
            print("❌ 전화번호(라벨 기반) 추출 실패 → 오인 방지로 건너뜀")
            debug_modal_inputs(modal)
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

    except Exception as e:
        print("에러 발생:", e)
        time.sleep(1)