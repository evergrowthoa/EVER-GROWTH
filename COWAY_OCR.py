from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import time
import re
import threading
import queue
import traceback

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
# 주기/안전 옵션
# ===========================
SIGN_SCAN_INTERVAL_SEC = 20    # 인증완료 스캔 주기(원하면 10~60)
STOP_ON_ERROR = True           # 예상 밖 오류면 중단

# ===========================
# 크롬 옵션
# ===========================
chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_experimental_option("detach", True)

driver = webdriver.Chrome(options=chrome_options)
wait = WebDriverWait(driver, 20)

# ===========================
# 작업 큐 (웹 → 에뮬레이터) : 인증발송 전용
# ===========================
auth_q = queue.Queue()

# 인증발송 성공한 건(= 인증완료에서 매칭할 대상)
# phone11 -> name
auth_sent_jobs = {}
jobs_lock = threading.Lock()

# 전자서명 플로우 진입(중복 방지)
sign_started = set()

# ===========================
# 중단 플래그
# ===========================
STOP_FLAG = False

def set_stop(reason: str):
    global STOP_FLAG
    STOP_FLAG = True
    print("🛑 자동화 중단:", reason)

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

def phone11_to_display(phone11: str) -> str:
    # 01076222590 -> 010-7622-2590
    if not phone11 or len(phone11) != 11:
        return ""
    return f"{phone11[0:3]}-{phone11[3:7]}-{phone11[7:11]}"

# ===========================
# 모달(라벨 기반) 값 추출
# ===========================
DEBUG_MODAL_ON_FAIL = True
_last_debug_ts = 0.0

def find_open_modal():
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
    name = find_input_near_label(modal, ["고객명", "이름", "성명"])
    birth = find_input_near_label(modal, ["생년월일", "주민", "생년"])
    phone_raw = find_input_near_label(modal, ["연락처", "휴대폰번호", "휴대폰", "전화번호", "전화"])
    account = find_input_near_label(modal, ["계좌", "은행", "계좌번호", "결제정보", "결제"])
    zipcode = find_input_near_label(modal, ["우편번호", "우편", "ZIP", "postcode"])

    phone11 = normalize_phone11_only_010(phone_raw)

    return {
        "name": (name or "").strip(),
        "birth": (birth or "").strip(),
        "phone_raw": (phone_raw or "").strip(),
        "phone11": (phone11 or "").strip(),
        "account": (account or "").strip(),
        "zipcode": (zipcode or "").strip(),
    }

# ===========================
# 에뮬레이터 제어(단일 워커에서만 호출)
# ===========================
def connect_emulator():
    d = u2.connect(ADB_SERIAL)
    d.implicitly_wait(5.0)
    return d

def ensure_on_mobile_order_home(d):
    return d(text="일반 주문하기").exists

def ensure_on_order_status_screen(d):
    return d(text="주문현황").exists

def goto_order_status(d) -> bool:
    """
    전자서명 스캔을 위해 주문현황 화면으로 이동
    - 가능하면 건드리지 않고,
    - 홈이면 '일반주문' 클릭으로 주문현황 진입 시도
    """
    if ensure_on_order_status_screen(d):
        return True

    if ensure_on_mobile_order_home(d):
        if d(text="일반주문").exists:
            d(text="일반주문").click()
            time.sleep(0.8)
            return ensure_on_order_status_screen(d)

    return False

def try_click_auth_done_item(d) -> bool:
    """
    주문현황 리스트에서 '인증완료' 항목(핑크)을 하나 클릭해서 상세로 진입.
    인증완료가 화면에 여러 개면 첫 번째만 클릭.
    """
    if not ensure_on_order_status_screen(d):
        return False

    if d(text="인증완료").exists:
        d(text="인증완료").click()
        time.sleep(0.8)
        return True

    return False

def match_detail_by_name_phone(d, name: str, phone11: str) -> bool:
    phone_disp = phone11_to_display(phone11)
    if not phone_disp:
        return False
    return d(text=name).exists and d(text=phone_disp).exists

def click_order_continue(d) -> bool:
    if d(text="주문 이어서 하기").exists:
        d(text="주문 이어서 하기").click()
        time.sleep(0.8)
        return True
    return False

def back_to_order_list(d) -> bool:
    for _ in range(4):
        if ensure_on_order_status_screen(d):
            return True
        d.press("back")
        time.sleep(0.5)
    return ensure_on_order_status_screen(d)

def send_auth_request(d, name: str, phone11: str) -> bool:
    """
    인증발송(A)
    """
    if not ensure_on_mobile_order_home(d):
        print("❌ 에뮬레이터: '일반 주문하기' 화면이 아닙니다.")
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

    if not clicked_auth and btn.exists:
        btn.click()
        time.sleep(0.6)
        print("✅ 에뮬레이터: 본인인증 요청 클릭(강제) 시도 완료")
        clicked_auth = True

    if not clicked_auth:
        print("❌ 에뮬레이터: '본인인증 요청' 버튼을 찾지 못했습니다.")
        return False

    # 확인 팝업 [발송] 클릭
    send_clicked = False
    for _ in range(40):
        if d(text="발송").exists:
            d(text="발송").click()
            time.sleep(0.6)
            print("✅ 에뮬레이터: 확인 팝업 [발송] 클릭 완료")
            send_clicked = True
            break
        time.sleep(0.2)

    if not send_clicked:
        print("⚠️ 에뮬레이터: [발송] 팝업이 안 보였음(환경 차이 가능)")

    return True

def try_start_sign_flow(d) -> bool:
    """
    전자서명(B) 시작점:
    - 주문현황에서 인증완료 하나 열기
    - 상세 화면에서 (이름+연락처)로 auth_sent_jobs 중 미처리 건과 매칭
    - 매칭되면 '주문 이어서 하기' 클릭하고 sign_started에 기록
    - 매칭 안 되면 뒤로가기 후 종료
    """
    if not goto_order_status(d):
        return False

    if not try_click_auth_done_item(d):
        return False

    matched_phone = ""
    matched_name = ""

    with jobs_lock:
        candidates = [(p, n) for (p, n) in auth_sent_jobs.items() if p not in sign_started]

    for phone11, name in candidates:
        if match_detail_by_name_phone(d, name, phone11):
            matched_phone = phone11
            matched_name = name
            break

    if not matched_phone:
        back_to_order_list(d)
        return False

    ok = click_order_continue(d)
    if not ok:
        print("❌ 에뮬레이터: 인증완료 상세에서 '주문 이어서 하기' 버튼을 못 찾음")
        if STOP_ON_ERROR:
            set_stop("전자서명 시작 실패(버튼 없음)")
        return False

    with jobs_lock:
        sign_started.add(matched_phone)

    print(f"✅ 전자서명 플로우 진입(시작): {matched_name} / {matched_phone}")

    # 여기서부터 제품/색상/약정/관리/전자서명발송 단계로 이어 붙이면 됨.
    # 현재는 '주문 이어서 하기' 진입까지만 안전하게 처리.
    return True

# ===========================
# 에뮬레이터 단일 워커 (겹침 방지 핵심)
# ===========================
def emulator_main_loop():
    if not DO_EMULATOR:
        print("ℹ️ 에뮬레이터 자동화가 꺼져 있습니다(DO_EMULATOR=False).")
        while True:
            time.sleep(1.0)

    d = connect_emulator()
    print("✅ 에뮬레이터 연결 완료:", ADB_SERIAL)

    last_sign_scan = 0.0

    while True:
        try:
            if STOP_FLAG:
                time.sleep(1.0)
                continue

            # 1) 전자서명(B) 스캔 우선
            now = time.time()
            if now - last_sign_scan >= SIGN_SCAN_INTERVAL_SEC:
                last_sign_scan = now
                try_start_sign_flow(d)

            # 2) 전자서명 시작할 게 없을 때만 인증발송(A) 처리
            try:
                job = auth_q.get_nowait()
            except queue.Empty:
                job = None

            if job is not None:
                try:
                    name = job.get("name", "")
                    phone11 = job.get("phone11", "")
                    print(f"🚀 인증발송 작업 시작: {name} / {phone11}")

                    ok = send_auth_request(d, name, phone11)
                    if ok:
                        with jobs_lock:
                            auth_sent_jobs[phone11] = name
                        print(f"✅ 인증발송 성공: {name} / {phone11}")
                    else:
                        print(f"❌ 인증발송 실패: {name} / {phone11}")
                        if STOP_ON_ERROR:
                            set_stop(f"인증발송 실패: {name}/{phone11}")

                finally:
                    auth_q.task_done()

            time.sleep(0.2)

        except Exception as e:
            print("❌ 에뮬레이터 메인 루프 에러:", e)
            traceback.print_exc()
            if STOP_ON_ERROR:
                set_stop("에뮬레이터 메인 루프 예외")
            time.sleep(1.0)

t_emul = threading.Thread(target=emulator_main_loop, daemon=True)
t_emul.start()

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
        if STOP_FLAG:
            time.sleep(1.0)
            continue

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

        auth_q.put({
            "name": name,
            "phone11": phone11,
            "birth": birth,
            "account": account,
            "zipcode": zipcode,
        })

        time.sleep(1.2)

    except Exception as e:
        print("에러 발생:", e)
        traceback.print_exc()
        if STOP_ON_ERROR:
            set_stop("웹 모달 루프 예외")
        time.sleep(1.0)