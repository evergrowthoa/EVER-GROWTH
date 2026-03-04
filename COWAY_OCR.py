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
ADB_SERIAL = "emulator-5554"
DO_EMULATOR = True

# ===========================
# 디지털세일즈 앱 설정 (런처 복구용)
# ===========================
DIGITAL_SALES_APP_NAME = "디지털세일즈"
DIGITAL_SALES_PACKAGE = ""   # 패키지명 알면 넣으면 더 안정적(모르면 빈칸 유지)

# ===========================
# 주기/안전 옵션
# ===========================
SIGN_SCAN_INTERVAL_SEC = 20
ORDER_REFRESH_INTERVAL_SEC = 60
STOP_ON_ERROR = True
AUTH_RETRY_MAX = 5

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
# phone11 -> job(dict)
auth_sent_jobs = {}
jobs_lock = threading.Lock()

# 전자서명 플로우 진입(중복 방지)
sign_started = set()

# ===========================
# 중단 플래그
# ===========================
STOP_FLAG = False

def notify(msg: str):
    # TODO: 텔레그램/슬랙 붙일 자리 (지금은 출력만)
    print("🔔 알림:", msg)

def set_stop(reason: str):
    global STOP_FLAG
    STOP_FLAG = True
    print("🛑 자동화 중단:", reason)
    notify(reason)

# ===========================
# 공통 유틸
# ===========================
def normalize_digits(s: str) -> str:
    return re.sub(r"\D", "", str(s or ""))

def normalize_phone11_only_010(s: str) -> str:
    digits = normalize_digits(s)
    if digits.startswith("82") and len(digits) >= 12 and digits[2:4] == "10":
        digits = "0" + digits[2:]
    if len(digits) == 11 and digits.startswith("010"):
        return digits
    return ""

def phone11_to_display(phone11: str) -> str:
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

    print("----- [DEBUG] 모달 input 목록(앞 60개) -----")
    try:
        els = modal.find_elements(By.CSS_SELECTOR, "input")
        for i, el in enumerate(els[:60]):
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

def click_text_center(d, txt: str, y_min_ratio: float = 0.0, y_max_ratio: float = 1.0) -> bool:
    """
    텍스트를 '클릭 가능한지'에 의존하지 않고,
    해당 텍스트 요소의 bounds 중앙 좌표를 클릭한다.
    """
    try:
        w, h = d.window_size()
        objs = d(text=txt).all()
        for o in objs:
            try:
                info = o.info
                b = info.get("bounds", {})
                left = int(b.get("left", 0))
                top = int(b.get("top", 0))
                right = int(b.get("right", 0))
                bottom = int(b.get("bottom", 0))
                cx = (left + right) // 2
                cy = (top + bottom) // 2

                if cy < int(h * y_min_ratio) or cy > int(h * y_max_ratio):
                    continue

                d.click(cx, cy)
                return True
            except Exception:
                continue
    except Exception:
        pass

    try:
        if d(text=txt).exists:
            d(text=txt).click()
            return True
    except Exception:
        return False

    return False

def handle_exit_popup_cancel(d) -> bool:
    """
    '모바일 주문을 종료하시겠습니까?' 팝업이 뜨면 무조건 취소해서
    작업이 끊기지 않게 막는다.
    """
    try:
        if d(text="취소").exists and d(text="확인").exists and d(textContains="모바일 주문").exists and d(textContains="종료").exists:
            click_text_center(d, "취소", 0.40, 0.98)
            time.sleep(0.6)
            return True
    except Exception:
        return False
    return False

def click_bottom_sheet_option(d, txt: str) -> bool:
    """
    진행상태 '하단 패널(바텀시트)' 안에 있는 옵션을 정확히 누르기.
    - 같은 txt가 리스트에도 있을 수 있으니,
      화면 아래쪽(하단 45% 이후)에 있는 txt만 클릭한다.
    """
    try:
        w, h = d.window_size()
        objs = d(text=txt).all()
        best = None
        best_top = 10**9

        for o in objs:
            try:
                b = o.info.get("bounds", {})
                top = int(b.get("top", 0))
                if top < int(h * 0.45):
                    continue
                if top < best_top:
                    best_top = top
                    best = o
            except Exception:
                continue

        if best is None:
            return False

        b = best.info.get("bounds", {})
        cx = (int(b.get("left", 0)) + int(b.get("right", 0))) // 2
        cy = (int(b.get("top", 0)) + int(b.get("bottom", 0))) // 2
        d.click(cx, cy)
        time.sleep(0.8)
        return True

    except Exception:
        return False

def open_digital_sales_app(d) -> bool:
    if d(text="모바일 주문").exists or d(text="전체메뉴").exists or d(text="라이프스토리").exists:
        return True

    if DIGITAL_SALES_PACKAGE:
        try:
            d.app_start(DIGITAL_SALES_PACKAGE)
            time.sleep(2.0)
            if d(text="모바일 주문").exists or d(text="전체메뉴").exists:
                return True
        except Exception:
            pass

    try:
        d.press("home")
        time.sleep(0.8)
    except Exception:
        pass

    if d(text=DIGITAL_SALES_APP_NAME).exists:
        click_text_center(d, DIGITAL_SALES_APP_NAME)
        time.sleep(2.0)
        return True

    return False

def ensure_on_mobile_order_home(d) -> bool:
    """
    주문현황에 있어도 인증발송이 가능하도록:
    주문현황 X 클릭 → 종료 팝업 '확인' 클릭 → 모바일 주문 홈
    """
    for _ in range(10):
        if d(text="일반 주문하기").exists:
            return True

        if d(text="주문현황").exists:
            try:
                w, h = d.window_size()
                d.click(int(w * 0.965), int(h * 0.075))
                time.sleep(0.3)
            except Exception:
                pass

            for _ in range(10):
                if d(text="확인").exists and d(text="취소").exists:
                    click_text_center(d, "확인", 0.45, 0.95)
                    time.sleep(0.8)
                    break
                time.sleep(0.2)
            continue

        if d(text="확인").exists and d(text="취소").exists:
            click_text_center(d, "확인", 0.45, 0.95)
            time.sleep(0.8)
            continue

        if d(text="모바일 주문").exists:
            click_text_center(d, "모바일 주문", 0.70, 0.98)
            time.sleep(1.0)
            if d(text="일반 주문하기").exists:
                return True

        opened = open_digital_sales_app(d)
        if opened and d(text="모바일 주문").exists:
            click_text_center(d, "모바일 주문", 0.70, 0.98)
            time.sleep(1.0)
            if d(text="일반 주문하기").exists:
                return True

        time.sleep(0.4)

    return d(text="일반 주문하기").exists

def ensure_on_order_status_screen(d):
    return d(text="주문현황").exists

def goto_order_status(d) -> bool:
    if ensure_on_order_status_screen(d):
        return True
    if not ensure_on_mobile_order_home(d):
        return False
    if d(text="일반주문").exists:
        click_text_center(d, "일반주문", 0.35, 0.70)
        time.sleep(1.0)
        return ensure_on_order_status_screen(d)
    return False

def click_refresh_on_order_status(d) -> bool:
    """
    우측 원형 새로고침 버튼은 텍스트가 없으니 좌표로 누른다.
    (종료 팝업이 뜨면 무조건 취소)
    """
    try:
        if not d(text="주문현황").exists:
            return False

        handle_exit_popup_cancel(d)

        w, h = d.window_size()
        for ry in (0.265, 0.285, 0.305):
            d.click(int(w * 0.955), int(h * ry))
            time.sleep(0.6)
            handle_exit_popup_cancel(d)
            return True
    except Exception:
        return False
    return False

def ensure_filter_auth_done(d) -> bool:
    """
    ✅ 진행상태 필터 패널에서 '인증완료'를 반드시 선택.
    - back 사용 금지(종료 팝업 트리거)
    - 하단 패널 영역의 '인증완료'만 클릭
    - 패널 닫기는 상단 빈공간 클릭
    """
    if not ensure_on_order_status_screen(d):
        return False

    handle_exit_popup_cancel(d)

    if d(text="진행상태").exists:
        click_text_center(d, "진행상태", 0.20, 0.45)
        time.sleep(0.7)
    else:
        try:
            w, h = d.window_size()
            d.click(int(w * 0.18), int(h * 0.335))
            time.sleep(0.7)
        except Exception:
            return False

    if not d(text="인증완료").exists:
        return False

    ok = click_bottom_sheet_option(d, "인증완료")
    if not ok:
        return False

    try:
        w, h = d.window_size()
        d.click(int(w * 0.50), int(h * 0.18))
        time.sleep(0.4)
    except Exception:
        pass

    handle_exit_popup_cancel(d)
    return True

def click_auth_done_buttons_in_list(d):
    objs = []
    try:
        w, h = d.window_size()
        for o in d(text="인증완료").all():
            try:
                b = o.info.get("bounds", {})
                top = int(b.get("top", 0))
                if top >= int(h * 0.40):
                    objs.append((top, o))
            except Exception:
                continue
        objs.sort(key=lambda x: x[0])
    except Exception:
        pass
    return [o for _, o in objs]

def match_detail_by_name_phone_strict(d, name: str, phone11: str) -> bool:
    if not name or len(name.strip()) < 2:
        return False
    phone_disp = phone11_to_display(phone11)
    if not phone_disp:
        return False
    return d(text=name).exists and d(text=phone_disp).exists

def back_to_order_list(d) -> bool:
    for _ in range(4):
        if d(text="주문현황").exists:
            return True
        try:
            d.press("back")
        except Exception:
            pass
        time.sleep(0.5)
        handle_exit_popup_cancel(d)
    return d(text="주문현황").exists

def click_order_continue(d) -> bool:
    if d(text="주문 이어서 하기").exists:
        click_text_center(d, "주문 이어서 하기", 0.35, 0.90)
        time.sleep(0.8)
        return True
    return False

def send_auth_request(d, name: str, phone11: str):
    if not ensure_on_mobile_order_home(d):
        return (False, "NOT_READY")

    click_text_center(d, "일반 주문하기", 0.20, 0.55)
    time.sleep(0.8)

    if not d(text="주문접수").exists and not d(text="본인인증 요청").exists:
        return (False, "NOT_ORDER_FORM")

    if d(text="개인").exists:
        click_text_center(d, "개인", 0.15, 0.45)
        time.sleep(0.2)

    edits = d(className="android.widget.EditText")
    if edits.count < 2:
        return (False, "NO_EDITTEXT")

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
            clicked_auth = True
            break
        time.sleep(0.2)

    if not clicked_auth and btn.exists:
        btn.click()
        time.sleep(0.6)
        clicked_auth = True

    if not clicked_auth:
        return (False, "NO_AUTH_BTN")

    for _ in range(40):
        if d(text="발송").exists:
            click_text_center(d, "발송", 0.45, 0.95)
            time.sleep(0.6)
            break
        time.sleep(0.2)

    return (True, "OK")

def try_start_sign_flow(d) -> bool:
    if not goto_order_status(d):
        return False

    if not ensure_filter_auth_done(d):
        return False

    click_refresh_on_order_status(d)

    with jobs_lock:
        candidates = [(p, j) for (p, j) in auth_sent_jobs.items() if p not in sign_started]
    if not candidates:
        return False

    buttons = click_auth_done_buttons_in_list(d)
    if not buttons:
        return False

    for btn in buttons[:5]:
        try:
            b = btn.info.get("bounds", {})
            cx = (int(b.get("left", 0)) + int(b.get("right", 0))) // 2
            cy = (int(b.get("top", 0)) + int(b.get("bottom", 0))) // 2
            d.click(cx, cy)
            time.sleep(0.9)

            matched_phone = ""
            matched_job = None

            for phone11, job in candidates:
                if match_detail_by_name_phone_strict(d, job.get("name", ""), phone11):
                    matched_phone = phone11
                    matched_job = job
                    break

            if not matched_phone:
                back_to_order_list(d)
                continue

            ok = click_order_continue(d)
            if not ok:
                if STOP_ON_ERROR:
                    set_stop("매칭 성공했는데 '주문 이어서 하기' 버튼을 못 찾음")
                return False

            with jobs_lock:
                sign_started.add(matched_phone)

            notify(f"전자서명 진입(매칭 OK): {matched_job.get('name','')} / {matched_phone}")
            return True

        except Exception:
            back_to_order_list(d)
            continue

    return False

# ===========================
# 에뮬레이터 단일 워커 (겹침 방지)
# ===========================
def emulator_main_loop():
    if not DO_EMULATOR:
        print("ℹ️ 에뮬레이터 자동화가 꺼져 있습니다(DO_EMULATOR=False).")
        while True:
            time.sleep(1.0)

    d = connect_emulator()
    print("✅ 에뮬레이터 연결 완료:", ADB_SERIAL)

    last_sign_scan = 0.0
    last_refresh = 0.0

    while True:
        try:
            if STOP_FLAG:
                time.sleep(1.0)
                continue

            now = time.time()
            auth_pending = not auth_q.empty()

            if (not auth_pending) and (now - last_refresh >= ORDER_REFRESH_INTERVAL_SEC):
                if goto_order_status(d):
                    ensure_filter_auth_done(d)
                    click_refresh_on_order_status(d)
                last_refresh = now

            if now - last_sign_scan >= SIGN_SCAN_INTERVAL_SEC:
                last_sign_scan = now
                try_start_sign_flow(d)

            try:
                job = auth_q.get_nowait()
            except queue.Empty:
                job = None

            if job is not None:
                try:
                    name = job.get("name", "")
                    phone11 = job.get("phone11", "")
                    retry = int(job.get("_retry", 0))
                    print(f"🚀 인증발송 작업 시작: {name} / {phone11}")

                    ok, reason = send_auth_request(d, name, phone11)
                    if ok:
                        with jobs_lock:
                            auth_sent_jobs[phone11] = job
                        print(f"✅ 인증발송 성공: {name} / {phone11}")

                        if goto_order_status(d):
                            ensure_filter_auth_done(d)
                            click_refresh_on_order_status(d)

                    else:
                        if reason == "NOT_READY" and retry < AUTH_RETRY_MAX:
                            retry += 1
                            job["_retry"] = retry
                            print(f"⚠️ 인증발송 재시도 ({retry}/{AUTH_RETRY_MAX}) : {name} / {phone11}")
                            auth_q.put(job)
                        else:
                            if STOP_ON_ERROR:
                                set_stop(f"인증발송 실패: {name}/{phone11} ({reason})")

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

        name = data.get("name", "")
        birth = data.get("birth", "")
        phone_raw = data.get("phone_raw", "")
        phone11 = data.get("phone11", "")
        account = data.get("account", "")
        zipcode = data.get("zipcode", "")

        if not phone11:
            print("❌ 전화번호 추출 실패 → 건너뜀")
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
            "_retry": 0,
        })

        time.sleep(1.2)

    except Exception as e:
        print("에러 발생:", e)
        traceback.print_exc()
        if STOP_ON_ERROR:
            set_stop("웹 모달 루프 예외")
        time.sleep(1.0)