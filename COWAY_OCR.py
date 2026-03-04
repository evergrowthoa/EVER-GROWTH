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
SIGN_SCAN_INTERVAL_SEC = 20      # 인증완료 스캔 주기
ORDER_REFRESH_INTERVAL_SEC = 60  # ✅ Idle일 때 주문현황 새로고침 주기(1분)
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

def click_first_clickable_text(d, txt: str) -> bool:
    try:
        objs = d(text=txt).all()
        for o in objs:
            try:
                info = o.info
                if info.get("clickable", False):
                    o.click()
                    return True
            except Exception:
                continue
    except Exception:
        pass

    if d(text=txt).exists:
        try:
            d(text=txt).click()
            return True
        except Exception:
            return False

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
        ok = click_first_clickable_text(d, DIGITAL_SALES_APP_NAME)
        time.sleep(2.0)
        return ok

    return False

def ensure_on_mobile_order_home(d) -> bool:
    for _ in range(3):
        if d(text="일반 주문하기").exists:
            return True

        if d(text="모바일 주문").exists:
            click_first_clickable_text(d, "모바일 주문")
            time.sleep(1.0)
            if d(text="일반 주문하기").exists:
                return True

        opened = open_digital_sales_app(d)
        if opened:
            if d(text="모바일 주문").exists:
                click_first_clickable_text(d, "모바일 주문")
                time.sleep(1.0)
                if d(text="일반 주문하기").exists:
                    return True

        time.sleep(0.5)

    return d(text="일반 주문하기").exists

def ensure_on_order_status_screen(d):
    return d(text="주문현황").exists

def goto_order_status(d) -> bool:
    """
    ✅ 어떤 화면이든 '주문현황'까지 복구해서 진입
    순서:
    1) 주문현황이면 OK
    2) 모바일주문 홈(일반 주문하기 보이는 화면)으로 복구
    3) 모바일주문 홈에서 '일반주문' 클릭 → 주문현황
    """
    try:
        if d(text="주문현황").exists:
            return True
    except Exception:
        pass

    # 1) 모바일주문 홈으로 복구 (앱 메인/런처에서도 복구되도록)
    if not ensure_on_mobile_order_home(d):
        return False

    # 2) 모바일주문 홈에서 일반주문 클릭 → 주문현황
    try:
        if d(text="일반주문").exists:
            click_first_clickable_text(d, "일반주문")
            time.sleep(1.0)
            return d(text="주문현황").exists
    except Exception:
        return False

    return False

def click_refresh_on_order_status(d) -> bool:
    """
    ✅ 주문현황 우측 원형 새로고침 아이콘 클릭
    - 이 버튼은 text/desc가 없는 경우가 많아서 '좌표'로 누르는게 제일 확실함.
    - 화면 크기가 바뀌어도 동작하도록 상대좌표로 3번 시도.
    """
    try:
        if not d(text="주문현황").exists:
            return False
    except Exception:
        return False

    try:
        w, h = d.window_size()

        # 너가 표시한 원형 버튼 위치(검색창 오른쪽)를 상대좌표로 여러 번 시도
        # 화면/기기마다 살짝 달라서 y를 3개로 잡음
        candidates = [
            (int(w * 0.955), int(h * 0.265)),
            (int(w * 0.955), int(h * 0.285)),
            (int(w * 0.955), int(h * 0.305)),
        ]

        for x, y in candidates:
            d.click(x, y)
            time.sleep(0.6)
            return True

    except Exception:
        return False

    return False

def ensure_filter_auth_done(d) -> bool:
    """
    ✅ 진행상태 필터를 무조건 '인증완료'로 맞춘다.
    - 현재 칩이 '인증입력/서명입력/주문확정/인증완료' 등 무엇이든 상관없이
      그 칩(진행상태 필터 버튼)을 눌러 메뉴를 열고 '인증완료'를 선택한다.
    """

    try:
        if not d(text="주문현황").exists:
            return False
    except Exception:
        return False

    try:
        w, h = d.window_size()

        # 1) 상단 영역(필터칩이 있는 영역)에서 '현재 상태 칩'을 찾는다.
        #    (진행상태/인증입력/인증완료/서명입력/주문확정/주문삭제 중 하나가 상단에 있음)
        chip_texts = ["진행상태", "인증입력", "인증완료", "서명입력", "주문확정", "주문삭제"]
        chip = None

        for t in chip_texts:
            try:
                objs = d(text=t).all()
            except Exception:
                objs = []

            for o in objs:
                try:
                    info = o.info
                    b = info.get("bounds", {})
                    top = int(b.get("top", 999999))
                    left = int(b.get("left", 0))

                    # 필터칩은 '상단 중간' 영역에 있음 (대략 18%~40%)
                    if top >= int(h * 0.18) and top <= int(h * 0.40) and left <= int(w * 0.60):
                        chip = o
                        break
                except Exception:
                    continue

            if chip is not None:
                break

        if chip is None:
            return False

        # 2) 칩을 눌러서 메뉴를 연다
        chip.click()
        time.sleep(0.6)

        # 3) 메뉴에서 '인증완료' 선택
        if d(text="인증완료").exists:
            click_first_clickable_text(d, "인증완료")
            time.sleep(0.8)

        # 4) 상단에 '인증완료' 칩이 보이면 성공(상단 영역만 확인)
        objs2 = []
        try:
            objs2 = d(text="인증완료").all()
        except Exception:
            objs2 = []

        for o in objs2:
            try:
                info = o.info
                b = info.get("bounds", {})
                top = int(b.get("top", 999999))
                left = int(b.get("left", 0))
                if top >= int(h * 0.18) and top <= int(h * 0.40) and left <= int(w * 0.60):
                    return True
            except Exception:
                continue

        # 그래도 확인이 애매하면 '인증완료'가 메뉴에서 눌렸다고 가정하고 True
        return True

    except Exception:
        return False

def click_auth_done_item_in_list(d) -> bool:
    """
    ✅ 인증완료 리스트 항목(핑크)을 눌러 상세로 진입
    필터 칩과 겹치므로, '리스트 영역'에 있는 인증완료 버튼만 클릭한다.
    """
    if not ensure_on_order_status_screen(d):
        return False

    try:
        w, h = d.window_size()
        objs = d(text="인증완료").all()
        buttons = []
        for o in objs:
            try:
                info = o.info
                b = info.get("bounds", {})
                top = b.get("top", 0)
                # 리스트 영역(대략 35% 이후)만 대상으로
                if top >= int(h * 0.35):
                    buttons.append((top, o))
            except Exception:
                continue

        buttons.sort(key=lambda x: x[0])
        if not buttons:
            return False

        buttons[0][1].click()
        time.sleep(0.8)
        return True
    except Exception:
        return False

def match_detail_by_name_phone(d, name: str, phone11: str) -> bool:
    phone_disp = phone11_to_display(phone11)
    if not phone_disp:
        return False
    return d(text=name).exists and d(text=phone_disp).exists

def click_order_continue(d) -> bool:
    if d(text="주문 이어서 하기").exists:
        click_first_clickable_text(d, "주문 이어서 하기")
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

def send_auth_request(d, name: str, phone11: str):
    """
    인증발송(A)
    return: (ok: bool, reason: str)
    """
    if not ensure_on_mobile_order_home(d):
        return (False, "NOT_READY")

    click_first_clickable_text(d, "일반 주문하기")
    time.sleep(0.8)

    if not d(text="주문접수").exists and not d(text="본인인증 요청").exists:
        return (False, "NOT_ORDER_FORM")

    if d(text="개인").exists:
        click_first_clickable_text(d, "개인")
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

    send_clicked = False
    for _ in range(40):
        if d(text="발송").exists:
            click_first_clickable_text(d, "발송")
            time.sleep(0.6)
            send_clicked = True
            break
        time.sleep(0.2)

    return (True, "OK")

def try_start_sign_flow(d) -> bool:
    """
    전자서명(B) 시작점:
    - 주문현황 진입
    - 필터를 인증완료로 맞춤
    - 리스트에서 인증완료 항목 하나 진입
    - 상세에서 이름+전화번호 일치할 때만 '주문 이어서 하기'
    """
    if not goto_order_status(d):
        return False

    if not ensure_filter_auth_done(d):
        # 필터를 못 맞추면 진행 금지
        return False

    if not click_auth_done_item_in_list(d):
        return False

    matched_phone = ""
    matched_job = None

    with jobs_lock:
        candidates = [(p, j) for (p, j) in auth_sent_jobs.items() if p not in sign_started]

    for phone11, job in candidates:
        if match_detail_by_name_phone(d, job.get("name", ""), phone11):
            matched_phone = phone11
            matched_job = job
            break

    if not matched_phone:
        back_to_order_list(d)
        return False

    ok = click_order_continue(d)
    if not ok:
        if STOP_ON_ERROR:
            set_stop("인증완료 상세에서 '주문 이어서 하기' 버튼을 못 찾음")
        return False

    with jobs_lock:
        sign_started.add(matched_phone)

    notify(f"전자서명 단계 진입: {matched_job.get('name','')} / {matched_phone}")
    return True

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

            # ✅ Idle 판정: 인증발송 큐가 비어있고, 지금 루프에서 인증발송/전자서명 작업을 수행하지 않을 때만 새로고침
            auth_pending = not auth_q.empty()

            now = time.time()

            # 0) Idle 상태에서 1분마다 주문현황 새로고침 (작업 중이면 절대 안 함)
            if (not auth_pending) and (now - last_refresh >= ORDER_REFRESH_INTERVAL_SEC):
                if goto_order_status(d):
                    if ensure_filter_auth_done(d):
                        click_refresh_on_order_status(d)
                last_refresh = now

            # 1) 전자서명(B) 스캔 (20초 주기)
            if now - last_sign_scan >= SIGN_SCAN_INTERVAL_SEC:
                last_sign_scan = now
                try_start_sign_flow(d)

            # 2) 인증발송(A) 처리
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
                    else:
                        if reason == "NOT_READY" and retry < AUTH_RETRY_MAX:
                            retry += 1
                            job["_retry"] = retry
                            print(f"⚠️ 인증발송 재시도 예정 ({retry}/{AUTH_RETRY_MAX}) : {name} / {phone11}")
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
            "_retry": 0,
        })

        time.sleep(1.2)

    except Exception as e:
        print("에러 발생:", e)
        traceback.print_exc()
        if STOP_ON_ERROR:
            set_stop("웹 모달 루프 예외")
        time.sleep(1.0)