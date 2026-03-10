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

BUILD_ID = "COWAY_OCR_BUILD_2026-03-05_006"
print("✅ BUILD:", BUILD_ID)

USER_ID = "evergrowth"
USER_PW = "wkrlfjqm1!"

ADB_SERIAL = "emulator-5554"
DO_EMULATOR = True

DIGITAL_SALES_APP_NAME = "디지털세일즈"
DIGITAL_SALES_PACKAGE = ""

STOP_ON_ERROR = True
AUTH_RETRY_MAX = 3

ORDER_REENTER_INTERVAL_SEC = 90
SEARCH_BATCH_PER_CYCLE = 3
SEARCH_LOOP_SLEEP_SEC = 2.0

chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_experimental_option("detach", True)

driver = webdriver.Chrome(options=chrome_options)
wait = WebDriverWait(driver, 20)

auth_q = queue.Queue()
auth_sent_jobs = {}
jobs_lock = threading.Lock()
sign_started = set()

SIGN_IN_PROGRESS = False
STOP_FLAG = False

def notify(msg: str):
    print("🔔", msg)

def set_stop(reason: str):
    global STOP_FLAG
    STOP_FLAG = True
    print("🛑 자동화 중단:", reason)
    notify(reason)

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

def compute_check_interval(auth_sent_at: float) -> int:
    now = time.time()
    elapsed = max(0, now - auth_sent_at)
    if elapsed < 10 * 60:
        return 120
    if elapsed < 60 * 60:
        return 300
    if elapsed < 6 * 60 * 60:
        return 900
    if elapsed < 24 * 60 * 60:
        return 1800
    return 3600

# ---------------------------
# Selenium modal
# ---------------------------
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
    def _read_value(el):
        try:
            v = (el.get_attribute("value") or "").strip()
            if v:
                return v
        except Exception:
            pass

        try:
            v = (el.get_attribute("textContent") or "").strip()
            if v:
                return v
        except Exception:
            pass

        try:
            v = (el.get_attribute("innerText") or "").strip()
            if v:
                return v
        except Exception:
            pass

        try:
            v = (el.text or "").strip()
            if v:
                return v
        except Exception:
            pass

        return ""

    for kw in label_keywords:
        xpaths = [
            f".//label[contains(normalize-space(.), '{kw}')]/following::*[((self::input or self::textarea or self::select) and not(@type='hidden'))][1]",
            f".//*[contains(normalize-space(.), '{kw}')]/following::*[((self::input or self::textarea or self::select) and not(@type='hidden'))][1]",
            f".//tr[.//*[contains(normalize-space(.), '{kw}')]]//*[ (self::input or self::textarea or self::select) and not(@type='hidden') ][1]",
            f".//label[contains(normalize-space(.), '{kw}')]/following::*[self::td or self::span or self::div or self::p][1]",
            f".//*[contains(normalize-space(.), '{kw}')]/following::*[self::td or self::span or self::div or self::p][1]",
        ]

        for xp in xpaths:
            try:
                el = modal.find_element(By.XPATH, xp)
                if el and el.is_displayed():
                    v = _read_value(el)
                    if v and v != kw and len(v) <= 300:
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
            cand = container.find_elements(
                By.XPATH,
                ".//*[( (self::input or self::textarea or self::select) and not(@type='hidden') ) or self::td or self::span or self::div or self::p]"
            )
            for el in cand:
                try:
                    if el.is_displayed():
                        v = _read_value(el)
                        if v and v != kw and len(v) <= 300:
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

    product_name = find_input_near_label(modal, ["상품명", "제품명", "품목"])
    model_name = find_input_near_label(modal, ["모델명", "모델코드", "상품코드", "모델", "제품모델", "상품모델"])
    color_raw = find_input_near_label(modal, ["색상", "컬러", "색깔", "color", "Color"])

    address_basic = find_input_near_label(
        modal,
        ["기본주소", "기본 주소", "설치주소", "설치 주소", "도로명주소", "도로명 주소"]
    )
    if not address_basic:
        address_basic = find_input_near_label(modal, ["주소"])

    address_detail = find_input_near_label(modal, ["상세주소", "상세 주소", "나머지주소"])

    if address_basic and address_detail and address_basic == address_detail:
        address_basic = ""

    manage_raw = find_input_near_label(
        modal,
        ["관리", "관리유형", "관리 유형", "관리방식", "관리 방식", "방문주기", "관리주기", "관리형태", "방문관리"]
    )
    contract_raw = find_input_near_label(
        modal,
        ["약정", "약정기간", "의무사용기간", "의무 사용기간", "계약기간", "사용기간"]
    )

    discount_raw = find_input_near_label(
        modal,
        ["할인", "할인유형", "할인 유형", "반값", "프로모션", "혜택"]
    )
    amount_raw = find_input_near_label(
        modal,
        ["월요금", "예상월요금", "예상 월요금", "월 렌탈료", "렌탈료", "월금액", "월 금액", "납부금액", "월 납부금액", "금액"]
    )

    phone11 = normalize_phone11_only_010(phone_raw)

    return {
        "name": (name or "").strip(),
        "birth": (birth or "").strip(),
        "phone_raw": (phone_raw or "").strip(),
        "phone11": (phone11 or "").strip(),
        "account": (account or "").strip(),
        "zipcode": (zipcode or "").strip(),
        "product_name": (product_name or "").strip(),
        "model_name": (model_name or "").strip(),
        "color_raw": (color_raw or "").strip(),
        "address": (address_basic or "").strip(),
        "address_basic": (address_basic or "").strip(),
        "address_detail": (address_detail or "").strip(),
        "manage_raw": (manage_raw or "").strip(),
        "contract_raw": (contract_raw or "").strip(),
        "discount_raw": (discount_raw or "").strip(),
        "amount_raw": (amount_raw or "").strip(),
    }
# ---------------------------
# uiautomator2 helpers
# ---------------------------
def connect_emulator():
    d = u2.connect(ADB_SERIAL)
    d.implicitly_wait(5.0)
    return d

def click_text_center(d, txt: str, y_min_ratio: float = 0.0, y_max_ratio: float = 1.0) -> bool:
    try:
        w, h = d.window_size()
        objs = d(text=txt).all()
        for o in objs:
            try:
                b = o.info.get("bounds", {})
                cx = (int(b.get("left", 0)) + int(b.get("right", 0))) // 2
                cy = (int(b.get("top", 0)) + int(b.get("bottom", 0))) // 2
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

def enable_fast_ime(d):
    try:
        d.set_fastinput_ime(True)
    except Exception:
        pass

def type_into_edittext(d, edit_obj, text: str) -> bool:
    """
    ✅ 한글 입력 3단계(검증 포함)
    변경:
    - 클릭 횟수/속도 완화
    - refind 횟수 줄임
    - 로딩 중일 때는 입력 중단
    """
    def _get_now(obj):
        try:
            return (obj.get_text() or "").strip()
        except Exception:
            try:
                info = obj.info or {}
                return str(info.get("text") or "").strip()
            except Exception:
                return ""

    def _is_loading_overlay():
        loading_texts = [
            "조회중입니다",
            "잠시만 기다려주세요",
            "조회중입니다.",
            "잠시만 기다려주세요.",
        ]
        for t in loading_texts:
            try:
                if d(textContains=t).exists:
                    return True
            except Exception:
                pass
        return False

    try:
        if _is_loading_overlay():
            print("⏳ [type_into_edittext] 로딩중이라 입력 보류")
            return False

        obj = edit_obj
        b = obj.info.get("bounds", {})
        left = int(b.get("left", 0))
        top = int(b.get("top", 0))
        right = int(b.get("right", 0))
        bottom = int(b.get("bottom", 0))

        cy = (top + bottom) // 2
        focus_x = left + max(20, (right - left) // 6)

        # ✅ 클릭 1회만 하고 충분히 대기
        d.click(focus_x, cy)
        time.sleep(0.35)

        if _is_loading_overlay():
            print("⏳ [type_into_edittext] 포커스 직후 로딩감지")
            return False

        # 이전 값 비우기
        try:
            obj.set_text("")
            time.sleep(0.25)
        except Exception:
            pass

        # 1) set_text
        try:
            obj.set_text(text)
            time.sleep(0.50)
            got = _get_now(obj)
            print(f"🔍 [type_into_edittext] set_text 결과: '{got}'")
            if got and (got == text or text in got):
                return True
        except Exception as e:
            print("⚠️ [type_into_edittext] set_text 실패:", e)

        if _is_loading_overlay():
            print("⏳ [type_into_edittext] set_text 후 로딩감지")
            return False

        # 2) send_keys
        try:
            d.set_fastinput_ime(True)
        except Exception:
            pass

        try:
            d.click(focus_x, cy)
            time.sleep(0.30)
            d.send_keys(text, clear=True)
            time.sleep(0.55)
            got = _get_now(obj)
            print(f"🔍 [type_into_edittext] send_keys 결과: '{got}'")
            if got and (got == text or text in got):
                return True
        except Exception as e:
            print("⚠️ [type_into_edittext] send_keys 실패:", e)

        if _is_loading_overlay():
            print("⏳ [type_into_edittext] send_keys 후 로딩감지")
            return False

        # 3) clipboard paste
        try:
            d.set_clipboard(text)
            time.sleep(0.20)

            d.long_click(focus_x, cy, 0.7)
            time.sleep(0.55)

            paste_candidates = ["붙여넣기", "붙여 넣기", "Paste", "PASTE"]
            clicked = False

            for t in paste_candidates:
                if d(text=t).exists:
                    d(text=t).click()
                    clicked = True
                    time.sleep(0.50)
                    break

            if not clicked:
                if d(textContains="붙여").exists:
                    d(textContains="붙여").click()
                    clicked = True
                    time.sleep(0.50)
                elif d(textContains="Paste").exists:
                    d(textContains="Paste").click()
                    clicked = True
                    time.sleep(0.50)

            got = _get_now(obj)
            print(f"🔍 [type_into_edittext] clipboard 결과: '{got}' / clicked={clicked}")
            if got and (got == text or text in got):
                return True
        except Exception as e:
            print("⚠️ [type_into_edittext] clipboard 실패:", e)

        print("❌ [type_into_edittext] 최종 입력 실패")
        return False

    except Exception as e:
        print("❌ [type_into_edittext] 예외:", e)
        return False

def restart_digital_sales_app(d) -> bool:
    try:
        if DIGITAL_SALES_PACKAGE:
            try:
                d.app_stop(DIGITAL_SALES_PACKAGE)
                time.sleep(0.8)
            except Exception:
                pass
            try:
                d.app_start(DIGITAL_SALES_PACKAGE)
                time.sleep(2.0)
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
    except Exception:
        return False

def ensure_mobile_order_home(d) -> bool:
    for _ in range(10):
        if d(text="일반 주문하기").exists:
            return True
        if d(text="모바일 주문").exists:
            click_text_center(d, "모바일 주문", 0.70, 0.98)
            time.sleep(1.0)
            if d(text="일반 주문하기").exists:
                return True
        time.sleep(0.4)
    return d(text="일반 주문하기").exists

def is_unexpected_digital_sales_home(d) -> bool:
    """
    ✅ 작업 중 갑자기 디지털세일즈 홈으로 튄 상태 감지
    예:
    - 상단 타이틀이 '디지털세일즈'
    - 주문현황/주문접수 화면이 아님
    - 모바일 주문 홈('일반 주문하기')도 아님
    """
    try:
        if d(text="주문현황").exists:
            return False
    except Exception:
        pass

    try:
        if d(text="주문접수").exists or d(text="본인인증 요청").exists:
            return False
    except Exception:
        pass

    try:
        if d(text="일반 주문하기").exists:
            return False
    except Exception:
        pass

    try:
        if d(text="디지털세일즈").exists:
            return True
    except Exception:
        pass

    return False

def recover_from_unexpected_home(d, context: str, pause_sec: float = 1.5) -> bool:
    """
    ✅ 작업 중 홈 화면 이탈 감지 시
    - 알림
    - 잠시 대기
    - 모바일 주문 홈으로 복귀
    """
    if not is_unexpected_digital_sales_home(d):
        return True

    msg = f"예상 밖 홈화면 이동 감지: {context} → 잠시 대기 후 재진행"
    print("⚠️", msg)
    notify(msg)

    time.sleep(pause_sec)

    try:
        if d(text="모바일 주문").exists:
            click_text_center(d, "모바일 주문", 0.70, 0.98)
            time.sleep(1.0)
    except Exception:
        pass

    ok = ensure_mobile_order_home(d)
    if ok:
        print(f"✅ [recover_from_unexpected_home] 모바일 주문 홈 복구 성공 ({context})")
    else:
        print(f"❌ [recover_from_unexpected_home] 모바일 주문 홈 복구 실패 ({context})")
    return ok

def is_tab_selected(d, tab_text: str) -> bool:
    """
    이 앱은 selected/checked 값이 잘 안 잡히는 경우가 많아서
    현재는 신뢰하지 않는다.
    """
    return False

def ensure_general_tab(d, force_click: bool = False) -> bool:
    """
    ✅ 주문현황에서는 '일반' 탭만 사용
    - 좌표 fallback 금지
    - 상단 탭 영역의 '일반' 텍스트 객체만 직접 클릭
    - 검증은 '고객/상호' 존재로 판단
    """
    if not d(text="주문현황").exists:
        return False

    def _is_general_ready():
        try:
            return d(text="고객/상호").exists
        except Exception:
            return False

    def _debug_mode_label():
        try:
            has_customer_company = d(text="고객/상호").exists
        except Exception:
            has_customer_company = False

        try:
            has_customer_only = d(text="고객").exists
        except Exception:
            has_customer_only = False

        return f"고객/상호={has_customer_company}, 고객={has_customer_only}"

    if _is_general_ready():
        print("✅ [ensure_general_tab] 이미 일반 탭 상태")
        return True

    if not force_click:
        print(f"⚠️ [ensure_general_tab] 일반 탭 아님 / {_debug_mode_label()}")
        return False

    try:
        h = d.window_size()[1]

        objs = d(text="일반")
        try:
            cnt = objs.count
        except Exception:
            cnt = 0

        if cnt <= 0:
            print(f"❌ [ensure_general_tab] 상단 '일반' 텍스트 객체를 찾지 못함 / {_debug_mode_label()}")
            return False

        for i in range(cnt):
            try:
                o = objs[i]
                b = o.info.get("bounds", {})
                left = int(b.get("left", 0))
                top = int(b.get("top", 0))
                right = int(b.get("right", 0))
                bottom = int(b.get("bottom", 0))

                # 상단 탭 줄에 있는 '일반'만 허용
                if top < int(h * 0.04) or top > int(h * 0.22):
                    continue

                cx = (left + right) // 2
                cy = (top + bottom) // 2

                d.click(cx, cy)
                time.sleep(0.45)

                if _is_general_ready():
                    print(f"✅ [ensure_general_tab] 일반 탭 텍스트 클릭 성공 ({cx}, {cy})")
                    return True
            except Exception as e:
                print(f"⚠️ [ensure_general_tab] 일반 탭 객체 클릭 실패[{i}]: {e}")
                continue

        print(f"❌ [ensure_general_tab] 일반 탭 복구 실패 / {_debug_mode_label()}")
        return False

    except Exception as e:
        print("❌ [ensure_general_tab] 예외:", e)
        return False

def enter_order_status(d) -> bool:
    """
    ✅ 모바일 주문 홈에서 '주문 이어하기 > 일반주문' 리스트로 진입
    핵심:
    - x는 '일반주문' 텍스트 bounds 기준이 아니라 화면 비율 기준으로 고정
    - y만 일반주문 줄 중심값 사용
    - 사용자가 표시한 빨간 박스(건수 + 파란 화살표) 영역을 직접 누름
    """
    if d(text="주문현황").exists:
        return True

    if not ensure_mobile_order_home(d):
        print("❌ [enter_order_status] 모바일 주문 홈 진입 실패")
        return False

    try:
        objs = d(text="일반주문")
        try:
            cnt = objs.count
        except Exception:
            cnt = 0

        print(f"🔍 [enter_order_status] 일반주문 텍스트 개수: {cnt}")

        target = None
        best_top = 10**9

        for i in range(cnt):
            try:
                o = objs[i]
                b = o.info.get("bounds", {})
                left = int(b.get("left", 0))
                top = int(b.get("top", 0))
                right = int(b.get("right", 0))
                bottom = int(b.get("bottom", 0))

                print(f"   - 일반주문[{i}] bounds=({left},{top},{right},{bottom})")

                if top < 300 or top > 1400:
                    continue

                if top < best_top:
                    best_top = top
                    target = (left, top, right, bottom)
            except Exception as e:
                print(f"⚠️ [enter_order_status] 일반주문 후보 확인 실패[{i}]: {e}")
                continue

        if target is None:
            print("❌ [enter_order_status] 주문 이어하기의 일반주문 행을 찾지 못함")
            return False

        left, top, right, bottom = target
        cy = (top + bottom) // 2
        w, h = d.window_size()

        # ✅ 빨간 박스 영역: 화면 오른쪽 84%~88% 부근
        candidate_points = [
            (int(w * 0.84), cy),
            (int(w * 0.86), cy),
            (int(w * 0.88), cy),
            (int(w * 0.84), cy - 8),
            (int(w * 0.86), cy - 8),
            (int(w * 0.88), cy - 8),
            (int(w * 0.84), cy + 8),
            (int(w * 0.86), cy + 8),
            (int(w * 0.88), cy + 8),
        ]

        tried = set()
        filtered_points = []
        for x, y in candidate_points:
            x = max(0, min(w - 10, int(x)))
            y = max(0, min(h - 10, int(y)))
            key = (x, y)
            if key not in tried:
                filtered_points.append((x, y))
                tried.add(key)

        for idx, (x, y) in enumerate(filtered_points, start=1):
            print(f"🎯 [enter_order_status] 건수/화살표 박스 클릭 시도 {idx}/{len(filtered_points)} ({x}, {y})")
            d.click(x, y)
            time.sleep(1.2)

            if d(text="주문현황").exists:
                ensure_general_tab(d, force_click=True)
                time.sleep(0.4)
                print("✅ [enter_order_status] 건수/화살표 박스 클릭으로 주문현황 진입 성공")
                return True

        # fallback 1: 빨간 박스보다 약간 왼쪽
        fallback_x1 = int(w * 0.82)
        print(f"🔁 [enter_order_status] 건수영역 왼쪽 fallback 클릭 ({fallback_x1}, {cy})")
        d.click(fallback_x1, cy)
        time.sleep(1.2)

        if d(text="주문현황").exists:
            ensure_general_tab(d, force_click=True)
            time.sleep(0.4)
            print("✅ [enter_order_status] 건수영역 왼쪽 fallback으로 주문현황 진입 성공")
            return True

        # fallback 2: 빨간 박스보다 약간 오른쪽
        fallback_x2 = int(w * 0.90)
        print(f"🔁 [enter_order_status] 건수영역 오른쪽 fallback 클릭 ({fallback_x2}, {cy})")
        d.click(fallback_x2, cy)
        time.sleep(1.2)

        if d(text="주문현황").exists:
            ensure_general_tab(d, force_click=True)
            time.sleep(0.4)
            print("✅ [enter_order_status] 건수영역 오른쪽 fallback으로 주문현황 진입 성공")
            return True

        print("❌ [enter_order_status] 주문현황 진입 실패")
        return False

    except Exception as e:
        print("❌ [enter_order_status] 예외:", e)
        return False

def exit_order_status_to_mobile_home(d) -> bool:
    """
    ✅ 주문현황 종료는 X 좌표보다 back 우선
    성공 조건:
    - 주문현황 텍스트가 사라짐
    - 일반 주문하기가 보임
    """
    def _is_mobile_home():
        try:
            return (not d(text="주문현황").exists) and d(text="일반 주문하기").exists
        except Exception:
            return False

    def _click_confirm_like():
        candidates = ["확인", "예", "닫기", "나가기", "OK", "확 인"]
        for t in candidates:
            try:
                if d(text=t).exists:
                    d(text=t).click()
                    time.sleep(0.8)
                    return True
            except Exception:
                pass

        for t in candidates:
            try:
                if d(textContains=t).exists:
                    d(textContains=t).click()
                    time.sleep(0.8)
                    return True
            except Exception:
                pass

        return False

    if _is_mobile_home():
        print("✅ [exit_order_status_to_mobile_home] 이미 모바일 주문 홈 상태")
        return True

    if not d(text="주문현황").exists:
        print("⚠️ [exit_order_status_to_mobile_home] 주문현황 화면이 아님")
        return _is_mobile_home()

    for attempt in range(3):
        print(f"🔄 [exit_order_status_to_mobile_home] 주문현황 종료 시도 {attempt + 1}/3")

        # 1) back 우선
        try:
            d.press("back")
            time.sleep(0.45)
        except Exception as e:
            print("⚠️ [exit_order_status_to_mobile_home] back 실패:", e)

        # 2) 확인류 팝업 처리
        clicked_confirm = False
        for _ in range(10):
            if _click_confirm_like():
                clicked_confirm = True
                print("✅ [exit_order_status_to_mobile_home] 확인류 팝업 처리 완료")
                break
            time.sleep(0.2)

        if not clicked_confirm:
            print("ℹ️ [exit_order_status_to_mobile_home] 확인류 팝업 없음 또는 미검출")

        # 3) 홈 복귀 판정
        for _ in range(20):
            if _is_mobile_home():
                print("✅ [exit_order_status_to_mobile_home] 모바일 주문 홈 복귀 성공")
                return True
            time.sleep(0.25)

        print("⚠️ [exit_order_status_to_mobile_home] 아직 모바일 주문 홈 복귀 안 됨")

    print("❌ [exit_order_status_to_mobile_home] 모바일 주문 홈 복귀 실패")
    return False

def refresh_order_status_by_reenter(d) -> bool:
    """
    ✅ 앱 종료 없이 주문현황 리스트를 강하게 갱신
    흐름:
    주문현황 X -> 확인 -> 모바일 주문 홈 -> 일반주문 재진입

    중요:
    - 실제로 '일반 주문하기' 화면까지 나갔는지 검증
    - 다시 주문현황으로 들어왔는지 검증
    """
    try:
        print("🔄 [refresh_order_status_by_reenter] 주문현황 재진입 갱신 시작")

        if not d(text="주문현황").exists:
            print("ℹ️ [refresh_order_status_by_reenter] 현재 주문현황 밖 → 바로 진입 시도")
            ok = enter_order_status(d)
            if not ok:
                print("❌ [refresh_order_status_by_reenter] 주문현황 진입 실패")
                return False

            if not d(text="주문현황").exists:
                print("❌ [refresh_order_status_by_reenter] enter_order_status 이후에도 주문현황 미확인")
                return False

            ensure_general_tab(d, force_click=True)
            print("✅ [refresh_order_status_by_reenter] 주문현황 진입으로 갱신 완료")
            return True

        ok = exit_order_status_to_mobile_home(d)
        if not ok:
            print("❌ [refresh_order_status_by_reenter] 주문현황 종료 실패")
            return False

        if d(text="주문현황").exists:
            print("❌ [refresh_order_status_by_reenter] 종료 후에도 아직 주문현황 화면임")
            return False

        if not d(text="일반 주문하기").exists:
            print("❌ [refresh_order_status_by_reenter] 종료 후 일반 주문하기 화면 미확인")
            return False

        print("✅ [refresh_order_status_by_reenter] 모바일 주문 홈 확인 완료")

        ok = enter_order_status(d)
        if not ok:
            print("❌ [refresh_order_status_by_reenter] 재진입 실패")
            return False

        if not d(text="주문현황").exists:
            print("❌ [refresh_order_status_by_reenter] 재진입 후에도 주문현황 미확인")
            return False

        if not ensure_general_tab(d, force_click=True):
            print("❌ [refresh_order_status_by_reenter] 재진입 후 일반 탭 복구 실패")
            return False

        print("✅ [refresh_order_status_by_reenter] 주문현황 재진입 갱신 완료")
        return True

    except Exception as e:
        print("❌ [refresh_order_status_by_reenter] 예외:", e)
        return False
# ---------------------------
# Search (드롭다운 클릭 금지)
# ---------------------------
def dismiss_search_mode_dropdown_if_open(d):
    try:
        if d(text="연락처").exists and d(text="고객/상호").exists:
            w, h = d.window_size()
            d.click(int(w * 0.60), int(h * 0.12))
            time.sleep(0.3)
    except Exception:
        pass

def find_search_edittext(d):
    """
    ✅ 가장 넓은 검색어 입력(EditText) 선택
    - 현재 uiautomator2 환경에서는 .all() 대신 count + index 방식 사용
    - 상단 검색영역 후보 중 가장 넓은 입력칸을 사용
    """
    try:
        w, h = d.window_size()
        best = None
        best_w = 0

        edits = d(className="android.widget.EditText")
        cnt = edits.count
        print(f"🔍 [find_search_edittext] EditText 개수: {cnt}")

        for i in range(cnt):
            try:
                e = edits[i]
                b = e.info.get("bounds", {})
                left = int(b.get("left", 0))
                top = int(b.get("top", 0))
                right = int(b.get("right", 0))
                bottom = int(b.get("bottom", 0))
                bw = right - left

                print(f"   - EditText[{i}] bounds=({left},{top},{right},{bottom}) width={bw}")

                if top < int(h * 0.10) or top > int(h * 0.42):
                    continue

                if bw < int(w * 0.30):
                    continue

                if bw > best_w:
                    best_w = bw
                    best = e
            except Exception as e:
                print(f"⚠️ [find_search_edittext] EditText[{i}] 확인 실패: {e}")
                continue

        if best is None:
            print("❌ [find_search_edittext] 상단 검색영역 후보를 찾지 못함")
        else:
            try:
                b = best.info.get("bounds", {})
                print(f"✅ [find_search_edittext] 선택된 bounds={b}")
            except Exception:
                pass

        return best
    except Exception as e:
        print("❌ [find_search_edittext] 예외:", e)
        return None

def trigger_search(d, edit_obj):
    """
    ✅ 검색 실행을 확실히:
    - 입력칸 '바깥쪽 오른쪽' (돋보기 아이콘 위치)을 클릭
    - 그 다음 Enter도 1번 누름
    """
    try:
        w, h = d.window_size()
        b = edit_obj.info.get("bounds", {})
        right = int(b.get("right", 0))
        top = int(b.get("top", 0))
        bottom = int(b.get("bottom", 0))
        cy = (top + bottom) // 2

        # ✅ 입력칸 내부가 아니라 '오른쪽 바깥'을 찍는다
        x = min(w - 5, right + 25)
        d.click(x, cy)
        time.sleep(0.35)
    except Exception:
        pass

    try:
        d.press("enter")
        time.sleep(0.35)
    except Exception:
        pass

    return True

def get_status_badges_in_results(d, target_status: str = ""):
    """
    ✅ 검색 결과 리스트에서 화면에 보이는 상태 배지들을 위에서부터 수집
    target_status="인증완료" 처럼 주면 해당 상태만 수집
    """
    status_texts = [target_status] if target_status else ["인증완료", "인증입력", "서명입력", "주문확정", "주문삭제", "주문불가"]

    try:
        w, h = d.window_size()
        found = []

        for st in status_texts:
            objs = d(text=st)
            try:
                cnt = objs.count
            except Exception:
                cnt = 0

            for i in range(cnt):
                try:
                    o = objs[i]
                    info = o.info or {}
                    b = info.get("bounds", {})

                    left = int(b.get("left", 0))
                    top = int(b.get("top", 0))
                    right = int(b.get("right", 0))
                    bottom = int(b.get("bottom", 0))

                    if top < int(h * 0.20):
                        continue
                    if bottom > int(h * 0.95):
                        continue

                    found.append({
                        "status": st,
                        "bounds": (left, top, right, bottom),
                    })
                except Exception:
                    continue

        found.sort(key=lambda x: (x["bounds"][1], x["bounds"][0]))

        if target_status:
            print(f"✅ [get_status_badges_in_results] {target_status} 후보 {len(found)}개")
            for idx, item in enumerate(found, start=1):
                print(f"   - 후보 {idx}: status={item['status']} bounds={item['bounds']}")

        return found

    except Exception as e:
        print("❌ [get_status_badges_in_results] 예외:", e)
        return []

def get_first_status_badge_in_results(d):
    """
    ✅ 검색 결과 리스트에서 가장 위에 보이는 상태 배지 1개를 반환
    """
    try:
        found = get_status_badges_in_results(d)
        if not found:
            return ("", None)

        first = found[0]
        print(f"✅ [get_first_status_badge_in_results] 첫 상태 배지 탐지: {first['status']} / bounds={first['bounds']}")
        return (first["status"], first)

    except Exception as e:
        print("❌ [get_first_status_badge_in_results] 예외:", e)
        return ("", None)

def open_detail_by_status_badge(d, badge_obj) -> bool:
    try:
        if isinstance(badge_obj, dict):
            left, top, right, bottom = badge_obj.get("bounds", (0, 0, 0, 0))
        elif isinstance(badge_obj, tuple) and len(badge_obj) == 4:
            left, top, right, bottom = badge_obj
        else:
            info = badge_obj.info or {}
            b = info.get("bounds", {})
            left = int(b.get("left", 0))
            top = int(b.get("top", 0))
            right = int(b.get("right", 0))
            bottom = int(b.get("bottom", 0))

        cx = (int(left) + int(right)) // 2
        cy = (int(top) + int(bottom)) // 2
        d.click(cx, cy)
        time.sleep(0.8)
        return True
    except Exception as e:
        print("❌ [open_detail_by_status_badge] 예외:", e)
        return False

def match_detail_by_name_phone(d, name: str, phone11: str) -> bool:
    if not name or len(name.strip()) < 2:
        return False
    disp = phone11_to_display(phone11)
    if not disp:
        return False
    return d(text=name).exists and d(text=disp).exists

def back_to_order_status(d) -> bool:
    for _ in range(4):
        if d(text="주문현황").exists:
            return True
        try:
            d.press("back")
        except Exception:
            pass
        time.sleep(0.4)
    return d(text="주문현황").exists

def wait_until_order_status_ready(d, timeout_sec: float = 8.0) -> str:
    """
    리턴값:
    - "ready": 주문현황 검색 가능 상태
    - "home": 로딩 중/대기 중 홈화면 이탈 감지
    - "timeout": 주문현황에 머물렀지만 로딩이 너무 오래 감
    """
    def _is_loading_overlay():
        loading_texts = [
            "조회중입니다",
            "잠시만 기다려주세요",
            "조회중입니다.",
            "잠시만 기다려주세요.",
        ]

        for t in loading_texts:
            try:
                if d(textContains=t).exists:
                    return True
            except Exception:
                pass

        try:
            pbs = d(className="android.widget.ProgressBar")
            if pbs.count > 0:
                for i in range(pbs.count):
                    try:
                        info = pbs[i].info or {}
                        b = info.get("bounds", {})
                        top = int(b.get("top", 0))
                        bottom = int(b.get("bottom", 0))
                        cy = (top + bottom) // 2
                        h = d.window_size()[1]
                        if int(h * 0.25) <= cy <= int(h * 0.80):
                            return True
                    except Exception:
                        continue
        except Exception:
            pass

        return False

    end_at = time.time() + timeout_sec
    stable_ok_count = 0

    while time.time() < end_at:
        try:
            if is_unexpected_digital_sales_home(d):
                print("⚠️ [wait_until_order_status_ready] 로딩 중 홈화면 이탈 감지")
                return "home"

            if not d(text="주문현황").exists:
                stable_ok_count = 0
                time.sleep(0.25)
                continue

            if _is_loading_overlay():
                print("⏳ [wait_until_order_status_ready] 주문현황 로딩중...")
                stable_ok_count = 0
                time.sleep(0.40)
                continue

            edits = d(className="android.widget.EditText")
            cnt = edits.count

            if cnt >= 2:
                edit = find_search_edittext(d)
                if edit is not None:
                    stable_ok_count += 1
                    if stable_ok_count >= 3:
                        print("✅ [wait_until_order_status_ready] 주문현황 로딩 완료")
                        time.sleep(0.45)
                        return "ready"
                    time.sleep(0.25)
                    continue

            stable_ok_count = 0
        except Exception:
            stable_ok_count = 0

        time.sleep(0.25)

    if is_unexpected_digital_sales_home(d):
        print("⚠️ [wait_until_order_status_ready] timeout 직전 홈화면 이탈 감지")
        return "home"

    print("⚠️ [wait_until_order_status_ready] 주문현황 로딩 대기 timeout")
    return "timeout"

def check_one_job_status_by_search(d, job: dict) -> str:
    """
    ✅ 리스트 검색 중
    - 홈화면 이탈이면 복구 후 재진입
    - 로딩 timeout이면 재진입 갱신 후 같은 고객 다시 검색
    """
    for attempt in range(3):
        if is_unexpected_digital_sales_home(d):
            if not recover_from_unexpected_home(d, f"상태확인 시작 직전 {job['name']}"):
                return ""

        if not enter_order_status(d):
            print("❌ [status_check] 주문현황 진입 실패")
            return ""

        if is_unexpected_digital_sales_home(d):
            if attempt < 2 and recover_from_unexpected_home(d, f"주문현황 진입 직후 {job['name']}"):
                continue
            return ""

        ready_state = wait_until_order_status_ready(d, timeout_sec=8.0)

        if ready_state == "home":
            print(f"⚠️ [status_check] 로딩 중 홈화면 이탈 → 복구 후 재시도: {job['name']}")
            notify(f"로딩 중 홈화면 이탈 감지 → 재시도: {job['name']}")
            if attempt < 2 and recover_from_unexpected_home(d, f"주문현황 로딩중 홈이탈 {job['name']}"):
                time.sleep(0.8)
                continue
            return ""

        if ready_state == "timeout":
            print(f"⚠️ [status_check] 주문현황 로딩 timeout → 재진입 후 재시도: {job['name']}")
            notify(f"주문현황 로딩 timeout → 재진입 재시도: {job['name']}")
            if attempt < 2:
                ok_refresh = refresh_order_status_by_reenter(d)
                if ok_refresh:
                    time.sleep(1.0)
                    continue
            return ""

        if ready_state != "ready":
            print("❌ [status_check] 주문현황 준비상태 판정 실패")
            return ""

        time.sleep(0.50)

        if not ensure_general_tab(d, force_click=True):
            print("❌ [status_check] 일반 탭 복구 실패 → 이번 검색 중단")
            return ""

        try:
            if not d(text="고객/상호").exists:
                print("❌ [status_check] 일반 탭 검증 실패(고객/상호 미표시) → 이번 검색 중단")
                return ""
        except Exception:
            print("❌ [status_check] 일반 탭 검증 예외 → 이번 검색 중단")
            return ""

        time.sleep(0.30)
        dismiss_search_mode_dropdown_if_open(d)
        time.sleep(0.25)

        edit = find_search_edittext(d)
        if edit is None:
            print("❌ [status_check] 검색어 입력 EditText를 찾지 못함")
            return ""

        query = job["name"]
        print(f"🔎 [status_check] 검색 시작: {query}")

        ok = type_into_edittext(d, edit, query)
        if not ok:
            print(f"❌ [status_check] 검색어 입력 실패: {query}")
            return ""

        if is_unexpected_digital_sales_home(d):
            if attempt < 2 and recover_from_unexpected_home(d, f"검색어 입력 중 {job['name']}"):
                time.sleep(0.8)
                continue
            return ""

        time.sleep(0.45)
        trigger_search(d, edit)
        time.sleep(1.30)

        if is_unexpected_digital_sales_home(d):
            print(f"⚠️ [status_check] 검색 실행 후 홈화면 이탈 → 재시도: {job['name']}")
            notify(f"검색 실행 후 홈화면 이탈 → 재시도: {job['name']}")
            if attempt < 2 and recover_from_unexpected_home(d, f"검색 실행 중 {job['name']}"):
                time.sleep(0.8)
                continue
            return ""

        time.sleep(0.80)

        ready_badges = get_status_badges_in_results(d, target_status="인증완료")
        if ready_badges:
            print(f"✅ [status_check] 인증완료 후보 {len(ready_badges)}개 감지")
            print("🧾 [status_check] 검색 결과 상태: 인증완료")
            return "인증완료"

        st, badge = get_first_status_badge_in_results(d)
        print(f"🧾 [status_check] 검색 결과 상태: {st or 'NONE'}")
        return st

    return ""

def try_open_ready_sign_detail(d, job: dict) -> bool:
    """
    ✅ 인증완료 상세 진입
    - 검색 결과에서 '인증완료' 핑크 배지 클릭
    - 상세 고객명 + 연락처가 웹 저장값과 완전 일치해야만 통과
    - 일치 시 '주문 이어서 하기' 클릭
    - 상품검색 화면에서 모델명 검색
    - 색상 일치 후보가 정확히 1개일 때만 선택
    - 판매구분에서 '렌탈' 선택
    - 관리유형 = 모달의 관리와 일치해야 함
    - 의무사용기간 = 모달의 약정과 일치해야 함
    - 상품 담기 → 할인정보 입력
    - 할인 페이지에서 반값/총금액 검증
    - 결제정보 선택 페이지에서 정기결제 수단 선택 클릭 후 '추가' 버튼까지 진행
    """
    global SIGN_IN_PROGRESS

    def _abort(reason: str) -> bool:
        print("🛑 [ready_sign]", reason)
        set_stop(reason)
        return False

    def _normalize_color_keyword(raw: str) -> str:
        s = str(raw or "").strip().lower()

        pairs = [
            ("화이트", "화이트"),
            ("white", "화이트"),
            ("베이지", "베이지"),
            ("beige", "베이지"),
            ("그레이", "그레이"),
            ("gray", "그레이"),
            ("grey", "그레이"),
            ("블루", "블루"),
            ("blue", "블루"),
            ("핑크", "핑크"),
            ("pink", "핑크"),
            ("블랙", "블랙"),
            ("black", "블랙"),
            ("실버", "실버"),
            ("silver", "실버"),
        ]

        for key, val in pairs:
            if key in s:
                return val
        return ""

    def _normalize_manage_target(raw: str) -> str:
        s = re.sub(r"\s+", "", str(raw or "")).lower()
        if not s:
            return ""

        if "자가" in s or "셀프" in s or "self" in s:
            return "자가"

        m = re.search(r"(\d+)\s*개?월", s)
        if m:
            return f"{m.group(1)}M"

        m = re.search(r"(\d+)\s*m", s)
        if m:
            return f"{m.group(1)}M"

        if "2m" in s:
            return "2M"
        if "4m" in s:
            return "4M"

        return str(raw or "").strip()

    def _normalize_contract_target(raw: str) -> str:
        s = re.sub(r"\s+", "", str(raw or ""))
        if not s:
            return ""

        m = re.search(r"(\d+)년", s)
        if m:
            return f"{m.group(1)}년"

        m = re.search(r"(\d+)개월", s)
        if m:
            months = int(m.group(1))
            if months % 12 == 0:
                return f"{months // 12}년"
            return f"{months}개월"

        return str(raw or "").strip()

    def _normalize_discount_target(raw: str) -> str:
        s = re.sub(r"\s+", "", str(raw or ""))
        if not s:
            return ""

        if "반값" in s:
            m = re.search(r"(\d+)\s*개?월", str(raw or ""))
            if m:
                return f"{m.group(1)}개월반값"
            m = re.search(r"(\d+)", s)
            if m:
                return f"{m.group(1)}개월반값"
            return "반값"

        return s

    def _normalize_amount_digits(raw: str) -> str:
        return normalize_digits(raw or "")

    def _scan_clickable_texts(candidates, y_min_ratio: float = 0.20, y_max_ratio: float = 0.98):
        try:
            w, h = d.window_size()
            found = []
            seen = set()

            for cls in ["android.widget.TextView", "android.widget.Button"]:
                objs = d(className=cls)
                try:
                    cnt = objs.count
                except Exception:
                    cnt = 0

                for i in range(cnt):
                    try:
                        obj = objs[i]
                        info = obj.info or {}
                        txt = str(info.get("text") or "").strip()
                        if not txt:
                            continue

                        if not any(c and c in txt for c in candidates):
                            continue

                        b = info.get("bounds", {})
                        left = int(b.get("left", 0))
                        top = int(b.get("top", 0))
                        right = int(b.get("right", 0))
                        bottom = int(b.get("bottom", 0))

                        if top < int(h * y_min_ratio):
                            continue
                        if bottom > int(h * y_max_ratio):
                            continue

                        key = (txt, left, top, right, bottom)
                        if key in seen:
                            continue
                        seen.add(key)

                        found.append({
                            "text": txt,
                            "bounds": (left, top, right, bottom),
                            "class_name": cls,
                        })
                    except Exception:
                        continue

            found.sort(key=lambda x: (x["bounds"][1], x["bounds"][0]))
            return found
        except Exception as e:
            print("❌ [ready_sign] 텍스트 스캔 예외:", e)
            return []

    def _click_first_text_candidate(candidates, y_min_ratio: float = 0.20, y_max_ratio: float = 0.98):
        found = _scan_clickable_texts(candidates, y_min_ratio=y_min_ratio, y_max_ratio=y_max_ratio)
        if not found:
            return None

        item = found[0]
        left, top, right, bottom = item["bounds"]
        cx = (left + right) // 2
        cy = (top + bottom) // 2
        d.click(cx, cy)
        time.sleep(2.0)
        return item

    def _find_product_search_edittext():
        try:
            w, h = d.window_size()
            edits = d(className="android.widget.EditText")
            cnt = edits.count

            best = None
            best_w = 0

            for i in range(cnt):
                try:
                    e = edits[i]
                    b = e.info.get("bounds", {})
                    left = int(b.get("left", 0))
                    top = int(b.get("top", 0))
                    right = int(b.get("right", 0))
                    bottom = int(b.get("bottom", 0))
                    bw = right - left

                    if top < int(h * 0.10) or top > int(h * 0.35):
                        continue

                    if bw > best_w:
                        best_w = bw
                        best = e
                except Exception:
                    continue

            return best
        except Exception:
            return None

    def _wait_product_search_ready(timeout_sec: float = 8.0) -> bool:
        end_at = time.time() + timeout_sec

        while time.time() < end_at:
            try:
                if is_unexpected_digital_sales_home(d):
                    return False

                has_title = d(text="상품검색").exists or d(text="주문접수").exists
                has_search_btn = d(text="검색").exists
                edit = _find_product_search_edittext()

                if has_title and has_search_btn and edit is not None:
                    print("✅ [ready_sign] 상품검색 화면 준비 완료")
                    time.sleep(0.4)
                    return True
            except Exception:
                pass

            time.sleep(0.25)

        print("❌ [ready_sign] 상품검색 화면 준비 실패")
        return False

    def _search_product_by_model(model_query: str) -> bool:
        if not _wait_product_search_ready(timeout_sec=8.0):
            return False

        edit = _find_product_search_edittext()
        if edit is None:
            print("❌ [ready_sign] 상품검색 입력창을 찾지 못함")
            return False

        print(f"🔎 [ready_sign] 모델명 검색: {model_query}")
        ok = type_into_edittext(d, edit, model_query)
        if not ok:
            print(f"❌ [ready_sign] 모델명 입력 실패: {model_query}")
            return False

        time.sleep(0.35)

        if d(text="검색").exists:
            click_text_center(d, "검색", 0.10, 0.35)
        else:
            try:
                d.press("enter")
            except Exception:
                return False

        print("⏳ [ready_sign] 제품검색 결과 안정화 대기 3.5초")
        time.sleep(3.5)
        return True

    def _collect_product_candidates(model_query: str):
        try:
            w, h = d.window_size()

            query_norm = re.sub(r"\s+", "", str(model_query or "")).upper()
            model_keys = [query_norm]
            if "_" in query_norm:
                model_keys.append(query_norm.split("_")[0])

            tvs = d(className="android.widget.TextView")
            cnt = tvs.count

            candidates = []
            seen = set()

            for i in range(cnt):
                try:
                    tv = tvs[i]
                    info = tv.info or {}
                    txt = str(info.get("text") or "").strip()
                    if not txt:
                        continue

                    if "," not in txt:
                        continue

                    txt_norm = re.sub(r"\s+", "", txt).upper()
                    if not any(k and k in txt_norm for k in model_keys):
                        continue

                    b = info.get("bounds", {})
                    left = int(b.get("left", 0))
                    top = int(b.get("top", 0))
                    right = int(b.get("right", 0))
                    bottom = int(b.get("bottom", 0))

                    if top < int(h * 0.18) or bottom > int(h * 0.92):
                        continue

                    if txt in seen:
                        continue
                    seen.add(txt)

                    parts = [p.strip() for p in txt.split(",")]
                    color_text = parts[-1] if len(parts) >= 3 else txt

                    candidates.append({
                        "text": txt,
                        "color_text": color_text,
                        "obj": tv,
                        "bounds": (left, top, right, bottom),
                    })
                except Exception:
                    continue

            print(f"✅ [ready_sign] 상품후보 수집: {len(candidates)}개")
            for c in candidates:
                print("   - 후보:", c["text"])

            return candidates

        except Exception as e:
            print("❌ [ready_sign] 상품후보 수집 예외:", e)
            return []

    def _is_search_return_popup_open() -> bool:
        popup_fragments = [
            "상품 검색 화면으로 이동",
            "상품검색 화면으로 이동",
            "이동하시겠습니까",
        ]

        for frag in popup_fragments:
            try:
                if d(text=frag).exists:
                    return True
            except Exception:
                pass

            try:
                if d(textContains=frag).exists:
                    return True
            except Exception:
                pass

        return False

    def _dismiss_search_return_popup_if_open() -> bool:
        if not _is_search_return_popup_open():
            return False

        print("⚠️ [ready_sign] 상품검색 복귀 팝업 감지 → [취소] 클릭 후 계속 진행")

        clicked = False

        try:
            if d(text="취소").exists:
                d(text="취소").click()
                clicked = True
        except Exception:
            clicked = False

        if not clicked:
            try:
                if d(textContains="취소").exists:
                    d(textContains="취소").click()
                    clicked = True
            except Exception:
                clicked = False

        if not clicked:
            try:
                clicked = click_text_center(d, "취소", 0.45, 0.98)
            except Exception:
                clicked = False

        time.sleep(2.0)

        if _is_search_return_popup_open():
            print("⚠️ [ready_sign] 팝업이 아직 남아있음")
        else:
            print("✅ [ready_sign] 팝업 취소 처리 완료")

        return clicked

    def _choose_and_click_product(model_query: str, color_raw: str) -> bool:
        desired_color = _normalize_color_keyword(color_raw)
        if not desired_color:
            return _abort(f"전자서명 중단: 색상 추출 실패 / 고객={job['name']} / 원본색상={color_raw}")

        candidates = _collect_product_candidates(model_query)
        if not candidates:
            return _abort(f"전자서명 중단: 상품검색 결과 없음 / 고객={job['name']} / 모델={model_query}")

        matched = []
        for c in candidates:
            full_text = f"{c['text']} {c['color_text']}"
            if desired_color in full_text:
                matched.append(c)

        if len(matched) == 0:
            return _abort(
                f"전자서명 중단: 색상 일치 후보 없음 / 고객={job['name']} / 모델={model_query} / 색상={color_raw}"
            )

        if len(matched) > 1:
            joined = " | ".join([m["text"] for m in matched])
            return _abort(
                f"전자서명 중단: 색상 후보 {len(matched)}개로 모호함 / 고객={job['name']} / 색상={color_raw} / 후보={joined}"
            )

        chosen = matched[0]
        left, top, right, bottom = chosen["bounds"]
        cx = (left + right) // 2
        cy = (top + bottom) // 2

        print(f"✅ [ready_sign] 최종 상품 선택: {chosen['text']}")
        d.click(cx, cy)

        print("⏳ [ready_sign] 색상 선택 후 주문접수/상품선택 전환 대기 10.0초")
        time.sleep(10.0)

        if _is_search_return_popup_open():
            _dismiss_search_return_popup_if_open()

        if not _wait_product_option_ready(timeout_sec=12.0):
            return _abort(f"전자서명 중단: 상품선택 화면 준비 실패 / 고객={job['name']} / 모델={model_query}")

        return True

    def _wait_product_option_ready(timeout_sec: float = 12.0) -> bool:
        end_at = time.time() + timeout_sec
        stable_ok_count = 0

        while time.time() < end_at:
            try:
                if is_unexpected_digital_sales_home(d):
                    return False

                if _is_search_return_popup_open():
                    _dismiss_search_return_popup_if_open()
                    stable_ok_count = 0
                    time.sleep(0.5)
                    continue

                has_title = d(text="상품선택").exists or d(text="주문접수").exists
                has_sale_group = d(text="판매구분").exists

                if has_title and has_sale_group:
                    stable_ok_count += 1
                    if stable_ok_count >= 2:
                        print("✅ [ready_sign] 판매구분 화면 준비 완료")
                        time.sleep(0.5)
                        return True
                else:
                    stable_ok_count = 0

            except Exception:
                stable_ok_count = 0

            time.sleep(0.5)

        print("❌ [ready_sign] 판매구분 화면 준비 실패")
        return False

    def _select_rental_option() -> bool:
        if not _wait_product_option_ready(timeout_sec=12.0):
            return _abort(f"전자서명 중단: 판매구분 화면 진입 실패 / 고객={job['name']}")

        section_open_attempted = False

        for attempt in range(6):
            if _is_search_return_popup_open():
                _dismiss_search_return_popup_if_open()

            has_rental = False
            has_cash = False

            try:
                has_rental = d(text="렌탈").exists
            except Exception:
                has_rental = False

            try:
                has_cash = d(text="일시불").exists
            except Exception:
                has_cash = False

            if has_rental and has_cash:
                ok_click = click_text_center(d, "렌탈", 0.35, 0.90)
                if not ok_click:
                    return _abort(f"전자서명 중단: 렌탈 옵션 클릭 실패 / 고객={job['name']}")

                time.sleep(2.0)
                print(f"✅ [ready_sign] 렌탈 선택 완료: {job['name']}")
                return True

            if (not section_open_attempted) and d(text="판매구분").exists:
                print(f"ℹ️ [ready_sign] 판매구분 옵션 재확인/펼치기 시도 {attempt + 1}/6")
                click_text_center(d, "판매구분", 0.35, 0.90)
                time.sleep(2.0)

                if _is_search_return_popup_open():
                    _dismiss_search_return_popup_if_open()

                section_open_attempted = True
                continue

            print(f"⚠️ [ready_sign] 판매구분 옵션 재조회 {attempt + 1}/6")
            time.sleep(2.0)

        return _abort(f"전자서명 중단: 판매구분 렌탈/일시불 옵션 미검출 / 고객={job['name']}")

    def _select_manage_type(manage_raw: str) -> bool:
        target = _normalize_manage_target(manage_raw)
        if not target:
            return _abort(f"전자서명 중단: 모달 관리 추출 실패 / 고객={job['name']} / 원본관리={manage_raw}")

        if target == "자가":
            candidates = ["자가관리", "자가", "셀프"]
        else:
            candidates = [f"방문관리-{target}", target, target.replace("M", "개월"), target.replace("M", "개월관리")]

        print(f"🔎 [ready_sign] 관리유형 선택 목표: {target}")

        item = _click_first_text_candidate(candidates, y_min_ratio=0.40, y_max_ratio=0.92)
        if item is not None:
            print(f"✅ [ready_sign] 관리유형 선택 완료: {item['text']}")
            return True

        if d(text="관리유형").exists:
            click_text_center(d, "관리유형", 0.35, 0.90)
            time.sleep(2.0)

        item = _click_first_text_candidate(candidates, y_min_ratio=0.40, y_max_ratio=0.95)
        if item is not None:
            print(f"✅ [ready_sign] 관리유형 선택 완료: {item['text']}")
            return True

        try:
            w, h = d.window_size()
            d.swipe(int(w * 0.50), int(h * 0.82), int(w * 0.50), int(h * 0.60), 0.20)
            time.sleep(0.6)
        except Exception:
            pass

        item = _click_first_text_candidate(candidates, y_min_ratio=0.35, y_max_ratio=0.98)
        if item is not None:
            print(f"✅ [ready_sign] 관리유형 선택 완료: {item['text']}")
            return True

        return _abort(f"전자서명 중단: 관리유형 일치 옵션 없음 / 고객={job['name']} / 모달관리={manage_raw} / 목표={target}")

    def _select_contract_period(contract_raw: str) -> bool:
        target = _normalize_contract_target(contract_raw)
        if not target:
            return _abort(f"전자서명 중단: 모달 약정 추출 실패 / 고객={job['name']} / 원본약정={contract_raw}")

        print(f"🔎 [ready_sign] 의무사용기간 선택 목표: {target}")

        item = _click_first_text_candidate([target], y_min_ratio=0.45, y_max_ratio=0.98)
        if item is not None:
            print(f"✅ [ready_sign] 의무사용기간 선택 완료: {item['text']}")
            return True

        if d(text="의무사용기간").exists:
            click_text_center(d, "의무사용기간", 0.45, 0.98)
            time.sleep(2.0)

        item = _click_first_text_candidate([target], y_min_ratio=0.45, y_max_ratio=0.98)
        if item is not None:
            print(f"✅ [ready_sign] 의무사용기간 선택 완료: {item['text']}")
            return True

        try:
            w, h = d.window_size()
            d.swipe(int(w * 0.50), int(h * 0.82), int(w * 0.50), int(h * 0.58), 0.20)
            time.sleep(0.6)
        except Exception:
            pass

        item = _click_first_text_candidate([target], y_min_ratio=0.40, y_max_ratio=0.98)
        if item is not None:
            print(f"✅ [ready_sign] 의무사용기간 선택 완료: {item['text']}")
            return True

        return _abort(f"전자서명 중단: 의무사용기간 일치 옵션 없음 / 고객={job['name']} / 모달약정={contract_raw} / 목표={target}")

    def _collect_visible_text_items(y_min_ratio: float = 0.0, y_max_ratio: float = 1.0):
        try:
            w, h = d.window_size()
            found = []
            seen = set()

            for cls in ["android.widget.TextView", "android.widget.Button"]:
                objs = d(className=cls)
                try:
                    cnt = objs.count
                except Exception:
                    cnt = 0

                for i in range(cnt):
                    try:
                        obj = objs[i]
                        info = obj.info or {}
                        txt = str(info.get("text") or "").strip()
                        if not txt:
                            continue

                        b = info.get("bounds", {})
                        left = int(b.get("left", 0))
                        top = int(b.get("top", 0))
                        right = int(b.get("right", 0))
                        bottom = int(b.get("bottom", 0))
                        cy = (top + bottom) // 2

                        if cy < int(h * y_min_ratio) or cy > int(h * y_max_ratio):
                            continue

                        key = (txt, left, top, right, bottom, cls)
                        if key in seen:
                            continue
                        seen.add(key)

                        found.append({
                            "text": txt,
                            "bounds": (left, top, right, bottom),
                            "class_name": cls,
                        })
                    except Exception:
                        continue

            found.sort(key=lambda x: (x["bounds"][1], x["bounds"][0]))
            return found
        except Exception as e:
            print("❌ [ready_sign] 화면 텍스트 수집 예외:", e)
            return []

    def _wait_discount_page_ready(timeout_sec: float = 12.0) -> bool:
        end_at = time.time() + timeout_sec
        stable_ok_count = 0

        while time.time() < end_at:
            try:
                if is_unexpected_digital_sales_home(d):
                    return False

                has_title = d(text="할인 선택").exists or d(textContains="할인 선택").exists
                has_next = d(text="다음").exists

                if has_title and has_next:
                    stable_ok_count += 1
                    if stable_ok_count >= 2:
                        print("✅ [ready_sign] 할인정보 입력 화면 준비 완료")
                        time.sleep(0.5)
                        return True
                else:
                    stable_ok_count = 0
            except Exception:
                stable_ok_count = 0

            time.sleep(0.4)

        print("❌ [ready_sign] 할인정보 입력 화면 준비 실패")
        return False

    def _scroll_discount_down_once() -> bool:
        try:
            w, h = d.window_size()
            d.swipe(int(w * 0.50), int(h * 0.82), int(w * 0.50), int(h * 0.34), 0.25)
            time.sleep(1.4)
            return True
        except Exception:
            return False

    def _discount_page_has_target(target_discount: str) -> bool:
        items = _collect_visible_text_items(0.10, 0.96)

        for it in items:
            norm = re.sub(r"\s+", "", it["text"])
            if not norm:
                continue

            if target_discount == "반값":
                if "반값" in norm:
                    return True
            else:
                if target_discount and target_discount in norm:
                    return True

        return False

    def _verify_total_amount_row(expected_amount_digits: str) -> bool:
        items = _collect_visible_text_items(0.20, 0.99)

        merged_texts = []
        seen = set()

        for it in items:
            norm = re.sub(r"\s+", "", str(it["text"] or ""))
            if norm and norm not in seen:
                merged_texts.append(norm)
                seen.add(norm)

        xml = ""
        try:
            xml = d.dump_hierarchy()
        except Exception:
            xml = ""

        if xml:
            try:
                xml_texts = re.findall(r'text="([^"]*)"', xml)
                for raw in xml_texts:
                    norm = re.sub(r"\s+", "", str(raw or ""))
                    if norm and norm not in seen:
                        merged_texts.append(norm)
                        seen.add(norm)
            except Exception:
                pass

            try:
                xml_descs = re.findall(r'content-desc="([^"]*)"', xml)
                for raw in xml_descs:
                    norm = re.sub(r"\s+", "", str(raw or ""))
                    if norm and norm not in seen:
                        merged_texts.append(norm)
                        seen.add(norm)
            except Exception:
                pass

        has_total_label = any("총금액" in t for t in merged_texts)

        has_receipt_zero = False
        if any("수납0원" in t for t in merged_texts):
            has_receipt_zero = True
        elif any(("수납" in t and normalize_digits(t) == "0") for t in merged_texts):
            has_receipt_zero = True
        else:
            has_receipt_word = any("수납" in t for t in merged_texts)
            has_zero_won = any(
                (t == "0원") or ("원" in t and normalize_digits(t) == "0")
                for t in merged_texts
            )
            if has_receipt_word and has_zero_won:
                has_receipt_zero = True

        amount_match = False
        for t in merged_texts:
            digits = normalize_digits(t)
            if digits and digits == expected_amount_digits and "원" in t:
                amount_match = True
                break

        if not amount_match and expected_amount_digits:
            try:
                expected_num = int(expected_amount_digits)
                expected_won = f"{expected_num:,}원"
                expected_won_month = f"{expected_num:,}원/월"
                joined = " ".join(merged_texts)

                if expected_won_month in joined or expected_won in joined:
                    amount_match = True
            except Exception:
                pass

        print(
            f"🔎 [ready_sign] 총금액 검증 / 총금액표시={has_total_label} / 수납0원={has_receipt_zero} / 금액일치={amount_match} / 기대금액={expected_amount_digits}"
        )

        return has_total_label and has_receipt_zero and amount_match

    def _click_add_product_and_go_discount() -> bool:
        if not d(text="상품 담기").exists:
            return _abort(f"전자서명 중단: 상품 담기 버튼 미검출 / 고객={job['name']}")

        click_text_center(d, "상품 담기", 0.85, 0.99)
        time.sleep(2.0)

        for _ in range(24):
            if is_unexpected_digital_sales_home(d):
                return _abort(f"전자서명 중단: 상품 담기 후 홈이탈 / 고객={job['name']}")

            has_add_more = False
            has_discount_btn = False

            try:
                has_add_more = d(text="상품 추가하기").exists
            except Exception:
                has_add_more = False

            try:
                has_discount_btn = d(text="할인정보 입력").exists
            except Exception:
                has_discount_btn = False

            if has_add_more or has_discount_btn:
                if not has_discount_btn:
                    return _abort(f"전자서명 중단: 할인정보 입력 버튼 미검출 / 고객={job['name']}")

                ok_click = click_text_center(d, "할인정보 입력", 0.35, 0.90)
                if not ok_click:
                    return _abort(f"전자서명 중단: 할인정보 입력 버튼 클릭 실패 / 고객={job['name']}")

                time.sleep(2.0)

                if not _wait_discount_page_ready(timeout_sec=12.0):
                    return _abort(f"전자서명 중단: 할인정보 입력 화면 진입 실패 / 고객={job['name']}")

                print(f"✅ [ready_sign] 상품 담기 후 할인정보 입력 화면 진입 완료: {job['name']}")
                return True

            time.sleep(0.25)

        return _abort(f"전자서명 중단: 상품 담기 완료 팝업 미검출 / 고객={job['name']}")

    def _verify_discount_and_amount_then_next(discount_raw: str, amount_raw: str) -> bool:
        target_discount = _normalize_discount_target(discount_raw)
        expected_amount_digits = _normalize_amount_digits(amount_raw)

        if not target_discount:
            return _abort(f"전자서명 중단: 모달 할인 추출 실패 / 고객={job['name']} / 원본할인={discount_raw}")

        if not expected_amount_digits:
            return _abort(f"전자서명 중단: 모달 금액 추출 실패 / 고객={job['name']} / 원본금액={amount_raw}")

        if not _wait_discount_page_ready(timeout_sec=12.0):
            return _abort(f"전자서명 중단: 할인정보 입력 화면 준비 실패 / 고객={job['name']}")

        found_discount = False
        total_ok = False

        for step in range(8):
            if is_unexpected_digital_sales_home(d):
                return _abort(f"전자서명 중단: 할인정보 입력 단계 홈이탈 / 고객={job['name']}")

            if not found_discount and _discount_page_has_target(target_discount):
                found_discount = True
                print(f"✅ [ready_sign] 할인 일치 확인: {target_discount}")

            if _verify_total_amount_row(expected_amount_digits):
                total_ok = True

            if found_discount and total_ok:
                if not d(text="다음").exists:
                    return _abort(f"전자서명 중단: 할인정보 입력 페이지 다음 버튼 미검출 / 고객={job['name']}")

                ok_click = click_text_center(d, "다음", 0.90, 0.99)
                if not ok_click:
                    return _abort(f"전자서명 중단: 할인정보 입력 페이지 다음 버튼 클릭 실패 / 고객={job['name']}")

                time.sleep(2.0)
                print(f"✅ [ready_sign] 할인/총금액 검증 완료 후 다음 클릭: {job['name']}")
                return True

            if step < 7:
                print(f"ℹ️ [ready_sign] 할인정보 하단 검증 스크롤 {step + 1}/7")
                _scroll_discount_down_once()

        if not found_discount:
            return _abort(
                f"전자서명 중단: 할인 불일치 / 고객={job['name']} / 모달할인={discount_raw} / 목표={target_discount}"
            )

        return _abort(
            f"전자서명 중단: 총 금액 또는 수납0원 불일치 / 고객={job['name']} / 모달금액={amount_raw}"
        )

    def _wait_payment_info_ready(timeout_sec: float = 12.0) -> bool:
        end_at = time.time() + timeout_sec
        stable_ok_count = 0

        while time.time() < end_at:
            try:
                if is_unexpected_digital_sales_home(d):
                    return False

                has_title = d(text="결제정보 선택").exists or d(textContains="결제정보 선택").exists
                has_next = d(text="다음").exists
                has_method = (
                    d(text="정기결제수단").exists
                    or d(text="정기결제 수단").exists
                    or d(text="정기결제 수단 선택").exists
                    or d(textContains="정기결제").exists
                )

                if has_title and has_next and has_method:
                    stable_ok_count += 1
                    if stable_ok_count >= 2:
                        print("✅ [ready_sign] 결제정보 선택 화면 준비 완료")
                        time.sleep(0.5)
                        return True
                else:
                    stable_ok_count = 0
            except Exception:
                stable_ok_count = 0

            time.sleep(0.4)

        print("❌ [ready_sign] 결제정보 선택 화면 준비 실패")
        return False

    def _click_payment_method_selector() -> bool:
        try:
            if d(text="정기결제 수단 선택").exists:
                ok = click_text_center(d, "정기결제 수단 선택", 0.20, 0.60)
                time.sleep(1.5)
                return ok
        except Exception:
            pass

        label_obj = None

        try:
            if d(text="정기결제수단").exists:
                label_obj = d(text="정기결제수단")
            elif d(text="정기결제 수단").exists:
                label_obj = d(text="정기결제 수단")
            elif d(textContains="정기결제").exists:
                label_obj = d(textContains="정기결제")
        except Exception:
            label_obj = None

        if label_obj is not None:
            try:
                w, h = d.window_size()
                b = label_obj.info.get("bounds", {})
                bottom = int(b.get("bottom", 0))
                field_y = min(h - 20, bottom + max(55, int(h * 0.03)))

                points = [
                    (int(w * 0.50), field_y),
                    (int(w * 0.82), field_y),
                    (int(w * 0.65), field_y),
                ]

                for x, y in points:
                    d.click(x, y)
                    time.sleep(1.5)
                    return True
            except Exception:
                pass

        return False

    def _wait_payment_add_sheet_and_click_add(timeout_sec: float = 6.0) -> bool:
        end_at = time.time() + timeout_sec

        while time.time() < end_at:
            if is_unexpected_digital_sales_home(d):
                return False

            has_sheet = False
            try:
                has_sheet = (
                    d(text="결제정보").exists
                    or d(textContains="등록된 결제 수단").exists
                    or d(textContains="결제 수단을 추가").exists
                )
            except Exception:
                has_sheet = False

            if has_sheet:
                try:
                    if d(text="추가").exists:
                        d(text="추가").click()
                        time.sleep(2.0)
                        print(f"✅ [ready_sign] 결제수단 추가 버튼 클릭 완료: {job['name']}")
                        return True
                except Exception:
                    pass

                try:
                    if d(textContains="추가").exists:
                        d(textContains="추가").click()
                        time.sleep(2.0)
                        print(f"✅ [ready_sign] 결제수단 추가 버튼 클릭 완료: {job['name']}")
                        return True
                except Exception:
                    pass

            time.sleep(0.25)

        return False

    def _open_payment_method_and_click_add() -> bool:
        if not _wait_payment_info_ready(timeout_sec=12.0):
            return _abort(f"전자서명 중단: 결제정보 선택 화면 진입 실패 / 고객={job['name']}")

        for attempt in range(3):
            print(f"🔎 [ready_sign] 정기결제 수단 선택 열기 시도 {attempt + 1}/3")

            ok_click = _click_payment_method_selector()
            if not ok_click:
                time.sleep(1.0)

            if _wait_payment_add_sheet_and_click_add(timeout_sec=6.0):
                print(f"✅ [ready_sign] 결제수단 추가 팝업 처리 완료: {job['name']}")
                return True

            time.sleep(1.0)

        return _abort(f"전자서명 중단: 결제수단 추가 팝업 또는 추가 버튼 미검출 / 고객={job['name']}")

    for attempt in range(2):
        if is_unexpected_digital_sales_home(d):
            if not recover_from_unexpected_home(d, f"상세진입 시작 직전 {job['name']}"):
                return False

        if not enter_order_status(d):
            print("❌ [ready_sign] 주문현황 진입 실패")
            return False

        if is_unexpected_digital_sales_home(d):
            if attempt == 0 and recover_from_unexpected_home(d, f"상세진입 주문현황 진입 직후 {job['name']}"):
                continue
            return False

        ready_state = wait_until_order_status_ready(d, timeout_sec=8.0)
        if ready_state == "home":
            if attempt == 0 and recover_from_unexpected_home(d, f"상세진입 로딩중 홈이탈 {job['name']}"):
                time.sleep(0.8)
                continue
            return False

        if ready_state == "timeout":
            if attempt == 0:
                ok_refresh = refresh_order_status_by_reenter(d)
                if ok_refresh:
                    time.sleep(1.0)
                    continue
            return False

        if not ensure_general_tab(d, force_click=True):
            print("❌ [ready_sign] 일반 탭 복구 실패")
            return False

        dismiss_search_mode_dropdown_if_open(d)

        edit = find_search_edittext(d)
        if edit is None:
            print("❌ [ready_sign] 검색어 입력 EditText를 찾지 못함")
            return False

        query = job["name"]
        print(f"🔎 [ready_sign] 인증완료 상세진입 검색: {query}")

        ok = type_into_edittext(d, edit, query)
        if not ok:
            print(f"❌ [ready_sign] 검색어 입력 실패: {query}")
            return False

        if is_unexpected_digital_sales_home(d):
            if attempt == 0 and recover_from_unexpected_home(d, f"상세진입 검색어 입력 중 {job['name']}"):
                continue
            return False

        time.sleep(0.45)
        trigger_search(d, edit)
        time.sleep(1.20)

        if is_unexpected_digital_sales_home(d):
            if attempt == 0 and recover_from_unexpected_home(d, f"상세진입 검색 실행 중 {job['name']}"):
                continue
            return False

        time.sleep(0.80)
        ready_badges = get_status_badges_in_results(d, target_status="인증완료")
        print(f"🧾 [ready_sign] 인증완료 후보 수: {len(ready_badges)}")

        if not ready_badges:
            return False

        matched_badge = None

        for idx, badge in enumerate(ready_badges, start=1):
            print(f"✅ [ready_sign] 인증완료 후보 확인 {idx}/{len(ready_badges)}: bounds={badge['bounds']}")

            ok_open = open_detail_by_status_badge(d, badge)
            if not ok_open:
                print(f"⚠️ [ready_sign] 인증완료 후보 열기 실패 {idx}/{len(ready_badges)}")
                continue

            time.sleep(0.90)

            if is_unexpected_digital_sales_home(d):
                if attempt == 0 and recover_from_unexpected_home(d, f"상세진입 상세열기 중 {job['name']}"):
                    time.sleep(0.8)
                    matched_badge = "__RETRY__"
                    break
                return False

            if match_detail_by_name_phone(d, job["name"], job["phone11"]):
                print(f"✅ [ready_sign] 상세 고객명/전화번호 매칭 성공 {idx}/{len(ready_badges)}")
                matched_badge = badge
                break

            print(f"⚠️ [ready_sign] 상세 불일치 {idx}/{len(ready_badges)} → 다음 인증완료 후보 확인")
            print(f"   - 기대 이름/번호: {job['name']} / {job['phone11']}")
            back_to_order_status(d)

            ready_state_after_back = wait_until_order_status_ready(d, timeout_sec=8.0)
            if ready_state_after_back == "home":
                if attempt == 0 and recover_from_unexpected_home(d, f"상세불일치 복귀중 홈이탈 {job['name']}"):
                    time.sleep(0.8)
                    matched_badge = "__RETRY__"
                    break
                return False

            if ready_state_after_back == "timeout":
                if attempt == 0:
                    ok_refresh = refresh_order_status_by_reenter(d)
                    if ok_refresh:
                        time.sleep(1.0)
                        matched_badge = "__RETRY__"
                        break
                return False

            if not ensure_general_tab(d, force_click=True):
                return False

            time.sleep(0.30)

        if matched_badge == "__RETRY__":
            continue

        if matched_badge is None:
            notify(f"인증완료 후보 {len(ready_badges)}개 모두 전화번호 불일치: {job['name']} / {job['phone11']}")
            print("❌ [ready_sign] 인증완료 후보 전체 확인했지만 전화번호 일치 없음")
            return False

        search_model = (job.get("model_name") or job.get("product_name") or "").strip()
        color_raw = (job.get("color_raw") or "").strip()
        manage_raw = (job.get("manage_raw") or "").strip()
        contract_raw = (job.get("contract_raw") or "").strip()
        discount_raw = (job.get("discount_raw") or "").strip()
        amount_raw = (job.get("amount_raw") or "").strip()

        if not search_model:
            return _abort(f"전자서명 중단: 모달 모델명 추출 실패 / 고객={job['name']}")

        if not color_raw:
            return _abort(f"전자서명 중단: 모달 색상 추출 실패 / 고객={job['name']} / 모델={search_model}")

        if not manage_raw:
            return _abort(f"전자서명 중단: 모달 관리 추출 실패 / 고객={job['name']} / 모델={search_model}")

        if not contract_raw:
            return _abort(f"전자서명 중단: 모달 약정 추출 실패 / 고객={job['name']} / 모델={search_model}")

        if not discount_raw:
            return _abort(f"전자서명 중단: 모달 할인 추출 실패 / 고객={job['name']} / 모델={search_model}")

        if not amount_raw:
            return _abort(f"전자서명 중단: 모달 금액 추출 실패 / 고객={job['name']} / 모델={search_model}")

        if not d(text="주문 이어서 하기").exists:
            return _abort(f"전자서명 중단: 주문 이어서 하기 버튼 미검출 / 고객={job['name']}")

        click_text_center(d, "주문 이어서 하기", 0.20, 0.85)
        time.sleep(2.0)

        if is_unexpected_digital_sales_home(d):
            return _abort(f"전자서명 중단: 주문 이어서 하기 클릭 후 홈이탈 / 고객={job['name']}")

        if not _search_product_by_model(search_model):
            return _abort(f"전자서명 중단: 모델 검색 실패 / 고객={job['name']} / 모델={search_model}")

        if is_unexpected_digital_sales_home(d):
            return _abort(f"전자서명 중단: 모델 검색 후 홈이탈 / 고객={job['name']} / 모델={search_model}")

        if not _choose_and_click_product(search_model, color_raw):
            return False

        if is_unexpected_digital_sales_home(d):
            return _abort(f"전자서명 중단: 상품 선택 후 홈이탈 / 고객={job['name']} / 모델={search_model}")

        if not _select_rental_option():
            return False

        if not _select_manage_type(manage_raw):
            return False

        if not _select_contract_period(contract_raw):
            return False

        if not _click_add_product_and_go_discount():
            return False

        if not _verify_discount_and_amount_then_next(discount_raw, amount_raw):
            return False

        if not _open_payment_method_and_click_add():
            return False

        SIGN_IN_PROGRESS = True
        sign_started.add(job["phone11"])
        notify(
            f"전자서명 할인/결제수단추가 단계 완료: {job['name']} / {job['phone11']} / 모델={search_model} / 색상={color_raw} / 관리={manage_raw} / 약정={contract_raw} / 할인={discount_raw} / 금액={amount_raw}"
        )
        print(f"✅ [ready_sign] 상품검색/색상/렌탈/관리/약정/할인검증/결제수단추가 완료: {job['name']}")
        return True

    return False
# ---------------------------
# 인증발송
# ---------------------------
def send_auth_request(d, job: dict):
    """
    ✅ 인증발송 중 홈화면 이탈이 생기면
    모바일 주문 홈으로 복구 후 '주문접수부터' 다시 시도
    """
    for attempt in range(2):
        if not restart_digital_sales_app(d):
            return (False, "APP_RESTART_FAIL")

        if not ensure_mobile_order_home(d):
            return (False, "NOT_READY")

        if is_unexpected_digital_sales_home(d):
            if not recover_from_unexpected_home(d, "인증발송 시작 직전"):
                return (False, "UNEXPECTED_HOME")

        click_text_center(d, "일반 주문하기", 0.20, 0.55)
        time.sleep(0.8)

        if is_unexpected_digital_sales_home(d):
            if attempt == 0 and recover_from_unexpected_home(d, "인증발송 진입 직후"):
                continue
            return (False, "UNEXPECTED_HOME")

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
        edits[0].set_text(job["name"])
        time.sleep(0.2)

        if is_unexpected_digital_sales_home(d):
            if attempt == 0 and recover_from_unexpected_home(d, f"인증발송 이름입력 중 {job['name']}"):
                continue
            return (False, "UNEXPECTED_HOME")

        edits = d(className="android.widget.EditText")
        if edits.count < 2:
            return (False, "NO_EDITTEXT")

        edits[1].click()
        time.sleep(0.1)
        edits[1].set_text(job["phone11"])
        time.sleep(0.2)

        if is_unexpected_digital_sales_home(d):
            if attempt == 0 and recover_from_unexpected_home(d, f"인증발송 연락처입력 중 {job['name']}"):
                continue
            return (False, "UNEXPECTED_HOME")

        btn = d(text="본인인증 요청")
        clicked_req = False
        for _ in range(20):
            if is_unexpected_digital_sales_home(d):
                break
            try:
                if btn.exists and btn.info.get("enabled", False):
                    btn.click()
                    time.sleep(0.6)
                    clicked_req = True
                    break
            except Exception:
                pass
            time.sleep(0.2)

        if is_unexpected_digital_sales_home(d):
            if attempt == 0 and recover_from_unexpected_home(d, f"인증발송 버튼대기 중 {job['name']}"):
                continue
            return (False, "UNEXPECTED_HOME")

        if not clicked_req:
            return (False, "REQ_BUTTON_FAIL")

        if not d(text="발송").exists:
            for _ in range(40):
                if is_unexpected_digital_sales_home(d):
                    break
                if d(text="발송").exists:
                    break
                time.sleep(0.2)

        if is_unexpected_digital_sales_home(d):
            if attempt == 0 and recover_from_unexpected_home(d, f"인증발송 팝업대기 중 {job['name']}"):
                continue
            return (False, "UNEXPECTED_HOME")

        if d(text="발송").exists:
            click_text_center(d, "발송", 0.45, 0.95)
            time.sleep(0.6)

        if is_unexpected_digital_sales_home(d):
            if attempt == 0 and recover_from_unexpected_home(d, f"인증발송 최종단계 중 {job['name']}"):
                continue
            return (False, "UNEXPECTED_HOME")

        return (True, "OK")

    return (False, "UNEXPECTED_HOME")

# ---------------------------
# emulator loop
# ---------------------------
def emulator_main_loop():
    global SIGN_IN_PROGRESS

    if not DO_EMULATOR:
        print("ℹ️ DO_EMULATOR=False")
        return

    d = connect_emulator()
    print("✅ 에뮬레이터 연결 완료:", ADB_SERIAL)

    last_reenter = 0.0

    while True:
        try:
            if STOP_FLAG:
                time.sleep(1.0)
                continue

            if SIGN_IN_PROGRESS:
                time.sleep(0.8)
                continue

            # 1) 인증발송 우선
            try:
                job = auth_q.get_nowait()
            except queue.Empty:
                job = None

            if job is not None:
                try:
                    retry = int(job.get("_retry", 0))
                    print(f"🚀 인증발송 작업 시작: {job['name']} / {job['phone11']}")

                    ok, reason = send_auth_request(d, job)
                    if ok:
                        now = time.time()
                        job["auth_sent_at"] = now
                        job["next_check_at"] = now + 60
                        job["last_check_at"] = 0.0
                        job["last_status"] = ""
                        with jobs_lock:
                            auth_sent_jobs[job["phone11"]] = job
                        print(f"✅ 인증발송 성공: {job['name']} / {job['phone11']} (최초 상태체크 60초 후)")
                        notify(f"인증발송 성공: {job['name']} / {job['phone11']} (최초 상태체크 60초 후)")
                    else:
                        if reason in ["APP_RESTART_FAIL", "NOT_READY", "UNEXPECTED_HOME"] and retry < AUTH_RETRY_MAX:
                            retry += 1
                            job["_retry"] = retry
                            print(f"⚠️ 인증발송 재시도 ({retry}/{AUTH_RETRY_MAX}) : {job['name']} / {job['phone11']} ({reason})")
                            auth_q.put(job)
                        else:
                            if STOP_ON_ERROR:
                                set_stop(f"인증발송 실패: {job['name']} / {job['phone11']} ({reason})")
                finally:
                    auth_q.task_done()

                time.sleep(0.2)
                continue

            # 2) 대기목록이 없으면 아무것도 안 함
            with jobs_lock:
                pending = [j for (p, j) in auth_sent_jobs.items() if p not in sign_started]
            if not pending:
                time.sleep(1.2)
                continue

            now = time.time()

            # 3) 백오프 + 배치 검색 대상 선정
            pending.sort(key=lambda x: x.get("next_check_at", 0.0))
            due = [j for j in pending if j.get("next_check_at", 0.0) <= now]
            if not due:
                time.sleep(SEARCH_LOOP_SLEEP_SEC)
                continue

            batch = due[:SEARCH_BATCH_PER_CYCLE]

            # 4) ✅ 실제 체크할 대상이 있을 때만
            #    앱 종료 없이 "주문현황 X -> 확인 -> 홈 -> 일반주문 재진입" 으로 갱신
            need_refresh = (now - last_reenter >= ORDER_REENTER_INTERVAL_SEC) or True
            if need_refresh:
                ok_refresh = refresh_order_status_by_reenter(d)
                if not ok_refresh:
                    print("❌ [emulator_main_loop] 재진입 갱신 실패 → 이번 배치 건너뜀")
                    time.sleep(SEARCH_LOOP_SLEEP_SEC)
                    continue
                last_reenter = time.time()
                time.sleep(0.5)

            for j in batch:
                if not auth_q.empty():
                    break

                print("🔎 검색 시도:", j["name"])

                st = check_one_job_status_by_search(d, j)
                j["last_check_at"] = time.time()
                j["last_status"] = st

                interval = compute_check_interval(j["auth_sent_at"])
                j["next_check_at"] = time.time() + interval

                with jobs_lock:
                    auth_sent_jobs[j["phone11"]] = j

                print(f"🧾 상태체크: {j['name']} / {j['phone11']} => {st or 'NONE'} (다음 {interval}s)")

                if st == "인증완료":
                    ok = try_open_ready_sign_detail(d, j)
                    if ok:
                        notify(f"✅ 인증완료 확인(매칭 OK): {j['name']} / {j['phone11']} → 관리/약정/상품담기 완료")
                    else:
                        back_to_order_status(d)

                time.sleep(0.6)

            time.sleep(SEARCH_LOOP_SLEEP_SEC)

        except Exception as e:
            print("❌ 에뮬레이터 루프 에러:", e)
            traceback.print_exc()
            if STOP_ON_ERROR:
                set_stop("에뮬레이터 루프 예외")
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

processed_phones = set()

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
        product_name = data.get("product_name", "")
        model_name = data.get("model_name", "")
        color_raw = data.get("color_raw", "")
        address = data.get("address", "")
        address_basic = data.get("address_basic", "") or address
        address_detail = data.get("address_detail", "")
        manage_raw = data.get("manage_raw", "")
        contract_raw = data.get("contract_raw", "")
        discount_raw = data.get("discount_raw", "")
        amount_raw = data.get("amount_raw", "")

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
        print("상품명:", product_name)
        print("모델명:", model_name)
        print("색상:", color_raw)
        print("기본주소:", address_basic)
        print("상세주소:", address_detail)
        print("관리:", manage_raw)
        print("약정:", contract_raw)
        print("할인:", discount_raw)
        print("금액:", amount_raw)
        print("-" * 40)

        auth_q.put({
            "name": name,
            "phone11": phone11,
            "birth": birth,
            "account": account,
            "zipcode": zipcode,
            "product_name": product_name,
            "model_name": model_name,
            "color_raw": color_raw,
            "address": address_basic,
            "address_basic": address_basic,
            "address_detail": address_detail,
            "manage_raw": manage_raw,
            "contract_raw": contract_raw,
            "discount_raw": discount_raw,
            "amount_raw": amount_raw,
            "_retry": 0,
        })

        time.sleep(1.2)

    except Exception as e:
        print("에러 발생:", e)
        traceback.print_exc()
        if STOP_ON_ERROR:
            set_stop("웹 모달 루프 예외")
        time.sleep(1.0)