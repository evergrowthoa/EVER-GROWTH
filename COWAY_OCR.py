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
    1) set_text
    2) fast IME send_keys
    3) clipboard + (붙여넣기/붙여 넣기/Paste) 버튼 클릭

    추가:
    - 검색창은 가운데보다 '왼쪽 안쪽' 클릭이 더 안정적일 수 있음
    - 포커스를 2회 확보
    - 매 단계마다 실제 입력값 검증
    """
    def _refind():
        try:
            fresh = find_search_edittext(d)
            return fresh if fresh is not None else edit_obj
        except Exception:
            return edit_obj

    def _get_now(obj):
        try:
            return (obj.get_text() or "").strip()
        except Exception:
            try:
                info = obj.info or {}
                return str(info.get("text") or "").strip()
            except Exception:
                return ""

    try:
        obj = _refind()
        b = obj.info.get("bounds", {})
        left = int(b.get("left", 0))
        top = int(b.get("top", 0))
        right = int(b.get("right", 0))
        bottom = int(b.get("bottom", 0))

        cy = (top + bottom) // 2

        # ✅ 가운데 클릭 대신 왼쪽 안쪽 클릭
        focus_x = left + max(20, (right - left) // 6)

        # 1차 포커스
        d.click(focus_x, cy)
        time.sleep(0.2)

        # 2차 포커스
        d.click(focus_x, cy)
        time.sleep(0.2)

        obj = _refind()

        # 혹시 이전 값이 남아 있으면 먼저 비우기 시도
        try:
            obj.set_text("")
            time.sleep(0.15)
        except Exception:
            pass

        # 1) set_text
        try:
            obj = _refind()
            obj.set_text(text)
            time.sleep(0.35)
            got = _get_now(obj)
            print(f"🔍 [type_into_edittext] set_text 결과: '{got}'")
            if got and (got == text or text in got):
                return True
        except Exception as e:
            print("⚠️ [type_into_edittext] set_text 실패:", e)

        # 2) fast IME + send_keys
        try:
            d.set_fastinput_ime(True)
        except Exception:
            pass

        try:
            obj = _refind()
            d.click(focus_x, cy)
            time.sleep(0.15)
            d.send_keys(text, clear=True)
            time.sleep(0.35)
            got = _get_now(obj)
            print(f"🔍 [type_into_edittext] send_keys 결과: '{got}'")
            if got and (got == text or text in got):
                return True
        except Exception as e:
            print("⚠️ [type_into_edittext] send_keys 실패:", e)

        # 3) clipboard paste
        try:
            obj = _refind()
            d.set_clipboard(text)
            time.sleep(0.15)

            d.long_click(focus_x, cy, 0.7)
            time.sleep(0.45)

            paste_candidates = ["붙여넣기", "붙여 넣기", "Paste", "PASTE"]
            clicked = False

            for t in paste_candidates:
                if d(text=t).exists:
                    d(text=t).click()
                    clicked = True
                    time.sleep(0.35)
                    break

            if not clicked:
                if d(textContains="붙여").exists:
                    d(textContains="붙여").click()
                    clicked = True
                    time.sleep(0.35)
                elif d(textContains="Paste").exists:
                    d(textContains="Paste").click()
                    clicked = True
                    time.sleep(0.35)

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

def is_tab_selected(d, tab_text: str) -> bool:
    """
    이 앱은 selected/checked 값이 잘 안 잡히는 경우가 많아서
    현재는 신뢰하지 않는다.
    """
    return False

def ensure_general_tab(d, force_click: bool = False) -> bool:
    """
    ✅ 주문현황에서 '일반' 탭 복구 + 실제 복구 여부 검증
    검증 기준:
    - 일반 탭이면 좌측 검색 드롭다운이 '고객/상호'
    - 코라솔 쪽이면 '고객'으로 보이는 경우가 있음
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
        print(f"⚠️ [ensure_general_tab] 일반 탭 아님 / { _debug_mode_label() }")
        return False

    try:
        w, h = d.window_size()

        # 1차: 상단의 '일반' 텍스트 객체 자체를 직접 클릭 시도
        objs = d(text="일반")
        try:
            cnt = objs.count
        except Exception:
            cnt = 0

        for i in range(cnt):
            try:
                o = objs[i]
                b = o.info.get("bounds", {})
                left = int(b.get("left", 0))
                top = int(b.get("top", 0))
                right = int(b.get("right", 0))
                bottom = int(b.get("bottom", 0))

                # 상단 탭 영역에 있는 '일반'만 사용
                if top < int(h * 0.05) or top > int(h * 0.22):
                    continue

                cx = (left + right) // 2
                cy = (top + bottom) // 2

                d.click(cx, cy)
                time.sleep(0.45)

                if _is_general_ready():
                    print(f"✅ [ensure_general_tab] 일반 탭 텍스트 클릭 성공 ({cx}, {cy})")
                    return True
            except Exception:
                continue

        # 2차: 텍스트 클릭이 안 먹으면 좌표 fallback
        candidate_points = [
            (int(w * 0.18), int(h * 0.105)),
            (int(w * 0.20), int(h * 0.105)),
            (int(w * 0.16), int(h * 0.105)),
            (int(w * 0.18), int(h * 0.115)),
            (int(w * 0.20), int(h * 0.115)),
            (int(w * 0.16), int(h * 0.115)),
        ]

        for x, y in candidate_points:
            try:
                d.click(x, y)
                time.sleep(0.45)

                if _is_general_ready():
                    print(f"✅ [ensure_general_tab] 일반 탭 좌표 클릭 성공 ({x}, {y})")
                    return True
            except Exception:
                continue

        print(f"❌ [ensure_general_tab] 일반 탭 복구 실패 / { _debug_mode_label() }")
        return False

    except Exception as e:
        print("❌ [ensure_general_tab] 예외:", e)
        return False

def enter_order_status(d) -> bool:
    if d(text="주문현황").exists:
        return True

    if not ensure_mobile_order_home(d):
        return False

    if d(text="일반주문").exists:
        click_text_center(d, "일반주문", 0.25, 0.80)
        time.sleep(1.0)

        ok = d(text="주문현황").exists
        if ok:
            ensure_general_tab(d, force_click=True)
            time.sleep(0.4)
        return ok

    return d(text="주문현황").exists

def exit_order_status_to_mobile_home(d) -> bool:
    """
    ✅ 주문현황에서 '진짜' 모바일 주문 홈으로 빠져나오는지 확인
    성공 조건:
    - 주문현황 텍스트가 사라져야 함
    - 일반 주문하기가 보여야 함
    """
    if not d(text="주문현황").exists:
        return d(text="일반 주문하기").exists

    for attempt in range(2):
        print(f"🔄 [exit_order_status_to_mobile_home] 주문현황 종료 시도 {attempt + 1}/2")

        try:
            w, h = d.window_size()
            d.click(int(w * 0.965), int(h * 0.075))
            time.sleep(0.35)
        except Exception as e:
            print("⚠️ [exit_order_status_to_mobile_home] X 클릭 실패:", e)

        confirmed = False
        for _ in range(15):
            try:
                if d(text="확인").exists and d(text="취소").exists:
                    click_text_center(d, "확인", 0.45, 0.95)
                    time.sleep(0.9)
                    confirmed = True
                    break
            except Exception:
                pass
            time.sleep(0.2)

        if confirmed:
            print("✅ [exit_order_status_to_mobile_home] 확인 팝업 처리 완료")
        else:
            print("ℹ️ [exit_order_status_to_mobile_home] 확인 팝업 없음 또는 미검출")

        for _ in range(20):
            try:
                in_order_status = d(text="주문현황").exists
            except Exception:
                in_order_status = False

            try:
                on_mobile_home = d(text="일반 주문하기").exists
            except Exception:
                on_mobile_home = False

            if (not in_order_status) and on_mobile_home:
                print("✅ [exit_order_status_to_mobile_home] 모바일 주문 홈 복귀 성공")
                return True

            time.sleep(0.25)

        print("⚠️ [exit_order_status_to_mobile_home] 아직 주문현황에서 완전히 빠지지 못함")

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

def get_first_status_badge_in_results(d):
    status_texts = ["인증완료", "인증입력", "서명입력", "주문확정", "주문삭제", "주문불가"]
    try:
        w, h = d.window_size()
        found = []
        for st in status_texts:
            for o in d(text=st).all():
                try:
                    b = o.info.get("bounds", {})
                    top = int(b.get("top", 0))
                    if top < int(h * 0.35):
                        continue
                    found.append((top, st, o))
                except Exception:
                    continue
        if not found:
            return ("", None)
        found.sort(key=lambda x: x[0])
        return (found[0][1], found[0][2])
    except Exception:
        return ("", None)

def open_detail_by_status_badge(d, badge_obj) -> bool:
    try:
        b = badge_obj.info.get("bounds", {})
        cx = (int(b.get("left", 0)) + int(b.get("right", 0))) // 2
        cy = (int(b.get("top", 0)) + int(b.get("bottom", 0))) // 2
        d.click(cx, cy)
        time.sleep(0.8)
        return True
    except Exception:
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

def check_one_job_status_by_search(d, job: dict) -> str:
    if not enter_order_status(d):
        print("❌ [status_check] 주문현황 진입 실패")
        return ""

    # ✅ 검색 직전마다 무조건 일반 탭으로 복구
    if not ensure_general_tab(d, force_click=True):
        print("❌ [status_check] 일반 탭 복구 실패")
        return ""

    dismiss_search_mode_dropdown_if_open(d)

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

    trigger_search(d, edit)
    st, badge = get_first_status_badge_in_results(d)
    print(f"🧾 [status_check] 검색 결과 상태: {st or 'NONE'}")
    return st

def try_open_ready_sign_detail(d, job: dict) -> bool:
    if not enter_order_status(d):
        print("❌ [ready_sign] 주문현황 진입 실패")
        return False

    # ✅ 상세 진입 검색 전에도 무조건 일반 탭으로 복구
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

    trigger_search(d, edit)

    st, badge = get_first_status_badge_in_results(d)
    print(f"🧾 [ready_sign] 검색 결과 상태: {st or 'NONE'}")

    if st != "인증완료" or badge is None:
        return False

    open_detail_by_status_badge(d, badge)

    if match_detail_by_name_phone(d, job["name"], job["phone11"]):
        print("✅ [ready_sign] 상세 고객명/전화번호 매칭 성공")
        return True

    print("❌ [ready_sign] 상세 고객명/전화번호 매칭 실패 → 주문현황으로 복귀")
    back_to_order_status(d)
    return False

# ---------------------------
# 인증발송
# ---------------------------
def send_auth_request(d, job: dict):
    if not restart_digital_sales_app(d):
        return (False, "APP_RESTART_FAIL")

    if not ensure_mobile_order_home(d):
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
    edits[0].set_text(job["name"])

    edits[1].click()
    time.sleep(0.1)
    edits[1].set_text(job["phone11"])

    btn = d(text="본인인증 요청")
    for _ in range(20):
        if btn.exists and btn.info.get("enabled", False):
            btn.click()
            time.sleep(0.6)
            break
        time.sleep(0.2)

    if not d(text="발송").exists:
        for _ in range(40):
            if d(text="발송").exists:
                break
            time.sleep(0.2)

    if d(text="발송").exists:
        click_text_center(d, "발송", 0.45, 0.95)
        time.sleep(0.6)

    return (True, "OK")

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
                        job["next_check_at"] = now + 10
                        job["last_check_at"] = 0.0
                        job["last_status"] = ""
                        with jobs_lock:
                            auth_sent_jobs[job["phone11"]] = job
                        print(f"✅ 인증발송 성공: {job['name']} / {job['phone11']}")
                        notify(f"인증발송 성공: {job['name']} / {job['phone11']}")
                    else:
                        if reason in ["APP_RESTART_FAIL", "NOT_READY"] and retry < AUTH_RETRY_MAX:
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
                        notify(f"✅ 인증완료 확인(매칭 OK): {j['name']} / {j['phone11']}  → 전자서명 단계로 진행 가능")
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