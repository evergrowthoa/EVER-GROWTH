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
SIGN_SCAN_INTERVAL_SEC = 20    # 인증완료 스캔 주기
STOP_ON_ERROR = True           # 예상 밖 오류면 중단
AUTH_RETRY_MAX = 5             # 인증발송 화면복구 재시도 횟수

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

def normalize_model_code(s: str) -> str:
    raw = (str(s or "").strip().upper()).replace(" ", "")
    # CHPI-7430N 같은 패턴 우선 추출
    m = re.search(r"[A-Z]{2,}[A-Z0-9]*-\d+[A-Z0-9]*", raw)
    if m:
        return m.group(0)
    # 하이픈 없는 케이스도 대응
    m2 = re.search(r"[A-Z]{2,}\d+[A-Z0-9]*", raw)
    if m2:
        return m2.group(0)
    return raw

def normalize_color_raw(s: str) -> str:
    return (str(s or "").strip()).replace(" ", "")

def color_rule_and_keyword(color_raw: str):
    """
    올앤업 색상 -> 앱 검색결과 색상 판단 규칙
    반환: (mode, keyword)
      mode:
        - "contains": keyword 포함이면 OK (화이트/블루/핑크/베이지 등)
        - "ambiguous_gray": 올앤업이 '그레이'처럼 애매하면 절대 진행 금지(중단)
        - "exact": keyword가 그대로 포함되어야 함
    """
    c = normalize_color_raw(color_raw)

    if not c:
        return ("exact", "")

    # 자주 나오는 색상 카테고리
    if "화이트" in c:
        return ("contains", "화이트")  # 아이스 화이트 OK
    if "블루" in c:
        return ("contains", "블루")
    if "핑크" in c:
        return ("contains", "핑크")
    if "베이지" in c:
        return ("contains", "베이지")

    # 그레이는 위험: 페블/차콜/아이스 등으로 갈릴 수 있음
    if "그레이" in c:
        # 올앤업에서 '페블그레이/차콜그레이/아이스그레이'처럼 구체적으로 적힌 경우만 진행 가능
        if any(k in c for k in ["페블", "차콜", "아이스"]):
            # 예: 페블그레이 -> keyword=페블 (또는 전체 문자열 포함)
            return ("contains", c)  # 구체 문자열 포함이면 OK
        return ("ambiguous_gray", "그레이")

    # 그 외는 그대로 포함되어야 함
    return ("contains", c)

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

    # ✅ 전자서명 단계에 필요한 값 추가
    model_raw = find_input_near_label(modal, ["모델명", "모델", "제품코드", "상품코드"])
    color_raw = find_input_near_label(modal, ["색상", "컬러"])

    phone11 = normalize_phone11_only_010(phone_raw)
    model_code = normalize_model_code(model_raw)
    color_norm = normalize_color_raw(color_raw)

    return {
        "name": (name or "").strip(),
        "birth": (birth or "").strip(),
        "phone_raw": (phone_raw or "").strip(),
        "phone11": (phone11 or "").strip(),
        "account": (account or "").strip(),
        "zipcode": (zipcode or "").strip(),
        "model_raw": (model_raw or "").strip(),
        "model_code": (model_code or "").strip(),
        "color_raw": (color_raw or "").strip(),
        "color_norm": (color_norm or "").strip(),
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
    # 앱 내부로 보이면 OK
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
    # 목표: 모바일 주문 화면(일반 주문하기 보이는 화면)
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
    # 주문현황으로 이동: 모바일 주문 홈 복구 → 일반주문 클릭
    if ensure_on_order_status_screen(d):
        return True
    if not ensure_on_mobile_order_home(d):
        return False
    if d(text="일반주문").exists:
        click_first_clickable_text(d, "일반주문")
        time.sleep(0.8)
        return ensure_on_order_status_screen(d)
    return False

def try_click_auth_done_item(d) -> bool:
    # 주문현황 리스트에서 '인증완료' 하나 클릭
    if not ensure_on_order_status_screen(d):
        return False
    if d(text="인증완료").exists:
        click_first_clickable_text(d, "인증완료")
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

def wait_for_product_search_screen(d) -> bool:
    # 3번 캡처: 주문접수/상품검색/검색 버튼이 보이면 OK
    for _ in range(40):
        if d(text="상품검색").exists and d(text="검색").exists:
            return True
        time.sleep(0.2)
    return False

def do_product_search(d, model_code: str) -> bool:
    # 상품검색 화면에서 모델명 입력 후 검색
    edits = d(className="android.widget.EditText")
    if edits.count < 1:
        return False

    edits[0].click()
    time.sleep(0.1)
    edits[0].set_text(model_code)
    time.sleep(0.2)

    if d(text="검색").exists:
        click_first_clickable_text(d, "검색")
        time.sleep(0.8)
        return True

    return False

def collect_model_lines(d, model_code: str):
    """
    결과 리스트에서 model_code가 들어있는 '라인 텍스트'들을 수집
    예: "CHPI-7430N_WT, 113981, 아이스 화이트"
    """
    items = []
    try:
        nodes = d(textContains=model_code).all()
    except Exception:
        nodes = []

    for n in nodes:
        try:
            t = (n.get_text() or "").strip()
            if not t:
                continue
            if model_code not in t:
                continue
            items.append((t, n))
        except Exception:
            continue

    # 중복 제거(텍스트 기준)
    uniq = {}
    for t, n in items:
        uniq[t] = n
    return [(t, uniq[t]) for t in uniq.keys()]

def parse_model_line(line_text: str):
    """
    "CHPI-7430N_WT, 113981, 아이스 화이트" -> (model_token, color_text)
    """
    parts = [p.strip() for p in line_text.split(",")]
    model_token = parts[0] if len(parts) >= 1 else ""
    color_text = parts[-1] if len(parts) >= 2 else ""
    return model_token, color_text

def choose_one_product_candidate(model_code: str, color_raw: str, lines):
    """
    lines: [(line_text, node), ...]
    안전 규칙:
    - model_token은 model_code로 시작해야 함
    - 색상 규칙:
      - 화이트/핑크/블루/베이지: 포함이면 OK
      - 그레이(애매): 올앤업이 '그레이'만이면 즉시 중단
    - 최종 후보가 1개일 때만 진행, 아니면 중단
    """
    mode, keyword = color_rule_and_keyword(color_raw)

    # 그레이 애매하면 무조건 중단
    if mode == "ambiguous_gray":
        return ("STOP", f"색상 애매(올앤업='{color_raw}') → 그레이 계열은 자동 선택 금지", None, None)

    candidates = []
    for line_text, node in lines:
        model_token, color_text = parse_model_line(line_text)

        if not model_token.startswith(model_code):
            continue

        ct = (color_text or "").replace(" ", "")
        if keyword:
            kw = keyword.replace(" ", "")
            if kw not in ct:
                continue

        candidates.append((line_text, node, model_token, color_text))

    if len(candidates) == 0:
        return ("STOP", f"검색 결과에 일치 항목 없음 (model={model_code}, color='{color_raw}')", None, None)

    # 2개 이상이면 위험 → 중단
    if len(candidates) >= 2:
        # 어떤 후보가 있었는지 메시지로 남김
        preview = " | ".join([c[0] for c in candidates[:5]])
        return ("STOP", f"일치 후보가 {len(candidates)}개로 애매함 → 중단. 후보: {preview}", None, None)

    return ("OK", "선택 가능", candidates[0][1], candidates[0][0])

def click_candidate_node(node) -> bool:
    try:
        node.click()
        return True
    except Exception:
        try:
            # 클릭 실패 시 좌표 클릭(그래도 UI요소 기반)
            node.click_exists(timeout=0.1)
            return True
        except Exception:
            return False

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

    # 확인 팝업 [발송] 클릭
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
    전자서명(B) - 여기까지 구현:
    1) 주문현황에서 인증완료 진입
    2) 상세 화면에서 이름+전화번호 일치 확인(필수)
    3) 주문 이어서 하기 클릭
    4) 상품검색 화면에서 모델코드 검색
    5) 결과에서 모델/색상 조건이 '딱 1개'일 때만 클릭
    6) 다음 단계 미구현 → 안전 중단
    """
    if not goto_order_status(d):
        return False

    if not try_click_auth_done_item(d):
        return False

    # 상세 화면에서 우리 건 매칭
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

    # 주문 이어서 하기
    if not click_order_continue(d):
        if STOP_ON_ERROR:
            set_stop("인증완료 상세에서 '주문 이어서 하기' 버튼을 못 찾음")
        return False

    # 상품검색 화면 대기
    if not wait_for_product_search_screen(d):
        if STOP_ON_ERROR:
            set_stop("상품검색 화면 진입 실패(시간초과)")
        return False

    model_code = normalize_model_code(matched_job.get("model_code", ""))
    color_raw = matched_job.get("color_raw", "")

    if not model_code:
        set_stop(f"올앤업 모델명 없음 → 중단 (phone={matched_phone})")
        return False

    # 모델 검색
    ok_search = do_product_search(d, model_code)
    if not ok_search:
        set_stop(f"상품검색 입력/검색 실패 → 중단 (model={model_code})")
        return False

    # 결과 대기: model_code가 화면에 뜰 때까지
    found = False
    for _ in range(60):
        if d(textContains=model_code).exists:
            found = True
            break
        time.sleep(0.2)

    if not found:
        set_stop(f"검색 결과에 모델이 안 보임 → 중단 (model={model_code})")
        return False

    # 모델 라인 수집 및 후보 선택
    lines = collect_model_lines(d, model_code)
    status, msg, node, picked_line = choose_one_product_candidate(model_code, color_raw, lines)

    if status != "OK":
        set_stop(msg)
        return False

    # 최종 클릭
    clicked = click_candidate_node(node)
    if not clicked:
        set_stop(f"상품 후보 클릭 실패 → 중단 (line={picked_line})")
        return False

    # 중복 방지 기록
    with jobs_lock:
        sign_started.add(matched_phone)

    notify(f"상품 선택 완료(다음 단계 미구현으로 중단): {matched_job.get('name','')} / {matched_phone} / {model_code} / {color_raw}")
    set_stop("안전 중단: 상품 선택까지 완료. 다음 단계(약정/관리/전자서명)는 아직 미구현.")
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

        model_raw = data.get("model_raw", "")
        model_code = data.get("model_code", "")
        color_raw = data.get("color_raw", "")

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
        print("모델명 원문:", model_raw)
        print("모델코드:", model_code)
        print("색상:", color_raw)
        print("계좌:", account)
        print("우편번호:", zipcode)
        print("-" * 40)

        auth_q.put({
            "name": name,
            "phone11": phone11,
            "birth": birth,
            "account": account,
            "zipcode": zipcode,
            "model_raw": model_raw,
            "model_code": model_code,
            "color_raw": color_raw,
            "_retry": 0,
        })

        time.sleep(1.2)

    except Exception as e:
        print("에러 발생:", e)
        traceback.print_exc()
        if STOP_ON_ERROR:
            set_stop("웹 모달 루프 예외")
        time.sleep(1.0)