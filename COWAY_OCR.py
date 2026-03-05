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

BUILD_ID = "COWAY_OCR_BUILD_2026-03-05_003"
print("✅ BUILD:", BUILD_ID)

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
# 디지털세일즈 앱 설정
#  - 가능하면 패키지명을 넣어야 force-stop/app_start가 안정적
# ===========================
DIGITAL_SALES_APP_NAME = "디지털세일즈"
DIGITAL_SALES_PACKAGE = ""   # 예: "com.xxx.yyy" (모르면 빈칸 유지)

# ===========================
# 운영 옵션
# ===========================
STOP_ON_ERROR = True
AUTH_RETRY_MAX = 3

# 주문현황 갱신(재진입) 주기
ORDER_REENTER_INTERVAL_SEC = 90   # 60~180 추천

# 검색 배치: 한 사이클에 몇 건만 검색할지(10~20건 있어도 2~3개씩만)
SEARCH_BATCH_PER_CYCLE = 3

# 검색 루프 주기(너무 짧게 돌리면 화면 흔들림)
SEARCH_LOOP_SLEEP_SEC = 2.0

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
auth_q = queue.Queue()

# 인증발송 성공한 건(대기목록)
# phone11 -> job(dict)
auth_sent_jobs = {}
jobs_lock = threading.Lock()

# 전자서명 진입(중복 방지)
sign_started = set()

# 전자서명 진행중이면 인증발송은 후순위(큐에 쌓기만)
SIGN_IN_PROGRESS = False

# 중단 플래그
STOP_FLAG = False

def notify(msg: str):
    # TODO: 텔레그램/슬랙 붙일 자리
    print("🔔", msg)

def set_stop(reason: str):
    global STOP_FLAG
    STOP_FLAG = True
    print("🛑 자동화 중단:", reason)
    notify(reason)

# ===========================
# 유틸
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

def compute_check_interval(auth_sent_at: float) -> int:
    """
    대기시간이 길어질수록 체크 간격을 늘려서(백오프) 10~20건도 부담 없이 운영.
    """
    now = time.time()
    elapsed = max(0, now - auth_sent_at)

    if elapsed < 10 * 60:      # 0~10분
        return 120            # 2분
    if elapsed < 60 * 60:      # 10~60분
        return 300            # 5분
    if elapsed < 6 * 60 * 60:  # 1~6시간
        return 900            # 15분
    if elapsed < 24 * 60 * 60: # 6~24시간
        return 1800           # 30분
    return 3600               # 1시간

# ===========================
# Selenium: 모달 추출
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
# uiautomator2: 공통 클릭/앱 이동
# ===========================
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

def restart_digital_sales_app(d) -> bool:
    """
    인증발송 시작시: 앱 상태 꼬임 방지용 강제 재시작(권장)
    """
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

        # fallback: 홈으로 나가서 아이콘 클릭
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
    """
    디지털세일즈 앱 내부에서 '모바일 주문' 탭 → '일반 주문하기' 화면
    """
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

def enter_order_status(d) -> bool:
    """
    모바일 주문 홈에서 '일반주문'을 눌러 주문현황 진입
    """
    if d(text="주문현황").exists:
        return True

    if not ensure_mobile_order_home(d):
        return False

    if d(text="일반주문").exists:
        click_text_center(d, "일반주문", 0.25, 0.80)
        time.sleep(1.0)
        return d(text="주문현황").exists

    return d(text="주문현황").exists

def exit_order_status_to_mobile_home(d) -> bool:
    """
    주문현황 우측 상단 X → 팝업 '확인' → 모바일 주문 화면으로 복귀
    """
    if not d(text="주문현황").exists:
        return True

    try:
        w, h = d.window_size()
        d.click(int(w * 0.965), int(h * 0.075))  # X
        time.sleep(0.3)
    except Exception:
        pass

    # 종료 팝업 확인
    for _ in range(15):
        if d(text="확인").exists and d(text="취소").exists:
            click_text_center(d, "확인", 0.45, 0.95)
            time.sleep(0.8)
            break
        time.sleep(0.2)

    # 모바일 주문 홈 도달 확인
    for _ in range(20):
        if d(text="일반 주문하기").exists or d(text="모바일 주문").exists:
            return True
        time.sleep(0.2)

    return d(text="일반 주문하기").exists or d(text="모바일 주문").exists

def refresh_order_status_by_reenter(d) -> bool:
    """
    ✅ 새로고침 아이콘 대신 '재진입'으로 갱신(안정)
    """
    if not d(text="주문현황").exists:
        # 이미 주문현황이 아니면 그냥 진입 시도
        return enter_order_status(d)

    ok = exit_order_status_to_mobile_home(d)
    if not ok:
        return False
    return enter_order_status(d)

# ===========================
# 주문현황: 검색 방식(대기목록 1건씩)
# ===========================
def set_search_mode_contact_if_possible(d) -> bool:
    """
    드롭다운이 '고객/상호' 또는 '고객'이면 눌러서 '연락처'로 변경 시도.
    실패해도 안전하게 False 반환(이 경우 이름 검색으로 fallback)
    """
    if not d(text="주문현황").exists:
        return False

    # 드롭다운 현재 라벨 후보
    label_candidates = ["고객/상호", "고객"]
    opened = False
    for lb in label_candidates:
        if d(text=lb).exists:
            click_text_center(d, lb, 0.15, 0.40)
            time.sleep(0.4)
            opened = True
            break

    if not opened:
        return False

    if d(text="연락처").exists:
        click_text_center(d, "연락처", 0.10, 0.60)
        time.sleep(0.4)
        return True

    return False

def find_search_edittext(d):
    """
    주문현황 상단 검색창 EditText 찾기(상단 15%~35% 범위)
    """
    try:
        w, h = d.window_size()
        edits = d(className="android.widget.EditText").all()
        for e in edits:
            try:
                b = e.info.get("bounds", {})
                top = int(b.get("top", 0))
                if top < int(h * 0.12) or top > int(h * 0.35):
                    continue
                return e
            except Exception:
                continue
    except Exception:
        return None
    return None

def click_search_icon_near_input(d, edit_obj):
    """
    돋보기 아이콘은 텍스트가 없을 수 있어 입력창 오른쪽을 좌표 클릭
    """
    try:
        w, h = d.window_size()
        b = edit_obj.info.get("bounds", {})
        cy = (int(b.get("top", 0)) + int(b.get("bottom", 0))) // 2
        d.click(int(w * 0.955), cy)
        time.sleep(0.6)
        return True
    except Exception:
        return False

def get_first_status_badge_in_results(d):
    """
    검색 결과에서 첫 번째 row의 상태 뱃지 텍스트를 추출
    (인증입력/인증완료/서명입력/주문확정/주문삭제 등)
    """
    status_texts = ["인증완료", "인증입력", "서명입력", "주문확정", "주문삭제", "주문불가"]
    try:
        w, h = d.window_size()
        found = []
        for st in status_texts:
            for o in d(text=st).all():
                try:
                    b = o.info.get("bounds", {})
                    top = int(b.get("top", 0))
                    # 결과 리스트 영역만
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
    """
    주문현황에서 1건 검색 → 상태 확인.
    반환: 상태 텍스트(없으면 "")
    """
    if not enter_order_status(d):
        return ""

    # 검색 모드: 연락처로 바꾸기 시도(안 되면 이름검색)
    use_contact = set_search_mode_contact_if_possible(d)

    edit = find_search_edittext(d)
    if edit is None:
        return ""

    query = job["phone11"] if use_contact else job["name"]

    try:
        edit.click()
        time.sleep(0.1)
        edit.set_text(query)
        time.sleep(0.3)
    except Exception:
        return ""

    # 돋보기 클릭(검색 확정)
    click_search_icon_near_input(d, edit)

    # 상태 뱃지 읽기
    st, badge = get_first_status_badge_in_results(d)
    return st

def try_open_ready_sign_detail(d, job: dict) -> bool:
    """
    상태가 인증완료로 보일 때:
    - 인증완료 배지 클릭 → 상세 진입
    - 상세에서 이름+전화번호 완전일치면 True
    """
    if not enter_order_status(d):
        return False

    # 연락처 검색 모드 시도
    use_contact = set_search_mode_contact_if_possible(d)

    edit = find_search_edittext(d)
    if edit is None:
        return False

    query = job["phone11"] if use_contact else job["name"]

    try:
        edit.click()
        time.sleep(0.1)
        edit.set_text(query)
        time.sleep(0.3)
    except Exception:
        return False

    click_search_icon_near_input(d, edit)

    st, badge = get_first_status_badge_in_results(d)
    if st != "인증완료" or badge is None:
        return False

    open_detail_by_status_badge(d, badge)

    if match_detail_by_name_phone(d, job["name"], job["phone11"]):
        return True

    # 내 건 아니면 즉시 복귀
    back_to_order_status(d)
    return False

# ===========================
# 인증발송(A)
# ===========================
def send_auth_request(d, job: dict):
    """
    인증발송은 앱 상태 꼬임 방지 위해 '강제 재시작' 후 진행
    """
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
        # 팝업이 늦는 경우 대기
        for _ in range(40):
            if d(text="발송").exists:
                break
            time.sleep(0.2)

    if d(text="발송").exists:
        click_text_center(d, "발송", 0.45, 0.95)
        time.sleep(0.6)

    return (True, "OK")

# ===========================
# 에뮬레이터 메인 루프
# ===========================
def emulator_main_loop():
    global SIGN_IN_PROGRESS

    if not DO_EMULATOR:
        print("ℹ️ DO_EMULATOR=False")
        return

    d = connect_emulator()
    print("✅ 에뮬레이터 연결 완료:", ADB_SERIAL)

    last_reenter = 0.0
    rr_index = 0

    while True:
        try:
            if STOP_FLAG:
                time.sleep(1.0)
                continue

            # 전자서명 진행중이면 인증발송 후순위
            if SIGN_IN_PROGRESS:
                time.sleep(0.8)
                continue

            # 1) 인증발송 큐 우선 처리
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
                        # 대기목록 저장(백오프 메타 포함)
                        now = time.time()
                        job["auth_sent_at"] = now
                        job["next_check_at"] = now + 60  # 첫 체크는 1분 뒤
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

            # 2) 대기목록이 없으면 아무것도 안 함(화면 흔들림 방지)
            with jobs_lock:
                pending = [j for (p, j) in auth_sent_jobs.items() if p not in sign_started]
            if not pending:
                time.sleep(1.2)
                continue

            now = time.time()

            # 3) 주문현황 “재진입 갱신” (주기)
            if now - last_reenter >= ORDER_REENTER_INTERVAL_SEC:
                # 주문현황이든 뭐든 상관없이 재진입으로 갱신
                refresh_order_status_by_reenter(d)
                last_reenter = now

            # 4) 라운드로빈 + 백오프: 이번 사이클에 2~3건만 검색
            pending.sort(key=lambda x: x.get("next_check_at", 0.0))
            due = [j for j in pending if j.get("next_check_at", 0.0) <= now]

            if not due:
                time.sleep(SEARCH_LOOP_SLEEP_SEC)
                continue

            batch = due[:SEARCH_BATCH_PER_CYCLE]

            for j in batch:
                # 다시 인증발송이 들어오면 즉시 빠져나가 후순위 보장
                if not auth_q.empty():
                    break

                st = check_one_job_status_by_search(d, j)
                j["last_check_at"] = time.time()
                j["last_status"] = st

                # 다음 체크 시각 갱신(백오프)
                interval = compute_check_interval(j["auth_sent_at"])
                j["next_check_at"] = time.time() + interval

                with jobs_lock:
                    auth_sent_jobs[j["phone11"]] = j

                print(f"🧾 상태체크: {j['name']} / {j['phone11']} => {st or 'NONE'} (다음 {interval}s)")

                if st == "인증완료":
                    # 상세 진입 + 이름/번호 완전일치 검증
                    ok = try_open_ready_sign_detail(d, j)
                    if ok:
                        # 여기서부터 전자서명 자동화 이어붙일 자리
                        notify(f"✅ 인증완료 확인(매칭 OK): {j['name']} / {j['phone11']}  → 전자서명 단계로 진행 가능")
                        # 지금은 안전하게 자동 진행은 하지 않고, 다음 단계 코딩 시 여기에 붙임
                        # sign_started.add(j['phone11']) 를 하면 중복 체크를 막을 수 있음
                        # (원하면 지금부터 시작 처리로 바꿔줄게)
                        # 예: sign_started.add(j['phone11'])

                        # 안전: 상세 화면에 머물지 않게 주문현황로 복귀
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

# ===========================
# 3. 중복 방지
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