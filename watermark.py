import os
import re
import time
import json
import queue
import sqlite3
import threading
import traceback
import urllib.request
import urllib.error

import uiautomator2 as u2
from slack_sdk import WebClient
from slack_sdk.errors import SlackApiError


BUILD_ID = "DIGITAL_SALES_SLACK_WATERMARK_2026-03-18_002"
print("✅ BUILD:", BUILD_ID)

# =========================
# 기본 설정
# =========================
ADB_SERIAL = "emulator-5554"
DIGITAL_SALES_APP_NAME = "디지털세일즈"
DIGITAL_SALES_PACKAGE = ""

SLACK_BOT_TOKEN = os.environ.get("SLACK_BOT_TOKEN", "").strip()
SLACK_CHANNEL_ID = os.environ.get("SLACK_CHANNEL_ID", "C0AM6V8LEF8").strip()

ENABLE_TELEGRAM_NOTIFY = True
TELEGRAM_BOT_TOKEN = os.environ.get("TG_BOT_TOKEN", "8777418618:AAGNdHu1C6Lz5yQJn8Xq1hCxWhEUp8Ao4O8").strip()
TELEGRAM_CHAT_ID_ME = os.environ.get("TG_CHAT_ID_ME", "8759444041").strip()

NOTIFY_PREFIX = "[디지털세일즈 Slack 조회]"

POLL_INTERVAL_SEC = 3.0
SEARCH_NAME = "정재경"
DETAIL_EXPECT_NAME = "정재경"
DETAIL_EXPECT_PHONE11 = "01043271050"
TARGET_STATUS_TEXT = "설치입력"

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "digital_sales_slack_watermark.sqlite3")

STOP_FLAG = False
LAST_STOP_REASON = ""
ADB_DEVICE = None

job_q = queue.Queue()
db_lock = threading.RLock()
runtime_lock = threading.RLock()

slack_client = WebClient(token=SLACK_BOT_TOKEN) if SLACK_BOT_TOKEN else None


# =========================
# 공통 알림
# =========================
def _post_json(url: str, payload: dict, timeout_sec: float = 10.0):
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(
        url,
        data=body,
        headers={"Content-Type": "application/json; charset=utf-8"},
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=timeout_sec) as resp:
        return resp.read().decode("utf-8", errors="ignore")


def _send_telegram_message(chat_id: str, text: str) -> bool:
    token = str(TELEGRAM_BOT_TOKEN or "").strip()
    chat_id = str(chat_id or "").strip()

    if not ENABLE_TELEGRAM_NOTIFY:
        return False
    if not token or not chat_id:
        return False

    url = f"https://api.telegram.org/bot{token}/sendMessage"
    payload = {
        "chat_id": chat_id,
        "text": text[:4000],
    }

    try:
        _post_json(url, payload, timeout_sec=10.0)
        return True
    except Exception as e:
        print("⚠️ [notify] Telegram 전송 실패:", e)
        return False


def _build_notify_text(msg: str, level: str = "info") -> str:
    icon = "🔔"
    lv = str(level or "info").strip().lower()

    if lv in ["error", "err", "fail", "stop"]:
        icon = "🛑"
    elif lv in ["success", "done", "ok"]:
        icon = "✅"
    elif lv in ["warn", "warning", "hold"]:
        icon = "⚠️"
    elif lv in ["progress", "status"]:
        icon = "📌"

    now_str = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
    return f"{icon} {NOTIFY_PREFIX}\n시간: {now_str}\n메시지: {msg}"


def notify(msg: str, level: str = "info"):
    text = _build_notify_text(msg, level=level)
    print("🔔", msg)
    _send_telegram_message(TELEGRAM_CHAT_ID_ME, text)


def notify_error(msg: str):
    notify(msg, level="error")


def notify_progress(msg: str):
    notify(msg, level="progress")


def notify_success(msg: str):
    notify(msg, level="success")


def set_stop(reason: str):
    global STOP_FLAG, LAST_STOP_REASON
    STOP_FLAG = True
    LAST_STOP_REASON = str(reason or "").strip()
    print("🛑 자동화 중단:", reason)
    notify_error(reason)


def clear_stop():
    global STOP_FLAG, LAST_STOP_REASON
    STOP_FLAG = False
    LAST_STOP_REASON = ""
    print("✅ 자동화 정지 해제")


# =========================
# DB
# =========================
def _db_conn():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    with db_lock:
        conn = _db_conn()
        try:
            conn.execute("""
                CREATE TABLE IF NOT EXISTS slack_seen (
                    message_ts TEXT PRIMARY KEY,
                    raw_text TEXT DEFAULT '',
                    matched_watermark8 TEXT DEFAULT '',
                    created_at REAL DEFAULT 0
                )
            """)

            conn.execute("""
                CREATE TABLE IF NOT EXISTS jobs (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    message_ts TEXT UNIQUE,
                    channel_id TEXT DEFAULT '',
                    thread_ts TEXT DEFAULT '',
                    watermark8 TEXT DEFAULT '',
                    status TEXT DEFAULT '',
                    last_error TEXT DEFAULT '',
                    created_at REAL DEFAULT 0,
                    updated_at REAL DEFAULT 0
                )
            """)

            conn.execute("CREATE INDEX IF NOT EXISTS idx_jobs_status ON jobs(status)")
            conn.execute("CREATE INDEX IF NOT EXISTS idx_jobs_updated_at ON jobs(updated_at DESC)")
            conn.commit()
        finally:
            conn.close()


def db_has_seen_message(message_ts: str) -> bool:
    with db_lock:
        conn = _db_conn()
        try:
            cur = conn.execute("SELECT 1 FROM slack_seen WHERE message_ts = ?", (message_ts,))
            return cur.fetchone() is not None
        finally:
            conn.close()


def db_mark_seen_message(message_ts: str, raw_text: str, watermark8: str):
    now = time.time()
    with db_lock:
        conn = _db_conn()
        try:
            conn.execute("""
                INSERT OR IGNORE INTO slack_seen (message_ts, raw_text, matched_watermark8, created_at)
                VALUES (?, ?, ?, ?)
            """, (
                str(message_ts or "").strip(),
                str(raw_text or "").strip(),
                str(watermark8 or "").strip(),
                now,
            ))
            conn.commit()
        finally:
            conn.close()


def db_create_job(message_ts: str, channel_id: str, thread_ts: str, watermark8: str):
    now = time.time()
    with db_lock:
        conn = _db_conn()
        try:
            conn.execute("""
                INSERT OR IGNORE INTO jobs (
                    message_ts, channel_id, thread_ts, watermark8,
                    status, last_error, created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                str(message_ts or "").strip(),
                str(channel_id or "").strip(),
                str(thread_ts or "").strip(),
                str(watermark8 or "").strip(),
                "대기",
                "",
                now,
                now,
            ))
            conn.commit()
        finally:
            conn.close()


def db_get_job_by_message_ts(message_ts: str):
    with db_lock:
        conn = _db_conn()
        try:
            cur = conn.execute("SELECT * FROM jobs WHERE message_ts = ?", (str(message_ts or "").strip(),))
            row = cur.fetchone()
            return dict(row) if row else None
        finally:
            conn.close()


def db_update_job_status(message_ts: str, status: str, last_error: str = ""):
    with db_lock:
        conn = _db_conn()
        try:
            conn.execute("""
                UPDATE jobs
                SET status = ?, last_error = ?, updated_at = ?
                WHERE message_ts = ?
            """, (
                str(status or "").strip(),
                str(last_error or "").strip(),
                time.time(),
                str(message_ts or "").strip(),
            ))
            conn.commit()
        finally:
            conn.close()


def db_list_waiting_jobs():
    with db_lock:
        conn = _db_conn()
        try:
            cur = conn.execute("""
                SELECT *
                FROM jobs
                WHERE status IN ('대기')
                ORDER BY created_at ASC, id ASC
            """)
            return [dict(r) for r in cur.fetchall()]
        finally:
            conn.close()


# =========================
# Slack
# =========================
def extract_watermark8(text: str) -> str:
    s = str(text or "").strip()
    m = re.fullmatch(r"물마크\s+(\d{8})", s)
    if not m:
        return ""
    return m.group(1)


def slack_reply_in_thread(channel_id: str, thread_ts: str, text: str):
    if not slack_client:
        print("⚠️ Slack 토큰 미설정 → Slack 답장 생략:", text)
        return

    try:
        slack_client.chat_postMessage(
            channel=channel_id,
            thread_ts=thread_ts,
            text=text,
        )
    except SlackApiError as e:
        print("⚠️ Slack 답장 실패:", e)


def slack_poll_loop():
    if not slack_client or not SLACK_CHANNEL_ID:
        print("⚠️ Slack 설정 없음 → Slack polling 미실행")
        return

    print("✅ Slack polling 시작:", SLACK_CHANNEL_ID)

    while True:
        try:
            if STOP_FLAG:
                time.sleep(1.0)
                continue

            resp = slack_client.conversations_history(
                channel=SLACK_CHANNEL_ID,
                limit=20,
            )

            messages = resp.get("messages", []) or []
            messages = list(reversed(messages))

            for msg in messages:
                message_ts = str(msg.get("ts") or "").strip()
                raw_text = str(msg.get("text") or "").strip()
                subtype = str(msg.get("subtype") or "").strip()
                bot_id = str(msg.get("bot_id") or "").strip()
                thread_ts = str(msg.get("thread_ts") or "").strip() or message_ts

                if not message_ts:
                    continue

                if db_has_seen_message(message_ts):
                    continue

                if subtype:
                    db_mark_seen_message(message_ts, raw_text, "")
                    continue

                if bot_id:
                    db_mark_seen_message(message_ts, raw_text, "")
                    continue

                watermark8 = extract_watermark8(raw_text)
                db_mark_seen_message(message_ts, raw_text, watermark8)

                if not watermark8:
                    continue

                print(f"📥 Slack 조회요청 감지: ts={message_ts} / 물마크={watermark8}")

                db_create_job(
                    message_ts=message_ts,
                    channel_id=SLACK_CHANNEL_ID,
                    thread_ts=thread_ts,
                    watermark8=watermark8,
                )

                row = db_get_job_by_message_ts(message_ts)
                if row:
                    job_q.put(dict(row))
                    slack_reply_in_thread(
                        channel_id=SLACK_CHANNEL_ID,
                        thread_ts=thread_ts,
                        text=f"조회 접수 / 물마크 {watermark8}",
                    )
                    notify_progress(f"Slack 조회 접수: 물마크 {watermark8}")

        except Exception as e:
            print("❌ Slack polling 예외:", e)
            traceback.print_exc()
            notify_error(f"Slack polling 예외: {e}")

        time.sleep(POLL_INTERVAL_SEC)


# =========================
# uiautomator2 / 디지털세일즈 공통
# =========================
def normalize_digits(s: str) -> str:
    return re.sub(r"\D", "", str(s or ""))


def phone11_to_display(phone11: str) -> str:
    if not phone11 or len(phone11) != 11:
        return ""
    return f"{phone11[0:3]}-{phone11[3:7]}-{phone11[7:11]}"


def connect_emulator():
    global ADB_DEVICE
    d = u2.connect(ADB_SERIAL)
    d.implicitly_wait(5.0)
    ADB_DEVICE = d
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


def type_into_edittext(d, edit_obj, text: str) -> bool:
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
        b = edit_obj.info.get("bounds", {})
        left = int(b.get("left", 0))
        top = int(b.get("top", 0))
        right = int(b.get("right", 0))
        bottom = int(b.get("bottom", 0))
        cy = (top + bottom) // 2
        focus_x = left + max(20, (right - left) // 6)

        d.click(focus_x, cy)
        time.sleep(0.35)

        try:
            edit_obj.set_text("")
            time.sleep(0.25)
        except Exception:
            pass

        try:
            edit_obj.set_text(text)
            time.sleep(0.50)
            got = _get_now(edit_obj)
            if got and (got == text or text in got):
                return True
        except Exception:
            pass

        try:
            d.set_fastinput_ime(True)
        except Exception:
            pass

        try:
            d.click(focus_x, cy)
            time.sleep(0.30)
            d.send_keys(text, clear=True)
            time.sleep(0.55)
            got = _get_now(edit_obj)
            if got and (got == text or text in got):
                return True
        except Exception:
            pass

        return False
    except Exception:
        return False


def restart_digital_sales_app(d) -> bool:
    try:
        if DIGITAL_SALES_PACKAGE:
            try:
                d.app_stop(DIGITAL_SALES_PACKAGE)
                time.sleep(1.0)
            except Exception:
                pass
            try:
                d.app_start(DIGITAL_SALES_PACKAGE)
                time.sleep(2.5)
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
            time.sleep(2.5)
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
    try:
        if d(text="주문현황").exists:
            return False
    except Exception:
        pass

    try:
        if d(text="주문접수").exists:
            return False
    except Exception:
        pass

    try:
        if d(text="설치정보").exists:
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
    if not is_unexpected_digital_sales_home(d):
        return True

    msg = f"예상 밖 홈화면 이동 감지: {context} → 잠시 대기 후 재진행"
    print("⚠️", msg)
    notify(msg, level="warn")

    time.sleep(pause_sec)

    try:
        if d(text="모바일 주문").exists:
            click_text_center(d, "모바일 주문", 0.70, 0.98)
            time.sleep(1.0)
    except Exception:
        pass

    ok = ensure_mobile_order_home(d)
    if ok:
        print(f"✅ [recover] 모바일 주문 홈 복구 성공 ({context})")
    else:
        print(f"❌ [recover] 모바일 주문 홈 복구 실패 ({context})")
    return ok


def ensure_general_tab(d, force_click: bool = False) -> bool:
    if not d(text="주문현황").exists:
        return False

    def _is_general_ready():
        try:
            return d(text="고객/상호").exists
        except Exception:
            return False

    if _is_general_ready():
        return True

    if not force_click:
        return False

    try:
        h = d.window_size()[1]
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

                if top < int(h * 0.04) or top > int(h * 0.22):
                    continue

                cx = (left + right) // 2
                cy = (top + bottom) // 2
                d.click(cx, cy)
                time.sleep(0.45)

                if _is_general_ready():
                    return True
            except Exception:
                continue
    except Exception:
        pass

    return False


def enter_order_status(d) -> bool:
    if d(text="주문현황").exists:
        return True

    if not ensure_mobile_order_home(d):
        return False

    try:
        objs = d(text="일반주문")
        try:
            cnt = objs.count
        except Exception:
            cnt = 0

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

                if top < 300 or top > 1400:
                    continue

                if top < best_top:
                    best_top = top
                    target = (left, top, right, bottom)
            except Exception:
                continue

        if target is None:
            return False

        left, top, right, bottom = target
        cy = (top + bottom) // 2
        w, h = d.window_size()

        candidate_points = [
            (int(w * 0.84), cy),
            (int(w * 0.86), cy),
            (int(w * 0.88), cy),
        ]

        for x, y in candidate_points:
            d.click(x, y)
            time.sleep(1.2)

            if d(text="주문현황").exists:
                ensure_general_tab(d, force_click=True)
                time.sleep(0.4)
                return True

        return False
    except Exception:
        return False


def wait_until_order_status_ready(d, timeout_sec: float = 8.0) -> str:
    end_at = time.time() + timeout_sec
    stable_ok_count = 0

    while time.time() < end_at:
        try:
            if is_unexpected_digital_sales_home(d):
                return "home"

            if not d(text="주문현황").exists:
                stable_ok_count = 0
                time.sleep(0.25)
                continue

            edits = d(className="android.widget.EditText")
            cnt = edits.count

            if cnt >= 2:
                stable_ok_count += 1
                if stable_ok_count >= 3:
                    time.sleep(0.45)
                    return "ready"
                time.sleep(0.25)
                continue

            stable_ok_count = 0
        except Exception:
            stable_ok_count = 0

        time.sleep(0.25)

    if is_unexpected_digital_sales_home(d):
        return "home"

    return "timeout"


def find_search_edittext(d):
    try:
        w, h = d.window_size()
        best = None
        best_w = 0

        edits = d(className="android.widget.EditText")
        cnt = edits.count

        for i in range(cnt):
            try:
                e = edits[i]
                b = e.info.get("bounds", {})
                left = int(b.get("left", 0))
                top = int(b.get("top", 0))
                right = int(b.get("right", 0))
                bw = right - left

                if top < int(h * 0.10) or top > int(h * 0.42):
                    continue

                if bw < int(w * 0.30):
                    continue

                if bw > best_w:
                    best_w = bw
                    best = e
            except Exception:
                continue

        return best
    except Exception:
        return None


def trigger_search(d, edit_obj):
    try:
        w, h = d.window_size()
        b = edit_obj.info.get("bounds", {})
        right = int(b.get("right", 0))
        top = int(b.get("top", 0))
        bottom = int(b.get("bottom", 0))
        cy = (top + bottom) // 2

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


def get_status_badges_in_results(d, target_status: str):
    try:
        h = d.window_size()[1]
        found = []

        objs = d(text=target_status)
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
                    "status": target_status,
                    "bounds": (left, top, right, bottom),
                })
            except Exception:
                continue

        found.sort(key=lambda x: (x["bounds"][1], x["bounds"][0]))
        return found
    except Exception:
        return []


def open_detail_by_status_badge(d, badge_obj) -> bool:
    try:
        left, top, right, bottom = badge_obj.get("bounds", (0, 0, 0, 0))
        cx = (int(left) + int(right)) // 2
        cy = (int(top) + int(bottom)) // 2
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

    try:
        return d(text=name).exists and d(text=disp).exists
    except Exception:
        return False


def wait_install_info_ready(d, timeout_sec: float = 12.0) -> bool:
    end_at = time.time() + timeout_sec
    stable_ok_count = 0

    while time.time() < end_at:
        try:
            if is_unexpected_digital_sales_home(d):
                return False

            has_title = d(text="설치정보").exists or d(textContains="설치정보").exists
            has_phone = d(text="휴대폰번호").exists or d(textContains="휴대폰번호").exists
            has_address = d(text="주소").exists or d(textContains="주소").exists

            if has_title and has_phone and has_address:
                stable_ok_count += 1
                if stable_ok_count >= 2:
                    time.sleep(0.5)
                    return True
            else:
                stable_ok_count = 0
        except Exception:
            stable_ok_count = 0

        time.sleep(0.4)

    return False


def open_install_env_info(d) -> bool:
    try:
        if d(text="미입력").exists:
            ok = click_text_center(d, "미입력", 0.45, 0.85)
            time.sleep(1.0)
            return ok
    except Exception:
        pass

    label_obj = None
    try:
        if d(text="설치환경정보").exists:
            label_obj = d(text="설치환경정보")
        elif d(textContains="설치환경정보").exists:
            label_obj = d(textContains="설치환경정보")
    except Exception:
        label_obj = None

    if label_obj is not None:
        try:
            w, h = d.window_size()
            b = label_obj.info.get("bounds", {})
            bottom = int(b.get("bottom", 0))
            field_y = min(h - 10, bottom + max(50, int(h * 0.03)))
            for x in [int(w * 0.50), int(w * 0.72), int(w * 0.86)]:
                d.click(x, field_y)
                time.sleep(1.0)
                try:
                    if d(textContains="타사제품 반환여부").exists or d(text="입력완료").exists:
                        return True
                except Exception:
                    pass
        except Exception:
            pass

    return False


# =========================
# 실제 조회 플로우
# =========================
def process_one_job(d, job: dict):
    watermark8 = str(job.get("watermark8") or "").strip()
    message_ts = str(job.get("message_ts") or "").strip()
    channel_id = str(job.get("channel_id") or "").strip()
    thread_ts = str(job.get("thread_ts") or "").strip()

    def _fail(reason: str):
        db_update_job_status(message_ts, "오류", reason)
        slack_reply_in_thread(
            channel_id=channel_id,
            thread_ts=thread_ts,
            text=f"조회번호 {watermark8} / 오류\n{reason}",
        )
        notify_error(f"조회 실패 / 물마크 {watermark8} / {reason}")
        return False

    print(f"🚀 작업 시작: 물마크={watermark8}")
    db_update_job_status(message_ts, "처리중", "")

    if not restart_digital_sales_app(d):
        return _fail("디지털세일즈 앱 재시작 실패")

    if not ensure_mobile_order_home(d):
        return _fail("모바일 주문 홈 진입 실패")

    if is_unexpected_digital_sales_home(d):
        if not recover_from_unexpected_home(d, f"작업시작 직전 / 물마크 {watermark8}"):
            return _fail("예상 밖 홈화면 복구 실패")

    if not enter_order_status(d):
        return _fail("주문현황 진입 실패")

    ready_state = wait_until_order_status_ready(d, timeout_sec=8.0)
    if ready_state == "home":
        return _fail("주문현황 로딩 중 홈화면 이탈")
    if ready_state != "ready":
        return _fail("주문현황 준비 실패")

    if not ensure_general_tab(d, force_click=True):
        return _fail("일반 탭 복구 실패")

    edit = find_search_edittext(d)
    if edit is None:
        return _fail("검색 입력칸 찾기 실패")

    print(f"🔎 고객 검색: {SEARCH_NAME}")
    if not type_into_edittext(d, edit, SEARCH_NAME):
        return _fail(f"검색어 입력 실패: {SEARCH_NAME}")

    trigger_search(d, edit)
    time.sleep(1.3)

    if is_unexpected_digital_sales_home(d):
        return _fail("검색 후 홈화면 이탈")

    badges = get_status_badges_in_results(d, TARGET_STATUS_TEXT)
    if not badges:
        return _fail(f"{TARGET_STATUS_TEXT} 상태 배지 미검출")

    matched = False
    for idx, badge in enumerate(badges, start=1):
        print(f"✅ 상태 후보 확인 {idx}/{len(badges)}")

        if not open_detail_by_status_badge(d, badge):
            continue

        time.sleep(0.8)

        if match_detail_by_name_phone(d, DETAIL_EXPECT_NAME, DETAIL_EXPECT_PHONE11):
            matched = True
            print("✅ 상세 이름/번호 일치")
            break

        try:
            d.press("back")
            time.sleep(0.8)
        except Exception:
            pass

    if not matched:
        return _fail("설치입력 상세에서 이름/번호 일치 항목을 찾지 못함")

    try:
        if not d(text="주문 이어서 하기").exists:
            return _fail("주문 이어서 하기 버튼 미검출")
    except Exception:
        return _fail("주문 이어서 하기 버튼 확인 실패")

    if not click_text_center(d, "주문 이어서 하기", 0.20, 0.85):
        return _fail("주문 이어서 하기 클릭 실패")

    time.sleep(2.0)

    if is_unexpected_digital_sales_home(d):
        return _fail("주문 이어서 하기 후 홈화면 이탈")

    if not wait_install_info_ready(d, timeout_sec=12.0):
        return _fail("설치정보 화면 진입 실패")

    if not open_install_env_info(d):
        return _fail("설치환경정보 클릭 실패")

    db_update_job_status(message_ts, "완료", "")
    slack_reply_in_thread(
        channel_id=channel_id,
        thread_ts=thread_ts,
        text=f"조회번호 {watermark8} / 설치환경정보 진입 완료",
    )
    notify_success(f"조회 완료 / 물마크 {watermark8} / 설치환경정보 진입 완료")
    return True


# =========================
# 워커
# =========================
def preload_waiting_jobs():
    rows = db_list_waiting_jobs()
    for row in rows:
        job_q.put(dict(row))


def emulator_worker_loop():
    d = connect_emulator()
    print("✅ 에뮬레이터 연결 완료:", ADB_SERIAL)

    while True:
        try:
            if STOP_FLAG:
                time.sleep(1.0)
                continue

            try:
                job = job_q.get(timeout=1.0)
            except queue.Empty:
                continue

            try:
                process_one_job(d, job)
            except Exception as e:
                print("❌ 작업 처리 예외:", e)
                traceback.print_exc()

                message_ts = str(job.get("message_ts") or "").strip()
                channel_id = str(job.get("channel_id") or "").strip()
                thread_ts = str(job.get("thread_ts") or "").strip()
                watermark8 = str(job.get("watermark8") or "").strip()

                db_update_job_status(message_ts, "오류", str(e))
                slack_reply_in_thread(
                    channel_id=channel_id,
                    thread_ts=thread_ts,
                    text=f"조회번호 {watermark8} / 오류\n{e}",
                )
                notify_error(f"작업 예외 / 물마크 {watermark8} / {e}")

            finally:
                try:
                    job_q.task_done()
                except Exception:
                    pass

        except Exception as e:
            print("❌ 워커 루프 예외:", e)
            traceback.print_exc()
            notify_error(f"워커 루프 예외: {e}")
            time.sleep(1.0)


# =========================
# 실행
# =========================
def main():
    print("🚀 시작")
    init_db()
    preload_waiting_jobs()
    clear_stop()

    t_worker = threading.Thread(target=emulator_worker_loop, daemon=True)
    t_worker.start()

    t_slack = threading.Thread(target=slack_poll_loop, daemon=True)
    t_slack.start()

    while True:
        time.sleep(1.0)


if __name__ == "__main__":
    main()