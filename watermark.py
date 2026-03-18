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
import tkinter as tk
from tkinter import ttk

import uiautomator2 as u2
from slack_sdk import WebClient
from slack_sdk.errors import SlackApiError


BUILD_ID = "DIGITAL_SALES_SLACK_WATERMARK_2026-03-18_003"
print("✅ BUILD:", BUILD_ID)

# =========================
# 기본 설정
# =========================
ADB_SERIAL = "emulator-5554"
DIGITAL_SALES_APP_NAME = "디지털세일즈"
DIGITAL_SALES_PACKAGE = ""

SLACK_BOT_TOKEN = os.environ.get("SLACK_BOT_TOKEN", "").strip()
SLACK_CHANNEL_ID = os.environ.get("SLACK_CHANNEL_ID", "").strip()

ENABLE_TELEGRAM_NOTIFY = True
TELEGRAM_BOT_TOKEN = os.environ.get("TG_BOT_TOKEN", "").strip()
TELEGRAM_CHAT_ID_ME = os.environ.get("TG_CHAT_ID_ME", "").strip()

NOTIFY_PREFIX = "[디지털세일즈 Slack 조회]"

POLL_INTERVAL_SEC = 3.0
SEARCH_NAME = "정재경"
DETAIL_EXPECT_NAME = "정재경"
DETAIL_EXPECT_PHONE11 = "01043271050"
TARGET_STATUS_TEXT = "설치입력"

DB_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "digital_sales_slack_watermark.sqlite3"
)

STOP_FLAG = False
LAST_STOP_REASON = ""
ADB_DEVICE = None

job_q = queue.Queue()
db_lock = threading.RLock()

slack_client = WebClient(token=SLACK_BOT_TOKEN) if SLACK_BOT_TOKEN else None
SLACK_USER_NAME_CACHE = {}


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
                    slack_user_id TEXT DEFAULT '',
                    slack_user_name TEXT DEFAULT '',
                    watermark8 TEXT DEFAULT '',
                    status TEXT DEFAULT '',
                    last_error TEXT DEFAULT '',
                    replied_at REAL DEFAULT 0,
                    reply_done INTEGER DEFAULT 0,
                    created_at REAL DEFAULT 0,
                    updated_at REAL DEFAULT 0
                )
            """)

            existing_cols = set()
            cur = conn.execute("PRAGMA table_info(jobs)")
            for row in cur.fetchall():
                try:
                    existing_cols.add(str(row["name"]))
                except Exception:
                    try:
                        existing_cols.add(str(row[1]))
                    except Exception:
                        pass

            alter_columns = [
                ("slack_user_id", "TEXT DEFAULT ''"),
                ("slack_user_name", "TEXT DEFAULT ''"),
                ("replied_at", "REAL DEFAULT 0"),
                ("reply_done", "INTEGER DEFAULT 0"),
            ]

            for col_name, col_def in alter_columns:
                if col_name not in existing_cols:
                    conn.execute(f"ALTER TABLE jobs ADD COLUMN {col_name} {col_def}")

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


def db_create_job(
    message_ts: str,
    channel_id: str,
    thread_ts: str,
    slack_user_id: str,
    slack_user_name: str,
    watermark8: str
):
    now = time.time()
    with db_lock:
        conn = _db_conn()
        try:
            conn.execute("""
                INSERT OR IGNORE INTO jobs (
                    message_ts, channel_id, thread_ts,
                    slack_user_id, slack_user_name, watermark8,
                    status, last_error, replied_at, reply_done,
                    created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                str(message_ts or "").strip(),
                str(channel_id or "").strip(),
                str(thread_ts or "").strip(),
                str(slack_user_id or "").strip(),
                str(slack_user_name or "").strip(),
                str(watermark8 or "").strip(),
                "대기",
                "",
                0,
                0,
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
            cur = conn.execute(
                "SELECT * FROM jobs WHERE message_ts = ?",
                (str(message_ts or "").strip(),)
            )
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


def db_mark_job_replied(message_ts: str):
    now = time.time()
    with db_lock:
        conn = _db_conn()
        try:
            conn.execute("""
                UPDATE jobs
                SET replied_at = ?, reply_done = 1, updated_at = ?
                WHERE message_ts = ?
            """, (
                now,
                now,
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


def db_list_recent_jobs(limit: int = 300):
    with db_lock:
        conn = _db_conn()
        try:
            cur = conn.execute("""
                SELECT *
                FROM jobs
                ORDER BY created_at DESC, id DESC
                LIMIT ?
            """, (int(limit),))
            return [dict(r) for r in cur.fetchall()]
        finally:
            conn.close()


def db_delete_jobs_by_message_ts_list(message_ts_list):
    targets = []
    for x in list(message_ts_list or []):
        s = str(x or "").strip()
        if s:
            targets.append(s)

    if not targets:
        return 0

    with db_lock:
        conn = _db_conn()
        try:
            placeholders = ",".join(["?"] * len(targets))
            cur = conn.execute(
                f"DELETE FROM jobs WHERE message_ts IN ({placeholders})",
                tuple(targets)
            )
            conn.commit()
            return int(cur.rowcount or 0)
        finally:
            conn.close()


def fmt_ts(ts_value):
    try:
        ts = float(ts_value or 0)
        if ts <= 0:
            return ""
        return time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(ts))
    except Exception:
        return ""


# =========================
# Slack
# =========================
def extract_watermark8(text: str) -> str:
    s = str(text or "").strip()
    m = re.fullmatch(r"물마크\s?(\d{8})", s)
    if not m:
        return ""
    return m.group(1)


def slack_resolve_user_name(slack_user_id: str) -> str:
    slack_user_id = str(slack_user_id or "").strip()
    if not slack_user_id:
        return ""

    cached = SLACK_USER_NAME_CACHE.get(slack_user_id, "")
    if cached:
        return cached

    if not slack_client:
        return ""

    try:
        resp = slack_client.users_info(user=slack_user_id)
        user = resp.get("user", {}) or {}
        profile = user.get("profile", {}) or {}

        display_name = str(profile.get("display_name") or "").strip()
        real_name = str(profile.get("real_name") or "").strip()
        name = str(user.get("name") or "").strip()

        resolved = display_name or real_name or name or slack_user_id
        SLACK_USER_NAME_CACHE[slack_user_id] = resolved
        return resolved
    except SlackApiError as e:
        print("⚠️ Slack 사용자명 조회 실패:", e)
        return ""
    except Exception as e:
        print("⚠️ Slack 사용자명 조회 예외:", e)
        return ""


def slack_reply_in_thread(channel_id: str, thread_ts: str, text: str):
    if not slack_client:
        print("⚠️ Slack 토큰 미설정 → Slack 답장 생략:", text)
        return False

    try:
        slack_client.chat_postMessage(
            channel=channel_id,
            thread_ts=thread_ts,
            text=text,
        )
        return True
    except SlackApiError as e:
        print("⚠️ Slack 답장 실패:", e)
        return False


def slack_upload_file_and_reply(channel_id: str, thread_ts: str, text: str, file_path: str):
    if not slack_client:
        print("⚠️ Slack 토큰 미설정 → 파일 업로드 생략:", text)
        return False

    try:
        if file_path and os.path.isfile(file_path):
            slack_client.files_upload_v2(
                channel=channel_id,
                thread_ts=thread_ts,
                initial_comment=text,
                file=file_path,
                title=os.path.basename(file_path),
            )
        else:
            slack_client.chat_postMessage(
                channel=channel_id,
                thread_ts=thread_ts,
                text=text,
            )
        return True
    except SlackApiError as e:
        print("⚠️ Slack 파일 업로드/답장 실패:", e)
        try:
            fallback_text = text
            if file_path and os.path.isfile(file_path):
                fallback_text += f"\n(캡처 업로드 실패: {os.path.basename(file_path)})"
            slack_client.chat_postMessage(
                channel=channel_id,
                thread_ts=thread_ts,
                text=fallback_text,
            )
            return True
        except Exception as e2:
            print("⚠️ Slack fallback 텍스트 답장도 실패:", e2)
            return False


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
                slack_user_id = str(msg.get("user") or "").strip()
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
                    print(f"⚠️ Slack 양식불일치: ts={message_ts} / text={raw_text}")
                    slack_reply_in_thread(
                        channel_id=SLACK_CHANNEL_ID,
                        thread_ts=thread_ts,
                        text="조회요청 양식확인부탁드립니다.",
                    )
                    continue

                slack_user_name = slack_resolve_user_name(slack_user_id)

                print(
                    f"📥 Slack 조회요청 감지: ts={message_ts} / 물마크={watermark8} "
                    f"/ user={slack_user_id} / name={slack_user_name}"
                )

                db_create_job(
                    message_ts=message_ts,
                    channel_id=SLACK_CHANNEL_ID,
                    thread_ts=thread_ts,
                    slack_user_id=slack_user_id,
                    slack_user_name=slack_user_name,
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


def reset_digital_sales_to_idle(d) -> bool:
    try:
        if DIGITAL_SALES_PACKAGE:
            try:
                d.app_stop(DIGITAL_SALES_PACKAGE)
                time.sleep(1.2)
            except Exception:
                pass

            try:
                d.app_start(DIGITAL_SALES_PACKAGE)
                time.sleep(2.5)
            except Exception:
                return False
        else:
            try:
                d.press("home")
                time.sleep(0.8)
            except Exception:
                pass

            if not d(text=DIGITAL_SALES_APP_NAME).exists:
                return False

            if not click_text_center(d, DIGITAL_SALES_APP_NAME):
                return False

            time.sleep(2.5)

        return ensure_mobile_order_home(d)
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


def screenshots_dir():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    path = os.path.join(base_dir, "screenshots")
    os.makedirs(path, exist_ok=True)
    return path


def save_device_screenshot(d, prefix: str, watermark8: str) -> str:
    stamp = time.strftime("%Y%m%d_%H%M%S", time.localtime())
    filename = f"{prefix}_{watermark8}_{stamp}.png"
    local_path = os.path.join(screenshots_dir(), filename)
    remote_path = f"/sdcard/{filename}"

    def _is_valid_png(path: str) -> bool:
        try:
            if not os.path.isfile(path):
                return False
            if os.path.getsize(path) <= 0:
                return False
            with open(path, "rb") as f:
                header = f.read(8)
            return header.startswith(b"\x89PNG")
        except Exception:
            return False

    for attempt in range(1, 4):
        try:
            try:
                if os.path.isfile(local_path):
                    os.remove(local_path)
            except Exception:
                pass

            try:
                d.shell(["rm", "-f", remote_path])
            except Exception:
                pass

            shot_ok = False

            try:
                d.shell(["screencap", "-p", remote_path])
                shot_ok = True
            except Exception as e:
                print(f"⚠️ 캡처 {attempt}차 screencap 1단계 실패:", e)
                shot_ok = False

            if not shot_ok:
                try:
                    d.shell(["sh", "-c", f"screencap -p {remote_path}"])
                    shot_ok = True
                except Exception as e:
                    print(f"⚠️ 캡처 {attempt}차 screencap 2단계 실패:", e)
                    shot_ok = False

            if not shot_ok:
                time.sleep(0.6)
                continue

            time.sleep(0.8)

            pulled = False
            try:
                d.pull(remote_path, local_path)
                pulled = True
            except Exception as e:
                print(f"⚠️ 캡처 {attempt}차 pull 실패:", e)
                pulled = False

            try:
                d.shell(["rm", "-f", remote_path])
            except Exception:
                pass

            if pulled and _is_valid_png(local_path):
                print("📸 캡처 저장:", local_path)
                return local_path

            print(f"⚠️ 캡처 {attempt}차 실패: 로컬 PNG 유효성 불일치")
            time.sleep(0.6)

        except Exception as e:
            print(f"⚠️ 캡처 {attempt}차 전체 예외:", e)
            time.sleep(0.6)

    try:
        img = d.screenshot()
        if img is not None:
            img.save(local_path)
            if _is_valid_png(local_path):
                print("📸 캡처 저장(폴백):", local_path)
                return local_path
    except Exception as e:
        print("⚠️ 캡처 폴백 실패:", e)

    print("⚠️ 캡처 최종 실패")
    return ""


def wait_install_env_ready(d, timeout_sec: float = 8.0) -> bool:
    end_at = time.time() + timeout_sec
    stable_ok_count = 0

    while time.time() < end_at:
        try:
            if is_unexpected_digital_sales_home(d):
                return False

            has_return = (
                d(textContains="타사제품 반환여부").exists
                or d(text="미반환").exists
                or d(text="반환").exists
            )
            has_watermark = d(textContains="물마크 번호").exists
            has_done = d(text="입력완료").exists or d(textContains="입력완료").exists

            if has_return and has_watermark and has_done:
                stable_ok_count += 1
                if stable_ok_count >= 2:
                    return True
            else:
                stable_ok_count = 0
        except Exception:
            stable_ok_count = 0

        time.sleep(0.3)

    return False


def _read_edit_obj_text(obj) -> str:
    try:
        return (obj.get_text() or "").strip()
    except Exception:
        try:
            info = obj.info or {}
            return str(info.get("text") or "").strip()
        except Exception:
            return ""


def _find_first_label_obj(d, label_candidates):
    for lab in label_candidates:
        try:
            if d(text=lab).exists:
                return d(text=lab)
        except Exception:
            pass

    for lab in label_candidates:
        try:
            if d(textContains=lab).exists:
                return d(textContains=lab)
        except Exception:
            pass

    return None


def _find_edittext_near_label(d, label_candidates, y_tolerance: int = 110, y_below: int = 220):
    label_obj = _find_first_label_obj(d, label_candidates)

    try:
        edits = d(className="android.widget.EditText")
        cnt = edits.count
    except Exception:
        cnt = 0

    best = None
    best_score = None

    label_left = 0
    label_right = 0
    label_cy = 0
    label_bottom = 0

    if label_obj is not None:
        try:
            b = label_obj.info.get("bounds", {})
            label_left = int(b.get("left", 0))
            label_right = int(b.get("right", 0))
            label_top = int(b.get("top", 0))
            label_bottom = int(b.get("bottom", 0))
            label_cy = (label_top + label_bottom) // 2
        except Exception:
            label_obj = None

    for i in range(cnt):
        try:
            e = edits[i]
            b = e.info.get("bounds", {})
            left = int(b.get("left", 0))
            top = int(b.get("top", 0))
            right = int(b.get("right", 0))
            bottom = int(b.get("bottom", 0))
            cy = (top + bottom) // 2
            bw = right - left

            if bw < 160:
                continue

            score = None

            if label_obj is not None:
                if abs(cy - label_cy) <= y_tolerance and left >= max(0, label_left - 20):
                    score = abs(cy - label_cy) + abs(left - label_right)
                elif top >= label_bottom - 10 and top <= label_bottom + y_below:
                    score = 1000 + abs(top - label_bottom)
            else:
                score = top

            if score is not None and (best is None or score < best_score):
                best = e
                best_score = score
        except Exception:
            continue

    return best


def select_bottom_sheet_option(d, option_text: str, timeout_sec: float = 6.0) -> bool:
    end_at = time.time() + timeout_sec

    while time.time() < end_at:
        try:
            if d(text=option_text).exists:
                d(text=option_text).click()
                time.sleep(0.8)
                return True
        except Exception:
            pass

        try:
            if d(textContains=option_text).exists:
                d(textContains=option_text).click()
                time.sleep(0.8)
                return True
        except Exception:
            pass

        time.sleep(0.2)

    return False


def fill_watermark_number(d, watermark8: str) -> bool:
    edit = _find_edittext_near_label(d, ["물마크 번호", "물마크번호"], y_tolerance=100, y_below=200)
    if edit is None:
        return False

    ok = type_into_edittext(d, edit, watermark8)
    if not ok:
        return False

    time.sleep(0.6)

    current_text = _read_edit_obj_text(edit)
    current_digits = normalize_digits(current_text)
    return current_digits == watermark8


def click_text_if_exists(d, txt: str, y_min_ratio: float = 0.0, y_max_ratio: float = 1.0) -> bool:
    try:
        if d(text=txt).exists:
            return click_text_center(d, txt, y_min_ratio, y_max_ratio)
    except Exception:
        pass

    try:
        if d(textContains=txt).exists:
            d(textContains=txt).click()
            return True
    except Exception:
        pass

    return False


def choose_install_env_values(d, watermark8: str) -> bool:
    if not wait_install_env_ready(d, timeout_sec=8.0):
        return False

    if not click_text_if_exists(d, "제조사 선택", 0.10, 0.45):
        return False
    if not select_bottom_sheet_option(d, "SK"):
        return False

    if not click_text_if_exists(d, "년도", 0.10, 0.45):
        return False
    if not select_bottom_sheet_option(d, "2025년"):
        return False

    if not click_text_if_exists(d, "월", 0.10, 0.45):
        return False
    if not select_bottom_sheet_option(d, "1월"):
        return False

    if not click_text_if_exists(d, "제품형태 선택", 0.10, 0.55):
        return False
    if not select_bottom_sheet_option(d, "데스크탑"):
        return False

    if not click_text_if_exists(d, "제품종류 선택", 0.10, 0.60):
        return False
    if not select_bottom_sheet_option(d, "얼음정수기"):
        return False

    if not click_text_if_exists(d, "설치형태 선택", 0.10, 0.70):
        return False
    if not select_bottom_sheet_option(d, "직수형"):
        return False

    if not fill_watermark_number(d, watermark8):
        return False

    if not click_text_if_exists(d, "다중시설 선택", 0.40, 0.80):
        return False
    if not select_bottom_sheet_option(d, "대상 아님"):
        return False

    return True


def is_duplicate_watermark_popup_open(d) -> bool:
    popup_fragments = [
        "이미 접수 된 물마크 번호",
        "이미 접수된 물마크 번호",
        "본칙의심 모니터링 대상",
        "해당 번호로 주문 진행 시",
    ]

    for frag in popup_fragments:
        try:
            if d(textContains=frag).exists:
                return True
        except Exception:
            pass

    return False


def click_input_complete(d) -> bool:
    try:
        if d(text="입력완료").exists:
            d(text="입력완료").click()
            time.sleep(1.2)
            return True
    except Exception:
        pass

    try:
        if d(textContains="입력완료").exists:
            d(textContains="입력완료").click()
            time.sleep(1.2)
            return True
    except Exception:
        pass

    return click_text_center(d, "입력완료", 0.85, 0.99)


# =========================
# 실제 조회 플로우
# =========================
def process_one_job(d, job: dict):
    watermark8 = str(job.get("watermark8") or "").strip()
    message_ts = str(job.get("message_ts") or "").strip()
    channel_id = str(job.get("channel_id") or "").strip()
    thread_ts = str(job.get("thread_ts") or "").strip()
    slack_user_id = str(job.get("slack_user_id") or "").strip()

    def _mention_text():
        if slack_user_id:
            return f"<@{slack_user_id}>"
        return ""

    def _final_reset():
        try:
            ok = reset_digital_sales_to_idle(d)
            if ok:
                print("✅ 작업 종료 후 디지털세일즈 대기상태 복귀 완료")
            else:
                print("⚠️ 작업 종료 후 디지털세일즈 대기상태 복귀 실패")
        except Exception as e:
            print("⚠️ 작업 종료 후 디지털세일즈 리셋 예외:", e)

    def _send_final_error(reason: str):
        db_update_job_status(message_ts, "오류", reason)
        reply_text = f"조회번호 {watermark8} / 오류\n{reason}"
        mention = _mention_text()
        if mention:
            reply_text += f"\n{mention}"

        ok_sent = slack_reply_in_thread(
            channel_id=channel_id,
            thread_ts=thread_ts,
            text=reply_text,
        )
        if ok_sent:
            db_mark_job_replied(message_ts)

        notify_error(f"조회 실패 / 물마크 {watermark8} / {reason}")
        _final_reset()
        return False

    print(f"🚀 작업 시작: 물마크={watermark8}")
    db_update_job_status(message_ts, "처리중", "")

    last_reason = ""

    for attempt in range(1, 4):
        print(f"🔄 작업 재시도 {attempt}/3 / 물마크={watermark8}")

        if not restart_digital_sales_app(d):
            last_reason = "디지털세일즈 앱 재시작 실패"
            time.sleep(1.0)
            continue

        if not ensure_mobile_order_home(d):
            last_reason = "모바일 주문 홈 진입 실패"
            time.sleep(1.0)
            continue

        if is_unexpected_digital_sales_home(d):
            if not recover_from_unexpected_home(d, f"작업시작 직전 / 물마크 {watermark8}"):
                last_reason = "예상 밖 홈화면 복구 실패"
                time.sleep(1.0)
                continue

        if not enter_order_status(d):
            last_reason = "주문현황 진입 실패"
            time.sleep(1.0)
            continue

        ready_state = wait_until_order_status_ready(d, timeout_sec=12.0)
        if ready_state == "home":
            last_reason = "주문현황 로딩 중 홈화면 이탈"
            time.sleep(1.0)
            continue
        if ready_state != "ready":
            last_reason = "주문현황 준비 실패"
            time.sleep(1.0)
            continue

        if not ensure_general_tab(d, force_click=True):
            last_reason = "일반 탭 복구 실패"
            time.sleep(1.0)
            continue

        edit = find_search_edittext(d)
        if edit is None:
            last_reason = "검색 입력칸 찾기 실패"
            time.sleep(1.0)
            continue

        print(f"🔎 고객 검색: {SEARCH_NAME}")
        if not type_into_edittext(d, edit, SEARCH_NAME):
            last_reason = f"검색어 입력 실패: {SEARCH_NAME}"
            time.sleep(1.0)
            continue

        trigger_search(d, edit)
        time.sleep(1.8)

        if is_unexpected_digital_sales_home(d):
            last_reason = "검색 후 홈화면 이탈"
            time.sleep(1.0)
            continue

        badges = get_status_badges_in_results(d, TARGET_STATUS_TEXT)
        if not badges:
            last_reason = f"{TARGET_STATUS_TEXT} 상태 배지 미검출"
            time.sleep(1.0)
            continue

        matched = False
        for idx, badge in enumerate(badges, start=1):
            print(f"✅ 상태 후보 확인 {idx}/{len(badges)}")

            if not open_detail_by_status_badge(d, badge):
                continue

            time.sleep(1.0)

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
            last_reason = "설치입력 상세에서 이름/번호 일치 항목을 찾지 못함"
            time.sleep(1.0)
            continue

        try:
            if not d(text="주문 이어서 하기").exists:
                last_reason = "주문 이어서 하기 버튼 미검출"
                time.sleep(1.0)
                continue
        except Exception:
            last_reason = "주문 이어서 하기 버튼 확인 실패"
            time.sleep(1.0)
            continue

        if not click_text_center(d, "주문 이어서 하기", 0.20, 0.85):
            last_reason = "주문 이어서 하기 클릭 실패"
            time.sleep(1.0)
            continue

        time.sleep(2.0)

        if is_unexpected_digital_sales_home(d):
            last_reason = "주문 이어서 하기 후 홈화면 이탈"
            time.sleep(1.0)
            continue

        if not wait_install_info_ready(d, timeout_sec=12.0):
            last_reason = "설치정보 화면 진입 실패"
            time.sleep(1.0)
            continue

        if not open_install_env_info(d):
            last_reason = "설치환경정보 클릭 실패"
            time.sleep(1.0)
            continue

        if not wait_install_env_ready(d, timeout_sec=8.0):
            last_reason = "설치환경정보 화면 준비 실패"
            time.sleep(1.0)
            continue

        if not choose_install_env_values(d, watermark8):
            last_reason = "설치환경정보 값 선택/입력 실패"
            time.sleep(1.0)
            continue

        capture_path = save_device_screenshot(d, "watermark_filled", watermark8)

        if not click_input_complete(d):
            last_reason = "입력완료 클릭 실패"
            time.sleep(1.0)
            continue

        time.sleep(1.5)

        if is_duplicate_watermark_popup_open(d):
            reply_text = "🚨물마크 중복입니다🚨"
            mention = _mention_text()
            if mention:
                reply_text += f"\n{mention}"

            ok_sent = slack_upload_file_and_reply(
                channel_id=channel_id,
                thread_ts=thread_ts,
                text=reply_text,
                file_path=capture_path,
            )
            if ok_sent:
                db_mark_job_replied(message_ts)

            db_update_job_status(message_ts, "중복팝업", "🚨물마크 중복 팝업🚨")
            notify_progress(f"🚨물마크 중복 팝업🚨 / 물마크 {watermark8}")
            _final_reset()
            return True

        reply_text = "💠물마크 사용가능💠"
        mention = _mention_text()
        if mention:
            reply_text += f"\n{mention}"

        ok_sent = slack_upload_file_and_reply(
            channel_id=channel_id,
            thread_ts=thread_ts,
            text=reply_text,
            file_path=capture_path,
        )
        if ok_sent:
            db_mark_job_replied(message_ts)

        db_update_job_status(message_ts, "완료", "")
        notify_success(f"물마크 {watermark8} / 💠물마크 사용가능💠")
        _final_reset()
        return True

    return _send_final_error(last_reason or "재시도 후에도 처리 실패")


# =========================
# 로그 창
# =========================
def start_log_window_thread():
    def _run():
        root = tk.Tk()
        root.title("디지털세일즈 물마크 진행현황")
        root.geometry("1520x760")

        checked_ids = set()

        top_wrap = ttk.Frame(root, padding=8)
        top_wrap.pack(fill="both", expand=True)

        ctrl = ttk.Frame(top_wrap)
        ctrl.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        top = ttk.Frame(top_wrap)
        top.grid(row=1, column=0, sticky="nsew")

        top_wrap.rowconfigure(1, weight=1)
        top_wrap.columnconfigure(0, weight=1)

        info_var = tk.StringVar(value="체크 후 [선택삭제]를 누르면 기록이 삭제됩니다.")

        ttk.Label(ctrl, textvariable=info_var).pack(side="left", padx=(0, 10))

        def _delete_checked():
            targets = list(checked_ids)
            if not targets:
                info_var.set("삭제할 항목이 없습니다.")
                return

            deleted = db_delete_jobs_by_message_ts_list(targets)

            for iid in list(targets):
                checked_ids.discard(iid)

            info_var.set(f"{deleted}건 삭제 완료")
            _refresh_once()

        ttk.Button(ctrl, text="선택삭제", command=_delete_checked).pack(side="right")

        cols = (
            "delete_mark",
            "requester",
            "processed",
            "duplicated",
            "watermark8",
            "asked_at",
            "replied_at",
            "reply_done",
            "status",
            "last_error",
        )

        tree = ttk.Treeview(top, columns=cols, show="headings", height=25)

        tree.heading("delete_mark", text="삭제")
        tree.heading("requester", text="요청자")
        tree.heading("processed", text="처리완료여부")
        tree.heading("duplicated", text="중복여부")
        tree.heading("watermark8", text="물마크번호")
        tree.heading("asked_at", text="문의시각")
        tree.heading("replied_at", text="답변시각")
        tree.heading("reply_done", text="답변완료여부")
        tree.heading("status", text="현재상태")
        tree.heading("last_error", text="오류/비고")

        tree.column("delete_mark", width=60, anchor="center")
        tree.column("requester", width=180, anchor="center")
        tree.column("processed", width=100, anchor="center")
        tree.column("duplicated", width=80, anchor="center")
        tree.column("watermark8", width=110, anchor="center")
        tree.column("asked_at", width=150, anchor="center")
        tree.column("replied_at", width=150, anchor="center")
        tree.column("reply_done", width=100, anchor="center")
        tree.column("status", width=100, anchor="center")
        tree.column("last_error", width=420, anchor="w")

        vsb = ttk.Scrollbar(top, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(top, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        top.rowconfigure(0, weight=1)
        top.columnconfigure(0, weight=1)

        def _toggle_check(iid: str):
            if not iid:
                return

            if iid in checked_ids:
                checked_ids.discard(iid)
            else:
                checked_ids.add(iid)

            _refresh_once()

        def _on_tree_click(event):
            region = tree.identify("region", event.x, event.y)
            if region != "cell":
                return

            col = tree.identify_column(event.x)
            row_id = tree.identify_row(event.y)

            if not row_id:
                return

            if col == "#1":
                _toggle_check(row_id)

        tree.bind("<Button-1>", _on_tree_click)

        def _refresh_once():
            rows = db_list_recent_jobs(limit=500)
            existing_ids = set(tree.get_children())

            db_ids = set()

            for row in rows:
                iid = str(row.get("message_ts") or row.get("id") or "")
                if not iid:
                    continue

                db_ids.add(iid)

                requester = str(row.get("slack_user_name") or "").strip()
                if not requester:
                    requester = str(row.get("slack_user_id") or "").strip()
                    if requester:
                        requester = f"<@{requester}>"

                status_value = str(row.get("status") or "").strip()

                values = (
                    "☑" if iid in checked_ids else "☐",
                    requester,
                    "Y" if status_value in ["완료", "중복팝업", "오류"] else "N",
                    "Y" if status_value == "중복팝업" else "N",
                    row.get("watermark8") or "",
                    fmt_ts(row.get("created_at")),
                    fmt_ts(row.get("replied_at")),
                    "Y" if int(row.get("reply_done") or 0) == 1 else "N",
                    status_value,
                    row.get("last_error") or "",
                )

                if iid in existing_ids:
                    tree.item(iid, values=values)
                    existing_ids.discard(iid)
                else:
                    tree.insert("", "end", iid=iid, values=values)

            for iid in existing_ids:
                tree.delete(iid)

            for iid in list(checked_ids):
                if iid not in db_ids:
                    checked_ids.discard(iid)

        def _refresh_loop():
            _refresh_once()
            root.after(1500, _refresh_loop)

        _refresh_loop()
        root.mainloop()

    t = threading.Thread(target=_run, daemon=True)
    t.start()

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
                ok_sent = slack_reply_in_thread(
                    channel_id=channel_id,
                    thread_ts=thread_ts,
                    text=f"조회번호 {watermark8} / 오류\n{e}",
                )
                if ok_sent:
                    db_mark_job_replied(message_ts)

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
    start_log_window_thread()

    t_worker = threading.Thread(target=emulator_worker_loop, daemon=True)
    t_worker.start()

    t_slack = threading.Thread(target=slack_poll_loop, daemon=True)
    t_slack.start()

    while True:
        time.sleep(1.0)


if __name__ == "__main__":
    main()