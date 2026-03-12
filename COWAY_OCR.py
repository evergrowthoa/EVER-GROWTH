from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import time
import re
import os
import base64
import shutil
import subprocess
import sqlite3
import threading
import queue
import traceback
import tkinter as tk
from tkinter import ttk

import uiautomator2 as u2

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

STATUS_NEW = "신규"
STATUS_AUTH_SENT = "인증발송"
STATUS_DONE = "완료"
STATUS_CANCELLED = "취소"
STATUS_HOLD = "보류"
STATUS_DELETED = "삭제"

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "coway_jobs.sqlite3")

chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_experimental_option("detach", True)

driver = webdriver.Chrome(options=chrome_options)
wait = WebDriverWait(driver, 20)

auth_q = queue.Queue()
auth_sent_jobs = {}
jobs_lock = threading.Lock()
db_lock = threading.RLock()
sign_started = set()
queued_new_phones = set()

SIGN_IN_PROGRESS = False
STOP_FLAG = False
LAST_STOP_REASON = ""

def notify(msg: str):
    print("🔔", msg)

def set_stop(reason: str):
    global STOP_FLAG, LAST_STOP_REASON
    STOP_FLAG = True
    LAST_STOP_REASON = str(reason or "").strip()
    print("🛑 자동화 중단:", reason)
    notify(reason)

def clear_stop():
    global STOP_FLAG, LAST_STOP_REASON
    STOP_FLAG = False
    LAST_STOP_REASON = ""
    print("✅ 자동화 정지 해제")

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

def _db_conn():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn

def _fmt_ts(ts_value):
    try:
        ts = float(ts_value or 0)
        if ts <= 0:
            return ""
        return time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(ts))
    except Exception:
        return ""

def init_db():
    with db_lock:
        conn = _db_conn()
        try:
            conn.execute("""
                CREATE TABLE IF NOT EXISTS jobs (
                    phone11 TEXT PRIMARY KEY,
                    name TEXT DEFAULT '',
                    birth TEXT DEFAULT '',
                    account TEXT DEFAULT '',
                    zipcode TEXT DEFAULT '',
                    product_name TEXT DEFAULT '',
                    model_name TEXT DEFAULT '',
                    color_raw TEXT DEFAULT '',
                    address TEXT DEFAULT '',
                    address_basic TEXT DEFAULT '',
                    address_detail TEXT DEFAULT '',
                    manage_raw TEXT DEFAULT '',
                    contract_raw TEXT DEFAULT '',
                    discount_raw TEXT DEFAULT '',
                    amount_raw TEXT DEFAULT '',
                    pickup_request_raw TEXT DEFAULT '',
                    special_note_raw TEXT DEFAULT '',
                    third_brand TEXT DEFAULT '',
                    third_install_mfg TEXT DEFAULT '',
                    third_product_type TEXT DEFAULT '',
                    third_product_kind TEXT DEFAULT '',
                    third_install_shape TEXT DEFAULT '',
                    third_watermark_no TEXT DEFAULT '',
                    status TEXT DEFAULT '',
                    auth_sent_at REAL DEFAULT 0,
                    next_check_at REAL DEFAULT 0,
                    last_check_at REAL DEFAULT 0,
                    last_status TEXT DEFAULT '',
                    last_error TEXT DEFAULT '',
                    note TEXT DEFAULT '',
                    created_at REAL DEFAULT 0,
                    updated_at REAL DEFAULT 0
                )
            """)

            alter_columns = [
                ("pickup_request_raw", "TEXT DEFAULT ''"),
                ("special_note_raw", "TEXT DEFAULT ''"),
                ("third_brand", "TEXT DEFAULT ''"),
                ("third_install_mfg", "TEXT DEFAULT ''"),
                ("third_product_type", "TEXT DEFAULT ''"),
                ("third_product_kind", "TEXT DEFAULT ''"),
                ("third_install_shape", "TEXT DEFAULT ''"),
                ("third_watermark_no", "TEXT DEFAULT ''"),
            ]

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

            for col_name, col_def in alter_columns:
                if col_name not in existing_cols:
                    conn.execute(f"ALTER TABLE jobs ADD COLUMN {col_name} {col_def}")

            conn.execute("CREATE INDEX IF NOT EXISTS idx_jobs_status ON jobs(status)")
            conn.execute("CREATE INDEX IF NOT EXISTS idx_jobs_updated_at ON jobs(updated_at DESC)")
            conn.commit()
        finally:
            conn.close()

def db_get_job(phone11: str):
    if not phone11:
        return None
    with db_lock:
        conn = _db_conn()
        try:
            cur = conn.execute("SELECT * FROM jobs WHERE phone11 = ?", (phone11,))
            row = cur.fetchone()
            return dict(row) if row else None
        finally:
            conn.close()

def db_list_all_jobs():
    with db_lock:
        conn = _db_conn()
        try:
            cur = conn.execute("""
                SELECT *
                FROM jobs
                ORDER BY updated_at DESC, created_at DESC, phone11 DESC
            """)
            return [dict(r) for r in cur.fetchall()]
        finally:
            conn.close()

def db_upsert_job_from_modal(data: dict):
    phone11 = str(data.get("phone11") or "").strip()
    if not phone11:
        return None

    now = time.time()
    existing = db_get_job(phone11)

    payload = {
        "name": str(data.get("name") or "").strip(),
        "birth": str(data.get("birth") or "").strip(),
        "account": str(data.get("account") or "").strip(),
        "zipcode": str(data.get("zipcode") or "").strip(),
        "product_name": str(data.get("product_name") or "").strip(),
        "model_name": str(data.get("model_name") or "").strip(),
        "color_raw": str(data.get("color_raw") or "").strip(),
        "address": str(data.get("address") or "").strip(),
        "address_basic": str(data.get("address_basic") or "").strip(),
        "address_detail": str(data.get("address_detail") or "").strip(),
        "manage_raw": str(data.get("manage_raw") or "").strip(),
        "contract_raw": str(data.get("contract_raw") or "").strip(),
        "discount_raw": str(data.get("discount_raw") or "").strip(),
        "amount_raw": str(data.get("amount_raw") or "").strip(),
        "pickup_request_raw": str(data.get("pickup_request_raw") or "").strip(),
        "special_note_raw": str(data.get("special_note_raw") or "").strip(),
        "third_brand": str(data.get("third_brand") or "").strip(),
        "third_install_mfg": str(data.get("third_install_mfg") or "").strip(),
        "third_product_type": str(data.get("third_product_type") or "").strip(),
        "third_product_kind": str(data.get("third_product_kind") or "").strip(),
        "third_install_shape": str(data.get("third_install_shape") or "").strip(),
        "third_watermark_no": str(data.get("third_watermark_no") or "").strip(),
    }

    with db_lock:
        conn = _db_conn()
        try:
            if existing:
                existing_status = str(existing.get("status") or "").strip()

                if existing_status == STATUS_DELETED:
                    conn.execute("""
                        UPDATE jobs
                        SET name = ?,
                            birth = ?,
                            account = ?,
                            zipcode = ?,
                            product_name = ?,
                            model_name = ?,
                            color_raw = ?,
                            address = ?,
                            address_basic = ?,
                            address_detail = ?,
                            manage_raw = ?,
                            contract_raw = ?,
                            discount_raw = ?,
                            amount_raw = ?,
                            pickup_request_raw = ?,
                            special_note_raw = ?,
                            third_brand = ?,
                            third_install_mfg = ?,
                            third_product_type = ?,
                            third_product_kind = ?,
                            third_install_shape = ?,
                            third_watermark_no = ?,
                            status = ?,
                            auth_sent_at = 0,
                            next_check_at = 0,
                            last_check_at = 0,
                            last_status = '',
                            last_error = '',
                            updated_at = ?
                        WHERE phone11 = ?
                    """, (
                        payload["name"],
                        payload["birth"],
                        payload["account"],
                        payload["zipcode"],
                        payload["product_name"],
                        payload["model_name"],
                        payload["color_raw"],
                        payload["address"],
                        payload["address_basic"],
                        payload["address_detail"],
                        payload["manage_raw"],
                        payload["contract_raw"],
                        payload["discount_raw"],
                        payload["amount_raw"],
                        payload["pickup_request_raw"],
                        payload["special_note_raw"],
                        payload["third_brand"],
                        payload["third_install_mfg"],
                        payload["third_product_type"],
                        payload["third_product_kind"],
                        payload["third_install_shape"],
                        payload["third_watermark_no"],
                        STATUS_NEW,
                        now,
                        phone11,
                    ))
                else:
                    conn.execute("""
                        UPDATE jobs
                        SET name = ?,
                            birth = ?,
                            account = ?,
                            zipcode = ?,
                            product_name = ?,
                            model_name = ?,
                            color_raw = ?,
                            address = ?,
                            address_basic = ?,
                            address_detail = ?,
                            manage_raw = ?,
                            contract_raw = ?,
                            discount_raw = ?,
                            amount_raw = ?,
                            pickup_request_raw = ?,
                            special_note_raw = ?,
                            third_brand = ?,
                            third_install_mfg = ?,
                            third_product_type = ?,
                            third_product_kind = ?,
                            third_install_shape = ?,
                            third_watermark_no = ?,
                            updated_at = ?
                        WHERE phone11 = ?
                    """, (
                        payload["name"],
                        payload["birth"],
                        payload["account"],
                        payload["zipcode"],
                        payload["product_name"],
                        payload["model_name"],
                        payload["color_raw"],
                        payload["address"],
                        payload["address_basic"],
                        payload["address_detail"],
                        payload["manage_raw"],
                        payload["contract_raw"],
                        payload["discount_raw"],
                        payload["amount_raw"],
                        payload["pickup_request_raw"],
                        payload["special_note_raw"],
                        payload["third_brand"],
                        payload["third_install_mfg"],
                        payload["third_product_type"],
                        payload["third_product_kind"],
                        payload["third_install_shape"],
                        payload["third_watermark_no"],
                        now,
                        phone11,
                    ))
            else:
                conn.execute("""
                    INSERT INTO jobs (
                        phone11, name, birth, account, zipcode,
                        product_name, model_name, color_raw,
                        address, address_basic, address_detail,
                        manage_raw, contract_raw, discount_raw, amount_raw,
                        pickup_request_raw, special_note_raw,
                        third_brand, third_install_mfg, third_product_type,
                        third_product_kind, third_install_shape, third_watermark_no,
                        status, auth_sent_at, next_check_at, last_check_at,
                        last_status, last_error, note, created_at, updated_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    phone11,
                    payload["name"],
                    payload["birth"],
                    payload["account"],
                    payload["zipcode"],
                    payload["product_name"],
                    payload["model_name"],
                    payload["color_raw"],
                    payload["address"],
                    payload["address_basic"],
                    payload["address_detail"],
                    payload["manage_raw"],
                    payload["contract_raw"],
                    payload["discount_raw"],
                    payload["amount_raw"],
                    payload["pickup_request_raw"],
                    payload["special_note_raw"],
                    payload["third_brand"],
                    payload["third_install_mfg"],
                    payload["third_product_type"],
                    payload["third_product_kind"],
                    payload["third_install_shape"],
                    payload["third_watermark_no"],
                    STATUS_NEW,
                    0,
                    0,
                    0,
                    "",
                    "",
                    "",
                    now,
                    now,
                ))
            conn.commit()
        finally:
            conn.close()

    return db_get_job(phone11)

def db_mark_auth_sent(job: dict):
    phone11 = str(job.get("phone11") or "").strip()
    if not phone11:
        return None

    now = time.time()
    with db_lock:
        conn = _db_conn()
        try:
            conn.execute("""
                UPDATE jobs
                SET status = ?,
                    auth_sent_at = ?,
                    next_check_at = ?,
                    last_check_at = 0,
                    last_status = '',
                    last_error = '',
                    updated_at = ?
                WHERE phone11 = ?
            """, (
                STATUS_AUTH_SENT,
                now,
                now + 60,
                now,
                phone11,
            ))
            conn.commit()
        finally:
            conn.close()

    return db_get_job(phone11)

def db_update_check_state(phone11: str, last_status: str, next_check_at: float, last_check_at: float):
    if not phone11:
        return
    with db_lock:
        conn = _db_conn()
        try:
            conn.execute("""
                UPDATE jobs
                SET last_status = ?,
                    next_check_at = ?,
                    last_check_at = ?,
                    updated_at = ?
                WHERE phone11 = ?
            """, (
                str(last_status or "").strip(),
                float(next_check_at or 0),
                float(last_check_at or 0),
                time.time(),
                phone11,
            ))
            conn.commit()
        finally:
            conn.close()

def db_mark_done(phone11: str):
    if not phone11:
        return
    with db_lock:
        conn = _db_conn()
        try:
            conn.execute("""
                UPDATE jobs
                SET status = ?,
                    next_check_at = 0,
                    last_error = '',
                    updated_at = ?
                WHERE phone11 = ?
            """, (
                STATUS_DONE,
                time.time(),
                phone11,
            ))
            conn.commit()
        finally:
            conn.close()

def db_mark_hold(phone11: str, reason: str):
    if not phone11:
        return
    with db_lock:
        conn = _db_conn()
        try:
            conn.execute("""
                UPDATE jobs
                SET status = ?,
                    last_error = ?,
                    updated_at = ?
                WHERE phone11 = ?
            """, (
                STATUS_HOLD,
                str(reason or "").strip(),
                time.time(),
                phone11,
            ))
            conn.commit()
        finally:
            conn.close()

def db_set_note(phone11: str, note: str):
    if not phone11:
        return
    with db_lock:
        conn = _db_conn()
        try:
            conn.execute("""
                UPDATE jobs
                SET note = ?,
                    updated_at = ?
                WHERE phone11 = ?
            """, (
                str(note or "").strip(),
                time.time(),
                phone11,
            ))
            conn.commit()
        finally:
            conn.close()

def db_apply_manual_status(phone11: str, status_value: str):
    if not phone11:
        return

    now = time.time()
    status_value = str(status_value or "").strip()
    if status_value not in [STATUS_NEW, STATUS_AUTH_SENT, STATUS_DONE, STATUS_CANCELLED, STATUS_HOLD, STATUS_DELETED]:
        return

    row = db_get_job(phone11)
    if row is None:
        return

    if status_value == STATUS_DELETED:
        db_soft_delete_job(phone11)
        return

    next_check_at = 0
    if status_value == STATUS_AUTH_SENT:
        next_check_at = now
    elif status_value == STATUS_NEW:
        next_check_at = 0

    with db_lock:
        conn = _db_conn()
        try:
            conn.execute("""
                UPDATE jobs
                SET status = ?,
                    next_check_at = ?,
                    last_error = CASE WHEN ? IN (?, ?) THEN '' ELSE last_error END,
                    updated_at = ?
                WHERE phone11 = ?
            """, (
                status_value,
                next_check_at,
                status_value,
                STATUS_NEW,
                STATUS_AUTH_SENT,
                now,
                phone11,
            ))
            conn.commit()
        finally:
            conn.close()

def db_soft_delete_job(phone11: str):
    if not phone11:
        return

    row = db_get_job(phone11)
    if row is None:
        return

    now = time.time()
    stamp = _fmt_ts(now)
    prev_status = str(row.get("status") or "").strip()
    prev_last_status = str(row.get("last_status") or "").strip()
    old_note = str(row.get("note") or "").strip()

    delete_line = f"[삭제이력 {stamp}] 이전상태={prev_status} 마지막상태={prev_last_status}"
    new_note = delete_line if not old_note else (old_note + "\n" + delete_line)

    with db_lock:
        conn = _db_conn()
        try:
            conn.execute("""
                UPDATE jobs
                SET status = ?,
                    auth_sent_at = 0,
                    next_check_at = 0,
                    last_check_at = 0,
                    last_status = '',
                    last_error = '',
                    note = ?,
                    updated_at = ?
                WHERE phone11 = ?
            """, (
                STATUS_DELETED,
                new_note,
                now,
                phone11,
            ))
            conn.commit()
        finally:
            conn.close()

    try:
        queued_new_phones.discard(phone11)
    except Exception:
        pass

    try:
        processed_phones.discard(phone11)
    except Exception:
        pass

    try:
        sign_started.discard(phone11)
    except Exception:
        pass

    with jobs_lock:
        auth_sent_jobs.pop(phone11, None)

def db_resume_job(phone11: str):
    row = db_get_job(phone11)
    if row is None:
        return
    auth_sent_at = float(row.get("auth_sent_at") or 0)
    if auth_sent_at > 0:
        db_apply_manual_status(phone11, STATUS_AUTH_SENT)
        db_force_next_check(phone11)
    else:
        db_apply_manual_status(phone11, STATUS_NEW)

def db_force_next_check(phone11: str):
    if not phone11:
        return
    with db_lock:
        conn = _db_conn()
        try:
            conn.execute("""
                UPDATE jobs
                SET status = CASE WHEN auth_sent_at > 0 THEN ? ELSE status END,
                    next_check_at = ?,
                    updated_at = ?
                WHERE phone11 = ?
            """, (
                STATUS_AUTH_SENT,
                time.time(),
                time.time(),
                phone11,
            ))
            conn.commit()
        finally:
            conn.close()

def sync_runtime_state_from_db():
    rows = db_list_all_jobs()

    active_auth = {}
    active_new = set()
    active_done = set()

    for row in rows:
        phone11 = str(row.get("phone11") or "").strip()
        status = str(row.get("status") or "").strip()

        if not phone11:
            continue

        if status == STATUS_NEW:
            active_new.add(phone11)
            if phone11 not in queued_new_phones:
                auth_q.put(dict(row))
                queued_new_phones.add(phone11)

        elif status == STATUS_AUTH_SENT:
            active_auth[phone11] = dict(row)

        elif status == STATUS_DONE:
            active_done.add(phone11)

    for p in list(queued_new_phones):
        if p not in active_new:
            queued_new_phones.discard(p)

    with jobs_lock:
        for p in list(auth_sent_jobs.keys()):
            if p not in active_auth:
                auth_sent_jobs.pop(p, None)
        for p, row in active_auth.items():
            auth_sent_jobs[p] = row

    for p in list(sign_started):
        if p not in active_done:
            sign_started.discard(p)
    for p in active_done:
        sign_started.add(p)

def start_record_window_thread():
    def _run():
        root = tk.Tk()
        root.title("COWAY 진행기록")
        root.geometry("1280x620")

        top = ttk.Frame(root, padding=8)
        top.pack(fill="both", expand=True)

        cols = ("name", "phone11", "status", "last_status", "last_error", "note", "updated_at")
        tree = ttk.Treeview(top, columns=cols, show="headings", height=20)

        tree.heading("name", text="고객명")
        tree.heading("phone11", text="연락처")
        tree.heading("status", text="현재 상태")
        tree.heading("last_status", text="마지막 상태")
        tree.heading("last_error", text="마지막 오류")
        tree.heading("note", text="메모")
        tree.heading("updated_at", text="수정시각")

        tree.column("name", width=110, anchor="center")
        tree.column("phone11", width=120, anchor="center")
        tree.column("status", width=90, anchor="center")
        tree.column("last_status", width=100, anchor="center")
        tree.column("last_error", width=280, anchor="w")
        tree.column("note", width=220, anchor="w")
        tree.column("updated_at", width=150, anchor="center")

        vsb = ttk.Scrollbar(top, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)

        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")

        top.rowconfigure(0, weight=1)
        top.columnconfigure(0, weight=1)

        ctrl = ttk.Frame(root, padding=8)
        ctrl.pack(fill="x")

        selected_phone = {"value": ""}

        ttk.Label(ctrl, text="상태").grid(row=0, column=0, padx=4, pady=4, sticky="w")
        status_var = tk.StringVar(value=STATUS_NEW)
        status_combo = ttk.Combobox(
            ctrl,
            textvariable=status_var,
            state="readonly",
            values=[STATUS_NEW, STATUS_AUTH_SENT, STATUS_DONE, STATUS_CANCELLED, STATUS_HOLD, STATUS_DELETED],
            width=12
        )
        status_combo.grid(row=0, column=1, padx=4, pady=4, sticky="w")

        ttk.Label(ctrl, text="메모").grid(row=0, column=2, padx=4, pady=4, sticky="w")
        note_entry = ttk.Entry(ctrl, width=50)
        note_entry.grid(row=0, column=3, padx=4, pady=4, sticky="we")
        ctrl.columnconfigure(3, weight=1)

        def _selected_phone():
            items = tree.selection()
            if not items:
                return ""
            return str(items[0])

        def _refresh():
            current_phone = _selected_phone() or selected_phone["value"]

            rows = db_list_all_jobs()
            existing_ids = set(tree.get_children())

            for row in rows:
                iid = str(row.get("phone11") or "")
                values = (
                    row.get("name") or "",
                    row.get("phone11") or "",
                    row.get("status") or "",
                    row.get("last_status") or "",
                    row.get("last_error") or "",
                    row.get("note") or "",
                    _fmt_ts(row.get("updated_at")),
                )
                if iid in existing_ids:
                    tree.item(iid, values=values)
                    existing_ids.discard(iid)
                else:
                    tree.insert("", "end", iid=iid, values=values)

            for iid in existing_ids:
                tree.delete(iid)

            if current_phone and tree.exists(current_phone):
                tree.selection_set(current_phone)
                tree.focus(current_phone)
                tree.see(current_phone)

            root.after(2000, _refresh)

        def _load_selected(*args):
            phone11 = _selected_phone()
            selected_phone["value"] = phone11
            if not phone11:
                return
            row = db_get_job(phone11)
            if not row:
                return
            status_var.set(row.get("status") or STATUS_NEW)
            note_entry.delete(0, "end")
            note_entry.insert(0, row.get("note") or "")

        def _save_note():
            phone11 = _selected_phone()
            if not phone11:
                return
            db_set_note(phone11, note_entry.get().strip())
            sync_runtime_state_from_db()
            _load_selected()

        def _apply_status():
            phone11 = _selected_phone()
            if not phone11:
                return
            db_set_note(phone11, note_entry.get().strip())
            db_apply_manual_status(phone11, status_var.get().strip())
            clear_stop()
            sync_runtime_state_from_db()
            _load_selected()

        def _resume():
            phone11 = _selected_phone()
            if not phone11:
                return
            db_set_note(phone11, note_entry.get().strip())
            db_resume_job(phone11)
            clear_stop()
            sync_runtime_state_from_db()
            _load_selected()

        def _cancel():
            phone11 = _selected_phone()
            if not phone11:
                return
            db_set_note(phone11, note_entry.get().strip())
            db_apply_manual_status(phone11, STATUS_CANCELLED)
            clear_stop()
            sync_runtime_state_from_db()
            _load_selected()

        def _delete():
            phone11 = _selected_phone()
            if not phone11:
                return
            db_set_note(phone11, note_entry.get().strip())
            db_soft_delete_job(phone11)
            clear_stop()
            sync_runtime_state_from_db()
            _load_selected()

        def _hold():
            phone11 = _selected_phone()
            if not phone11:
                return
            db_set_note(phone11, note_entry.get().strip())
            db_apply_manual_status(phone11, STATUS_HOLD)
            clear_stop()
            sync_runtime_state_from_db()
            _load_selected()

        def _done():
            phone11 = _selected_phone()
            if not phone11:
                return
            db_set_note(phone11, note_entry.get().strip())
            db_apply_manual_status(phone11, STATUS_DONE)
            clear_stop()
            sync_runtime_state_from_db()
            _load_selected()

        def _force_check():
            phone11 = _selected_phone()
            if not phone11:
                return
            db_force_next_check(phone11)
            clear_stop()
            sync_runtime_state_from_db()
            _load_selected()

        ttk.Button(ctrl, text="메모저장", command=_save_note).grid(row=0, column=4, padx=4, pady=4)
        ttk.Button(ctrl, text="상태적용", command=_apply_status).grid(row=0, column=5, padx=4, pady=4)
        ttk.Button(ctrl, text="재개", command=_resume).grid(row=0, column=6, padx=4, pady=4)
        ttk.Button(ctrl, text="취소", command=_cancel).grid(row=0, column=7, padx=4, pady=4)
        ttk.Button(ctrl, text="삭제", command=_delete).grid(row=0, column=8, padx=4, pady=4)
        ttk.Button(ctrl, text="보류", command=_hold).grid(row=0, column=9, padx=4, pady=4)
        ttk.Button(ctrl, text="완료", command=_done).grid(row=0, column=10, padx=4, pady=4)
        ttk.Button(ctrl, text="다음확인 즉시", command=_force_check).grid(row=0, column=11, padx=4, pady=4)

        tree.bind("<<TreeviewSelect>>", _load_selected)

        _refresh()
        root.mainloop()

    t = threading.Thread(target=_run, daemon=True)
    t.start()

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
    def _find_special_note_text():
        try:
            ta = modal.find_element(By.XPATH, ".//textarea")
            if ta and ta.is_displayed():
                v = (ta.get_attribute("value") or "").strip()
                if v:
                    return v
                v = (ta.text or "").strip()
                if v:
                    return v
        except Exception:
            pass

        try:
            box = modal.find_element(
                By.XPATH,
                ".//*[contains(normalize-space(.), '특이사항')]/following::*[self::textarea or self::div or self::p][1]"
            )
            if box and box.is_displayed():
                v = (box.get_attribute("value") or "").strip()
                if v:
                    return v
                v = (box.get_attribute("textContent") or "").strip()
                if v:
                    return v
                v = (box.text or "").strip()
                if v:
                    return v
        except Exception:
            pass

        return ""

    def _is_checkbox_checked_near_label(label_keywords):
        for kw in label_keywords:
            xpaths = [
                f".//*[contains(normalize-space(.), '{kw}')]/preceding::*[self::input[@type='checkbox']][1]",
                f".//*[contains(normalize-space(.), '{kw}')]/ancestor::*[self::label or self::div or self::td or self::tr][1]//*[self::input[@type='checkbox']][1]",
                f".//label[contains(normalize-space(.), '{kw}')]//*[self::input[@type='checkbox']][1]",
                f".//tr[.//*[contains(normalize-space(.), '{kw}')]]//*[self::input[@type='checkbox']][1]",
            ]

            for xp in xpaths:
                try:
                    el = modal.find_element(By.XPATH, xp)
                    checked = (el.get_attribute("checked") or "").strip().lower()
                    selected = (el.get_attribute("selected") or "").strip().lower()
                    aria_checked = (el.get_attribute("aria-checked") or "").strip().lower()

                    if checked in ["true", "checked", "1", "on"]:
                        return True
                    if selected in ["true", "selected", "1", "on"]:
                        return True
                    if aria_checked in ["true", "checked", "1", "on"]:
                        return True

                    cls = (el.get_attribute("class") or "").strip().lower()
                    if "checked" in cls or "selected" in cls:
                        return True
                except Exception:
                    pass

        return False

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

    third_brand = find_input_near_label(modal, ["타사 브랜드", "타사브랜드"])
    third_install_mfg = find_input_near_label(
        modal,
        ["설치 또는 제조년월", "설치또는제조년월", "설치/제조년월", "설치 년월", "제조년월"]
    )
    third_product_type = find_input_near_label(modal, ["제품 타입", "제품타입"])
    third_product_kind = find_input_near_label(modal, ["제품 종류", "제품종류"])
    third_install_shape = find_input_near_label(modal, ["설치 형태", "설치형태"])
    third_watermark_no = find_input_near_label(
        modal,
        ["물마크 번호", "물마크번호", "워터마크 번호", "워터마크번호"]
    )

    pickup_request_raw = ""
    if _is_checkbox_checked_near_label(["기존 제품 수거", "기존제품수거"]):
        pickup_request_raw = "기존 제품 수거요청"

    special_note_raw = _find_special_note_text()

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
        "pickup_request_raw": (pickup_request_raw or "").strip(),
        "special_note_raw": (special_note_raw or "").strip(),
        "third_brand": (third_brand or "").strip(),
        "third_install_mfg": (third_install_mfg or "").strip(),
        "third_product_type": (third_product_type or "").strip(),
        "third_product_kind": (third_product_kind or "").strip(),
        "third_install_shape": (third_install_shape or "").strip(),
        "third_watermark_no": (third_watermark_no or "").strip(),
    }
# ---------------------------
# uiautomator2 helpers
# ---------------------------
def connect_emulator():
    d = u2.connect(ADB_SERIAL)
    d.implicitly_wait(5.0)
    return d

_ADB_EXE_CACHE = ""
ADB_DEVICE = None

def connect_emulator():
    global ADB_DEVICE
    d = u2.connect(ADB_SERIAL)
    d.implicitly_wait(5.0)
    ADB_DEVICE = d
    return d

def _resolve_adb_executable() -> str:
    candidates = []

    for env_key in ["ADB_EXE", "ADB_PATH"]:
        v = str(os.environ.get(env_key) or "").strip().strip('"')
        if v:
            candidates.append(v)

    for cmd in ["adb", "adb.exe"]:
        found = shutil.which(cmd)
        if found:
            candidates.append(found)

    sdk_roots = [
        str(os.environ.get("ANDROID_SDK_ROOT") or "").strip(),
        str(os.environ.get("ANDROID_HOME") or "").strip(),
        os.path.join(str(os.environ.get("LOCALAPPDATA") or "").strip(), "Android", "Sdk"),
        os.path.join(os.path.expanduser("~"), "AppData", "Local", "Android", "Sdk"),
    ]

    for root in sdk_roots:
        if not root:
            continue
        candidates.append(os.path.join(root, "platform-tools", "adb.exe"))
        candidates.append(os.path.join(root, "platform-tools", "adb"))

    seen = set()
    normalized = []

    for c in candidates:
        cc = str(c or "").strip().strip('"')
        if not cc:
            continue
        cc = os.path.normpath(cc)
        if cc in seen:
            continue
        seen.add(cc)
        normalized.append(cc)

    for path in normalized:
        if os.path.isfile(path):
            return path

    return ""

def _normalize_u2_shell_result(res):
    output = ""
    exit_code = 0

    if hasattr(res, "output"):
        output = str(getattr(res, "output") or "").strip()
        exit_code = getattr(res, "exit_code", 0)
    elif isinstance(res, tuple):
        if len(res) >= 1:
            output = str(res[0] or "").strip()
        if len(res) >= 2:
            try:
                exit_code = int(res[1])
            except Exception:
                exit_code = 0 if str(res[1]).strip() in ["", "0", "None"] else 1
    else:
        output = str(res or "").strip()
        exit_code = 0

    ok = (exit_code in [0, None, "0"])
    return ok, output, ""

def _uia2_shell_run(args, timeout_sec: float = 10.0):
    global ADB_DEVICE

    if ADB_DEVICE is None:
        return False, "", "uiautomator2 device not ready"

    arg_list = [str(a) for a in list(args)]
    last_err = ""

    try:
        res = ADB_DEVICE.shell(arg_list, timeout=timeout_sec)
        ok, out, err = _normalize_u2_shell_result(res)
        if ok:
            return True, out, ""
        last_err = out or err or "uiautomator2 shell non-zero"
    except TypeError:
        pass
    except Exception as e:
        last_err = str(e)

    try:
        res = ADB_DEVICE.shell(" ".join(arg_list))
        ok, out, err = _normalize_u2_shell_result(res)
        if ok:
            return True, out, ""
        last_err = out or err or last_err or "uiautomator2 shell non-zero"
    except Exception as e:
        if last_err:
            last_err = f"{last_err} / {e}"
        else:
            last_err = str(e)

    return False, "", last_err or "uiautomator2 shell failed"

def adb_run(args, timeout_sec: float = 10.0):
    global _ADB_EXE_CACHE

    raw_args = [str(a) for a in list(args)]
    shell_args = list(raw_args)

    if shell_args and shell_args[0].lower() == "shell":
        shell_args = shell_args[1:]

    ok, out, err = _uia2_shell_run(shell_args, timeout_sec=timeout_sec)
    if ok:
        return ok, out, err

    if not _ADB_EXE_CACHE:
        _ADB_EXE_CACHE = _resolve_adb_executable()

    if not _ADB_EXE_CACHE:
        return False, "", f"uiautomator2 shell 실패 / adb.exe 경로도 찾지 못함 / {err}"

    try:
        result = subprocess.run(
            [_ADB_EXE_CACHE, "-s", ADB_SERIAL] + raw_args,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="ignore",
            timeout=timeout_sec,
        )
        return (
            result.returncode == 0,
            (result.stdout or "").strip(),
            (result.stderr or "").strip(),
        )
    except Exception as e:
        return False, "", str(e)

def ensure_adb_keyboard_ime() -> bool:
    ok, out, err = adb_run(["shell", "ime", "set", "com.android.adbkeyboard/.AdbIME"], timeout_sec=8.0)
    if not ok:
        print("⚠️ [adb_keyboard] ime set 실패:", err or out)
        return False
    time.sleep(0.4)
    return True

def adb_keyboard_clear_text() -> bool:
    ok, out, err = adb_run(["shell", "am", "broadcast", "-a", "ADB_CLEAR_TEXT"], timeout_sec=8.0)
    if not ok:
        print("⚠️ [adb_keyboard] clear_text 실패:", err or out)
        return False
    time.sleep(0.4)
    return True

def adb_keyboard_input_text(raw_text: str) -> bool:
    raw_text = str(raw_text or "")
    try:
        msg_b64 = base64.b64encode(raw_text.encode("utf-8")).decode("ascii")
    except Exception as e:
        print("⚠️ [adb_keyboard] base64 인코딩 실패:", e)
        return False

    ok, out, err = adb_run(
        ["shell", "am", "broadcast", "-a", "ADB_INPUT_B64", "--es", "msg", msg_b64],
        timeout_sec=10.0
    )
    if ok:
        time.sleep(0.8)
        return True

    print("⚠️ [adb_keyboard] ADB_INPUT_B64 실패:", err or out)

    ok2, out2, err2 = adb_run(
        ["shell", "am", "broadcast", "-a", "ADB_INPUT_TEXT", "--es", "msg", raw_text],
        timeout_sec=10.0
    )
    if ok2:
        time.sleep(0.8)
        return True

    print("⚠️ [adb_keyboard] ADB_INPUT_TEXT 실패:", err2 or out2)
    return False

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
            print("❌ [enter_order_status] 주문 이어하기의 일반주문 행을 찾지 못함")
            return False

        left, top, right, bottom = target
        cy = (top + bottom) // 2
        w, h = d.window_size()

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

        for x, y in filtered_points:
            d.click(x, y)
            time.sleep(1.2)

            if d(text="주문현황").exists:
                ensure_general_tab(d, force_click=True)
                time.sleep(0.4)
                print("✅ [enter_order_status] 주문현황 진입 성공")
                return True

        fallback_x1 = int(w * 0.82)
        d.click(fallback_x1, cy)
        time.sleep(1.2)

        if d(text="주문현황").exists:
            ensure_general_tab(d, force_click=True)
            time.sleep(0.4)
            print("✅ [enter_order_status] 주문현황 진입 성공")
            return True

        fallback_x2 = int(w * 0.90)
        d.click(fallback_x2, cy)
        time.sleep(1.2)

        if d(text="주문현황").exists:
            ensure_general_tab(d, force_click=True)
            time.sleep(0.4)
            print("✅ [enter_order_status] 주문현황 진입 성공")
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

        for i in range(cnt):
            try:
                e = edits[i]
                b = e.info.get("bounds", {})
                left = int(b.get("left", 0))
                top = int(b.get("top", 0))
                right = int(b.get("right", 0))
                bottom = int(b.get("bottom", 0))
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

        if best is None:
            print("❌ [find_search_edittext] 상단 검색 입력칸을 찾지 못함")

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
    status_texts = [target_status] if target_status else [
        "인증완료",
        "설치입력",
        "결제입력",
        "할인입력",
        "인증입력",
        "서명입력",
        "주문확정",
        "주문삭제",
        "주문불가",
    ]

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
    loading_log_count = 0

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
                loading_log_count += 1
                if loading_log_count % 3 == 1:
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

def try_open_ready_sign_detail(d, job: dict, entry_status: str = "인증완료") -> bool:
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
            ("그린", "그린"),
            ("green", "그린"),
            ("민트", "그린"),
            ("mint", "그린"),
            ("블루", "블루"),
            ("blue", "블루"),
            ("핑크", "핑크"),
            ("pink", "핑크"),
            ("블랙", "블랙"),
            ("black", "블랙"),
            ("실버", "실버"),
            ("silver", "실버"),
            ("브라운", "브라운"),
            ("brown", "브라운"),
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

        print("⏳ [ready_sign] 색상 선택 후 주문접수/상품선택 전환 대기 12.0초")
        time.sleep(12.0)

        if _is_search_return_popup_open():
            _dismiss_search_return_popup_if_open()

        if not _wait_product_option_ready(timeout_sec=16.0):
            return _abort(f"전자서명 중단: 상품선택 화면 준비 실패 / 고객={job['name']} / 모델={model_query}")

        return True

    def _has_sale_group_options_visible() -> bool:
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

        return has_rental and has_cash

    def _click_sale_group_expand_arrow() -> bool:
        label_obj = None

        try:
            if d(text="판매구분").exists:
                label_obj = d(text="판매구분")
            elif d(textContains="판매구분").exists:
                label_obj = d(textContains="판매구분")
        except Exception:
            label_obj = None

        if label_obj is None:
            return False

        try:
            w, h = d.window_size()
            b = label_obj.info.get("bounds", {})
            top = int(b.get("top", 0))
            bottom = int(b.get("bottom", 0))
            cy = (top + bottom) // 2

            candidate_points = [
                (int(w * 0.95), cy),
                (int(w * 0.92), cy),
                (int(w * 0.89), cy),
                (int(w * 0.85), cy),
                (int(w * 0.80), cy),
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

            for x, y in filtered_points:
                d.click(x, y)
                time.sleep(1.8)

                if _has_sale_group_options_visible():
                    print("✅ [ready_sign] 판매구분 옵션 펼치기 완료")
                    return True

            return False

        except Exception:
            return False

    def _wait_product_option_ready(timeout_sec: float = 16.0) -> bool:
        end_at = time.time() + timeout_sec
        stable_ok_count = 0

        while time.time() < end_at:
            try:
                if is_unexpected_digital_sales_home(d):
                    return False

                if _is_search_return_popup_open():
                    _dismiss_search_return_popup_if_open()
                    stable_ok_count = 0
                    time.sleep(0.6)
                    continue

                has_title = d(text="상품선택").exists or d(text="주문접수").exists
                has_sale_group = d(text="판매구분").exists or d(textContains="판매구분").exists

                if has_title and has_sale_group:
                    stable_ok_count += 1
                    if stable_ok_count >= 2:
                        print("✅ [ready_sign] 판매구분 화면 준비 완료")
                        time.sleep(0.8)
                        return True
                else:
                    stable_ok_count = 0

            except Exception:
                stable_ok_count = 0

            time.sleep(0.6)

        print("❌ [ready_sign] 판매구분 화면 준비 실패")
        return False

    def _select_rental_option() -> bool:
        if not _wait_product_option_ready(timeout_sec=16.0):
            return _abort(f"전자서명 중단: 판매구분 화면 진입 실패 / 고객={job['name']}")

        for attempt in range(8):
            if _is_search_return_popup_open():
                _dismiss_search_return_popup_if_open()

            if _has_sale_group_options_visible():
                ok_click = click_text_center(d, "렌탈", 0.35, 0.90)
                if not ok_click:
                    return _abort(f"전자서명 중단: 렌탈 옵션 클릭 실패 / 고객={job['name']}")

                time.sleep(2.0)
                print(f"✅ [ready_sign] 렌탈 선택 완료: {job['name']}")
                return True

            print(f"ℹ️ [ready_sign] 판매구분 옵션 펼치기 시도 {attempt + 1}/8")

            expanded = _click_sale_group_expand_arrow()
            if expanded:
                time.sleep(1.2)

                if _has_sale_group_options_visible():
                    ok_click = click_text_center(d, "렌탈", 0.35, 0.90)
                    if not ok_click:
                        return _abort(f"전자서명 중단: 렌탈 옵션 클릭 실패 / 고객={job['name']}")

                    time.sleep(2.0)
                    print(f"✅ [ready_sign] 렌탈 선택 완료: {job['name']}")
                    return True

            time.sleep(1.5)

        return _abort(f"전자서명 중단: 판매구분 렌탈/일시불 옵션 미검출 / 고객={job['name']}")

    def _select_manage_type(manage_raw: str) -> bool:
        target = _normalize_manage_target(manage_raw)
        if not target:
            return _abort(f"전자서명 중단: 모달 관리 추출 실패 / 고객={job['name']} / 원본관리={manage_raw}")

        if target == "자가":
            exact_candidates = ["자가관리", "자가", "셀프"]
            loose_candidates = ["자가관리", "자가", "셀프"]
        else:
            exact_candidates = [
                f"방문관리-{target}",
                target,
                target.replace("M", "개월"),
                target.replace("M", "개월관리"),
            ]
            loose_candidates = list(exact_candidates)

        print(f"🔎 [ready_sign] 관리유형 선택 목표: {target}")

        def _norm(s: str) -> str:
            return re.sub(r"\s+", "", str(s or "")).strip()

        def _pick_manage_candidate(y_min_ratio: float, y_max_ratio: float):
            items = _collect_visible_text_items(y_min_ratio, y_max_ratio)
            exact_norms = [_norm(x) for x in exact_candidates]
            loose_norms = [_norm(x) for x in loose_candidates]

            filtered = []
            for it in items:
                txt = str(it.get("text") or "").strip()
                txt_norm = _norm(txt)
                if not txt_norm:
                    continue

                # 안내문/설명문/긴 문장 제외
                if len(txt_norm) > 16:
                    continue
                if "선택할수있습니다" in txt_norm:
                    continue
                if "방문하여관리를받거나" in txt_norm:
                    continue
                if "유형을선택" in txt_norm:
                    continue

                filtered.append({
                    "text": txt,
                    "text_norm": txt_norm,
                    "bounds": it["bounds"],
                    "class_name": it.get("class_name", ""),
                })

            # 1) 완전일치 우선
            for want in exact_norms:
                for it in filtered:
                    if it["text_norm"] == want:
                        return it

            # 2) 부분일치는 짧은 후보만 허용
            for want in loose_norms:
                for it in filtered:
                    if want and want in it["text_norm"]:
                        return it

            return None

        def _click_manage_item(item) -> bool:
            try:
                left, top, right, bottom = item["bounds"]
                cx = (left + right) // 2
                cy = (top + bottom) // 2
                d.click(cx, cy)
                time.sleep(1.2)
                return True
            except Exception:
                return False

        item = _pick_manage_candidate(0.40, 0.92)
        if item is not None:
            if _click_manage_item(item):
                print(f"✅ [ready_sign] 관리유형 선택 완료: {item['text']}")
                return True

        if d(text="관리유형").exists:
            click_text_center(d, "관리유형", 0.35, 0.90)
            time.sleep(1.2)

        item = _pick_manage_candidate(0.40, 0.95)
        if item is not None:
            if _click_manage_item(item):
                print(f"✅ [ready_sign] 관리유형 선택 완료: {item['text']}")
                return True

        try:
            w, h = d.window_size()
            d.swipe(int(w * 0.50), int(h * 0.82), int(w * 0.50), int(h * 0.60), 0.20)
            time.sleep(0.6)
        except Exception:
            pass

        item = _pick_manage_candidate(0.35, 0.98)
        if item is not None:
            if _click_manage_item(item):
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

    def _dismiss_payment_info_notice_if_open() -> bool:
        notice_fragments = [
            "기본 설정 됩니다",
            "사용 가능한 주문 건에만",
            "등록한 결제정보 중",
        ]

        found = False

        for frag in notice_fragments:
            try:
                if d(text=frag).exists:
                    found = True
                    break
            except Exception:
                pass

            try:
                if d(textContains=frag).exists:
                    found = True
                    break
            except Exception:
                pass

        if not found:
            return False

        print("⚠️ [ready_sign] 결제정보 기본설정 안내 팝업 감지 → [확인] 클릭")

        clicked = False

        try:
            if d(text="확인").exists:
                d(text="확인").click()
                clicked = True
        except Exception:
            clicked = False

        if not clicked:
            try:
                if d(textContains="확인").exists:
                    d(textContains="확인").click()
                    clicked = True
            except Exception:
                clicked = False

        if not clicked:
            try:
                clicked = click_text_center(d, "확인", 0.45, 0.98)
            except Exception:
                clicked = False

        time.sleep(1.8)
        return clicked

    def _wait_payment_info_ready(timeout_sec: float = 12.0) -> bool:
        end_at = time.time() + timeout_sec
        stable_ok_count = 0

        while time.time() < end_at:
            try:
                if is_unexpected_digital_sales_home(d):
                    return False

                if _dismiss_payment_info_notice_if_open():
                    stable_ok_count = 0
                    time.sleep(0.4)
                    continue

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
                ok = click_text_center(d, "정기결제 수단 선택", 0.20, 0.70)
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

    def _wait_payment_method_add_page_ready(timeout_sec: float = 10.0) -> bool:
        end_at = time.time() + timeout_sec
        stable_ok_count = 0

        while time.time() < end_at:
            try:
                if is_unexpected_digital_sales_home(d):
                    return False

                has_title = d(text="결제수단 추가").exists or d(textContains="결제수단 추가").exists
                has_card_tab = d(text="카드이체").exists or d(textContains="카드이체").exists
                has_bank_tab = d(text="은행이체").exists or d(textContains="은행이체").exists

                if has_title and (has_card_tab or has_bank_tab):
                    stable_ok_count += 1
                    if stable_ok_count >= 2:
                        print("✅ [ready_sign] 결제수단 추가 화면 준비 완료")
                        time.sleep(0.5)
                        return True
                else:
                    stable_ok_count = 0
            except Exception:
                stable_ok_count = 0

            time.sleep(0.4)

        print("❌ [ready_sign] 결제수단 추가 화면 준비 실패")
        return False

    def _is_card_expiry_like(account_raw: str) -> bool:
        s = re.sub(r"\s+", "", str(account_raw or ""))
        if not s:
            return False

        return bool(
            re.search(r"(?<!\d)\d{2}/\d{2}(?!\d)", s)
            or re.search(r"(?<!\d)\d{4}/\d{2}(?!\d)", s)
            or re.search(r"(?<!\d)\d{2}/\d{4}(?!\d)", s)
        )

    def _extract_bank_name_from_account(account_raw: str) -> str:
        s = re.sub(r"\s+", "", str(account_raw or ""))
        if not s:
            return ""

        s = re.sub(r"[\d\*\-_/]", "", s)
        s = re.sub(r"[^가-힣A-Za-z]", "", s)
        return s.strip()

    def _extract_account_digits(account_raw: str) -> str:
        return normalize_digits(account_raw or "")

    def _read_edit_obj_text(obj) -> str:
        try:
            return (obj.get_text() or "").strip()
        except Exception:
            try:
                info = obj.info or {}
                return str(info.get("text") or "").strip()
            except Exception:
                return ""

    def _find_first_label_obj(label_candidates):
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

    def _find_edittext_near_label(label_candidates, y_tolerance: int = 110, y_below: int = 220):
        label_obj = _find_first_label_obj(label_candidates)

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

    def _find_top_widest_edittext(y_min_ratio: float = 0.05, y_max_ratio: float = 0.25):
        try:
            w, h = d.window_size()
            edits = d(className="android.widget.EditText")
            cnt = edits.count
        except Exception:
            return None

        best = None
        best_w = 0
        best_top = 10**9

        for i in range(cnt):
            try:
                e = edits[i]
                b = e.info.get("bounds", {})
                left = int(b.get("left", 0))
                top = int(b.get("top", 0))
                right = int(b.get("right", 0))
                bottom = int(b.get("bottom", 0))
                bw = right - left

                if top < int(h * y_min_ratio) or bottom > int(h * y_max_ratio):
                    continue

                if bw > best_w or (bw == best_w and top < best_top):
                    best = e
                    best_w = bw
                    best_top = top
            except Exception:
                continue

        return best

    def _find_lowest_widest_edittext(y_min_ratio: float = 0.20, y_max_ratio: float = 0.85):
        try:
            w, h = d.window_size()
            edits = d(className="android.widget.EditText")
            cnt = edits.count
        except Exception:
            return None

        best = None
        best_top = -1
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

                if top < int(h * y_min_ratio) or bottom > int(h * y_max_ratio):
                    continue

                if top > best_top or (top == best_top and bw > best_w):
                    best = e
                    best_top = top
                    best_w = bw
            except Exception:
                continue

        return best

    def _click_bank_transfer_tab_if_needed(account_raw: str) -> bool:
        if _is_card_expiry_like(account_raw):
            print(f"ℹ️ [ready_sign] 카드 유효기간 패턴 감지 → 카드이체 탭 유지 / 원본결제정보={account_raw}")
            return True

        try:
            if d(text="은행이체").exists:
                ok = click_text_center(d, "은행이체", 0.05, 0.20)
                time.sleep(2.0)
                if ok:
                    print("✅ [ready_sign] 은행이체 탭 선택 완료")
                    return True
        except Exception:
            pass

        try:
            if d(textContains="은행이체").exists:
                d(textContains="은행이체").click()
                time.sleep(2.0)
                print("✅ [ready_sign] 은행이체 탭 선택 완료")
                return True
        except Exception:
            pass

        return False

    def _open_bank_picker() -> bool:
        try:
            if d(text="은행입력").exists:
                ok = click_text_center(d, "은행입력", 0.12, 0.45)
                time.sleep(1.8)
                return ok
        except Exception:
            pass

        try:
            if d(textContains="은행입력").exists:
                d(textContains="은행입력").click()
                time.sleep(1.8)
                return True
        except Exception:
            pass

        label_obj = None

        try:
            if d(text="은행").exists:
                label_obj = d(text="은행")
            elif d(textContains="은행").exists:
                label_obj = d(textContains="은행")
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
                    time.sleep(1.8)
                    return True
            except Exception:
                pass

        return False

    def _collect_visible_bank_candidates():
        items = _collect_visible_text_items(0.12, 0.96)

        found = []
        seen = set()

        skip_texts = [
            "은행선택",
            "결제수단추가",
            "카드이체",
            "은행이체",
            "취소",
            "추가하기",
            "추가",
            "은행입력",
            "계좌번호",
            "이체일",
            "명의",
            "명의자",
            "법정생년월일",
        ]

        for it in items:
            txt = re.sub(r"\s+", "", str(it["text"] or ""))
            if not txt:
                continue
            if txt in seen:
                continue
            if len(txt) > 20:
                continue
            if any(skip in txt for skip in skip_texts):
                continue

            seen.add(txt)
            found.append({
                "text": txt,
                "bounds": it["bounds"],
                "class_name": it["class_name"],
            })

        return found

    def _bank_match_score(query_norm: str, cand_norm: str) -> int:
        if not query_norm or not cand_norm:
            return 0

        cand_has_security = ("증권" in cand_norm) or ("투자" in cand_norm)
        query_has_security = ("증권" in query_norm) or ("투자" in query_norm)

        if cand_has_security and not query_has_security:
            return 0

        if cand_has_security and query_has_security:
            if cand_norm == query_norm:
                return 100
            return 0

        if cand_norm == query_norm:
            return 95

        q_simple = query_norm.replace("은행", "").replace("뱅크", "")
        c_simple = cand_norm.replace("은행", "").replace("뱅크", "")

        if q_simple and c_simple and q_simple == c_simple:
            return 90

        if query_norm in cand_norm:
            return 80

        if q_simple and q_simple in cand_norm:
            return 75

        if cand_norm in query_norm:
            return 70

        if c_simple and c_simple in query_norm:
            return 65

        return 0

    def _choose_bank_from_picker(account_raw: str) -> bool:
        bank_query = _extract_bank_name_from_account(account_raw)
        if not bank_query:
            return _abort(f"전자서명 중단: 결제정보에서 은행명 추출 실패 / 고객={job['name']} / 원본결제정보={account_raw}")

        query_norm = re.sub(r"\s+", "", bank_query)
        print(f"🔎 [ready_sign] 은행 선택 목표: {bank_query}")

        for step in range(12):
            candidates = _collect_visible_bank_candidates()

            best_item = None
            best_score = 0

            for c in candidates:
                cand_norm = re.sub(r"\s+", "", c["text"])
                score = _bank_match_score(query_norm, cand_norm)
                if score > best_score:
                    best_score = score
                    best_item = c

            if best_item is not None and best_score > 0:
                left, top, right, bottom = best_item["bounds"]
                cx = (left + right) // 2
                cy = (top + bottom) // 2
                d.click(cx, cy)
                time.sleep(2.0)
                print(f"✅ [ready_sign] 은행 선택 완료: {best_item['text']} / 원본결제정보={account_raw}")
                return True

            if step < 11:
                print(f"ℹ️ [ready_sign] 은행 리스트 스크롤 {step + 1}/11")
                try:
                    w, h = d.window_size()
                    d.swipe(int(w * 0.50), int(h * 0.84), int(w * 0.50), int(h * 0.25), 0.25)
                    time.sleep(1.2)
                except Exception:
                    pass

        return _abort(f"전자서명 중단: 은행 리스트에서 일치 은행 미검출 / 고객={job['name']} / 원본결제정보={account_raw} / 목표은행={bank_query}")

    def _fill_bank_account_number(account_raw: str) -> bool:
        account_digits = _extract_account_digits(account_raw)
        if not account_digits:
            return _abort(f"전자서명 중단: 결제정보 숫자 추출 실패 / 고객={job['name']} / 원본결제정보={account_raw}")

        edit = _find_edittext_near_label(["계좌번호"], y_tolerance=80, y_below=180)
        if edit is None:
            return _abort(f"전자서명 중단: 계좌번호 입력칸 미검출 / 고객={job['name']}")

        current_digits = normalize_digits(_read_edit_obj_text(edit))
        if current_digits == account_digits:
            print(f"✅ [ready_sign] 계좌번호 이미 입력 일치: {account_digits}")
            return True

        ok = type_into_edittext(d, edit, account_digits)
        if not ok:
            return _abort(f"전자서명 중단: 계좌번호 입력 실패 / 고객={job['name']} / 계좌번호={account_digits}")

        time.sleep(1.0)

        current_digits = normalize_digits(_read_edit_obj_text(edit))
        if current_digits != account_digits:
            return _abort(f"전자서명 중단: 계좌번호 입력값 불일치 / 고객={job['name']} / 기대={account_digits} / 현재={current_digits}")

        print(f"✅ [ready_sign] 계좌번호 입력 완료: {account_digits}")
        return True

    def _click_payment_method_submit() -> bool:
        try:
            if d(text="추가하기").exists:
                ok = click_text_center(d, "추가하기", 0.82, 0.99)
                time.sleep(2.0)
                return ok
        except Exception:
            pass

        try:
            if d(textContains="추가하기").exists:
                d(textContains="추가하기").click()
                time.sleep(2.0)
                return True
        except Exception:
            pass

        return False

    def _click_confirm_if_exists() -> bool:
        try:
            if d(text="확인").exists:
                d(text="확인").click()
                time.sleep(1.2)
                return True
        except Exception:
            pass

        try:
            if d(textContains="확인").exists:
                d(textContains="확인").click()
                time.sleep(1.2)
                return True
        except Exception:
            pass

        try:
            ok = click_text_center(d, "확인", 0.45, 0.98)
            if ok:
                time.sleep(1.2)
            return ok
        except Exception:
            return False

    def _click_cancel_if_exists() -> bool:
        try:
            if d(text="취소").exists:
                d(text="취소").click()
                time.sleep(1.2)
                return True
        except Exception:
            pass

        try:
            if d(textContains="취소").exists:
                d(textContains="취소").click()
                time.sleep(1.2)
                return True
        except Exception:
            pass

        try:
            ok = click_text_center(d, "취소", 0.45, 0.98)
            if ok:
                time.sleep(1.2)
            return ok
        except Exception:
            return False

    def _collect_popup_texts(y_min_ratio: float = 0.20, y_max_ratio: float = 0.82):
        items = _collect_visible_text_items(y_min_ratio, y_max_ratio)
        texts = []
        seen = set()

        for it in items:
            txt = re.sub(r"\s+", "", str(it["text"] or ""))
            if not txt:
                continue
            if txt in seen:
                continue
            seen.add(txt)
            texts.append(txt)

        return texts

    def _is_prev_stage_move_popup_open() -> bool:
        popup_fragments = [
            "이전단계로이동",
            "이전 단계로 이동",
            "설정하신상품의할인정보및설치정보는저장되지않아",
            "선택한상품만유지되며",
        ]

        texts = _collect_popup_texts(0.20, 0.82)
        joined = " ".join(texts)

        for frag in popup_fragments:
            if re.sub(r"\s+", "", frag) in joined:
                return True

        return False

    def _dismiss_prev_stage_move_popup_if_open() -> bool:
        if not _is_prev_stage_move_popup_open():
            return False

        print("⚠️ [ready_sign] 이전 단계 이동 팝업 감지 → [취소] 클릭 후 현재 단계 유지")
        clicked = _click_cancel_if_exists()
        time.sleep(1.2)
        return clicked

    def _is_payment_error_popup_open() -> bool:
        texts = _collect_popup_texts(0.20, 0.82)
        joined = " ".join(texts)

        if not joined:
            return False

        error_fragments = [
            "오류",
            "센터계좌체계오류",
            "결제정보",
            "계좌",
            "입력값",
            "불가",
            "실패",
        ]

        if _is_prev_stage_move_popup_open():
            return False

        for frag in error_fragments:
            if re.sub(r"\s+", "", frag) in joined:
                return True

        return False

    def _popup_message_for_log() -> str:
        texts = _collect_popup_texts(0.20, 0.82)
        ignore_fragments = [
            "확인",
            "취소",
            "추가하기",
            "추가",
            "결제수단추가",
            "카드이체",
            "은행이체",
        ]

        picked = []
        for txt in texts:
            if any(re.sub(r"\s+", "", frag) in txt for frag in ignore_fragments):
                continue
            picked.append(txt)

        return " | ".join(picked[:3]).strip() or "popup"

    def _wait_after_payment_submit(timeout_sec: float = 10.0) -> bool:
        end_at = time.time() + timeout_sec

        while time.time() < end_at:
            if is_unexpected_digital_sales_home(d):
                return _abort(f"전자서명 중단: 결제수단 추가 후 홈이탈 / 고객={job['name']}")

            if _dismiss_prev_stage_move_popup_if_open():
                time.sleep(0.8)
                continue

            if _is_payment_error_popup_open():
                popup_msg = _popup_message_for_log()
                _click_confirm_if_exists()
                return _abort(f"결제정보 오류 / 고객={job['name']} / 팝업={popup_msg}")

            if _wait_payment_info_ready(timeout_sec=1.5):
                print(f"✅ [ready_sign] 결제수단 추가 후 결제정보 선택 화면 복귀 확인: {job['name']}")
                return True

            if _wait_install_info_ready(timeout_sec=1.5):
                print(f"✅ [ready_sign] 결제수단 추가 후 설치정보 화면 진입 확인: {job['name']}")
                return True

            time.sleep(0.3)

        return _abort(f"전자서명 중단: 결제수단 추가 후 다음 화면 판정 실패 / 고객={job['name']}")

    def _click_payment_info_next() -> bool:
        if not _wait_payment_info_ready(timeout_sec=10.0):
            return False

        for attempt in range(3):
            _dismiss_payment_info_notice_if_open()
            _dismiss_prev_stage_move_popup_if_open()

            try:
                if d(text="다음").exists:
                    ok = click_text_center(d, "다음", 0.88, 0.99)
                elif d(textContains="다음").exists:
                    d(textContains="다음").click()
                    ok = True
                else:
                    ok = False
            except Exception:
                ok = False

            if not ok:
                time.sleep(0.8)
                continue

            time.sleep(2.0)

            if _dismiss_prev_stage_move_popup_if_open():
                print(f"⚠️ [ready_sign] 결제정보 다음 클릭 후 이전단계 팝업 발생 → 재시도 {attempt + 1}/3")
                time.sleep(0.8)
                continue

            if _wait_install_info_ready(timeout_sec=4.0):
                print(f"✅ [ready_sign] 결제정보 선택 화면 다음 클릭 완료: {job['name']}")
                return True

            if _wait_payment_info_ready(timeout_sec=1.5):
                print(f"⚠️ [ready_sign] 결제정보 다음 클릭 후 아직 같은 화면 → 재시도 {attempt + 1}/3")
                time.sleep(0.8)
                continue

        return False

    def _dismiss_existing_install_info_popup_if_open() -> bool:
        popup_fragments = [
            "설치처정보가있습니다",
            "설치처 정보가 있습니다",
            "해당정보로주문및설치를진행",
            "해당 정보로 주문 및 설치를 진행",
        ]

        items = _collect_visible_text_items(0.25, 0.80)
        found = False

        for it in items:
            txt = re.sub(r"\s+", "", str(it["text"] or ""))
            if any(re.sub(r"\s+", "", frag) in txt for frag in popup_fragments):
                found = True
                break

        if not found:
            return False

        print("⚠️ [ready_sign] 기존 설치처 정보 팝업 감지 → [취소] 클릭 후 직접 입력 진행")

        clicked = _click_cancel_if_exists()
        time.sleep(1.8)
        return clicked

    def _wait_install_info_ready(timeout_sec: float = 12.0) -> bool:
        end_at = time.time() + timeout_sec
        stable_ok_count = 0

        while time.time() < end_at:
            try:
                if is_unexpected_digital_sales_home(d):
                    return False

                if _dismiss_existing_install_info_popup_if_open():
                    stable_ok_count = 0
                    time.sleep(0.4)
                    continue

                has_title = d(text="설치정보").exists or d(textContains="설치정보").exists
                has_phone = d(text="휴대폰번호").exists or d(textContains="휴대폰번호").exists
                has_address = d(text="주소").exists or d(textContains="주소").exists

                if has_title and has_phone and has_address:
                    stable_ok_count += 1
                    if stable_ok_count >= 2:
                        print("✅ [ready_sign] 설치정보 화면 준비 완료")
                        time.sleep(0.5)
                        return True
                else:
                    stable_ok_count = 0
            except Exception:
                stable_ok_count = 0

            time.sleep(0.4)

        print("❌ [ready_sign] 설치정보 화면 준비 실패")
        return False

    def _fill_install_phone(phone11: str) -> bool:
        if not phone11:
            return _abort(f"전자서명 중단: 휴대폰번호 값 없음 / 고객={job['name']}")

        edit = _find_edittext_near_label(["휴대폰번호"], y_tolerance=80, y_below=160)
        if edit is None:
            return _abort(f"전자서명 중단: 휴대폰번호 입력칸 미검출 / 고객={job['name']}")

        current_digits = normalize_digits(_read_edit_obj_text(edit))
        if current_digits == phone11:
            print(f"✅ [ready_sign] 휴대폰번호 이미 입력 일치: {phone11}")
            return True

        ok = type_into_edittext(d, edit, phone11)
        if not ok:
            return _abort(f"전자서명 중단: 휴대폰번호 입력 실패 / 고객={job['name']} / 휴대폰번호={phone11}")

        time.sleep(1.0)

        current_digits = normalize_digits(_read_edit_obj_text(edit))
        if current_digits != phone11:
            return _abort(f"전자서명 중단: 휴대폰번호 입력값 불일치 / 고객={job['name']} / 기대={phone11} / 현재={current_digits}")

        print(f"✅ [ready_sign] 휴대폰번호 입력 완료: {phone11}")
        return True

    def _click_install_address_input_button() -> bool:
        try:
            if d(text="주소 입력").exists:
                ok = click_text_center(d, "주소 입력", 0.25, 0.70)
                time.sleep(1.8)
                return ok
        except Exception:
            pass

        try:
            if d(textContains="주소 입력").exists:
                d(textContains="주소 입력").click()
                time.sleep(1.8)
                return True
        except Exception:
            pass

        return False

    def _wait_address_management_ready(timeout_sec: float = 10.0) -> bool:
        end_at = time.time() + timeout_sec
        stable_ok_count = 0

        while time.time() < end_at:
            try:
                if is_unexpected_digital_sales_home(d):
                    return False

                has_title = d(text="주소지 관리").exists or d(textContains="주소지 관리").exists
                has_add = d(text="새 주소지 추가").exists or d(textContains="새 주소지 추가").exists

                if has_title and has_add:
                    stable_ok_count += 1
                    if stable_ok_count >= 2:
                        print("✅ [ready_sign] 주소지 관리 화면 준비 완료")
                        time.sleep(0.4)
                        return True
                else:
                    stable_ok_count = 0
            except Exception:
                stable_ok_count = 0

            time.sleep(0.4)

        print("❌ [ready_sign] 주소지 관리 화면 준비 실패")
        return False

    def _click_new_address_add() -> bool:
        try:
            if d(text="새 주소지 추가").exists:
                ok = click_text_center(d, "새 주소지 추가", 0.30, 0.80)
                time.sleep(1.8)
                return ok
        except Exception:
            pass

        try:
            if d(textContains="새 주소지 추가").exists:
                d(textContains="새 주소지 추가").click()
                time.sleep(1.8)
                return True
        except Exception:
            pass

        return False

    def _wait_address_search_ready(timeout_sec: float = 10.0) -> bool:
        end_at = time.time() + timeout_sec
        stable_ok_count = 0

        while time.time() < end_at:
            try:
                if is_unexpected_digital_sales_home(d):
                    return False

                has_title = d(text="주소입력").exists or d(textContains="주소입력").exists
                has_search = d(text="검색").exists or d(textContains="검색").exists
                edit = _find_top_widest_edittext(0.05, 0.25)

                if has_title and has_search and edit is not None:
                    stable_ok_count += 1
                    if stable_ok_count >= 2:
                        print("✅ [ready_sign] 주소 검색 화면 준비 완료")
                        time.sleep(0.4)
                        return True
                else:
                    stable_ok_count = 0
            except Exception:
                stable_ok_count = 0

            time.sleep(0.4)

        print("❌ [ready_sign] 주소 검색 화면 준비 실패")
        return False

    def _build_address_search_query(address_basic: str) -> str:
        s = str(address_basic or "").strip()
        if not s:
            return ""

        s = re.sub(r"\([^)]*\)", "", s).strip()
        s = re.sub(r"\s+", " ", s).strip()

        if "," in s:
            s = s.split(",")[0].strip()

        if len(s) > 30:
            s = s[:30].rstrip()

        return s

    def _search_basic_address(address_basic: str) -> bool:
        if not address_basic:
            return _abort(f"전자서명 중단: 기본주소 값 없음 / 고객={job['name']}")

        if not _wait_address_search_ready(timeout_sec=10.0):
            return False

        edit = _find_top_widest_edittext(0.05, 0.25)
        if edit is None:
            return _abort(f"전자서명 중단: 주소 검색 입력칸 미검출 / 고객={job['name']}")

        search_query = _build_address_search_query(address_basic)
        if not search_query:
            return _abort(f"전자서명 중단: 주소 검색용 기본주소 가공 실패 / 고객={job['name']} / 기본주소={address_basic}")

        print(f"🔎 [ready_sign] 기본주소 검색: {search_query}")
        ok = type_into_edittext(d, edit, search_query)
        if not ok:
            return _abort(f"전자서명 중단: 기본주소 검색어 입력 실패 / 고객={job['name']} / 기본주소={search_query}")

        time.sleep(0.5)

        try:
            if d(text="검색").exists:
                click_text_center(d, "검색", 0.05, 0.25)
            elif d(textContains="검색").exists:
                d(textContains="검색").click()
            else:
                d.press("enter")
        except Exception:
            return _abort(f"전자서명 중단: 주소 검색 실행 실패 / 고객={job['name']}")

        time.sleep(3.0)
        return True

    def _wait_postcode_result_ready(timeout_sec: float = 12.0, expected_zipcode: str = "") -> bool:
        end_at = time.time() + timeout_sec
        expect_zip = normalize_digits(expected_zipcode or "")

        while time.time() < end_at:
            try:
                if is_unexpected_digital_sales_home(d):
                    return False

                has_title = d(text="우편번호찾기").exists or d(textContains="우편번호찾기").exists

                has_zip = False
                if expect_zip:
                    items = _collect_visible_text_items(0.05, 0.55)
                    for it in items:
                        if normalize_digits(it["text"]) == expect_zip:
                            has_zip = True
                            break

                if has_title or has_zip:
                    print("✅ [ready_sign] 우편번호 검색 결과 화면 준비 완료")
                    time.sleep(0.4)
                    return True
            except Exception:
                pass

            time.sleep(0.4)

        print("❌ [ready_sign] 우편번호 검색 결과 화면 준비 실패")
        return False

    def _normalize_addr_cmp(s: str) -> str:
        s = str(s or "")
        s = re.sub(r"\s+", "", s)
        s = re.sub(r"\([^)]*\)", "", s)
        s = s.replace(",", "")
        return s.strip()

    def _extract_addr_road_base(s: str) -> str:
        s = str(s or "").strip()
        if not s:
            return ""

        if "(" in s:
            s = s.split("(", 1)[0].strip()

        if "," in s:
            s = s.split(",", 1)[0].strip()

        s = re.sub(r"\s+", "", s)
        return s.strip()

    def _click_matching_postcode_result(zipcode: str, address_basic: str) -> bool:
        expected_zip = normalize_digits(zipcode or "")
        expected_addr = _normalize_addr_cmp(address_basic or "")
        expected_road = _extract_addr_road_base(address_basic or "")

        if not expected_zip:
            return _abort(f"전자서명 중단: 우편번호 값 없음 / 고객={job['name']}")

        if not expected_addr:
            return _abort(f"전자서명 중단: 기본주소 값 없음 / 고객={job['name']}")

        items = _collect_visible_text_items(0.05, 0.60)

        zip_found = False
        for it in items:
            if normalize_digits(it["text"]) == expected_zip:
                zip_found = True
                break

        if not zip_found:
            return _abort(f"전자서명 중단: 주소 검색 결과 우편번호 불일치 / 고객={job['name']} / 기대우편번호={expected_zip}")

        best_item = None
        best_score = 0

        for it in items:
            raw = str(it["text"] or "").strip()
            norm = _normalize_addr_cmp(raw)
            road = _extract_addr_road_base(raw)

            if not norm:
                continue

            if any(skip in norm for skip in ["우편번호", "도로명", "지번", "영문보기", "지도", "Poweredby", "카카오"]):
                continue

            score = 0

            if expected_road and road:
                if road == expected_road:
                    score = 120
                elif road.startswith(expected_road):
                    score = 110
                elif expected_road.startswith(road):
                    score = 100

            if score == 0:
                if norm == expected_addr:
                    score = 95
                elif norm.startswith(expected_addr):
                    score = 90
                elif expected_addr.startswith(norm):
                    score = 85

            if score > best_score:
                best_score = score
                best_item = it

        if best_item is None or best_score <= 0:
            return _abort(
                f"전자서명 중단: 주소 검색 결과 기본주소 불일치 / 고객={job['name']} / 기대주소={address_basic} / 기대우편번호={expected_zip}"
            )

        left, top, right, bottom = best_item["bounds"]
        cx = (left + right) // 2
        cy = (top + bottom) // 2

        d.click(cx, cy)
        time.sleep(2.0)
        print(f"✅ [ready_sign] 우편번호/기본주소 일치 결과 선택 완료: {best_item['text']}")
        return True

    def _find_focused_edittext():
        try:
            obj = d(focused=True, className="android.widget.EditText")
            if obj.exists:
                return obj
        except Exception:
            pass
        return None

    def _get_detail_submit_top():
        try:
            if d(text="주소 입력").exists:
                b = d(text="주소 입력").info.get("bounds", {})
                return int(b.get("top", 0))
            if d(textContains="주소 입력").exists:
                b = d(textContains="주소 입력").info.get("bounds", {})
                return int(b.get("top", 0))
        except Exception:
            pass
        return None

    def _get_address_result_bottom():
        result_bottom = 0
        for label in ["우편번호", "도로명", "지번"]:
            try:
                objs = d(text=label)
                cnt = objs.count
            except Exception:
                cnt = 0

            for i in range(cnt):
                try:
                    b = objs[i].info.get("bounds", {})
                    bottom = int(b.get("bottom", 0))
                    if bottom > result_bottom:
                        result_bottom = bottom
                except Exception:
                    continue
        return result_bottom

    def _get_edit_meta_text(edit_obj):
        try:
            info = edit_obj.info or {}
        except Exception:
            info = {}

        parts = [
            str(info.get("text") or ""),
            str(info.get("hintText") or ""),
            str(info.get("contentDescription") or ""),
            str(info.get("resourceName") or ""),
        ]
        return " ".join([p for p in parts if p]).strip()

    def _is_detail_address_candidate(edit_obj):
        try:
            w, h = d.window_size()
            info = edit_obj.info or {}
            b = info.get("bounds", {})

            left = int(b.get("left", 0))
            top = int(b.get("top", 0))
            right = int(b.get("right", 0))
            bottom = int(b.get("bottom", 0))
            width = right - left
            height = bottom - top
            cy = (top + bottom) // 2

            if left < 0 or top < 0 or right <= left or bottom <= top:
                return False

            if cy < int(h * 0.22) or cy > int(h * 0.72):
                return False

            if width < int(w * 0.55):
                return False

            if height < 45:
                return False

            result_bottom = _get_address_result_bottom()
            if result_bottom and top <= result_bottom + 12:
                return False

            submit_top = _get_detail_submit_top()
            if submit_top is not None and bottom >= submit_top:
                return False

            return True
        except Exception:
            return False

    def _find_detail_address_edittext():
        try:
            submit_top = _get_detail_submit_top()
            edits = d(className="android.widget.EditText")
            try:
                cnt = edits.count
            except Exception:
                cnt = 0

            best_hint = None
            best_hint_score = None
            best_geo = None
            best_geo_score = None

            for i in range(cnt):
                try:
                    e = edits[i]
                    if not _is_detail_address_candidate(e):
                        continue

                    info = e.info or {}
                    b = info.get("bounds", {})
                    bottom = int(b.get("bottom", 0))

                    score = 999999
                    if submit_top is not None:
                        score = abs(submit_top - bottom)

                    meta_norm = re.sub(r"\s+", "", _get_edit_meta_text(e))

                    if any(k in meta_norm for k in ["상세주소", "필수입력", "아파트명", "건물명"]):
                        if best_hint is None or score < best_hint_score:
                            best_hint = e
                            best_hint_score = score
                    else:
                        if best_geo is None or score < best_geo_score:
                            best_geo = e
                            best_geo_score = score
                except Exception:
                    continue

            if best_hint is not None:
                return best_hint
            return best_geo
        except Exception:
            return None

    def _handle_input_method_picker_if_open() -> bool:
        picker_open = False

        for kw in ["입력 방법 선택", "입력방법선택"]:
            try:
                if d(text=kw).exists or d(textContains=kw).exists:
                    picker_open = True
                    break
            except Exception:
                pass

        if not picker_open:
            return True

        print("⚠️ [ready_sign] 입력 방법 선택 팝업 감지 → AdbKeyboard 선택 시도")

        for kw in ["AdbKeyboard", "ADBKeyboard", "Adb Keyboard"]:
            try:
                if d(text=kw).exists:
                    d(text=kw).click()
                    time.sleep(0.8)
                    return True
            except Exception:
                pass

            try:
                if d(textContains=kw).exists:
                    d(textContains=kw).click()
                    time.sleep(0.8)
                    return True
            except Exception:
                pass

        try:
            d.press("back")
            time.sleep(0.5)
        except Exception:
            pass

        print("⚠️ [ready_sign] 입력 방법 선택 팝업에서 AdbKeyboard 미검출 → 기본 입력 방식 유지")
        return True

    def _same_edit_bounds(a, b) -> bool:
        if a is None or b is None:
            return False

        try:
            ba = a.info.get("bounds", {})
            bb = b.info.get("bounds", {})

            return (
                int(ba.get("left", -1)) == int(bb.get("left", -2))
                and int(ba.get("top", -1)) == int(bb.get("top", -2))
                and int(ba.get("right", -1)) == int(bb.get("right", -2))
                and int(ba.get("bottom", -1)) == int(bb.get("bottom", -2))
            )
        except Exception:
            return False

    def _normalize_detail_value_for_check(s: str) -> str:
        return re.sub(r"\s+", "", str(s or "")).strip()

    def _click_detail_address_field(edit=None) -> bool:
        if edit is None:
            edit = _find_detail_address_edittext()

        try:
            w, h = d.window_size()
        except Exception:
            w, h = (1080, 1920)

        click_points = []
        tried = set()

        def _add_point(x, y):
            try:
                x = max(5, min(w - 5, int(x)))
                y = max(5, min(h - 5, int(y)))
                key = (x, y)
                if key not in tried:
                    tried.add(key)
                    click_points.append((x, y))
            except Exception:
                pass

        if edit is not None:
            try:
                b = edit.info.get("bounds", {})
                left = int(b.get("left", 0))
                top = int(b.get("top", 0))
                right = int(b.get("right", 0))
                bottom = int(b.get("bottom", 0))

                width = max(1, right - left)
                height = max(1, bottom - top)
                cy = (top + bottom) // 2

                _add_point((left + right) // 2, cy)
                _add_point(left + max(30, int(width * 0.12)), cy)
                _add_point(left + max(70, int(width * 0.28)), cy)
                _add_point(left + max(110, int(width * 0.42)), cy)
                _add_point((left + right) // 2, top + max(12, int(height * 0.35)))
            except Exception:
                pass

        try:
            result_bottom = _get_address_result_bottom()
        except Exception:
            result_bottom = 0

        try:
            submit_top = _get_detail_submit_top()
        except Exception:
            submit_top = None

        if submit_top is not None and submit_top > 0:
            if result_bottom <= 0:
                result_bottom = int(h * 0.24)

            field_top = max(result_bottom + 18, int(h * 0.26))
            field_bottom = min(submit_top - 18, int(h * 0.62))

            if field_bottom > field_top:
                mid_y = (field_top + field_bottom) // 2

                _add_point(int(w * 0.50), mid_y)
                _add_point(int(w * 0.22), mid_y)
                _add_point(int(w * 0.35), mid_y)
                _add_point(int(w * 0.65), mid_y)
                _add_point(int(w * 0.50), field_top + max(8, int((field_bottom - field_top) * 0.30)))

        if not click_points:
            print("⚠️ [detail_address] 상세주소칸 클릭 좌표 생성 실패")
            return False

        for idx, (x, y) in enumerate(click_points, start=1):
            try:
                d.click(x, y)
                time.sleep(0.28)

                focused = _find_focused_edittext()
                if focused is not None:
                    if edit is not None and _same_edit_bounds(focused, edit):
                        print(f"✅ [detail_address] 상세주소칸 포커스 확인 완료: ({x}, {y})")
                        return True

                    try:
                        if _is_detail_address_candidate(focused):
                            print(f"✅ [detail_address] 상세주소 후보 포커스 전환 확인: ({x}, {y})")
                            return True
                    except Exception:
                        pass

                if idx == 1:
                    try:
                        d.long_click(x, y, 0.35)
                        time.sleep(0.22)
                        focused = _find_focused_edittext()
                        if focused is not None:
                            if edit is not None and _same_edit_bounds(focused, edit):
                                print(f"✅ [detail_address] 상세주소칸 롱클릭 포커스 확인 완료: ({x}, {y})")
                                return True
                            try:
                                if _is_detail_address_candidate(focused):
                                    print(f"✅ [detail_address] 상세주소 후보 롱클릭 포커스 전환 확인: ({x}, {y})")
                                    return True
                            except Exception:
                                pass
                    except Exception:
                        pass

            except Exception:
                continue

        print("⚠️ [detail_address] 포커스 확인은 실패했지만 상세주소칸 좌표 클릭은 수행함")
        return True

    def _click_address_submit_button() -> bool:
        try:
            if d(text="주소 입력").exists:
                b = d(text="주소 입력").info.get("bounds", {})
                left = int(b.get("left", 0))
                top = int(b.get("top", 0))
                right = int(b.get("right", 0))
                bottom = int(b.get("bottom", 0))
                cx = (left + right) // 2
                cy = (top + bottom) // 2
                d.click(cx, cy)
                time.sleep(1.2)
                return True
        except Exception:
            pass

        try:
            if d(textContains="주소 입력").exists:
                b = d(textContains="주소 입력").info.get("bounds", {})
                left = int(b.get("left", 0))
                top = int(b.get("top", 0))
                right = int(b.get("right", 0))
                bottom = int(b.get("bottom", 0))
                cx = (left + right) // 2
                cy = (top + bottom) // 2
                d.click(cx, cy)
                time.sleep(1.2)
                return True
        except Exception:
            pass

        return False

    def _is_address_detail_screen_now() -> bool:
        try:
            has_title = d(text="주소입력").exists or d(textContains="주소입력").exists
        except Exception:
            has_title = False

        try:
            has_submit = d(text="주소 입력").exists or d(textContains="주소 입력").exists
        except Exception:
            has_submit = False

        detail_edit = _find_detail_address_edittext()
        return has_title and has_submit and detail_edit is not None

    def _wait_address_detail_ready(timeout_sec: float = 12.0) -> bool:
        end_at = time.time() + timeout_sec
        stable_ok_count = 0

        while time.time() < end_at:
            try:
                if is_unexpected_digital_sales_home(d):
                    return False

                if _is_address_detail_screen_now():
                    stable_ok_count += 1
                    if stable_ok_count >= 2:
                        print("✅ [ready_sign] 주소 상세입력 화면 준비 완료")
                        time.sleep(0.4)
                        return True
                else:
                    stable_ok_count = 0
            except Exception:
                stable_ok_count = 0

            time.sleep(0.4)

        print("❌ [ready_sign] 주소 상세입력 화면 준비 실패")
        return False

    def _fill_detail_address_and_submit(address_detail: str) -> bool:
        if not address_detail:
            return _abort(f"전자서명 중단: 상세주소 값 없음 / 고객={job['name']}")

        if not _wait_address_detail_ready(timeout_sec=12.0):
            return False

        expected_raw = str(address_detail).strip()
        expected_norm = _normalize_detail_value_for_check(expected_raw)
        expected_digits = normalize_digits(expected_raw)

        for attempt in range(3):
            candidate = _find_detail_address_edittext()
            if candidate is None:
                return _abort(f"전자서명 중단: 상세주소 입력칸 미검출 / 고객={job['name']}")

            clicked = _click_detail_address_field(candidate)
            if not clicked:
                print("⚠️ [detail_address] 상세주소칸 클릭 확정 실패 → 그래도 입력 시도 진행")
                if attempt == 2:
                    return _abort(f"전자서명 중단: 상세주소칸 클릭 실패 / 고객={job['name']}")
                time.sleep(0.5)

            if not _handle_input_method_picker_if_open():
                if attempt == 2:
                    return _abort(f"전자서명 중단: AdbKeyboard 선택 실패 / 고객={job['name']}")
                time.sleep(0.8)
                continue

            focused = _find_focused_edittext()
            target_edit = None

            if focused is not None:
                if _same_edit_bounds(focused, candidate):
                    target_edit = focused
                else:
                    try:
                        if _is_detail_address_candidate(focused):
                            target_edit = focused
                    except Exception:
                        target_edit = None

            if target_edit is None:
                target_edit = candidate

            current_before = _read_edit_obj_text(target_edit)
            print(f"🔍 [detail_address] 입력 전 현재값: '{current_before}'")

            ok_input = type_into_edittext(d, target_edit, expected_raw)

            refind = _find_focused_edittext()
            if refind is None:
                refind = _find_detail_address_edittext()
            if refind is None:
                refind = target_edit

            current_text = _read_edit_obj_text(refind)
            current_norm = _normalize_detail_value_for_check(current_text)
            current_digits = normalize_digits(current_text)

            print(f"🔍 [detail_address] type_into_edittext 결과: '{current_text}'")

            value_ok = False
            if current_norm == expected_norm:
                value_ok = True
            elif expected_digits and current_digits == expected_digits:
                value_ok = True

            if (not ok_input) or (not value_ok):
                print("⚠️ [detail_address] 일반 입력 검증 실패 → AdbKeyboard fallback 시도")

                reclicked = _click_detail_address_field(refind if refind is not None else candidate)
                if not reclicked:
                    print("⚠️ [detail_address] 상세주소칸 재포커스 확정 실패 → ADB fallback은 계속 시도")
                    if attempt < 2:
                        time.sleep(0.4)

                adb_keyboard_clear_text()
                time.sleep(0.2)

                ok_adb = adb_keyboard_input_text(expected_raw)
                if not ok_adb:
                    if attempt == 2:
                        return _abort(f"전자서명 중단: 상세주소 입력 실패 / 고객={job['name']} / 기대={expected_raw}")
                    time.sleep(0.8)
                    continue

                time.sleep(0.6)

                refind = _find_focused_edittext()
                if refind is None:
                    refind = _find_detail_address_edittext()
                if refind is None:
                    refind = candidate

                current_text = _read_edit_obj_text(refind)
                current_norm = _normalize_detail_value_for_check(current_text)
                current_digits = normalize_digits(current_text)

                print(f"🔍 [detail_address] adb_keyboard 결과: '{current_text}'")

                value_ok = False
                if current_norm == expected_norm:
                    value_ok = True
                elif expected_digits and current_digits == expected_digits:
                    value_ok = True

            if not value_ok:
                if attempt == 2:
                    return _abort(
                        f"전자서명 중단: 상세주소 입력값 불일치 / 고객={job['name']} / 기대={expected_raw} / 현재={current_text}"
                    )
                print(f"⚠️ [ready_sign] 상세주소 입력값 불일치 → 재시도 {attempt + 1}/3")
                time.sleep(0.8)
                continue

            ok_click = _click_address_submit_button()
            if not ok_click:
                if attempt == 2:
                    return _abort(f"전자서명 중단: 상세주소 주소입력 버튼 클릭 실패 / 고객={job['name']}")
                time.sleep(0.8)
                continue

            time.sleep(1.5)

            if not _is_address_detail_screen_now():
                if _wait_install_info_ready(timeout_sec=4.0):
                    print(f"✅ [ready_sign] 상세주소 입력 및 주소 등록 완료: {address_detail}")
                    return True

            if _is_address_detail_screen_now():
                print(f"⚠️ [ready_sign] 상세주소 입력 후 아직 같은 화면 → 재시도 {attempt + 1}/3")
                time.sleep(1.0)
                continue

            if attempt == 2:
                return _abort(f"전자서명 중단: 상세주소 입력 후 화면 전환 실패 / 고객={job['name']}")

        return _abort(f"전자서명 중단: 상세주소 입력 처리 최종 실패 / 고객={job['name']}")

    def _is_saved_address_confirm_popup_open() -> bool:
        popup_fragments = [
            "선택한 주소를 다시 한번 확인해 주시기 바랍니다",
            "선택한주소를다시한번확인해주시기바랍니다",
            "해당 주소로 주문 및 설치를 진행하시겠습니까",
            "해당주소로주문및설치를진행하시겠습니까",
        ]

        texts = _collect_popup_texts(0.20, 0.86)
        joined = " ".join(texts)

        if not joined:
            return False

        for frag in popup_fragments:
            if re.sub(r"\s+", "", frag) in joined:
                return True

        return False

    def _click_saved_new_address_card() -> bool:
        label_obj = None

        try:
            if d(text="신규주소지").exists:
                label_obj = d(text="신규주소지")
            elif d(textContains="신규주소지").exists:
                label_obj = d(textContains="신규주소지")
        except Exception:
            label_obj = None

        if label_obj is None:
            return False

        try:
            w, h = d.window_size()
            b = label_obj.info.get("bounds", {})
            left = int(b.get("left", 0))
            top = int(b.get("top", 0))
            right = int(b.get("right", 0))
            bottom = int(b.get("bottom", 0))

            card_y = min(h - 10, max(10, (top + bottom) // 2 + 26))
            candidate_points = [
                (int(w * 0.50), card_y),
                (int(w * 0.38), card_y),
                (int(w * 0.65), card_y),
                (max(20, left + 120), card_y),
                (max(20, right + 80), card_y),
            ]

            tried = set()
            filtered_points = []
            for x, y in candidate_points:
                x = max(5, min(w - 5, int(x)))
                y = max(5, min(h - 5, int(y)))
                key = (x, y)
                if key not in tried:
                    filtered_points.append((x, y))
                    tried.add(key)

            for x, y in filtered_points:
                d.click(x, y)
                time.sleep(0.8)

                if _is_saved_address_confirm_popup_open():
                    print(f"✅ [ready_sign] 신규주소지 카드 클릭 완료: ({x}, {y})")
                    return True

            print("⚠️ [ready_sign] 신규주소지 텍스트는 찾았지만 확인 팝업이 안 뜸")
            return False

        except Exception as e:
            print("⚠️ [ready_sign] 신규주소지 카드 클릭 예외:", e)
            return False

    def _handle_saved_address_selection_if_needed() -> bool:
        for attempt in range(4):
            if _is_saved_address_confirm_popup_open():
                print("✅ [ready_sign] 신규주소지 확인 팝업 감지")
                if not _click_confirm_if_exists():
                    if attempt == 3:
                        return _abort(f"전자서명 중단: 신규주소지 확인 팝업 확인 버튼 클릭 실패 / 고객={job['name']}")
                    time.sleep(0.8)
                    continue

                time.sleep(1.5)

                if _wait_install_info_ready(timeout_sec=6.0):
                    print("✅ [ready_sign] 신규주소지 확인 완료 후 설치정보 화면 복귀")
                    return True

            if _wait_install_info_ready(timeout_sec=1.2):
                return True

            if not _wait_address_management_ready(timeout_sec=2.5):
                if attempt == 3:
                    return _abort(f"전자서명 중단: 주소지 관리/설치정보 화면 판정 실패 / 고객={job['name']}")
                time.sleep(0.6)
                continue

            clicked = _click_saved_new_address_card()
            if not clicked:
                if attempt == 3:
                    return _abort(f"전자서명 중단: 신규주소지 선택 실패 / 고객={job['name']}")
                time.sleep(0.8)
                continue

            time.sleep(0.8)

            if _is_saved_address_confirm_popup_open():
                print("✅ [ready_sign] 신규주소지 확인 팝업 감지")
                if not _click_confirm_if_exists():
                    if attempt == 3:
                        return _abort(f"전자서명 중단: 신규주소지 확인 팝업 확인 버튼 클릭 실패 / 고객={job['name']}")
                    time.sleep(0.8)
                    continue

                time.sleep(1.5)

                if _wait_install_info_ready(timeout_sec=6.0):
                    print("✅ [ready_sign] 신규주소지 선택/확인 완료 후 설치정보 화면 복귀")
                    return True

            if attempt < 3:
                time.sleep(0.8)

        return _abort(f"전자서명 중단: 신규주소지 선택 후 설치정보 화면 복귀 실패 / 고객={job['name']}")

    def _fill_install_home_phone(phone11: str) -> bool:
        if not phone11:
            return _abort(f"전자서명 중단: 전화번호 값 없음 / 고객={job['name']}")

        def _find_install_home_phone_edit():
            try:
                w, h = d.window_size()
            except Exception:
                w, h = (1080, 1920)

            best = None
            best_score = None

            try:
                edits = d(className="android.widget.EditText")
                cnt = edits.count
            except Exception:
                cnt = 0

            for i in range(cnt):
                try:
                    e = edits[i]
                    info = e.info or {}
                    b = info.get("bounds", {})

                    left = int(b.get("left", 0))
                    top = int(b.get("top", 0))
                    right = int(b.get("right", 0))
                    bottom = int(b.get("bottom", 0))
                    width = right - left

                    if top < int(h * 0.22) or bottom > int(h * 0.60):
                        continue

                    if width < int(w * 0.45):
                        continue

                    meta = _get_edit_meta_text(e)
                    meta_norm = re.sub(r"\s+", "", str(meta or ""))

                    if "자택번호를입력해주세요" in meta_norm or "선택" in meta_norm:
                        score = top
                        if best is None or score < best_score:
                            best = e
                            best_score = score
                        continue

                    current_text = _read_edit_obj_text(e)
                    current_digits = normalize_digits(current_text)

                    if current_digits:
                        continue

                    if "휴대폰번호" in meta_norm:
                        continue

                    score = 100000 + top
                    if best is None or score < best_score:
                        best = e
                        best_score = score

                except Exception:
                    continue

            return best

        edit = _find_install_home_phone_edit()
        if edit is None:
            return _abort(f"전자서명 중단: 전화번호 입력칸 미검출 / 고객={job['name']}")

        before_text = _read_edit_obj_text(edit)
        before_digits = normalize_digits(before_text)
        print(f"🔍 [ready_sign] 전화번호 입력 전 현재값: '{before_text}'")

        if before_digits and before_digits == phone11:
            print(f"✅ [ready_sign] 전화번호 이미 입력 일치: {phone11}")
            return True

        ok = type_into_edittext(d, edit, phone11)
        if not ok:
            return _abort(f"전자서명 중단: 전화번호 입력 실패 / 고객={job['name']} / 전화번호={phone11}")

        time.sleep(0.8)

        current_text = _read_edit_obj_text(edit)
        current_digits = normalize_digits(current_text)
        print(f"🔍 [ready_sign] 전화번호 입력 후 현재값: '{current_text}'")

        if current_digits != phone11:
            return _abort(f"전자서명 중단: 전화번호 입력값 불일치 / 고객={job['name']} / 기대={phone11} / 현재={current_digits}")

        print(f"✅ [ready_sign] 전화번호 입력 완료: {phone11}")
        return True

    def _wait_install_place_category_sheet_ready(timeout_sec: float = 6.0) -> bool:
        end_at = time.time() + timeout_sec

        while time.time() < end_at:
            try:
                if is_unexpected_digital_sales_home(d):
                    return False

                has_title = d(text="설치처 분류").exists or d(textContains="설치처 분류").exists
                has_home = d(text="가정").exists or d(textContains="가정").exists

                if has_title and has_home:
                    print("✅ [ready_sign] 설치처 분류 선택 시트 준비 완료")
                    time.sleep(0.3)
                    return True
            except Exception:
                pass

            time.sleep(0.25)

        print("❌ [ready_sign] 설치처 분류 선택 시트 준비 실패")
        return False

    def _open_install_place_category_selector() -> bool:
        try:
            if d(text="설치처 분류 선택").exists:
                ok = click_text_center(d, "설치처 분류 선택", 0.35, 0.78)
                time.sleep(1.2)
                return ok
        except Exception:
            pass

        try:
            if d(textContains="설치처 분류 선택").exists:
                d(textContains="설치처 분류 선택").click()
                time.sleep(1.2)
                return True
        except Exception:
            pass

        label_obj = None

        try:
            if d(text="설치처 분류").exists:
                label_obj = d(text="설치처 분류")
            elif d(textContains="설치처 분류").exists:
                label_obj = d(textContains="설치처 분류")
        except Exception:
            label_obj = None

        if label_obj is not None:
            try:
                w, h = d.window_size()
                b = label_obj.info.get("bounds", {})
                bottom = int(b.get("bottom", 0))
                field_y = min(h - 10, bottom + max(52, int(h * 0.03)))

                candidate_points = [
                    (int(w * 0.50), field_y),
                    (int(w * 0.72), field_y),
                    (int(w * 0.85), field_y),
                ]

                for x, y in candidate_points:
                    d.click(x, y)
                    time.sleep(1.0)
                    if _wait_install_place_category_sheet_ready(timeout_sec=1.5):
                        return True
            except Exception:
                pass

        return False

    def _select_install_place_home() -> bool:
        for attempt in range(3):
            if not _open_install_place_category_selector():
                if attempt == 2:
                    return _abort(f"전자서명 중단: 설치처 분류 선택 열기 실패 / 고객={job['name']}")
                time.sleep(0.8)
                continue

            if not _wait_install_place_category_sheet_ready(timeout_sec=4.0):
                if attempt == 2:
                    return _abort(f"전자서명 중단: 설치처 분류 선택 시트 미검출 / 고객={job['name']}")
                time.sleep(0.8)
                continue

            clicked = False

            try:
                if d(text="가정").exists:
                    d(text="가정").click()
                    clicked = True
            except Exception:
                clicked = False

            if not clicked:
                try:
                    if d(textContains="가정").exists:
                        d(textContains="가정").click()
                        clicked = True
                except Exception:
                    clicked = False

            if not clicked:
                try:
                    clicked = click_text_center(d, "가정", 0.58, 0.96)
                except Exception:
                    clicked = False

            if not clicked:
                if attempt == 2:
                    return _abort(f"전자서명 중단: 설치처 분류 '가정' 클릭 실패 / 고객={job['name']}")
                time.sleep(0.8)
                continue

            time.sleep(1.2)

            try:
                if d(text="가정").exists or d(textContains="가정").exists:
                    print("✅ [ready_sign] 설치처 분류 선택 완료: 가정")
                    return True
            except Exception:
                pass

            if _wait_install_info_ready(timeout_sec=2.0):
                print("✅ [ready_sign] 설치처 분류 선택 완료: 가정")
                return True

        return _abort(f"전자서명 중단: 설치처 분류 '가정' 선택 처리 실패 / 고객={job['name']}")

    def _needs_fast_install_request(note_text: str) -> bool:
        s = re.sub(r"\s+", "", str(note_text or ""))
        return ("가장빠른설치요청" in s) or ("빠른설치요청" in s)

    def _open_install_datetime_selector() -> bool:
        try:
            if d(text="설치일시 선택").exists:
                ok = click_text_center(d, "설치일시 선택", 0.35, 0.80)
                time.sleep(1.2)
                return ok
        except Exception:
            pass

        try:
            if d(textContains="설치일시 선택").exists:
                d(textContains="설치일시 선택").click()
                time.sleep(1.2)
                return True
        except Exception:
            pass

        label_obj = None
        try:
            if d(text="설치희망일시").exists:
                label_obj = d(text="설치희망일시")
            elif d(textContains="설치희망일시").exists:
                label_obj = d(textContains="설치희망일시")
        except Exception:
            label_obj = None

        if label_obj is not None:
            try:
                w, h = d.window_size()
                b = label_obj.info.get("bounds", {})
                bottom = int(b.get("bottom", 0))
                field_y = min(h - 10, bottom + max(50, int(h * 0.03)))
                for x in [int(w * 0.50), int(w * 0.74), int(w * 0.86)]:
                    d.click(x, field_y)
                    time.sleep(1.0)
                    if d(textContains="방문 희망일시 선택").exists or d(text="선택완료").exists:
                        return True
            except Exception:
                pass

        return False

    def _wait_fast_install_sheet_ready(timeout_sec: float = 6.0) -> bool:
        end_at = time.time() + timeout_sec
        while time.time() < end_at:
            try:
                if is_unexpected_digital_sales_home(d):
                    return False

                has_title = d(text="방문 희망일시 선택").exists or d(textContains="방문 희망일시 선택").exists
                has_more = d(text="더보기").exists or d(textContains="더보기").exists
                has_deadline = d(text="마감").exists or d(textContains="마감").exists
                has_holiday = d(text="휴일").exists or d(textContains="휴일").exists

                has_date_like = False
                items = _collect_visible_text_items(0.08, 0.72)
                for it in items:
                    txt = str(it.get("text") or "").strip()
                    if re.fullmatch(r"\d{1,2}\.\d{1,2}", txt):
                        has_date_like = True
                        break

                if has_title and (has_more or has_deadline or has_holiday or has_date_like):
                    print("✅ [ready_sign] 방문 희망일시 선택 시트 준비 완료")
                    time.sleep(0.3)
                    return True
            except Exception:
                pass

            time.sleep(0.25)

        print("❌ [ready_sign] 방문 희망일시 선택 시트 준비 실패")
        return False

    def _pick_earliest_visit_slot_and_done() -> bool:
        if not _open_install_datetime_selector():
            return _abort(f"전자서명 중단: 설치희망일시 선택 열기 실패 / 고객={job['name']}")

        print("⏳ [ready_sign] 방문 희망일시 배정판 로딩 대기 5.0초")
        time.sleep(5.0)

        if not _wait_fast_install_sheet_ready(timeout_sec=5.0):
            return _abort(f"전자서명 중단: 방문 희망일시 선택 시트 미검출 / 고객={job['name']}")

        items = _collect_visible_text_items(0.10, 0.74)

        slot_candidates = []
        for it in items:
            txt = str(it["text"] or "").strip()
            if txt not in ["○", "◯"]:
                continue

            left, top, right, bottom = it["bounds"]

            slot_candidates.append({
                "text": txt,
                "bounds": (left, top, right, bottom),
            })

        slot_candidates.sort(key=lambda x: (x["bounds"][1], x["bounds"][0]))

        if not slot_candidates:
            return _abort(f"전자서명 중단: 방문 희망일시 선택 가능 슬롯 미검출 / 고객={job['name']}")

        chosen = slot_candidates[0]
        left, top, right, bottom = chosen["bounds"]
        cx = (left + right) // 2
        cy = (top + bottom) // 2
        d.click(cx, cy)
        time.sleep(1.0)

        clicked_done = False

        for _ in range(8):
            try:
                if d(text="선택완료").exists:
                    d(text="선택완료").click()
                    clicked_done = True
                    break
            except Exception:
                pass

            try:
                if d(textContains="선택완료").exists:
                    d(textContains="선택완료").click()
                    clicked_done = True
                    break
            except Exception:
                pass

            try:
                if click_text_center(d, "선택완료", 0.82, 0.99):
                    clicked_done = True
                    break
            except Exception:
                pass

            time.sleep(0.3)

        if not clicked_done:
            return _abort(f"전자서명 중단: 방문 희망일시 선택완료 클릭 실패 / 고객={job['name']}")

        time.sleep(1.2)

        if not _wait_install_info_ready(timeout_sec=6.0):
            return _abort(f"전자서명 중단: 설치희망일시 선택 후 설치정보 화면 복귀 실패 / 고객={job['name']}")

        print("✅ [ready_sign] 가장 빠른 설치희망일시 선택 완료")
        return True

    def _open_install_env_info() -> bool:
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
                    if d(textContains="타사제품 반환여부").exists or d(text="입력완료").exists:
                        return True
            except Exception:
                pass

        return False

    def _wait_install_env_ready(timeout_sec: float = 6.0) -> bool:
        end_at = time.time() + timeout_sec
        while time.time() < end_at:
            try:
                if is_unexpected_digital_sales_home(d):
                    return False
                has_done = d(text="입력완료").exists or d(textContains="입력완료").exists
                has_return = d(textContains="타사제품 반환여부").exists or d(text="미반환").exists or d(text="반환").exists
                has_multi = d(textContains="다중시설").exists
                if has_done and has_return and has_multi:
                    print("✅ [ready_sign] 설치환경정보 화면 준비 완료")
                    time.sleep(0.3)
                    return True
            except Exception:
                pass
            time.sleep(0.25)
        print("❌ [ready_sign] 설치환경정보 화면 준비 실패")
        return False

    def _wait_multifacility_sheet_ready(timeout_sec: float = 5.0) -> bool:
        end_at = time.time() + timeout_sec
        while time.time() < end_at:
            try:
                has_title = d(text="다중시설 선택").exists or d(textContains="다중시설 선택").exists
                has_target = d(text="대상 아님").exists or d(textContains="대상 아님").exists
                if has_title and has_target:
                    print("✅ [ready_sign] 다중시설 선택 시트 준비 완료")
                    time.sleep(0.2)
                    return True
            except Exception:
                pass
            time.sleep(0.2)
        print("❌ [ready_sign] 다중시설 선택 시트 준비 실패")
        return False

    def _fill_install_env_info(pickup_request_raw: str) -> bool:
        if not _open_install_env_info():
            return _abort(f"전자서명 중단: 설치환경정보 열기 실패 / 고객={job['name']}")

        if not _wait_install_env_ready(timeout_sec=5.0):
            return _abort(f"전자서명 중단: 설치환경정보 화면 미검출 / 고객={job['name']}")

        if str(pickup_request_raw or "").strip():
            clicked_return = False
            try:
                if d(text="반환").exists:
                    d(text="반환").click()
                    clicked_return = True
            except Exception:
                clicked_return = False

            if not clicked_return:
                try:
                    if d(textContains="반환").exists:
                        d(textContains="반환").click()
                        clicked_return = True
                except Exception:
                    clicked_return = False

            if not clicked_return:
                return _abort(f"전자서명 중단: 타사제품 반환여부 '반환' 클릭 실패 / 고객={job['name']}")

            time.sleep(0.6)
            print("✅ [ready_sign] 타사제품 반환여부 선택 완료: 반환")

        opened_multi = False
        try:
            if d(text="다중시설 선택").exists:
                d(text="다중시설 선택").click()
                opened_multi = True
        except Exception:
            opened_multi = False

        if not opened_multi:
            try:
                if d(textContains="다중시설 선택").exists:
                    d(textContains="다중시설 선택").click()
                    opened_multi = True
            except Exception:
                opened_multi = False

        if not opened_multi:
            return _abort(f"전자서명 중단: 다중시설 선택 열기 실패 / 고객={job['name']}")

        time.sleep(0.8)

        if not _wait_multifacility_sheet_ready(timeout_sec=4.0):
            return _abort(f"전자서명 중단: 다중시설 선택 시트 미검출 / 고객={job['name']}")

        clicked_target = False
        try:
            if d(text="대상 아님").exists:
                d(text="대상 아님").click()
                clicked_target = True
        except Exception:
            clicked_target = False

        if not clicked_target:
            try:
                if d(textContains="대상 아님").exists:
                    d(textContains="대상 아님").click()
                    clicked_target = True
            except Exception:
                clicked_target = False

        if not clicked_target:
            return _abort(f"전자서명 중단: 다중시설 '대상 아님' 클릭 실패 / 고객={job['name']}")

        time.sleep(0.8)

        clicked_done = False
        try:
            if d(text="입력완료").exists:
                d(text="입력완료").click()
                clicked_done = True
        except Exception:
            clicked_done = False

        if not clicked_done:
            try:
                if d(textContains="입력완료").exists:
                    d(textContains="입력완료").click()
                    clicked_done = True
            except Exception:
                clicked_done = False

        if not clicked_done:
            try:
                clicked_done = click_text_center(d, "입력완료", 0.85, 0.99)
            except Exception:
                clicked_done = False

        if not clicked_done:
            return _abort(f"전자서명 중단: 설치환경정보 입력완료 클릭 실패 / 고객={job['name']}")

        time.sleep(1.2)

        if not _wait_install_info_ready(timeout_sec=6.0):
            return _abort(f"전자서명 중단: 설치환경정보 입력 후 설치정보 화면 복귀 실패 / 고객={job['name']}")

        print("✅ [ready_sign] 설치환경정보 입력 완료")
        return True

    def _fill_install_memo(note_text: str) -> bool:
        note_text = str(note_text or "").strip()
        if not note_text:
            print("ℹ️ [ready_sign] 설치메모 입력 생략(특이사항 없음)")
            return True

        edit = _find_edittext_near_label(["설치메모"], y_tolerance=100, y_below=260)
        if edit is None:
            return _abort(f"전자서명 중단: 설치메모 입력칸 미검출 / 고객={job['name']}")

        ok = type_into_edittext(d, edit, note_text)
        if not ok:
            return _abort(f"전자서명 중단: 설치메모 입력 실패 / 고객={job['name']}")

        time.sleep(0.6)

        current_text = _read_edit_obj_text(edit)
        current_norm = re.sub(r"\s+", "", str(current_text or ""))
        expect_norm = re.sub(r"\s+", "", note_text)

        if current_norm != expect_norm:
            return _abort(
                f"전자서명 중단: 설치메모 입력값 불일치 / 고객={job['name']} / 기대={note_text} / 현재={current_text}"
            )

        print("✅ [ready_sign] 설치메모 입력 완료")
        return True

    def _fill_install_info_address(
        phone11: str,
        zipcode: str,
        address_basic: str,
        address_detail: str,
        pickup_request_raw: str,
        special_note_raw: str,
    ) -> bool:
        if not _wait_install_info_ready(timeout_sec=12.0):
            return _abort(f"전자서명 중단: 설치정보 화면 진입 실패 / 고객={job['name']}")

        _dismiss_existing_install_info_popup_if_open()

        if not _fill_install_phone(phone11):
            return False

        if not _click_install_address_input_button():
            return _abort(f"전자서명 중단: 설치정보 주소 입력 버튼 클릭 실패 / 고객={job['name']}")

        if not _wait_address_management_ready(timeout_sec=10.0):
            return _abort(f"전자서명 중단: 주소지 관리 화면 진입 실패 / 고객={job['name']}")

        if not _click_new_address_add():
            return _abort(f"전자서명 중단: 새 주소지 추가 버튼 클릭 실패 / 고객={job['name']}")

        if not _search_basic_address(address_basic):
            return False

        if not _wait_postcode_result_ready(timeout_sec=12.0, expected_zipcode=zipcode):
            return _abort(f"전자서명 중단: 우편번호 검색 결과 화면 진입 실패 / 고객={job['name']}")

        if not _click_matching_postcode_result(zipcode, address_basic):
            return False

        if not _fill_detail_address_and_submit(address_detail):
            return False

        if not _handle_saved_address_selection_if_needed():
            return False

        if not _wait_install_info_ready(timeout_sec=12.0):
            return _abort(f"전자서명 중단: 주소 입력 후 설치정보 화면 복귀 실패 / 고객={job['name']}")

        if not _fill_install_home_phone(phone11):
            return False

        if not _select_install_place_home():
            return False

        if not _wait_install_info_ready(timeout_sec=12.0):
            return _abort(f"전자서명 중단: 설치처 분류 선택 후 설치정보 화면 확인 실패 / 고객={job['name']}")

        if _needs_fast_install_request(special_note_raw):
            if not _pick_earliest_visit_slot_and_done():
                return False

        if not _fill_install_env_info(pickup_request_raw):
            return False

        if not _fill_install_memo(special_note_raw):
            return False

        if not _wait_install_info_ready(timeout_sec=12.0):
            return _abort(f"전자서명 중단: 설치메모 입력 후 설치정보 화면 확인 실패 / 고객={job['name']}")

        print(f"✅ [ready_sign] 설치정보 주소/전화번호/설치처분류/설치희망일시/설치환경/설치메모 입력 완료: {job['name']}")
        return True

    def _try_resume_from_current_screen(
        account_raw: str,
        discount_raw: str,
        amount_raw: str,
        phone11: str,
        zipcode: str,
        address_basic: str,
        address_detail: str,
        pickup_request_raw: str,
        special_note_raw: str,
    ):
        if _wait_install_info_ready(timeout_sec=3.0):
            print(f"✅ [ready_sign] 주문 이어서 하기 후 설치정보 화면 진입 감지: {job['name']}")
            return _fill_install_info_address(
                phone11,
                zipcode,
                address_basic,
                address_detail,
                pickup_request_raw,
                special_note_raw,
            )

        if _wait_payment_info_ready(timeout_sec=3.0):
            print(f"✅ [ready_sign] 주문 이어서 하기 후 결제정보 선택 화면 진입 감지: {job['name']}")

            if not account_raw:
                return _abort(f"전자서명 중단: 모달 결제정보 추출 실패 / 고객={job['name']}")

            if not _open_payment_method_and_click_add(account_raw):
                return False

            return _fill_install_info_address(
                phone11,
                zipcode,
                address_basic,
                address_detail,
                pickup_request_raw,
                special_note_raw,
            )

        if _wait_discount_page_ready(timeout_sec=3.0):
            print(f"✅ [ready_sign] 주문 이어서 하기 후 할인정보 입력 화면 진입 감지: {job['name']}")

            if not discount_raw:
                return _abort(f"전자서명 중단: 모달 할인 추출 실패 / 고객={job['name']}")

            if not amount_raw:
                return _abort(f"전자서명 중단: 모달 금액 추출 실패 / 고객={job['name']}")

            if not account_raw:
                return _abort(f"전자서명 중단: 모달 결제정보 추출 실패 / 고객={job['name']}")

            if not _verify_discount_and_amount_then_next(discount_raw, amount_raw):
                return False

            if not _open_payment_method_and_click_add(account_raw):
                return False

            return _fill_install_info_address(
                phone11,
                zipcode,
                address_basic,
                address_detail,
                pickup_request_raw,
                special_note_raw,
            )

        return None
    
    def _open_payment_method_and_click_add(account_raw: str) -> bool:
        if not _wait_payment_info_ready(timeout_sec=12.0):
            return _abort(f"전자서명 중단: 결제정보 선택 화면 진입 실패 / 고객={job['name']}")

        _dismiss_payment_info_notice_if_open()

        for attempt in range(3):
            print(f"🔎 [ready_sign] 정기결제 수단 선택 열기 시도 {attempt + 1}/3")

            _dismiss_payment_info_notice_if_open()

            ok_click = _click_payment_method_selector()
            if not ok_click:
                time.sleep(1.0)

            if _wait_payment_add_sheet_and_click_add(timeout_sec=6.0):
                print(f"✅ [ready_sign] 결제수단 추가 바텀시트 처리 완료: {job['name']}")

                if not _wait_payment_method_add_page_ready(timeout_sec=10.0):
                    return _abort(f"전자서명 중단: 결제수단 추가 화면 진입 실패 / 고객={job['name']}")

                if _is_card_expiry_like(account_raw):
                    return _abort(f"전자서명 중단: 카드이체 결제정보는 아직 미구현 / 고객={job['name']} / 원본결제정보={account_raw}")

                if not _click_bank_transfer_tab_if_needed(account_raw):
                    return _abort(f"전자서명 중단: 은행이체 탭 선택 실패 / 고객={job['name']} / 원본결제정보={account_raw}")

                if not _open_bank_picker():
                    return _abort(f"전자서명 중단: 은행입력 선택창 열기 실패 / 고객={job['name']} / 원본결제정보={account_raw}")

                if not _choose_bank_from_picker(account_raw):
                    return False

                if not _fill_bank_account_number(account_raw):
                    return False

                if not _click_payment_method_submit():
                    return _abort(f"전자서명 중단: 결제수단 추가하기 버튼 클릭 실패 / 고객={job['name']}")

                if not _wait_after_payment_submit(timeout_sec=10.0):
                    return False

                if _wait_install_info_ready(timeout_sec=2.0):
                    print(f"✅ [ready_sign] 결제수단 추가 후 이미 설치정보 화면 진입: {job['name']}")
                    return True

                if not _click_payment_info_next():
                    return _abort(f"전자서명 중단: 결제정보 선택 화면 다음 버튼 클릭 실패 / 고객={job['name']}")

                print(f"✅ [ready_sign] 결제수단 추가 완료 후 결제정보 다음 클릭 완료: {job['name']}")
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

        target_status = str(entry_status or "").strip()
        if not target_status:
            target_status = "인증완료"

        target_badges = get_status_badges_in_results(d, target_status=target_status)
        print(f"🧾 [ready_sign] {target_status} 후보 수: {len(target_badges)}")

        if not target_badges:
            return False

        matched_badge = None

        for idx, badge in enumerate(target_badges, start=1):
            print(f"✅ [ready_sign] {target_status} 후보 확인 {idx}/{len(target_badges)}")

            ok_open = open_detail_by_status_badge(d, badge)
            if not ok_open:
                print(f"⚠️ [ready_sign] {target_status} 후보 열기 실패 {idx}/{len(target_badges)}")
                continue

            time.sleep(0.90)

            if is_unexpected_digital_sales_home(d):
                if attempt == 0 and recover_from_unexpected_home(d, f"상세진입 상세열기 중 {job['name']}"):
                    time.sleep(0.8)
                    matched_badge = "__RETRY__"
                    break
                return False

            if match_detail_by_name_phone(d, job["name"], job["phone11"]):
                print(f"✅ [ready_sign] 상세 고객명/전화번호 매칭 성공 {idx}/{len(target_badges)}")
                matched_badge = badge
                break

            print(f"⚠️ [ready_sign] 상세 불일치 {idx}/{len(target_badges)} → 다음 {target_status} 후보 확인")
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
            notify(f"{target_status} 후보 {len(target_badges)}개 모두 전화번호 불일치: {job['name']} / {job['phone11']}")
            print(f"❌ [ready_sign] {target_status} 후보 전체 확인했지만 전화번호 일치 없음")
            return False

        search_model = (job.get("model_name") or job.get("product_name") or "").strip()
        color_raw = (job.get("color_raw") or "").strip()
        manage_raw = (job.get("manage_raw") or "").strip()
        contract_raw = (job.get("contract_raw") or "").strip()
        discount_raw = (job.get("discount_raw") or "").strip()
        amount_raw = (job.get("amount_raw") or "").strip()
        account_raw = (job.get("account") or "").strip()
        phone11 = (job.get("phone11") or "").strip()
        zipcode = (job.get("zipcode") or "").strip()
        address_basic = (job.get("address_basic") or job.get("address") or "").strip()
        address_detail = (job.get("address_detail") or "").strip()
        pickup_request_raw = (job.get("pickup_request_raw") or "").strip()
        special_note_raw = (job.get("special_note_raw") or "").strip()

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

        if not account_raw:
            return _abort(f"전자서명 중단: 모달 결제정보 추출 실패 / 고객={job['name']} / 모델={search_model}")

        if not phone11:
            return _abort(f"전자서명 중단: 모달 연락처 추출 실패 / 고객={job['name']} / 모델={search_model}")

        if not zipcode:
            return _abort(f"전자서명 중단: 모달 우편번호 추출 실패 / 고객={job['name']} / 모델={search_model}")

        if not address_basic:
            return _abort(f"전자서명 중단: 모달 기본주소 추출 실패 / 고객={job['name']} / 모델={search_model}")

        if not address_detail:
            return _abort(f"전자서명 중단: 모달 상세주소 추출 실패 / 고객={job['name']} / 모델={search_model}")

        if not d(text="주문 이어서 하기").exists:
            return _abort(f"전자서명 중단: 주문 이어서 하기 버튼 미검출 / 고객={job['name']}")

        click_text_center(d, "주문 이어서 하기", 0.20, 0.85)
        time.sleep(2.0)

        if is_unexpected_digital_sales_home(d):
            return _abort(f"전자서명 중단: 주문 이어서 하기 클릭 후 홈이탈 / 고객={job['name']}")

        resume_result = _try_resume_from_current_screen(
            account_raw=account_raw,
            discount_raw=discount_raw,
            amount_raw=amount_raw,
            phone11=phone11,
            zipcode=zipcode,
            address_basic=address_basic,
            address_detail=address_detail,
            pickup_request_raw=pickup_request_raw,
            special_note_raw=special_note_raw,
        )

        if resume_result is True:
            SIGN_IN_PROGRESS = True
            sign_started.add(job["phone11"])
            notify(
                f"전자서명 재진행 완료: {job['name']} / {job['phone11']} / 시작상태={target_status}"
            )
            print(f"✅ [ready_sign] 현재 단계({target_status})부터 재진행 완료: {job['name']}")
            return True

        if resume_result is False:
            return False

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

        if not _open_payment_method_and_click_add(account_raw):
            return False

        if not _fill_install_info_address(
            phone11,
            zipcode,
            address_basic,
            address_detail,
            pickup_request_raw,
            special_note_raw,
        ):
            return False

        SIGN_IN_PROGRESS = True
        sign_started.add(job["phone11"])
        notify(
            f"전자서명 결제정보/설치주소 단계 완료: {job['name']} / {job['phone11']} / 결제정보={account_raw} / 우편번호={zipcode} / 기본주소={address_basic} / 상세주소={address_detail}"
        )
        print(f"✅ [ready_sign] 상품검색/색상/렌탈/관리/약정/할인검증/결제정보추가/계좌번호입력/설치주소입력 완료: {job['name']}")
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
    global SIGN_IN_PROGRESS, LAST_STOP_REASON

    if not DO_EMULATOR:
        print("ℹ️ DO_EMULATOR=False")
        return

    d = connect_emulator()
    print("✅ 에뮬레이터 연결 완료:", ADB_SERIAL)

    last_reenter = 0.0

    while True:
        try:
            sync_runtime_state_from_db()

            if STOP_FLAG:
                time.sleep(1.0)
                continue

            if SIGN_IN_PROGRESS:
                time.sleep(0.8)
                continue

            try:
                job = auth_q.get_nowait()
            except queue.Empty:
                job = None

            if job is not None:
                try:
                    phone11 = str(job.get("phone11") or "").strip()
                    queued_new_phones.discard(phone11)

                    fresh = db_get_job(phone11)
                    if not fresh or fresh.get("status") != STATUS_NEW:
                        auth_q.task_done()
                        time.sleep(0.1)
                        continue

                    job = dict(fresh)
                    retry = int(job.get("_retry", 0))
                    LAST_STOP_REASON = ""

                    print(f"🚀 인증발송 작업 시작: {job['name']} / {job['phone11']}")

                    ok, reason = send_auth_request(d, job)
                    if ok:
                        saved = db_mark_auth_sent(job)
                        if saved:
                            with jobs_lock:
                                auth_sent_jobs[saved["phone11"]] = dict(saved)
                        print(f"✅ 인증발송 성공: {job['name']} / {job['phone11']} (최초 상태체크 60초 후)")
                        notify(f"인증발송 성공: {job['name']} / {job['phone11']} (최초 상태체크 60초 후)")
                    else:
                        if reason in ["APP_RESTART_FAIL", "NOT_READY", "UNEXPECTED_HOME"] and retry < AUTH_RETRY_MAX:
                            retry += 1
                            job["_retry"] = retry
                            print(f"⚠️ 인증발송 재시도 ({retry}/{AUTH_RETRY_MAX}) : {job['name']} / {job['phone11']} ({reason})")
                            auth_q.put(job)
                            queued_new_phones.add(phone11)
                        else:
                            db_mark_hold(phone11, f"인증발송 실패: {reason}")
                            if STOP_ON_ERROR:
                                set_stop(f"인증발송 실패: {job['name']} / {job['phone11']} ({reason})")
                finally:
                    try:
                        auth_q.task_done()
                    except Exception:
                        pass

                time.sleep(0.2)
                continue

            with jobs_lock:
                pending = [j for (p, j) in auth_sent_jobs.items() if p not in sign_started]

            if not pending:
                time.sleep(1.2)
                continue

            now = time.time()

            pending.sort(key=lambda x: float(x.get("next_check_at") or 0))
            due = [j for j in pending if float(j.get("next_check_at") or 0) <= now]
            if not due:
                time.sleep(SEARCH_LOOP_SLEEP_SEC)
                continue

            batch = due[:SEARCH_BATCH_PER_CYCLE]

            need_refresh = True
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

                phone11 = str(j.get("phone11") or "").strip()
                fresh = db_get_job(phone11)
                if not fresh:
                    continue

                if fresh.get("status") in [STATUS_CANCELLED, STATUS_HOLD, STATUS_DONE]:
                    with jobs_lock:
                        auth_sent_jobs.pop(phone11, None)
                    continue

                j = dict(fresh)

                print("🔎 검색 시도:", j["name"])

                LAST_STOP_REASON = ""
                st = check_one_job_status_by_search(d, j)

                last_check_at = time.time()
                auth_sent_at = float(j.get("auth_sent_at") or 0)
                interval = compute_check_interval(auth_sent_at) if auth_sent_at > 0 else 120
                next_check_at = time.time() + interval

                db_update_check_state(phone11, st, next_check_at, last_check_at)

                saved = db_get_job(phone11)
                if saved:
                    with jobs_lock:
                        auth_sent_jobs[phone11] = dict(saved)

                print(f"🧾 상태체크: {j['name']} / {j['phone11']} => {st or 'NONE'} (다음 {interval}s)")

                if st in ["인증완료", "할인입력", "결제입력", "설치입력"]:
                    LAST_STOP_REASON = ""
                    ok = try_open_ready_sign_detail(d, j, entry_status=st)
                    if ok:
                        db_mark_done(phone11)
                        with jobs_lock:
                            auth_sent_jobs.pop(phone11, None)
                        sign_started.add(phone11)
                        notify(f"✅ 진행 재개/완료: {j['name']} / {j['phone11']} / 시작상태={st}")
                    else:
                        if LAST_STOP_REASON:
                            db_mark_hold(phone11, LAST_STOP_REASON)
                        back_to_order_status(d)

                time.sleep(0.6)

            time.sleep(SEARCH_LOOP_SLEEP_SEC)

        except Exception as e:
            print("❌ 에뮬레이터 루프 에러:", e)
            traceback.print_exc()
            if STOP_ON_ERROR:
                set_stop("에뮬레이터 루프 예외")
            time.sleep(1.0)

init_db()
sync_runtime_state_from_db()
start_record_window_thread()

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
        pickup_request_raw = data.get("pickup_request_raw", "")
        special_note_raw = data.get("special_note_raw", "")
        third_brand = data.get("third_brand", "")
        third_install_mfg = data.get("third_install_mfg", "")
        third_product_type = data.get("third_product_type", "")
        third_product_kind = data.get("third_product_kind", "")
        third_install_shape = data.get("third_install_shape", "")
        third_watermark_no = data.get("third_watermark_no", "")

        if not phone11:
            print("❌ 전화번호 추출 실패 → 건너뜀")
            debug_modal_inputs(modal)
            time.sleep(0.8)
            continue

        if phone11 in processed_phones:
            time.sleep(0.8)
            continue

        processed_phones.add(phone11)

        row = db_upsert_job_from_modal({
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
            "pickup_request_raw": pickup_request_raw,
            "special_note_raw": special_note_raw,
            "third_brand": third_brand,
            "third_install_mfg": third_install_mfg,
            "third_product_type": third_product_type,
            "third_product_kind": third_product_kind,
            "third_install_shape": third_install_shape,
            "third_watermark_no": third_watermark_no,
        })

        current_status = row.get("status", STATUS_NEW) if row else STATUS_NEW

        print("\n✅ 고객 감지/DB 저장")
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
        print("기존 제품 수거:", pickup_request_raw)
        print("특이사항:", special_note_raw)
        print("타사정보-브랜드:", third_brand)
        print("타사정보-설치/제조년월:", third_install_mfg)
        print("타사정보-제품타입:", third_product_type)
        print("타사정보-제품종류:", third_product_kind)
        print("타사정보-설치형태:", third_install_shape)
        print("타사정보-물마크번호:", third_watermark_no)
        print("현재상태:", current_status)
        print("-" * 40)

        if row and current_status == STATUS_NEW and phone11 not in queued_new_phones:
            auth_q.put(dict(row))
            queued_new_phones.add(phone11)

        sync_runtime_state_from_db()
        time.sleep(1.2)

    except Exception as e:
        print("에러 발생:", e)
        traceback.print_exc()
        if STOP_ON_ERROR:
            set_stop("웹 모달 루프 예외")
        time.sleep(1.0)