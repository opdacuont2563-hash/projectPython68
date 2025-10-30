# -*- coding: utf-8 -*-
"""
SurgiBot Server — Tkinter UI + Flask API (Waitress) + Google Sheets (graceful fallback)
- หน้าจอใหญ่ด้วย Tkinter (ID แสดง HN แบบ mask XXX)
- API:
    • GET  /api/health
    • GET  /api/list[?token=SECRET]
         - token ถูกต้อง  -> ส่งรายการพร้อม hn_full
         - ไม่ส่ง/ผิด     -> ส่งรายการแบบปลอดภัย (ID = HN ที่ mask แล้ว, ไม่ส่ง hn_full)
    • GET  /api/list_full?token=SECRET  (compat; เทียบเท่า /api/list?token=SECRET)
    • POST /api/update { token, action, [patient_id|or+queue], status, eta_minutes?, hn? }
- เก็บ hn เต็มไว้ใน patient_data และ push เข้า snapshot ให้ client (ส่ง/ไม่ส่งตาม token)
- จัดตาราง “เสียงประกาศ” ทุก ANNOUNCE_MIN นาทีแบบยึดเวลาคงที่ (เช่น :00, :20, :40)
- NEW:
    • ปุ่ม "ลบแถวที่เลือก" (ลบหลายแถวได้)
    • เปิดโปรแกรมเต็มจออัตโนมัติ (Windows: zoomed; อื่น ๆ fullscreen)
    • Logic อัตโนมัติ:
        - ครบ 1 ชม.จาก "กำลังพักฟื้น" -> "พักฟื้นครบแล้ว" -> ~3 นาที -> "กำลังส่งกลับตึก"
        - เมื่อ "เลื่อนการผ่าตัด" ให้ประกาศเรียกญาติ 2 รอบ (ไทยก่อน–อังกฤษทีหลังในทุกรอบ)
        - (ใหม่) เมื่อเข้าสู่ "กำลังส่งกลับตึก" ครบ ~3 นาที -> ลบรายการอัตโนมัติ (เปิด/ปิดได้จาก UI)
    • เพิ่มชุดคำแปลสถานะผ่าตัดเป็นภาษาอังกฤษ และปรับประกาศทุกจุดเป็น “ไทยก่อน -> เว้น -> อังกฤษ”
      โดยใช้เสียง/ความเร็วเดียวกับระบบ popup QR (กด TTS ผ่าน gTTS + pygame)
    • โหลด Service Account แบบเสถียร + Fallback: หากยังไม่ตั้งค่า credentials จะรันต่อได้โดยปิดฟีเจอร์ Sheets ชั่วคราว
"""

import tkinter as tk
from tkinter import messagebox, ttk
from datetime import datetime, timedelta
import winsound
import threading
import gspread
from google.oauth2 import service_account
import json
import os
from pathlib import Path
import queue
from flask import Flask, request, jsonify
from waitress import serve
import sys
import time  # ใช้จับเวลารอให้เสียงเล่นจนจบ

from .config import CONFIG
from .logging_setup import get_logger
from .workers.audio_worker import AudioWorker

logger = get_logger(__name__)

SURGIBOT_SECRET = CONFIG.secret

API_HOST = CONFIG.api_host
API_PORT = CONFIG.api_port

# รอบประกาศเสียง (นาที) — server จะ sync ค่านี้ไปชีต "Config"
ANNOUNCE_MIN = CONFIG.announce_interval_minutes

# ข้อความประกาศสาธารณะ (ไทย/อังกฤษ)
PUBLIC_ANNOUNCEMENT_TH = (
    "ท่านใดที่ต้องการเดินทางไปยังจุดอื่นหรือไม่ได้อยู่ที่จุดรอผ่าตัดนี้ "
    "ท่านสามารถสแกนคิวอาร์โค้ดที่แสดงที่หน้าจอเพื่อติดตามสถานะการผ่าตัดแบบเรียลไทม์ออนไลน์ได้ตลอดเวลา "
    "โดยติดตามจากรหัสผู้ป่วยที่ท่านได้รับไปค่ะ ขอบคุณค่ะ"
)
PUBLIC_ANNOUNCEMENT_EN = (
    "If you need to go to another area or cannot remain in this surgical waiting area, "
    "please scan the QR code on the screen to follow the surgery status in real time. "
    "Use the patient code you were given. Thank you."
)

# แผนที่คำแปลสถานะ (ไทย -> อังกฤษ)
STATUS_EN = {
    "รอผ่าตัด": "waiting for surgery",
    "กำลังผ่าตัด": "in surgery",
    "กำลังพักฟื้น": "in recovery",
    "พักฟื้นครบแล้ว": "recovery complete",
    "กำลังส่งกลับตึก": "being transferred back to the ward",
    "เลื่อนการผ่าตัด": "surgery postponed",
}

# ตั้งค่าการประกาศเมื่อเลื่อนผ่าตัด
POSTPONED_REPEAT = 2     # จำนวนรอบประกาศ
POSTPONED_GAP_SEC = 8    # เวลาห่างแต่ละรอบ (วินาที) — เริ่มนับหลัง “เล่นจบ” ไทย+อังกฤษแล้ว
BILINGUAL_PAUSE_MS = 600 # พักระหว่างเวอร์ชันไทย -> อังกฤษ ภายใน 1 รอบ

# หน่วงเวลาจาก "พักฟื้นครบแล้ว" -> "กำลังส่งกลับตึก"
AUTO_DISCHARGE_DELAY_MIN = 3

# (ใหม่) ลบอัตโนมัติหลังเข้าสถานะ "กำลังส่งกลับตึก"
AUTO_DELETE_AFTER_DISCHARGE_MIN = CONFIG.auto_delete_minutes

# ===================== Google Sheets (robust loader + graceful fallback) =====================
SHEETS_ENABLED = False
_gspread_client = None
_sheet = None
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
SPREADSHEET_ID = CONFIG.google_sheet_id

def _normalize_sa_info(raw: str | dict) -> dict:
    """รับ JSON string หรือ dict ของ Service Account แล้ว normalize private_key ให้ถูกฟอร์แมต"""
    if isinstance(raw, str):
        data = json.loads(raw)
    else:
        data = dict(raw)

    pk = data.get("private_key", "")
    if "\\n" in pk and "-----BEGIN" in pk:
        pk = pk.replace("\\n", "\n")
    pk = pk.strip()
    if pk.startswith("-----BEGIN PRIVATE KEY-----") and not pk.endswith("-----END PRIVATE KEY-----"):
        if "-----END PRIVATE KEY-----" in pk:
            pk = pk.split("-----END PRIVATE KEY-----")[0] + "-----END PRIVATE KEY-----"
    data["private_key"] = pk
    return data

def _load_service_account_credentials():
    """
    โหลดคีย์จาก (ตามลำดับ):
      1) ENV SURGIBOT_GCP_CREDENTIALS_JSON  (วาง JSON ทั้งก้อน)
      2) ENV SURGIBOT_GCP_CREDENTIALS_FILE  (พาธไฟล์ .json)
      3) ENV SURGIBOT_EMBEDDED_CREDENTIALS_JSON (วาง JSON ไว้ใน ENV นี้)
    """
    env_json = (CONFIG.gcp_credentials_json or "").strip()
    env_file = (CONFIG.gcp_credentials_file or "").strip()
    embedded = (CONFIG.embedded_credentials_json or "").strip()

    sa_info = None
    if env_json:
        sa_info = _normalize_sa_info(env_json)
    elif env_file and Path(env_file).exists():
        sa_info = _normalize_sa_info(Path(env_file).read_text(encoding="utf-8"))
    elif embedded:
        sa_info = _normalize_sa_info(embedded)
    else:
        payload = CONFIG.google_credentials_payload()
        if payload:
            sa_info = _normalize_sa_info(payload)
        else:
            raise RuntimeError(
                "No valid service account credentials provided. "
                "Set SURGIBOT_GCP_CREDENTIALS_JSON or SURGIBOT_GCP_CREDENTIALS_FILE or SURGIBOT_EMBEDDED_CREDENTIALS_JSON."
            )

    pk = sa_info.get("private_key", "")
    if not (pk.startswith("-----BEGIN PRIVATE KEY-----") and "-----END PRIVATE KEY-----" in pk):
        raise ValueError("Invalid private_key PEM format. Please re-check your service account JSON.")

    creds = service_account.Credentials.from_service_account_info(sa_info).with_scopes(SCOPES)
    return creds

def init_sheets():
    """พยายามเชื่อม Google Sheets ถ้าไม่ได้ให้ fallback เฉย ๆ"""
    global SHEETS_ENABLED, _gspread_client, _sheet
    try:
        creds = _load_service_account_credentials()
        _gspread_client = gspread.authorize(creds)
        _sheet = _gspread_client.open_by_key(SPREADSHEET_ID).sheet1
        SHEETS_ENABLED = True
        logger.info("[Sheets] Connected and enabled.")
    except Exception as e:
        SHEETS_ENABLED = False
        _gspread_client = None
        _sheet = None
        message = f"[Sheets] Disabled (reason: {e}). Running without Sheets."
        if "No valid service account credentials" in str(e):
            logger.info(message)
        else:
            logger.warning(message)

def sync_config_to_sheet():
    """อัปเดตชีต Config: ANNOUNCE_MIN (ถ้าเปิด Sheets เท่านั้น)"""
    if not SHEETS_ENABLED:
        return
    try:
        try:
            cfg = _gspread_client.open_by_key(SPREADSHEET_ID).worksheet("Config")
        except gspread.exceptions.WorksheetNotFound:
            ss = _gspread_client.open_by_key(SPREADSHEET_ID)
            cfg = ss.add_worksheet(title="Config", rows=10, cols=4)
        cfg.clear()
        cfg.update("A1:B1", [["ANNOUNCE_MIN", ANNOUNCE_MIN]])
    except Exception as e:
        logger.warning("[Sheets] Config sync error: %s", e)

def _update_next_announce_to_sheet(next_dt: datetime):
    """เขียนค่า NEXT_ANNOUNCE_ISO และ SERVER_NOW_ISO ลงแท็บ Config (ถ้าเปิด Sheets เท่านั้น)"""
    if not SHEETS_ENABLED:
        return
    try:
        ss = _gspread_client.open_by_key(SPREADSHEET_ID)
        try:
            cfg = ss.worksheet("Config")
        except gspread.exceptions.WorksheetNotFound:
            cfg = ss.add_worksheet(title="Config", rows=10, cols=4)

        now_iso = datetime.now().replace(microsecond=0).isoformat()
        next_iso = next_dt.replace(microsecond=0).isoformat()

        cfg.update("A1:B1", [["ANNOUNCE_MIN", ANNOUNCE_MIN]])
        cfg.update("A2:B2", [["NEXT_ANNOUNCE_ISO", next_iso]])
        cfg.update("A3:B3", [["SERVER_NOW_ISO", now_iso]])
    except Exception as e:
        logger.warning("[Sheets] update next announce error: %s", e)

# ===================== Helpers =====================
def _fmt_td(td: timedelta) -> str:
    total_seconds = int(abs(td.total_seconds()))
    h = total_seconds // 3600
    m = (total_seconds % 3600) // 60
    s = total_seconds % 60
    return f"{h:02d}:{m:02d}:{s:02d}"

def mask_hn(hn: str):
    """แสดง 6 ตัวแรก + XXX (เช่น 590166XXX)"""
    if isinstance(hn, str) and len(hn) >= 3:
        return hn[:-3] + "XXX"
    return hn

# ===== คำนวณเวลาถึงรอบถัดไปตาม ANNOUNCE_MIN (ยึดเวลาคงที่) =====
def ms_until_next_boundary(interval_min: int) -> int:
    now = datetime.now()
    sec_from_hour = now.minute * 60 + now.second
    step = max(1, int(interval_min)) * 60
    next_slot_sec = ((sec_from_hour // step) + 1) * step
    delta_sec = next_slot_sec - sec_from_hour
    if delta_sec <= 0:
        delta_sec += step
    return int(delta_sec * 1000)

# ===================== Queue & API App =====================
incoming_queue = queue.Queue()

# snapshot เก็บครบ (รวม hn_full) แต่จะตัดก่อนส่งถ้า token ไม่ถูก
server_snapshot = {"items": []}
_snapshot_lock = threading.Lock()

audio_worker = AudioWorker()

def _build_public_payload(include_hn_full: bool) -> dict:
    with _snapshot_lock:
        if include_hn_full:
            return json.loads(json.dumps(server_snapshot, ensure_ascii=False))
        safe_items = []
        for it in server_snapshot.get("items", []):
            nz = dict(it)
            nz.pop("hn_full", None)
            safe_items.append(nz)
        return {"items": safe_items}

def update_snapshot_from_dict(patient_data: dict):
    rows = []
    now = datetime.now()
    for pid, d in patient_data.items():
        ts = d.get("timestamp")
        eta_m = d.get("eta_minutes")
        eta_iso, remaining = None, None
        if ts and isinstance(ts, datetime) and isinstance(eta_m, int):
            eta_dt = ts + timedelta(minutes=eta_m)
            eta_iso = eta_dt.isoformat()
            remaining = int((eta_dt - now).total_seconds())

        hn_full = d.get("hn")
        masked = mask_hn(hn_full) if hn_full else d.get("id")

        rows.append({
            "id": masked,
            "hn_full": hn_full,
            "patient_id": pid,
            "status": d.get("status", ""),
            "timestamp": ts.isoformat() if ts else None,
            "eta_minutes": eta_m if isinstance(eta_m, int) else None,
            "eta_time": eta_iso,
            "eta_remaining_seconds": remaining
        })
    with _snapshot_lock:
        server_snapshot["items"] = sorted(rows, key=lambda x: str(x.get("patient_id","")))

flask_app = Flask(__name__)

@flask_app.route("/api/health", methods=["GET"])
def api_health():
    return jsonify({"ok": True, "ts": datetime.utcnow().isoformat() + "Z"})


@flask_app.route("/healthz", methods=["GET"])
def healthz():
    return jsonify({"ok": True})

@flask_app.route("/api/list", methods=["GET"])
def api_list():
    token = request.args.get("token", "")
    authed = token == SURGIBOT_SECRET
    return jsonify(_build_public_payload(include_hn_full=authed))

@flask_app.route("/api/list_full", methods=["GET"])
def api_list_full():
    token = request.args.get("token", "")
    if token != SURGIBOT_SECRET:
        return jsonify({"ok": False, "error": "unauthorized"}), 401
    return jsonify(_build_public_payload(include_hn_full=True))

@flask_app.route("/api/update", methods=["POST"])
def api_update():
    """
    { "token":"...", "action":"add|edit|delete",
      "or":"OR1", "queue":"0-2" | "patient_id":"OR1-0-2",
      "status":"กำลังผ่าตัด", "eta_minutes": 90, "hn":"590166994" }
    """
    try:
        data = request.get_json(force=True)
    except Exception:
        return jsonify({"ok": False, "error": "invalid json"}), 400

    token = data.get("token", "")
    if token != SURGIBOT_SECRET:
        return jsonify({"ok": False, "error": "unauthorized"}), 401

    action = (data.get("action") or "").strip().lower()
    pid = data.get("patient_id") or f"{data.get('or','')}-{data.get('queue','')}"
    status = data.get("status")
    eta_minutes = data.get("eta_minutes", None)
    hn = (data.get("hn") or "").strip()

    if action not in ("add", "edit", "delete"):
        return jsonify({"ok": False, "error": "invalid action"}), 400
    if not pid or pid == "-":
        return jsonify({"ok": False, "error": "missing patient_id"}), 400

    if eta_minutes is not None:
        try:
            eta_minutes = int(eta_minutes)
            if eta_minutes < 0:
                eta_minutes = None
        except Exception:
            eta_minutes = None

    if hn and (not hn.isdigit() or len(hn) != 9):
        return jsonify({"ok": False, "error": "HN must be 9 digits"}), 400

    incoming_queue.put({
        "action": action,
        "patient_id": pid,
        "status": status,
        "eta_minutes": eta_minutes,
        "hn": hn if hn else None
    })
    return jsonify({"ok": True, "queued": True, "patient_id": pid})

def _run_api_server():
    serve(flask_app, host=API_HOST, port=API_PORT, threads=4)

# ===================== Theme (ย่อ) =====================
TAG_STYLES_LIGHT = {
    "waiting":           {"background": "#FFF4CC", "foreground": "#5D4037"},
    "surgery":           {"background": "#D0EBFF", "foreground": "#0B3C5D"},
    "recovery":          {"background": "#D3F9D8", "foreground": "#1B5E20"},
    "recovery_complete": {"background": "#B2F2E8", "foreground": "#0F5132"},
    "discharge":         {"background": "#E5DEFF", "foreground": "#3D2C8D"},
    "postponed":         {"background": "#ECEFF1", "foreground": "#37474F"},
}
PULSE_LIGHT_A, PULSE_LIGHT_B = "#D0EBFF", "#E3F3FF"

# ===================== Tkinter App =====================
class SurgeryStatusApp(tk.Frame):
    def __init__(self, root):
        super().__init__(root)
        self.root = root
        self.root.title("Surgery Status Tracker")

        # เริ่มแบบเต็มหน้าจออัตโนมัติ
        try:
            if os.name == "nt":  # Windows
                self.root.state("zoomed")
            else:
                self.root.attributes("-fullscreen", True)
        except Exception:
            self.root.geometry("1600x900")

        self.root.resizable(True, True)

        # โครงสร้างข้อมูล:
        # patient_data[patient_id] = {
        #   "id": int|str,
        #   "status": str (TH),
        #   "timestamp": datetime,
        #   "eta_minutes": int|None,
        #   "hn": str|None,
        #   "auto_to_discharge_at": datetime|None,
        #   "auto_delete_at": datetime|None,
        # }
        self.patient_data = {}
        self.id_counter = 1

        self.root.configure(bg="#f0f4f8")
        self.root.option_add("*TCombobox*Listbox.Font", ("Prompt", 14))

        self.current_theme = "light"
        self._pulse_on = False

        # Header
        title_frame = tk.Frame(root, bg="#1f4e79", pady=14, padx=20)
        title_frame.pack(pady=5, fill="x")
        tk.Label(title_frame, text="ติดตามสถานะการผ่าตัดโรงพยาบาลหนองบัวลำภู",
                 font=("Prompt", 34, "bold"), fg="white", bg="#1f4e79").pack()

        # Inputs (ย่อ)
        input_frame = tk.Frame(self.root, pady=5, bg="#f0f4f8")
        input_frame.pack(pady=5, padx=5, fill="x")

        tk.Label(input_frame, text="รหัสผู้ป่วย:", font=("Prompt", 14), bg="#f0f4f8").grid(row=0, column=0, padx=5, sticky="e")
        self.or_var = tk.StringVar()
        self.patient_id_combobox = ttk.Combobox(input_frame, textvariable=self.or_var,
                                                values=["OR1", "OR2", "OR3", "OR4", "OR5"],
                                                font=("Prompt", 14), width=8)
        self.patient_id_combobox.grid(row=0, column=1, padx=5)

        tk.Label(input_frame, text="คิวผู้ป่วย:", font=("Prompt", 14), bg="#f0f4f8").grid(row=0, column=2, padx=5, sticky="e")
        self.queue_var = tk.StringVar()
        self.queue_combobox = ttk.Combobox(input_frame, textvariable=self.queue_var,
                                           values=["0-1", "0-2", "0-3", "0-4", "0-5"],
                                           font=("Prompt", 14), width=5)
        self.queue_combobox.grid(row=0, column=3, padx=5)

        tk.Label(input_frame, text="สถานะการผ่าตัด:", font=("Prompt", 14), bg="#f0f4f8").grid(row=0, column=4, padx=5, sticky="e")
        self.status_var = tk.StringVar()
        self.status_combobox = ttk.Combobox(input_frame, textvariable=self.status_var,
                                            values=["รอผ่าตัด", "กำลังผ่าตัด", "กำลังพักฟื้น", "กำลังส่งกลับตึก", "เลื่อนการผ่าตัด", "พักฟื้นครบแล้ว"],
                                            font=("Prompt", 14), width=15)
        self.status_combobox.grid(row=0, column=5, padx=5)

        # ปุ่มเพิ่ม & ปุ่มลบ & เช็กบ็อกซ์ลบอัตโนมัติ
        btn_frame = tk.Frame(input_frame, bg="#f0f4f8")
        btn_frame.grid(row=0, column=6, padx=10, sticky="w")

        add_btn = ttk.Button(btn_frame, text="เพิ่มผู้ป่วย", command=self.add_patient)
        add_btn.pack(side="left", padx=(0, 6))

        del_btn = ttk.Button(btn_frame, text="ลบแถวที่เลือก", command=self.delete_selected)
        del_btn.pack(side="left")

        # (ใหม่) เช็กบ็อกซ์ เปิด/ปิด ลบอัตโนมัติหลังส่งกลับตึก
        self.auto_delete_enabled = tk.BooleanVar(value=True)
        chk = ttk.Checkbutton(input_frame,
                              text=f"ลบอัตโนมัติหลังส่งกลับตึก (~{AUTO_DELETE_AFTER_DISCHARGE_MIN} นาที)",
                              variable=self.auto_delete_enabled)
        chk.grid(row=0, column=7, padx=10, sticky="w")

        # ตาราง
        table_frame = tk.Frame(root, bg="#f0f4f8", bd=2, relief="groove")
        table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        style = ttk.Style()
        style.configure("Treeview", font=("Prompt", 28), rowheight=56, background="#ffffff",
                        fieldbackground="#ffffff")
        style.configure("Treeview.Heading", font=("Prompt", 26), padding=6)

        self.tree = ttk.Treeview(
            table_frame,
            columns=("ID", "Patient ID", "Status", "Elapsed", "ETA"),
            show='headings',
            selectmode="extended"
        )
        self.tree.heading("ID", text="ID")
        self.tree.heading("Patient ID", text="รหัสผู้ป่วย (Patient ID)")
        self.tree.heading("Status", text="สถานะ (Status)")
        self.tree.heading("Elapsed", text="เวลาเดินไป (Elapsed)")
        self.tree.heading("ETA", text="เวลาคาดเสร็จ (ETA)")

        self.tree.column("ID", width=160, anchor='center', minwidth=140)
        self.tree.column("Patient ID", width=420, anchor='center', minwidth=360)
        self.tree.column("Status", width=420, anchor='center', minwidth=360)
        self.tree.column("Elapsed", width=260, anchor='center', minwidth=220)
        self.tree.column("ETA", width=340, anchor='center', minwidth=300)

        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side='right', fill='y')
        self.tree.pack(fill=tk.BOTH, expand=True)

        # ช็อตคัท Delete
        self.tree.bind("<Delete>", lambda e: self.delete_selected())

        # Start sounds
        try:
            winsound.Beep(1000, 300)
        except Exception:
            pass

        self.apply_tag_styles()
        self.update_timers()
        self.root.after(200, self.process_incoming_updates)

        # ตั้งตารางเสียงประกาศแบบยึดเวลาคงที่ (ไทย -> อังกฤษ)
        schedule_next_public_announcement(self)

        # ออกจาก fullscreen ด้วย ESC (non-Windows)
        self.root.bind("<Escape>", self._exit_fullscreen)

    # ----- styles -----
    def apply_tag_styles(self):
        for tag, style in TAG_STYLES_LIGHT.items():
            self.tree.tag_configure(tag, background=style["background"], foreground=style["foreground"])

    # ----- sheets / announcement -----
    def sync_with_google_sheets(self):
        if not SHEETS_ENABLED:
            return
        try:
            _sheet.clear()
            _sheet.append_row(["ID(mask)", "PatientID", "Status", "StartTime", "ETA(min)", "ETA_Time"])
            now_str = datetime.now().strftime("%H:%M")
            for patient_id, data in self.patient_data.items():
                status = data.get("status", "")
                start_time_str = data.get("timestamp").strftime("%H:%M") if data.get("timestamp") else ""
                eta_min = data.get("eta_minutes")
                eta_time_str = ""
                if status == "กำลังพักฟื้น":
                    eta_time_str = now_str  # ETA ปัจจุบัน
                    eta_min = ""
                else:
                    if isinstance(eta_min, int) and data.get("timestamp"):
                        eta_dt = data["timestamp"] + timedelta(minutes=eta_min)
                        eta_time_str = eta_dt.strftime("%H:%M")
                _sheet.append_row([
                    mask_hn(data.get("hn")) or data.get("id"),
                    patient_id,
                    status,
                    start_time_str,
                    eta_min or "",
                    eta_time_str
                ])
        except Exception as e:
            logger.warning("[Sheets] Sync error: %s", e)

    # ===== ตัวช่วยอ่านรหัสผู้ป่วย =====
    def _format_pid_th(self, patient_id: str) -> str:
        """อ่าน OR1-05 เป็น 'โออาหนึ่งศูนย์ห้า'"""
        th_digit = {'0': 'ศูนย์', '1': 'หนึ่ง', '2': 'สอง', '3': 'สาม', '4': 'สี่', '5': 'ห้า', '6': 'หก', '7': 'เจ็ด',
                    '8': 'แปด', '9': 'เก้า'}
        th_letter = {
            'O': 'โอ', 'o': 'โอ', 'R': 'อา', 'r': 'อา',
            'A': 'เอ', 'a': 'เอ', 'B': 'บี', 'b': 'บี', 'C': 'ซี', 'c': 'ซี',
            'D': 'ดี', 'd': 'ดี', 'E': 'อี', 'e': 'อี', 'F': 'เอฟ', 'f': 'เอฟ',
            'G': 'จี', 'g': 'จี', 'H': 'เฮช', 'h': 'เฮช', 'I': 'ไอ', 'i': 'ไอ',
            'J': 'เจ', 'j': 'เจ', 'K': 'เค', 'k': 'เค', 'L': 'แอล', 'l': 'แอล',
            'M': 'เอ็ม', 'm': 'เอ็ม', 'N': 'เอ็น', 'n': 'เอ็น', 'P': 'พี', 'p': 'พี',
            'Q': 'คิว', 'q': 'คิว', 'S': 'เอส', 's': 'เอส', 'T': 'ที', 't': 'ที',
            'U': 'ยู', 'u': 'ยู', 'V': 'วี', 'v': 'วี', 'W': 'ดับเบิลยู', 'w': 'ดับเบิลยู',
            'X': 'เอ็กซ์', 'x': 'เอ็กซ์', 'Y': 'วาย', 'y': 'วาย', 'Z': 'แซด', 'z': 'แซด',
        }
        cleaned = ''.join(ch for ch in patient_id if ch not in '-–—_ ')
        parts = []
        for ch in cleaned:
            if ch.isdigit():
                parts.append(th_digit.get(ch, ch))
            elif ch in th_letter:
                parts.append(th_letter[ch])
            else:
                parts.append(ch)
        return ''.join(parts)

    def _format_pid_en(self, patient_id: str) -> str:
        """อ่านรหัสภาษาอังกฤษทีละตัว เช่น OR105 → 'O R one zero five'"""
        en_digit = {'0': 'zero', '1': 'one', '2': 'two', '3': 'three', '4': 'four', '5': 'five',
                    '6': 'six', '7': 'seven', '8': 'eight', '9': 'nine'}
        cleaned = ''.join(ch for ch in patient_id if ch not in '-–—_ ')
        tokens = []
        for ch in cleaned:
            if ch.isdigit():
                tokens.append(en_digit[ch])
            elif ch.isalpha():
                tokens.append(ch.upper())
            else:
                tokens.append(ch)
        return ' '.join(tokens)

    def _speak_bilingual_async(self, th_text: str, en_text: str, pause_ms: int = BILINGUAL_PAUSE_MS):
        """เล่นประกาศ 2 ภาษา: ไทย -> เว้น -> อังกฤษ (ไม่บล็อก UI)"""
        audio_worker.enqueue_bilingual(th_text, en_text, pause_ms)

    def _build_status_messages(self, patient_id: str, status_th: str):
        """สร้างข้อความประกาศสถานะ (ไทย/อังกฤษ)"""
        th_msg = f"สถานะของผู้ป่วยรหัส {patient_id} ขณะนี้อยู่ที่ {status_th}"
        en_status = STATUS_EN.get(status_th, "updated")
        if en_status in ("in surgery", "in recovery"):
            en_msg = f"The status of patient ID {patient_id} is now {en_status}."
        elif en_status == "waiting for surgery":
            en_msg = f"Patient ID {patient_id} is now waiting for surgery."
        elif en_status == "recovery complete":
            en_msg = f"Patient ID {patient_id} has completed recovery."
        elif en_status == "being transferred back to the ward":
            en_msg = f"Patient ID {patient_id} is being transferred back to the ward."
        elif en_status == "surgery postponed":
            en_msg = f"The surgery for patient ID {patient_id} has been postponed."
        else:
            en_msg = f"The status of patient ID {patient_id} has been updated."
        return th_msg, en_msg

    def play_public_bilingual(self):
        """ประกาศสาธารณะ: ไทย -> อังกฤษ"""
        self._speak_bilingual_async(PUBLIC_ANNOUNCEMENT_TH, PUBLIC_ANNOUNCEMENT_EN)

    def play_status_announcement(self, patient_id: str, status_th: str):
        """ประกาศสถานะผู้ป่วย: ไทย -> อังกฤษ"""
        th_msg, en_msg = self._build_status_messages(patient_id, status_th)
        self._speak_bilingual_async(th_msg, en_msg)

    def play_postponed_announcement(self, patient_id: str):
        """
        ประกาศเมื่อเลื่อนผ่าตัด: 2 รอบ
        รอบละ 2 ภาษา (ไทย -> อังกฤษ) และรอให้ “จบจริง” ก่อนเริ่มรอบถัดไป
        ใช้การอ่านรหัสแบบลบขีดกลาง: OR1-05 -> โออาหนึ่งศูนย์ห้า (TH), O R one zero five (EN)
        """
        def _runner():
            pid_th = self._format_pid_th(patient_id)
            pid_en = self._format_pid_en(patient_id)

            th_msg = (
                f"เรียนญาติผู้ป่วยรหัส {pid_th} "
                f"วันนี้มีความจำเป็นต้องปรับเวลาเข้าห้องผ่าตัด "
                f"กรุณามาพบเจ้าหน้าที่ที่หน้าห้องผ่าตัดเพื่อชี้แจงรายละเอียดและเวลานัดหมายใหม่ ขอบคุณค่ะ"
            )
            en_msg = (
                f"Attention, family of patient ID {pid_en}. "
                f"Please come to the operating room front desk to discuss a schedule change. Thank you."
            )
            for i in range(POSTPONED_REPEAT):
                audio_worker.enqueue_bilingual(th_msg, en_msg, BILINGUAL_PAUSE_MS)
                if i + 1 < POSTPONED_REPEAT:
                    time.sleep(max(0, POSTPONED_GAP_SEC))

        threading.Thread(target=_runner, daemon=True).start()

    # ----- helpers -----
    def _apply_status_change(self, patient_id, new_status, eta_minutes=None, announce=True):
        now = datetime.now()
        data = self.patient_data.get(patient_id, {})
        data["status"] = new_status
        data["timestamp"] = now

        # เดิม: จัดการธงส่งกลับตึกอัตโนมัติ
        if new_status == "พักฟื้นครบแล้ว":
            data["auto_to_discharge_at"] = now + timedelta(minutes=AUTO_DISCHARGE_DELAY_MIN)
        elif new_status == "กำลังส่งกลับตึก":
            data["auto_to_discharge_at"] = None
        else:
            data["auto_to_discharge_at"] = None if new_status != "พักฟื้นครบแล้ว" else data.get("auto_to_discharge_at")

        # ใหม่: ตั้ง/ล้างนาฬิกาลบอัตโนมัติเมื่อเข้าสถานะ "กำลังส่งกลับตึก" (ขึ้นกับเช็กบ็อกซ์)
        if new_status == "กำลังส่งกลับตึก":
            data["auto_delete_at"] = (now + timedelta(minutes=AUTO_DELETE_AFTER_DISCHARGE_MIN)
                                      if self.auto_delete_enabled.get() else None)
        else:
            data["auto_delete_at"] = None

        if eta_minutes is not None:
            try:
                data["eta_minutes"] = int(eta_minutes)
            except Exception:
                pass
        self.patient_data[patient_id] = data

        # UI + sheets + snapshot
        self._refresh_row(patient_id)
        try:
            self.sync_with_google_sheets()
        except Exception as e:
            logger.warning("[Sheets] sync after status change error: %s", e)
        update_snapshot_from_dict(self.patient_data)

        # ประกาศเสียง
        if announce:
            if new_status == "เลื่อนการผ่าตัด":
                self.play_postponed_announcement(patient_id)
            else:
                self.play_status_announcement(patient_id, new_status)

    def _apply_status_tag(self, tree_item_id, status_text):
        tag = None
        if status_text == "รอผ่าตัด":
            tag = "waiting"
        elif status_text == "กำลังผ่าตัด":
            tag = "surgery"
        elif status_text == "กำลังพักฟื้น":
            tag = "recovery"
        elif status_text == "พักฟื้นครบแล้ว":
            tag = "recovery_complete"
        elif status_text == "กำลังส่งกลับตึก":
            tag = "discharge"
        elif status_text == "เลื่อนการผ่าตัด":
            tag = "postponed"
        if tag:
            self.tree.item(tree_item_id, tags=(tag,))

    # ----- Delete UI -----
    def delete_selected(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("ลบข้อมูล", "กรุณาเลือกแถวที่ต้องการลบในตารางก่อน")
            return
        if not messagebox.askyesno("ยืนยันการลบ", f"ต้องการลบ {len(sel)} แถวที่เลือกหรือไม่?"):
            return

        removed = 0
        for iid in sel:
            vals = self.tree.item(iid, "values")
            if len(vals) < 2:
                continue
            patient_id = str(vals[1])
            if patient_id in self.patient_data:
                del self.patient_data[patient_id]
                removed += 1
            self.tree.delete(iid)

        if removed > 0:
            try:
                self.sync_with_google_sheets()
            except Exception as e:
                logger.warning("[Sheets] sync after delete error: %s", e)
            update_snapshot_from_dict(self.patient_data)

    # ----- CRUD (ย่อ) -----
    def add_patient(self):
        patient_id = f"{self.or_var.get()}-{self.queue_var.get()}"
        status = self.status_var.get()
        if not self.or_var.get() or not self.queue_var.get() or not status:
            messagebox.showerror("ข้อผิดพลาด", "กรุณากรอกข้อมูลให้ครบถ้วน")
            return
        if patient_id in self.patient_data:
            messagebox.showerror("ข้อผิดพลาด", "รหัสผู้ป่วยนี้มีอยู่แล้ว")
            return

        now = datetime.now()
        auto_delete_at = None
        if status == "กำลังส่งกลับตึก" and self.auto_delete_enabled.get():
            # แก้ไข: ถ้าเพิ่มด้วยสถานะนี้ ให้ตั้งนาฬิกาลบอัตโนมัติทันที
            auto_delete_at = now + timedelta(minutes=AUTO_DELETE_AFTER_DISCHARGE_MIN)

        self.patient_data[patient_id] = {
            "id": self.id_counter,
            "status": status,
            "timestamp": now,
            "auto_to_discharge_at": None,
            "auto_delete_at": auto_delete_at,  # แก้ใหม่
        }
        show_id = mask_hn(self.patient_data[patient_id].get("hn")) or self.id_counter
        iid = self.tree.insert("", "end", values=(show_id, patient_id, status, "", ""))
        self._apply_status_tag(iid, status)
        self.id_counter += 1
        self.sync_with_google_sheets()
        update_snapshot_from_dict(self.patient_data)
        # ประกาศเสียงตามสถานะ
        if status == "เลื่อนการผ่าตัด":
            self.play_postponed_announcement(patient_id)
        else:
            self.play_status_announcement(patient_id, status)

    # ----- Timers -----
    def update_timers(self):
        now = datetime.now()
        to_delete = []  # เก็บ patient_id ที่ครบกำหนดลบ

        for item in self.tree.get_children():
            values = self.tree.item(item, 'values')
            if len(values) < 3:
                continue
            patient_id = values[1]
            if patient_id not in self.patient_data:
                continue
            data = self.patient_data[patient_id]
            status = data.get("status")
            ts = data.get("timestamp")
            eta_m = data.get("eta_minutes")

            # อัตโนมัติ: กำลังพักฟื้น -> (ครบ 1 ชม.) -> พักฟื้นครบแล้ว
            if status == "กำลังพักฟื้น" and ts:
                end_dt = ts + timedelta(hours=1)
                if now >= end_dt:
                    self._apply_status_change(patient_id, "พักฟื้นครบแล้ว", announce=True)
                    status = "พักฟื้นครบแล้ว"

            # อัตโนมัติ: พักฟื้นครบแล้ว -> (~3 นาที) -> กำลังส่งกลับตึก
            if status == "พักฟื้นครบแล้ว":
                auto_dt = data.get("auto_to_discharge_at")
                if isinstance(auto_dt, datetime) and now >= auto_dt:
                    self._apply_status_change(patient_id, "กำลังส่งกลับตึก", announce=True)
                    status = "กำลังส่งกลับตึก"

            # ใหม่: ถ้ากำลังส่งกลับตึกและยังไม่มี auto_delete_at แต่เปิดลบอัตโนมัติ → ตั้งให้ทันที
            if status == "กำลังส่งกลับตึก" and self.auto_delete_enabled.get():
                if not isinstance(data.get("auto_delete_at"), datetime):
                    base_ts = ts or now
                    data["auto_delete_at"] = base_ts + timedelta(minutes=AUTO_DELETE_AFTER_DISCHARGE_MIN)

                del_at = data.get("auto_delete_at")
                if isinstance(del_at, datetime) and now >= del_at:
                    to_delete.append(patient_id)
                    continue  # ข้ามการอัปเดตคอลัมน์แสดงผล เพราะกำลังจะลบ

            # แสดงเวลา
            elapsed_text, eta_text = "", ""
            if ts:
                if status == "กำลังผ่าตัด":
                    elapsed = now - ts
                    elapsed_text = _fmt_td(elapsed)
                    if isinstance(eta_m, int):
                        eta_dt = ts + timedelta(minutes=eta_m)
                        remain = eta_dt - now
                        hhmm = eta_dt.strftime("%H:%M") + " น."
                        if remain.total_seconds() >= 0:
                            eta_text = f"{hhmm} • เหลือ {_fmt_td(remain)}"
                        else:
                            eta_text = f"{hhmm} • เกินเวลา {_fmt_td(remain)}".replace("-", "")
                elif status == "กำลังพักฟื้น":
                    end_dt = ts + timedelta(hours=1)
                    remain = end_dt - now
                    if remain.total_seconds() < 0:
                        remain = timedelta(seconds=0)
                    elapsed_text = _fmt_td(remain)  # นับถอยหลัง
                    eta_text = now.strftime("%H:%M") + " น."  # ETA ปัจจุบัน
                else:
                    elapsed_text = _fmt_td(now - ts)

            self.tree.item(item, values=(
                mask_hn(data.get("hn")) or data.get("id"),
                patient_id, status, elapsed_text, eta_text
            ))
            self._apply_status_tag(item, status)

        # ลบรายการที่ครบกำหนด
        if to_delete:
            for pid in to_delete:
                if pid in self.patient_data:
                    del self.patient_data[pid]
                self._remove_row(pid)
            try:
                self.sync_with_google_sheets()
            except Exception as e:
                logger.warning("[Sheets] sync after auto-delete error: %s", e)
            update_snapshot_from_dict(self.patient_data)

        self.root.after(1000, self.update_timers)

    # ----- Queue from API -----
    def process_incoming_updates(self):
        try:
            while True:
                msg = incoming_queue.get_nowait()
                action = msg.get("action")
                patient_id = msg.get("patient_id")
                status = msg.get("status")
                eta_minutes = msg.get("eta_minutes", None)
                hn = msg.get("hn", None)

                if action == "add":
                    if patient_id not in self.patient_data:
                        now = datetime.now()
                        auto_delete_at = None
                        if (status or "") == "กำลังส่งกลับตึก" and self.auto_delete_enabled.get():
                            # แก้ไข: เพิ่มเข้าใหม่พร้อมตั้งนาฬิกาลบ
                            auto_delete_at = now + timedelta(minutes=AUTO_DELETE_AFTER_DISCHARGE_MIN)

                        self.patient_data[patient_id] = {
                            "id": self.id_counter,
                            "status": status or "รอผ่าตัด",
                            "timestamp": now,
                            "auto_to_discharge_at": None,
                            "auto_delete_at": auto_delete_at,  # แก้ใหม่
                        }
                        if hn:
                            self.patient_data[patient_id]["hn"] = str(hn).strip()
                        if eta_minutes is not None:
                            try:
                                self.patient_data[patient_id]["eta_minutes"] = int(eta_minutes)
                            except Exception:
                                pass
                        show_id = mask_hn(self.patient_data[patient_id].get("hn")) or self.id_counter
                        iid = self.tree.insert("", "end", values=(show_id, patient_id, self.patient_data[patient_id]["status"], "", ""))
                        self._apply_status_tag(iid, self.patient_data[patient_id]["status"])
                        self.id_counter += 1
                        self.sync_with_google_sheets()
                        if self.patient_data[patient_id]["status"] == "เลื่อนการผ่าตัด":
                            self.play_postponed_announcement(patient_id)
                        else:
                            self.play_status_announcement(patient_id, self.patient_data[patient_id]["status"])
                        update_snapshot_from_dict(self.patient_data)
                    else:
                        if status and status != self.patient_data[patient_id].get("status"):
                            self._apply_status_change(patient_id, status, eta_minutes)
                        else:
                            if eta_minutes is not None:
                                try:
                                    self.patient_data[patient_id]["eta_minutes"] = int(eta_minutes)
                                except Exception:
                                    pass
                            if hn:
                                self.patient_data[patient_id]["hn"] = str(hn).strip()

                            # แก้ไข: ถ้ามีอยู่แล้วและสถานะเป็นกำลังส่งกลับตึก แต่ยังไม่มี auto_delete_at → ตั้งให้
                            if self.patient_data[patient_id].get("status") == "กำลังส่งกลับตึก" and \
                               self.auto_delete_enabled.get() and \
                               not isinstance(self.patient_data[patient_id].get("auto_delete_at"), datetime):
                                base_ts = self.patient_data[patient_id].get("timestamp") or datetime.now()
                                self.patient_data[patient_id]["auto_delete_at"] = base_ts + timedelta(minutes=AUTO_DELETE_AFTER_DISCHARGE_MIN)

                            self._refresh_row(patient_id)
                            self.sync_with_google_sheets()
                            update_snapshot_from_dict(self.patient_data)
                            if status:
                                if status == "เลื่อนการผ่าตัด":
                                    self.play_postponed_announcement(patient_id)
                                else:
                                    self.play_status_announcement(patient_id, self.patient_data[patient_id]["status"])

                elif action == "edit":
                    if patient_id in self.patient_data and (status or eta_minutes is not None or hn):
                        if status and status != self.patient_data[patient_id].get("status"):
                            self._apply_status_change(patient_id, status, eta_minutes)
                        else:
                            if eta_minutes is not None:
                                try:
                                    self.patient_data[patient_id]["eta_minutes"] = int(eta_minutes)
                                except Exception:
                                    pass
                            if hn:
                                self.patient_data[patient_id]["hn"] = str(hn).strip()

                            # ป้องกันเคส edit ข้อมูลอื่น ๆ ขณะสถานะเป็นกำลังส่งกลับตึก แต่ยังไม่ตั้งนาฬิกาลบ
                            if self.patient_data[patient_id].get("status") == "กำลังส่งกลับตึก" and \
                               self.auto_delete_enabled.get() and \
                               not isinstance(self.patient_data[patient_id].get("auto_delete_at"), datetime):
                                base_ts = self.patient_data[patient_id].get("timestamp") or datetime.now()
                                self.patient_data[patient_id]["auto_delete_at"] = base_ts + timedelta(minutes=AUTO_DELETE_AFTER_DISCHARGE_MIN)

                            self._refresh_row(patient_id)
                            self.sync_with_google_sheets()
                            update_snapshot_from_dict(self.patient_data)
                            if status:
                                if status == "เลื่อนการผ่าตัด":
                                    self.play_postponed_announcement(patient_id)
                                else:
                                    self.play_status_announcement(patient_id, self.patient_data[patient_id]["status"])

                elif action == "delete":
                    if patient_id in self.patient_data:
                        del self.patient_data[patient_id]
                        self._remove_row(patient_id)
                        self.sync_with_google_sheets()
                        update_snapshot_from_dict(self.patient_data)
        except queue.Empty:
            pass

        self.root.after(200, self.process_incoming_updates)

    def _refresh_row(self, patient_id):
        for iid in self.tree.get_children():
            vals = self.tree.item(iid, "values")
            if str(vals[1]) == str(patient_id):
                d = self.patient_data.get(patient_id, {})
                show_id = mask_hn(d.get("hn")) or d.get("id")
                status = d.get("status", "")
                self.tree.item(iid, values=(show_id, patient_id, status, "", ""))
                self._apply_status_tag(iid, status)
                return

    def _remove_row(self, patient_id):
        for iid in self.tree.get_children():
            vals = self.tree.item(iid, "values")
            if str(vals[1]) == str(patient_id):
                self.tree.delete(iid)
                return

    def _exit_fullscreen(self, event=None):
        try:
            if os.name != "nt":
                self.root.attributes("-fullscreen", False)
        except Exception:
            pass

# ===== ตั้งรอบประกาศเสียงตามเส้นเวลาแน่นอน (ไทย -> อังกฤษ) =====
def schedule_next_public_announcement(app_self: SurgeryStatusApp):
    def do_announce():
        try:
            app_self.play_public_bilingual()
        except Exception as e:
            logger.error("[announce] tts error: %s", e)

        # นัดรอบถัดไป + เขียนค่าไปชีต
        ms = ms_until_next_boundary(ANNOUNCE_MIN)
        next_dt = datetime.now() + timedelta(milliseconds=ms)
        try:
            _update_next_announce_to_sheet(next_dt)
        except Exception as e:
            logger.warning("[Sheets] next announce write error: %s", e)

        app_self.root.after(ms, lambda: schedule_next_public_announcement(app_self))

    ms0 = ms_until_next_boundary(ANNOUNCE_MIN)
    next_dt0 = datetime.now() + timedelta(milliseconds=ms0)
    try:
        _update_next_announce_to_sheet(next_dt0)
    except Exception as e:
        logger.warning("[Sheets] first next announce write error: %s", e)

    app_self.root.after(ms0, do_announce)

# ===================== main =====================
def main():
    # พยายามเชื่อม Google Sheets (ถ้าไม่สำเร็จ จะทำงานโหมดไม่มีชีต)
    init_sheets()

    threading.Thread(target=_run_api_server, daemon=True).start()
    try:
        sync_config_to_sheet()
    except Exception as e:
        logger.warning("Initial config sync failed: %s", e)

    root = tk.Tk()
    app = SurgeryStatusApp(root)
    app.pack()
    try:
        root.mainloop()
    finally:
        audio_worker.stop()

if __name__ == "__main__":
    main()
