# -*- coding: utf-8 -*-
"""
SurgiBot Client — PySide6 (revamped layout)
- Banner (Online/Offline + Reconnect + Settings)
- Top compact controls bar: Identify / Assign Room / Status & Timing
- Tabs: (1) Result Schedule Patient, (2) Status Operation Real Time
- Fix text overlapping in schedule delegate
"""

import os, sys, json, argparse
import math
import time
from pathlib import Path
from typing import Union, List, Dict
from datetime import datetime, timedelta, time as dtime, date as ddate

import requests

from PySide6 import QtCore, QtGui, QtWidgets
from PySide6.QtCore import QSettings, QUrl
from PySide6.QtGui import (
    QShortcut, QKeySequence, QIcon, QPixmap, QPainter,
    QLinearGradient, QColor, QImageReader
)
from PySide6.QtWidgets import QSystemTrayIcon, QSizePolicy, QFormLayout

try:
    from registry_patient_connect import make_search_combo, SCRUB_NURSES, SearchSelectAdder
except Exception:  # pragma: no cover - fallback without optional dependency
    def make_search_combo(options: list[str]) -> QtWidgets.QComboBox:
        cb = QtWidgets.QComboBox()
        cb.setEditable(True)
        cb.addItems([""] + list(options))
        cb.setInsertPolicy(QtWidgets.QComboBox.NoInsert)
        return cb

    SCRUB_NURSES: list[str] = []

    class SearchSelectAdder(QtWidgets.QWidget):
        itemsChanged = QtCore.Signal(list)

        def __init__(self, placeholder: str = "", suggestions: list[str] | None = None, parent=None):
            super().__init__(parent)
            layout = QtWidgets.QVBoxLayout(self)
            layout.setContentsMargins(0, 0, 0, 0)
            layout.setSpacing(6)

            row = QtWidgets.QHBoxLayout()
            row.setSpacing(6)
            self.combo = QtWidgets.QComboBox()
            self.combo.setEditable(True)
            self.combo.setInsertPolicy(QtWidgets.QComboBox.NoInsert)
            if self.combo.lineEdit():
                self.combo.lineEdit().setPlaceholderText(placeholder)
            row.addWidget(self.combo, 1)
            self.btn = QtWidgets.QPushButton("➕ เพิ่ม")
            row.addWidget(self.btn)
            layout.addLayout(row)

            self.list = QtWidgets.QListWidget()
            self.list.setSelectionMode(QtWidgets.QAbstractItemView.ExtendedSelection)
            layout.addWidget(self.list)

            self.set_suggestions(suggestions or [])
            self.btn.clicked.connect(self._add_current)
            if self.combo.lineEdit():
                self.combo.lineEdit().returnPressed.connect(self._add_current)
            self.list.model().rowsInserted.connect(lambda *_: self._emit())
            self.list.model().rowsRemoved.connect(lambda *_: self._emit())

        def _add_current(self):
            text = self.combo.currentText().strip()
            if text and text.lower() not in [self.list.item(i).text().lower() for i in range(self.list.count())]:
                self.list.addItem(text)
            self.combo.setCurrentIndex(0)
            self.combo.setEditText("")
            self._emit()

        def items(self) -> list[str]:
            return [self.list.item(i).text().strip() for i in range(self.list.count())]

        def clear(self):
            self.list.clear()
            self.combo.setCurrentIndex(0)
            self.combo.setEditText("")
            self._emit()

        def set_suggestions(self, suggestions: list[str]):
            opts = sorted({s for s in suggestions if s})
            self.combo.blockSignals(True)
            self.combo.clear()
            self.combo.addItem("")
            self.combo.addItems(opts)
            self.combo.blockSignals(False)
            comp = QtWidgets.QCompleter(opts)
            comp.setCaseSensitivity(QtCore.Qt.CaseInsensitive)
            comp.setFilterMode(QtCore.Qt.MatchContains)
            self.combo.setCompleter(comp)

        def _emit(self):
            self.itemsChanged.emit(self.items())

try:
    from icd10_catalog import diagnosis_suggestions, operation_suggestions
except Exception:  # pragma: no cover - keep UI responsive without catalog
    def operation_suggestions(_specialty: str | None = None) -> list[str]:
        return []

    def diagnosis_suggestions(_specialty: str | None = None, _ops: list[str] | None = None) -> list[str]:
        return []
from PySide6.QtWebSockets import QWebSocket

from .config import CONFIG
from .logging_setup import get_logger
from .workers.io_worker import SESSION_MANAGER, NetworkTask

logger = get_logger(__name__)

# ---------- Defaults ----------
DEFAULT_HOST = CONFIG.client_host
DEFAULT_PORT = CONFIG.client_port
DEFAULT_TOKEN = CONFIG.client_secret
DEFAULT_TIMEOUT = CONFIG.client_timeout

API_HEALTH = "/api/health"
API_UPDATE = "/api/update"
API_LIST   = "/api/list"
API_LIST_FULL = "/api/list_full"
API_WS = "/api/ws"

STATUS_CHOICES = ["รอผ่าตัด", "กำลังผ่าตัด", "กำลังพักฟื้น", "กำลังส่งกลับตึก", "เลื่อนการผ่าตัด"]
OR_CHOICES     = ["OR1", "OR2", "OR3", "OR4", "OR5", "OR6", "OR8"]
QUEUE_CHOICES  = ["0-1", "0-2", "0-3", "0-4", "0-5", "0-6", "0-7"]

STATUS_OP_START = "กำลังผ่าตัด"
STATUS_OP_END = "กำลังพักฟื้น"
STATUS_RETURNING = "กำลังส่งกลับตึก"

STATUS_COLORS = {
    "รอผ่าตัด": "#fde047",
    "กำลังผ่าตัด": "#ef4444",
    "กำลังพักฟื้น": "#22c55e",
    "กำลังส่งกลับตึก": "#a855f7",
    "เลื่อนการผ่าตัด": "#64748b",
}
OR_HEADER_COLORS = {
    "OR1": "#3b82f6",
    "OR2": "#10b981",
    "OR3": "#f59e0b",
    "OR4": "#ef4444",
    "OR5": "#8b5cf6",
    "OR6": "#0ea5e9",
    "OR8": "#64748b",
}

# ---- Auto purge (client-side) ----
AUTO_PURGE_MINUTES = CONFIG.client_auto_purge_minutes
AUTO_PURGE_STATUSES = {"กำลังส่งกลับตึก"}

# ---------- Shared schedule ----------
ORG_NAME    = "ORNBH"
APP_SHARED  = "SurgiBotShared"
OR_KEY      = "schedule/or_rooms"
ENTRIES_KEY = "schedule/entries"
SEQ_KEY     = "schedule/seq"

# ---------- Persistent monitor keys ----------
PERSIST_ORG = "ORNBH"
PERSIST_APP = "SurgiBotClient"
KEY_LAST_ROWS = "monitor/last_rows_json"
KEY_WAS_IN_MONITOR = "monitor/was_in_monitor_json"
KEY_CURRENT_MONITOR = "monitor/current_monitor_json"

# ---------- Working-hours helpers ----------
def _now_period(dt_val: datetime) -> str:
    start = dtime(8,30); end = dtime(16,30)
    return "in" if (start <= dt_val.time() < end) else "off"

def _period_label(code: str) -> str:
    return "ในเวลาราชการ" if code == "in" else "นอกเวลาราชการ"

class _SchedEntry:
    def __init__(self, d: Dict):
        known_keys = {
            "or", "date", "time", "hn", "name", "age", "dept", "doctor", "diags", "ops",
            "ward", "queue", "period", "case_size", "urgency", "assist1", "assist2",
            "scrub", "circulate", "time_start", "time_end", "status", "state",
            "returning_started_at", "version", "updated_at"
        }
        self.or_room = str(d.get("or","") or "")
        self.date = str(d.get("date","") or "")
        self.date_obj = _parse_date(self.date)
        self.time = str(d.get("time","") or "")
        self.hn = str(d.get("hn","") or "")
        self.name = str(d.get("name","") or "")
        age_val = d.get("age")
        self.age = str(age_val) if age_val not in (None, "") else ""
        self.dept = str(d.get("dept","") or "")
        self.doctor = str(d.get("doctor","") or "")
        self.diags = d.get("diags") or []
        self.ops = d.get("ops") or []
        self.ward = str(d.get("ward","") or "")
        self.queue = int(d.get("queue") or 1)
        self.period = str(d.get("period") or "in")
        self.case_size = str(d.get("case_size", "") or "")
        self.urgency = str(d.get("urgency", "Elective") or "Elective")
        self.assist1 = str(d.get("assist1", "") or "")
        self.assist2 = str(d.get("assist2", "") or "")
        self.scrub = str(d.get("scrub", "") or "")
        self.circulate = str(d.get("circulate", "") or "")
        self.time_start = str(d.get("time_start", "") or "")
        self.time_end = str(d.get("time_end", "") or "")
        self.status = str(d.get("status", "") or "")
        self.state = str(d.get("state", "") or "")
        self.returning_started_at = str(d.get("returning_started_at", "") or "")
        try:
            self.version = int(d.get("version") or 0)
        except Exception:
            self.version = 0
        self.updated_at = str(d.get("updated_at", "") or "")
        self._extra = {k: v for k, v in d.items() if k not in known_keys}

    def uid(self) -> str:
        return f"{self.or_room}|{self.hn}|{self.time}|{self.date}"

    def to_dict(self) -> Dict:
        payload = {
            "or": self.or_room,
            "date": self.date or (self.date_obj.isoformat() if self.date_obj else ""),
            "time": self.time,
            "hn": self.hn,
            "name": self.name,
            "age": int(self.age) if str(self.age).isdigit() else self.age,
            "dept": self.dept,
            "doctor": self.doctor,
            "diags": list(self.diags or []),
            "ops": list(self.ops or []),
            "ward": self.ward,
            "queue": int(self.queue or 0),
            "period": self.period,
            "case_size": self.case_size,
            "urgency": self.urgency,
            "assist1": self.assist1,
            "assist2": self.assist2,
            "scrub": self.scrub,
            "circulate": self.circulate,
            "time_start": self.time_start,
            "time_end": self.time_end,
            "status": self.status,
            "state": self.state,
            "returning_started_at": self.returning_started_at,
            "version": int(self.version or 0),
            "updated_at": self.updated_at,
        }
        payload.update(self._extra)
        return payload
class SharedScheduleReader:
    def __init__(self):
        self.s = QSettings(ORG_NAME, APP_SHARED)
        self._seq = int(self.s.value(SEQ_KEY, 0))
        self.or_rooms = self._load_or()
        self.entries = self._load_entries()
    def _load_or(self) -> List[str]:
        lst = self.s.value(OR_KEY, [])
        return [str(x) for x in (lst or [])]
    def _load_entries(self) -> List[_SchedEntry]:
        raw = self.s.value(ENTRIES_KEY, [])
        out = []
        if isinstance(raw, list):
            for d in raw:
                if isinstance(d, dict):
                    out.append(_SchedEntry(d))
        return out
    def seq(self) -> int:
        return int(self.s.value(SEQ_KEY, 0))
    def refresh_if_changed(self) -> bool:
        cur = self.seq()
        if cur != self._seq:
            self._seq = cur
            self.or_rooms = self._load_or()
            self.entries = self._load_entries()
            return True
        return False


class SharedScheduleModel(SharedScheduleReader):
    def __init__(self):
        super().__init__()

    def _save(self):
        payload = [e.to_dict() for e in self.entries]
        next_seq = int(self.s.value(SEQ_KEY, 0) or 0) + 1
        self.s.setValue(ENTRIES_KEY, payload)
        self.s.setValue(SEQ_KEY, next_seq)
        self.s.sync()
        self._seq = next_seq

    def touch_entry(self, entry: _SchedEntry):
        if entry is None:
            return False
        uid = entry.uid()
        for idx, existing in enumerate(self.entries):
            if existing is entry or existing.uid() == uid:
                self.entries[idx] = entry
                self._save()
                return True
        # ถ้าไม่พบ ให้เพิ่มใหม่เพื่อความปลอดภัย (เช่น registry เพิ่งเพิ่ม)
        self.entries.append(entry)
        self._save()
        return True

    def find_by_uid(self, uid: str) -> _SchedEntry | None:
        for entry in self.entries:
            if entry.uid() == uid:
                return entry
        return None

def _fmt_td(td: timedelta) -> str:
    total = int(abs(td.total_seconds()))
    h = total // 3600
    m = (total % 3600) // 60
    s = total % 60
    return f"{h:02d}:{m:02d}:{s:02d}"

def _parse_iso(ts: str):
    if not isinstance(ts, str) or not ts: return None
    try:
        return datetime.fromisoformat(ts.replace("Z",""))
    except Exception:
        return None

def _parse_date(date_str: str):
    if not isinstance(date_str, str):
        return None
    txt = date_str.strip()
    if not txt:
        return None
    cleaned = txt.replace("Z", "")
    try:
        return datetime.fromisoformat(cleaned).date()
    except ValueError:
        pass
    formats = (
        "%Y-%m-%d",
        "%Y/%m/%d",
        "%d/%m/%Y",
        "%d-%m-%Y",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%d %H:%M:%S",
        "%Y/%m/%d %H:%M",
        "%Y/%m/%d %H:%M:%S",
        "%d/%m/%Y %H:%M",
        "%d/%m/%Y %H:%M:%S",
        "%Y-%m-%dT%H:%M",
        "%Y-%m-%dT%H:%M:%S",
    )
    for fmt in formats:
        try:
            return datetime.strptime(cleaned, fmt).date()
        except ValueError:
            continue
    if cleaned.startswith(datetime.now().date().isoformat()):
        return datetime.now().date()
    return None

# ---------- HTTP ----------
class SurgiBotClientHTTP:
    def __init__(self, host=DEFAULT_HOST, port=DEFAULT_PORT, token=DEFAULT_TOKEN, timeout=DEFAULT_TIMEOUT):
        self.base = f"http://{host}:{port}"
        self.token = token
        self.timeout = timeout or CONFIG.request_timeout
        self.sess = SESSION_MANAGER.get()

    def health(self):
        r = self.sess.get(self.base + API_HEALTH, timeout=self.timeout, headers={"Accept":"application/json"})
        r.raise_for_status(); return r.json()

    def send_update(self, action, or_room=None, queue=None, status=None, patient_id=None, eta_minutes=None, hn=None):
        payload = {"token": self.token, "action": action}
        if patient_id:
            payload["patient_id"] = str(patient_id)
        else:
            if or_room: payload["or"] = str(or_room)
            if queue:   payload["queue"] = str(queue)
        if status is not None: payload["status"] = str(status)
        if hn: payload["hn"] = str(hn)
        if eta_minutes is not None and str(eta_minutes).strip() != "":
            try: payload["eta_minutes"] = int(eta_minutes)
            except Exception: pass
        r = self.sess.post(self.base + API_UPDATE, json=payload, timeout=self.timeout, headers={"Accept":"application/json"})
        try:
            data = r.json()
        except Exception:
            data = {"ok": False, "error": f"HTTP {r.status_code}", "text": r.text}
        if r.status_code >= 400:
            raise requests.HTTPError(json.dumps(data, ensure_ascii=False))
        return data

    def _wrap_items(self, data):
        if isinstance(data, list): return {"items": data}
        if isinstance(data, dict):
            for k in ("items","data","table","rows","list"):
                if k in data and isinstance(data[k], list): return {"items": data[k]}
            for v in data.values():
                if isinstance(v, list): return {"items": v}
            return data
        return {"items": []}

    def list_items(self):
        try:
            r = self.sess.get(f"{self.base}{API_LIST_FULL}?token={self.token}", timeout=self.timeout, headers={"Accept":"application/json"})
            if r.status_code == 200: return self._wrap_items(r.json())
        except Exception:
            pass
        try:
            r = self.sess.get(self.base + API_LIST, timeout=self.timeout, headers={"Accept":"application/json"})
            if r.status_code == 200: return self._wrap_items(r.json())
        except Exception:
            pass
        return {"items": []}


class SubmitError(Exception):
    def __init__(self, original: Exception, context: Dict):
        super().__init__(str(original))
        self.original = original
        self.context = context


# ---------- Local model ----------
class LocalTableModel:
    def __init__(self):
        self.rows, self._seq = [], 1
    def _find(self, pid):
        for i, r in enumerate(self.rows):
            if r["patient_id"] == pid: return i
        return -1
    def add_or_edit(self, pid, status, timestamp=None, eta_minutes=None, hn=None):
        i = self._find(pid)
        if i >= 0:
            self.rows[i]["status"] = status
            if timestamp is not None: self.rows[i]["timestamp"] = timestamp
            if eta_minutes is not None: self.rows[i]["eta_minutes"] = eta_minutes
            if hn is not None: self.rows[i]["hn_full"] = hn
            return self.rows[i]["id"]
        rid = self._seq; self._seq += 1
        self.rows.append({"id": hn or rid, "hn_full": hn, "patient_id": pid, "status": status,
                          "timestamp": timestamp, "eta_minutes": eta_minutes})
        return rid
    def delete(self, pid):
        i = self._find(pid)
        if i >= 0: self.rows.pop(i)

# ---------- UI helpers ----------
class FlowLayout(QtWidgets.QLayout):
    """A layout that arranges widgets in a flowing manner."""

    def __init__(self, parent=None, margin: int = -1, spacing: int | None = None):
        super().__init__(parent)
        self._items: list[QtWidgets.QLayoutItem] = []
        if parent is not None and margin >= 0:
            self.setContentsMargins(margin, margin, margin, margin)
        if spacing is not None:
            self.setSpacing(spacing)

    def addItem(self, item: QtWidgets.QLayoutItem) -> None:
        self._items.append(item)

    def count(self) -> int:
        return len(self._items)

    def itemAt(self, index: int) -> QtWidgets.QLayoutItem | None:
        return self._items[index] if 0 <= index < len(self._items) else None

    def takeAt(self, index: int) -> QtWidgets.QLayoutItem | None:
        if 0 <= index < len(self._items):
            return self._items.pop(index)
        return None

    def expandingDirections(self) -> QtCore.Qt.Orientations:
        return QtCore.Qt.Orientations(QtCore.Qt.Orientation(0))

    def hasHeightForWidth(self) -> bool:
        return True

    def heightForWidth(self, width: int) -> int:
        return self._do_layout(QtCore.QRect(0, 0, width, 0), True)

    def setGeometry(self, rect: QtCore.QRect) -> None:
        super().setGeometry(rect)
        self._do_layout(rect, False)

    def sizeHint(self) -> QtCore.QSize:
        return self.minimumSize()

    def minimumSize(self) -> QtCore.QSize:
        size = QtCore.QSize()
        margins = self.contentsMargins()
        for item in self._items:
            size = size.expandedTo(item.minimumSize())
        size += QtCore.QSize(margins.left() + margins.right(), margins.top() + margins.bottom())
        return size

    def _do_layout(self, rect: QtCore.QRect, test_only: bool) -> int:
        margins = self.contentsMargins()
        effective = rect.adjusted(margins.left(), margins.top(), -margins.right(), -margins.bottom())
        x = effective.x()
        y = effective.y()
        line_height = 0
        max_width = effective.width()

        space_x = self.spacing()
        if space_x < 0:
            space_x = self.smartSpacing(QtWidgets.QStyle.PM_LayoutHorizontalSpacing)
        space_y = self.spacing()
        if space_y < 0:
            space_y = self.smartSpacing(QtWidgets.QStyle.PM_LayoutVerticalSpacing)

        for item in self._items:
            widget = item.widget()
            if widget and not widget.isVisible():
                continue
            size = item.sizeHint()
            next_x = x + size.width()
            if (next_x - effective.x() > max_width) and line_height > 0:
                x = effective.x()
                y = y + line_height + space_y
                next_x = x + size.width()
                line_height = 0
            if not test_only:
                item.setGeometry(QtCore.QRect(QtCore.QPoint(x, y), size))
            x = next_x + space_x
            line_height = max(line_height, size.height())

        return y + line_height - rect.y() + margins.top() + margins.bottom()


class SimpleToast(QtWidgets.QLabel):
    def __init__(self, parent: QtWidgets.QWidget | None = None):
        super().__init__("", parent)
        self.setObjectName("SimpleToast")
        self.setAlignment(QtCore.Qt.AlignCenter)
        self.setStyleSheet(
            "QLabel#SimpleToast{background:rgba(15,23,42,0.92);color:#fff;"
            "padding:8px 14px;border-radius:12px;font-weight:600;}"
        )
        self._timer = QtCore.QTimer(self)
        self._timer.setSingleShot(True)
        self._timer.timeout.connect(self.hide)
        self.hide()

    def show_toast(self, text: str, msec: int = 2200):
        self.setText(text)
        self.adjustSize()
        if self.parent():
            parent_rect = self.parent().rect()
            x = parent_rect.right() - self.width() - 32
            y = parent_rect.bottom() - self.height() - 32
            self.move(max(16, x), max(16, y))
        self.show()
        self.raise_()
        self._timer.start(max(600, msec))

    def smartSpacing(self, pm: QtWidgets.QStyle.PixelMetric) -> int:
        parent = self.parent()
        if parent is None:
            return 0
        if isinstance(parent, QtWidgets.QWidget):
            return parent.style().pixelMetric(pm, None, parent)
        return parent.spacing()


class ShadowButton(QtWidgets.QPushButton):
    def __init__(self, text="", color="#2dd4bf", parent=None):
        super().__init__(text, parent)
        self.base_color = QtGui.QColor(color)
        self.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.setMinimumHeight(36)
        self.setStyleSheet(f"QPushButton{{border:none;color:white;padding:6px 10px;border-radius:10px;font-weight:600;background:{self.base_color.name()};}}")
        sh = QtWidgets.QGraphicsDropShadowEffect(self); sh.setBlurRadius(14); sh.setXOffset(0); sh.setYOffset(4); sh.setColor(QtGui.QColor(0,0,0,64))
        self.setGraphicsEffect(sh)


class PostOpDialog(QtWidgets.QDialog):
    def __init__(self, entry: _SchedEntry, parent: QtWidgets.QWidget | None = None):
        super().__init__(parent)
        self.entry = entry
        self.specialty_key = (entry.dept or "Surgery").strip() or "Surgery"
        self.setWindowTitle(f"บันทึกหลังผ่าตัด — HN {entry.hn}")
        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(12)

        grid = QtWidgets.QGridLayout()
        grid.setHorizontalSpacing(8)
        grid.setVerticalSpacing(8)

        grid.addWidget(QtWidgets.QLabel("Assist 1"), 0, 0)
        self.assist1 = make_search_combo(SCRUB_NURSES)
        self.assist1.setEditText(entry.assist1)
        grid.addWidget(self.assist1, 0, 1)

        grid.addWidget(QtWidgets.QLabel("Assist 2"), 0, 2)
        self.assist2 = make_search_combo(SCRUB_NURSES)
        self.assist2.setEditText(entry.assist2)
        grid.addWidget(self.assist2, 0, 3)

        grid.addWidget(QtWidgets.QLabel("Scrub"), 1, 0)
        self.scrub = make_search_combo(SCRUB_NURSES)
        self.scrub.setEditText(entry.scrub)
        grid.addWidget(self.scrub, 1, 1)

        grid.addWidget(QtWidgets.QLabel("Circulate"), 1, 2)
        self.circulate = make_search_combo(SCRUB_NURSES)
        self.circulate.setEditText(entry.circulate)
        grid.addWidget(self.circulate, 1, 3)

        row = 2
        op_label = QtWidgets.QLabel("Operation (หลังผ่าตัด)")
        grid.addWidget(op_label, row, 0, 1, 4)
        row += 1
        self.op_adder = SearchSelectAdder(
            "ค้นหา/เลือก Operation...",
            suggestions=operation_suggestions(self.specialty_key),
        )
        for op_text in (entry.ops or []):
            self.op_adder.list.addItem(op_text)
        grid.addWidget(self.op_adder, row, 0, 1, 4)
        row += 1

        dx_label = QtWidgets.QLabel("Diagnosis (หลังผ่าตัด)")
        grid.addWidget(dx_label, row, 0, 1, 4)
        row += 1
        self.dx_adder = SearchSelectAdder(
            "ค้นหา ICD-10 ...",
            suggestions=diagnosis_suggestions(self.specialty_key, entry.ops or []),
        )
        for dx_text in (entry.diags or []):
            self.dx_adder.list.addItem(dx_text)
        grid.addWidget(self.dx_adder, row, 0, 1, 4)

        self.op_adder.itemsChanged.connect(self._refresh_dx_suggest)

        layout.addLayout(grid)

        btn = QtWidgets.QPushButton("💾 บันทึกหลังผ่าตัด")
        btn.setProperty("variant", "primary")
        btn.clicked.connect(self.accept)
        layout.addWidget(btn, 0, QtCore.Qt.AlignRight)

    def _refresh_dx_suggest(self, _items: list[str]):
        suggestions = diagnosis_suggestions(self.specialty_key, self.op_adder.items())
        self.dx_adder.set_suggestions(suggestions)

    def values(self) -> Dict:
        return {
            "assist1": self.assist1.currentText().strip(),
            "assist2": self.assist2.currentText().strip(),
            "scrub": self.scrub.currentText().strip(),
            "circulate": self.circulate.currentText().strip(),
            "ops": self.op_adder.items(),
            "diags": self.dx_adder.items(),
        }

class Card(QtWidgets.QFrame):
    def __init__(self, title="", parent=None):
        super().__init__(parent)
        self.setObjectName("Card")
        self.setStyleSheet("""
            QFrame#Card { background:#ffffff; border-radius:14px; border:1px solid #e6e6ef; }
            QLabel[role="title"]{ font-size:14px; font-weight:800; color:#0f172a; }
        """)
        lay = QtWidgets.QVBoxLayout(self)
        lay.setContentsMargins(10,10,10,10); lay.setSpacing(8)
        self.title = QtWidgets.QLabel(title); self.title.setProperty("role","title")
        lay.addWidget(self.title)
        self.body = QtWidgets.QWidget()
        self._grid = QtWidgets.QGridLayout(); self._grid.setContentsMargins(0,0,0,0)
        self._grid.setHorizontalSpacing(6); self._grid.setVerticalSpacing(6)
        self.body.setLayout(self._grid); lay.addWidget(self.body)

        shadow = QtWidgets.QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(20); shadow.setXOffset(0); shadow.setYOffset(6)
        shadow.setColor(QtGui.QColor(15,23,42,36))
        self.setGraphicsEffect(shadow)

    def grid(self): return self.body.layout()

class GlassCard(QtWidgets.QFrame):
    def __init__(self, title: str, subtitle: str = "", icon: str = "",
                 parent=None, accent: str = "#2563eb", header_bg: str | None = None):
        super().__init__(parent)
        self.setObjectName("GlassCard")
        header_fill = header_bg or _rgba(accent, 0.12)
        self.setStyleSheet(f"""
            QFrame#GlassCard {{
                background:#ffffff;
                border-radius:14px;
                border:1px solid #e6e6ef;
            }}
            QLabel[role="card-title"]{{ font-size:13.5px; font-weight:800; color:#0f172a; }}
            QLabel[role="card-sub"]  {{ font-size:11px;  color:#64748b; }}
        """)
        shadow = QtWidgets.QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(20); shadow.setXOffset(0); shadow.setYOffset(6)
        shadow.setColor(QtGui.QColor(15,23,42,36)); self.setGraphicsEffect(shadow)

        lay = QtWidgets.QVBoxLayout(self); lay.setContentsMargins(10,10,10,10); lay.setSpacing(8)
        headerFrame = QtWidgets.QFrame(); headerFrame.setObjectName("HeaderCapsule")
        headerFrame.setStyleSheet(f"""
            QFrame#HeaderCapsule {{
                background:{header_fill};
                border:1px solid #e2e8f0;
                border-radius:10px;
                padding:6px 10px;
                border-left:6px solid {accent};
            }}
        """)
        hh = QtWidgets.QHBoxLayout(headerFrame); hh.setContentsMargins(8,4,8,4); hh.setSpacing(8)
        badge = QtWidgets.QLabel(icon or "•"); badge.setFixedWidth(18); hh.addWidget(badge, 0)
        tbox = QtWidgets.QVBoxLayout(); tbox.setSpacing(0)
        ttl = QtWidgets.QLabel(title); ttl.setProperty("role","card-title")
        sub = QtWidgets.QLabel(subtitle); sub.setProperty("role","card-sub")
        tbox.addWidget(ttl); tbox.addWidget(sub); hh.addLayout(tbox, 1)
        lay.addWidget(headerFrame)
        self.body = QtWidgets.QWidget()
        self.grid = QtWidgets.QGridLayout(self.body)
        self.grid.setContentsMargins(0,0,0,0); self.grid.setHorizontalSpacing(6); self.grid.setVerticalSpacing(6)
        lay.addWidget(self.body)

from PySide6.QtGui import QColor
def _rgba(hex_color: str, a: float) -> str:
    c = QColor(hex_color)
    return f"rgba({c.red()},{c.green()},{c.blue()},{a})"

class ElevatedCard(QtWidgets.QFrame):
    def __init__(self, title: str, icon: str = "📦",
                 accent: str = "#2563eb", bg: str = "#ffffff",
                 header_bg: str | None = None, parent=None):
        super().__init__(parent)
        self.setObjectName("ElevatedCard")
        header_fill = header_bg or _rgba(accent, 0.12)
        self.setStyleSheet(f"""
            QFrame#ElevatedCard {{
                background:{bg};
                border-radius:14px;
                border:1px solid #e6e6ef;
            }}
            QLabel[role="x-title"] {{ font-size:14.5px; font-weight:900; color:#0f172a; }}
            QFrame#XHeader {{
                background:{header_fill};
                border:1px solid #e2e8f0;
                border-radius:10px; padding:6px 10px;
                border-left:6px solid {accent};
            }}
            QLabel#XBadge {{
                background:#ffffff; border:1px solid #e5e7eb; border-radius:9px;
                min-width:18px; max-width:18px; min-height:18px; max-height:18px;
                qproperty-alignment: 'AlignCenter';
            }}
        """)
        shadow = QtWidgets.QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(24); shadow.setXOffset(0); shadow.setYOffset(10)
        shadow.setColor(QtGui.QColor(15,23,42,36)); self.setGraphicsEffect(shadow)
        lay = QtWidgets.QVBoxLayout(self); lay.setContentsMargins(10,10,10,10); lay.setSpacing(10)
        header = QtWidgets.QFrame(); header.setObjectName("XHeader")
        hl = QtWidgets.QHBoxLayout(header); hl.setContentsMargins(10,6,10,6); hl.setSpacing(10)
        badge = QtWidgets.QLabel(icon or "•"); badge.setObjectName("XBadge")
        hl.addWidget(badge, 0, QtCore.Qt.AlignVCenter)
        ttl = QtWidgets.QLabel(title); ttl.setProperty("role", "x-title"); hl.addWidget(ttl, 1, QtCore.Qt.AlignVCenter)
        lay.addWidget(header)
        self.body = QtWidgets.QWidget()
        self._grid = QtWidgets.QGridLayout(self.body)
        self._grid.setContentsMargins(0,0,0,0); self._grid.setHorizontalSpacing(6); self._grid.setVerticalSpacing(6)
        lay.addWidget(self.body)
    def grid(self): return self._grid

class ElideDelegate(QtWidgets.QStyledItemDelegate):
    def __init__(self, mode=QtCore.Qt.ElideRight, parent=None):
        super().__init__(parent); self._mode = mode
    def paint(self, painter, option, index):
        option = QtWidgets.QStyleOptionViewItem(option); self.initStyleOption(option, index)
        option.textElideMode = self._mode; option.features &= ~QtWidgets.QStyleOptionViewItem.WrapText
        super().paint(painter, option, index)

# ---------- Schedule delegate (wrap + watermark + column lines) ----------
class ScheduleDelegate(QtWidgets.QStyledItemDelegate):
    def __init__(self, tree: QtWidgets.QTreeWidget):
        super().__init__(tree)
        self._tree = tree

    def paint(self, painter, option, index):
        super().paint(painter, option, index)

        try:
            item = self._tree.itemFromIndex(index)
            is_child = bool(item and item.parent() is not None)
            if is_child and index.column() < (self._tree.columnCount() - 1):
                painter.save()
                painter.setPen(QtGui.QPen(QtGui.QColor("#eef2f7")))
                x = option.rect.right()
                painter.drawLine(x, option.rect.top(), x, option.rect.bottom())
                painter.restore()
        except Exception:
            pass

# ---------- Icon helpers ----------

def _read_png_safe(path: Path) -> QPixmap:
    f = QtCore.QFile(str(path))
    if not f.open(QtCore.QIODevice.ReadOnly): return QPixmap()
    rd = QImageReader(f, b"png"); img = rd.read(); f.close()
    return QPixmap.fromImage(img) if not img.isNull() else QPixmap()

def _icon_from_png(p: Path) -> QIcon:
    f = QtCore.QFile(str(p))
    if not f.open(QtCore.QIODevice.ReadOnly): return QIcon()
    rd = QImageReader(f, b"png"); img = rd.read(); f.close()
    return QIcon(QPixmap.fromImage(img)) if not img.isNull() else QIcon()

def _draw_fallback_icon(size=256) -> QIcon:
    pm = QPixmap(size, size); pm.fill(QtCore.Qt.transparent); p = QPainter(pm); p.setRenderHint(QPainter.Antialiasing, True)
    grad = QLinearGradient(0,0,size,size); grad.setColorAt(0.0,QColor("#eaf2ff")); grad.setColorAt(0.6,QColor("#e6f7ff")); grad.setColorAt(1.0,QColor("#eefcf8"))
    p.setBrush(grad); p.setPen(QtCore.Qt.NoPen); p.drawEllipse(8,8,size-16,size-16); p.end(); return QIcon(pm)

def _load_app_icon() -> QIcon:
    here = Path(__file__).resolve().parent; assets = here / "assets"
    for p in [assets/"app.ico", here/"app.ico", assets/"app.png", here/"app.png"]:
        if p.exists():
            if p.suffix.lower()==".ico":
                ico = QIcon(str(p))
                if not ico.isNull(): return ico
            else:
                ico = _icon_from_png(p)
                if not ico.isNull(): return ico
    return _draw_fallback_icon(256)

# ---------- Banner ----------
class HeroBanner(QtWidgets.QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._title = "SurgiBot Client — Operating Room Nongbualamphu Hospital"
        self.setMinimumHeight(56); self.setMaximumHeight(60)
        lay = QtWidgets.QHBoxLayout(self); lay.setContentsMargins(10,8,10,8); lay.setSpacing(8)
        self.logoLabel = QtWidgets.QLabel(); self.logoLabel.setFixedSize(34,34); lay.addWidget(self.logoLabel, 0, QtCore.Qt.AlignVCenter)
        self.lblTitle = QtWidgets.QLabel(self._title); f = self.lblTitle.font(); f.setPointSize(14); f.setBold(True); self.lblTitle.setFont(f)
        lay.addWidget(self.lblTitle,1)
        self.rightBox = QtWidgets.QHBoxLayout(); self.rightBox.setSpacing(6); lay.addLayout(self.rightBox,0)

    def setTitle(self, text: str): self.lblTitle.setText(text)
    def setRight(self, widget: QtWidgets.QWidget):
        while self.rightBox.count():
            it = self.rightBox.takeAt(0); w = it.widget()
            if w: w.setParent(None)
        self.rightBox.addWidget(widget)
    def setLogo(self, path: Union[Path, str], size: int = 34, radius: int = 8):
        p = Path(path); f = QtCore.QFile(str(p))
        if not f.open(QtCore.QIODevice.ReadOnly): self.logoLabel.clear(); return
        rd = QImageReader(f, b"png"); img = rd.read(); f.close()
        if img.isNull(): self.logoLabel.clear(); return
        pm = QPixmap.fromImage(img).scaled(size,size,QtCore.Qt.KeepAspectRatioByExpanding,QtCore.Qt.SmoothTransformation)
        canvas = QPixmap(size,size); canvas.fill(QtCore.Qt.transparent); painter = QPainter(canvas); painter.setRenderHint(QPainter.Antialiasing,True)
        pathp = QtGui.QPainterPath(); pathp.addRoundedRect(0,0,size,size,radius,radius); painter.setClipPath(pathp); painter.drawPixmap(0,0,pm); painter.end()
        self.logoLabel.setPixmap(canvas)


class WaveBanner(QtWidgets.QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("WaveBanner")
        self.setMinimumHeight(90)
        self._t = 0.0

        self._timer = QtCore.QTimer(self)
        self._timer.timeout.connect(self._tick)
        self._timer.start(16)

        shadow = QtWidgets.QGraphicsDropShadowEffect(blurRadius=24, xOffset=0, yOffset=8)
        shadow.setColor(QtGui.QColor(15, 23, 42, 40))
        self.setGraphicsEffect(shadow)

        self.setStyleSheet(
            """
            #WaveBanner {
                border:1px solid #dbeafe;
                border-radius:14px;
                background: transparent;
            }
            QLabel#Title { font-weight:800; font-size:20px; color:#0f172a; }
            QLabel#Pill  {
                background:#ffffff;
                border:1px solid #e5e7eb;
                border-radius:10px;
                padding:4px 10px;
                font-weight:600;
            }
            QPushButton#Reconnect { color:white; background:#475569; border:none; border-radius:10px; padding:6px 12px; }
            QPushButton#Health    { color:white; background:#3b82f6; border:none; border-radius:10px; padding:6px 12px; }
            QPushButton#Settings  { color:white; background:#06b6d4; border:none; border-radius:10px; padding:6px 12px; }
            QPushButton#Reconnect:hover { background:#334155; }
            QPushButton#Health:hover    { background:#2563eb; }
            QPushButton#Settings:hover  { background:#0891b2; }
            """
        )

        lay = QtWidgets.QHBoxLayout(self)
        lay.setContentsMargins(18, 12, 18, 14)
        lay.setSpacing(12)

        logo = QtWidgets.QLabel()
        here = Path(__file__).resolve().parent
        logo_path = here / "MascotAlert.png"
        if logo_path.exists():
            pm = QPixmap(str(logo_path)).scaled(34, 34, QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation)
        else:
            pm = QtGui.QPixmap(34, 34)
            pm.fill(QtCore.Qt.transparent)
            painter = QtGui.QPainter(pm)
            painter.setRenderHint(QtGui.QPainter.Antialiasing)
            painter.setBrush(QtGui.QColor("#64748b"))
            painter.setPen(QtCore.Qt.NoPen)
            painter.drawEllipse(0, 0, 34, 34)
            painter.end()
        logo.setPixmap(pm)

        title = QtWidgets.QLabel("SurgiBot Client — Operating Room Nongbualamphu Hospital")
        title.setObjectName("Title")

        lay.addWidget(logo)
        lay.addWidget(title)
        lay.addStretch(1)

        self._pill_base = "background:#ffffff;border:1px solid #e5e7eb;border-radius:10px;padding:4px 10px;font-weight:600;"

        self.status_label = QtWidgets.QLabel("  • Offline  ")
        self.status_label.setObjectName("Pill")
        self.status_label.setStyleSheet(f"{self._pill_base}color:#ef4444;")

        self.btn_reconnect = QtWidgets.QPushButton("Reconnect")
        self.btn_reconnect.setObjectName("Reconnect")
        self.btn_health = QtWidgets.QPushButton("Health")
        self.btn_health.setObjectName("Health")
        self.btn_settings = QtWidgets.QPushButton("Settings")
        self.btn_settings.setObjectName("Settings")

        for w in (self.status_label, self.btn_reconnect, self.btn_health, self.btn_settings):
            lay.addWidget(w)

    def pill_base_style(self) -> str:
        return self._pill_base

    def _tick(self):
        self._t += 0.03
        self.update()

    def paintEvent(self, event: QtGui.QPaintEvent):
        r = self.rect()
        painter = QtGui.QPainter(self)
        painter.setRenderHint(QtGui.QPainter.Antialiasing)

        path = QtGui.QPainterPath()
        path.addRoundedRect(r.adjusted(0, 0, -1, -1), 14, 14)
        painter.setClipPath(path)

        grad = QtGui.QLinearGradient(r.topLeft(), r.topRight())
        grad.setColorAt(0.0, QtGui.QColor("#eef2ff"))
        grad.setColorAt(1.0, QtGui.QColor("#e0f2fe"))
        painter.fillRect(r, grad)

        def draw_wave(ampl, wave_len, phase, color, height_ratio):
            y0 = r.height() * height_ratio
            wave = QtGui.QPainterPath()
            wave.moveTo(r.left(), r.bottom())
            wave.lineTo(r.left(), y0)
            x = r.left()
            while x <= r.right():
                y = y0 - ampl * math.sin((x / wave_len) + self._t * phase)
                wave.lineTo(x, y)
                x += 3
            wave.lineTo(r.right(), r.bottom())
            painter.fillPath(wave, QtGui.QColor(color))

        draw_wave(14, 55.0, 1.2, "#c7d2fe", 0.86)
        draw_wave(10, 75.0, 0.9, "#93c5fd", 0.78)

        pen = QtGui.QPen(QtGui.QColor("#dbeafe"))
        pen.setWidth(1)
        painter.setPen(pen)
        painter.drawPath(path)

# ---------- Main ----------
class Main(QtWidgets.QWidget):
    def __init__(self, host, port, token):
        super().__init__()
        self.cli = SurgiBotClientHTTP(host, port, token)
        self.model = LocalTableModel()
        self.rows_cache = []
        self.sched = SharedScheduleModel()
        self.ws: QWebSocket|None = None
        self.ws_connected = False
        self.tray = None
        self._last_states = {}
        self._last_selected_uid = ""
        self._suppress_status_change = False
        self.toast = SimpleToast(self)
        self._thread_pool = QtCore.QThreadPool.globalInstance()
        self._refresh_timer = QtCore.QTimer(self)
        self._refresh_timer.setSingleShot(True)
        self._refresh_timer.timeout.connect(self._start_refresh_task)
        self._refresh_inflight = False
        self._refresh_requested = False

        # Monitor knowledge
        self.monitor_ready = False
        self._was_in_monitor: set[str] = set()
        self._current_monitor_hn: set[str] = set()

        # Connection hysteresis
        self._conn_ok_streak = 0
        self._conn_fail_streak = 0
        self._conn_state: bool | None = None

        # Throttled persistence & scheduled rendering
        self._persist_timer = QtCore.QTimer(self)
        self._persist_timer.setSingleShot(True)
        self._persist_timer.setInterval(7000)
        self._persist_timer.timeout.connect(lambda: self._save_persisted_monitor_state(self.rows_cache))

        self._render_pending = False
        self._render_timer = QtCore.QTimer(self)
        self._render_timer.setSingleShot(True)
        self._render_timer.setInterval(150)
        self._render_timer.timeout.connect(self._flush_scheduled_render)

        self._last_autofit_at = 0.0

        self.setWindowTitle("SurgiBot Client — Modern (PySide6)")
        self.resize(1440, 900)
        self._build_ui()
        self._load_settings()

        # ---------- load persisted monitor state BEFORE first refresh ----------
        self._load_persisted_monitor_state()

        if self.rows_cache:
            self.monitor_ready = True
            self._rebuild(self.rows_cache)

        # Barcode
        self.scan_enabled = True; self._scan_buf = ""; self._scan_timeout_ms = 120
        self._scan_timer = QtCore.QTimer(self); self._scan_timer.setSingleShot(True); self._scan_timer.timeout.connect(self._finalize_scan_if_any)
        self.installEventFilter(self)

        self._ensure_tray()
        self._refresh(prefer_server=True)

        self._time_tick = QtCore.QTimer(self)
        self._time_tick.timeout.connect(self._update_time_cells)
        self._time_tick.start(1000)
        self._pull = QtCore.QTimer(self); self._pull.timeout.connect(lambda: self._refresh(True)); self._pull.start(CONFIG.client_refresh_interval_ms)
        self._sched_timer = QtCore.QTimer(self); self._sched_timer.timeout.connect(self._check_schedule_seq); self._sched_timer.start(1000)
        self._start_websocket()

    # ---------- Settings dialog ----------
    def _open_settings_dialog(self):
        dlg = QtWidgets.QDialog(self)
        dlg.setWindowTitle("Server Settings")
        lay = QtWidgets.QVBoxLayout(dlg)

        form = QFormLayout()
        host = QtWidgets.QLineEdit(self.ent_host.text())
        port = QtWidgets.QLineEdit(self.ent_port.text())
        token = QtWidgets.QLineEdit(self.ent_token.text()); token.setEchoMode(QtWidgets.QLineEdit.Password)
        form.addRow("Host", host)
        form.addRow("Port", port)
        form.addRow("Token", token)
        lay.addLayout(form)

        btns = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Save | QtWidgets.QDialogButtonBox.Cancel)
        lay.addWidget(btns)
        def _save():
            self.ent_host.setText(host.text().strip())
            self.ent_port.setText(port.text().strip())
            self.ent_token.setText(token.text().strip())
            self._save_settings()
            self._on_reconnect_clicked()
            dlg.accept()
        btns.accepted.connect(_save)
        btns.rejected.connect(dlg.reject)
        dlg.exec()

    def _capture_or_expand_state(self):
        try:
            st = {}
            topc = self.tree_sched.topLevelItemCount()
            for i in range(topc):
                it = self.tree_sched.topLevelItem(i)
                key = self._or_item_label(it)
                if key:
                    st[key] = it.isExpanded()
            self._or_expand_state = st
        except Exception:
            pass

    def _apply_or_expand_state(self, item: QtWidgets.QTreeWidgetItem):
        key = self._or_item_label(item)
        expanded = self._or_expand_state.get(key, True)
        item.setExpanded(bool(expanded))

    def _or_item_label(self, item: QtWidgets.QTreeWidgetItem | None) -> str:
        if item is None:
            return ""
        text = (item.text(0) or "").strip()
        if not text:
            cached_title = item.data(0, QtCore.Qt.UserRole + 201)
            if cached_title:
                text = str(cached_title).strip()
        if not text:
            cached_or = item.data(0, QtCore.Qt.UserRole + 200)
            if cached_or:
                text = str(cached_or).strip()
        return text

    def _or_card_widget(self, title: str, accent: str) -> QtWidgets.QWidget:
        w = QtWidgets.QFrame(); w.setObjectName("OrCard")
        c = QtGui.QColor(accent)
        dark = c.darker(130).name(); mid = c.name(); bar = c.lighter(110).name()
        w.setStyleSheet(f"""
        QFrame#OrCard {{
            background: qlineargradient(x1:0,y1:0, x2:0,y2:1, stop:0 {dark}, stop:1 {mid});
            border-radius: 12px; border: 1px solid rgba(255,255,255,0.20);
        }}
        QLabel[role="or-title"] {{ color:#fff; font-weight:900; font-size:15px; }}
        QLabel[role="or-sub"]   {{ color:rgba(255,255,255,0.90); font-weight:600; font-size:11px; }}
        """)
        w.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        w.setMinimumHeight(44)
        lay = QtWidgets.QHBoxLayout(w); lay.setContentsMargins(12, 8, 12, 8); lay.setSpacing(10)
        barf = QtWidgets.QFrame(); barf.setFixedWidth(6); barf.setStyleSheet(f"background:{bar}; border-radius:3px;")
        lay.addWidget(barf, 0, QtCore.Qt.AlignVCenter)
        box = QtWidgets.QVBoxLayout(); box.setSpacing(0)
        lbl = QtWidgets.QLabel(title); lbl.setProperty("role", "or-title"); lbl.setWordWrap(False)
        lbl.setMinimumWidth(140); lbl.setAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)
        sub = QtWidgets.QLabel("ห้องผ่าตัด"); sub.setProperty("role", "or-sub"); sub.setWordWrap(False); sub.setAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)
        box.addWidget(lbl); box.addWidget(sub); lay.addLayout(box, 1)
        shadow = QtWidgets.QGraphicsDropShadowEffect(w); shadow.setBlurRadius(24); shadow.setXOffset(0); shadow.setYOffset(8); shadow.setColor(QtGui.QColor(15, 23, 42, 48))
        w.setGraphicsEffect(shadow)
        return w

    def _style_or_group_header(self, item: QtWidgets.QTreeWidgetItem, bg_hex: str = "#eef2ff"):
        try:
            cols = self.tree_sched.columnCount()
            item.setFlags((item.flags() | QtCore.Qt.ItemIsEnabled) & ~QtCore.Qt.ItemIsSelectable)
            f = self.tree_sched.font(); f.setBold(True); f.setPointSize(f.pointSize() + 1)
            fg = QtGui.QBrush(QtGui.QColor("#1e293b")); bg = QtGui.QBrush(QtGui.QColor(bg_hex))
            for c in range(cols):
                item.setFont(c, f); item.setForeground(c, fg); item.setBackground(c, bg)
            item.setSizeHint(0, QtCore.QSize(item.sizeHint(0).width(), 34))
        except Exception:
            pass

    # ------ Header pulse helpers ------
    def _ensure_sched_pulser(self):
        if hasattr(self, "_sched_pulser"): return
        self._sched_pulser = {"items": [], "phase": 0}
        self._sched_timer2 = QtCore.QTimer(self)
        self._sched_timer2.timeout.connect(self._tick_sched_pulse)
        self._sched_timer2.start(60)

    def _clear_sched_pulser(self):
        if hasattr(self, "_sched_pulser"):
            self._sched_pulser["items"].clear()

    def _register_or_header_for_pulse(self, item: QtWidgets.QTreeWidgetItem, color_hex: str):
        self._ensure_sched_pulser()
        base = QtGui.QColor(color_hex)
        f = self.tree_sched.font(); f.setBold(True); item.setFont(0, f)
        item.setForeground(0, QtGui.QBrush(base.darker(140)))
        self._sched_pulser["items"].append((item, base))

    def _tick_sched_pulse(self):
        if not hasattr(self, "_sched_pulser"): return
        import math
        self._sched_pulser["phase"] = (self._sched_pulser["phase"] + 1) % 120
        k = (1.0 + math.sin(self._sched_pulser["phase"] / 120.0 * 2.0 * math.pi)) * 0.5
        alive_items = []
        for item, base in list(self._sched_pulser["items"]):
            try:
                _ = item.text(0)
            except RuntimeError:
                continue
            if item.treeWidget() is None:
                continue
            bg = QtGui.QColor(base); bg.setAlpha(int(40 + k * 80))
            brush = QtGui.QBrush(bg)
            try:
                for c in range(self.tree_sched.columnCount()):
                    item.setBackground(c, brush)
                alive_items.append((item, base))
            except RuntimeError:
                continue
        self._sched_pulser["items"] = alive_items
        self.tree_sched.viewport().update()

    # ----------- Monitor helpers -----------
    def _extract_hn_from_row(self, r: dict) -> str:
        hn = str(r.get("hn_full") or "").strip()
        if hn and hn.isdigit() and len(hn) == 9:
            return hn
        _id = str(r.get("id") or "").strip()
        if _id.isdigit() and len(_id) == 9:
            return _id
        return ""

    def _is_hn_in_monitor(self, hn: str) -> bool:
        if not hn: return False
        return hn in self._current_monitor_hn

    def _should_auto_purge(self, row: dict) -> bool:
        st = str(row.get("status") or "")
        if st not in AUTO_PURGE_STATUSES:
            return False
        ts = _parse_iso(row.get("timestamp"))
        if not ts:
            return False
        return (datetime.now() - ts) >= timedelta(minutes=AUTO_PURGE_MINUTES)

    # ----------- UI reactions -----------
    def _on_sched_item_clicked_from_selection(self):
        it = self.tree_sched.currentItem()
        if it is not None:
            self._on_sched_item_clicked(it, 0)

    def _on_sched_item_clicked(self, item: QtWidgets.QTreeWidgetItem, column: int):
        try:
            if item is None or item.parent() is None:
                return
            if not (item.flags() & QtCore.Qt.ItemIsEnabled):
                return

            hn = (item.text(3) or "").strip()
            entry = item.data(0, QtCore.Qt.UserRole)
            if hn and hn.isdigit() and len(hn) == 9:
                self.ent_hn.setText(hn)

            if isinstance(entry, _SchedEntry):
                self._last_selected_uid = entry.uid()
                self._set_status_combo(entry.status or None)
                self._auto_assign_or_queue_by_doctor(entry)
            else:
                self._last_selected_uid = ""

            or_room = (item.parent().text(0) or "").strip()
            if or_room:
                i = self.cb_or.findText(or_room)
                if i >= 0: self.cb_or.setCurrentIndex(i)

            q_raw = (item.text(19) or "").strip()
            if q_raw.startswith("0-"):
                qi = self.cb_q.findText(q_raw)
                if qi >= 0:
                    self.cb_q.setCurrentIndex(qi)
            elif q_raw.isdigit():
                q_label = f"0-{q_raw}"
                qi = self.cb_q.findText(q_label)
                if qi >= 0:
                    self.cb_q.setCurrentIndex(qi)

            if self._is_hn_in_monitor(hn): self.rb_edit.setChecked(True)
            else: self.rb_add.setChecked(True)

            self._update_action_styles()
            self.cb_status.setFocus()
        except Exception:
            pass

    def _auto_assign_or_queue_by_doctor(self, entry: "_SchedEntry") -> bool:
        """Auto-select OR/Queue based on doctor order within the same OR and day."""
        try:
            if not isinstance(entry, _SchedEntry):
                return False
            if not hasattr(self, "sched") or not hasattr(self, "cb_or") or not hasattr(self, "cb_q"):
                return False

            or_room = (entry.or_room or "").strip()
            doctor = (entry.doctor or "").strip()
            target_hn = (entry.hn or "").strip()
            if not or_room or not doctor or not target_hn:
                return False

            target_date = entry.date_obj
            if not target_date:
                target_date = _parse_date(entry.date) if getattr(entry, "date", None) else None
            if not target_date:
                target_date = datetime.now().date()

            def _coerce_date(candidate: "_SchedEntry"):
                if getattr(candidate, "date_obj", None):
                    return candidate.date_obj
                raw = getattr(candidate, "date", "")
                return _parse_date(raw) or target_date

            def _queue_value(candidate: "_SchedEntry") -> int:
                try:
                    return int(getattr(candidate, "queue", 0) or 0)
                except Exception:
                    return 0

            def _time_minutes(candidate: "_SchedEntry") -> int:
                txt = str(getattr(candidate, "time", "") or "").strip()
                if not txt:
                    return 24 * 60 + 1
                parts = txt.split(":")
                try:
                    hh = int(parts[0])
                    mm = int(parts[1]) if len(parts) > 1 else 0
                except Exception:
                    return 24 * 60 + 1
                hh = max(0, min(23, hh))
                mm = max(0, min(59, mm))
                return hh * 60 + mm

            same_entries: list[_SchedEntry] = []
            for candidate in getattr(self.sched, "entries", []):
                if not isinstance(candidate, _SchedEntry):
                    continue
                if (candidate.or_room or "").strip() != or_room:
                    continue
                if (candidate.doctor or "").strip() != doctor:
                    continue
                if _coerce_date(candidate) != target_date:
                    continue
                same_entries.append(candidate)

            if not same_entries:
                return False

            same_entries.sort(
                key=lambda cand: (
                    0,
                    _queue_value(cand),
                    _time_minutes(cand),
                )
                if _queue_value(cand) > 0
                else (
                    1,
                    _time_minutes(cand),
                    cand.uid(),
                )
            )

            position: int | None = None
            for idx, candidate in enumerate(same_entries, start=1):
                if (candidate.hn or "").strip() == target_hn:
                    position = idx
                    break

            if position is None:
                position = len(same_entries)

            idx_or = self.cb_or.findText(or_room)
            if idx_or < 0:
                self.cb_or.addItem(or_room)
                idx_or = self.cb_or.findText(or_room)
            if idx_or >= 0:
                self.cb_or.setCurrentIndex(idx_or)

            queue_label = f"0-{position}"
            idx_queue = self.cb_q.findText(queue_label)
            if idx_queue < 0:
                self.cb_q.addItem(queue_label)
                idx_queue = self.cb_q.findText(queue_label)
            if idx_queue >= 0:
                self.cb_q.setCurrentIndex(idx_queue)

            return True
        except Exception:
            return False

    def _make_form_label(self, text: str) -> QtWidgets.QLabel:
        lbl = QtWidgets.QLabel(text)
        lbl.setAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
        lbl.setStyleSheet("padding-right: 10px; color:#0f172a;")
        return lbl

    def _autofit_schedule_columns(self):
        tree = getattr(self, "tree_sched", None)
        if tree is None:
            return
        hdr = tree.header()
        if hdr is None:
            return
        now = time.monotonic()
        if now - getattr(self, "_last_autofit_at", 0.0) < 10.0:
            return
        self._last_autofit_at = now
        hbar = tree.horizontalScrollBar()
        vbar = tree.verticalScrollBar()
        old_h = hbar.value() if hbar is not None else 0
        old_v = vbar.value() if vbar is not None else 0
        hdr.setStretchLastSection(False)
        for c in (0, 1, 2, 4, 9, 19):
            try:
                tree.resizeColumnToContents(c)
            except Exception:
                break
        if hbar is not None or vbar is not None:
            def _restore_after_autofit():
                if hbar is not None:
                    hbar.setValue(min(old_h, hbar.maximum()))
                if vbar is not None:
                    vbar.setValue(min(old_v, vbar.maximum()))

            QtCore.QTimer.singleShot(0, _restore_after_autofit)

    def _build_header_frame(self) -> QtWidgets.QFrame:
        banner = WaveBanner(self)
        self.status_chip = banner.status_label
        self._status_pill_base = banner.pill_base_style()
        self.btn_reconnect = banner.btn_reconnect
        self.btn_health = banner.btn_health
        self.btn_settings = banner.btn_settings

        self.btn_reconnect.clicked.connect(self._on_reconnect_clicked)
        self.btn_health.clicked.connect(self._on_health)
        self.btn_settings.clicked.connect(self._open_settings_dialog)

        return banner

    def _build_ui(self):
        root = QtWidgets.QVBoxLayout(self)
        root.setContentsMargins(10, 6, 10, 10)
        root.setSpacing(12)
        self.setStyleSheet("""
            QWidget { font-family:'Segoe UI','Inter','Noto Sans',system-ui; font-size:12pt; color:#0f172a; }
            QComboBox, QLineEdit { padding:5px 8px; border-radius:8px; border:1px solid #e5e7eb; background:#f8fafc; min-height:32px; }
            QHeaderView::section { background:#f1f5f9; border:none; padding:6px; font-weight:700; color:#0f172a; }
            QTableWidget { background:white; border:1px solid #e6e6ef; border-radius:12px; gridline-color:#e6e6ef; selection-background-color:#e0f2fe; }
            QTableView::item { height:34px; } QTreeView::item { height:34px; }
        """)

        # --- Hidden connection fields (used by Settings dialog & client) ---
        self.ent_host = QtWidgets.QLineEdit(DEFAULT_HOST); self.ent_host.setVisible(False)
        self.ent_port = QtWidgets.QLineEdit(str(DEFAULT_PORT)); self.ent_port.setVisible(False)
        self.ent_token = QtWidgets.QLineEdit(DEFAULT_TOKEN); self.ent_token.setVisible(False)

        header_banner = self._build_header_frame()
        root.addWidget(header_banner)
        self._set_chip(False)

        # --- Top workflow cards (compact with restored headers) ---
        CARD_CSS = """
QGroupBox {
    background:#f8fafc;
    border:1px solid #e5e7eb;
    border-radius:12px;
    padding:22px 12px 12px 12px;
    border-left:4px solid #3b82f6;
}
QGroupBox::title {
    subcontrol-origin: margin;
    left:16px; top:-6px;
    padding:6px 14px;
    border-radius:10px;
    background:qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 #2563eb, stop:1 #1d4ed8);
    color:#fff; font-weight:800; font-size:15px; letter-spacing:.3px;
    border:1px solid #1e40af;
}
QLabel{ color:#0f172a; }
QLineEdit, QComboBox {
    padding:6px 8px; border:1px solid #e5e7eb; border-radius:8px; background:#fff; min-height:30px;
}
QCheckBox { color:#0f172a; }
"""

        top = QtWidgets.QHBoxLayout()
        top.setContentsMargins(12, 0, 12, 8)
        top.setSpacing(16)
        top.setAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignTop)

        # Identify Patient (HN only)
        card_ident = QtWidgets.QGroupBox("1. Identify Patient")
        card_ident.setStyleSheet(CARD_CSS)
        card_ident.setSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Minimum)

        grid_ident = QtWidgets.QGridLayout(card_ident)
        grid_ident.setHorizontalSpacing(10)
        grid_ident.setVerticalSpacing(6)
        grid_ident.setContentsMargins(10, 12, 10, 10)

        lbl_hn = QtWidgets.QLabel("HN (9 หลัก)")
        self.ent_hn = QtWidgets.QLineEdit()
        self.ent_hn.setPlaceholderText("เช่น 590166994")
        self.ent_hn.setMaxLength(9)
        self.ent_hn.setValidator(QtGui.QIntValidator(0, 999999999, self))
        self.ent_hn.setMinimumWidth(160)
        self.ent_pid = QtWidgets.QLineEdit()
        self.ent_pid.setVisible(False)

        grid_ident.addWidget(lbl_hn, 0, 0)
        grid_ident.addWidget(self.ent_hn, 0, 1)

        self.chk_scan = QtWidgets.QCheckBox("โหมดสแกนบาร์โค้ด HN")
        self.chk_scan.setChecked(True)
        self.chk_scan.stateChanged.connect(lambda s: setattr(self, "scan_enabled", bool(s)))
        self.lbl_scan_state = QtWidgets.QLabel("Scanner: Ready")
        self.lbl_scan_state.setStyleSheet("color:#16a34a;font-weight:600;")

        scan_line = QtWidgets.QHBoxLayout()
        scan_line.setContentsMargins(0, 0, 0, 0)
        scan_line.setSpacing(6)
        scan_line.addWidget(self.chk_scan, 0)
        scan_line.addWidget(self.lbl_scan_state, 0)
        scan_line.addStretch(1)
        grid_ident.addLayout(scan_line, 1, 0, 1, 2)

        # Assign Room (OR & Queue same row)
        card_or = QtWidgets.QGroupBox("2. Assign Room")
        card_or.setStyleSheet(CARD_CSS)
        card_or.setSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Minimum)

        grid_or = QtWidgets.QGridLayout(card_or)
        grid_or.setHorizontalSpacing(10)
        grid_or.setVerticalSpacing(6)
        grid_or.setContentsMargins(10, 12, 10, 10)

        lbl_or = QtWidgets.QLabel("OR")
        self.cb_or = QtWidgets.QComboBox()
        self.cb_or.addItems(OR_CHOICES)
        self.cb_or.setSizeAdjustPolicy(QtWidgets.QComboBox.AdjustToContents)
        lbl_q = QtWidgets.QLabel("Queue")
        self.cb_q = QtWidgets.QComboBox()
        self.cb_q.addItems(QUEUE_CHOICES)
        self.cb_q.setSizeAdjustPolicy(QtWidgets.QComboBox.AdjustToContents)

        grid_or.addWidget(lbl_or, 0, 0)
        grid_or.addWidget(self.cb_or, 0, 1)
        grid_or.addWidget(lbl_q, 0, 2)
        grid_or.addWidget(self.cb_q, 0, 3)
        grid_or.setColumnStretch(1, 1)
        grid_or.setColumnStretch(3, 1)

        # Status & Timing
        card_stat = QtWidgets.QGroupBox("3. Status & Timing")
        card_stat.setStyleSheet(CARD_CSS)
        card_stat.setSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Minimum)

        form_stat = QtWidgets.QFormLayout(card_stat)
        form_stat.setContentsMargins(10, 12, 10, 10)
        form_stat.setHorizontalSpacing(10)
        form_stat.setVerticalSpacing(6)
        form_stat.setLabelAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)

        lbl_status = QtWidgets.QLabel("Status")
        self.cb_status = QtWidgets.QComboBox()
        self.cb_status.addItems(STATUS_CHOICES)
        self.cb_status.setSizeAdjustPolicy(QtWidgets.QComboBox.AdjustToContents)
        self.ent_eta = QtWidgets.QLineEdit()
        self.ent_eta.setPlaceholderText("เช่น 90")
        self.lbl_eta = QtWidgets.QLabel("เวลาโดยประมาณในการผ่าตัด (นาที)")

        form_stat.addRow(lbl_status, self.cb_status)
        form_stat.addRow(self.lbl_eta, self.ent_eta)
        self.cb_status.currentTextChanged.connect(self._on_status_combo_changed)
        self._toggle_eta_visibility()

        # Action card
        action = QtWidgets.QGroupBox("Action")
        action.setStyleSheet(CARD_CSS)
        action.setSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Minimum)

        act_layout = QtWidgets.QHBoxLayout(action)
        act_layout.setContentsMargins(10, 12, 10, 10)
        act_layout.setSpacing(8)

        def mk_btn(text, corner):
            btn = QtWidgets.QPushButton(text)
            btn.setCheckable(True)
            btn.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
            radius = {
                "left": "border-top-left-radius:8px;border-bottom-left-radius:8px;",
                "mid": "border-radius:0;",
                "right": "border-top-right-radius:8px;border-bottom-right-radius:8px;",
            }[corner]
            btn.setMinimumHeight(32)
            btn.setMinimumWidth(96)
            btn.setProperty("cornerCSS", radius)
            btn.setStyleSheet(
                f"QPushButton{{padding:6px 10px;border:1px solid #e5e7eb;background:#f8fafc;color:#0f172a;font-weight:700;{radius}}}"
                f"QPushButton:hover{{background:#eef2f7;}}"
            )
            return btn

        self.rb_add = mk_btn("➕ เพิ่ม", "left")
        self.rb_edit = mk_btn("✏️ แก้ไข", "mid")
        self.rb_del = mk_btn("🗑️ ลบ", "right")
        for btn in (self.rb_add, self.rb_edit, self.rb_del):
            act_layout.addWidget(btn)

        self.action_group = QtWidgets.QButtonGroup(self)
        self.action_group.setExclusive(True)
        for btn in (self.rb_add, self.rb_edit, self.rb_del):
            self.action_group.addButton(btn)

        self.rb_add.setChecked(True)
        for btn in (self.rb_add, self.rb_edit, self.rb_del):
            btn.toggled.connect(self._update_action_styles)

        self.btn_send = ShadowButton("🚀 ส่งคำสั่ง", "#10b981")
        self.btn_send.setMinimumWidth(130)
        self.btn_send.clicked.connect(self._on_send)
        act_layout.addWidget(self.btn_send)
        act_layout.addStretch(1)

        for card in (card_ident, card_or, card_stat, action):
            card.setMinimumWidth(220)
            top.addWidget(card)

        top.addStretch(1)
        root.addLayout(top)

        # --- Tabs (Schedule + Monitor) ---
        self.tabs = QtWidgets.QTabWidget()
        self.tabs.setTabPosition(QtWidgets.QTabWidget.North)
        self.tabs.setDocumentMode(True)

        # Schedule
        self.card_sched = ElevatedCard(
            "Result Schedule (Private) — จาก Registry",
            icon="🗂", accent="#0ea5e9", bg="#ffffff", header_bg=_rgba("#0ea5e9", 0.12)
        )
        gs = self.card_sched.grid(); gs.setContentsMargins(0,0,0,0)
        self.tree_sched = QtWidgets.QTreeWidget()
        self.tree_sched.setColumnCount(21)
        self.tree_sched.setHeaderLabels([
            "บันทึก",
            "ช่วงเวลา",
            "OR/เวลา",
            "HN",
            "ชื่อ-สกุล",
            "อายุ",
            "Diagnosis",
            "Operation",
            "แพทย์",
            "Ward",
            "ขนาดเคส",
            "แผนก",
            "Assist1",
            "Assist2",
            "Scrub",
            "Circulate",
            "สถานะ",
            "เริ่ม",
            "จบ",
            "คิว",
            "ประเภทเคส",
        ])
        self.tree_sched.setUniformRowHeights(False)
        hdr = self.tree_sched.header()
        hdr.setStretchLastSection(False)
        hdr.setDefaultAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)
        hdr.setFixedHeight(42)
        for i in range(21):
            hdr.setSectionResizeMode(i, QtWidgets.QHeaderView.ResizeToContents)
        self.tree_sched.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAsNeeded)
        self.tree_sched.setHorizontalScrollMode(QtWidgets.QAbstractItemView.ScrollPerPixel)
        self.tree_sched.setVerticalScrollMode(QtWidgets.QAbstractItemView.ScrollPerPixel)
        self.tree_sched.setTextElideMode(QtCore.Qt.ElideNone)
        self.tree_sched.setWordWrap(False)
        self.tree_sched.setObjectName("ScheduleTree")
        self.tree_sched.setAlternatingRowColors(True)
        self.tree_sched.setStyleSheet("""
        QTreeWidget#ScheduleTree {
            background:#ffffff; border:1px solid #e6e6ef; border-radius:12px; gridline-color:#e6e6ef;
        }
        QTreeWidget#ScheduleTree QHeaderView::section {
            background: qlineargradient(x1:0,y1:0, x2:0,y2:1, stop:0 #1e3a8a, stop:1 #1e40af);
            color:#ffffff; font-weight:800; padding:8px 10px; border-top:0px; border-bottom:2px solid #0b153f; border-left:1px solid rgba(255,255,255,0.25);
        }
        QTreeWidget#ScheduleTree QHeaderView::section:first { border-top-left-radius:8px; }
        QTreeWidget#ScheduleTree QHeaderView::section:last  { border-top-right-radius:8px; }
        QTreeWidget#ScheduleTree::item { padding:6px 8px; border-bottom:1px solid #e9edf3; }
        QTreeWidget#ScheduleTree::item:selected { background:#e0f2fe; color:#0f172a; }
        """)
        self.tree_sched.setItemDelegate(ScheduleDelegate(self.tree_sched))
        gs.addWidget(self.tree_sched, 0, 0, 1, 1)
        self.tree_sched.itemClicked.connect(self._on_sched_item_clicked)
        self.tree_sched.itemSelectionChanged.connect(self._on_sched_item_clicked_from_selection)
        self.tree_sched.setStyleSheet(self.tree_sched.styleSheet() + "\nQTreeView::item{ min-height: 34px; }")

        # --- OR sticky badge overlay ---
        self._orSticky = QtWidgets.QFrame(self.tree_sched.viewport())
        self._orSticky.setAttribute(QtCore.Qt.WA_TransparentForMouseEvents, True)
        self._orSticky.hide()
        self._orSticky.setStyleSheet("""
QFrame {
    background: qlineargradient(x1:0,y1:0, x2:1,y2:1, stop:0 #4f86ff, stop:1 #2f64e9);
    border-radius: 14px;
    padding: 6px 12px;
    color: #fff;
    border: 0px solid transparent;
}
QLabel { color:#fff; font-weight: 900; }
""")
        sticky_layout = QtWidgets.QHBoxLayout(self._orSticky)
        sticky_layout.setContentsMargins(12, 6, 12, 6)
        self._orStickyLabel = QtWidgets.QLabel("OR")
        sticky_layout.addWidget(self._orStickyLabel)

        self.tree_sched.verticalScrollBar().valueChanged.connect(self._update_or_sticky)
        self.tree_sched.horizontalScrollBar().valueChanged.connect(self._update_or_sticky)
        self.tree_sched.itemExpanded.connect(lambda *_: self._update_or_sticky())
        self.tree_sched.itemCollapsed.connect(lambda *_: self._update_or_sticky())
        self.tree_sched.viewport().installEventFilter(self)

        # Monitor
        self.card_table = ElevatedCard(
            "Result (Monitor) — จากเซิร์ฟเวอร์",
            icon="📺", accent="#8b5cf6", bg="#ffffff", header_bg=_rgba("#8b5cf6", 0.12)
        )
        gt = self.card_table.grid(); gt.setContentsMargins(0,0,0,0)
        self.table = QtWidgets.QTableWidget(0,4)
        self.table.setWordWrap(False); self.table.setItemDelegate(ElideDelegate(QtCore.Qt.ElideRight, self.table))
        self.table.setHorizontalHeaderLabels(["ID","รหัสผู้ป่วย (Patient ID)","สถานะ (Status)","เวลา (Elapsed / เวลาคาดเสร็จ)"])
        th = self.table.horizontalHeader(); th.setStretchLastSection(True); th.setDefaultAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)
        for col,mode in [(0,QtWidgets.QHeaderView.ResizeToContents),(1,QtWidgets.QHeaderView.Stretch),(2,QtWidgets.QHeaderView.ResizeToContents),(3,QtWidgets.QHeaderView.ResizeToContents)]:
            th.setSectionResizeMode(col, mode)
        self.table.verticalHeader().setDefaultSectionSize(34)
        gt.addWidget(self.table,1,0,1,1)
        self.table.itemSelectionChanged.connect(self._on_table_select)

        # Tabs
        self.tabs.addTab(self.card_sched, "Result Schedule Patient")
        self.tabs.addTab(self.card_table, "Status Operation Real Time")
        root.addWidget(self.tabs, 1)

        # Shortcuts
        QShortcut(QKeySequence("Alt+S"), self, self._on_send)
        QShortcut(QKeySequence("Alt+H"), self, self._on_health)
        QShortcut(QKeySequence("Alt+R"), self, lambda: self._refresh(True))
        self._schedule_render()

    # ---------- Helper styles ----------
    def _update_action_styles(self):
        pal = { self.rb_add:"#10b981", self.rb_edit:"#3b82f6", self.rb_del:"#f43f5e" }
        for btn, color in pal.items():
            btn.setStyleSheet(
                f"QPushButton{{padding:6px 12px;border:1px solid "
                f"{color if btn.isChecked() else '#e5e7eb'};"
                f"background:{color if btn.isChecked() else '#f8fafc'};"
                f"color:{'#fff' if btn.isChecked() else '#0f172a'};font-weight:800;}}"
                f"QPushButton:hover{{background:{color if btn.isChecked() else '#eef2f7'};}}"
            )

    # ---------- Settings ----------
    def _load_settings(self):
        s = QSettings("ORNBH", "SurgiBotClient")
        self.ent_host.setText(s.value("host", self.ent_host.text()))
        self.ent_port.setText(s.value("port", self.ent_port.text()))
        self.ent_token.setText(s.value("token", self.ent_token.text()))
        if g := s.value("geometry"):
            try: self.restoreGeometry(g)
            except Exception: pass

    def _save_settings(self):
        s = QSettings("ORNBH", "SurgiBotClient")
        s.setValue("host", self.ent_host.text()); s.setValue("port", self.ent_port.text())
        s.setValue("token", self.ent_token.text()); s.setValue("geometry", self.saveGeometry())

    # ---------- Persist monitor state ----------
    def _save_persisted_monitor_state(self, rows: List[dict]):
        try:
            s = QSettings(PERSIST_ORG, PERSIST_APP)
            s.setValue(KEY_LAST_ROWS, json.dumps(rows, ensure_ascii=False))
            s.setValue(KEY_WAS_IN_MONITOR, json.dumps(sorted(list(self._was_in_monitor))))
            s.setValue(KEY_CURRENT_MONITOR, json.dumps(sorted(list(self._current_monitor_hn))))
        except Exception:
            pass

    def _load_persisted_monitor_state(self):
        try:
            s = QSettings(PERSIST_ORG, PERSIST_APP)
            last_rows_json = s.value(KEY_LAST_ROWS, "")
            was_json = s.value(KEY_WAS_IN_MONITOR, "")
            cur_json = s.value(KEY_CURRENT_MONITOR, "")
            if isinstance(last_rows_json, bytes): last_rows_json = last_rows_json.decode("utf-8", "ignore")
            if isinstance(was_json, bytes): was_json = was_json.decode("utf-8", "ignore")
            if isinstance(cur_json, bytes): cur_json = cur_json.decode("utf-8", "ignore")
            if last_rows_json:
                try:
                    rows = json.loads(last_rows_json)
                    if isinstance(rows, list):
                        self.rows_cache = rows[:]
                except Exception:
                    pass
            if was_json:
                try:
                    arr = json.loads(was_json)
                    if isinstance(arr, list):
                        self._was_in_monitor = set(str(x) for x in arr if isinstance(x, (str,int)))
                except Exception:
                    pass
            if cur_json:
                try:
                    arr = json.loads(cur_json)
                    if isinstance(arr, list):
                        self._current_monitor_hn = set(str(x) for x in arr if isinstance(x, (str,int)))
                except Exception:
                    pass
        finally:
            self.monitor_ready = True
            self._schedule_render()

    def closeEvent(self, e):
        self._save_settings(); self._save_persisted_monitor_state(self.rows_cache)
        if self.ws:
            try: self.ws.close()
            except Exception: pass
        super().closeEvent(e)

    # ---------- Small helpers ----------
    def _toggle_eta_visibility(self):
        is_op = (self.cb_status.currentText() == "กำลังผ่าตัด")
        self.lbl_eta.setVisible(is_op); self.ent_eta.setVisible(is_op); self.ent_eta.setEnabled(is_op)
        if not is_op: self.ent_eta.clear()

    def _set_status_combo(self, status: str | None):
        self._suppress_status_change = True
        try:
            if status:
                idx = self.cb_status.findText(status)
                if idx >= 0:
                    self.cb_status.setCurrentIndex(idx)
                    self._toggle_eta_visibility()
                    return
            self.cb_status.setCurrentIndex(0)
        finally:
            self._suppress_status_change = False
        self._toggle_eta_visibility()

    def _on_status_combo_changed(self, text: str):
        self._toggle_eta_visibility()
        if self._suppress_status_change:
            return
        entry = self._get_active_schedule_entry()
        if entry is None:
            return
        self._apply_status_change(entry, text)

    def _get_active_schedule_entry(self) -> _SchedEntry | None:
        if not self._last_selected_uid:
            return None
        return self.sched.find_by_uid(self._last_selected_uid)

    def _apply_status_change(self, entry: _SchedEntry, new_status: str):
        changed = False
        now_hm = datetime.now().strftime("%H:%M")

        if new_status == STATUS_OP_START:
            if not entry.time_start:
                entry.time_start = now_hm
                changed = True
            if entry.state in ("scheduled", "in_or", "operation_ended", "postop_pending", "") or not entry.state:
                entry.state = "operation_started"
                changed = True

        elif new_status == STATUS_OP_END:
            if not entry.time_end:
                entry.time_end = now_hm
                changed = True
            if entry.state in ("operation_started", "in_or", "scheduled", "") or not entry.state:
                entry.state = "operation_ended"
                changed = True

        elif new_status == STATUS_RETURNING:
            if not entry.time_end:
                QtWidgets.QMessageBox.warning(
                    self,
                    "ยังผ่าตัดไม่จบ",
                    "ยังไม่มีเวลา 'จบผ่าตัด' ระบบฝั่งแม่จะไม่เริ่มนับ 3 นาทีจนกว่าจะมีเวลาจบ",
                )
            entry.state = "returning_to_ward"
            entry.returning_started_at = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
            changed = True

        if entry.status != new_status:
            entry.status = new_status
            changed = True

        if changed:
            entry.version = int(entry.version or 0) + 1
            entry.updated_at = datetime.now().isoformat()
            self.sched.touch_entry(entry)
            self._schedule_render()
            self._flash_row_by_uid(entry.uid())
            self._set_status_combo(entry.status)

    def _reset_form(self):
        self.ent_hn.clear(); self.ent_pid.clear(); self.ent_eta.clear()
        self._set_status_combo(None); self.cb_q.setCurrentIndex(0)
        self.ent_hn.setFocus()
        self.lbl_scan_state.setText("Scanner: Ready"); self.lbl_scan_state.setStyleSheet("color:#16a34a;font-weight:600;")

    def _flash_row_by_uid(self, uid: str):
        if not uid:
            return
        tree = getattr(self, "tree_sched", None)
        if tree is None:
            return
        matches: list[QtWidgets.QTreeWidgetItem] = []
        for i in range(tree.topLevelItemCount()):
            parent = tree.topLevelItem(i)
            if parent is None:
                continue
            for j in range(parent.childCount()):
                child = parent.child(j)
                entry = child.data(0, QtCore.Qt.UserRole)
                if isinstance(entry, _SchedEntry) and entry.uid() == uid:
                    matches.append(child)

        if not matches:
            return

        highlight = QtGui.QBrush(QtGui.QColor("#fef08a"))
        for item in matches:
            for col in range(tree.columnCount()):
                item.setBackground(col, highlight)

        def _clear():
            for item in matches:
                self._style_schedule_item(item, False)

        QtCore.QTimer.singleShot(1200, _clear)

    def _restore_selected_schedule_item(self):
        if not self._last_selected_uid:
            return
        tree = getattr(self, "tree_sched", None)
        if tree is None:
            return
        for i in range(tree.topLevelItemCount()):
            parent = tree.topLevelItem(i)
            if parent is None:
                continue
            for j in range(parent.childCount()):
                child = parent.child(j)
                entry = child.data(0, QtCore.Qt.UserRole)
                if isinstance(entry, _SchedEntry) and entry.uid() == self._last_selected_uid:
                    tree.setCurrentItem(child)
                    tree.scrollToItem(child, QtWidgets.QAbstractItemView.PositionAtCenter)
                    return

    def _on_postop_clicked(self):
        entry = self._get_active_schedule_entry()
        if entry is None:
            QtWidgets.QMessageBox.information(self, "ยังไม่ได้เลือก", "กรุณาเลือกเคสจากตารางก่อน")
            return
        self._open_postop_dialog(entry)

    def _open_postop_dialog(self, entry: _SchedEntry):
        dlg = PostOpDialog(entry, self)
        if dlg.exec() != QtWidgets.QDialog.Accepted:
            return
        values = dlg.values()
        changed = False
        for field in ("assist1", "assist2", "scrub", "circulate"):
            new_val = values.get(field, "")
            if getattr(entry, field, "") != new_val:
                setattr(entry, field, new_val)
                changed = True

        diags = values.get("diags") or []
        if diags and list(entry.diags or []) != diags:
            entry.diags = diags
            changed = True

        ops = values.get("ops") or []
        if ops and list(entry.ops or []) != ops:
            entry.ops = ops
            changed = True

        if not changed:
            return

        entry.version = int(entry.version or 0) + 1
        entry.updated_at = datetime.now().isoformat()
        self.sched.touch_entry(entry)
        self._schedule_render()
        self._flash_row_by_uid(entry.uid())
        self.toast.show_toast("บันทึกหลังผ่าตัดเรียบร้อย")

    def _open_postop_by_uid(self, uid: str):
        if not uid:
            return
        for entry in self.sched.entries:
            if isinstance(entry, _SchedEntry) and entry.uid() == uid:
                self._open_postop_dialog(entry)
                break

    def _make_postop_button(self, uid: str) -> QtWidgets.QPushButton:
        btn = QtWidgets.QPushButton("💾 บันทึกหลังผ่าตัด")
        btn.setCursor(QtCore.Qt.PointingHandCursor)
        btn.setFocusPolicy(QtCore.Qt.NoFocus)
        btn.setStyleSheet(
            """
            QPushButton{
                background:#fb923c; color:#111; border:1px solid #f97316;
                border-radius:12px; padding:6px 10px; font-weight:800;
            }
            QPushButton:hover{ background:#f59e0b; }
            """
        )
        btn.clicked.connect(lambda *_: self._open_postop_by_uid(uid))
        return btn

    def _incomplete(self, entry: _SchedEntry) -> bool:
        if not (entry.time_start and entry.time_end):
            return True
        if not (entry.scrub or entry.circulate or entry.assist1 or entry.assist2):
            return True
        if not (entry.ops or entry.diags):
            return True
        return False

    def _first_visible_item(self) -> QtWidgets.QTreeWidgetItem | None:
        tree = getattr(self, "tree_sched", None)
        if tree is None:
            return None
        viewport = tree.viewport()
        if viewport is None:
            return None
        top_y = 0
        for i in range(tree.topLevelItemCount()):
            parent = tree.topLevelItem(i)
            if parent is None:
                continue
            rect = tree.visualItemRect(parent)
            if rect.bottom() >= top_y:
                if rect.top() <= top_y <= rect.bottom():
                    return parent
                for j in range(parent.childCount()):
                    child = parent.child(j)
                    if child is None:
                        continue
                    child_rect = tree.visualItemRect(child)
                    if child_rect.bottom() >= top_y:
                        return child
                return parent
        return None

    def _update_or_sticky(self):
        tree = getattr(self, "tree_sched", None)
        sticky = getattr(self, "_orSticky", None)
        if tree is None or sticky is None:
            return

        item = self._first_visible_item()
        if item is None:
            sticky.hide()
            return

        parent = item
        while parent.parent() is not None:
            parent = parent.parent()

        or_text = self._or_item_label(parent)
        if not or_text:
            sticky.hide()
            return

        rect = tree.visualItemRect(parent)
        if not rect.isValid():
            sticky.hide()
            return

        self._orStickyLabel.setText(or_text)
        sticky.adjustSize()
        width = max(120, sticky.sizeHint().width())
        height = 32
        x = 8
        y = max(4, rect.top() + 6)
        sticky.setGeometry(x, y, width, height)
        sticky.show()
        sticky.raise_()

    def _set_chip(self, ok: bool):
        base = getattr(self, "_status_pill_base", "background:#ffffff;border:1px solid #e5e7eb;border-radius:10px;padding:4px 10px;font-weight:600;")
        color = "#16a34a" if ok else "#ef4444"
        text = "  • Online  " if ok else "  • Offline  "
        self.status_chip.setText(text)
        self.status_chip.setStyleSheet(f"{base}color:{color};")
        self._conn_state = ok

    def _bump_conn(self, ok: bool):
        if ok:
            self._conn_ok_streak += 1
            self._conn_fail_streak = 0
            if self._conn_state is not True and self._conn_ok_streak >= 2:
                self._set_chip(True)
            self._conn_ok_streak = min(self._conn_ok_streak, 3)
        else:
            self._conn_fail_streak += 1
            self._conn_ok_streak = 0
            if self._conn_state is not False and self._conn_fail_streak >= 2:
                self._set_chip(False)
            self._conn_fail_streak = min(self._conn_fail_streak, 3)

    def _client(self):
        try:
            h = self.ent_host.text().strip() or DEFAULT_HOST
            p = int(self.ent_port.text()) if self.ent_port.text().strip() else DEFAULT_PORT
            t = self.ent_token.text().strip() or DEFAULT_TOKEN
            return SurgiBotClientHTTP(h, p, t)
        except Exception:
            return self.cli

    def _on_health(self):
        def _check():
            return self._client().health()

        task = NetworkTask(_check, self, "_on_health_success", "_on_health_error")
        self._thread_pool.start(task)

    @QtCore.Slot(object)
    def _on_health_success(self, _payload: object):
        self._bump_conn(True)
        QtWidgets.QToolTip.showText(QtGui.QCursor.pos(), "Health OK")

    @QtCore.Slot(object)
    def _on_health_error(self, err: object):
        self._bump_conn(False)
        QtWidgets.QMessageBox.warning(self, "เชื่อมต่อไม่ได้", "กรุณา check IP Address ให้ตรงกับเครื่อง Server ด้วยครับ")
        logger.warning("Health check failed: %s", err)

    # ---------- Data extraction & render helpers ----------
    def _extract_rows(self, payload):
        """Normalize payload from API/websocket into monitor row dicts."""
        src = []
        if isinstance(payload, list):
            src = payload
        elif isinstance(payload, dict):
            for k in ("items", "data", "table", "rows", "list"):
                if k in payload and isinstance(payload[k], list):
                    src = payload[k]
                    break
            else:
                src = next((v for v in payload.values() if isinstance(v, list)), [])

        rows = []
        for i, it in enumerate(src, start=1):
            if not isinstance(it, dict):
                continue

            hn_full = str(it.get("hn_full") or it.get("hn") or "").strip()

            pid = str(
                it.get("patient_id")
                or it.get("pid")
                or it.get("queue_id")
                or ""
            ).strip()
            if not pid:
                or_room = str(it.get("or") or it.get("or_room") or "").strip()
                q = str(it.get("queue") or it.get("q") or "").strip()
                if or_room and q:
                    pid = f"{or_room}-{q}"
                else:
                    pid = f"row-{i}"

            status_raw = str(
                it.get("status")
                or it.get("state")
                or it.get("operation_status")
                or it.get("op_status")
                or ""
            ).strip().lower()
            status_map = {
                "รอผ่าตัด": "รอผ่าตัด", "waiting": "รอผ่าตัด",
                "queued": "รอผ่าตัด", "pending": "รอผ่าตัด",
                "กำลังผ่าตัด": "กำลังผ่าตัด", "operating": "กำลังผ่าตัด",
                "in operation": "กำลังผ่าตัด", "in_operation": "กำลังผ่าตัด",
                "in-surgery": "กำลังผ่าตัด", "surgery": "กำลังผ่าตัด",
                "ongoing": "กำลังผ่าตัด",
                "กำลังพักฟื้น": "กำลังพักฟื้น", "recovery": "กำลังพักฟื้น",
                "pacu": "กำลังพักฟื้น", "post-op": "กำลังพักฟื้น",
                "post_operation": "กำลังพักฟื้น",
                "กำลังส่งกลับตึก": "กำลังส่งกลับตึก", "sending back": "กำลังส่งกลับตึก",
                "transfer": "กำลังส่งกลับตึก", "returning": "กำลังส่งกลับตึก",
                "เลื่อนการผ่าตัด": "เลื่อนการผ่าตัด", "postponed": "เลื่อนการผ่าตัด",
                "deferred": "เลื่อนการผ่าตัด", "canceled": "เลื่อนการผ่าตัด",
                "cancelled": "เลื่อนการผ่าตัด",
            }
            if status_raw in status_map:
                status = status_map[status_raw]
            else:
                try:
                    idx = int(status_raw)
                    map_idx = ["รอผ่าตัด", "กำลังผ่าตัด", "กำลังพักฟื้น", "กำลังส่งกลับตึก", "เลื่อนการผ่าตัด"]
                    status = map_idx[idx] if 0 <= idx < len(map_idx) else "รอผ่าตัด"
                except Exception:
                    status = "รอผ่าตัด"

            ts_val = (
                it.get("timestamp")
                or it.get("ts")
                or it.get("updated_at")
                or it.get("created_at")
                or it.get("time")
                or ""
            )
            ts_iso = ""
            try:
                if isinstance(ts_val, (int, float)):
                    ts_iso = datetime.fromtimestamp(float(ts_val)).isoformat(timespec="seconds")
                elif isinstance(ts_val, str) and ts_val.strip():
                    ts_iso = ts_val
            except Exception:
                ts_iso = ""
            if not _parse_iso(ts_iso):
                ts_iso = datetime.now().isoformat(timespec="seconds")

            eta_raw = it.get("eta_minutes", it.get("eta", it.get("eta_min", None)))
            try:
                eta_minutes = int(eta_raw) if str(eta_raw).strip() != "" else None
            except Exception:
                eta_minutes = None

            rid = it.get("id") or (hn_full if hn_full else pid) or i

            rows.append({
                "id": str(rid),
                "hn_full": hn_full if hn_full else None,
                "patient_id": str(pid),
                "status": status,
                "timestamp": ts_iso,
                "eta_minutes": eta_minutes,
            })
        return rows

    def _render_time_cell(self, row: dict) -> str:
        status = row.get("status", "")
        ts_iso = row.get("timestamp")
        eta_min = row.get("eta_minutes")
        ts = _parse_iso(ts_iso)

        if status == "กำลังผ่าตัด" and ts:
            now = datetime.now()
            elapsed = now - ts
            base = _fmt_td(elapsed)
            if eta_min is not None:
                try:
                    eta_dt = ts + timedelta(minutes=int(eta_min))
                    remain = eta_dt - now
                    flag = "เหลือ" if remain.total_seconds() >= 0 else "เกินเวลา"
                    return f"{base} / ETA {eta_min} นาที ({flag} {_fmt_td(remain)})"
                except Exception:
                    return base
            return base

        if ts and status in ("กำลังพักฟื้น", "พักฟื้นครบแล้ว", "กำลังส่งกลับตึก", "เลื่อนการผ่าตัด"):
            return _fmt_td(datetime.now() - ts)

        return ""

    def _ensure_tray(self):
        if self.tray is None:
            self.tray = QSystemTrayIcon(_load_app_icon(), self)
            self.tray.setToolTip("SurgiBot Client")
            self.tray.show()

    def _rebuild(self, rows):
        # 1) แจ้งเตือนใน tray เมื่อสถานะเปลี่ยน
        new_map = {}
        for r in rows or []:
            pid, st = r.get("patient_id", ""), r.get("status", "")
            if pid:
                new_map[pid] = st
                prev = self._last_states.get(pid)
                if prev is not None and prev != st and self.tray:
                    self.tray.showMessage("SurgiBot", f"{pid} → {st}", QSystemTrayIcon.Information, 3000)
        self._last_states = new_map

        # 2) บันทึก cache และเปิดโหมด monitor
        self.rows_cache = rows if isinstance(rows, list) else []
        self.monitor_ready = True

        # เก็บว่า HN ใดเคยอยู่ใน monitor แล้ว (ใช้กับการขีด + watermark)
        for r in self.rows_cache:
            hn_all = self._extract_hn_from_row(r)
            if hn_all:
                self._was_in_monitor.add(hn_all)

        # ตัดรายการออกตามกติกา auto-purge (ฝั่ง client)
        visible_rows = [r for r in self.rows_cache if not self._should_auto_purge(r)]

        # อัปเดตรายชื่อ HN ที่ "ยังอยู่" ใน monitor ตอนนี้
        current = set()
        for r in visible_rows:
            hn = self._extract_hn_from_row(r)
            if hn:
                current.add(hn)
        self._current_monitor_hn = current

        # 3) วาดตาราง Monitor
        self.table.setRowCount(0)
        for r in visible_rows:
            row = self.table.rowCount()
            self.table.insertRow(row)

            # ID
            self.table.setItem(row, 0, QtWidgets.QTableWidgetItem(str(r.get("id", ""))))

            # Patient ID
            self.table.setItem(row, 1, QtWidgets.QTableWidgetItem(str(r.get("patient_id", ""))))

            # สถานะ + สีพื้นตามสถานะ
            status_item = QtWidgets.QTableWidgetItem(str(r.get("status", "")))
            col = STATUS_COLORS.get(r.get("status", ""))
            if col:
                status_item.setBackground(QtGui.QBrush(QtGui.QColor(col)))
                fg = "#ffffff" if r.get("status") in ("กำลังผ่าตัด", "กำลังส่งกลับตึก",
                                                      "เลื่อนการผ่าตัด") else "#000000"
                status_item.setForeground(QtGui.QBrush(QtGui.QColor(fg)))
            self.table.setItem(row, 2, status_item)

            # เวลาแสดงผล
            self.table.setItem(row, 3, QtWidgets.QTableWidgetItem(self._render_time_cell(r)))

        # 4) วาดตาราง Schedule
        self._schedule_render()

        # 5) persist state
        self._schedule_persist()

    def _update_time_cells(self):
        if not getattr(self, "table", None):
            return
        if not self.rows_cache:
            return
        try:
            by_pid: dict[str, dict] = {}
            by_id: dict[str, dict] = {}
            for raw in self.rows_cache:
                pid = str(raw.get("patient_id", "")).strip()
                rid = str(raw.get("id", "")).strip()
                if pid:
                    by_pid[pid] = raw
                if rid:
                    by_id[rid] = raw

            for row_idx in range(self.table.rowCount()):
                pid_item = self.table.item(row_idx, 1)
                rid_item = self.table.item(row_idx, 0)
                pid = pid_item.text().strip() if pid_item else ""
                rid = rid_item.text().strip() if rid_item else ""
                raw = by_pid.get(pid) or by_id.get(rid)
                if raw is None:
                    continue
                new_text = self._render_time_cell(raw)
                cell = self.table.item(row_idx, 3)
                if cell is None:
                    self.table.setItem(row_idx, 3, QtWidgets.QTableWidgetItem(new_text))
                elif cell.text() != new_text:
                    cell.setText(new_text)
        except Exception:
            pass

    def _schedule_persist(self):
        if self._persist_timer.isActive():
            return
        self._persist_timer.start()

    def _refresh(self, prefer_server=True):
        if not prefer_server:
            self._bump_conn(False)
            self._rebuild(self.model.rows)
            return

        self._refresh_requested = True
        if self._refresh_inflight:
            return
        self._refresh_timer.start(CONFIG.client_debounce_ms)

    def _start_refresh_task(self):
        if self._refresh_inflight or not self._refresh_requested:
            return
        self._refresh_inflight = True
        self._refresh_requested = False
        task = NetworkTask(lambda: self._client().list_items(), self, "_on_refresh_success", "_on_refresh_error")
        self._thread_pool.start(task)

    @QtCore.Slot(object)
    def _on_refresh_success(self, payload: object):
        self._refresh_inflight = False
        rows = self._extract_rows(payload)
        if rows is not None:
            self._rebuild(rows)
            self._bump_conn(True)
        else:
            self._rebuild(self.model.rows)
        if self._refresh_requested:
            self._refresh_timer.start(CONFIG.client_debounce_ms)

    @QtCore.Slot(object)
    def _on_refresh_error(self, err: object):
        self._refresh_inflight = False
        self._bump_conn(False)
        self._rebuild(self.model.rows)
        logger.warning("Refresh failed: %s", err)
        if self._refresh_requested:
            self._refresh_timer.start(CONFIG.client_debounce_ms)

    # ---------- WebSocket ----------
    def _ws_url(self):
        host = self.ent_host.text().strip() or DEFAULT_HOST
        port = int(self.ent_port.text().strip() or DEFAULT_PORT)
        token = self.ent_token.text().strip() or DEFAULT_TOKEN
        return f"ws://{host}:{port}{API_WS}?token={token}"

    def _start_websocket(self):
        if self.ws:
            try:
                self.ws.close()
            except Exception:
                pass
            self.ws = None
        try:
            self.ws = QWebSocket()
            self.ws.errorOccurred.connect(self._ws_error)
            self.ws.connected.connect(self._ws_connected)
            self.ws.disconnected.connect(self._ws_disconnected)
            self.ws.textMessageReceived.connect(self._on_ws_message)
            self.ws.open(QUrl(self._ws_url()))
        except Exception:
            self._ws_disconnected()

    def _ws_connected(self):
        self.ws_connected = True
        self._bump_conn(True)
        if self._pull.isActive():
            self._pull.stop()

    def _ws_disconnected(self):
        self.ws_connected = False
        self._bump_conn(False)
        if not self._pull.isActive():
            self._pull.start(2000)

    def _ws_error(self, err):
        self._bump_conn(False)
        self._ws_disconnected()

    def _on_ws_message(self, msg: str):
        try:
            payload = json.loads(msg)
            rows = self._extract_rows(payload)
            if rows is not None:
                self._rebuild(rows)
        except Exception:
            pass

    def _on_reconnect_clicked(self):
        self.cli = self._client()
        self._save_settings()
        self._on_health()
        self._refresh(True)
        self._start_websocket()

    # ---------- Barcode ----------
    def _finalize_scan_if_any(self):
        if not self._scan_buf:
            return
        digits = "".join(ch for ch in self._scan_buf if ch.isdigit())
        self._scan_buf = ""
        if not digits:
            return
        if len(digits) >= 9:
            hn9 = digits[-9:]
            self.ent_hn.setText(hn9)
            QtWidgets.QApplication.beep()
            self.lbl_scan_state.setText("Scanner: HN captured")
            self.lbl_scan_state.setStyleSheet("color:#2563eb;font-weight:600;")
            self.cb_status.setFocus()
        else:
            self.lbl_scan_state.setText("Scanner: Waiting")
            self.lbl_scan_state.setStyleSheet("color:#16a34a;font-weight:600;")

    def eventFilter(self, obj, event):
        tree = getattr(self, "tree_sched", None)
        viewport = tree.viewport() if tree else None
        if viewport is not None and obj is viewport:
            if event.type() in (
                QtCore.QEvent.Resize,
                QtCore.QEvent.Paint,
                QtCore.QEvent.Show,
                QtCore.QEvent.UpdateRequest,
            ):
                QtCore.QTimer.singleShot(0, self._update_or_sticky)

        if event.type() == QtCore.QEvent.KeyPress and self.scan_enabled:
            key = event.key()
            text = event.text() or ""
            if key in (QtCore.Qt.Key_Return, QtCore.Qt.Key_Enter):
                self._scan_timer.stop()
                self._finalize_scan_if_any()
                return True
            if text and text.isprintable():
                if not self._scan_timer.isActive():
                    self._scan_buf = ""
                self._scan_buf += text
                self._scan_timer.start(self._scan_timeout_ms)
                return False
        return super().eventFilter(obj, event)

    # ---------- Table selection ----------
    def _on_table_select(self):
        try:
            row = self.table.currentRow()
            if row < 0:
                return
            pid_item = self.table.item(row, 1)
            st_item = self.table.item(row, 2)
            id_item = self.table.item(row, 0)

            pid = (pid_item.text().strip() if pid_item else "")
            st = (st_item.text().strip() if st_item else "")
            hid = (id_item.text().strip() if id_item else "")

            if pid:
                self.ent_pid.setText(pid)
            if st:
                self._set_status_combo(st)
            else:
                self._set_status_combo(None)
            if hid.isdigit() and len(hid) == 9:
                self.ent_hn.setText(hid)

            self.rb_edit.setChecked(True)
            self._update_action_styles()
        except Exception:
            pass

    # ---------- Actions ----------
    def _on_send(self):
        action = "add" if self.rb_add.isChecked() else ("edit" if self.rb_edit.isChecked() else "delete")
        pid = self.ent_pid.text().strip() or None
        or_room = None if pid else self.cb_or.currentText()
        q = None if pid else self.cb_q.currentText()
        status = self.cb_status.currentText() if action in ("add", "edit") else None

        hn = self.ent_hn.text().strip()
        if action in ("add", "edit") and (not hn or len(hn) != 9 or not hn.isdigit()):
            QtWidgets.QMessageBox.warning(self, "ข้อมูลไม่ครบ", "กรุณากรอก HN 9 หลักให้ถูกต้อง")
            return

        eta_minutes = None
        if self.cb_status.currentText() == "กำลังผ่าตัด":
            eta_val = self.ent_eta.text().strip()
            eta_minutes = int(eta_val) if eta_val.isdigit() else None

        eff_pid = pid or f"{or_room}-{q}"
        ts_iso = datetime.now().isoformat()
        context = {
            "action": action,
            "eff_pid": eff_pid,
            "status": status,
            "ts_iso": ts_iso,
            "eta_minutes": eta_minutes,
            "hn": hn if action != "delete" else None,
            "patient_id": pid,
            "or_room": or_room,
            "queue": q,
        }

        def _send():
            try:
                resp = self._client().send_update(
                    action=action,
                    or_room=or_room,
                    queue=q,
                    status=status,
                    patient_id=pid,
                    eta_minutes=eta_minutes,
                    hn=hn if action != "delete" else None,
                )
            except requests.exceptions.RequestException as exc:
                raise SubmitError(exc, context)
            return {"response": resp, "context": context}

        task = NetworkTask(_send, self, "_on_submit_success", "_on_submit_error")
        self._thread_pool.start(task)

    @QtCore.Slot(object)
    def _on_submit_success(self, payload: object):
        data = payload if isinstance(payload, dict) else {}
        ctx = data.get("context", {}) if isinstance(data, dict) else {}
        action = ctx.get("action")
        eff_pid = ctx.get("eff_pid")
        status = ctx.get("status")
        ts_iso = ctx.get("ts_iso") or datetime.now().isoformat()
        eta_minutes = ctx.get("eta_minutes")
        hn = ctx.get("hn")
        if action == "delete":
            if eff_pid:
                self.model.delete(eff_pid)
        else:
            self.model.add_or_edit(eff_pid, status or "", ts_iso, eta_minutes, hn=hn)
        self._bump_conn(True)
        self._refresh(True)
        self._reset_form()

    @QtCore.Slot(object)
    def _on_submit_error(self, err: object):
        ctx = getattr(err, "context", None)
        if not isinstance(ctx, dict):
            ctx = {}
        action = ctx.get("action")
        eff_pid = ctx.get("eff_pid")
        status = ctx.get("status")
        ts_iso = ctx.get("ts_iso") or datetime.now().isoformat()
        eta_minutes = ctx.get("eta_minutes")
        hn = ctx.get("hn")
        if action == "delete":
            if eff_pid:
                self.model.delete(eff_pid)
        else:
            self.model.add_or_edit(eff_pid, status or "", ts_iso, eta_minutes, hn=hn)
        self._bump_conn(False)
        self._refresh(False)
        self._reset_form()
        logger.warning("Send update failed: %s", getattr(err, "original", err))

    # ---------- Schedule ----------

    def _schedule_render(self):
        if self._render_pending:
            return
        self._render_pending = True
        self._render_timer.start()

    def _flush_scheduled_render(self):
        self._render_pending = False
        try:
            self._render_schedule_tree()
        finally:
            if self.monitor_ready:
                self._update_schedule_completion_markers()

    def _render_schedule_tree(self):
        """วาด Result Schedule ให้ตรงกับ Registry + เคารพสถานะพับ/ขยายของผู้ใช้"""
        tree = getattr(self, "tree_sched", None)
        if tree is None:
            return

        self._capture_or_expand_state()

        hbar = tree.horizontalScrollBar()
        vbar = tree.verticalScrollBar()
        old_h = hbar.value() if hbar is not None else 0
        old_v = vbar.value() if vbar is not None else 0

        tree.setUpdatesEnabled(False)
        try:
            self._clear_sched_pulser()
            tree.clear()

            today = datetime.now().date()

            def _is_today(entry: _SchedEntry) -> bool:
                if entry.date_obj:
                    return entry.date_obj == today
                if entry.date:
                    return entry.date.strip().startswith(today.isoformat())
                return True

            groups: dict[str, list[_SchedEntry]] = {}

            def should_show(e: _SchedEntry) -> bool:
                if not _is_today(e):
                    return False
                if not self.ws_connected:
                    return True
                return True

            for e in self.sched.entries:
                if should_show(e):
                    groups.setdefault(e.or_room or "-", []).append(e)

            order = self.sched.or_rooms or []

            def room_key(x: str):
                return (order.index(x) if x in order else 999, x)

            def row_sort_key(e: _SchedEntry):
                q = int(e.queue or 0)
                if q > 0:
                    return (0, q, "")
                return (1, 0, e.time or "99:99")

            for orr in sorted(groups.keys(), key=room_key):
                if not groups[orr]:
                    continue

                parent = QtWidgets.QTreeWidgetItem([""] * tree.columnCount())
                header_title = f"{orr}  ห้องผ่าตัด"
                parent.setText(0, header_title)
                parent.setData(0, QtCore.Qt.UserRole + 200, orr)
                parent.setData(0, QtCore.Qt.UserRole + 201, header_title)
                parent.setFirstColumnSpanned(True)
                tree.addTopLevelItem(parent)

                self._style_or_group_header(parent, "#eef2ff")
                parent.setFlags((parent.flags() | QtCore.Qt.ItemIsEnabled) & ~QtCore.Qt.ItemIsSelectable)

                accent = OR_HEADER_COLORS.get(orr, "#64748b")
                tree.setItemWidget(parent, 0, self._or_card_widget(orr, accent))

                self._apply_or_expand_state(parent)

                for e in sorted(groups[orr], key=row_sort_key):
                    display_or = e.or_room or orr or "-"
                    display_time = e.time or "-"
                    if display_or == "-" and display_time == "-":
                        or_cell = "-"
                    elif display_or == "-":
                        or_cell = display_time
                    elif display_time == "-":
                        or_cell = display_or
                    else:
                        or_cell = f"{display_or} / {display_time}"
                    row = QtWidgets.QTreeWidgetItem([
                        "",
                        _period_label(e.period),
                        or_cell,
                        e.hn,
                        (e.name or "-"),
                        (str(e.age) if e.age not in (None, "") else "-"),
                        (", ".join(e.diags) if getattr(e, "diags", None) else "-"),
                        (", ".join(e.ops) if getattr(e, "ops", None) else "-"),
                        (e.doctor or "-"),
                        (e.ward or "-"),
                        (e.case_size or "-"),
                        (e.dept or "-"),
                        (e.assist1 or "-"),
                        (e.assist2 or "-"),
                        (e.scrub or "-"),
                        (e.circulate or "-"),
                        (e.status or "-"),
                        (e.time_start or "-"),
                        (e.time_end or "-"),
                        (str(e.queue) if str(getattr(e, "queue", "0")).isdigit() and int(getattr(e, "queue", "0")) > 0 else "ตามเวลา"),
                        (e.urgency or "Elective"),
                    ])
                    row.setData(0, QtCore.Qt.UserRole, e)
                    parent.addChild(row)

                    status_text = (e.status or "").strip()
                    if status_text:
                        col_hex = STATUS_COLORS.get(status_text)
                        if col_hex:
                            status_col = 16
                            row.setBackground(status_col, QtGui.QBrush(QtGui.QColor(col_hex)))
                            fg = "#ffffff" if status_text in {"กำลังผ่าตัด", "กำลังส่งกลับตึก", "เลื่อนการผ่าตัด"} else "#000000"
                            row.setForeground(status_col, QtGui.QBrush(QtGui.QColor(fg)))

                    if self._incomplete(e):
                        tree.setItemWidget(row, 0, self._make_postop_button(e.uid()))
        finally:
            tree.setUpdatesEnabled(True)

            def _restore_scroll():
                if hbar is not None:
                    hbar.setValue(min(old_h, hbar.maximum()))
                if vbar is not None:
                    vbar.setValue(min(old_v, vbar.maximum()))

            QtCore.QTimer.singleShot(0, _restore_scroll)

        QtCore.QTimer.singleShot(0, self._autofit_schedule_columns)
        QtCore.QTimer.singleShot(0, self._update_or_sticky)
        QtCore.QTimer.singleShot(0, self._restore_selected_schedule_item)
    def _update_schedule_completion_markers(self):
        return

    def _style_schedule_item(self, item: QtWidgets.QTreeWidgetItem, completed: bool):
        cols = self.tree_sched.columnCount()
        for c in range(cols):
            item.setForeground(c, QtGui.QBrush())
            item.setBackground(c, QtGui.QBrush())
            f = self.tree_sched.font()
            f.setStrikeOut(False)
            item.setFont(c, f)
        item.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)

    def _check_schedule_seq(self):
        if self.sched.refresh_if_changed():
            self._schedule_render()
        # ไม่บังคับ expandAll เพื่อคงสถานะพับ/ขยายของผู้ใช้
        QtCore.QTimer.singleShot(0, self._autofit_schedule_columns)


# ---------- main (module level) ----------
def run_gui_pyside6(host, port, token):
    app = QtWidgets.QApplication([])
    app.setWindowIcon(_load_app_icon())
    ui = Main(host, port, token)
    ui.setWindowIcon(_load_app_icon())
    ui.showMaximized()
    app.exec()


def build_parser():
    p = argparse.ArgumentParser(
        description="SurgiBot Client (+ETA + HN + Barcode Scan + WebSocket + Tray + Private Schedule Viewer)",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter
    )
    p.add_argument("--host", default=DEFAULT_HOST)
    p.add_argument("--port", default=DEFAULT_PORT, type=int)
    p.add_argument("--token", default=DEFAULT_TOKEN)
    p.add_argument("--gui", action="store_true")
    p.add_argument("--ui", choices=["ttk", "pyside6"], default=os.getenv("SURGIBOT_CLIENT_UI", "pyside6"))

    sub = p.add_subparsers(dest="cmd", required=False)
    sub.add_parser("health")
    sub.add_parser("list")

    addp = sub.add_parser("add")
    addp.add_argument("--hn", required=True)
    addp.add_argument("--or", dest="or_room", choices=OR_CHOICES, required=True)
    addp.add_argument("--queue", choices=QUEUE_CHOICES, required=True)
    addp.add_argument("--status", choices=STATUS_CHOICES, required=True)
    addp.add_argument("--eta", dest="eta", type=int)

    edp = sub.add_parser("edit")
    g = edp.add_mutually_exclusive_group(required=True)
    g.add_argument("--patient-id", dest="patient_id")
    g.add_argument("--or", dest="or_room", choices=OR_CHOICES)
    edp.add_argument("--queue", choices=QUEUE_CHOICES)
    edp.add_argument("--status", choices=STATUS_CHOICES)
    edp.add_argument("--eta", dest="eta", type=int)
    edp.add_argument("--hn")

    delp = sub.add_parser("delete")
    g2 = delp.add_mutually_exclusive_group(required=True)
    g2.add_argument("--patient-id", dest="patient_id")
    g2.add_argument("--or", dest="or_room", choices=OR_CHOICES)
    delp.add_argument("--queue", choices=QUEUE_CHOICES)
    return p


def main():
    parser = build_parser()
    args = parser.parse_args()
    if args.cmd is None:
        run_gui_pyside6(args.host, args.port, args.token)
        return

    cli = SurgiBotClientHTTP(args.host, args.port, args.token)
    try:
        if args.cmd == "health":
            print(json.dumps(cli.health(), ensure_ascii=False, indent=2))
        elif args.cmd == "list":
            print(json.dumps(cli.list_items(), ensure_ascii=False, indent=2))
        elif args.cmd == "add":
            print(json.dumps(cli.send_update("add",
                                             or_room=args.or_room, queue=args.queue,
                                             status=args.status, eta_minutes=args.eta, hn=args.hn),
                             ensure_ascii=False, indent=2))
        elif args.cmd == "edit":
            if args.patient_id:
                print(json.dumps(cli.send_update("edit",
                                                 patient_id=args.patient_id,
                                                 status=args.status, eta_minutes=args.eta, hn=args.hn),
                                 ensure_ascii=False, indent=2))
            else:
                if not args.or_room or not args.queue:
                    raise SystemExit("--or และ --queue จำเป็นเมื่อไม่ได้ส่ง --patient-id")
                print(json.dumps(cli.send_update("edit",
                                                 or_room=args.or_room, queue=args.queue,
                                                 status=args.status, eta_minutes=args.eta, hn=args.hn),
                                 ensure_ascii=False, indent=2))
        elif args.cmd == "delete":
            if args.patient_id:
                print(json.dumps(cli.send_update("delete", patient_id=args.patient_id),
                                 ensure_ascii=False, indent=2))
            else:
                if not args.or_room or not args.queue:
                    raise SystemExit("--or และ --queue จำเป็นเมื่อไม่ได้ส่ง --patient-id")
                print(json.dumps(cli.send_update("delete",
                                                 or_room=args.or_room, queue=args.queue),
                                 ensure_ascii=False, indent=2))
        else:
            print("Unknown command", file=sys.stderr)
    except requests.HTTPError as he:
        print(f"[HTTP ERROR] {he}", file=sys.stderr)
        sys.exit(1)
    except SystemExit as se:
        print(str(se), file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"[ERROR] {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()