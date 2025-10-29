"""Network/background worker utilities for SurgiBot."""
from __future__ import annotations

import queue
import threading
from typing import Any, Callable

import requests
from requests.adapters import HTTPAdapter, Retry

from ..config import CONFIG
from ..logging_setup import get_logger

logger = get_logger(__name__)


class SessionManager:
    """Maintain a shared requests.Session with retries."""

    def __init__(self) -> None:
        self._lock = threading.Lock()
        self._session: Optional[requests.Session] = None

    def get(self) -> requests.Session:
        with self._lock:
            if self._session is None:
                self._session = requests.Session()
                retries = Retry(
                    total=CONFIG.request_retries,
                    connect=CONFIG.request_retries,
                    read=CONFIG.request_retries,
                    backoff_factor=CONFIG.request_backoff,
                    status_forcelist=(429, 500, 502, 503, 504),
                    allowed_methods=frozenset(["GET", "POST"]),
                )
                adapter = HTTPAdapter(max_retries=retries)
                self._session.mount("http://", adapter)
                self._session.mount("https://", adapter)
            return self._session


SESSION_MANAGER = SessionManager()


class RequestExecutor:
    """Execute blocking request callables in a background thread."""

    def __init__(self) -> None:
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._queue: "queue.Queue[tuple[Callable[[], Any], Callable[[Any], None], Callable[[BaseException], None]]]]" = queue.Queue()
        self._thread.start()

    def submit(
        self,
        fn: Callable[[], Any],
        on_success: Callable[[Any], None],
        on_error: Callable[[BaseException], None],
    ) -> None:
        self._queue.put((fn, on_success, on_error))

    def _run(self) -> None:
        while True:
            try:
                fn, on_success, on_error = self._queue.get()
                try:
                    result = fn()
                except BaseException as exc:  # pragma: no cover - worker thread
                    logger.error("Request worker exception: %s", exc)
                    on_error(exc)
                else:
                    on_success(result)
            except Exception as exc:  # pragma: no cover - worker thread
                logger.error("Unexpected worker loop error: %s", exc)


try:
    from PySide6 import QtCore
except Exception:  # pragma: no cover - server usage without PySide6
    QtCore = None  # type: ignore


if QtCore:

    class NetworkTask(QtCore.QRunnable):
        def __init__(self, fn: Callable[[], Any], receiver: "QtCore.QObject", success_slot: str, error_slot: str) -> None:
            super().__init__()
            self.fn = fn
            self.receiver = receiver
            self.success_slot = success_slot
            self.error_slot = error_slot

        def run(self) -> None:  # pragma: no cover - executed in thread pool
            try:
                result = self.fn()
            except BaseException as exc:
                QtCore.QMetaObject.invokeMethod(
                    self.receiver,
                    self.error_slot,
                    QtCore.Qt.QueuedConnection,
                    QtCore.Q_ARG(object, exc),
                )
            else:
                QtCore.QMetaObject.invokeMethod(
                    self.receiver,
                    self.success_slot,
                    QtCore.Qt.QueuedConnection,
                    QtCore.Q_ARG(object, result),
                )


__all__ = ["SESSION_MANAGER", "NetworkTask", "SessionManager"]
