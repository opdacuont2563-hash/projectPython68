"""Audio worker for SurgiBot announcements."""
from __future__ import annotations

import hashlib
import queue
import threading
import time
from pathlib import Path
from typing import Optional, Tuple

import pygame
from gtts import gTTS

from ..config import CONFIG
from ..logging_setup import get_logger

logger = get_logger(__name__)


class AudioWorker:
    """Background worker that serializes audio playback and caching."""

    def __init__(self, cache_dir: Optional[Path] = None) -> None:
        self.cache_dir = cache_dir or CONFIG.audio_cache_dir
        self.cache_dir.mkdir(parents=True, exist_ok=True)
        self._queue: "queue.Queue[Tuple[str, str, int]]" = queue.Queue()
        self._lock = threading.Lock()
        self._stop = threading.Event()
        self._last_text: Optional[Tuple[str, float]] = None
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()
        try:
            pygame.mixer.init()
        except Exception as exc:  # pragma: no cover - audio backend optional
            logger.warning("pygame mixer init failed: %s", exc)

    def enqueue_bilingual(self, th_text: str, en_text: str, pause_ms: int) -> None:
        if not th_text and not en_text:
            return
        now = time.monotonic()
        with self._lock:
            if self._last_text and self._last_text[0] == th_text + en_text:
                if now - self._last_text[1] < CONFIG.announcement_ttl_seconds:
                    logger.debug("Skipping duplicate announcement within TTL")
                    return
            self._last_text = (th_text + en_text, now)
        self._queue.put((th_text, en_text, pause_ms))

    def stop(self) -> None:
        self._stop.set()
        self._queue.put(("", "", 0))
        if self._thread.is_alive():
            self._thread.join(timeout=1.5)
        try:
            pygame.mixer.quit()
        except Exception:
            pass

    def _run(self) -> None:
        while not self._stop.is_set():
            try:
                th_text, en_text, pause_ms = self._queue.get(timeout=0.5)
            except queue.Empty:
                continue
            if self._stop.is_set():
                break
            if not (th_text or en_text):
                continue
            try:
                self._play_segment(th_text, "th")
                if pause_ms:
                    time.sleep(max(0, pause_ms) / 1000.0)
                self._play_segment(en_text, "en")
            except Exception as exc:
                logger.error("Audio worker error: %s", exc)

    def _play_segment(self, text: str, lang: str) -> None:
        if not text:
            return
        filename = self._cache_path(text, lang)
        if not filename.exists():
            gTTS(text=text, lang=lang).save(str(filename))
        try:
            pygame.mixer.music.load(str(filename))
            pygame.mixer.music.play()
            start = time.time()
            while pygame.mixer.music.get_busy():
                if time.time() - start > 120:
                    logger.warning("Audio playback exceeded maximum duration; forcing stop")
                    break
                time.sleep(0.1)
            pygame.mixer.music.stop()
        except Exception as exc:  # pragma: no cover - optional audio backend
            logger.error("Playback error: %s", exc)

    def _cache_path(self, text: str, lang: str) -> Path:
        digest = hashlib.sha256(f"{lang}:{text}".encode("utf-8")).hexdigest()
        return self.cache_dir / f"{digest}.mp3"


__all__ = ["AudioWorker"]
