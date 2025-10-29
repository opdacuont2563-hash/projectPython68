"""Central logging configuration for SurgiBot."""
from __future__ import annotations

import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Optional

from .config import CONFIG

_LOGGERS: dict[str, logging.Logger] = {}


def configure_logging(force: bool = False) -> None:
    if _LOGGERS and not force:
        return

    log_dir: Path = CONFIG.log_dir
    log_dir.mkdir(parents=True, exist_ok=True)
    log_path = log_dir / "surgibot.log"

    fmt = "%(asctime)s %(levelname)s %(name)s %(message)s"
    formatter = logging.Formatter(fmt)

    root_logger = logging.getLogger()
    root_logger.setLevel(logging.INFO)

    rotating = RotatingFileHandler(log_path, maxBytes=2_000_000, backupCount=5, encoding="utf-8")
    rotating.setFormatter(formatter)

    console = logging.StreamHandler()
    console.setFormatter(formatter)

    root_logger.handlers.clear()
    root_logger.addHandler(rotating)
    root_logger.addHandler(console)


def get_logger(name: Optional[str] = None) -> logging.Logger:
    configure_logging()
    logger_name = name or "surgibot"
    if logger_name not in _LOGGERS:
        _LOGGERS[logger_name] = logging.getLogger(logger_name)
    return _LOGGERS[logger_name]


__all__ = ["get_logger", "configure_logging"]
