"""Configuration helpers for SurgiBot components."""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Optional
import json
import os

try:
    from dotenv import load_dotenv  # type: ignore
except ImportError:  # pragma: no cover - optional dependency
    load_dotenv = None  # type: ignore

_ENV_LOADED = False


def _load_env_files() -> None:
    global _ENV_LOADED
    if _ENV_LOADED:
        return
    env_paths = [Path.cwd() / ".env", Path(__file__).resolve().parents[2] / ".env"]
    for path in env_paths:
        if path.exists() and load_dotenv:
            load_dotenv(path)
    _ENV_LOADED = True


@dataclass(frozen=True)
class SurgiBotConfig:
    api_host: str = "0.0.0.0"
    api_port: int = 8088
    secret: str = "uTCoBelMyNfSSNmUulT_Kz6zrrCVkvD578MxEuLKZoaaXX0pVlpAD8toYHBxsFxI"
    client_host: str = "127.0.0.1"
    client_port: int = 8088
    client_timeout: float = 6.0
    client_base_url: str = "http://127.0.0.1:8088"
    client_refresh_interval_ms: int = 2000
    client_debounce_ms: int = 180
    client_auto_purge_minutes: int = 3
    announce_interval_minutes: int = 20
    auto_delete_minutes: int = 3
    google_sheet_id: str = "1dr6pCw8dEnCh_UYJXJzAFsthdZ8IsRBF3VlNw1AlvjI"
    gcp_credentials_json: Optional[str] = None
    gcp_credentials_file: Optional[str] = None
    embedded_credentials_json: Optional[str] = None
    announcement_ttl_seconds: int = 5
    audio_cache_dir: Path = Path("assets/cache")
    data_dir: Path = Path("data")
    request_timeout: float = 6.0
    request_retries: int = 3
    request_backoff: float = 0.35
    log_dir: Path = Path("logs")

    @property
    def api_base_url(self) -> str:
        return f"http://{self.api_host}:{self.api_port}"

    @property
    def client_secret(self) -> str:
        return self.secret

    def google_credentials_payload(self) -> Optional[Dict[str, Any]]:
        raw = self.gcp_credentials_json or self.embedded_credentials_json
        if raw:
            try:
                return json.loads(raw)
            except json.JSONDecodeError:
                return None
        if self.gcp_credentials_file:
            path = Path(self.gcp_credentials_file)
            if path.exists():
                try:
                    return json.loads(path.read_text(encoding="utf-8"))
                except (OSError, json.JSONDecodeError):
                    return None
        return None


def load_config() -> SurgiBotConfig:
    _load_env_files()
    env = os.environ.get
    api_host = env("SURGIBOT_API_HOST", SurgiBotConfig.api_host)
    api_port = int(env("SURGIBOT_API_PORT", str(SurgiBotConfig.api_port)))
    secret = env("SURGIBOT_SECRET", SurgiBotConfig.secret)

    client_host = env("SURGIBOT_CLIENT_HOST", SurgiBotConfig.client_host)
    client_port = int(env("SURGIBOT_CLIENT_PORT", str(SurgiBotConfig.client_port)))
    client_timeout = float(env("SURGIBOT_CLIENT_TIMEOUT", str(SurgiBotConfig.client_timeout)))
    client_base = env("SURGIBOT_CLIENT_BASE", f"http://{client_host}:{client_port}")
    refresh_ms = int(env("SURGIBOT_CLIENT_REFRESH_MS", str(SurgiBotConfig.client_refresh_interval_ms)))
    debounce_ms = int(env("SURGIBOT_CLIENT_DEBOUNCE_MS", str(SurgiBotConfig.client_debounce_ms)))
    auto_purge_min = int(env("SURGIBOT_CLIENT_PURGE_MINUTES", str(SurgiBotConfig.client_auto_purge_minutes)))

    announce_min = int(env("SURGIBOT_ANNOUNCE_MINUTES", str(SurgiBotConfig.announce_interval_minutes)))
    auto_delete_min = int(env("SURGIBOT_AUTO_DELETE_MIN", str(SurgiBotConfig.auto_delete_minutes)))
    google_sheet_id = env("SURGIBOT_SPREADSHEET_ID", SurgiBotConfig.google_sheet_id)
    gcp_json = env("SURGIBOT_GCP_CREDENTIALS_JSON")
    gcp_file = env("SURGIBOT_GCP_CREDENTIALS_FILE")
    embedded_json = env("SURGIBOT_EMBEDDED_CREDENTIALS_JSON")
    announce_ttl = int(env("SURGIBOT_ANNOUNCE_TTL", str(SurgiBotConfig.announcement_ttl_seconds)))
    audio_cache_dir = Path(env("SURGIBOT_AUDIO_CACHE", str(SurgiBotConfig.audio_cache_dir)))
    data_dir = Path(env("SURGIBOT_DATA_DIR", str(SurgiBotConfig.data_dir)))
    request_timeout = float(env("SURGIBOT_REQUEST_TIMEOUT", str(SurgiBotConfig.request_timeout)))
    request_retries = int(env("SURGIBOT_REQUEST_RETRIES", str(SurgiBotConfig.request_retries)))
    request_backoff = float(env("SURGIBOT_REQUEST_BACKOFF", str(SurgiBotConfig.request_backoff)))
    log_dir = Path(env("SURGIBOT_LOG_DIR", str(SurgiBotConfig.log_dir)))

    cfg = SurgiBotConfig(
        api_host=api_host,
        api_port=api_port,
        secret=secret,
        client_host=client_host,
        client_port=client_port,
        client_timeout=client_timeout,
        client_base_url=client_base,
        client_refresh_interval_ms=refresh_ms,
        client_debounce_ms=debounce_ms,
        client_auto_purge_minutes=auto_purge_min,
        announce_interval_minutes=announce_min,
        auto_delete_minutes=auto_delete_min,
        google_sheet_id=google_sheet_id,
        gcp_credentials_json=gcp_json,
        gcp_credentials_file=gcp_file,
        embedded_credentials_json=embedded_json,
        announcement_ttl_seconds=announce_ttl,
        audio_cache_dir=audio_cache_dir,
        data_dir=data_dir,
        request_timeout=request_timeout,
        request_retries=request_retries,
        request_backoff=request_backoff,
        log_dir=log_dir,
    )
    cfg.audio_cache_dir.mkdir(parents=True, exist_ok=True)
    cfg.log_dir.mkdir(parents=True, exist_ok=True)
    cfg.data_dir.mkdir(parents=True, exist_ok=True)
    return cfg


CONFIG = load_config()

__all__ = ["CONFIG", "SurgiBotConfig", "load_config"]
