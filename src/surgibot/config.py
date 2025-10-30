"""Configuration helpers for SurgiBot components."""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Optional
from urllib.parse import urlparse
import json
import os

try:
    from dotenv import load_dotenv  # type: ignore
except ImportError:  # pragma: no cover - optional dependency
    load_dotenv = None  # type: ignore

_ENV_LOADED = False

_TRUE_VALUES = {"1", "true", "t", "yes", "y", "on"}
_FALSE_VALUES = {"0", "false", "f", "no", "n", "off"}
_LOCAL_HOST_SENTINELS = {"", "0.0.0.0", "127.0.0.1", "localhost"}


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
    secret: str = "8HDYAANLgTyjbBK4JPGx1ooZbVC86_OMJ9uEXBm3EZTidUVyzhGiReaksGA0ites"
    secret_from_env: bool = False
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
    runner_enabled: bool = False
    runner_port: int = 8777
    runner_base_url: str = ""

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


def _normalize_runner_base(
    base_url: Optional[str],
    *,
    runner_port: int,
    fallback_host: str,
    fallback_scheme: str,
) -> str:
    if base_url:
        candidate = base_url.strip()
        if candidate:
            if "://" not in candidate:
                candidate = f"{fallback_scheme}://{candidate}"
            parsed = urlparse(candidate)
            scheme = parsed.scheme or fallback_scheme
            host = parsed.hostname or ""
            if not host:
                host = candidate.split("://", 1)[-1].split("/", 1)[0]
            if host in _LOCAL_HOST_SENTINELS:
                host = "127.0.0.1"
            port = parsed.port or runner_port
            return f"{scheme}://{host}:{port}".rstrip("/")

    host = fallback_host or "127.0.0.1"
    if host in _LOCAL_HOST_SENTINELS:
        host = "127.0.0.1"
    return f"{fallback_scheme}://{host}:{runner_port}".rstrip("/")


def _parse_bool(value: Optional[str], default: bool = False) -> bool:
    if value is None:
        return default
    text = value.strip().lower()
    if not text:
        return default
    if text in _TRUE_VALUES:
        return True
    if text in _FALSE_VALUES:
        return False
    return default


def _clean_host(value: str) -> str:
    text = value.strip()
    if text.startswith("http://") or text.startswith("https://"):
        parsed = urlparse(text)
        host = parsed.hostname or ""
        return host or text
    return text


def load_config() -> SurgiBotConfig:
    _load_env_files()
    env = os.environ.get
    raw_api_host = env("SURGIBOT_API_HOST")
    api_host = _clean_host(raw_api_host or SurgiBotConfig.api_host)
    try:
        api_port = int((env("SURGIBOT_API_PORT") or str(SurgiBotConfig.api_port)).strip())
    except (TypeError, ValueError):
        api_port = SurgiBotConfig.api_port
    raw_secret = env("SURGIBOT_SECRET")
    secret_from_env = bool(raw_secret and raw_secret.strip())
    secret = (raw_secret or SurgiBotConfig.secret).strip() or SurgiBotConfig.secret

    raw_client_host = env("SURGIBOT_CLIENT_HOST")
    if raw_client_host is not None and raw_client_host.strip():
        client_host = _clean_host(raw_client_host)
    else:
        fallback_host = _clean_host(api_host)
        client_host = fallback_host if fallback_host not in _LOCAL_HOST_SENTINELS else SurgiBotConfig.client_host

    raw_client_port = env("SURGIBOT_CLIENT_PORT")
    if raw_client_port is not None and raw_client_port.strip():
        try:
            client_port = int(raw_client_port.strip())
        except ValueError:
            client_port = SurgiBotConfig.client_port
    else:
        client_port = api_port if api_port else SurgiBotConfig.client_port
    client_timeout = float(env("SURGIBOT_CLIENT_TIMEOUT", str(SurgiBotConfig.client_timeout)))
    raw_client_base = env("SURGIBOT_CLIENT_BASE")
    client_base = raw_client_base.strip() if raw_client_base and raw_client_base.strip() else f"http://{client_host}:{client_port}"
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

    try:
        runner_port = int(env("SURGIBOT_RUNNER_PORT", str(SurgiBotConfig.runner_port)))
    except ValueError:
        runner_port = SurgiBotConfig.runner_port
    parsed_client = urlparse(client_base if "://" in client_base else f"http://{client_base}")
    fallback_scheme = parsed_client.scheme or "http"
    fallback_host = parsed_client.hostname or client_host
    runner_base_env = env("SURGIBOT_RUNNER_BASE_URL")
    runner_enabled = _parse_bool(env("SURGIBOT_RUNNER_ENABLED"), bool((runner_base_env or "").strip()))
    runner_base = ""
    if runner_enabled:
        runner_base = _normalize_runner_base(
            runner_base_env,
            runner_port=runner_port,
            fallback_host=fallback_host,
            fallback_scheme=fallback_scheme,
        )

    cfg = SurgiBotConfig(
        api_host=api_host,
        api_port=api_port,
        secret=secret,
        secret_from_env=secret_from_env,
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
        runner_enabled=runner_enabled,
        runner_port=runner_port,
        runner_base_url=runner_base,
    )
    cfg.audio_cache_dir.mkdir(parents=True, exist_ok=True)
    cfg.log_dir.mkdir(parents=True, exist_ok=True)
    cfg.data_dir.mkdir(parents=True, exist_ok=True)
    return cfg


CONFIG = load_config()

__all__ = ["CONFIG", "SurgiBotConfig", "load_config"]
