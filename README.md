# SurgiBot

Refactored layout that keeps the original entry points (`surgibot_server.py`, `surgibot_client.py`,
`registry_patient_connect.py`, `icd10_catalog.py`) while moving the actual implementation into
`src/surgibot`. The legacy filenames now act as shims so existing launch commands continue to work.

## Project layout
```
project_root/
├─ assets/
├─ data/
├─ src/
│  └─ surgibot/
│     ├─ surgibot_server.py
│     ├─ surgibot_client.py
│     ├─ registry_patient_connect.py
│     ├─ icd10_catalog.py
│     ├─ config.py
│     ├─ logging_setup.py
│     ├─ workers/
│     │   ├─ audio_worker.py
│     │   └─ io_worker.py
│     └─ utils/
│         ├─ cache.py
│         └─ db.py
├─ Makefile
├─ requirements.txt
├─ pyproject.toml
├─ CHANGELOG.md
├─ README.md
└─ .env.example
```

## Quick start
```bash
# 1) Optional virtual environment
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate

# 2) Install dependencies
make install

# 3) Configure environment variables
cp .env.example .env
# edit .env to suit your deployment (API host, token, Google Sheets, etc.)

# 4) Start the server
make run-server

# 5) Start the client (new terminal)
make run-client
```

## Common tasks
```bash
make format   # black src
make lint     # ruff check src
```

## Environment variables
All configurable settings live in `src/surgibot/config.py` and read from the environment (via `.env`).
Important keys:

- `SURGIBOT_API_HOST`, `SURGIBOT_API_PORT`, `SURGIBOT_SECRET` – Flask/Waitress server settings.
- `SURGIBOT_CLIENT_HOST`, `SURGIBOT_CLIENT_PORT`, `SURGIBOT_CLIENT_TIMEOUT` – client defaults.
- `SURGIBOT_ANNOUNCE_MINUTES`, `SURGIBOT_AUTO_DELETE_MIN`, `SURGIBOT_ANNOUNCE_TTL` – announcement cadence.
- `SURGIBOT_AUDIO_CACHE`, `SURGIBOT_DATA_DIR`, `SURGIBOT_LOG_DIR` – local paths for caches and logs.
- `SURGIBOT_SPREADSHEET_ID`, `SURGIBOT_GCP_CREDENTIALS_*` – optional Google Sheets integration.

Defaults in `.env.example` preserve the previous behaviour so the application still works out-of-the-box
on `http://127.0.0.1:8088` using the historical secret token.

## Troubleshooting
- Ensure `assets/cache/` and `data/` are writable; the audio worker caches synthesized announcements here and
  SQLite databases use `data/` with WAL mode.
- Audio playback relies on `pygame`. On headless servers consider installing a dummy audio driver
  or disable announcements via configuration.
- Google Sheets is optional. If credentials are absent the server logs a warning and continues without syncing.

Refer to `CHANGELOG.md` for a summary of the recent refactor work.
