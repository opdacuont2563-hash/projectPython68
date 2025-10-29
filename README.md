# SurgiBot — Minimal Clean Pack

This pack contains only the core Python files you said you actually use, organized in one place:

- `registry_patient_connect.py` — shared UI helpers & lookup, used by the client.  
- `surgibot_client.py` — PySide6 client (OR monitor + schedule).  
- `surgibot_server.py` — Tkinter + Flask API server (with optional Google Sheets + announcements).  
- `icd10_catalog.py` — in‑memory catalog + user additions (no Excel dependency).

## Quick Start

```bash
# 1) Create & activate a venv (recommended)
python -m venv .venv
# Windows: .venv\Scripts\activate
# macOS/Linux: source .venv/bin/activate

# 2) Install deps
pip install -r requirements.txt

# 3) Configure environment (optional but recommended)
#    Copy .env.example to .env and edit values as needed.
#    The server reads SURGIBOT_* variables, the client uses SURGIBOT_CLIENT_*.
copy .env.example .env  # (Windows)
# or
cp .env.example .env    # (macOS/Linux)

# 4) Run server
python surgibot_server.py

# 5) Run client (in another terminal)
python surgibot_client.py
```

## Notes

- The client and server defaults are set so they can talk on `http://127.0.0.1:8088` using the same secret token.
- Google Sheets is **optional**. If credentials are not provided, the server will gracefully run without Sheets.
- Text‑to‑speech uses `gTTS` + `pygame` and `pyttsx3` on Windows. If you don't need audio announcements, you can remove those packages.
- If you later want a more modular structure, you can place these files under a package (e.g., `src/surgibot/`) and adjust imports to relative ones. For now, everything works with the current names.

## Minimal tree (this pack)

```
surgibot_minimal/
├─ icd10_catalog.py
├─ registry_patient_connect.py
├─ surgibot_client.py
├─ surgibot_server.py
├─ requirements.txt
└─ .env.example
```
