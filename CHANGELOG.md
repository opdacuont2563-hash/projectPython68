# Changelog

## Unreleased
- Reorganized codebase under `src/surgibot` with compatibility shims for legacy entry points.
- Introduced centralized configuration (`config.py`) and rotating logging (`logging_setup.py`).
- Added background workers for audio playback and HTTP requests to keep UI threads responsive.
- Implemented caching helpers for ICD10 lookups and SQLite utility layer with WAL defaults.
- Updated PySide6 client to debounce refreshes, run network calls off the GUI thread, and reuse shared sessions.
- Added `/healthz` endpoint and audio announcement throttling on the server.
- Created project tooling files (`requirements.txt`, `pyproject.toml`, `Makefile`) and refreshed documentation.
