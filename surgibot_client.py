"""Compatibility shim for legacy imports."""
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parent
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

if __name__ == "__main__":
    module = __import__(f"surgibot.{Path(__file__).stem}", fromlist=["main"])
    main = getattr(module, "main", None)
    if callable(main):
        main()
else:
    module = __import__(f"surgibot.{Path(__file__).stem}")
    globals().update({k: getattr(module, k) for k in dir(module) if not k.startswith("__")})
