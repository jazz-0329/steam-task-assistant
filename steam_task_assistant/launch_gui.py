from __future__ import annotations

import sys
import traceback
from pathlib import Path


APP_DIR = Path(__file__).resolve().parent
LOG_PATH = APP_DIR / "启动错误.log"


def main() -> None:
    try:
        sys.path.insert(0, str(APP_DIR))
        from app import main as app_main

        app_main()
    except Exception:
        LOG_PATH.write_text(traceback.format_exc(), encoding="utf-8")
        raise


if __name__ == "__main__":
    main()
