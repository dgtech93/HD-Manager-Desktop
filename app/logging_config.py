from __future__ import annotations

import logging
import os
import sys
from pathlib import Path


def setup_logging(level: str | int = "INFO") -> None:
    """Basic app-wide logging to console + rotating file."""

    if isinstance(level, str):
        log_level = getattr(logging, level.upper(), logging.INFO)
    else:
        log_level = int(level)

    # Avoid configuring twice
    root = logging.getLogger()
    if root.handlers:
        root.setLevel(log_level)
        return

    fmt = logging.Formatter(
        "%(asctime)s %(levelname)s %(name)s - %(message)s"
    )

    root.setLevel(log_level)

    console = logging.StreamHandler()
    console.setLevel(log_level)
    console.setFormatter(fmt)
    root.addHandler(console)

    try:
        if getattr(sys, "frozen", False):
            local_appdata = os.getenv("LOCALAPPDATA")
            if local_appdata:
                logs_dir = Path(local_appdata) / "HDManagerDesktop" / "logs"
            else:
                base = Path(__file__).resolve().parent.parent
                logs_dir = base / "logs"
        else:
            base = Path(__file__).resolve().parent.parent
            logs_dir = base / "logs"
        logs_dir.mkdir(parents=True, exist_ok=True)
        file_path = logs_dir / "app.log"
        file_handler = logging.FileHandler(file_path, encoding="utf-8")
        file_handler.setLevel(log_level)
        file_handler.setFormatter(fmt)
        root.addHandler(file_handler)
    except OSError:
        # If filesystem is not writable, skip file logging.
        return

