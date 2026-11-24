"""Shared helpers: paths, safe unique file reservation, logging, OS integration."""
from __future__ import annotations

import datetime as dt
import logging
import logging.handlers
import os
import queue
import re
import subprocess
import sys
from pathlib import Path

LOGGER_NAME = "reg"
_INVALID = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
_RESERVED = {"CON", "PRN", "AUX", "NUL", *(f"COM{i}" for i in range(1, 10)), *(f"LPT{i}" for i in range(1, 10))}


def app_dir() -> Path:
    """Folder that holds output/, configs/, logs/ (next to the EXE when frozen)."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def sanitize_prefix(prefix: str) -> str:
    p = _INVALID.sub("_", (prefix or "").strip()).strip(" .")
    if not p:
        p = "random_data"
    if p.upper() in _RESERVED:
        p = "_" + p
    return p[:80]


def reserve_files(out_dir, prefix: str, exts, now: dt.datetime | None = None) -> list[Path]:
    """Atomically create empty placeholder files `<prefix>_<seq>_<timestamp>.<ext>`.

    Uses open(..., 'x') so an existing file is NEVER overwritten. The sequence
    number continues from the highest one already present for this prefix.
    """
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    prefix = sanitize_prefix(prefix)
    pat = re.compile(rf"^{re.escape(prefix)}_(\d+)_\d{{8}}_\d{{6}}\.[A-Za-z0-9]+$")
    seq = 0
    for name in os.listdir(out_dir):
        m = pat.match(name)
        if m:
            seq = max(seq, int(m.group(1)))
    seq += 1
    while True:
        stamp = (now or dt.datetime.now()).strftime("%Y%m%d_%H%M%S")
        stem = f"{prefix}_{seq:05d}_{stamp}"
        created: list[Path] = []
        try:
            for ext in exts:
                p = out_dir / f"{stem}.{ext}"
                with open(p, "x"):
                    pass
                created.append(p)
            return created
        except FileExistsError:
            for p in created:
                try:
                    p.unlink()
                except OSError:
                    pass
            seq += 1


def open_path(path) -> None:
    path = str(path)
    if sys.platform.startswith("win"):
        os.startfile(path)  # type: ignore[attr-defined]  # noqa: S606
    elif sys.platform == "darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])


def fmt_duration(seconds: float) -> str:
    s = int(max(0, seconds))
    return f"{s // 3600:02d}:{(s % 3600) // 60:02d}:{s % 60:02d}"


class QueueLogHandler(logging.Handler):
    """Forwards log lines to the GUI event queue as ('log', text)."""

    def __init__(self, q: "queue.Queue"):
        super().__init__(level=logging.INFO)
        self.q = q
        self.setFormatter(logging.Formatter("%(asctime)s  %(levelname)-7s %(message)s", "%H:%M:%S"))

    def emit(self, record: logging.LogRecord) -> None:
        try:
            self.q.put(("log", self.format(record).split("\n")[0]))
        except Exception:  # never let logging break the app
            pass


def setup_logging(log_dir) -> logging.Logger:
    log_dir = Path(log_dir)
    log_dir.mkdir(parents=True, exist_ok=True)
    logger = logging.getLogger(LOGGER_NAME)
    logger.setLevel(logging.DEBUG)
    if not any(isinstance(h, logging.handlers.RotatingFileHandler) for h in logger.handlers):
        fh = logging.handlers.RotatingFileHandler(log_dir / "app.log", maxBytes=1_000_000, backupCount=3, encoding="utf-8")
        fh.setFormatter(logging.Formatter("%(asctime)s %(levelname)s [%(threadName)s] %(message)s"))
        logger.addHandler(fh)
    return logger
