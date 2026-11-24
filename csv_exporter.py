"""CSV export using Python's csv module (UTF-8 with BOM so Excel opens it correctly)."""
from __future__ import annotations

import csv
import datetime as dt

from generators import Dataset, GenerationAborted


def _fmt(v, col):
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, dt.datetime):
        return v.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(v, dt.date):
        return v.isoformat()
    if isinstance(v, float):
        return f"{v:.{max(col.decimals, 0)}f}" if col.kind in ("decimal", "currency") else repr(v)
    if isinstance(v, str) and v[:1] in ("=", "@"):
        return "'" + v        # neutralise spreadsheet formula injection
    return v


def export_csv(ds: Dataset, path, settings: dict | None = None, progress=None, abort=None) -> None:
    n = len(ds.rows)
    step = max(1, n // 100)
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(ds.headers)
        for i, row in enumerate(ds.rows):
            if i % step == 0:
                if abort is not None and abort.is_set():
                    raise GenerationAborted()
                if progress:
                    progress(i / max(1, n))
            w.writerow([_fmt(v, c) for v, c in zip(row, ds.cols)])
    if progress:
        progress(1.0)
