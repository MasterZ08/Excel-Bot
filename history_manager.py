"""Persistent generation history (JSON Lines: cheap appends, survives restarts)."""
from __future__ import annotations

import csv
import json
import os
import threading
from pathlib import Path

MAX_ENTRIES = 20000
FIELDS = ["timestamp", "filename", "category", "records", "format", "status", "message", "path"]


class HistoryManager:
    def __init__(self, path):
        self.path = Path(path)
        self._lock = threading.RLock()
        self._entries: list = []
        self._load()

    def _load(self) -> None:
        entries = []
        try:
            with open(self.path, encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if not line:
                        continue
                    try:
                        entries.append(json.loads(line))
                    except ValueError:
                        continue          # skip a damaged line, keep the rest
        except OSError:
            pass
        if len(entries) > MAX_ENTRIES:
            entries = entries[-MAX_ENTRIES:]
            self._entries = entries
            self._rewrite()
        self._entries = entries

    def _rewrite(self) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        tmp = self.path.with_name(self.path.name + ".tmp")
        with open(tmp, "w", encoding="utf-8") as f:
            for e in self._entries:
                f.write(json.dumps(e, ensure_ascii=False) + "\n")
        os.replace(tmp, self.path)

    def add(self, entry: dict) -> None:
        with self._lock:
            self._entries.append(entry)
            self.path.parent.mkdir(parents=True, exist_ok=True)
            with open(self.path, "a", encoding="utf-8") as f:
                f.write(json.dumps(entry, ensure_ascii=False) + "\n")
            if len(self._entries) > MAX_ENTRIES + 500:
                self._entries = self._entries[-MAX_ENTRIES:]
                self._rewrite()

    def all(self) -> list:
        with self._lock:
            return list(self._entries)

    def clear(self) -> None:
        with self._lock:
            self._entries = []
            self._rewrite()

    def export_csv(self, dest) -> int:
        with self._lock, open(dest, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=FIELDS, extrasaction="ignore")
            w.writeheader()
            for e in self._entries:
                w.writerow(e)
            return len(self._entries)
