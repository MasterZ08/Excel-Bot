"""Background generation: Worker thread + Engine controller.

The GUI never blocks: it talks to the Engine, and the worker reports back through a
thread-safe queue of tuples:
  ("state", "running"|"paused"|"stopping"|"idle")   ("status", text)
  ("progress", fraction, text)                      ("tick", runtime_s, countdown_s|None)
  ("cycle_done", info)  ("history", entry)  ("error", text)  ("notice", text)
"""
from __future__ import annotations

import datetime as dt
import logging
import os
import queue
import threading
import time
from pathlib import Path

from config_manager import MAX_DELAY, MIN_DELAY, ConfigError, ConfigManager, SettingsStore, validate_settings, validate_columns
from csv_exporter import export_csv
from excel_exporter import export_xlsx
from generators import DataGenerator, GenerationAborted
from history_manager import HistoryManager
from utils import LOGGER_NAME, reserve_files

log = logging.getLogger(LOGGER_NAME)
MAX_CONSECUTIVE_FAILURES = 10     # after this many failed cycles in a row the worker pauses itself
EXTS = {"XLSX": ["xlsx"], "CSV": ["csv"], "Both": ["xlsx", "csv"]}


class Worker(threading.Thread):
    def __init__(self, engine: "Engine", mode: str):
        super().__init__(name=f"GenWorker-{mode}", daemon=True)
        self.eng = engine
        self.mode = mode                       # "continuous" | "single"
        self.stop_evt = threading.Event()      # finish current file, then stop
        self.abort_evt = threading.Event()     # abandon the file in progress (close / force stop)
        self.run_evt = threading.Event()       # set = running, cleared = paused
        self.run_evt.set()
        self.trigger_evt = threading.Event()   # "Randomize again": skip the countdown
        self.state = "running"
        self._t0 = time.monotonic()
        self._paused_total = 0.0
        self._pause_started = None
        self._fails = 0
        self._last_emit = 0.0

    # ----------------------------------------------------------- helpers
    def runtime(self) -> float:
        now = time.monotonic()
        paused = self._paused_total + ((now - self._pause_started) if self._pause_started else 0.0)
        return max(0.0, now - self._t0 - paused)

    def _tick(self, countdown=None, force=False) -> None:
        now = time.monotonic()
        if force or now - self._last_emit >= 0.2:
            self._last_emit = now
            self.eng.emit("tick", self.runtime(), countdown)

    def _wait_resume(self) -> bool:
        """Block while paused. Returns False if a stop was requested."""
        entered = False
        while not self.run_evt.is_set():
            if self.stop_evt.is_set():
                break
            if not entered:
                entered = True
                self._pause_started = time.monotonic()
                self.state = "paused"
                self.eng.emit("state", "paused")
                self.eng.emit("status", "Paused")
            self.run_evt.wait(0.1)
            self._tick(None)
        if entered:
            self._paused_total += time.monotonic() - (self._pause_started or time.monotonic())
            self._pause_started = None
            if not self.stop_evt.is_set():
                self.state = "running"
                self.eng.emit("state", "running")
        return not self.stop_evt.is_set()

    def auto_pause(self, reason: str) -> None:
        self.run_evt.clear()
        self.eng.emit("notice", reason)

    def _wait_delay(self) -> bool:
        """Countdown to the next file. Stop and Randomize-Again wake it instantly."""
        self.eng.emit("status", "Waiting for next file")
        last, waited = time.monotonic(), 0.0
        while True:
            if self.stop_evt.is_set():
                return False
            if self.trigger_evt.is_set():
                self.trigger_evt.clear()
                return True
            if not self.run_evt.is_set():
                if not self._wait_resume():
                    return False
                self.eng.emit("status", "Waiting for next file")
                last = time.monotonic()           # paused time does not count
                continue
            delay = max(MIN_DELAY, min(MAX_DELAY, float(self.eng.store.get("delay"))))
            now = time.monotonic()
            waited += now - last
            last = now
            remaining = delay - waited
            if remaining <= 0:
                return True
            self._tick(remaining)
            self.stop_evt.wait(min(0.1, remaining))

    # ----------------------------------------------------------- main loop
    def run(self) -> None:
        eng = self.eng
        eng.emit("state", "running")
        try:
            while True:
                if not self._wait_resume():
                    break
                self._cycle()
                if self.stop_evt.is_set() or self.mode == "single":
                    break
                if not eng.store.get("continuous"):
                    break
                if self._fails >= MAX_CONSECUTIVE_FAILURES:
                    self._fails = 0
                    self.auto_pause(f"Paused after {MAX_CONSECUTIVE_FAILURES} failed files in a row. "
                                    "Check the Log tab, fix the problem, then press RESUME.")
                if not self._wait_delay():
                    break
        except Exception:                                   # pragma: no cover - last-resort guard
            log.exception("Worker crashed")
            eng.emit("error", "Worker stopped unexpectedly - see logs/app.log")
        finally:
            self.state = "idle"
            self._tick(None, force=True)
            if self.abort_evt.is_set():
                eng.emit("status", "Stopped (in-progress file discarded)")
            elif self.stop_evt.is_set():
                eng.emit("status", "Stopped")
            else:
                eng.emit("status", "Finished" if self._fails == 0 else "Finished with errors")
            eng.emit("state", "idle")

    # ----------------------------------------------------------- one file
    def _progress_cb(self, base: float, span: float, label: str):
        def cb(frac: float) -> None:
            now = time.monotonic()
            if frac >= 1.0 or now - self._last_emit >= 0.1:
                self._last_emit = now
                self.eng.emit("progress", base + span * frac, label)
                self.eng.emit("tick", self.runtime(), None)
        return cb

    def _cycle(self) -> None:
        eng = self.eng
        s, cols = eng.store.snapshot()
        category, n = s["category"], s["records"]
        started = time.monotonic()
        cycle = eng.next_cycle()
        eng.emit("progress", 0.0, "Starting")
        eng.emit("status", f"Generating {n:,} records")
        history_rows: list = []
        try:
            errs = validate_settings(s) + (validate_columns(cols) if category == "Custom Dataset" else [])
            if errs:
                raise ConfigError(" ".join(errs))
            out_dir = Path(s["output_dir"])
            out_dir.mkdir(parents=True, exist_ok=True)
            exts = EXTS[s["export_format"]]
            gen_span = 0.5
            ds = eng.generator.generate(category, s, cols, cycle=cycle,
                                        progress=self._progress_cb(0.0, gen_span, "Generating data"),
                                        abort=self.abort_evt)
            eng.emit("status", "Saving file")
            paths = reserve_files(out_dir, s["prefix"], exts)
            saved = 0
            try:
                for k, p in enumerate(paths):
                    span = (1 - gen_span) / len(paths)
                    cb = self._progress_cb(gen_span + k * span, span, f"Writing {p.suffix[1:].upper()}")
                    tmp = p.with_name(p.name + ".part")
                    try:
                        if p.suffix == ".xlsx":
                            export_xlsx(ds, tmp, s, cb, self.abort_evt)
                        else:
                            export_csv(ds, tmp, s, cb, self.abort_evt)
                        os.replace(tmp, p)
                        saved += 1
                        history_rows.append((p, "Success", ""))
                    except GenerationAborted:
                        raise
                    except Exception as e:                  # one bad file must not lose the others
                        log.exception("Failed writing %s", p)
                        history_rows.append((p, "Failed", f"{type(e).__name__}: {e}"))
                        eng.emit("error", f"Could not write {p.name}: {e}")
                    finally:
                        try:
                            tmp.unlink()
                        except OSError:
                            pass
            finally:
                for p in paths:                             # remove placeholders that never got content
                    if not any(p == hp for hp, st, _ in history_rows if st == "Success"):
                        try:
                            p.unlink()
                        except OSError:
                            pass
            del ds
            if saved:
                self._fails = 0
            else:
                self._fails += 1
            try:
                eng.cfg.save_state(eng.generator.seq)
            except OSError:
                log.warning("Could not save ID counters")
            eng.emit("progress", 1.0, "Complete" if saved else "Failed")
            eng.emit("cycle_done", {"records": n if saved else 0, "files": saved,
                                    "seconds": time.monotonic() - started, "category": category})
            if saved:
                log.info("Generated %s record(s) -> %s", f"{n:,}", ", ".join(p.name for p, st, _ in history_rows if st == "Success"))
        except GenerationAborted:
            history_rows = [r for r in history_rows if r[1] == "Success"] + \
                [(Path(s["output_dir"]) / "(cancelled)", "Cancelled", "Stopped before the file was finished")]
            log.info("Generation cancelled")
        except ConfigError as e:
            self._fails += 1
            history_rows = [(Path(s["output_dir"]) / "(not created)", "Failed", str(e))]
            log.error("Invalid configuration: %s", e)
            eng.emit("error", str(e))
            eng.emit("progress", 0.0, "Invalid settings")
            if self.mode == "continuous":
                self.auto_pause(f"Generation paused: {e}\nFix the setting, then press RESUME.")
        except Exception as e:
            self._fails += 1
            history_rows = [(Path(s["output_dir"]) / "(not created)", "Failed", f"{type(e).__name__}: {e}")]
            log.exception("Generation failed")
            eng.emit("error", f"{type(e).__name__}: {e}")
            eng.emit("progress", 0.0, "Failed - will retry next cycle" if self.mode == "continuous" else "Failed")
        for p, status, msg in history_rows:
            entry = {"timestamp": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "filename": p.name,
                     "category": category, "records": n, "format": p.suffix[1:].upper() or s["export_format"],
                     "status": status, "message": msg, "path": str(p)}
            try:
                eng.history.add(entry)
            except OSError:
                log.warning("Could not write history")
            eng.emit("history", entry)


class Engine:
    """What the GUI talks to. All methods are cheap and non-blocking."""

    def __init__(self, store: SettingsStore, history: HistoryManager, events: queue.Queue, cfg: ConfigManager):
        self.store, self.history, self.events, self.cfg = store, history, events, cfg
        self.generator = DataGenerator(cfg.load_state())
        self.worker: Worker | None = None
        self._cycle = 0
        self._lock = threading.Lock()

    def emit(self, *event) -> None:
        self.events.put(event)

    def next_cycle(self) -> int:
        with self._lock:
            c = self._cycle
            self._cycle += 1
            return c

    @property
    def active(self) -> bool:
        return self.worker is not None and self.worker.is_alive()

    @property
    def mode(self) -> str | None:
        return self.worker.mode if self.active else None

    def start(self, mode: str = "continuous", reset_cycle: bool = True) -> bool:
        """START (mode from the Continuous toggle) or GENERATE ONE ('single')."""
        if self.worker is not None and self.worker.is_alive():
            self.worker.join(0.5)          # it may be finishing a moment after reporting "idle"
        if self.active:
            return False
        if reset_cycle:
            self._cycle = 0
        self.worker = Worker(self, mode)
        self.worker.start()
        return True

    def start_run(self) -> bool:
        """START button: continuous if the toggle is on, otherwise exactly one file."""
        return self.start("continuous" if self.store.get("continuous") else "single", reset_cycle=True)

    def generate_one(self) -> bool:
        return self.start("single", reset_cycle=False)

    def pause(self) -> None:
        if self.active:
            self.worker.run_evt.clear()

    def resume(self) -> None:
        if self.active:
            self.worker.run_evt.set()

    def randomize(self) -> bool:
        """Immediately generate another dataset with the current settings."""
        if self.active:
            if self.worker.mode == "continuous" and self.worker.run_evt.is_set():
                self.worker.trigger_evt.set()
                return True
            return False
        return self.start("single", reset_cycle=False)

    def stop(self, force: bool = False) -> None:
        w = self.worker
        if w is None or not w.is_alive():
            return
        first = not w.stop_evt.is_set()
        w.stop_evt.set()
        w.run_evt.set()                # wake it if paused
        if force:
            w.abort_evt.set()
        if first:
            self.emit("state", "stopping")
            self.emit("status", "Stopping (finishing current file)")

    def shutdown(self, timeout: float = 5.0) -> None:
        """Safely terminate the worker (used when the window closes)."""
        w = self.worker
        if w is not None and w.is_alive():
            w.stop_evt.set()
            w.abort_evt.set()
            w.run_evt.set()
            w.join(timeout)
