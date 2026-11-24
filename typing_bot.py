"""Excel Typing Bot - fills an open Excel sheet cell by cell at a human-looking pace.

Hobby / demo / testing tool. Run:  python typing_bot.py   (or run_bot.bat)

Modes
  Excel (COM)  - writes into the open workbook through Excel itself (most reliable, Windows + desktop Excel).
                 Optional "animate typing" shows each value appearing letter by letter.
  Keystrokes   - presses real keys (Tab / Enter) in whatever window has focus (needs pyautogui).

How to use: pick a data source, press START, then click the FIRST cell in Excel during the countdown.
Stop at any time with the STOP button, the ESC key, or (keystroke mode) by moving the mouse to a screen corner.
"""
from __future__ import annotations

import csv
import queue
import random
import sys
import threading
import time
import unicodedata
from pathlib import Path


# ------------------------------------------------------------------ data sources
def load_csv(path) -> list:
    with open(path, newline="", encoding="utf-8-sig") as f:
        return [row for row in csv.reader(f)]


def _fmt(v) -> str:
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float):
        return f"{v:.2f}"
    if hasattr(v, "isoformat"):
        return v.isoformat()
    return str(v)


def sample_rows(category: str, n: int, with_header: bool = True) -> list:
    """Rows from the generator project (generators.py must sit next to this file)."""
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    import config_manager as cm
    from generators import DataGenerator
    settings, cols = cm.ConfigManager().defaults()
    ds = DataGenerator().generate(category, settings, cols, records=n)
    rows = [[_fmt(v) for v in r] for r in ds.rows]
    return ([ds.headers] if with_header else []) + rows


def to_excel_value(text: str):
    """Give Excel real numbers instead of text where a value is plainly numeric."""
    t = text.strip()
    if t == "":
        return ""
    try:
        if t.lstrip("-").isdigit() and not (len(t.lstrip("-")) > 1 and t.lstrip("-").startswith("0")):
            return int(t)
        return float(t) if t.replace(".", "", 1).lstrip("-").isdigit() and t.count(".") == 1 else text
    except ValueError:
        return text


# ------------------------------------------------------------------ pacing
class Pacer:
    def __init__(self, seconds_per_cell: float, randomness: float, stop: threading.Event):
        self.base, self.rnd, self.stop = seconds_per_cell, randomness, stop

    def sleep(self, seconds: float) -> bool:
        """Interruptible sleep. Returns False if a stop was requested."""
        end = time.monotonic() + seconds
        while time.monotonic() < end:
            if self.stop.is_set() or esc_pressed():
                self.stop.set()
                return False
            time.sleep(min(0.03, max(0.0, end - time.monotonic())))
        return not self.stop.is_set()

    def cell_delay(self) -> float:
        d = self.base * random.uniform(1 - self.rnd, 1 + self.rnd)
        if random.random() < 0.06 * self.rnd:          # an occasional "thinking" pause
            d += random.uniform(0.8, 2.2) * max(self.base, 0.3)
        return max(0.02, d)


def esc_pressed() -> bool:
    if sys.platform.startswith("win"):
        import ctypes
        return bool(ctypes.windll.user32.GetAsyncKeyState(0x1B) & 0x8000)
    return False


# ------------------------------------------------------------------ backends
class ExcelCom:
    name = "Excel (COM)"

    def __init__(self, animate: bool):
        import pythoncom  # noqa: F401
        import win32com.client
        self.app = win32com.client.GetActiveObject("Excel.Application")
        cell = self.app.ActiveCell
        if cell is None:
            raise RuntimeError("Open a workbook in Excel and click a starting cell first.")
        self.r0, self.c0 = cell.Row, cell.Column
        self.sheet = self.app.ActiveSheet
        self.animate = animate

    def enter(self, r, c, text, pacer: Pacer) -> bool:
        cell = self.sheet.Cells(self.r0 + r, self.c0 + c)
        cell.Select()
        if self.animate and 0 < len(text) <= 40 and not text.replace(".", "").replace("-", "").isdigit():
            step = max(0.02, pacer.base / (len(text) + 2))
            for i in range(1, len(text) + 1):
                cell.Value = text[:i]
                if not pacer.sleep(step * random.uniform(0.6, 1.4)):
                    return False
        cell.Value = to_excel_value(text)
        return True

    def end_row(self, r, ncols) -> None:
        self.sheet.Cells(self.r0 + r + 1, self.c0).Select()


class Keystrokes:
    name = "Keystrokes"

    def __init__(self, animate: bool):
        import pyautogui
        pyautogui.FAILSAFE = True
        pyautogui.PAUSE = 0
        self.pg = pyautogui
        self.animate = animate

    @staticmethod
    def _ascii(text: str) -> str:
        return unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode()

    def enter(self, r, c, text, pacer: Pacer) -> bool:
        text = self._ascii(text)
        if self.animate:
            for ch in text:
                self.pg.write(ch)
                if not pacer.sleep(random.uniform(0.02, 0.09)):
                    return False
        else:
            self.pg.write(text)
        self.pg.press("tab")
        return True

    def end_row(self, r, ncols) -> None:
        self.pg.press("enter")          # Enter after Tabs returns to the row's first column


# ------------------------------------------------------------------ runner
def run_bot(rows: list, backend, pacer: Pacer, on_progress=None) -> str:
    total = sum(len(r) for r in rows) or 1
    done = 0
    for r, row in enumerate(rows):
        for c, text in enumerate(row):
            if pacer.stop.is_set():
                return "stopped"
            if not backend.enter(r, c, text, pacer):
                return "stopped"
            done += 1
            if on_progress:
                on_progress(done, total)
            if not pacer.sleep(pacer.cell_delay()):
                return "stopped"
        backend.end_row(r, len(row))
        if not pacer.sleep(pacer.cell_delay() * 1.5):     # moving to the next line
            return "stopped"
    return "finished"


# ------------------------------------------------------------------ GUI (stdlib tkinter)
def main() -> None:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
    from config_manager import CATEGORIES

    root = tk.Tk()
    root.title("Excel Typing Bot")
    root.geometry("520x560")
    q: queue.Queue = queue.Queue()
    stop = threading.Event()
    v = {
        "source": tk.StringVar(value="Sample data"), "csv": tk.StringVar(), "cat": tk.StringVar(value=CATEGORIES[1]),
        "rows": tk.StringVar(value="20"), "header": tk.BooleanVar(value=True), "mode": tk.StringVar(value="Excel (COM)"),
        "speed": tk.DoubleVar(value=0.8), "rnd": tk.DoubleVar(value=0.5), "anim": tk.BooleanVar(value=True),
        "count": tk.StringVar(value="5"),
    }
    pad = {"padx": 12, "pady": 4, "sticky": "ew"}
    f = ttk.Frame(root, padding=8)
    f.pack(fill="both", expand=True)
    f.columnconfigure(1, weight=1)

    def row(i, label, widget):
        ttk.Label(f, text=label).grid(row=i, column=0, sticky="w", padx=12)
        widget.grid(row=i, column=1, **pad)

    row(0, "Data source", ttk.Combobox(f, textvariable=v["source"], values=["Sample data", "CSV file"], state="readonly"))
    csvf = ttk.Frame(f)
    ttk.Entry(csvf, textvariable=v["csv"]).pack(side="left", fill="x", expand=True)
    ttk.Button(csvf, text="Browse…", command=lambda: v["csv"].set(
        filedialog.askopenfilename(filetypes=[("CSV", "*.csv")]) or v["csv"].get())).pack(side="left", padx=4)
    row(1, "CSV file", csvf)
    row(2, "Sample category", ttk.Combobox(f, textvariable=v["cat"], values=[c for c in CATEGORIES if c != "Custom Dataset"], state="readonly"))
    row(3, "Sample rows", ttk.Entry(f, textvariable=v["rows"]))
    row(4, "Include header row", ttk.Checkbutton(f, variable=v["header"]))
    row(5, "Mode", ttk.Combobox(f, textvariable=v["mode"], values=["Excel (COM)", "Keystrokes"], state="readonly"))
    row(6, "Seconds per cell", ttk.Scale(f, from_=0.1, to=3.0, variable=v["speed"]))
    row(7, "Randomness (0-1)", ttk.Scale(f, from_=0.0, to=1.0, variable=v["rnd"]))
    row(8, "Animate typing", ttk.Checkbutton(f, variable=v["anim"]))
    row(9, "Countdown (s)", ttk.Entry(f, textvariable=v["count"]))
    status = tk.StringVar(value="Ready. Press START, then click the first Excel cell.")
    ttk.Label(f, textvariable=status, wraplength=470, justify="left").grid(row=10, column=0, columnspan=2, padx=12, pady=10, sticky="w")
    bar = ttk.Progressbar(f, maximum=100)
    bar.grid(row=11, column=0, columnspan=2, padx=12, sticky="ew")
    btns = ttk.Frame(f)
    btns.grid(row=12, column=0, columnspan=2, pady=12)
    start_btn = ttk.Button(btns, text="START")
    stop_btn = ttk.Button(btns, text="STOP", command=lambda: stop.set(), state="disabled")
    start_btn.pack(side="left", padx=6)
    stop_btn.pack(side="left", padx=6)
    ttk.Label(f, text="Stop: STOP button, ESC key, or mouse to a screen corner (keystroke mode).",
              foreground="gray").grid(row=13, column=0, columnspan=2, padx=12)

    def worker(rows, mode, speed, rnd, anim, countdown):
        try:
            for s in range(countdown, 0, -1):
                q.put(("status", f"Click the first Excel cell now… starting in {s}s"))
                if stop.wait(1):
                    q.put(("done", "Stopped.")); return
            backend = ExcelCom(anim) if mode == "Excel (COM)" else Keystrokes(anim)
            q.put(("status", f"Typing with {backend.name}…"))
            res = run_bot(rows, backend, Pacer(speed, rnd, stop), lambda d, t: q.put(("prog", 100 * d / t)))
            q.put(("done", "Finished." if res == "finished" else "Stopped."))
        except ImportError as e:
            q.put(("done", f"Missing package: {e.name}. Run: pip install -r requirements-bot.txt"))
        except Exception as e:
            q.put(("done", f"Error: {e}"))

    def start():
        try:
            if v["source"].get() == "CSV file":
                rows = load_csv(v["csv"].get())
                if not v["header"].get() and rows:
                    rows = rows[1:]
            else:
                rows = sample_rows(v["cat"].get(), max(1, min(5000, int(v["rows"].get()))), v["header"].get())
            if not rows:
                raise ValueError("No data to type.")
            countdown = max(1, min(60, int(v["count"].get())))
        except Exception as e:
            messagebox.showerror("Excel Typing Bot", str(e)); return
        stop.clear()
        start_btn.state(["disabled"]); stop_btn.state(["!disabled"])
        bar["value"] = 0
        threading.Thread(target=worker, daemon=True, args=(
            rows, v["mode"].get(), v["speed"].get(), v["rnd"].get(), v["anim"].get(), countdown)).start()

    start_btn.configure(command=start)

    def poll():
        try:
            while True:
                k, val = q.get_nowait()
                if k == "status": status.set(val)
                elif k == "prog": bar["value"] = val
                else:
                    status.set(val); start_btn.state(["!disabled"]); stop_btn.state(["disabled"])
        except queue.Empty:
            pass
        root.after(100, poll)

    poll()
    root.protocol("WM_DELETE_WINDOW", lambda: (stop.set(), root.destroy()))
    root.mainloop()


if __name__ == "__main__":
    main()
