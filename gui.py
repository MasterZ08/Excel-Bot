"""CustomTkinter dashboard. All heavy work happens in worker.py; this module only
reads/writes widgets on the Tk thread and polls a thread-safe event queue."""
from __future__ import annotations

import logging
import math
import os
import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk

from config_manager import (APP_NAME, APPEARANCES, CATEGORIES, CUSTOM_TYPES, DEFAULTS, FORMATS, ID_MODES, LOCALES,
                            MAX_DELAY, MAX_RECORDS, MIN_DELAY, TYPE_LABELS, ConfigManager, SettingsStore,
                            default_custom_columns, normalize_column, normalize_settings, validate_columns,
                            validate_settings)
from generators import HAVE_FAKER, DataGenerator
from history_manager import HistoryManager
from utils import LOGGER_NAME, QueueLogHandler, fmt_duration, open_path
from worker import Engine

GREEN, GREEN_H = "#2e7d32", "#1b5e20"
RED, RED_H = "#c62828", "#8e0000"
AMBER, AMBER_H = "#ef6c00", "#bf360c"
BOOL_KEYS = {k for k, v in DEFAULTS.items() if isinstance(v, bool)}
LABEL_TO_TYPE = {v: k for k, v in TYPE_LABELS.items()}
EDITOR_FIELDS = [   # (field, label, types that use it)
    ("min", "Min", ("text", "integer", "decimal", "currency", "date")),
    ("max", "Max", ("text", "integer", "decimal", "currency", "date")),
    ("decimals", "Decimal places", ("decimal",)),
    ("id_mode", "ID mode", ("id",)),
    ("prefix", "Prefix", ("id",)),
    ("suffix", "Suffix", ("id",)),
    ("width", "Digits", ("id",)),
    ("true_pct", "% True", ("boolean",)),
    ("values", "Values (one per line)", ("custom_list",)),
]
MINMAX_LABELS = {"text": ("Min words", "Max words"), "integer": ("Min", "Max"), "decimal": ("Min", "Max"),
                 "currency": ("Min", "Max"), "date": ("From (YYYY-MM-DD)", "To (YYYY-MM-DD)")}
log = logging.getLogger(LOGGER_NAME)


def fmt_cell(v) -> str:
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float):
        return f"{v:,.2f}"
    if hasattr(v, "isoformat"):
        return v.isoformat()
    return str(v)


class App(ctk.CTk):
    def __init__(self, cfg: ConfigManager):
        super().__init__()
        self.cfg = cfg
        settings, columns = cfg.load_settings()
        ctk.set_appearance_mode(settings["appearance"].lower())
        ctk.set_default_color_theme("blue")
        self.title(APP_NAME)
        self.geometry("1400x880")
        self.minsize(1120, 720)

        self.store = SettingsStore(settings, columns)
        self.columns = [dict(c) for c in columns]
        self.events: queue.Queue = queue.Queue()
        self.history = HistoryManager(cfg.history_path)
        self.engine = Engine(self.store, self.history, self.events, cfg)
        log.addHandler(QueueLogHandler(self.events))
        self.preview_gen = DataGenerator()
        self._preview_busy = False
        self._preview_idx = 0
        self.run_state = "idle"
        self._loading = False
        self._closing = False
        self.sel_idx = None
        self._ed_loading = False
        self.files_done = 0
        self.records_done = 0
        self._hist_by_iid: dict = {}

        self.vars: dict = {}
        for k, v in settings.items():
            self.vars[k] = tk.BooleanVar(self, value=v) if k in BOOL_KEYS else tk.StringVar(self, value=str(v))

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self._build_sidebar()
        self._build_main()
        self._style_trees()
        for var in self.vars.values():
            var.trace_add("write", self._on_var_change)
        self._refresh_col_tree(select=0 if self.columns else None)
        self._load_history_tree()
        self._update_buttons()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self._poll_id = self.after(80, self._poll_events)
        self.after(300, self._refresh_preview)
        log.info("Ready. Data source: %s", self.preview_gen.engine_name)

    # ================================================================== layout
    def _build_sidebar(self) -> None:
        side = ctk.CTkScrollableFrame(self, width=350, corner_radius=0, label_text="")
        side.grid(row=0, column=0, sticky="nsew")
        side.grid_columnconfigure(0, weight=1)
        self.side = side
        bold = ctk.CTkFont(size=13, weight="bold")

        ctk.CTkLabel(side, text="Random Excel\nData Generator", font=ctk.CTkFont(size=22, weight="bold"),
                     justify="left").grid(row=0, column=0, sticky="w", padx=14, pady=(14, 2))
        ctk.CTkLabel(side, text="Synthetic data only - no real personal data", text_color="gray60",
                     font=ctk.CTkFont(size=11)).grid(row=1, column=0, sticky="w", padx=14, pady=(0, 8))

        # --- controls
        box = ctk.CTkFrame(side)
        box.grid(row=2, column=0, sticky="ew", padx=10, pady=4)
        box.grid_columnconfigure((0, 1), weight=1)
        mk = lambda text, cmd, fg=None, hv=None: ctk.CTkButton(  # noqa: E731
            box, text=text, command=cmd, height=36, font=bold, **({"fg_color": fg, "hover_color": hv} if fg else {}))
        self.btn_start = mk("▶  START", self._on_start, GREEN, GREEN_H)
        self.btn_stop = mk("■  STOP", self._on_stop, RED, RED_H)
        self.btn_pause = mk("❚❚  PAUSE", self._on_pause, AMBER, AMBER_H)
        self.btn_resume = mk("▶  RESUME", self._on_resume)
        self.btn_one = mk("Generate One File", self._on_generate_one)
        self.btn_rand = mk("Randomize Again", self._on_randomize)
        for i, b in enumerate((self.btn_start, self.btn_stop, self.btn_pause, self.btn_resume)):
            b.grid(row=i // 2, column=i % 2, sticky="ew", padx=5, pady=5)
        self.btn_one.grid(row=2, column=0, sticky="ew", padx=5, pady=5)
        self.btn_rand.grid(row=2, column=1, sticky="ew", padx=5, pady=5)

        # --- settings form
        self._r = 10
        self._section(side, "OUTPUT")
        self._entry_row(side, "Output folder", "output_dir", wide=True)
        row = ctk.CTkFrame(side, fg_color="transparent")
        row.grid(row=self._r, column=0, sticky="ew", padx=10)
        self._r += 1
        ctk.CTkButton(row, text="Browse…", width=90, command=self._browse).pack(side="left", padx=(4, 4))
        ctk.CTkButton(row, text="Open folder", width=100, command=self._open_output).pack(side="left")
        self._entry_row(side, "Filename prefix", "prefix")

        self._section(side, "DATASET")
        self._menu_row(side, "Category", "category", CATEGORIES, self._on_category)
        self._seg_row(side, "Export format", "export_format", FORMATS)
        self._entry_row(side, "Records per file", "records")
        self._entry_row(side, f"Delay (s, {MIN_DELAY}-{MAX_DELAY})", "delay")
        self._switch_row(side, "Continuous generation", "continuous")

        self._section(side, "RANDOMIZATION")
        self._seg_row(side, "ID style", "id_mode", ID_MODES)
        self._entry_row(side, "ID prefix", "id_prefix")
        self._entry_row(side, "ID suffix", "id_suffix")
        self._switch_row(side, "Fixed random seed", "use_seed")
        self._entry_row(side, "Seed", "seed")
        self._switch_row(side, "Allow duplicate records", "allow_duplicates")
        self._entry_row(side, "Duplicates (%)", "duplicate_pct")
        self._entry_row(side, "Date from (YYYY-MM-DD)", "date_start")
        self._entry_row(side, "Date to (blank = today)", "date_end")
        self._entry_row(side, "Amount min (blank = default)", "num_min")
        self._entry_row(side, "Amount max (blank = default)", "num_max")
        self._entry_row(side, "Age min (Personal)", "age_min")
        self._entry_row(side, "Age max (Personal)", "age_max")
        self._menu_row(side, "Locale", "locale", LOCALES, lambda _v: self._refresh_preview())
        self._entry_row(side, "Currency symbol", "currency_symbol")
        self._switch_row(side, "Summary + breakdown sheets", "include_summary")

        self._section(side, "CONFIGURATION")
        cf = ctk.CTkFrame(side, fg_color="transparent")
        cf.grid(row=self._r, column=0, sticky="ew", padx=10)
        cf.grid_columnconfigure((0, 1), weight=1)
        self._r += 1
        for i, (t, c) in enumerate([("Save config…", self._save_config), ("Load config…", self._load_config),
                                    ("Restore defaults", self._restore_defaults), ("Reset ID counters", self._reset_ids)]):
            ctk.CTkButton(cf, text=t, height=30, command=c, fg_color=("gray70", "gray30"),
                          hover_color=("gray60", "gray40"), text_color=("black", "white")
                          ).grid(row=i // 2, column=i % 2, sticky="ew", padx=4, pady=4)
        src = "Faker is active" if HAVE_FAKER else "Faker not installed - using built-in generator"
        ctk.CTkLabel(side, text=src, text_color="gray60", font=ctk.CTkFont(size=11)).grid(
            row=self._r, column=0, sticky="w", padx=14, pady=(10, 14))

    def _section(self, parent, text) -> None:
        ctk.CTkLabel(parent, text=text, font=ctk.CTkFont(size=11, weight="bold"), text_color="gray60").grid(
            row=self._r, column=0, sticky="w", padx=14, pady=(14, 2))
        self._r += 1

    def _row_frame(self, parent, label):
        f = ctk.CTkFrame(parent, fg_color="transparent")
        f.grid(row=self._r, column=0, sticky="ew", padx=10, pady=2)
        f.grid_columnconfigure(1, weight=1)
        self._r += 1
        ctk.CTkLabel(f, text=label, anchor="w").grid(row=0, column=0, sticky="w", padx=(4, 8))
        return f

    def _entry_row(self, parent, label, key, wide=False) -> None:
        if wide:
            ctk.CTkLabel(parent, text=label, anchor="w").grid(row=self._r, column=0, sticky="w", padx=14)
            self._r += 1
            ctk.CTkEntry(parent, textvariable=self.vars[key]).grid(row=self._r, column=0, sticky="ew", padx=14, pady=2)
            self._r += 1
            return
        f = self._row_frame(parent, label)
        ctk.CTkEntry(f, textvariable=self.vars[key], width=130).grid(row=0, column=1, sticky="e")

    def _menu_row(self, parent, label, key, values, command=None) -> None:
        f = self._row_frame(parent, label)
        ctk.CTkOptionMenu(f, values=list(values), variable=self.vars[key], width=170, command=command).grid(
            row=0, column=1, sticky="e")

    def _seg_row(self, parent, label, key, values) -> None:
        f = self._row_frame(parent, label)
        ctk.CTkSegmentedButton(f, values=list(values), variable=self.vars[key]).grid(row=0, column=1, sticky="e")

    def _switch_row(self, parent, label, key) -> None:
        f = ctk.CTkFrame(parent, fg_color="transparent")
        f.grid(row=self._r, column=0, sticky="ew", padx=10, pady=2)
        self._r += 1
        ctk.CTkSwitch(f, text=label, variable=self.vars[key], onvalue=True, offvalue=False).pack(anchor="w", padx=4)

    def _build_main(self) -> None:
        main = ctk.CTkFrame(self, fg_color="transparent")
        main.grid(row=0, column=1, sticky="nsew", padx=(8, 12), pady=10)
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(3, weight=1)

        head = ctk.CTkFrame(main, fg_color="transparent")
        head.grid(row=0, column=0, sticky="ew")
        head.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(head, text="Dashboard", font=ctk.CTkFont(size=24, weight="bold")).grid(row=0, column=0, sticky="w")
        ctk.CTkLabel(head, text="Theme").grid(row=0, column=1, padx=(0, 6))
        ctk.CTkOptionMenu(head, values=APPEARANCES, variable=self.vars["appearance"], width=100,
                          command=self._on_appearance).grid(row=0, column=2)

        cards = ctk.CTkFrame(main, fg_color="transparent")
        cards.grid(row=1, column=0, sticky="ew", pady=(8, 4))
        cards.grid_columnconfigure((0, 1, 2, 3, 4), weight=1, uniform="card")
        self.lbl_status = self._card(cards, 0, "STATUS", "Idle", 18)
        self.lbl_files = self._card(cards, 1, "FILES GENERATED", "0")
        self.lbl_records = self._card(cards, 2, "TOTAL RECORDS", "0")
        self.lbl_runtime = self._card(cards, 3, "RUNTIME", "00:00:00")
        self.lbl_next = self._card(cards, 4, "NEXT FILE IN", "-")

        prog = ctk.CTkFrame(main, fg_color="transparent")
        prog.grid(row=2, column=0, sticky="ew", pady=(2, 8))
        prog.grid_columnconfigure(0, weight=1)
        self.bar = ctk.CTkProgressBar(prog, height=14)
        self.bar.grid(row=0, column=0, sticky="ew", padx=4)
        self.bar.set(0)
        self.lbl_progress = ctk.CTkLabel(prog, text="Ready", anchor="w", text_color="gray60")
        self.lbl_progress.grid(row=1, column=0, sticky="w", padx=4)
        self.lbl_error = ctk.CTkLabel(prog, text="", anchor="w", text_color="#e57373")
        self.lbl_error.grid(row=2, column=0, sticky="w", padx=4)

        tabs = ctk.CTkTabview(main)
        tabs.grid(row=3, column=0, sticky="nsew")
        self.tabs = tabs
        self._build_preview_tab(tabs.add("Preview"))
        self._build_custom_tab(tabs.add("Custom Dataset"))
        self._build_history_tab(tabs.add("History"))
        self._build_log_tab(tabs.add("Log"))

    def _card(self, parent, col, title, value, size=22):
        f = ctk.CTkFrame(parent, corner_radius=10)
        f.grid(row=0, column=col, sticky="nsew", padx=4)
        ctk.CTkLabel(f, text=title, font=ctk.CTkFont(size=10, weight="bold"), text_color="gray60").pack(
            anchor="w", padx=12, pady=(8, 0))
        lbl = ctk.CTkLabel(f, text=value, font=ctk.CTkFont(size=size, weight="bold"), anchor="w")
        lbl.pack(anchor="w", padx=12, pady=(0, 10))
        return lbl

    # ---------------------------------------------------------------- tabs
    def _make_tree(self, parent, columns, row=1, heights=None):
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(row, weight=1)
        wrap = ctk.CTkFrame(parent, fg_color="transparent")
        wrap.grid(row=row, column=0, sticky="nsew", pady=4)
        wrap.grid_columnconfigure(0, weight=1)
        wrap.grid_rowconfigure(0, weight=1)
        tree = ttk.Treeview(wrap, columns=columns, show="headings", selectmode="browse")
        ys = ttk.Scrollbar(wrap, orient="vertical", command=tree.yview)
        xs = ttk.Scrollbar(wrap, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        tree.grid(row=0, column=0, sticky="nsew")
        ys.grid(row=0, column=1, sticky="ns")
        xs.grid(row=1, column=0, sticky="ew")
        return tree

    def _build_preview_tab(self, tab) -> None:
        bar = ctk.CTkFrame(tab, fg_color="transparent")
        bar.grid(row=0, column=0, sticky="ew")
        ctk.CTkButton(bar, text="⟳  Refresh preview (new random values)", command=self._refresh_preview).pack(side="left")
        self.lbl_preview = ctk.CTkLabel(bar, text="", text_color="gray60")
        self.lbl_preview.pack(side="left", padx=12)
        self.preview_tree = self._make_tree(tab, ("c",))

    def _build_history_tab(self, tab) -> None:
        bar = ctk.CTkFrame(tab, fg_color="transparent")
        bar.grid(row=0, column=0, sticky="ew")
        for t, c in [("Open file", self._open_selected), ("Open folder", self._open_selected_folder),
                     ("Export history…", self._export_history), ("Clear history", self._clear_history)]:
            ctk.CTkButton(bar, text=t, width=120, command=c).pack(side="left", padx=(0, 8))
        cols = ("time", "file", "category", "records", "format", "status")
        self.hist_tree = self._make_tree(tab, cols)
        for c, w, t in [("time", 150, "Date / time"), ("file", 360, "Filename"), ("category", 150, "Category"),
                        ("records", 80, "Records"), ("format", 70, "Format"), ("status", 90, "Status")]:
            self.hist_tree.heading(c, text=t)
            self.hist_tree.column(c, width=w, anchor="w", stretch=(c == "file"))
        self.hist_tree.bind("<Double-1>", lambda _e: self._open_selected())

    def _build_log_tab(self, tab) -> None:
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)
        self.log_box = ctk.CTkTextbox(tab, font=ctk.CTkFont(family="Consolas", size=12), wrap="none")
        self.log_box.grid(row=0, column=0, sticky="nsew")
        self.log_box.configure(state="disabled")

    def _build_custom_tab(self, tab) -> None:
        tab.grid_columnconfigure(0, weight=3)
        tab.grid_columnconfigure(1, weight=2)
        tab.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(tab, text="Columns used when Category = Custom Dataset. Edits apply automatically when you "
                               "select another column or press Apply.", text_color="gray60", anchor="w",
                     wraplength=900, justify="left").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 4))
        left = ctk.CTkFrame(tab, fg_color="transparent")
        left.grid(row=1, column=0, sticky="nsew", padx=(0, 6))
        left.grid_columnconfigure(0, weight=1)
        left.grid_rowconfigure(0, weight=1)
        self.col_tree = self._make_tree(left, ("name", "type", "detail"), row=0)
        for c, w, t in [("name", 170, "Column"), ("type", 90, "Type"), ("detail", 260, "Range / values")]:
            self.col_tree.heading(c, text=t)
            self.col_tree.column(c, width=w, anchor="w")
        self.col_tree.bind("<<TreeviewSelect>>", self._on_col_select)
        bar = ctk.CTkFrame(left, fg_color="transparent")
        bar.grid(row=1, column=0, sticky="ew")
        for t, c in [("Add", self._add_col), ("Duplicate", self._dup_col), ("Remove", self._del_col),
                     ("▲ Up", lambda: self._move_col(-1)), ("▼ Down", lambda: self._move_col(1)),
                     ("Reset to sample", self._reset_cols)]:
            ctk.CTkButton(bar, text=t, width=84, height=30, command=c).pack(side="left", padx=(0, 6), pady=4)

        ed = ctk.CTkFrame(tab)
        ed.grid(row=1, column=1, sticky="nsew")
        ed.grid_columnconfigure(1, weight=1)
        self.ed_vars = {k: tk.StringVar(self) for k in ("name", "type", "min", "max", "decimals", "id_mode",
                                                        "prefix", "suffix", "width", "true_pct")}
        ctk.CTkLabel(ed, text="Column editor", font=ctk.CTkFont(size=14, weight="bold")).grid(
            row=0, column=0, columnspan=2, sticky="w", padx=12, pady=(10, 6))
        ctk.CTkLabel(ed, text="Name").grid(row=1, column=0, sticky="w", padx=12, pady=3)
        ctk.CTkEntry(ed, textvariable=self.ed_vars["name"]).grid(row=1, column=1, sticky="ew", padx=12, pady=3)
        ctk.CTkLabel(ed, text="Data type").grid(row=2, column=0, sticky="w", padx=12, pady=3)
        ctk.CTkOptionMenu(ed, values=[TYPE_LABELS[t] for t in CUSTOM_TYPES], variable=self.ed_vars["type"],
                          command=lambda _v: self._update_editor_visibility()).grid(row=2, column=1, sticky="ew", padx=12, pady=3)
        self.ed_rows: dict = {}
        for i, (field, label, types) in enumerate(EDITOR_FIELDS, start=3):
            lab = ctk.CTkLabel(ed, text=label)
            if field == "values":
                w = ctk.CTkTextbox(ed, height=130)
                self.ed_values = w
            elif field == "id_mode":
                w = ctk.CTkOptionMenu(ed, values=["Global", "Random", "Sequential"], variable=self.ed_vars[field])
            else:
                w = ctk.CTkEntry(ed, textvariable=self.ed_vars[field])
            lab.grid(row=i, column=0, sticky="nw", padx=12, pady=3)
            w.grid(row=i, column=1, sticky="ew", padx=12, pady=3)
            self.ed_rows[field] = (lab, w, types)
        ctk.CTkButton(ed, text="Apply changes", command=self._apply_editor).grid(
            row=20, column=0, columnspan=2, sticky="ew", padx=12, pady=12)

    # ================================================================== settings plumbing
    def _on_var_change(self, *_a) -> None:
        if not self._loading and not self._closing:
            self._sync_store()

    def _sync_store(self) -> None:
        raw = {k: v.get() for k, v in self.vars.items()}
        base, _ = self.store.snapshot()
        self.store.set_settings(normalize_settings(raw, base=base))

    def _push_vars(self, s: dict) -> None:
        self._loading = True
        try:
            for k, v in s.items():
                if k in self.vars:
                    self.vars[k].set(v if k in BOOL_KEYS else str(v))
        finally:
            self._loading = False

    def _save_settings(self) -> None:
        try:
            s, cols = self.store.snapshot()
            self.cfg.save_settings(s, cols)
        except OSError as e:
            log.warning("Could not save settings: %s", e)

    def _preflight(self) -> bool:
        self._commit_editor()
        self._sync_store()
        s, cols = self.store.snapshot()
        self._push_vars(s)                       # show clamped values (e.g. delay 0 -> 1)
        errs = validate_settings(s)
        if s["category"] == "Custom Dataset":
            errs += validate_columns(cols)
        if not errs:
            try:
                Path(s["output_dir"]).mkdir(parents=True, exist_ok=True)
                probe = Path(s["output_dir"]) / ".write_test"
                probe.write_text("x")
                probe.unlink()
            except OSError as e:
                errs.append(f"Cannot write to the output folder: {e}")
        if errs:
            messagebox.showerror("Check your settings", "\n".join(errs))
            return False
        return True

    # ================================================================== controls
    def _on_start(self) -> None:
        if not self._preflight():
            return
        self.files_done = self.records_done = 0
        self._update_stats()
        self.lbl_error.configure(text="")
        self._save_settings()
        if not self.engine.start_run():
            messagebox.showinfo(APP_NAME, "The previous run is still shutting down - try again in a moment.")

    def _on_generate_one(self) -> None:
        if self._preflight():
            self.engine.generate_one()

    def _on_randomize(self) -> None:
        if not self._preflight():
            return
        self.engine.randomize()
        self._refresh_preview()

    def _on_pause(self) -> None:
        self.engine.pause()
        self.lbl_status.configure(text="Pausing…")

    def _on_resume(self) -> None:
        self.engine.resume()

    def _on_stop(self) -> None:
        self.engine.stop(force=(self.run_state == "stopping"))   # a second click forces it
        self._save_settings()

    def _update_buttons(self) -> None:
        st, mode = self.run_state, self.engine.mode
        idle = st == "idle"
        on = lambda b, flag: b.configure(state="normal" if flag else "disabled")  # noqa: E731
        on(self.btn_start, idle)
        on(self.btn_stop, not idle)
        on(self.btn_pause, st == "running" and mode == "continuous")
        on(self.btn_resume, st == "paused")
        on(self.btn_one, idle)
        on(self.btn_rand, idle or (st == "running" and mode == "continuous"))
        self.btn_stop.configure(text="■  FORCE STOP" if st == "stopping" else "■  STOP")

    def _on_category(self, _v=None) -> None:
        self._refresh_preview()

    def _on_appearance(self, value) -> None:
        ctk.set_appearance_mode(value.lower())
        self.after(150, self._style_trees)

    def _browse(self) -> None:
        d = filedialog.askdirectory(initialdir=self.vars["output_dir"].get() or str(self.cfg.base), title="Output folder")
        if d:
            self.vars["output_dir"].set(os.path.normpath(d))

    def _open_output(self) -> None:
        d = Path(self.vars["output_dir"].get())
        try:
            d.mkdir(parents=True, exist_ok=True)
            open_path(d)
        except OSError as e:
            messagebox.showerror(APP_NAME, f"Cannot open folder: {e}")

    # ---------------------------------------------------------------- config buttons
    def _save_config(self) -> None:
        self._commit_editor()
        self._sync_store()
        path = filedialog.asksaveasfilename(initialdir=self.cfg.configs_dir, defaultextension=".json",
                                            initialfile="dataset_config.json", filetypes=[("JSON", "*.json")])
        if not path:
            return
        s, cols = self.store.snapshot()
        try:
            self.cfg.save_dataset_config(path, s, cols)
            log.info("Saved config %s", path)
        except OSError as e:
            messagebox.showerror(APP_NAME, f"Could not save: {e}")

    def _load_config(self) -> None:
        path = filedialog.askopenfilename(initialdir=self.cfg.configs_dir, filetypes=[("JSON", "*.json")])
        if not path:
            return
        cur, _ = self.store.snapshot()
        try:
            s, cols = self.cfg.load_dataset_config(path, cur)
        except ValueError as e:
            messagebox.showerror(APP_NAME, str(e))
            return
        self._apply_loaded(s, cols if cols else self.columns)
        log.info("Loaded config %s", path)

    def _restore_defaults(self) -> None:
        if not messagebox.askyesno(APP_NAME, "Restore ALL settings (including output folder) and the sample custom columns to defaults?"):
            return
        s, cols = self.cfg.defaults()
        self._apply_loaded(s, cols)
        ctk.set_appearance_mode(s["appearance"].lower())
        self.after(150, self._style_trees)
        self._save_settings()

    def _apply_loaded(self, s: dict, cols: list) -> None:
        self._push_vars(s)
        self.columns = [dict(c) for c in cols]
        self.store.set_settings(s)
        self.store.set_columns(self.columns)
        self.sel_idx = None
        self._refresh_col_tree(select=0 if self.columns else None)
        self._refresh_preview()

    def _reset_ids(self) -> None:
        if messagebox.askyesno(APP_NAME, "Restart sequential ID counters from 1?"):
            self.engine.generator.reset_sequences()
            try:
                self.cfg.save_state({})
            except OSError:
                pass
            log.info("Sequential ID counters reset")

    # ================================================================== custom columns editor
    @staticmethod
    def _detail(c: dict) -> str:
        t = c["type"]
        if t in ("integer", "decimal", "currency", "text"):
            return f"{c['min'] or 'default'} – {c['max'] or 'default'}"
        if t == "date":
            return f"{c['min'] or 'global'} – {c['max'] or 'global'}"
        if t == "custom_list":
            return ", ".join(c["values"][:5]) + (" …" if len(c["values"]) > 5 else "")
        if t == "id":
            return f"{c['id_mode']} {c['prefix']}…{c['suffix']}"
        if t == "boolean":
            return f"{c['true_pct']:g}% true"
        return ""

    def _col_values(self, c: dict) -> tuple:
        return (c["name"], TYPE_LABELS[c["type"]], self._detail(c))

    def _refresh_col_tree(self, select=None) -> None:
        self._ed_loading = True
        try:
            self.col_tree.delete(*self.col_tree.get_children())
            for i, c in enumerate(self.columns):
                self.col_tree.insert("", "end", iid=str(i), values=self._col_values(c))
            if select is not None and 0 <= select < len(self.columns):
                self.col_tree.selection_set(str(select))
                self.col_tree.see(str(select))
        finally:
            self._ed_loading = False
        self.sel_idx = select if select is not None and 0 <= select < len(self.columns) else None
        self._load_editor(self.sel_idx)
        self.store.set_columns(self.columns)

    def _on_col_select(self, _e=None) -> None:
        if self._ed_loading:
            return
        sel = self.col_tree.selection()
        idx = int(sel[0]) if sel else None
        if idx == self.sel_idx:
            return
        self._commit_editor()
        self.sel_idx = idx
        self._load_editor(idx)

    def _load_editor(self, idx) -> None:
        self._ed_loading = True
        try:
            c = self.columns[idx] if idx is not None and 0 <= idx < len(self.columns) else normalize_column({"name": ""})
            self.ed_vars["name"].set(c["name"] if idx is not None else "")
            self.ed_vars["type"].set(TYPE_LABELS[c["type"]])
            for k in ("min", "max", "decimals", "id_mode", "prefix", "suffix", "width", "true_pct"):
                self.ed_vars[k].set(str(c[k]))
            self.ed_values.delete("1.0", "end")
            self.ed_values.insert("1.0", "\n".join(c["values"]))
        finally:
            self._ed_loading = False
        self._update_editor_visibility()

    def _update_editor_visibility(self) -> None:
        t = LABEL_TO_TYPE.get(self.ed_vars["type"].get(), "text")
        for field, (lab, w, types) in self.ed_rows.items():
            if t in types:
                lab.grid()
                w.grid()
            else:
                lab.grid_remove()
                w.grid_remove()
        if t in MINMAX_LABELS:
            self.ed_rows["min"][0].configure(text=MINMAX_LABELS[t][0])
            self.ed_rows["max"][0].configure(text=MINMAX_LABELS[t][1])

    def _commit_editor(self) -> None:
        idx = self.sel_idx
        if idx is None or not (0 <= idx < len(self.columns)):
            return
        d = {k: v.get() for k, v in self.ed_vars.items()}
        d["type"] = LABEL_TO_TYPE.get(d["type"], "text")
        d["values"] = self.ed_values.get("1.0", "end")
        new = normalize_column(d)
        if new != self.columns[idx]:
            self.columns[idx] = new
            self.col_tree.item(str(idx), values=self._col_values(new))
            self.store.set_columns(self.columns)

    def _apply_editor(self) -> None:
        self._commit_editor()
        self._load_editor(self.sel_idx)
        if self.vars["category"].get() == "Custom Dataset":
            self._refresh_preview()

    def _unique_name(self, base: str) -> str:
        names = {c["name"].casefold() for c in self.columns}
        n, name = 2, base
        while name.casefold() in names:
            name = f"{base} {n}"
            n += 1
        return name

    def _add_col(self) -> None:
        self._commit_editor()
        self.columns.append(normalize_column({"name": self._unique_name("New Column"), "type": "text"}))
        self._refresh_col_tree(select=len(self.columns) - 1)

    def _dup_col(self) -> None:
        self._commit_editor()
        if self.sel_idx is None:
            return
        c = dict(self.columns[self.sel_idx])
        c["name"] = self._unique_name(c["name"] + " copy")
        self.columns.insert(self.sel_idx + 1, normalize_column(c))
        self._refresh_col_tree(select=self.sel_idx + 1)

    def _del_col(self) -> None:
        if self.sel_idx is None:
            return
        i = self.sel_idx
        del self.columns[i]
        self.sel_idx = None
        self._refresh_col_tree(select=min(i, len(self.columns) - 1) if self.columns else None)

    def _move_col(self, d: int) -> None:
        self._commit_editor()
        i = self.sel_idx
        if i is None or not (0 <= i + d < len(self.columns)):
            return
        self.columns[i], self.columns[i + d] = self.columns[i + d], self.columns[i]
        self._refresh_col_tree(select=i + d)

    def _reset_cols(self) -> None:
        if messagebox.askyesno(APP_NAME, "Replace the custom columns with the sample set?"):
            self.columns = default_custom_columns()
            self.sel_idx = None
            self._refresh_col_tree(select=0)

    # ================================================================== preview / history / log
    def _refresh_preview(self) -> None:
        if self._preview_busy or self._closing:
            return
        self._commit_editor()
        self._sync_store()
        self._preview_busy = True
        s, cols = self.store.snapshot()
        idx = self._preview_idx
        self._preview_idx += 1
        self.lbl_preview.configure(text="Generating preview…")
        threading.Thread(target=self._preview_job, args=(s, cols, idx), daemon=True, name="Preview").start()

    def _preview_job(self, s, cols, idx) -> None:
        try:
            self.preview_gen.reset_sequences()
            ds = self.preview_gen.generate(s["category"], s, cols, records=20, cycle=idx)
            self.events.put(("preview", ds.headers, ds.rows, ds.category))
        except Exception as e:  # shown in the label; the app keeps working
            self.events.put(("preview_error", str(e)))
        finally:
            self._preview_busy = False

    def _show_preview(self, headers, rows, category) -> None:
        t = self.preview_tree
        t.delete(*t.get_children())
        t.configure(columns=headers)
        for h in headers:
            t.heading(h, text=h)
            t.column(h, width=max(90, min(260, 9 * len(h) + 30)), anchor="w", stretch=False)
        for i, r in enumerate(rows):
            t.insert("", "end", values=[fmt_cell(v) for v in r], tags=("odd" if i % 2 else "even",))
        self.lbl_preview.configure(text=f"{category} - first {len(rows)} records (preview only, not saved)")

    def _load_history_tree(self) -> None:
        self.hist_tree.delete(*self.hist_tree.get_children())
        self._hist_by_iid.clear()
        for e in reversed(self.history.all()[-2000:]):
            self._insert_history(e, end=True)

    def _insert_history(self, e: dict, end=False) -> None:
        tag = {"Success": "ok", "Failed": "bad", "Cancelled": "warn"}.get(e.get("status"), "")
        iid = self.hist_tree.insert("", "end" if end else 0, values=(
            e.get("timestamp", ""), e.get("filename", ""), e.get("category", ""), f"{e.get('records', 0):,}",
            e.get("format", ""), e.get("status", "")), tags=(tag,))
        self._hist_by_iid[iid] = e

    def _selected_entry(self):
        sel = self.hist_tree.selection()
        return self._hist_by_iid.get(sel[0]) if sel else None

    def _open_selected(self) -> None:
        e = self._selected_entry()
        if not e:
            return
        p = Path(e.get("path", ""))
        if p.is_file():
            try:
                open_path(p)
            except OSError as ex:
                messagebox.showerror(APP_NAME, f"Cannot open file: {ex}")
        else:
            messagebox.showinfo(APP_NAME, "That file no longer exists (or was never created).")

    def _open_selected_folder(self) -> None:
        e = self._selected_entry()
        folder = Path(e["path"]).parent if e and e.get("path") else Path(self.vars["output_dir"].get())
        if folder.is_dir():
            open_path(folder)

    def _export_history(self) -> None:
        path = filedialog.asksaveasfilename(defaultextension=".csv", initialfile="generation_history.csv",
                                            filetypes=[("CSV", "*.csv")])
        if path:
            try:
                n = self.history.export_csv(path)
                messagebox.showinfo(APP_NAME, f"Exported {n:,} history rows.")
            except OSError as e:
                messagebox.showerror(APP_NAME, f"Export failed: {e}")

    def _clear_history(self) -> None:
        if messagebox.askyesno(APP_NAME, "Clear the generation history? (Generated files are not deleted.)"):
            self.history.clear()
            self._load_history_tree()

    def _append_log(self, line: str) -> None:
        self.log_box.configure(state="normal")
        self.log_box.insert("end", line + "\n")
        if int(self.log_box.index("end-1c").split(".")[0]) > 2000:
            self.log_box.delete("1.0", "300.0")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def _style_trees(self) -> None:
        dark = ctk.get_appearance_mode() == "Dark"
        bg, fg = ("#2b2b2b", "#e6e6e6") if dark else ("#ffffff", "#1a1a1a")
        alt = "#333333" if dark else "#f2f6fb"
        head = "#1f538d" if dark else "#3a7ebf"
        st = ttk.Style(self)
        st.theme_use("clam")
        st.configure("Treeview", background=bg, foreground=fg, fieldbackground=bg, rowheight=24, borderwidth=0)
        st.configure("Treeview.Heading", background=head, foreground="white", relief="flat", font=("Segoe UI", 10, "bold"))
        st.map("Treeview", background=[("selected", "#1f6aa5")], foreground=[("selected", "white")])
        st.map("Treeview.Heading", background=[("active", head)])
        self.preview_tree.tag_configure("even", background=bg)
        self.preview_tree.tag_configure("odd", background=alt)
        self.hist_tree.tag_configure("ok", foreground="#66bb6a" if dark else "#2e7d32")
        self.hist_tree.tag_configure("bad", foreground="#ef5350" if dark else "#c62828")
        self.hist_tree.tag_configure("warn", foreground="#ffa726" if dark else "#e65100")

    # ================================================================== event pump
    def _update_stats(self) -> None:
        self.lbl_files.configure(text=f"{self.files_done:,}")
        self.lbl_records.configure(text=f"{self.records_done:,}")

    def _poll_events(self) -> None:
        if self._closing:
            return
        try:
            for _ in range(300):
                try:
                    ev = self.events.get_nowait()
                except queue.Empty:
                    break
                self._handle(ev)
        except Exception:
            log.exception("GUI event error")
        self._poll_id = self.after(80, self._poll_events)

    def _handle(self, ev: tuple) -> None:
        kind = ev[0]
        if kind == "state":
            self.run_state = ev[1]
            self._update_buttons()
            if ev[1] == "idle":
                self.lbl_next.configure(text="-")
                self._save_settings()
        elif kind == "status":
            self.lbl_status.configure(text=ev[1] if len(ev[1]) < 26 else ev[1][:24] + "…")
        elif kind == "progress":
            self.bar.set(max(0.0, min(1.0, ev[1])))
            self.lbl_progress.configure(text=f"{ev[1] * 100:.0f}%  -  {ev[2]}")
        elif kind == "tick":
            self.lbl_runtime.configure(text=fmt_duration(ev[1]))
            self.lbl_next.configure(text="-" if ev[2] is None else f"{math.ceil(ev[2])} s")
        elif kind == "cycle_done":
            self.files_done += ev[1]["files"]
            self.records_done += ev[1]["records"]
            self._update_stats()
            if ev[1]["files"]:
                self.lbl_error.configure(text="")
        elif kind == "history":
            self._insert_history(ev[1])
        elif kind == "error":
            self.lbl_error.configure(text="Last error: " + ev[1][:160])
        elif kind == "notice":
            self._update_buttons()
            messagebox.showwarning(APP_NAME, ev[1])
        elif kind == "log":
            self._append_log(ev[1])
        elif kind == "preview":
            self._show_preview(ev[1], ev[2], ev[3])
        elif kind == "preview_error":
            self.lbl_preview.configure(text="Preview unavailable: " + ev[1])

    # ================================================================== shutdown
    def _on_close(self) -> None:
        if self.engine.active and not messagebox.askyesno(APP_NAME, "Generation is still running. Stop and exit?"):
            return
        self._closing = True
        try:
            self._commit_editor()
            self._sync_store()
            self._save_settings()
        except Exception:
            pass
        self.engine.shutdown(5.0)
        try:
            self.after_cancel(self._poll_id)
        except Exception:
            pass
        self.destroy()
