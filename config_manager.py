"""Settings, custom-column definitions, JSON persistence."""
from __future__ import annotations

import copy
import datetime as dt
import json
import os
import threading
from pathlib import Path

from utils import app_dir

APP_NAME = "Random Excel Data Generator"
CATEGORIES = [
    "Personal Information", "Employees", "Business/Sales", "Financial",
    "Inventory", "Education", "Custom Dataset",
]
FORMATS = ["XLSX", "CSV", "Both"]
LOCALES = ["en_US", "en_GB", "en_PH", "en_AU", "en_CA", "en_IN", "fr_FR", "de_DE",
           "es_ES", "es_MX", "it_IT", "pt_BR", "nl_NL", "ja_JP", "ko_KR", "zh_CN"]
APPEARANCES = ["Dark", "Light", "System"]

MIN_DELAY, MAX_DELAY = 1, 3600
MAX_RECORDS = 1_000_000          # Excel's hard limit is 1,048,575 data rows
CUSTOM_TYPES = ["text", "integer", "decimal", "name", "email", "phone", "date",
                "id", "boolean", "address", "currency", "custom_list"]
TYPE_LABELS = {
    "text": "Text", "integer": "Integer", "decimal": "Decimal", "name": "Name", "email": "Email",
    "phone": "Phone", "date": "Date", "id": "ID", "boolean": "Boolean", "address": "Address",
    "currency": "Currency", "custom_list": "Custom list",
}
ID_MODES = ["Random", "Sequential"]


class ConfigError(Exception):
    """Invalid user configuration (not retryable until the user fixes it)."""


DEFAULTS: dict = {
    "output_dir": "",                 # filled with <app>/output at load time
    "prefix": "random_data",
    "category": "Personal Information",
    "export_format": "XLSX",
    "records": 100,
    "delay": 5,
    "continuous": True,
    "id_mode": "Random",
    "id_prefix": "",
    "id_suffix": "",
    "use_seed": False,
    "seed": 12345,
    "allow_duplicates": False,
    "duplicate_pct": 5.0,
    "date_start": "2015-01-01",
    "date_end": "",                   # blank = today
    "num_min": "",                    # blank = category default
    "num_max": "",
    "age_min": 18,
    "age_max": 80,
    "locale": "en_US",
    "currency_symbol": "$",
    "include_summary": True,
    "appearance": "Dark",
}


# ------------------------------------------------------------------ coercion
def _num(v, default, lo, hi, as_float=False):
    try:
        x = float(str(v).replace(",", "").strip())
    except (TypeError, ValueError):
        return default
    x = max(lo, min(hi, x))
    return x if as_float else int(x)


def _bool(v, default):
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return bool(v)
    if isinstance(v, str):
        if v.strip().lower() in ("1", "true", "yes", "on"):
            return True
        if v.strip().lower() in ("0", "false", "no", "off", ""):
            return False
    return default


def _choice(v, options, default):
    return v if v in options else default


def opt_float(v):
    """'' / None -> None, otherwise float (raises ValueError if not numeric)."""
    if v is None or str(v).strip() == "":
        return None
    return float(str(v).replace(",", "").strip())


def parse_date(s, default=None):
    s = (s or "").strip() if isinstance(s, str) else s
    if not s:
        return default
    if isinstance(s, dt.date):
        return s
    try:
        return dt.datetime.strptime(s, "%Y-%m-%d").date()
    except ValueError:
        raise ConfigError(f"Invalid date '{s}' (use YYYY-MM-DD)")


def normalize_settings(raw: dict | None, base: dict | None = None) -> dict:
    """Return a complete, clamped settings dict. Bad values fall back to `base`/defaults."""
    d = copy.deepcopy(DEFAULTS)
    if base:
        d.update({k: base[k] for k in DEFAULTS if k in base})
    raw = raw or {}
    g = lambda k: raw.get(k, d[k])  # noqa: E731
    d["output_dir"] = str(g("output_dir") or d["output_dir"] or "")
    d["prefix"] = str(g("prefix") or "").strip() or d["prefix"] or "random_data"
    d["category"] = _choice(g("category"), CATEGORIES, d["category"])
    d["export_format"] = _choice(g("export_format"), FORMATS, d["export_format"])
    d["records"] = _num(g("records"), d["records"], 1, MAX_RECORDS)
    d["delay"] = _num(g("delay"), d["delay"], MIN_DELAY, MAX_DELAY, as_float=True)
    if float(d["delay"]).is_integer():
        d["delay"] = int(d["delay"])
    d["continuous"] = _bool(g("continuous"), d["continuous"])
    d["id_mode"] = _choice(g("id_mode"), ID_MODES, d["id_mode"])
    d["id_prefix"] = str(g("id_prefix") or "")[:20]
    d["id_suffix"] = str(g("id_suffix") or "")[:20]
    d["use_seed"] = _bool(g("use_seed"), d["use_seed"])
    d["seed"] = _num(g("seed"), d["seed"], 0, 2**31 - 1)
    d["allow_duplicates"] = _bool(g("allow_duplicates"), d["allow_duplicates"])
    d["duplicate_pct"] = _num(g("duplicate_pct"), d["duplicate_pct"], 0, 50, as_float=True)
    d["date_start"] = str(g("date_start") or "").strip()
    d["date_end"] = str(g("date_end") or "").strip()
    d["num_min"] = "" if str(g("num_min")).strip() == "" else str(g("num_min")).strip()
    d["num_max"] = "" if str(g("num_max")).strip() == "" else str(g("num_max")).strip()
    d["age_min"] = _num(g("age_min"), d["age_min"], 0, 120)
    d["age_max"] = _num(g("age_max"), d["age_max"], 0, 120)
    d["locale"] = _choice(g("locale"), LOCALES, d["locale"])
    d["currency_symbol"] = str(g("currency_symbol") if g("currency_symbol") is not None else "$")[:3].replace('"', "")
    d["include_summary"] = _bool(g("include_summary"), d["include_summary"])
    d["appearance"] = _choice(g("appearance"), APPEARANCES, d["appearance"])
    return d


def validate_settings(s: dict) -> list[str]:
    """Human-readable problems that would make generation impossible."""
    errs: list[str] = []
    try:
        d0 = parse_date(s["date_start"], dt.date(2015, 1, 1))
        d1 = parse_date(s["date_end"], dt.date.today())
        if d0 > d1:
            errs.append("Date range: start is after end.")
    except ConfigError as e:
        errs.append(str(e))
    try:
        lo, hi = opt_float(s["num_min"]), opt_float(s["num_max"])
        if lo is not None and hi is not None and lo > hi:
            errs.append("Number range: min is greater than max.")
    except ValueError:
        errs.append("Number range must be numeric (or blank for defaults).")
    if s["age_min"] > s["age_max"]:
        errs.append("Age range: min is greater than max.")
    if not str(s["output_dir"]).strip():
        errs.append("Choose an output folder.")
    return errs


# --------------------------------------------------------------- custom columns
def normalize_column(d: dict) -> dict:
    t = d.get("type", "text")
    if t not in CUSTOM_TYPES:
        t = "text"
    vals = d.get("values", [])
    if isinstance(vals, str):
        vals = [v.strip() for v in vals.replace("\r", "").replace(",", "\n").split("\n")]
    vals = [str(v) for v in vals if str(v).strip() != ""]
    return {
        "name": (str(d.get("name", "")).strip() or "Column")[:100],
        "type": t,
        "min": "" if d.get("min") is None else str(d.get("min")).strip(),
        "max": "" if d.get("max") is None else str(d.get("max")).strip(),
        "decimals": _num(d.get("decimals", 2), 2, 0, 8),
        "values": vals,
        "prefix": str(d.get("prefix", ""))[:20],
        "suffix": str(d.get("suffix", ""))[:20],
        "id_mode": _choice(d.get("id_mode"), ["Global", "Random", "Sequential"], "Global"),
        "width": _num(d.get("width", 6), 6, 1, 20),
        "true_pct": _num(d.get("true_pct", 50), 50.0, 0, 100, as_float=True),
    }


def column_range(c: dict):
    """Numeric/date range with defaults for a normalized column -> (lo, hi) raw strings parsed."""
    t = c["type"]
    if t in ("integer", "decimal", "currency", "text"):
        dlo, dhi = {"integer": (0, 1000), "decimal": (0, 1000), "currency": (1, 1000), "text": (1, 6)}[t]
        lo = float(c["min"]) if c["min"] != "" else dlo
        hi = float(c["max"]) if c["max"] != "" else dhi
        return lo, hi
    return None


def validate_columns(cols: list[dict]) -> list[str]:
    errs: list[str] = []
    if not cols:
        return ["Custom dataset has no columns."]
    seen = set()
    for i, raw in enumerate(cols, 1):
        c = normalize_column(raw)
        key = c["name"].casefold()
        if key in seen:
            errs.append(f"Duplicate column name '{c['name']}'.")
        seen.add(key)
        try:
            r = column_range(c)
            if r and r[0] > r[1]:
                errs.append(f"'{c['name']}': min is greater than max.")
        except ValueError:
            errs.append(f"'{c['name']}': min/max must be numeric.")
        if c["type"] == "date":
            try:
                a, b = parse_date(c["min"]), parse_date(c["max"])
                if a and b and a > b:
                    errs.append(f"'{c['name']}': start date is after end date.")
            except ConfigError as e:
                errs.append(f"'{c['name']}': {e}")
        if c["type"] == "custom_list" and not c["values"]:
            errs.append(f"'{c['name']}': custom list needs at least one value.")
    return errs


def default_custom_columns() -> list[dict]:
    return [normalize_column(c) for c in [
        {"name": "Record ID", "type": "id", "prefix": "REC-", "id_mode": "Sequential", "width": 6},
        {"name": "Full Name", "type": "name"},
        {"name": "Email", "type": "email"},
        {"name": "Score", "type": "integer", "min": 0, "max": 100},
        {"name": "Rating", "type": "decimal", "min": 1, "max": 5, "decimals": 1},
        {"name": "Signup Date", "type": "date"},
        {"name": "Plan", "type": "custom_list", "values": ["Free", "Basic", "Pro", "Enterprise"]},
        {"name": "Active", "type": "boolean", "true_pct": 70},
    ]]


# ------------------------------------------------------------------ thread-safe store
class SettingsStore:
    """GUI writes, worker reads snapshots. Never touches Tk widgets from the worker."""

    def __init__(self, settings: dict, columns: list[dict]):
        self._lock = threading.RLock()
        self._s = copy.deepcopy(settings)
        self._c = copy.deepcopy(columns)

    def set_settings(self, s: dict) -> None:
        with self._lock:
            self._s = copy.deepcopy(s)

    def set_columns(self, c: list[dict]) -> None:
        with self._lock:
            self._c = copy.deepcopy(c)

    def get(self, key):
        with self._lock:
            return self._s[key]

    def snapshot(self) -> tuple[dict, list[dict]]:
        with self._lock:
            return copy.deepcopy(self._s), copy.deepcopy(self._c)


# ------------------------------------------------------------------ persistence
def _atomic_write_json(path: Path, data) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(path.name + ".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    os.replace(tmp, path)


class ConfigManager:
    def __init__(self, base=None):
        self.base = Path(base) if base else app_dir()
        self.configs_dir = self.base / "configs"
        self.logs_dir = self.base / "logs"
        self.default_output = self.base / "output"
        self.settings_path = self.configs_dir / "settings.json"
        self.history_path = self.configs_dir / "history.jsonl"
        self.state_path = self.configs_dir / "state.json"

    def ensure_dirs(self) -> None:
        for p in (self.configs_dir, self.logs_dir, self.default_output):
            p.mkdir(parents=True, exist_ok=True)

    def defaults(self) -> tuple[dict, list[dict]]:
        s = normalize_settings({"output_dir": str(self.default_output)})
        s["output_dir"] = str(self.default_output)
        return s, default_custom_columns()

    def load_settings(self) -> tuple[dict, list[dict]]:
        s, cols = self.defaults()
        try:
            with open(self.settings_path, encoding="utf-8") as f:
                data = json.load(f)
            s = normalize_settings(data.get("settings", {}), base=s)
            if not s["output_dir"]:
                s["output_dir"] = str(self.default_output)
            loaded = [normalize_column(c) for c in data.get("custom_columns", [])]
            if loaded:
                cols = loaded
        except FileNotFoundError:
            pass
        except (OSError, ValueError, AttributeError):
            pass  # corrupt file -> defaults
        return s, cols

    def save_settings(self, settings: dict, columns: list[dict]) -> None:
        _atomic_write_json(self.settings_path, {"settings": settings, "custom_columns": columns})

    def save_dataset_config(self, path, settings: dict, columns: list[dict]) -> None:
        _atomic_write_json(Path(path), {
            "format_version": 1, "app": APP_NAME,
            "settings": settings, "custom_columns": columns,
        })

    def load_dataset_config(self, path, current: dict) -> tuple[dict, list[dict]]:
        """Load a config file. Output folder and theme stay as currently set."""
        try:
            with open(path, encoding="utf-8") as f:
                data = json.load(f)
        except (OSError, ValueError) as e:
            raise ValueError(f"Cannot read config: {e}")
        if not isinstance(data, dict) or "settings" not in data and "custom_columns" not in data:
            raise ValueError("Not a valid generator config file.")
        incoming = dict(data.get("settings", {}))
        incoming.pop("output_dir", None)
        incoming.pop("appearance", None)
        s = normalize_settings(incoming, base=current)
        cols = [normalize_column(c) for c in data.get("custom_columns", [])]
        return s, cols

    # sequential-ID counters survive restarts so IDs stay unique across runs
    def load_state(self) -> dict:
        try:
            with open(self.state_path, encoding="utf-8") as f:
                d = json.load(f)
            return {str(k): int(v) for k, v in d.get("sequences", {}).items()}
        except (OSError, ValueError, AttributeError):
            return {}

    def save_state(self, sequences: dict) -> None:
        _atomic_write_json(self.state_path, {"sequences": sequences})
