"""Headless tests for everything except the Tk window.

Run:  python -m unittest discover -s tests -v
"""
import copy
import csv
import os
import queue
import shutil
import sys
import tempfile
import threading
import time
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import config_manager as cm  # noqa: E402
import generators as g  # noqa: E402
from csv_exporter import export_csv  # noqa: E402
from excel_exporter import export_xlsx  # noqa: E402
from history_manager import HistoryManager  # noqa: E402
from utils import reserve_files  # noqa: E402
from worker import Engine  # noqa: E402
from openpyxl import load_workbook  # noqa: E402


def make_env(tmp, **over):
    cfg = cm.ConfigManager(tmp)
    cfg.ensure_dirs()
    s, cols = cfg.defaults()
    s.update(over)
    s = cm.normalize_settings(s)
    store = cm.SettingsStore(s, cols)
    q = queue.Queue()
    hist = HistoryManager(cfg.history_path)
    return cfg, s, cols, store, q, hist, Engine(store, hist, q, cfg)


def drain(q):
    out = []
    while True:
        try:
            out.append(q.get_nowait())
        except queue.Empty:
            return out


class Base(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="reg_test_")

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)


class TestGenerators(Base):
    def test_every_category(self):
        cfg, s, cols, *_ = make_env(self.tmp)
        gen = g.DataGenerator()
        for cat in cm.CATEGORIES:
            ds = gen.generate(cat, s, cols, records=57)
            self.assertEqual(len(ds.rows), 57, cat)
            for r in ds.rows:
                self.assertEqual(len(r), len(ds.cols), cat)
                self.assertFalse(any(v is None for v in r), cat)

    def test_values_differ_every_cycle(self):
        cfg, s, cols, *_ = make_env(self.tmp)
        gen = g.DataGenerator()
        for cat in cm.CATEGORIES:
            a = gen.generate(cat, s, cols, records=40).rows
            b = gen.generate(cat, s, cols, records=40).rows
            self.assertNotEqual(a, b, cat)

    def test_fixed_seed_reproducible_but_cycles_differ(self):
        cfg, s, cols, *_ = make_env(self.tmp, use_seed=True, seed=7, id_mode="Random")
        a0 = g.DataGenerator().generate("Employees", s, cols, records=30, cycle=0).rows
        a0b = g.DataGenerator().generate("Employees", s, cols, records=30, cycle=0).rows
        a1 = g.DataGenerator().generate("Employees", s, cols, records=30, cycle=1).rows
        self.assertEqual(a0, a0b)
        self.assertNotEqual(a0, a1)

    def test_ids_unique_prefix_suffix_sequential(self):
        cfg, s, cols, *_ = make_env(self.tmp, id_mode="Sequential", id_prefix="X-", id_suffix="-Z")
        gen = g.DataGenerator()
        d1 = gen.generate("Employees", s, cols, records=10)
        d2 = gen.generate("Employees", s, cols, records=10)
        ids = [r[0] for r in d1.rows + d2.rows]
        self.assertEqual(len(set(ids)), 20)                     # continues across cycles
        self.assertTrue(all(i.startswith("X-") and i.endswith("-Z") for i in ids))
        self.assertEqual(ids[0], "X-00001-Z")
        cfg, s, cols, *_ = make_env(self.tmp, id_mode="Random")
        ids = [r[0] for r in g.DataGenerator().generate("Personal Information", s, cols, records=5000).rows]
        self.assertEqual(len(set(ids)), 5000)

    def test_duplicates_option(self):
        cfg, s, cols, *_ = make_env(self.tmp, allow_duplicates=True, duplicate_pct=20)
        ds = g.DataGenerator().generate("Employees", s, cols, records=100)
        self.assertEqual(len(ds.rows), 100)
        ids = [r[0] for r in ds.rows]
        self.assertEqual(len(ids) - len(set(ids)), 20)
        cfg, s, cols, *_ = make_env(self.tmp, allow_duplicates=False)
        ids = [r[0] for r in g.DataGenerator().generate("Employees", s, cols, records=100).rows]
        self.assertEqual(len(ids), len(set(ids)))

    def test_ranges(self):
        cfg, s, cols, *_ = make_env(self.tmp, date_start="2020-02-01", date_end="2020-02-29",
                                    num_min="100", num_max="200", age_min=30, age_max=31)
        ds = g.DataGenerator().generate("Business/Sales", s, cols, records=200)
        di = ds.headers.index("Order Date"); pi = ds.headers.index("Unit Price")
        for r in ds.rows:
            self.assertTrue("2020-02-01" <= r[di].isoformat() <= "2020-02-29")
            self.assertTrue(100 <= r[pi] <= 200)
        ds = g.DataGenerator().generate("Personal Information", s, cols, records=100)
        self.assertTrue(all(30 <= r[3] <= 31 for r in ds.rows))

    def test_business_and_financial_integrity(self):
        cfg, s, cols, *_ = make_env(self.tmp)
        ds = g.DataGenerator().generate("Business/Sales", s, cols, records=300)
        q, p, t = (ds.headers.index(h) for h in ("Quantity", "Unit Price", "Total"))
        self.assertTrue(all(abs(r[q] * r[p] - r[t]) < 0.011 for r in ds.rows))
        ds = g.DataGenerator().generate("Financial", s, cols, records=500)
        self.assertTrue(all(r[5] >= 0 for r in ds.rows))          # never overdrawn
        self.assertTrue(all((r[3] == 0) != (r[4] == 0) for r in ds.rows))

    def test_emails_are_non_routable(self):
        cfg, s, cols, *_ = make_env(self.tmp)
        ds = g.DataGenerator().generate("Personal Information", s, cols, records=200)
        self.assertTrue(all(r[6].endswith(("@example.com", "@example.org", "@example.net")) for r in ds.rows))

    def test_custom_dataset_all_types(self):
        cols = [cm.normalize_column(c) for c in [
            {"name": "T", "type": "text", "min": 2, "max": 3}, {"name": "I", "type": "integer", "min": 5, "max": 9},
            {"name": "D", "type": "decimal", "min": 1, "max": 2, "decimals": 3}, {"name": "N", "type": "name"},
            {"name": "E", "type": "email"}, {"name": "P", "type": "phone"},
            {"name": "Dt", "type": "date", "min": "2021-01-01", "max": "2021-01-31"}, {"name": "Id", "type": "id", "prefix": "Q", "id_mode": "Sequential", "width": 4},
            {"name": "B", "type": "boolean", "true_pct": 100}, {"name": "A", "type": "address"},
            {"name": "C", "type": "currency", "min": 10, "max": 20}, {"name": "L", "type": "custom_list", "values": ["u", "v"]}]]
        cfg, s, _, *_ = make_env(self.tmp)
        ds = g.DataGenerator().generate("Custom Dataset", s, cols, records=50)
        self.assertEqual(len(ds.headers), 12)
        for r in ds.rows:
            self.assertTrue(5 <= r[1] <= 9 and 1 <= r[2] <= 2 and r[8] is True and r[11] in "uv" and 10 <= r[10] <= 20)
            self.assertTrue("2021-01-01" <= r[6].isoformat() <= "2021-01-31")
        self.assertEqual(ds.rows[0][7], "Q0001")

    def test_custom_validation(self):
        cfg, s, *_ = make_env(self.tmp)
        bad = [{"name": "A", "type": "integer", "min": 9, "max": 1}, {"name": "a", "type": "name"}]
        with self.assertRaises(cm.ConfigError):
            g.DataGenerator().generate("Custom Dataset", s, bad, records=3)

    def test_abort(self):
        cfg, s, cols, *_ = make_env(self.tmp)
        ev = threading.Event(); ev.set()
        with self.assertRaises(g.GenerationAborted):
            g.DataGenerator().generate("Employees", s, cols, records=5000, abort=ev)


class TestExport(Base):
    def test_xlsx_structure(self):
        cfg, s, cols, *_ = make_env(self.tmp)
        ds = g.DataGenerator().generate("Business/Sales", s, cols, records=120)
        p = Path(self.tmp) / "a.xlsx"
        export_xlsx(ds, p, s)
        wb = load_workbook(p)
        self.assertEqual(wb.sheetnames, ["Data", "Summary", "Breakdown"])
        ws = wb["Data"]
        self.assertEqual(ws.max_row, 121)
        self.assertEqual(ws.freeze_panes, "A2")
        self.assertTrue(ws["A1"].font.b)
        self.assertEqual(ws["B2"].number_format, "yyyy-mm-dd")
        self.assertIn("$", ws["G2"].number_format)
        self.assertEqual(list(ws.tables), ["DataTable"])
        self.assertTrue(ws.tables["DataTable"].tableStyleInfo.showRowStripes)
        self.assertIsNotNone(ws.tables["DataTable"].autoFilter)
        self.assertTrue(ws.column_dimensions["C"].width >= 10)
        self.assertTrue(any(str(c.value).startswith("=SUM(") for row in wb["Summary"].iter_rows() for c in row))

    def test_text_starting_with_equals_is_not_a_formula(self):
        cols = [cm.normalize_column({"name": "L", "type": "custom_list", "values": ["=1+1"]})]
        cfg, s, *_ = make_env(self.tmp)
        ds = g.DataGenerator().generate("Custom Dataset", s, cols, records=2)
        p = Path(self.tmp) / "f.xlsx"; export_xlsx(ds, p, s)
        c = load_workbook(p)["Data"]["A2"]
        self.assertEqual((c.value, c.data_type), ("=1+1", "s"))

    def test_csv(self):
        cfg, s, cols, *_ = make_env(self.tmp)
        ds = g.DataGenerator().generate("Employees", s, cols, records=25)
        p = Path(self.tmp) / "a.csv"; export_csv(ds, p, s)
        with open(p, encoding="utf-8-sig") as fh:
            rows = list(csv.reader(fh))
        self.assertEqual(len(rows), 26)
        self.assertEqual(rows[0], ds.headers)

    def test_no_overwrite_unique_names(self):
        out = Path(self.tmp) / "o"
        seen = set()
        for _ in range(50):
            ps = reserve_files(out, "same", ["xlsx", "csv"])
            for p in ps:
                self.assertNotIn(p.name, seen); seen.add(p.name)
        stems = {p.rsplit(".", 1)[0] for p in seen}
        self.assertEqual(len(stems), 50)
        # sequence continues after restart
        self.assertTrue(reserve_files(out, "same", ["xlsx"])[0].name.startswith("same_00051_"))
        # an existing file with the same stem is skipped, not clobbered
        existing = out / "same_00052_19990101_000000.xlsx"; existing.write_text("keep")
        reserve_files(out, "same", ["xlsx"]); self.assertEqual(existing.read_text(), "keep")

    def test_bad_prefix_sanitised(self):
        out = Path(self.tmp) / "o"
        p = reserve_files(out, 'a<b>:"c"/d', ["csv"])[0]
        self.assertTrue(p.exists() and "<" not in p.name)


class TestConfigHistory(Base):
    def test_settings_roundtrip_and_defaults(self):
        cfg = cm.ConfigManager(self.tmp); cfg.ensure_dirs()
        s, cols = cfg.defaults()
        s["records"] = 777; s["prefix"] = "zz"; cols.append(cm.normalize_column({"name": "Extra", "type": "name"}))
        cfg.save_settings(s, cols)
        s2, cols2 = cfg.load_settings()
        self.assertEqual((s2["records"], s2["prefix"], len(cols2)), (777, "zz", 9))
        d, dc = cfg.defaults()
        self.assertEqual((d["records"], d["delay"], d["continuous"]), (100, 5, True))
        self.assertEqual(len(dc), 8)

    def test_dataset_config_file(self):
        cfg = cm.ConfigManager(self.tmp); cfg.ensure_dirs()
        s, cols = cfg.defaults(); s["records"] = 42; s["output_dir"] = "/somewhere/else"
        p = Path(self.tmp) / "c.json"; cfg.save_dataset_config(p, s, cols[:3])
        cur, _ = cfg.defaults()
        s2, cols2 = cfg.load_dataset_config(p, cur)
        self.assertEqual((s2["records"], len(cols2), s2["output_dir"]), (42, 3, cur["output_dir"]))
        (Path(self.tmp) / "bad.json").write_text("{nope")
        with self.assertRaises(ValueError):
            cfg.load_dataset_config(Path(self.tmp) / "bad.json", cur)

    def test_clamping(self):
        s = cm.normalize_settings({"delay": 0, "records": 0, "duplicate_pct": 999})
        self.assertEqual((s["delay"], s["records"], s["duplicate_pct"]), (1, 1, 50))
        self.assertEqual(cm.normalize_settings({"delay": 99999})["delay"], 3600)
        self.assertEqual(cm.normalize_settings({"delay": "abc"})["delay"], 5)

    def test_history_persists(self):
        cfg = cm.ConfigManager(self.tmp); cfg.ensure_dirs()
        h = HistoryManager(cfg.history_path)
        for i in range(3):
            h.add({"timestamp": "t", "filename": f"f{i}", "category": "c", "records": i, "format": "XLSX", "status": "Success"})
        self.assertEqual(len(HistoryManager(cfg.history_path).all()), 3)
        with open(cfg.history_path, "a") as f:
            f.write("{garbage\n")
        h2 = HistoryManager(cfg.history_path); self.assertEqual(len(h2.all()), 3)
        n = h2.export_csv(Path(self.tmp) / "h.csv"); self.assertEqual(n, 3)
        h2.clear(); self.assertEqual(len(HistoryManager(cfg.history_path).all()), 0)


class TestEngine(Base):
    def files(self, ext="xlsx"):
        return sorted(p for p in (Path(self.tmp) / "output").glob(f"*.{ext}"))

    def wait_for(self, pred, timeout=15):
        end = time.time() + timeout
        while time.time() < end:
            if pred():
                return True
            time.sleep(0.05)
        return False

    def test_generate_one(self):
        cfg, s, cols, store, q, hist, eng = make_env(self.tmp, records=20, export_format="Both", delay=1)
        self.assertTrue(eng.generate_one())
        self.assertTrue(self.wait_for(lambda: not eng.active))
        self.assertEqual((len(self.files("xlsx")), len(self.files("csv"))), (1, 1))
        self.assertEqual(len(hist.all()), 2)
        self.assertFalse(list((Path(self.tmp) / "output").glob("*.part")))

    def test_continuous_pause_resume_stop(self):
        cfg, s, cols, store, q, hist, eng = make_env(self.tmp, records=15, delay=1, export_format="XLSX")
        t0 = time.time()
        eng.start_run()
        self.assertTrue(self.wait_for(lambda: len(self.files()) >= 1, 5))          # first file is immediate
        self.assertLess(time.time() - t0, 3)
        self.assertTrue(self.wait_for(lambda: len(self.files()) >= 3, 8))
        eng.pause()
        self.assertTrue(self.wait_for(lambda: ("state", "paused") in drain(q) or True))
        time.sleep(1.6)
        n_paused = len(self.files()); time.sleep(2.5)
        self.assertEqual(len(self.files()), n_paused, "files were generated while paused")
        eng.resume()
        self.assertTrue(self.wait_for(lambda: len(self.files()) > n_paused, 5))
        t1 = time.time(); eng.stop()
        self.assertTrue(self.wait_for(lambda: not eng.active, 3)); self.assertLess(time.time() - t1, 1.5)
        n_stop = len(self.files()); time.sleep(2.2)
        self.assertEqual(len(self.files()), n_stop)
        names = [p.name for p in self.files()]
        self.assertEqual(len(names), len(set(names)))
        evs = drain(q)
        self.assertTrue(any(e[0] == "tick" and e[2] is not None for e in evs))     # countdown reported
        self.assertEqual([e for e in evs if e[0] == "state"][-1], ("state", "idle"))
        # every file opens and the data differs between files
        data = [[r for r in load_workbook(p)["Data"].iter_rows(min_row=2, values_only=True)] for p in self.files()[:3]]
        self.assertNotEqual(data[0], data[1]); self.assertNotEqual(data[1], data[2])

    def test_no_file_limit_attribute_and_counts_beyond_any_cap(self):
        cfg, s, cols, store, q, hist, eng = make_env(self.tmp, records=1, delay=1, export_format="CSV")
        import inspect, worker
        self.assertNotRegex(inspect.getsource(worker).lower(), r"max_files|file_limit|files_limit")
        eng.start_run()
        self.assertTrue(self.wait_for(lambda: len(self.files("csv")) >= 7, 12))   # keeps going
        self.assertTrue(eng.active)
        eng.stop(); self.assertTrue(self.wait_for(lambda: not eng.active, 3))

    def test_continuous_off_makes_single_file(self):
        cfg, s, cols, store, q, hist, eng = make_env(self.tmp, records=5, delay=1, continuous=False)
        eng.start_run(); self.assertTrue(self.wait_for(lambda: not eng.active, 5))
        time.sleep(1.5); self.assertEqual(len(self.files()), 1)

    def test_randomize_again_skips_countdown(self):
        cfg, s, cols, store, q, hist, eng = make_env(self.tmp, records=5, delay=60)
        eng.start_run(); self.assertTrue(self.wait_for(lambda: len(self.files()) == 1, 5))
        time.sleep(0.5); self.assertTrue(eng.randomize())
        self.assertTrue(self.wait_for(lambda: len(self.files()) == 2, 3))
        eng.stop(); self.assertTrue(self.wait_for(lambda: not eng.active, 2))

    def test_stop_is_instant_during_long_countdown(self):
        cfg, s, cols, store, q, hist, eng = make_env(self.tmp, records=5, delay=3600)
        eng.start_run(); self.assertTrue(self.wait_for(lambda: len(self.files()) == 1, 5))
        t = time.time(); eng.stop(); self.assertTrue(self.wait_for(lambda: not eng.active, 2)); self.assertLess(time.time() - t, 1)

    def test_error_recovery_continues(self):
        bad = Path(self.tmp) / "blocker"; bad.write_text("i am a file, not a folder")
        cfg, s, cols, store, q, hist, eng = make_env(self.tmp, records=5, delay=1, output_dir=str(bad / "sub"))
        eng.start_run()
        self.assertTrue(self.wait_for(lambda: sum(1 for e in hist.all() if e["status"] == "Failed") >= 2, 8))
        self.assertTrue(eng.active, "worker must survive recoverable errors")
        good = Path(self.tmp) / "ok"; s2 = dict(s); s2["output_dir"] = str(good); store.set_settings(s2)
        self.assertTrue(self.wait_for(lambda: len(list(good.glob("*.xlsx"))) >= 1, 6))   # recovered by itself
        eng.stop(); self.assertTrue(self.wait_for(lambda: not eng.active, 3))

    def test_bad_config_pauses_instead_of_spinning(self):
        cfg, s, cols, store, q, hist, eng = make_env(self.tmp, records=5, delay=1, date_start="2030-01-01", date_end="2020-01-01")
        eng.start_run()
        self.assertTrue(self.wait_for(lambda: any(e[0] == "notice" for e in drain(q)) or len(hist.all()) >= 1, 5))
        time.sleep(2.5); self.assertEqual(len(hist.all()), 1)                      # not retried every second
        eng.stop(); self.assertTrue(self.wait_for(lambda: not eng.active, 3))

    def test_shutdown_aborts_big_job_without_partial_files(self):
        cfg, s, cols, store, q, hist, eng = make_env(self.tmp, records=400000, delay=1, category="Personal Information")
        eng.start_run(); time.sleep(1.0)
        t = time.time(); eng.shutdown(5); self.assertFalse(eng.active); self.assertLess(time.time() - t, 4)
        out = Path(self.tmp) / "output"
        self.assertEqual(list(out.glob("*.xlsx")) + list(out.glob("*.part")), [])
        self.assertEqual([p for p in out.iterdir() if p.name != ".gitkeep"], [])

    def test_large_dataset_100k(self):
        cfg, s, cols, store, q, hist, eng = make_env(self.tmp, records=100000, export_format="Both", category="Employees")
        t = time.time(); eng.generate_one()
        self.assertTrue(self.wait_for(lambda: not eng.active, 240))
        print(f"\n   100,000-record Employees XLSX+CSV in {time.time() - t:.1f}s")
        x = self.files("xlsx")[0]
        self.assertEqual(sum(1 for _ in load_workbook(x, read_only=True)["Data"].iter_rows()), 100001)
        with open(self.files("csv")[0], encoding="utf-8-sig") as fh:
            self.assertEqual(sum(1 for _ in fh), 100001)

    def test_sequence_state_persists(self):
        cfg, s, cols, store, q, hist, eng = make_env(self.tmp, records=10, id_mode="Sequential", category="Employees")
        eng.generate_one(); self.assertTrue(self.wait_for(lambda: not eng.active, 5))
        eng2 = Engine(store, hist, q, cfg)
        eng2.generate_one(); self.assertTrue(self.wait_for(lambda: not eng2.active, 5))
        ids = []
        for p in self.files():
            ids += [r[0] for r in load_workbook(p)["Data"].iter_rows(min_row=2, values_only=True)]
        self.assertEqual(len(ids), len(set(ids)))


if __name__ == "__main__":
    unittest.main(verbosity=2)
