"""Regenerates the sample files in examples/ (fixed seed, 100 records each).

    python make_examples.py
"""
from pathlib import Path

import config_manager as cm
import csv_exporter as ce
import excel_exporter as xe
import generators as g

HERE = Path(__file__).resolve().parent


def main() -> None:
    cfg = cm.ConfigManager(HERE)
    s, cols = cfg.defaults()
    s.update(use_seed=True, seed=2024, records=100, date_end="2025-12-31")
    s = cm.normalize_settings(s)
    out = HERE / "examples"
    out.mkdir(exist_ok=True)
    gen = g.DataGenerator()
    for cat in cm.CATEGORIES:
        ds = gen.generate(cat, s, cols, records=100, cycle=0)
        slug = cat.lower().replace("/", "_").replace(" ", "_")
        xe.export_xlsx(ds, out / f"example_{slug}.xlsx", s)
        if cat in ("Employees", "Business/Sales", "Custom Dataset"):
            ce.export_csv(ds, out / f"example_{slug}.csv", s)
        print("wrote", slug)

    support = [
        {"name": "Ticket ID", "type": "id", "prefix": "TCK-", "id_mode": "Sequential", "width": 5},
        {"name": "Opened", "type": "date", "min": "2025-01-01", "max": "2025-12-31"},
        {"name": "Customer", "type": "name"}, {"name": "Contact Email", "type": "email"},
        {"name": "Priority", "type": "custom_list", "values": ["Low", "Medium", "High", "Urgent"]},
        {"name": "Status", "type": "custom_list", "values": ["Open", "In Progress", "Resolved", "Closed"]},
        {"name": "Hours Spent", "type": "decimal", "min": 0.25, "max": 16, "decimals": 2},
        {"name": "Billable Amount", "type": "currency", "min": 0, "max": 1500},
        {"name": "Escalated", "type": "boolean", "true_pct": 15},
    ]
    cs = dict(s, category="Custom Dataset", prefix="support_tickets", records=250, use_seed=False)
    cfg.save_dataset_config(HERE / "configs" / "example_support_tickets.json", cm.normalize_settings(cs),
                            [cm.normalize_column(c) for c in support])


if __name__ == "__main__":
    main()
