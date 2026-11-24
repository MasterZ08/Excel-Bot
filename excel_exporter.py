"""XLSX export with OpenPyXL (write-only mode so 100,000+ rows stay fast and low-memory).

Workbook layout
  Data       - Excel table (banded rows, filters), bold coloured header, frozen header,
               auto column widths, date / currency / number formats
  Summary    - metadata + live formulas (COUNTA/SUM/AVERAGE/MIN/MAX) over the Data sheet
  Breakdown  - counts per value for low-cardinality text columns (or a column dictionary)
"""
from __future__ import annotations

import warnings

from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

from generators import Dataset, GenerationAborted

FONT = "Arial"
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(name=FONT, bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(name=FONT, bold=True, size=14, color="1F4E78")
BOLD = Font(name=FONT, bold=True, size=10)
NORMAL = Font(name=FONT, size=10)
SUB_FILL = PatternFill("solid", fgColor="DDEBF7")
NUMERIC = ("int", "decimal", "currency")


def number_format(col, sym: str = "$"):
    if col.kind == "int":
        return "#,##0"
    if col.kind == "decimal":
        return "#,##0" + ("." + "0" * col.decimals if col.decimals > 0 else "")
    if col.kind == "currency":
        return f'"{sym}"#,##0.' + "0" * max(1, col.decimals)
    if col.kind == "date":
        return "yyyy-mm-dd"
    return None


def _cell(ws, value, fmt=None, font=None, fill=None, align=None):
    c = WriteOnlyCell(ws, value=value)   # strings starting with "=" are formulas (used by Summary)
    if fmt:
        c.number_format = fmt
    if font:
        c.font = font
    if fill:
        c.fill = fill
    if align:
        c.alignment = align
    return c


def _text_cell(ws, value, **kw):
    """A cell whose text is never interpreted as a formula."""
    c = _cell(ws, value, **kw)
    c.data_type = "s"
    return c


def _widths(ds: Dataset, sym: str) -> list:
    n = len(ds.rows)
    step = max(1, n // 2000)
    out = []
    for j, col in enumerate(ds.cols):
        m = len(col.name) + 4          # room for the filter arrow
        for r in ds.rows[::step]:
            v = r[j]
            if col.kind == "date":
                ln = 11
            elif col.kind == "currency":
                ln = len(f"{v:,.2f}") + len(sym) if v is not None else 0
            elif isinstance(v, float):
                ln = len(f"{v:,.{max(col.decimals, 0)}f}")
            else:
                ln = len(str(v)) if v is not None else 0
            m = max(m, ln + 2)
        out.append(min(max(m, 10), 60))
    return out


def export_xlsx(ds: Dataset, path, settings: dict, progress=None, abort=None) -> None:
    sym = settings.get("currency_symbol", "$")
    ncols, nrows = len(ds.cols), len(ds.rows)
    wb = Workbook(write_only=True)
    ws = wb.create_sheet("Data")
    summary = wb.create_sheet("Summary") if settings.get("include_summary", True) else None
    breakdown = wb.create_sheet("Breakdown") if summary is not None else None

    ws.freeze_panes = "A2"
    for j, w in enumerate(_widths(ds, sym), 1):
        ws.column_dimensions[get_column_letter(j)].width = w

    center = Alignment(horizontal="center", vertical="center")
    ws.append([_text_cell(ws, h, font=HEADER_FONT, fill=HEADER_FILL, align=center) for h in ds.headers])

    fmts = [number_format(c, sym) for c in ds.cols]
    plain = [f is None for f in fmts]
    step = max(1, nrows // 100)
    for i, row in enumerate(ds.rows):
        if i % step == 0:
            if abort is not None and abort.is_set():
                raise GenerationAborted()
            if progress:
                progress(0.9 * i / max(1, nrows))
        out = []
        for v, fmt, is_plain in zip(row, fmts, plain):
            if v is None:
                out.append(None)
            elif is_plain:
                out.append(_text_cell(ws, v) if isinstance(v, str) and v.startswith("=") else v)
            else:
                c = WriteOnlyCell(ws, value=v)
                c.number_format = fmt
                out.append(c)
        ws.append(out)

    last_col = get_column_letter(ncols)
    tab = Table(displayName="DataTable", ref=f"A1:{last_col}{nrows + 1}")
    tab._initialise_columns()
    for tc, name in zip(tab.tableColumns, ds.headers):
        tc.name = name
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True,
                                        showFirstColumn=False, showLastColumn=False)
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        ws.add_table(tab)
    if nrows:   # banding that also shows in LibreOffice / Google Sheets (Excel uses the table style too)
        band = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
        ws.conditional_formatting.add(f"A2:{last_col}{nrows + 1}", FormulaRule(formula=["MOD(ROW(),2)=0"], fill=band))

    if summary is not None:
        _write_summary(summary, ds, sym, nrows)
        _write_breakdown(breakdown, ds, nrows)
    if progress:
        progress(0.95)
    if abort is not None and abort.is_set():
        raise GenerationAborted()
    wb.save(str(path))
    if progress:
        progress(1.0)


def _write_summary(ws, ds: Dataset, sym: str, nrows: int) -> None:
    ws.column_dimensions["A"].width = 26
    for col in "BCDEF":
        ws.column_dimensions[col].width = 20
    ws.append([_text_cell(ws, "Dataset Summary", font=TITLE_FONT)])
    ws.append([])
    last = nrows + 1
    meta = [
        ("Category", ds.category),
        ("Generated at", ds.generated_at.strftime("%Y-%m-%d %H:%M:%S")),
        ("Records", f"=COUNTA(Data!A2:A{last})"),
        ("Columns", len(ds.cols)),
        ("Random seed", ds.seed if ds.seed is not None else "Random (not fixed)"),
        ("Locale", ds.locale),
        ("Data source", ds.engine_name + " - all values are synthetic"),
    ]
    for k, v in meta:
        ws.append([_text_cell(ws, k, font=BOLD), _cell(ws, v, fmt="#,##0" if isinstance(v, int) or str(v).startswith("=") else None, font=NORMAL)])
    ws.append([])
    ws.append([_text_cell(ws, h, font=HEADER_FONT, fill=HEADER_FILL) for h in ("Column", "Type", "Sum", "Average", "Min", "Max")])
    for j, col in enumerate(ds.cols, 1):
        L = get_column_letter(j)
        rng = f"Data!{L}2:{L}{last}"
        fmt = number_format(col, sym)
        row = [_text_cell(ws, col.name, font=NORMAL), _text_cell(ws, col.kind, font=NORMAL)]
        if col.kind in NUMERIC:
            row += [_cell(ws, f"=SUM({rng})", fmt=fmt, font=NORMAL), _cell(ws, f"=AVERAGE({rng})", fmt=fmt, font=NORMAL),
                    _cell(ws, f"=MIN({rng})", fmt=fmt, font=NORMAL), _cell(ws, f"=MAX({rng})", fmt=fmt, font=NORMAL)]
        elif col.kind == "date":
            row += [None, None, _cell(ws, f"=MIN({rng})", fmt=fmt, font=NORMAL), _cell(ws, f"=MAX({rng})", fmt=fmt, font=NORMAL)]
        ws.append(row)


def _write_breakdown(ws, ds: Dataset, nrows: int) -> None:
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 16
    ws.append([_text_cell(ws, "Breakdown by category", font=TITLE_FONT)])
    ws.append([])
    shown = 0
    for j, col in enumerate(ds.cols):
        if col.kind not in ("text", "bool") or shown >= 3:
            continue
        counts: dict = {}
        ok = True
        for r in ds.rows:
            v = r[j]
            counts[v] = counts.get(v, 0) + 1
            if len(counts) > 25:
                ok = False
                break
        if not ok or len(counts) < 2 or len(counts) > nrows * 0.6 and nrows > 10:
            continue
        shown += 1
        ws.append([_text_cell(ws, col.name, font=HEADER_FONT, fill=HEADER_FILL),
                   _text_cell(ws, "Count", font=HEADER_FONT, fill=HEADER_FILL),
                   _text_cell(ws, "% of records", font=HEADER_FONT, fill=HEADER_FILL)])
        items = sorted(counts.items(), key=lambda kv: (-kv[1], str(kv[0])))
        for val, cnt in items:
            ws.append([_text_cell(ws, str(val), font=NORMAL), _cell(ws, cnt, fmt="#,##0", font=NORMAL),
                       _cell(ws, cnt / max(1, nrows), fmt="0.0%", font=NORMAL)])
        ws.append([])
    if shown == 0:
        ws.append([_text_cell(ws, "No low-cardinality text columns in this dataset.", font=NORMAL)])
