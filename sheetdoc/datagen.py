"""Category 4 — data generation and analysis.

build_mock_dataset, explain_pivot_table_setup, chart_type_recommender,
plus the practice-workbook writer used by tutor.generate_practice_variant.
"""

from __future__ import annotations

import datetime as dt
import random
import statistics
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .report import Finding, Report
from .workbook import Book, Sheet

HEADER_FILL = PatternFill("solid", start_color="1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFF")
TITLE_FONT = Font(bold=True, size=13)
THIN = Side(style="thin", color="B0B0B0")

FIRST = [
    "Amara", "Ben", "Chloe", "Dmitri", "Elena", "Farid", "Grace", "Hugo", "Ines",
    "Jonas", "Kira", "Liam", "Mira", "Noah", "Omar", "Priya", "Quinn", "Rosa",
    "Sven", "Tara", "Ugo", "Vera", "Wes", "Xenia", "Yusuf", "Zoe", "Anja", "Bo",
]
LAST = [
    "Adler", "Bauer", "Castro", "Dunn", "Ergin", "Fischer", "Gomez", "Haas",
    "Iqbal", "Jensen", "Klein", "Lopez", "Meier", "Novak", "Okafor", "Petrov",
    "Rahman", "Schmidt", "Tanaka", "Ueda", "Vogel", "Weber", "Yilmaz", "Zhang",
]
REGIONS = ["North", "South", "East", "West"]
PRODUCTS = ["Laptop", "Monitor", "Keyboard", "Headset", "Docking Station", "Webcam"]
CHANNELS = ["Online", "Retail", "Partner"]
DEPTS = ["Sales", "Engineering", "Marketing", "Support", "Finance"]
CATEGORIES = ["Rent", "Utilities", "Salaries", "Marketing", "Travel", "Supplies", "Software"]
COURSES = ["Accounting", "Statistics", "Marketing", "Economics", "Management"]


# ==========================================================================
# 4.1 build_mock_dataset
# ==========================================================================

DATASETS = {
    "sales": "Order-level sales: date, region, channel, product, units, unit price, revenue",
    "payroll": "Employee payroll: name, department, hire date, hours, rate, gross pay",
    "grades": "Student gradebook: name, course, three assessments, attendance",
    "inventory": "Stock levels: SKU, product, warehouse, on hand, reorder point, unit cost",
    "expenses": "Monthly expenses: date, category, vendor, amount, budgeted amount",
    "budget": "Budget vs actual by category and month",
}


def build_mock_dataset(
    kind: str,
    out_path: str,
    rows: int = 20,
    seed: int | None = None,
    sheet_name: str | None = None,
    formatted: bool = True,
    append_to: str | None = None,
) -> Report:
    """Write a realistic practice dataset to a worksheet.

    Values are plausible and internally consistent (revenue really is units ×
    price), because a dataset with nonsense relationships teaches nothing when
    you then chart or pivot it.
    """
    kind = kind.lower().strip()
    rep = Report(title=f"Mock dataset — {kind}")
    if kind not in DATASETS:
        rep.add(Finding("error", "UNKNOWN_DATASET", None, None,
                        f"Unknown dataset {kind!r}",
                        "Available: " + ", ".join(f"{k} ({v})" for k, v in DATASETS.items()),
                        "Pick one of the datasets above."))
        return rep

    rng = random.Random(seed if seed is not None else random.randrange(1 << 30))

    if append_to:
        wb = openpyxl.load_workbook(append_to)
        target = Path(append_to)
    else:
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        target = Path(out_path).expanduser()

    title = sheet_name or kind.capitalize()
    if title in wb.sheetnames:
        del wb[title]
    ws = wb.create_sheet(title)

    headers, records, formats = _generate(kind, rows, rng)
    _write_table(ws, headers, records, formats, formatted=formatted)

    wb.save(target)
    rep.summary = {
        "file": str(target),
        "sheet": title,
        "rows": len(records),
        "columns": ", ".join(headers),
        "data range": f"A2:{get_column_letter(len(headers))}{len(records) + 1}",
    }
    rep.add(Finding("done", "DATASET_WRITTEN", None, title,
                    f"Wrote {len(records)} rows of {kind} data to {target.name}",
                    DATASETS[kind],
                    "Open it in Excel and try the exercises on it."))
    return rep


def _generate(kind: str, rows: int, rng: random.Random):
    """-> (headers, records, number_formats)"""
    if kind == "sales":
        headers = ["Order ID", "Order Date", "Region", "Channel", "Product",
                   "Units", "Unit Price", "Revenue"]
        formats = ["General", "mm/dd/yyyy", "General", "General", "General",
                   "#,##0", '$#,##0.00', '$#,##0.00']
        base = dt.date(2026, 1, 1)
        prices = {p: round(rng.uniform(29, 1450), 2) for p in PRODUCTS}
        recs = []
        for i in range(rows):
            product = rng.choice(PRODUCTS)
            units = rng.randint(1, 25)
            price = round(prices[product] * rng.uniform(0.95, 1.05), 2)
            recs.append([
                f"ORD-{1000 + i}",
                base + dt.timedelta(days=rng.randint(0, 200)),
                rng.choice(REGIONS),
                rng.choice(CHANNELS),
                product,
                units,
                price,
                round(units * price, 2),
            ])
        recs.sort(key=lambda r: r[1])
        return headers, recs, formats

    if kind == "payroll":
        headers = ["Employee ID", "Name", "Department", "Hire Date",
                   "Hours Worked", "Hourly Rate", "Gross Pay"]
        formats = ["General", "General", "General", "mm/dd/yyyy",
                   "#,##0.0", '$#,##0.00', '$#,##0.00']
        recs = []
        for i in range(rows):
            hours = round(rng.choice([35, 38, 40, 40, 40, 42, 45, 48]) + rng.random(), 1)
            rate = round(rng.uniform(16.5, 62.0), 2)
            recs.append([
                f"E{2001 + i}",
                f"{rng.choice(FIRST)} {rng.choice(LAST)}",
                rng.choice(DEPTS),
                dt.date(rng.randint(2018, 2025), rng.randint(1, 12), rng.randint(1, 28)),
                hours,
                rate,
                round(hours * rate, 2),
            ])
        return headers, recs, formats

    if kind == "grades":
        headers = ["Student ID", "Name", "Course", "Midterm", "Project",
                   "Final Exam", "Attendance %"]
        formats = ["General", "General", "General", "#,##0", "#,##0", "#,##0", "0%"]
        recs = []
        for i in range(rows):
            ability = rng.gauss(74, 12)
            recs.append([
                f"S{3001 + i}",
                f"{rng.choice(FIRST)} {rng.choice(LAST)}",
                rng.choice(COURSES),
                max(0, min(100, round(rng.gauss(ability, 8)))),
                max(0, min(100, round(rng.gauss(ability + 5, 7)))),
                max(0, min(100, round(rng.gauss(ability, 9)))),
                round(min(1.0, max(0.55, rng.gauss(0.9, 0.09))), 2),
            ])
        return headers, recs, formats

    if kind == "inventory":
        headers = ["SKU", "Product", "Warehouse", "On Hand", "Reorder Point",
                   "Unit Cost", "Stock Value"]
        formats = ["General", "General", "General", "#,##0", "#,##0",
                   '$#,##0.00', '$#,##0.00']
        recs = []
        for i in range(rows):
            on_hand = rng.randint(0, 480)
            cost = round(rng.uniform(4.5, 320), 2)
            recs.append([
                f"SKU-{4100 + i * 7}",
                rng.choice(PRODUCTS),
                rng.choice(["Mainz", "Berlin", "Hamburg"]),
                on_hand,
                rng.choice([25, 50, 75, 100]),
                cost,
                round(on_hand * cost, 2),
            ])
        return headers, recs, formats

    if kind == "expenses":
        headers = ["Date", "Category", "Vendor", "Amount", "Budgeted", "Variance"]
        formats = ["mm/dd/yyyy", "General", "General", '$#,##0.00', '$#,##0.00', '$#,##0.00']
        recs = []
        for i in range(rows):
            budgeted = round(rng.uniform(400, 9000), 2)
            amount = round(budgeted * rng.uniform(0.72, 1.34), 2)
            recs.append([
                dt.date(2026, rng.randint(1, 8), rng.randint(1, 28)),
                rng.choice(CATEGORIES),
                f"{rng.choice(LAST)} {rng.choice(['GmbH', 'Ltd', 'AG', 'Co'])}",
                amount,
                budgeted,
                round(budgeted - amount, 2),
            ])
        recs.sort(key=lambda r: r[0])
        return headers, recs, formats

    # budget
    headers = ["Month", "Category", "Budget", "Actual", "Variance", "Variance %"]
    formats = ["General", "General", '$#,##0', '$#,##0', '$#,##0', "0.0%"]
    months = ["January", "February", "March", "April", "May", "June"]
    recs = []
    for i in range(rows):
        budget = round(rng.uniform(2000, 25000), -2)
        actual = round(budget * rng.uniform(0.75, 1.28), -1)
        recs.append([
            months[i % len(months)],
            CATEGORIES[i % len(CATEGORIES)],
            budget,
            actual,
            budget - actual,
            round((budget - actual) / budget, 4),
        ])
    return headers, recs, formats


def _write_table(ws, headers, records, formats, formatted=True, start_row=1):
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=start_row, column=c, value=h)
        if formatted:
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = Border(bottom=THIN)
    for r, rec in enumerate(records, start_row + 1):
        for c, v in enumerate(rec, 1):
            cell = ws.cell(row=r, column=c, value=v)
            if formatted and c <= len(formats):
                cell.number_format = formats[c - 1]
    if formatted:
        for c, h in enumerate(headers, 1):
            letter = get_column_letter(c)
            longest = max([len(str(h))] + [len(str(rec[c - 1])) for rec in records[:60]])
            ws.column_dimensions[letter].width = min(max(longest + 4, 11), 22)
        ws.freeze_panes = ws.cell(row=start_row + 1, column=1)
        ws.auto_filter.ref = (
            f"A{start_row}:{get_column_letter(len(headers))}{start_row + len(records)}"
        )
    return start_row + len(records)


# ==========================================================================
# 4.2 explain_pivot_table_setup
# ==========================================================================


def explain_pivot_table_setup(
    book: Book, sheet_name: str | None = None, goal: str | None = None
) -> Report:
    """Read a real sheet and map its columns onto Rows / Columns / Values / Filters.

    The classification is mechanical: low-cardinality text is a dimension, dates
    are a time dimension, numerics that are not IDs are measures. That is exactly
    the judgement the PivotTable field list expects you to make.
    """
    rep = Report(title="Pivot table setup")
    sh = book.sheet(sheet_name)
    region = sh.main_region()
    if not region:
        rep.add(Finding("error", "NO_DATA", None, sh.name,
                        "No data block found on this sheet", "",
                        "A pivot table needs a rectangular table with one header row."))
        return rep

    hr = sh.header_row(region)
    if hr is None:
        rep.add(Finding("error", "NO_HEADER", region.a1, sh.name,
                        "No header row detected",
                        "Every pivot source needs one row of unique text labels at the top.",
                        "Add short column labels above your data."))
        return rep

    rows_f, cols_f, values_f, filters_f, skipped = [], [], [], [], []
    profiles = []
    for c in range(region.min_col, region.max_col + 1):
        p = sh.column_profile(c, region)
        if not p.header:
            skipped.append(get_column_letter(c))
            continue
        profiles.append(p)
        n = len(p.values) or 1
        distinct = p.distinct

        if p.kind == "date":
            cols_f.append((p, "date — group by Month/Quarter/Year after dropping it in"))
        elif p.kind == "number":
            if _is_identifier(p):
                filters_f.append((p, "looks like an identifier, not something to add up"))
            else:
                values_f.append((p, _agg_for(p)))
        elif p.kind in ("text", "mixed"):
            if distinct <= 1:
                skipped.append(p.letter)
            elif distinct <= 8:
                rows_f.append((p, f"{distinct} distinct values — compact enough for rows or columns"))
            elif distinct <= max(20, n * 0.3):
                rows_f.append((p, f"{distinct} distinct values — good for rows"))
            else:
                filters_f.append((p, f"{distinct} distinct values — too many to display, use as a filter"))

    rep.summary = {
        "source table": f"{sh.name}!{region.a1}",
        "header row": hr,
        "records": region.max_row - hr,
        "goal": goal or "(none given — recommendations are generic)",
    }

    if goal:
        _bias_toward_goal(goal, rows_f, cols_f, values_f, filters_f)

    def emit(area: str, items, note: str):
        if not items:
            rep.add(Finding("info", f"PIVOT_{area.upper()}_EMPTY", None, sh.name,
                            f"{area}: nothing obvious", note))
            return
        for p, why in items:
            rep.add(Finding("info", f"PIVOT_{area.upper()}", p.letter, sh.name,
                            f"{area} ▸ {p.header}", why))

    emit("Rows", rows_f[:3], "No low-cardinality text column found.")
    emit("Columns", cols_f[:2],
         "Optional. Leave Columns empty unless you want a cross-tab.")
    emit("Values", values_f[:4], "No numeric measure found — a pivot needs something to total.")
    emit("Filters", filters_f[:3], "Optional.")

    steps = [
        f"Click any cell inside {sh.name}!{region.a1}.",
        "Insert ▸ PivotTable ▸ New Worksheet ▸ OK.",
        f"Drag {rows_f[0][0].header if rows_f else '<a category>'} into the Rows box.",
        f"Drag {values_f[0][0].header if values_f else '<a number>'} into the Values box.",
        (f"Values defaults to Count for text and Sum for numbers. Click the field ▸ "
         f"Value Field Settings to change it to {values_f[0][1] if values_f else 'Sum'}."),
        "Value Field Settings ▸ Number Format sets the pivot's own formatting — the source "
        "sheet's formatting does not carry over.",
    ]
    if cols_f:
        steps.insert(4, f"Drag {cols_f[0][0].header} into Columns to break the totals out sideways.")
    for i, s in enumerate(steps, 1):
        rep.add(Finding("info", "PIVOT_STEP", None, sh.name, f"Step {i}", s))

    rep.notes.append(
        "Rows = what you group BY. Values = what you add UP. If a field ends up in Values "
        "showing 'Count of Region', you dropped a text field where a number belongs."
    )
    if skipped:
        rep.notes.append(f"Ignored columns: {', '.join(skipped)} (blank or single-valued).")
    return rep


def _is_identifier(p) -> bool:
    """Distinguish an ID column from a measure.

    All-distinct values alone prove nothing — revenue on 200 orders is all
    distinct too, and it is exactly what you want to total. An ID is either
    named like one, or a whole-number sequence that never repeats.
    """
    from .workbook import ID_WORDS

    if ID_WORDS.search(p.header):
        return True
    vals = p.numeric_values
    if len(vals) < 3 or p.cardinality_ratio < 0.99:
        return False
    if any(v != int(v) for v in vals):
        return False  # fractional values are measures, not identifiers
    ordered = sorted(int(v) for v in vals)
    span = ordered[-1] - ordered[0] + 1
    return span <= len(ordered) * 1.5  # a dense integer sequence


def _agg_for(p) -> str:
    h = p.header.lower()
    if any(w in h for w in ("rate", "%", "percent", "average", "avg", "score", "grade")):
        return "Average (rates and scores do not add up meaningfully)"
    if any(w in h for w in ("price", "cost per", "unit")):
        return "Average (a unit price should not be summed across orders)"
    return "Sum"


def _bias_toward_goal(goal: str, rows_f, cols_f, values_f, filters_f) -> None:
    g = goal.lower()
    def bump(bucket):
        for i, (p, why) in enumerate(bucket):
            if p.header.lower() in g or any(w in g for w in p.header.lower().split()):
                bucket.insert(0, bucket.pop(i))
                return
    for bucket in (rows_f, cols_f, values_f, filters_f):
        bump(bucket)


# ==========================================================================
# 4.3 chart_type_recommender
# ==========================================================================


def chart_type_recommender(
    book: Book, sheet_name: str | None = None, cell_range: str | None = None
) -> Report:
    """Recommend a chart type from the actual shape of the data.

    The decision rests on three questions, in order: is the x-axis time, are the
    two variables both continuous, and do the parts sum to a whole?
    """
    from .formula import parse_range

    rep = Report(title="Chart type recommendation")
    sh = book.sheet(sheet_name)
    region = sh.main_region()

    if cell_range:
        rng = parse_range(cell_range)
        if not rng:
            rep.add(Finding("error", "BAD_RANGE", None, sh.name,
                            f"Could not parse {cell_range!r}", "", "Use A1 notation like A1:D21."))
            return rep
        min_row, max_row = rng.min_row, rng.max_row
        min_col, max_col = rng.min_col, rng.max_col
    elif region:
        min_row, max_row = region.min_row, region.max_row
        min_col, max_col = region.min_col, region.max_col
    else:
        rep.add(Finding("error", "NO_DATA", None, sh.name, "No data found on this sheet", "", ""))
        return rep

    from .workbook import Region

    sub = Region(min_row, max_row, min_col, max_col)
    hr = sh.header_row(sub)
    profiles = [sh.column_profile(c, sub) for c in range(min_col, max_col + 1)]
    profiles = [p for p in profiles if p.kind != "empty"]
    if not profiles:
        rep.add(Finding("error", "NO_DATA", None, sh.name, "Nothing to chart in that range", "", ""))
        return rep

    dates = [p for p in profiles if p.kind == "date"]
    numbers = [p for p in profiles if p.kind == "number"]
    texts = [p for p in profiles if p.kind in ("text", "mixed")]
    n_records = max((len(p.values) for p in profiles), default=0)

    rep.summary = {
        "range": f"{sh.name}!{get_column_letter(min_col)}{min_row}:"
                 f"{get_column_letter(max_col)}{max_row}",
        "records": n_records,
        "date columns": ", ".join(p.header or p.letter for p in dates) or "none",
        "category columns": ", ".join(p.header or p.letter for p in texts) or "none",
        "numeric columns": ", ".join(p.header or p.letter for p in numbers) or "none",
    }

    recs: list[tuple[int, str, str, str]] = []  # (rank score, chart, why, how)

    # -- time series ------------------------------------------------------
    if dates and numbers:
        recs.append((
            100, "Line chart",
            f"'{dates[0].header}' is a date column and you have {len(numbers)} numeric "
            f"measure(s). Time on the x-axis with a continuous measure is the definition of "
            "a line chart — the line implies 'these points are connected in sequence', which "
            "is true for time and false for categories.",
            f"Select the date column and the measure ▸ Insert ▸ Line ▸ Line with Markers. "
            f"Sort by {dates[0].header} first or the line will zig-zag.",
        ))
        if len(numbers) > 1:
            recs.append((
                70, "Combo chart (columns + line)",
                "With two measures on different scales — say revenue in thousands and a "
                "percentage — one axis crushes the other flat.",
                "Insert ▸ Combo ▸ put the smaller-scale series on the Secondary Axis.",
            ))

    # -- category comparison ----------------------------------------------
    if texts and numbers:
        cat = min(texts, key=lambda p: p.distinct)
        k = cat.distinct
        if k <= 12:
            recs.append((
                95 if not dates else 60, "Clustered column chart",
                f"'{cat.header}' has {k} distinct categories and you are comparing a numeric "
                "measure across them. Columns are read by comparing heights, which is the most "
                "accurate visual comparison available.",
                f"Select '{cat.header}' and the measure ▸ Insert ▸ Column ▸ Clustered Column. "
                "Sort descending by value first — an unsorted bar chart is much harder to read.",
            ))
        else:
            recs.append((
                90, "Bar chart (horizontal)",
                f"'{cat.header}' has {k} categories. Horizontal bars give long labels room to "
                "breathe; vertical columns would force the labels to tilt.",
                "Insert ▸ Bar ▸ Clustered Bar, sorted by value.",
            ))

        # parts of a whole
        if k <= 6 and numbers:
            vals = numbers[0].numeric_values
            if vals and all(v >= 0 for v in vals):
                recs.append((
                    35, "Pie chart — only if these are parts of one whole",
                    f"You have {k} categories, which is few enough for a pie. But a pie is only "
                    "honest when the slices sum to a meaningful total (share of revenue, yes; "
                    "average score by department, no). Humans compare angles badly, so a column "
                    "chart is almost always the better answer.",
                    "If the assignment explicitly asks for a pie: Insert ▸ Pie ▸ 2-D Pie, then "
                    "add data labels showing percentages.",
                ))

    # -- correlation ------------------------------------------------------
    if len(numbers) >= 2:
        x, y = numbers[0], numbers[1]
        r = _correlation(x.numeric_values, y.numeric_values)
        strength = (
            "strong" if r is not None and abs(r) > 0.7
            else "moderate" if r is not None and abs(r) > 0.4
            else "weak"
        )
        score = 85 if (r is not None and abs(r) > 0.4 and not dates and not texts) else 50
        recs.append((
            score, "Scatter plot (XY)",
            f"'{x.header}' and '{y.header}' are both continuous numbers"
            + (f", with a {strength} correlation (r = {r:+.2f})" if r is not None else "")
            + ". A scatter plot is the only chart that puts a number on both axes — use it when "
            "you are asking whether two measures move together, not when you are comparing "
            "categories.",
            f"Select the two numeric columns ▸ Insert ▸ Scatter ▸ Scatter with only Markers. "
            "Add a trendline: click a point ▸ Add Trendline ▸ tick 'Display R-squared'.",
        ))

    # -- distribution ------------------------------------------------------
    if numbers and n_records >= 15:
        p = numbers[0]
        vals = p.numeric_values
        if len(vals) >= 15:
            recs.append((
                45, "Histogram",
                f"'{p.header}' has {len(vals)} values"
                + (f" spanning {min(vals):,.0f} to {max(vals):,.0f}" if vals else "")
                + ". If the question is 'how are these distributed' rather than 'how do "
                "categories compare', you want a histogram.",
                "Insert ▸ Statistic Chart ▸ Histogram, then right-click the x-axis ▸ Format "
                "Axis to set the bin width.",
            ))

    if not recs:
        rep.add(Finding("warning", "NO_RECOMMENDATION", None, sh.name,
                        "Not enough structure to recommend a chart",
                        "A chart needs at least one label column and one numeric column.",
                        "Select a range that includes headers, categories, and values."))
        return rep

    recs.sort(key=lambda t: -t[0])
    for i, (score, chart, why, how) in enumerate(recs):
        label = "BEST CHOICE" if i == 0 else f"Alternative {i}"
        rep.add(Finding("info", "CHART_REC", None, sh.name,
                        f"{label}: {chart}", why, how))

    rep.notes.append(
        "Whatever you pick: title the chart with the finding, not the field names "
        "('Revenue grew 34% in Q2' beats 'Revenue by Month'), label both axes with units, "
        "and start a bar-chart value axis at zero."
    )
    return rep


def _correlation(a: list[float], b: list[float]) -> float | None:
    n = min(len(a), len(b))
    if n < 3:
        return None
    a, b = a[:n], b[:n]
    try:
        if statistics.pstdev(a) == 0 or statistics.pstdev(b) == 0:
            return None
        return statistics.correlation(a, b)
    except Exception:
        return None


# ==========================================================================
# Practice workbook writer (used by tutor.generate_practice_variant)
# ==========================================================================


def write_practice_workbook(
    topic: str, out_path: str, rows: int, rng: random.Random, difficulty: str = "same"
) -> Report:
    """Build a 3-sheet practice file: Instructions, Data, Answer Key (hidden)."""
    from .tutor import VARIANTS

    meta = VARIANTS[topic]
    path = Path(out_path).expanduser()
    wb = openpyxl.Workbook()
    inst = wb.active
    inst.title = "Instructions"
    data = wb.create_sheet("Data")
    key = wb.create_sheet("Answer Key")

    tasks, answers, notes = _build_topic(topic, data, rows, rng, difficulty)

    # -- Instructions sheet --------------------------------------------------
    inst["A1"] = f"Practice: {meta['skill']}"
    inst["A1"].font = TITLE_FONT
    inst["A2"] = str(meta["why"])
    inst["A2"].font = Font(italic=True, color="595959")
    inst["A4"] = "Tasks"
    inst["A4"].font = Font(bold=True)
    r = 5
    for i, t in enumerate(tasks, 1):
        inst.cell(row=r, column=1, value=f"{i}.").font = Font(bold=True)
        inst.cell(row=r, column=2, value=t).alignment = Alignment(wrap_text=True, vertical="top")
        r += 1
    r += 1
    inst.cell(row=r, column=1, value="Watch out for").font = Font(bold=True)
    r += 1
    for n in notes:
        inst.cell(row=r, column=2, value="• " + n).alignment = Alignment(wrap_text=True)
        r += 1
    r += 1
    inst.cell(row=r, column=2,
              value="The 'Answer Key' sheet is hidden. Try the tasks first, then right-click "
                    "any sheet tab ▸ Unhide to check yourself.").font = Font(italic=True)
    inst.column_dimensions["A"].width = 5
    inst.column_dimensions["B"].width = 95

    # -- Answer key ----------------------------------------------------------
    key["A1"] = "Answer key — try the tasks before reading"
    key["A1"].font = TITLE_FONT
    key["A3"] = "Task"
    key["B3"] = "Formula"
    key["C3"] = "Why"
    for c in ("A3", "B3", "C3"):
        key[c].font = HEADER_FONT
        key[c].fill = HEADER_FILL
    for i, (formula, why) in enumerate(answers, 1):
        key.cell(row=3 + i, column=1, value=i)
        # Written as text so opening the key does not evaluate it in place.
        key.cell(row=3 + i, column=2, value=formula).font = Font(name="Menlo")
        key.cell(row=3 + i, column=3, value=why).alignment = Alignment(wrap_text=True)
    key.column_dimensions["A"].width = 6
    key.column_dimensions["B"].width = 46
    key.column_dimensions["C"].width = 62
    key.sheet_state = "hidden"

    wb.save(path)

    rep = Report(title=f"Practice variant — {meta['skill']}")
    rep.summary = {
        "file": str(path),
        "skill": str(meta["skill"]),
        "tests": str(meta["why"]),
        "rows of data": rows,
        "tasks": len(tasks),
    }
    for i, t in enumerate(tasks, 1):
        rep.add(Finding("info", "TASK", None, "Instructions", f"Task {i}", t))
    for n in notes:
        rep.notes.append(n)
    rep.notes.append("Answer key is on a hidden sheet — right-click a tab ▸ Unhide to reveal.")
    return rep


def _build_topic(topic, ws, rows, rng, difficulty):
    """Write the data sheet for a topic; return (tasks, answers, notes)."""
    if topic in ("vlookup", "xlookup", "index_match"):
        return _build_lookup(topic, ws, rows, rng)
    if topic in ("absolute", "percentage"):
        return _build_absolute(topic, ws, rows, rng)
    if topic == "payroll":
        return _build_payroll(ws, rows, rng)
    if topic == "sumif":
        return _build_sumif(ws, rows, rng)
    if topic == "if_nested":
        return _build_if(ws, rows, rng)
    if topic == "gradebook":
        return _build_gradebook(ws, rows, rng)
    return _build_budget(ws, rows, rng)


def _build_lookup(topic, ws, rows, rng):
    ws["A1"] = "Orders"
    ws["A1"].font = TITLE_FONT
    headers = ["Order ID", "Product Code", "Units", "Unit Price", "Line Total"]
    codes = [f"P-{100 + i}" for i in range(6)]
    names = ["Laptop", "Monitor", "Keyboard", "Headset", "Dock", "Webcam"]
    prices = [round(rng.uniform(35, 1400), 2) for _ in codes]

    recs = [[f"ORD-{2000 + i}", rng.choice(codes), rng.randint(1, 20), None, None]
            for i in range(rows)]
    _write_table(ws, headers, recs, ["General"] * 3 + ['$#,##0.00'] * 2, start_row=2)

    lut_col = 7  # column G
    ws.cell(row=2, column=lut_col, value="Product Code").font = HEADER_FONT
    ws.cell(row=2, column=lut_col).fill = HEADER_FILL
    ws.cell(row=2, column=lut_col + 1, value="Product Name").font = HEADER_FONT
    ws.cell(row=2, column=lut_col + 1).fill = HEADER_FILL
    ws.cell(row=2, column=lut_col + 2, value="Unit Price").font = HEADER_FONT
    ws.cell(row=2, column=lut_col + 2).fill = HEADER_FILL
    for i, (code, name, price) in enumerate(zip(codes, names, prices)):
        ws.cell(row=3 + i, column=lut_col, value=code)
        ws.cell(row=3 + i, column=lut_col + 1, value=name)
        ws.cell(row=3 + i, column=lut_col + 2, value=price).number_format = '$#,##0.00'
    for c in range(lut_col, lut_col + 3):
        ws.column_dimensions[get_column_letter(c)].width = 15

    last = 2 + rows
    lut = f"$G$3:$I${2 + len(codes)}"

    if topic == "vlookup":
        tasks = [
            f"In D3, use VLOOKUP to pull each order's Unit Price from the product table in G2:I{2 + len(codes)}. Fill down to D{last}.",
            f"In E3, calculate Line Total (Units × Unit Price). Fill down to E{last}.",
            "Add a column F called 'Product Name' and fill it from the same lookup table.",
            f"In E{last + 2}, total the Line Total column.",
            "Type an order with a product code that does not exist in the table. What does your formula return, and is that the behaviour you want?",
        ]
        answers = [
            (f"=VLOOKUP(B3,{lut},3,FALSE)",
             "Locked table, column 3 counted from the left edge of the table, FALSE for exact match."),
            ("=C3*D3", "Both references relative — each row uses its own values."),
            (f"=VLOOKUP(B3,{lut},2,FALSE)", "Same lookup, column 2 instead of 3."),
            (f"=SUM(E3:E{last})", "Range covers data rows only, not the header."),
            (f'=IFERROR(VLOOKUP(B3,{lut},3,FALSE),"Unknown code")',
             "#N/A is correct behaviour; IFERROR replaces it with a readable message."),
        ]
        notes = [
            "The lookup table must be locked with $ or it slides as you fill down.",
            "col_index_num counts from the left edge of table_array, not from column A.",
            "Leaving off the 4th argument silently switches to approximate match.",
        ]
    elif topic == "xlookup":
        tasks = [
            f"In D3, use XLOOKUP to pull Unit Price from the product table. Fill down to D{last}.",
            "Give your XLOOKUP a friendly message for codes that are not found.",
            f"In E3, calculate Line Total. Fill down to E{last}.",
            "Add column F with the Product Name, using XLOOKUP again.",
            "Deliberately make lookup_array one row longer than return_array. What error appears, and why?",
        ]
        answers = [
            (f"=XLOOKUP(B3,$G$3:$G${2 + len(codes)},$I$3:$I${2 + len(codes)})",
             "Both vectors span the same rows and are locked."),
            (f'=XLOOKUP(B3,$G$3:$G${2 + len(codes)},$I$3:$I${2 + len(codes)},"Not found")',
             "The 4th argument is XLOOKUP's built-in IFERROR."),
            ("=C3*D3", "Plain relative arithmetic."),
            (f"=XLOOKUP(B3,$G$3:$G${2 + len(codes)},$H$3:$H${2 + len(codes)},\"Not found\")",
             "Same lookup vector, different return vector."),
            ("(mismatched vectors)", "#VALUE! — XLOOKUP requires both vectors to be the same size."),
        ]
        notes = [
            "lookup_array and return_array must cover exactly the same rows.",
            "XLOOKUP defaults to exact match, unlike VLOOKUP.",
            "The 4th argument replaces #N/A without needing IFERROR.",
        ]
    else:
        tasks = [
            f"In D3, use INDEX + MATCH to pull Unit Price from the product table. Fill down to D{last}.",
            "Explain in a comment which range MATCH searches and which range INDEX returns from.",
            f"In E3, calculate Line Total. Fill down to E{last}.",
            "Change the INDEX range to start one row lower than the MATCH range. Check three answers — are they still right?",
            "Rebuild the same lookup so it returns the Product Code when given a Product Name (a leftward lookup).",
        ]
        answers = [
            (f"=INDEX($I$3:$I${2 + len(codes)},MATCH(B3,$G$3:$G${2 + len(codes)},0))",
             "Both ranges span rows 3 to " + str(2 + len(codes)) + "; the 0 forces exact match."),
            ("(explanation)",
             "MATCH searches G and returns a position number; INDEX counts that many rows into I."),
            ("=C3*D3", "Relative arithmetic."),
            ("(misaligned ranges)",
             "Every answer shifts by one row — the classic silent INDEX/MATCH bug."),
            (f"=INDEX($G$3:$G${2 + len(codes)},MATCH(<name>,$H$3:$H${2 + len(codes)},0))",
             "INDEX/MATCH can look left; VLOOKUP cannot."),
        ]
        notes = [
            "MATCH returns a position, not a value — the ranges must line up row for row.",
            "The 0 on the end of MATCH is what forces an exact match.",
        ]
    return tasks, answers, notes


def _build_absolute(topic, ws, rows, rng):
    ws["A1"] = "Assumptions"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = "Tax rate"
    ws["B2"] = round(rng.uniform(0.05, 0.22), 3)
    ws["B2"].number_format = "0.0%"
    ws["A3"] = "Discount rate"
    ws["B3"] = round(rng.uniform(0.03, 0.15), 3)
    ws["B3"].number_format = "0.0%"

    headers = ["Product", "List Price", "Discounted Price", "Tax", "Final Price", "% of Total"]
    recs = [[rng.choice(PRODUCTS) + f" {i + 1}", round(rng.uniform(25, 900), 2), None, None, None, None]
            for i in range(rows)]
    _write_table(ws, headers, recs, ["General"] + ['$#,##0.00'] * 4 + ["0.0%"], start_row=5)
    last = 5 + rows
    tasks = [
        f"In C6, calculate the discounted price (List Price × (1 − discount rate)). Fill down to C{last}.",
        f"In D6, calculate the tax on the discounted price. Fill down to D{last}.",
        f"In E6, calculate the final price. Fill down to E{last}.",
        f"In E{last + 1}, total the final prices.",
        f"In F6, calculate each product's share of that total. Fill down to F{last} and format as a percentage.",
        "Change B2 to a different rate. Every figure in columns D, E and F should update. If any does not, you hardcoded a number.",
    ]
    answers = [
        ("=B6*(1-$B$3)", "$B$3 is locked so every row uses the same discount rate."),
        ("=C6*$B$2", "Locked tax rate, relative price."),
        ("=C6+D6", "Both relative."),
        (f"=SUM(E6:E{last})", "Data rows only."),
        (f"=E6/$E${last + 1}", "The total must be locked; the numerator moves with each row."),
        ("(test)", "A live model recalculates from its assumption cells."),
    ]
    notes = [
        "The rate cells in B2 and B3 must be referenced absolutely ($B$2, $B$3).",
        "In the share-of-total formula, only the denominator is locked.",
        "Press ⌘T (F4 on Windows) with the cursor inside a reference to add the $ signs.",
    ]
    return tasks, answers, notes


def _build_payroll(ws, rows, rng):
    ws["A1"] = "Payroll assumptions"
    ws["A1"].font = TITLE_FONT
    ws["A2"], ws["B2"] = "Overtime multiplier", 1.5
    ws["A3"], ws["B3"] = "Standard hours", 40
    ws["A4"], ws["B4"] = "Tax rate", round(rng.uniform(0.18, 0.32), 3)
    ws["B4"].number_format = "0.0%"

    headers = ["Employee", "Department", "Hours", "Rate", "Regular Pay",
               "Overtime Pay", "Gross Pay", "Tax", "Net Pay"]
    recs = [[f"{rng.choice(FIRST)} {rng.choice(LAST)}", rng.choice(DEPTS),
             rng.choice([36, 38, 40, 40, 43, 46, 52]) + round(rng.random(), 1),
             round(rng.uniform(15, 55), 2), None, None, None, None, None]
            for _ in range(rows)]
    _write_table(ws, headers, recs,
                 ["General", "General", "#,##0.0"] + ['$#,##0.00'] * 6, start_row=6)
    last = 6 + rows
    tasks = [
        f"In E7, calculate Regular Pay: hours up to the standard, times the rate. Fill down to E{last}.",
        f"In F7, calculate Overtime Pay: hours above the standard, at the overtime multiplier.",
        f"In G7, calculate Gross Pay. In H7, the tax. In I7, Net Pay. Fill all three down.",
        f"In I{last + 1}, total the net pay.",
        "Format every money column as Accounting with 2 decimals.",
        "Add conditional formatting so any Gross Pay above 2,500 is highlighted green.",
    ]
    answers = [
        ("=MIN(C7,$B$3)*D7", "MIN caps regular hours at the standard; $B$3 locked."),
        ("=MAX(0,C7-$B$3)*D7*$B$2", "MAX(0,…) stops negative overtime for part-time staff."),
        ("=E7+F7 / =G7*$B$4 / =G7-H7", "Locked tax rate, everything else relative."),
        (f"=SUM(I7:I{last})", "Data rows only."),
        ("(formatting)", "Home ▸ Number ▸ Accounting."),
        ("(conditional formatting)", "Home ▸ Conditional Formatting ▸ Highlight Cells ▸ Greater Than."),
    ]
    notes = [
        "MIN and MAX are the clean way to split regular from overtime hours — no IF needed.",
        "All three assumption cells must be locked with $.",
    ]
    return tasks, answers, notes


def _build_sumif(ws, rows, rng):
    ws["A1"] = "Sales log"
    ws["A1"].font = TITLE_FONT
    headers = ["Date", "Region", "Rep", "Product", "Amount"]
    recs = [[dt.date(2026, rng.randint(1, 8), rng.randint(1, 28)), rng.choice(REGIONS),
             rng.choice(FIRST), rng.choice(PRODUCTS), round(rng.uniform(120, 8500), 2)]
            for _ in range(rows)]
    recs.sort(key=lambda r: r[0])
    _write_table(ws, headers, recs, ["mm/dd/yyyy", "General", "General", "General", '$#,##0.00'],
                 start_row=2)
    last = 2 + rows
    ws["G2"] = "Summary"
    ws["G2"].font = TITLE_FONT
    ws["G3"] = "Region"
    ws["H3"] = "Total Sales"
    ws["I3"] = "Order Count"
    ws["J3"] = "Average Order"
    for c in ("G3", "H3", "I3", "J3"):
        ws[c].font = HEADER_FONT
        ws[c].fill = HEADER_FILL
    for i, r in enumerate(REGIONS):
        ws.cell(row=4 + i, column=7, value=r)
    ws["G9"] = "Threshold"
    ws["H9"] = 2000
    ws["H9"].number_format = '$#,##0'
    ws["G10"] = "Large orders"
    for c in "GHIJ":
        ws.column_dimensions[c].width = 15

    tasks = [
        f"In H4, use SUMIF to total sales for the region named in G4. Fill down to H7.",
        "In I4, count the orders for that region. In J4, find the average order value.",
        "In H10, count how many orders exceed the threshold in H9 — reference the cell, do not type the number into the formula.",
        "Add a second summary that totals sales by region AND product, using SUMIFS.",
        "Change H9 to a different number. Your large-order count should update by itself.",
    ]
    answers = [
        (f"=SUMIF($B$3:$B${last},G4,$E$3:$E${last})",
         "Both ranges locked; the criteria cell G4 moves with the fill."),
        (f"=COUNTIF($B$3:$B${last},G4) and =AVERAGEIF($B$3:$B${last},G4,$E$3:$E${last})",
         "Same pattern for count and average."),
        (f'=COUNTIF($E$3:$E${last},">"&$H$9)',
         "The & joins the operator to the cell value, so the threshold stays visible in H9."),
        (f"=SUMIFS($E$3:$E${last},$B$3:$B${last},<region>,$D$3:$D${last},<product>)",
         "SUMIFS puts the sum range FIRST — the opposite of SUMIF."),
        ("(test)", "A referenced threshold makes the model live."),
    ]
    notes = [
        'Writing ">2000" as a literal string buries the threshold. Use ">"&$H$9 instead.',
        "SUMIF takes (range, criteria, sum_range); SUMIFS takes (sum_range, range1, criteria1, …). The order flips.",
        "Lock the data ranges so they do not slide as you fill the summary down.",
    ]
    return tasks, answers, notes


def _build_if(ws, rows, rng):
    ws["A1"] = "Grade scale"
    ws["A1"].font = TITLE_FONT
    scale = [("A", 90), ("B", 80), ("C", 70), ("D", 60), ("F", 0)]
    ws["A2"], ws["B2"] = "Grade", "Minimum score"
    ws["A2"].font = Font(bold=True)
    ws["B2"].font = Font(bold=True)
    for i, (g, m) in enumerate(scale):
        ws.cell(row=3 + i, column=1, value=g)
        ws.cell(row=3 + i, column=2, value=m)

    headers = ["Student", "Score", "Letter Grade", "Pass/Fail", "Honours"]
    recs = [[f"{rng.choice(FIRST)} {rng.choice(LAST)}",
             max(0, min(100, round(rng.gauss(74, 14)))), None, None, None]
            for _ in range(rows)]
    _write_table(ws, headers, recs, ["General", "#,##0", "General", "General", "General"],
                 start_row=9)
    last = 9 + rows
    tasks = [
        f"In C10, use nested IF to convert the score to a letter grade using the scale above. Fill down to C{last}.",
        f"In D10, return \"Pass\" for 60 or above and \"Fail\" below it.",
        "In E10, return \"Honours\" only for students scoring 90 or above AND with no score below 85 — use IF with AND.",
        "Rebuild task 1 using IFS instead of nested IF. Which version is easier to read?",
        "Test the exact boundary scores: 90, 89, 80, 79, 60, 59. Does every one land in the right band?",
    ]
    answers = [
        ('=IF(B10>=90,"A",IF(B10>=80,"B",IF(B10>=70,"C",IF(B10>=60,"D","F"))))',
         "Test from the highest band down. Reversing the order makes every score an F."),
        ('=IF(B10>=60,"Pass","Fail")', "Single condition."),
        ('=IF(AND(B10>=90,B10>=85),"Honours","")', "AND requires every condition to be true."),
        ('=IFS(B10>=90,"A",B10>=80,"B",B10>=70,"C",B10>=60,"D",TRUE,"F")',
         "IFS avoids the nesting. The final TRUE is the catch-all."),
        ("(boundary test)", "Boundaries are where >= vs > actually matters."),
    ]
    notes = [
        "Order matters: check the highest threshold first, or every score matches the first test.",
        ">= and > differ by exactly one boundary case — the score that equals the cut-off.",
        'Text inside a formula needs "quotes" or you get #NAME?.',
        "Bonus: this whole task is what VLOOKUP with approximate match (TRUE) is designed for.",
    ]
    return tasks, answers, notes


def _build_gradebook(ws, rows, rng):
    ws["A1"] = "Weights"
    ws["A1"].font = TITLE_FONT
    weights = [("Homework", 0.20), ("Midterm", 0.30), ("Project", 0.20), ("Final", 0.30)]
    for i, (name, w) in enumerate(weights):
        ws.cell(row=2, column=2 + i, value=name).font = Font(bold=True)
        c = ws.cell(row=3, column=2 + i, value=w)
        c.number_format = "0%"
    ws["A3"] = "Weight"
    ws["A3"].font = Font(bold=True)

    headers = ["Student", "Homework", "Midterm", "Project", "Final", "Weighted Average", "Grade"]
    recs = []
    for _ in range(rows):
        ability = rng.gauss(75, 11)
        recs.append([f"{rng.choice(FIRST)} {rng.choice(LAST)}"]
                    + [max(0, min(100, round(rng.gauss(ability, 8)))) for _ in range(4)]
                    + [None, None])
    _write_table(ws, headers, recs, ["General"] + ["#,##0"] * 4 + ["0.0", "General"], start_row=5)
    last = 5 + rows
    tasks = [
        f"In F6, calculate the weighted average using the weights in row 3. Fill down to F{last}.",
        "Do it a second way using SUMPRODUCT and compare the two answers.",
        f"In G6, convert the weighted average to a letter grade.",
        f"In F{last + 2}, find the class average, highest and lowest weighted scores.",
        "Highlight every weighted average below 60 in light red using conditional formatting.",
        "Change one weight in row 3 and confirm the whole column recalculates.",
    ]
    answers = [
        ("=B6*$B$3+C6*$C$3+D6*$D$3+E6*$E$3",
         "Each weight locked to row 3 with $ on the ROW only, so B$3 also works when dragged sideways."),
        ("=SUMPRODUCT($B$3:$E$3,B6:E6)",
         "SUMPRODUCT multiplies pairs and adds them — one function instead of four terms."),
        ('=IF(F6>=90,"A",IF(F6>=80,"B",IF(F6>=70,"C",IF(F6>=60,"D","F"))))', "Nested IF, highest first."),
        (f"=AVERAGE(F6:F{last}), =MAX(F6:F{last}), =MIN(F6:F{last})", "Data rows only."),
        ("(conditional formatting)", "Home ▸ Conditional Formatting ▸ Highlight Cells ▸ Less Than ▸ 60."),
        ("(test)", "If it does not update, a weight was typed into a formula."),
    ]
    notes = [
        "The weights row needs its ROW locked (B$3) so the formula survives being dragged sideways.",
        "SUMPRODUCT is the professional version of this formula — worth learning now.",
        "Check that the weights sum to 100% before you trust any answer.",
    ]
    return tasks, answers, notes


def _build_budget(ws, rows, rng):
    ws["A1"] = "Budget vs Actual"
    ws["A1"].font = TITLE_FONT
    headers = ["Month", "Category", "Budget", "Actual", "Variance", "Variance %", "Status"]
    months = ["January", "February", "March", "April", "May", "June"]
    recs = []
    for i in range(rows):
        budget = round(rng.uniform(1500, 22000), -2)
        actual = round(budget * rng.uniform(0.7, 1.35), -1)
        recs.append([months[i % len(months)], CATEGORIES[i % len(CATEGORIES)],
                     budget, actual, None, None, None])
    _write_table(ws, headers, recs,
                 ["General", "General", '$#,##0', '$#,##0', '$#,##0', "0.0%", "General"],
                 start_row=2)
    last = 2 + rows
    tasks = [
        f"In E3, calculate Variance (Budget − Actual). Fill down to E{last}.",
        f"In F3, calculate Variance % (variance as a share of budget). Format as a percentage.",
        'In G3, return "Over budget" when the variance is negative and "Within budget" otherwise.',
        "Apply conditional formatting so negative variances fill light red.",
        f"In E{last + 2}, total the variance. Format the totals row bold with a top border.",
        "Build a pivot table showing total Budget and total Actual by Category.",
    ]
    answers = [
        ("=C3-D3", "Both relative."),
        ("=E3/C3", "Relative; format the cell as a percentage rather than multiplying by 100."),
        ('=IF(E3<0,"Over budget","Within budget")', "Single IF."),
        ("(conditional formatting)", "Home ▸ Conditional Formatting ▸ Highlight Cells ▸ Less Than ▸ 0."),
        (f"=SUM(E3:E{last})", "Data rows only, not the totals row itself."),
        ("(pivot)", "Rows = Category; Values = Sum of Budget and Sum of Actual."),
    ]
    notes = [
        "Do not multiply by 100 for a percentage — apply Percentage format instead.",
        "Never include the totals cell inside the range it totals: that is a circular reference.",
        "Negative currency is conventionally shown in red parentheses via Accounting format.",
    ]
    return tasks, answers, notes
