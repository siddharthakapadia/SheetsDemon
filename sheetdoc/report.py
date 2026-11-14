"""Findings and rendering.

Every audit function returns a Report of Findings. Text rendering is the default
because a human reads it; --json exists so the agent can chain results.
"""

from __future__ import annotations

import json
import shutil
from dataclasses import dataclass, field, asdict
from typing import Any

# "pass" means a check passed and is hidden unless -v. "done" means an action was
# performed — it must always be visible, or a write command prints nothing at all.
SEVERITY_ORDER = {"error": 0, "warning": 1, "done": 2, "info": 3, "pass": 4}
SEVERITY_ICON = {"error": "✗", "warning": "▲", "info": "·", "pass": "✓", "done": "✓"}


@dataclass
class Finding:
    severity: str  # error | warning | info | pass
    code: str
    cell: str | None
    sheet: str | None
    message: str
    detail: str = ""
    fix: str = ""
    formula: str = ""
    count: int = 1
    data: dict[str, Any] = field(default_factory=dict)


@dataclass
class Report:
    title: str
    findings: list[Finding] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)
    summary: dict[str, Any] = field(default_factory=dict)

    def add(self, *args, **kwargs) -> Finding:
        f = args[0] if args and isinstance(args[0], Finding) else Finding(*args, **kwargs)
        self.findings.append(f)
        return f

    def extend(self, other: "Report") -> "Report":
        self.findings.extend(other.findings)
        self.notes.extend(other.notes)
        return self

    @property
    def counts(self) -> dict[str, int]:
        """Counted after collapsing, so one mistake filled down 20 rows counts once."""
        c = {"error": 0, "warning": 0, "info": 0, "pass": 0, "done": 0}
        for f in self.collapse():
            c[f.severity] = c.get(f.severity, 0) + 1
        return c

    @property
    def clean(self) -> bool:
        return self.counts["error"] == 0 and self.counts["warning"] == 0

    def sorted_findings(self) -> list[Finding]:
        return sorted(
            self.collapse(),
            key=lambda f: (SEVERITY_ORDER.get(f.severity, 9), f.sheet or "", _cellkey(f.cell)),
        )

    def collapse(self) -> list[Finding]:
        """Merge one issue repeated down a filled column into a single finding.

        A missing $ in a formula that was filled to twenty rows is one mistake,
        not twenty. Reporting it twenty times buries everything else.
        """
        groups: dict[tuple, list[Finding]] = {}
        order: list[tuple] = []
        for f in self.findings:
            key = (f.severity, f.code, f.sheet, f.message, f.detail, f.fix)
            if key not in groups:
                order.append(key)
            groups.setdefault(key, []).append(f)

        out: list[Finding] = []
        for key in order:
            members = groups[key]
            first = members[0]
            if len(members) == 1:
                out.append(first)
                continue
            cells = [m.cell for m in members if m.cell]
            merged = Finding(
                severity=first.severity,
                code=first.code,
                cell=_span(cells),
                sheet=first.sheet,
                message=first.message,
                detail=first.detail,
                fix=first.fix,
                formula=first.formula,
                count=len(members),
                data=first.data,
            )
            out.append(merged)
        return out

    def to_json(self) -> str:
        return json.dumps(
            {
                "title": self.title,
                "summary": self.summary,
                "counts": self.counts,
                "notes": self.notes,
                "findings": [asdict(f) for f in self.sorted_findings()],
            },
            indent=2,
            default=str,
        )

    def to_text(self, verbose: bool = False) -> str:
        width = min(shutil.get_terminal_size((100, 24)).columns, 100)
        lines: list[str] = []
        lines.append("═" * width)
        lines.append(f"  {self.title}")
        lines.append("═" * width)

        if self.summary:
            for k, v in self.summary.items():
                lines.append(f"  {k}: {v}")
            lines.append("")

        shown = [f for f in self.sorted_findings() if verbose or f.severity != "pass"]
        if not shown:
            lines.append("  ✓ No issues found.")
        else:
            current_sheet = object()
            for f in shown:
                if f.sheet != current_sheet:
                    current_sheet = f.sheet
                    if f.sheet:
                        lines.append(f"\n  ── {f.sheet} " + "─" * max(0, width - len(f.sheet) - 8))
                icon = SEVERITY_ICON.get(f.severity, "?")
                loc = f"{f.cell:>11}" if f.cell else " " * 11
                tail = f"  [{f.count} cells]" if f.count > 1 else ""
                lines.append(f"  {icon} {loc}  {f.message}{tail}")
                pad = " " * 16
                if f.formula:
                    lines.append(f"{pad}formula: {f.formula}")
                if f.detail:
                    for para in f.detail.split("\n"):
                        lines.append(f"{pad}{para}")
                if f.fix:
                    lines.append(f"{pad}→ fix: {f.fix}")

        c = self.counts
        lines.append("")
        lines.append("─" * width)
        tally = (
            f"  {c['error']} error(s)   {c['warning']} warning(s)   "
            f"{c['info']} note(s)   {c['pass']} passed"
        )
        if c["done"]:
            tally += f"   {c['done']} change(s) applied"
        lines.append(tally)
        if self.notes:
            lines.append("")
            for n in self.notes:
                lines.append(f"  note: {n}")
        return "\n".join(lines)


def _span(cells: list[str]) -> str:
    """Render a set of cells as a range when they are contiguous, else a short list."""
    import re

    if not cells:
        return ""
    parsed = []
    for c in cells:
        m = re.match(r"^([A-Za-z]+)([0-9]+)$", c)
        if not m:
            return cells[0] + (f" +{len(cells) - 1}" if len(cells) > 1 else "")
        from .formula import col_index

        parsed.append((col_index(m.group(1)), int(m.group(2)), c))
    parsed.sort(key=lambda t: (t[0], t[1]))
    cols = {p[0] for p in parsed}
    rows = {p[1] for p in parsed}
    if len(cols) == 1 and len(rows) == len(parsed):
        rs = sorted(rows)
        if rs[-1] - rs[0] == len(rs) - 1:
            return f"{parsed[0][2]}:{parsed[-1][2]}"
    if len(rows) == 1 and len(cols) == len(parsed):
        return f"{parsed[0][2]}:{parsed[-1][2]}"
    return ", ".join(p[2] for p in parsed[:3]) + (f" +{len(parsed) - 3}" if len(parsed) > 3 else "")


def _cellkey(cell: str | None) -> tuple[int, int]:
    if not cell:
        return (0, 0)
    import re

    m = re.match(r"([A-Za-z]+)([0-9]+)", cell)
    if not m:
        return (0, 0)
    from .formula import col_index

    return (int(m.group(2)), col_index(m.group(1)))


def render(report: Report, fmt: str = "text", verbose: bool = False) -> str:
    return report.to_json() if fmt == "json" else report.to_text(verbose=verbose)
