"""Session stats for the command-center tiles.

Every number here is derived from work that actually happened — cells the engine
inspected, findings it produced, hints revealed. A tile that invents a number is
worse than no tile, because it teaches the student to distrust the panel.
"""

from __future__ import annotations

import json
from datetime import date, timedelta
from pathlib import Path

STORE = Path.home() / ".config" / "sheets-doctor" / "stats.json"


class Stats:
    def __init__(self, store: Path = STORE):
        self.store = store
        self.cells_inspected = 0
        self.formulas_seen = 0
        self.findings = 0
        self.errors = 0
        self.audits = 0
        self.hints = 0
        self._load()

    # ---------------- persistence (streak only) ----------------

    def _load(self) -> None:
        self.streak = 0
        self.last_day: str | None = None
        try:
            d = json.loads(self.store.read_text())
            self.streak = int(d.get("streak", 0))
            self.last_day = d.get("last_day")
        except Exception:
            pass
        self._touch_day()

    def _touch_day(self) -> None:
        today = date.today().isoformat()
        if self.last_day == today:
            return
        yesterday = (date.today() - timedelta(days=1)).isoformat()
        self.streak = self.streak + 1 if self.last_day == yesterday else 1
        self.last_day = today
        self._save()

    def _save(self) -> None:
        try:
            self.store.parent.mkdir(parents=True, exist_ok=True)
            self.store.write_text(json.dumps(
                {"streak": self.streak, "last_day": self.last_day}))
        except OSError:
            pass

    # ---------------- recording ----------------

    def record_audit(self, data: dict) -> None:
        self._touch_day()
        self.audits += 1
        counts = data.get("counts") or {}
        self.errors += counts.get("error", 0)
        self.findings += sum(
            counts.get(k, 0) for k in ("error", "warning", "info")
        )
        summary = data.get("summary") or {}
        for key in ("formulas scanned", "ranges checked", "lookup calls checked",
                    "numeric columns checked", "filled runs analysed"):
            val = summary.get(key)
            if isinstance(val, int):
                self.cells_inspected += val
                if key == "formulas scanned":
                    self.formulas_seen += val

    def record_hint(self) -> None:
        self._touch_day()
        self.hints += 1

    # ---------------- reporting ----------------

    @property
    def formula_health(self) -> int | None:
        """Share of inspected items with no error finding, as a percentage."""
        if not self.cells_inspected:
            return None
        clean = max(0, self.cells_inspected - self.errors)
        return round(100 * clean / self.cells_inspected)

    def snapshot(self) -> dict:
        health = self.formula_health
        return {
            "cells_inspected": self.cells_inspected,
            "formula_health": health,
            "formula_health_label": (
                None if health is None
                else "good" if health >= 90
                else "needs work" if health >= 70
                else "shaky"
            ),
            "hints": self.hints,
            "streak": self.streak,
            "audits": self.audits,
            "findings": self.findings,
            "errors": self.errors,
        }


STATS = Stats()
