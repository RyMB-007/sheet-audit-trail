"""Render workbook diffs as GitHub-flavoured Markdown (and JSON)."""
from __future__ import annotations

import json
from typing import Any, Iterable, Optional

from .diff import KIND_LABEL, CellChange, Kind, WorkbookDiff, group_changes, summarize_ranges
from .model import Cell
from .refs import addr

MARKER = "<!-- sheet-audit-trail -->"
GITHUB_COMMENT_LIMIT = 65_000

SEV_ICON = {"high": "🔴", "medium": "🟠", "low": "🟡"}
KIND_ICON = {
    Kind.FORMULA_TO_VALUE: "🔴",
    Kind.FORMULA_CHANGED: "🟣",
    Kind.VALUE_TO_FORMULA: "🔵",
    Kind.INPUT_CHANGED: "🟢",
    Kind.ADDED: "➕",
    Kind.REMOVED: "➖",
}


# --------------------------------------------------------------------------- formatting


def fmt_value(v: Any) -> str:
    if v is None:
        return "(blank)"
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, int):
        return f"{v:,}"
    if isinstance(v, float):
        if abs(v) >= 1000:
            return f"{v:,.2f}".rstrip("0").rstrip(".")
        return f"{v:.6g}"
    s = str(v).replace("\n", " ")
    return s if len(s) <= 60 else s[:57] + "…"


def _esc(s: str) -> str:
    return s.replace("|", "\\|").replace("\n", " ").replace("\r", "")


def _esc_md(s: str) -> str:
    """Escape for a table cell that already contains `code spans`."""
    return s.replace("|", "\\|").replace("\n", " ")


def code(s: str, limit: int = 120) -> str:
    s = s.replace("\n", " ")
    if len(s) > limit:
        s = s[: limit - 1] + "…"
    fence = "``" if "`" in s else "`"
    pad = " " if fence == "``" else ""
    return _esc(f"{fence}{pad}{s}{pad}{fence}")


def _cell_md(cell: Optional[Cell], show_result: bool) -> str:
    if cell is None:
        return "_(blank)_"
    if cell.is_formula:
        out = code(cell.formula)
        if show_result and cell.value is not None:
            out += f"<br><sub>= {_esc(fmt_value(cell.value))}</sub>"
        return out
    return code(fmt_value(cell.value))


def _pct(p: Optional[float]) -> str:
    if p is None:
        return "—"
    arrow = "▲" if p > 0 else "▼" if p < 0 else ""
    if abs(p) >= 10:  # beyond ±1000% a multiplier reads better
        return f"{arrow} {p:+,.0%}"
    return f"{arrow} {p:+.2%}"


def _change_pct(rc) -> str:
    if (rc.pct is not None and abs(rc.pct) >= 10 and rc.old_value
            and isinstance(rc.new_value, (int, float)) and rc.new_value / rc.old_value > 0):
        return f"{'▲' if rc.pct > 0 else '▼'} ×{rc.new_value / rc.old_value:,.0f}"
    return _pct(rc.pct)


def _plural(n: int, word: str) -> str:
    return f"{n} {word}{'' if n == 1 else 's'}"


# --------------------------------------------------------------------------- sections


def _summary_line(wd: WorkbookDiff) -> str:
    c = wd.counts()
    parts = [f"{KIND_ICON[k]} {c[k]} {KIND_LABEL[k].lower()}" for k in Kind if c[k]]
    results = sum(len(s.results) for s in wd.sheets)
    if results:
        parts.append(f"🔁 {_plural(results, 'recalculated value')}")
    return " · ".join(parts) or "No cell-level changes"


def _flags_md(wd: WorkbookDiff, max_rows: int) -> list[str]:
    if not wd.flags:
        return []
    out = ["#### 🚩 Risk flags", "", "| | Location | Issue | Detail |", "|---|---|---|---|"]
    for f in wd.flags[:max_rows]:
        loc = f"{f.sheet}!{f.address}" if f.sheet and f.address else (f.sheet or f.address)
        out.append(f"| {SEV_ICON[f.severity]} | {code(loc) if loc else ''} | {_esc(f.title)} | {_esc_md(f.detail)} |")
    if len(wd.flags) > max_rows:
        out.append(f"\n_…and {len(wd.flags) - max_rows} more flags._")
    out.append("")
    return out


def _impact_md(wd: WorkbookDiff, max_rows: int) -> list[str]:
    out: list[str] = []
    if wd.key_outputs:
        out += ["#### 🎯 Impact on calculated values", "",
                "Formulas that did **not** change but now produce a different result:", "",
                "| Cell | Label | Before | After | Change |", "|---|---|---:|---:|---:|"]
        for rc in wd.key_outputs[:max_rows]:
            star = " ⭐" if rc.is_output else ""
            out.append(
                f"| {code(f'{rc.sheet}!{addr(*rc.pos)}')}{star} | {_esc(rc.label)} | {_esc(fmt_value(rc.old_value))} "
                f"| {_esc(fmt_value(rc.new_value))} | {_change_pct(rc)} |"
            )
        out += ["", "<sub>⭐ = final output (no other cell depends on it)</sub>", ""]
    if wd.downstream_count:
        out.append(f"<details><summary>🕸️ <b>{_plural(wd.downstream_count, 'cell')}</b> depend on the changed cells "
                   f"(directly or indirectly)</summary>\n")
        out += ["| Cell | Label |", "|---|---|"]
        for sheet, pos, label in wd.downstream[:max_rows]:
            out.append(f"| {code(f'{sheet}!{addr(*pos)}')} | {_esc(label)} |")
        if wd.downstream_count > max_rows:
            out.append(f"\n_…and {wd.downstream_count - max_rows} more._")
        out += ["", "</details>", ""]
    return out


def _change_row(group: list[CellChange], show_result: bool) -> str:
    ch = group[0]
    where = code(ch.address)
    label = ch.label
    if len(group) > 1:
        last = group[-1]
        where = code(f"{addr(*ch.new_pos)}:{addr(*last.new_pos)}") + f"<br><sub>{len(group)} cells</sub>"
        label = label.split(" · ")[0]  # row label; the per-column header no longer applies
    return (f"| {where} | {_esc(label)} | {KIND_ICON[ch.kind]} {KIND_LABEL[ch.kind]} "
            f"| {_cell_md(ch.old, show_result)} | {_cell_md(ch.new, show_result)} |")


def _sheets_md(wd: WorkbookDiff, max_rows: int) -> list[str]:
    out: list[str] = []
    changed = [s for s in wd.sheets if s.has_changes]
    if not changed:
        return out
    out += ["#### 🔍 Changes by sheet", ""]
    open_attr = " open" if wd.total_changes <= 40 else ""
    for sd in changed:
        notes = []
        if sd.status == "added":
            notes.append(f"➕ **New sheet** — {sd.cell_count} cells, {sd.formula_count} formulas")
        elif sd.status == "removed":
            notes.append(f"➖ **Sheet deleted** — {sd.cell_count} cells, {sd.formula_count} formulas")
        elif sd.status == "renamed":
            notes.append(f"✏️ Renamed from {code(sd.old_name)}")
        if sd.rows_inserted:
            notes.append(f"↕️ Rows inserted: {summarize_ranges(sd.rows_inserted)}")
        if sd.rows_deleted:
            notes.append(f"↕️ Rows deleted (old numbering): {summarize_ranges(sd.rows_deleted)}")
        if sd.cols_inserted:
            notes.append(f"↔️ Columns inserted at: {summarize_ranges(sd.cols_inserted)}")
        if sd.cols_deleted:
            notes.append(f"↔️ Columns deleted (old numbering): {summarize_ranges(sd.cols_deleted)}")
        if sd.state_change:
            notes.append(f"👁️ Visibility: {sd.state_change[0]} → **{sd.state_change[1]}**")
        title = f"<b>{_esc(sd.name)}</b>"
        if sd.changes:
            title += f" — {_plural(len(sd.changes), 'change')}"
        out.append(f"<details{open_attr}><summary>{title}</summary>\n")
        out += [f"- {n}" for n in notes]
        if notes:
            out.append("")
        if sd.changes:
            out += ["| Cell | Label | Change | Before | After |", "|---|---|---|---|---|"]
            # Most audit-relevant changes first, then by position.
            order = list(Kind)
            groups = sorted(group_changes(sd.changes), key=lambda g: order.index(g[0].kind))
            for g in groups[:max_rows]:
                out.append(_change_row(g, wd.has_cached_values))
            if len(groups) > max_rows:
                hidden = sum(len(g) for g in groups[max_rows:])
                out.append(f"\n_…and {hidden} more changed cells on this sheet._")
            out.append("")
        out += ["</details>", ""]
    return out


def _names_md(wd: WorkbookDiff) -> list[str]:
    rows = []
    for n, t in wd.names_added.items():
        rows.append(f"| {code(n)} | ➕ added | | {code(t)} |")
    for n, (a, b) in wd.names_changed.items():
        rows.append(f"| {code(n)} | ✏️ redefined | {code(a)} | {code(b)} |")
    for n, t in wd.names_removed.items():
        rows.append(f"| {code(n)} | ➖ removed | {code(t)} | |")
    if not rows:
        return []
    return ["#### 🏷️ Named ranges", "", "| Name | Change | Before | After |", "|---|---|---|---|", *rows, ""]


def render_workbook(wd: WorkbookDiff, max_rows: int = 100) -> str:
    title = f"### 📗 `{wd.path}`"
    if wd.status == "renamed":
        title += f" <sub>(renamed from `{wd.old_path}`)</sub>"
    out = [title, ""]
    if wd.error:
        out += [f"> ⚠️ Could not parse this workbook: {_esc(wd.error)}", ""]
        return "\n".join(out)
    if wd.status == "added":
        out.append("**New workbook.**")
    elif wd.status == "removed":
        out.append("**Workbook deleted.**")
    else:
        out.append(f"**Summary:** {_summary_line(wd)}")
    if not wd.has_cached_values and wd.status in ("modified", "renamed"):
        out.append("\n<sub>ℹ️ No cached calculation results found (file not saved by Excel/Sheets), "
                   "so value-impact analysis is limited to the dependency graph.</sub>")
    out.append("")
    out += _flags_md(wd, max_rows)
    out += _impact_md(wd, max_rows)
    out += _sheets_md(wd, max_rows)
    out += _names_md(wd)
    return "\n".join(out)


def render(diffs: Iterable[WorkbookDiff], max_rows: int = 100, footer: str = "") -> str:
    diffs = [d for d in diffs if d.has_changes]
    out = [MARKER, "## 📊 Spreadsheet Audit Trail", ""]
    if not diffs:
        out.append("No content changes detected in spreadsheet files. ✅")
        return "\n".join(out) + "\n"
    total_flags = sum(len(d.flags) for d in diffs)
    high = sum(d.flag_counts()["high"] for d in diffs)
    verdict = ("🔴 **Review required:** high-risk changes detected." if high
               else "🟠 Some changes need a second look." if total_flags else "✅ No risk flags raised.")
    out += [verdict, "",
            "| Workbook | Status | Cell changes | 🔴 High | 🟠 Medium | 🟡 Low |",
            "|---|---|---:|---:|---:|---:|"]
    for d in diffs:
        fc = d.flag_counts()
        out.append(f"| {code(d.path)} | {d.status} | {d.total_changes} | {fc['high']} | {fc['medium']} | {fc['low']} |")
    out.append("")
    for d in diffs:
        out += ["---", "", render_workbook(d, max_rows)]
    out += ["---", "<sub>Generated by Sheet Audit Trail · 🔴 formula→hardcoded · 🟣 formula changed · "
            "🔵 hardcoded→formula · 🟢 input changed · ➕ added · ➖ removed</sub>"]
    if footer:
        out.append(footer)
    return "\n".join(out) + "\n"


def render_fitting(diffs: list[WorkbookDiff], limit: int = GITHUB_COMMENT_LIMIT, footer: str = "") -> str:
    """Render, shrinking row limits until the report fits in a PR comment."""
    for rows in (100, 40, 15, 5):
        md = render(diffs, rows, footer)
        if len(md) <= limit:
            return md
    md = render(diffs, 0, footer)
    if len(md) > limit:
        md = md[: limit - 200] + "\n\n_…report truncated. See the workflow run summary for the full report._\n"
    return md


# --------------------------------------------------------------------------- JSON


def _cell_json(c: Optional[Cell]):
    if c is None:
        return None
    return {"formula": c.formula, "value": c.value}


def to_json(diffs: list[WorkbookDiff]) -> str:
    def wb(d: WorkbookDiff):
        return {
            "path": d.path, "status": d.status, "old_path": d.old_path, "error": d.error,
            "counts": {k.value: v for k, v in d.counts().items()},
            "flags": [vars(f) for f in d.flags],
            "names": {"added": d.names_added, "removed": d.names_removed,
                      "changed": {k: list(v) for k, v in d.names_changed.items()}},
            "downstream_count": d.downstream_count,
            "sheets": [{
                "name": s.name, "status": s.status, "old_name": s.old_name,
                "rows_inserted": s.rows_inserted, "rows_deleted": s.rows_deleted,
                "cols_inserted": s.cols_inserted, "cols_deleted": s.cols_deleted,
                "changes": [{
                    "kind": ch.kind.value, "cell": ch.address, "label": ch.label,
                    "old": _cell_json(ch.old), "new": _cell_json(ch.new),
                } for ch in s.changes],
                "recalculated": [{
                    "cell": addr(*r.pos), "label": r.label, "old": r.old_value, "new": r.new_value,
                    "is_output": r.is_output,
                } for r in s.results],
            } for s in d.sheets if s.has_changes],
        }

    return json.dumps([wb(d) for d in diffs], indent=2, default=str)
