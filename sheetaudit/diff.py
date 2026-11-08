"""Semantic diff between two workbook models, plus audit risk flags."""
from __future__ import annotations

import math
from collections import Counter
from dataclasses import dataclass, field
from enum import Enum
from typing import Any, Optional

from .align import AxisMap, align_sheets
from .deps import DepGraph
from .model import Cell, Sheet, Workbook
from .refs import (
    addr,
    error_tokens,
    has_external_link,
    numeric_constants,
    remap_formula,
    to_r1c1,
)

Pos = tuple[int, int]


class Kind(str, Enum):
    FORMULA_TO_VALUE = "formula_to_value"
    FORMULA_CHANGED = "formula_changed"
    VALUE_TO_FORMULA = "value_to_formula"
    INPUT_CHANGED = "input_changed"
    ADDED = "added"
    REMOVED = "removed"


KIND_LABEL = {
    Kind.FORMULA_TO_VALUE: "Formula → hardcoded",
    Kind.FORMULA_CHANGED: "Formula changed",
    Kind.VALUE_TO_FORMULA: "Hardcoded → formula",
    Kind.INPUT_CHANGED: "Hardcoded value changed",
    Kind.ADDED: "Added",
    Kind.REMOVED: "Removed",
}

SEVERITY_ORDER = {"high": 0, "medium": 1, "low": 2}
_BENIGN_CONSTANTS = {"0", "1"}


@dataclass
class CellChange:
    sheet: str
    kind: Kind
    old_pos: Optional[Pos]
    new_pos: Optional[Pos]
    old: Optional[Cell]
    new: Optional[Cell]
    label: str = ""

    @property
    def address(self) -> str:
        r, c = self.new_pos or self.old_pos
        moved = self.new_pos and self.old_pos and self.new_pos != self.old_pos
        return addr(r, c) + (f" (was {addr(*self.old_pos)})" if moved else "")


@dataclass
class ResultChange:
    """A formula whose text did not change but whose calculated value did."""

    sheet: str
    pos: Pos
    formula: str
    old_value: Any
    new_value: Any
    label: str = ""
    is_output: bool = False  # nothing else in the workbook depends on it

    @property
    def pct(self) -> Optional[float]:
        if _is_num(self.old_value) and _is_num(self.new_value) and self.old_value != 0:
            return (self.new_value - self.old_value) / abs(self.old_value)
        return None


@dataclass
class Flag:
    severity: str  # high | medium | low
    sheet: str
    address: str
    title: str
    detail: str = ""
    group: Optional[tuple] = None  # flags with the same group key are merged (e.g. B6:F6)
    pos: Optional[Pos] = None


def _merge_flags(flags: list[Flag]) -> list[Flag]:
    """Collapse the same issue copied across a row/column into one flag with a range."""
    out: list[Flag] = []
    groups: dict[tuple, list[Flag]] = {}
    for f in flags:
        if f.group is None or f.pos is None:
            out.append(f)
            continue
        key = (f.severity, f.sheet, f.title, f.group)
        if key not in groups:
            groups[key] = []
            out.append(f)  # placeholder keeps original ordering
        groups[key].append(f)
    merged: list[Flag] = []
    for f in out:
        if f.group is None or f.pos is None:
            merged.append(f)
            continue
        members = groups.get((f.severity, f.sheet, f.title, f.group))
        if members is None:
            continue
        del groups[(f.severity, f.sheet, f.title, f.group)]
        if len(members) == 1:
            merged.append(f)
            continue
        rows = {m.pos[0] for m in members}
        cols = {m.pos[1] for m in members}
        r1, r2, c1, c2 = min(rows), max(rows), min(cols), max(cols)
        if (r2 - r1 + 1) * (c2 - c1 + 1) == len(members):
            where = addr(r1, c1) + ":" + addr(r2, c2)
        else:
            where = ", ".join(addr(*m.pos) for m in members[:6]) + ("…" if len(members) > 6 else "")
        merged.append(Flag(f.severity, f.sheet, where, f.title,
                           f"{f.detail} (+{len(members) - 1} similar)", f.group, f.pos))
    return merged


@dataclass
class SheetDiff:
    name: str
    status: str  # modified | added | removed | renamed | unchanged
    old_name: Optional[str] = None
    changes: list[CellChange] = field(default_factory=list)
    results: list[ResultChange] = field(default_factory=list)
    rows_inserted: list[int] = field(default_factory=list)
    rows_deleted: list[int] = field(default_factory=list)
    cols_inserted: list[int] = field(default_factory=list)
    cols_deleted: list[int] = field(default_factory=list)
    state_change: Optional[tuple[str, str]] = None
    cell_count: int = 0
    formula_count: int = 0

    @property
    def has_changes(self) -> bool:
        return bool(
            self.changes or self.results or self.rows_inserted or self.rows_deleted
            or self.cols_inserted or self.cols_deleted or self.state_change
            or self.status in ("added", "removed", "renamed")
        )


@dataclass
class WorkbookDiff:
    path: str
    status: str  # modified | added | removed | renamed
    old_path: Optional[str] = None
    sheets: list[SheetDiff] = field(default_factory=list)
    names_added: dict[str, str] = field(default_factory=dict)
    names_removed: dict[str, str] = field(default_factory=dict)
    names_changed: dict[str, tuple[str, str]] = field(default_factory=dict)
    flags: list[Flag] = field(default_factory=list)
    downstream: list[tuple[str, Pos, str]] = field(default_factory=list)  # (sheet, pos, label)
    downstream_count: int = 0
    key_outputs: list[ResultChange] = field(default_factory=list)
    has_cached_values: bool = False
    error: Optional[str] = None

    def counts(self) -> Counter:
        c: Counter = Counter()
        for s in self.sheets:
            for ch in s.changes:
                c[ch.kind] += 1
        return c

    @property
    def total_changes(self) -> int:
        return sum(len(s.changes) for s in self.sheets)

    def flag_counts(self) -> Counter:
        return Counter(f.severity for f in self.flags)

    @property
    def has_changes(self) -> bool:
        return bool(
            self.error or any(s.has_changes for s in self.sheets) or self.names_added
            or self.names_removed or self.names_changed or self.status in ("added", "removed")
        )


# --------------------------------------------------------------------------- helpers


def _is_num(v: Any) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def values_equal(a: Any, b: Any) -> bool:
    if _is_num(a) and _is_num(b):
        if isinstance(a, float) and isinstance(b, float) and math.isnan(a) and math.isnan(b):
            return True
        return math.isclose(a, b, rel_tol=1e-9, abs_tol=1e-12)
    if isinstance(a, bool) != isinstance(b, bool):
        return False
    return a == b


def cell_label(sheet: Sheet, r: int, c: int, scan: int = 60) -> str:
    """Human label for a cell: nearest text to the left (row label) and above (column header)."""

    def text_at(rr, cc):
        cell = sheet.cells.get((rr, cc))
        if cell and not cell.is_formula and isinstance(cell.value, str) and cell.value.strip():
            if cell.kind == "text":
                return cell.value.strip()
        return None

    row_label = next((t for cc in range(c - 1, max(0, c - scan), -1) if (t := text_at(r, cc))), None)
    col_label = None
    # Column header: walk up the contiguous block above the cell. Finance headers
    # are often text ("Q1") or years (2026).
    for rr in range(r - 1, max(0, r - scan), -1):
        cell = sheet.cells.get((rr, c))
        if cell is None:
            break
        if t := text_at(rr, c):
            col_label = t
            break
        if not cell.is_formula and isinstance(cell.value, int) and 1900 <= cell.value <= 2100:
            col_label = str(cell.value)
            break
    parts = [p for p in (row_label, col_label) if p]
    label = " · ".join(dict.fromkeys(parts))
    return label[:80]


def _pair_sheets(old: Workbook, new: Workbook):
    pairs: list[tuple[Optional[Sheet], Optional[Sheet]]] = []
    new_by_lower = {n.lower(): s for n, s in new.sheets.items()}
    matched_new: set[str] = set()
    removed: list[Sheet] = []
    for name, s in old.sheets.items():
        ns = new_by_lower.get(name.lower())
        if ns is not None:
            pairs.append((s, ns))
            matched_new.add(ns.name)
        else:
            removed.append(s)
    added = [s for n, s in new.sheets.items() if n not in matched_new]

    # Rename detection: a removed and an added sheet with mostly identical content.
    def content(s: Sheet) -> set:
        return {(k, v.formula or repr(v.value)) for k, v in s.cells.items()}

    for rs in list(removed):
        best, best_score = None, 0.0
        rc = content(rs)
        for a in added:
            ac = content(a)
            union = len(rc | ac) or 1
            score = len(rc & ac) / union
            if score > best_score:
                best, best_score = a, score
        if best is not None and best_score >= 0.6:
            pairs.append((rs, best))
            removed.remove(rs)
            added.remove(best)
    pairs.extend((s, None) for s in removed)
    pairs.extend((None, s) for s in added)
    order = {n: i for i, n in enumerate(new.sheets)}
    pairs.sort(key=lambda p: order.get(p[1].name, 10**6) if p[1] else 10**6)
    return pairs


def _ranges(nums: list[int]) -> list[str]:
    out, start, prev = [], None, None
    for n in sorted(nums):
        if start is None:
            start = prev = n
        elif n == prev + 1:
            prev = n
        else:
            out.append(f"{start}" if start == prev else f"{start}–{prev}")
            start = prev = n
    if start is not None:
        out.append(f"{start}" if start == prev else f"{start}–{prev}")
    return out


# --------------------------------------------------------------------------- core


def diff_workbooks(
    old: Optional[Workbook], new: Optional[Workbook], path: str, old_path: Optional[str] = None
) -> WorkbookDiff:
    status = "added" if old is None else "removed" if new is None else "modified"
    if status == "modified" and old_path and old_path != path:
        status = "renamed"
    wd = WorkbookDiff(path=path, status=status, old_path=old_path)
    old = old or Workbook()
    new = new or Workbook()
    wd.has_cached_values = old.has_cached_values and new.has_cached_values

    pairs = _pair_sheets(old, new)
    maps: dict[str, tuple[AxisMap, AxisMap]] = {}
    rename: dict[str, str] = {}
    for os_, ns in pairs:
        if os_ and ns:
            maps[os_.name.lower()] = align_sheets(os_, ns)
            if os_.name != ns.name:
                rename[os_.name.lower()] = ns.name

    def mapper(sheet_name, r, c):
        key = (sheet_name or current[0]).lower()
        rmap, cmap = maps.get(key, (AxisMap(), AxisMap()))
        return (rmap.translate(r) if r is not None else None, cmap.translate(c) if c is not None else None)

    current = [""]
    any_shift = any(not (rm.identity and cm.identity) for rm, cm in maps.values())

    for os_, ns in pairs:
        if os_ is None:
            sd = SheetDiff(ns.name, "added", cell_count=len(ns.cells),
                           formula_count=sum(c.is_formula for c in ns.cells.values()))
            if status == "modified" or status == "renamed":
                wd.flags.append(Flag("low", ns.name, "", "Sheet added",
                                     f"{sd.cell_count} cells, {sd.formula_count} formulas"))
            wd.sheets.append(sd)
            continue
        if ns is None:
            sd = SheetDiff(os_.name, "removed", cell_count=len(os_.cells),
                           formula_count=sum(c.is_formula for c in os_.cells.values()))
            if status in ("modified", "renamed"):
                wd.flags.append(Flag("medium", os_.name, "", "Sheet deleted",
                                     f"{sd.cell_count} cells, {sd.formula_count} formulas removed"))
            wd.sheets.append(sd)
            continue

        current[0] = os_.name
        rmap, cmap = maps[os_.name.lower()]
        sd = SheetDiff(ns.name, "renamed" if os_.name != ns.name else "modified",
                       old_name=os_.name if os_.name != ns.name else None)
        if not rmap.identity:
            sd.rows_inserted, sd.rows_deleted = rmap.inserted, rmap.deleted
        if not cmap.identity:
            sd.cols_inserted, sd.cols_deleted = cmap.inserted, cmap.deleted
        if os_.state != ns.state:
            sd.state_change = (os_.state, ns.state)
            if ns.state != "visible":
                wd.flags.append(Flag("low" if ns.state == "hidden" else "medium", ns.name, "",
                                     f"Sheet made {ns.state}",
                                     "Hidden sheets can conceal assumptions from reviewers."))

        seen_new: set[Pos] = set()
        for (r, c), oc in os_.cells.items():
            nr, nc = rmap.get(r), cmap.get(c)
            if nr is None or nc is None:
                sd.changes.append(CellChange(ns.name, Kind.REMOVED, (r, c), None, oc, None))
                continue
            npos = (nr, nc)
            seen_new.add(npos)
            ncell = ns.cells.get(npos)
            if ncell is None:
                sd.changes.append(CellChange(ns.name, Kind.REMOVED, (r, c), npos, oc, None))
                continue
            _compare(sd, ns.name, (r, c), npos, oc, ncell, mapper if any_shift else None, rename)
        for npos, ncell in ns.cells.items():
            if npos not in seen_new:
                sd.changes.append(CellChange(ns.name, Kind.ADDED, None, npos, None, ncell))

        sd.changes.sort(key=lambda ch: ((ch.new_pos or ch.old_pos)[0], (ch.new_pos or ch.old_pos)[1]))
        for ch in sd.changes:
            src = ns if ch.new_pos and ch.new else os_
            cell = ch.new if src is ns else ch.old
            if cell is not None and cell.kind == "text":
                continue  # a label cell itself; its neighbours aren't meaningful labels
            ch.label = cell_label(src, *(ch.new_pos if src is ns else ch.old_pos))
        for rc in sd.results:
            rc.label = cell_label(ns, *rc.pos)
        wd.sheets.append(sd)

    _diff_names(wd, old, new, rename)
    _audit(wd, old, new, maps)
    wd.flags = _merge_flags(wd.flags)
    wd.flags.sort(key=lambda f: SEVERITY_ORDER[f.severity])
    return wd


def _compare(sd, sheet, opos, npos, oc: Cell, nc: Cell, mapper, rename):
    if oc.is_formula and nc.is_formula:
        expected = oc.formula
        if mapper is not None or rename:
            expected = remap_formula(oc.formula, mapper or (lambda s, r, c: (r, c)), rename)
        if expected != nc.formula:
            sd.changes.append(CellChange(sheet, Kind.FORMULA_CHANGED, opos, npos, oc, nc))
        elif not values_equal(oc.value, nc.value) and oc.value is not None and nc.value is not None:
            sd.results.append(ResultChange(sheet, npos, nc.formula, oc.value, nc.value))
    elif oc.is_formula:
        sd.changes.append(CellChange(sheet, Kind.FORMULA_TO_VALUE, opos, npos, oc, nc))
    elif nc.is_formula:
        sd.changes.append(CellChange(sheet, Kind.VALUE_TO_FORMULA, opos, npos, oc, nc))
    elif not values_equal(oc.value, nc.value):
        sd.changes.append(CellChange(sheet, Kind.INPUT_CHANGED, opos, npos, oc, nc))


def _diff_names(wd: WorkbookDiff, old: Workbook, new: Workbook, rename):
    for name, target in new.names.items():
        if name not in old.names:
            wd.names_added[name] = target
        else:
            expected = remap_formula("=" + old.names[name], lambda s, r, c: (r, c), rename)[1:]
            if expected.replace("'", "") != target.replace("'", ""):
                wd.names_changed[name] = (old.names[name], target)
                wd.flags.append(Flag("medium", "", name, "Named range redefined",
                                     f"`{old.names[name]}` → `{target}`"))
    for name, target in old.names.items():
        if name not in new.names:
            wd.names_removed[name] = target
            wd.flags.append(Flag("medium", "", name, "Named range deleted", f"was `{target}`"))


def _pattern_break(sheet: Sheet, r: int, c: int) -> Optional[str]:
    """Does this formula deviate from the formula its neighbours share?"""
    me = sheet.cells.get((r, c))
    if not me or not me.is_formula:
        return None
    mine = to_r1c1(me.formula, r, c)
    for a, b, axis in (((r, c - 1), (r, c + 1), "row"), ((r - 1, c), (r + 1, c), "column")):
        ca, cb = sheet.cells.get(a), sheet.cells.get(b)
        if ca and cb and ca.is_formula and cb.is_formula:
            fa, fb = to_r1c1(ca.formula, *a), to_r1c1(cb.formula, *b)
            if fa == fb and fa != mine:
                return f"neighbours {addr(*a)} and {addr(*b)} share one formula pattern; this cell differs ({axis})"
    return None


def _fmt(v: Any) -> str:
    from .report import fmt_value

    return fmt_value(v)


def _audit(wd: WorkbookDiff, old: Workbook, new: Workbook, maps):
    graph = DepGraph(new) if new.sheets else None
    seeds = []
    for sd in wd.sheets:
        if sd.status in ("added", "removed"):
            continue
        ns = new.sheets[sd.name]
        for ch in sd.changes:
            where = ch.address.split(" ")[0]
            if ch.new_pos:
                seeds.append((sd.name.lower(), *ch.new_pos))
            npos = ch.new_pos
            if ch.kind == Kind.FORMULA_TO_VALUE:
                wd.flags.append(Flag("high", sd.name, where, "Formula overwritten with a hardcoded value",
                                     f"`{ch.old.formula}` → `{_fmt(ch.new.value)}`",
                                     ("f2v", to_r1c1(ch.old.formula, *ch.old_pos)), npos))
            if ch.kind in (Kind.FORMULA_CHANGED, Kind.VALUE_TO_FORMULA, Kind.ADDED) and ch.new and ch.new.is_formula:
                f_new = ch.new.formula
                f_old = ch.old.formula if ch.old and ch.old.is_formula else ""
                g = ("f", to_r1c1(f_new, *npos))
                new_errs = error_tokens(f_new) - error_tokens(f_old)
                if new_errs:
                    wd.flags.append(Flag("high", sd.name, where, f"Broken reference {', '.join(sorted(new_errs))} in formula",
                                         f"`{f_new}`", g, npos))
                added_consts = Counter(numeric_constants(f_new)) - Counter(numeric_constants(f_old))
                consts = [k for k in added_consts if k not in _BENIGN_CONSTANTS]
                if consts:
                    wd.flags.append(Flag("medium", sd.name, where, "Hardcoded number inside formula",
                                         f"{', '.join('`'+k+'`' for k in consts)} in `{f_new}`", g, npos))
                if has_external_link(f_new) and not (f_old and has_external_link(f_old)):
                    wd.flags.append(Flag("medium", sd.name, where, "Link to external workbook", f"`{f_new}`", g, npos))
                pb = _pattern_break(ns, *ch.new_pos)
                if pb:
                    wd.flags.append(Flag("medium", sd.name, where, "Inconsistent formula", pb))
            if ch.kind == Kind.INPUT_CHANGED and _is_num(ch.old.value) and _is_num(ch.new.value):
                a, b = abs(ch.old.value), abs(ch.new.value)
                if a and b and (b / a >= 50 or a / b >= 50):
                    wd.flags.append(Flag("medium", sd.name, where, "Order-of-magnitude change to input",
                                         f"`{_fmt(ch.old.value)}` → `{_fmt(ch.new.value)}` "
                                         f"(×{b / a:,.4g}) — possible unit or % error"))
            # New calculation errors
            if ch.new and ch.new.is_formula and ch.new.is_error and not (ch.old and ch.old.is_error):
                wd.flags.append(Flag("high", sd.name, where, f"Formula now evaluates to {ch.new.value}",
                                     f"`{ch.new.formula}`", ("err", to_r1c1(ch.new.formula, *npos)), npos))
        for rc in sd.results:
            if isinstance(rc.new_value, str) and rc.new_value.startswith("#") and not (
                isinstance(rc.old_value, str) and rc.old_value.startswith("#")
            ):
                wd.flags.append(Flag("high", sd.name, addr(*rc.pos), f"Formula now evaluates to {rc.new_value}",
                                     f"`{rc.formula}`", ("err", to_r1c1(rc.formula, *rc.pos)), rc.pos))

    if graph is None:
        return
    down = graph.downstream(seeds)
    wd.downstream_count = len(down)
    shown = sorted(down, key=lambda k: (graph.has_dependents(k), k))  # final outputs first
    for key in shown[:200]:
        s = new.sheet(key[0])
        if s is not None:
            wd.downstream.append((s.name, key[1:], cell_label(s, key[1], key[2])))

    outputs = []
    for sd in wd.sheets:
        for rc in sd.results:
            rc.is_output = not graph.has_dependents((sd.name.lower(), *rc.pos))
            outputs.append(rc)
    outputs.sort(key=lambda rc: (not rc.is_output, -(abs(rc.pct) if rc.pct is not None else 0)))
    wd.key_outputs = outputs[:15]


def group_changes(changes: list[CellChange]) -> list[list[CellChange]]:
    """Group the same edit copied across a contiguous row/column block (e.g. B7:F7)."""

    def key(ch: CellChange):
        if ch.kind in (Kind.INPUT_CHANGED,) or not ch.new_pos:
            return None
        o = to_r1c1(ch.old.formula, *ch.old_pos) if ch.old and ch.old.is_formula else (
            repr(ch.old.value) if ch.old else None)
        n = to_r1c1(ch.new.formula, *ch.new_pos) if ch.new and ch.new.is_formula else (
            repr(ch.new.value) if ch.new else None)
        if not ((ch.old and ch.old.is_formula) or (ch.new and ch.new.is_formula)):
            return None
        return (ch.kind, o, n)

    buckets: dict = {}
    order: list = []
    for ch in changes:
        k = key(ch)
        if k is None:
            order.append([ch])
            continue
        if k not in buckets:
            buckets[k] = []
            order.append(buckets[k])
        buckets[k].append(ch)
    out: list[list[CellChange]] = []
    for grp in order:
        if len(grp) == 1:
            out.append(grp)
            continue
        # split into contiguous runs along a row or a column
        grp.sort(key=lambda ch: ch.new_pos)
        rows = {ch.new_pos[0] for ch in grp}
        cols = {ch.new_pos[1] for ch in grp}
        r1, r2, c1, c2 = min(rows), max(rows), min(cols), max(cols)
        if (len(rows) == 1 or len(cols) == 1) and (r2 - r1 + 1) * (c2 - c1 + 1) == len(grp):
            out.append(grp)
        else:
            out.extend([ch] for ch in grp)
    return out


def summarize_ranges(nums: list[int]) -> str:
    return ", ".join(_ranges(nums))
