"""Cell dependency graph: which formulas read from which cells."""
from __future__ import annotations

from collections import defaultdict, deque
from typing import Iterable

from .model import Workbook
from .refs import Ref, name_tokens, parse_refs, unquote_sheet

Key = tuple[str, int, int]  # (sheet name lower-cased, row, col)

_EXPAND_LIMIT = 2_000  # ranges smaller than this are expanded to individual cells


class DepGraph:
    def __init__(self, wb: Workbook):
        self.wb = wb
        self._point: dict[Key, set[Key]] = defaultdict(set)
        self._rects: dict[str, list[tuple[Ref, Key]]] = defaultdict(list)
        self.precedents: dict[Key, list[tuple[str, Ref]]] = {}
        name_refs = self._name_refs(wb)
        for sname, sheet in wb.sheets.items():
            for (r, c), cell in sheet.cells.items():
                if not cell.is_formula:
                    continue
                me: Key = (sname.lower(), r, c)
                targets: list[tuple[str, Ref]] = []
                for ref in parse_refs(cell.formula):
                    if not ref.external:
                        targets.append(((ref.sheet or sname).lower(), ref))
                for tok in name_tokens(cell.formula):
                    targets.extend(name_refs.get(f"{sname}!{tok}".upper(), name_refs.get(tok, [])))
                self.precedents[me] = targets
                for tsheet, ref in targets:
                    if ref.size <= _EXPAND_LIMIT:
                        for rr in range(ref.r1, ref.r2 + 1):
                            for cc in range(ref.c1, ref.c2 + 1):
                                self._point[(tsheet, rr, cc)].add(me)
                    else:
                        self._rects[tsheet].append((ref, me))

    @staticmethod
    def _name_refs(wb: Workbook) -> dict[str, list[tuple[str, Ref]]]:
        out: dict[str, list[tuple[str, Ref]]] = {}
        for name, target in wb.names.items():
            refs = []
            for ref in parse_refs("=" + target):
                if ref.sheet and not ref.external:
                    refs.append((ref.sheet.lower(), ref))
            out[name] = refs
        return out

    def dependents(self, key: Key) -> set[Key]:
        deps = set(self._point.get(key, ()))
        sheet, r, c = key
        for ref, me in self._rects.get(sheet, ()):
            if ref.contains(r, c):
                deps.add(me)
        return deps

    def has_dependents(self, key: Key) -> bool:
        if self._point.get(key):
            return True
        sheet, r, c = key
        return any(ref.contains(r, c) for ref, _ in self._rects.get(sheet, ()))

    def downstream(self, keys: Iterable[Key], limit: int = 100_000) -> set[Key]:
        """All cells that transitively depend on any of `keys` (excluding the keys)."""
        start = set(keys)
        seen: set[Key] = set()
        queue = deque(start)
        while queue and len(seen) < limit:
            for d in self.dependents(queue.popleft()):
                if d not in seen and d not in start:
                    seen.add(d)
                    queue.append(d)
        return seen


def sheet_display_name(wb: Workbook, lower: str) -> str:
    s = wb.sheet(lower)
    return s.name if s else unquote_sheet(lower)[0]
