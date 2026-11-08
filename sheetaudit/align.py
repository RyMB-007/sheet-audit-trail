"""Detect inserted/deleted rows and columns so they don't explode the diff.

Each row (and column) gets a position-independent signature built from its
cells (formulas in R1C1 form, constants by value). The old and new signature
sequences are aligned with difflib, the same way a text diff aligns lines.
"""
from __future__ import annotations

import bisect
from dataclasses import dataclass, field
from difflib import SequenceMatcher
from typing import Optional

from .model import Sheet
from .refs import to_r1c1

_MAX_AXIS = 20_000  # skip alignment on gigantic sheets (positional diff instead)


@dataclass
class AxisMap:
    mapping: dict[int, int] = field(default_factory=dict)  # old index -> new index
    inserted: list[int] = field(default_factory=list)  # new indices with no old counterpart
    deleted: list[int] = field(default_factory=list)  # old indices with no new counterpart
    identity: bool = True
    _keys: list[int] = field(default_factory=list, repr=False)

    def get(self, i: int) -> Optional[int]:
        if self.identity:
            return i
        return self.mapping.get(i)

    def translate(self, i: int) -> int:
        """Map any old index (even deleted / out of range) to the most plausible new one."""
        if self.identity:
            return i
        if i in self.mapping:
            return self.mapping[i]
        if not self._keys:
            self._keys = sorted(self.mapping)
        pos = bisect.bisect_left(self._keys, i) - 1
        if pos < 0:
            return i
        prev = self._keys[pos]
        return self.mapping[prev] + (i - prev)


def _cell_sig(cell, r: int, c: int):
    return ("f", to_r1c1(cell.formula, r, c)) if cell.is_formula else ("v", repr(cell.value))


def _signatures(sheet: Sheet, axis: int, length: int) -> tuple[list[tuple], list]:
    """Per row (axis=0) or column (axis=1): a content signature and a text label."""
    buckets: list[list] = [[] for _ in range(length)]
    for (r, c), cell in sheet.cells.items():
        idx, other = (r, c) if axis == 0 else (c, r)
        buckets[idx - 1].append((other, cell, r, c))
    sigs, labels = [], []
    for b in buckets:
        b.sort(key=lambda t: t[0])
        sigs.append(tuple(_cell_sig(cell, r, c) for _, cell, r, c in b))
        label = next((cell.value for _, cell, _, _ in b
                      if not cell.is_formula and isinstance(cell.value, str) and cell.value.strip()), None)
        labels.append(label)
    return sigs, labels


def _pair_block(am: "AxisMap", o_idx: range, n_idx: range, o_labels, n_labels) -> None:
    """Pair rows inside a changed block: first by matching labels, then by position."""
    if len(o_idx) == len(n_idx):
        am.mapping.update({i + 1: j + 1 for i, j in zip(o_idx, n_idx)})
        return
    ol = [o_labels[i] if o_labels[i] is not None else object() for i in o_idx]
    nl = [n_labels[j] if n_labels[j] is not None else object() for j in n_idx]
    sm = SequenceMatcher(None, ol, nl, autojunk=False)
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        n = min(i2 - i1, j2 - j1)
        for k in range(n):
            am.mapping[o_idx[i1 + k] + 1] = n_idx[j1 + k] + 1
        am.deleted.extend(o_idx[i] + 1 for i in range(i1 + n, i2))
        am.inserted.extend(n_idx[j] + 1 for j in range(j1 + n, j2))


def _align(old: tuple[list, list], new: tuple[list, list]) -> AxisMap:
    (old, o_labels), (new, n_labels) = old, new
    sm = SequenceMatcher(None, old, new, autojunk=False)
    am = AxisMap(identity=False)
    equal = 0
    structural = False
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            equal += i2 - i1
            for k in range(i2 - i1):
                am.mapping[i1 + k + 1] = j1 + k + 1
        else:
            _pair_block(am, range(i1, i2), range(j1, j2), o_labels, n_labels)
            if (i2 - i1) != (j2 - j1):
                structural = True
    # Only trust the alignment when there was a real insertion/deletion and the
    # rest of the sheet lines up well; otherwise compare positionally.
    if not structural or equal < 0.5 * min(len(old), len(new)):
        return AxisMap()
    if all(k == v for k, v in am.mapping.items()):
        return AxisMap()  # only appended/trailing rows: nothing shifted
    return am


def align_sheets(old: Sheet, new: Sheet) -> tuple[AxisMap, AxisMap]:
    rows_o, rows_n = old.max_row, new.max_row
    cols_o, cols_n = old.max_col, new.max_col
    if max(rows_o, rows_n) > _MAX_AXIS or max(cols_o, cols_n) > _MAX_AXIS:
        return AxisMap(), AxisMap()
    rmap = _align(_signatures(old, 0, rows_o), _signatures(new, 0, rows_n))
    cmap = _align(_signatures(old, 1, cols_o), _signatures(new, 1, cols_n))
    return rmap, cmap
