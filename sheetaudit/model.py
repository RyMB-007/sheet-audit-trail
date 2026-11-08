"""In-memory workbook model extracted from .xlsx/.xlsm files."""
from __future__ import annotations

import datetime as _dt
import io
import os
import warnings
from dataclasses import dataclass, field
from typing import Any, Optional, Union

from .refs import ERROR_TOKENS


@dataclass(frozen=True)
class Cell:
    formula: Optional[str]  # "=SUM(A1:A3)" or None for constants
    value: Any  # constant value, or the cached (last calculated) result of a formula

    @property
    def is_formula(self) -> bool:
        return self.formula is not None

    @property
    def is_error(self) -> bool:
        return isinstance(self.value, str) and self.value in ERROR_TOKENS

    @property
    def kind(self) -> str:
        if self.is_formula:
            return "formula"
        v = self.value
        if isinstance(v, bool):
            return "bool"
        if isinstance(v, (int, float)):
            return "number"
        if self.is_error:
            return "error"
        if isinstance(v, str) and len(v) == 10 and v[4] == "-" and v[7] == "-":
            return "date"
        return "text"


@dataclass
class Sheet:
    name: str
    cells: dict[tuple[int, int], Cell] = field(default_factory=dict)
    state: str = "visible"  # visible | hidden | veryHidden

    @property
    def max_row(self) -> int:
        return max((r for r, _ in self.cells), default=0)

    @property
    def max_col(self) -> int:
        return max((c for _, c in self.cells), default=0)


@dataclass
class Workbook:
    sheets: dict[str, Sheet] = field(default_factory=dict)
    names: dict[str, str] = field(default_factory=dict)  # NAME -> "Sheet!$A$1"
    has_cached_values: bool = False

    def sheet(self, name: str) -> Optional[Sheet]:
        if name in self.sheets:
            return self.sheets[name]
        low = name.lower()
        return next((s for n, s in self.sheets.items() if n.lower() == low), None)


def _normalize_value(v: Any) -> Any:
    if isinstance(v, _dt.datetime):
        return v.date().isoformat() if v.time() == _dt.time() else v.isoformat(sep=" ")
    if isinstance(v, (_dt.date, _dt.time)):
        return v.isoformat()
    if isinstance(v, _dt.timedelta):
        return str(v)
    if isinstance(v, float) and v.is_integer() and abs(v) < 2**53:
        return int(v)
    return v


def _formula_text(v: Any) -> Optional[str]:
    if isinstance(v, str):
        return v if v.startswith("=") and len(v) > 1 else None
    text = getattr(v, "text", None)  # ArrayFormula
    if isinstance(text, str):
        return text if text.startswith("=") else "=" + text
    if type(v).__name__ == "DataTableFormula":
        return "=TABLE()"
    return None


Source = Union[str, os.PathLike, bytes]


def load_workbook(source: Source) -> Workbook:
    """Parse an .xlsx/.xlsm file (path or raw bytes) into a Workbook model."""
    import openpyxl

    def open_(data_only: bool):
        f = io.BytesIO(source) if isinstance(source, bytes) else source
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            return openpyxl.load_workbook(f, read_only=True, data_only=data_only, keep_links=False)

    wb_f = open_(False)
    wb_v = open_(True)
    out = Workbook()
    try:
        for ws_f, ws_v in zip(wb_f.worksheets, wb_v.worksheets):
            sheet = Sheet(ws_f.title, state=getattr(ws_f, "sheet_state", "visible") or "visible")
            cached: dict[tuple[int, int], Any] = {}
            for row in ws_v.iter_rows():
                for c in row:
                    if getattr(c, "value", None) is not None:
                        cached[(c.row, c.column)] = c.value
            for row in ws_f.iter_rows():
                for c in row:
                    v = getattr(c, "value", None)
                    if v is None:
                        continue
                    key = (c.row, c.column)
                    formula = _formula_text(v)
                    if formula is not None:
                        cv = cached.get(key)
                        if cv is not None:
                            out.has_cached_values = True
                        sheet.cells[key] = Cell(formula, _normalize_value(cv))
                    else:
                        sheet.cells[key] = Cell(None, _normalize_value(v))
            out.sheets[sheet.name] = sheet
        out.names = _defined_names(wb_f)
    finally:
        wb_f.close()
        wb_v.close()
    return out


def _defined_names(wb) -> dict[str, str]:
    names: dict[str, str] = {}
    try:
        items = list(wb.defined_names.items())  # openpyxl >= 3.1
    except AttributeError:  # pragma: no cover - older openpyxl
        items = [(d.name, d) for d in wb.defined_names.definedName]
    for name, dn in items:
        if name.startswith("_xlnm."):
            continue
        names[name.upper()] = str(getattr(dn, "attr_text", dn.value))
    for ws in wb.worksheets:
        local = getattr(ws, "defined_names", None)
        if local:
            for name, dn in local.items():
                if not name.startswith("_xlnm."):
                    names[f"{ws.title}!{name}".upper()] = str(dn.attr_text)
    return names
