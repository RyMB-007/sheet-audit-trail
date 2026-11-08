"""Excel formula reference parsing.

A lightweight tokenizer that understands the parts of Excel formula syntax an
auditor cares about: cell/range references (with optional sheet and external
workbook prefixes), string literals, numeric literals and defined names.

It is intentionally regex-based rather than a full grammar: it is robust to
unknown functions and only needs to find references reliably.
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Callable, Iterator, Optional

MAX_ROW = 1_048_576
MAX_COL = 16_384

STRING_RE = re.compile(r'"(?:[^"]|"")*"')

_SHEET = r"(?:\[[^\]]+\])?(?:'(?:[^']|'')+'|[A-Za-z_À-￿][\w.À-￿]*)"
_COL = r"\$?[A-Z]{1,3}"
_ROW = r"\$?[0-9]{1,7}"

REF_RE = re.compile(
    rf"(?<![\w.$'\]À-￿])"
    rf"(?:(?P<sheet>{_SHEET})!)?"
    rf"(?:(?P<c1>{_COL})(?P<r1>{_ROW})(?::(?P<c2>{_COL})(?P<r2>{_ROW}))?"
    rf"|(?P<cc1>{_COL}):(?P<cc2>{_COL})"
    rf"|(?P<rr1>{_ROW}):(?P<rr2>{_ROW}))"
    rf"(?![\w(!])"
)

NUMBER_RE = re.compile(r"(?<![\w.$§])(\d+(?:\.\d+)?(?:[eE][+-]?\d+)?%?)(?![\w(.])")
NAME_RE = re.compile(r"(?<![\w.$§\\])([A-Za-z_\\][\w.]*)(?![\w(!])")
ERROR_TOKENS = ("#REF!", "#NAME?", "#VALUE!", "#DIV/0!", "#N/A", "#NUM!", "#NULL!")
_KEYWORDS = {"TRUE", "FALSE"}


def col_to_index(col: str) -> int:
    n = 0
    for ch in col.lstrip("$").upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def index_to_col(n: int) -> str:
    s = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return s


def addr(row: int, col: int) -> str:
    return f"{index_to_col(col)}{row}"


def unquote_sheet(raw: str) -> tuple[str, bool]:
    """Return (sheet name, is_external) from a raw sheet prefix."""
    s = raw
    if s.startswith("'") and s.endswith("'"):
        s = s[1:-1].replace("''", "'")
    external = False
    if s.startswith("["):
        external = True
        s = s[s.index("]") + 1:] if "]" in s else s
    return s, external


def quote_sheet(name: str) -> str:
    if re.fullmatch(r"[A-Za-z_][\w.]*", name) and not re.fullmatch(r"[A-Za-z]{1,3}\d+", name):
        return name
    return "'" + name.replace("'", "''") + "'"


@dataclass(frozen=True)
class Ref:
    sheet: Optional[str]  # None = same sheet as the formula
    external: bool
    r1: int
    c1: int
    r2: int
    c2: int

    @property
    def size(self) -> int:
        return (self.r2 - self.r1 + 1) * (self.c2 - self.c1 + 1)

    def contains(self, r: int, c: int) -> bool:
        return self.r1 <= r <= self.r2 and self.c1 <= c <= self.c2


def segments(formula: str) -> Iterator[tuple[bool, str]]:
    """Yield (is_string_literal, text) pieces of a formula."""
    pos = 0
    for m in STRING_RE.finditer(formula):
        if m.start() > pos:
            yield False, formula[pos:m.start()]
        yield True, m.group(0)
        pos = m.end()
    if pos < len(formula):
        yield False, formula[pos:]


def _sub_code(formula: str, fn: Callable[[str], str]) -> str:
    return "".join(text if is_str else fn(text) for is_str, text in segments(formula))


def _match_to_ref(m: re.Match) -> Ref:
    sheet, external = (None, False)
    if m.group("sheet"):
        sheet, external = unquote_sheet(m.group("sheet"))
    if m.group("c1"):
        r1, c1 = int(m.group("r1").lstrip("$")), col_to_index(m.group("c1"))
        if m.group("c2"):
            r2, c2 = int(m.group("r2").lstrip("$")), col_to_index(m.group("c2"))
        else:
            r2, c2 = r1, c1
    elif m.group("cc1"):
        c1, c2 = col_to_index(m.group("cc1")), col_to_index(m.group("cc2"))
        r1, r2 = 1, MAX_ROW
    else:
        r1, r2 = int(m.group("rr1").lstrip("$")), int(m.group("rr2").lstrip("$"))
        c1, c2 = 1, MAX_COL
    return Ref(sheet, external, min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2))


def parse_refs(formula: str) -> list[Ref]:
    refs: list[Ref] = []
    for is_str, text in segments(formula):
        if not is_str:
            refs.extend(_match_to_ref(m) for m in REF_RE.finditer(text))
    return refs


def _strip_refs(text: str) -> str:
    return REF_RE.sub("§", text)


def numeric_constants(formula: str) -> list[str]:
    """Numeric literals hardcoded inside a formula (e.g. the 1.05 in =B4*1.05)."""
    out: list[str] = []
    for is_str, text in segments(formula):
        if not is_str:
            out.extend(NUMBER_RE.findall(_strip_refs(text)))
    return out


def name_tokens(formula: str) -> set[str]:
    """Identifiers that might be defined names (upper-cased)."""
    out: set[str] = set()
    for is_str, text in segments(formula):
        if not is_str:
            for tok in NAME_RE.findall(_strip_refs(text)):
                if tok.upper() not in _KEYWORDS:
                    out.add(tok.upper())
    return out


def error_tokens(formula: str) -> set[str]:
    code = "".join(t for s, t in segments(formula) if not s)
    return {e for e in ERROR_TOKENS if e in code}


def has_external_link(formula: str) -> bool:
    return any(r.external for r in parse_refs(formula))


def _r1c1_part(cpart: str, rpart: str, row: int, col: int) -> str:
    out = ""
    if rpart:
        rv = int(rpart.lstrip("$"))
        out += f"R{rv}" if rpart.startswith("$") else f"R[{rv - row}]"
    if cpart:
        cv = col_to_index(cpart)
        out += f"C{cv}" if cpart.startswith("$") else f"C[{cv - col}]"
    return out


def to_r1c1(formula: str, row: int, col: int) -> str:
    """Position-independent form of a formula, used to detect copy-paste patterns."""

    def repl(m: re.Match) -> str:
        prefix = (m.group("sheet") + "!") if m.group("sheet") else ""
        if m.group("c1"):
            s = _r1c1_part(m.group("c1"), m.group("r1"), row, col)
            if m.group("c2"):
                s += ":" + _r1c1_part(m.group("c2"), m.group("r2"), row, col)
        elif m.group("cc1"):
            s = _r1c1_part(m.group("cc1"), "", row, col) + ":" + _r1c1_part(m.group("cc2"), "", row, col)
        else:
            s = _r1c1_part("", m.group("rr1"), row, col) + ":" + _r1c1_part("", m.group("rr2"), row, col)
        return prefix + s

    return _sub_code(formula, lambda t: REF_RE.sub(repl, t))


Mapper = Callable[[Optional[str], Optional[int], Optional[int]], tuple[Optional[int], Optional[int]]]


def remap_formula(formula: str, mapper: Mapper, rename: Optional[dict[str, str]] = None) -> str:
    """Rewrite references through a (sheet, row, col) -> (row, col) mapping.

    Used to translate an old formula into the coordinates of the new workbook
    after rows/columns were inserted or deleted, so that "=B5*2" which became
    "=B6*2" purely because a row was inserted is *not* reported as a change.
    """

    def fix_row(sheet, rpart):
        if not rpart:
            return rpart
        new, _ = mapper(sheet, int(rpart.lstrip("$")), None)
        return ("$" if rpart.startswith("$") else "") + str(new)

    def fix_col(sheet, cpart):
        if not cpart:
            return cpart
        _, new = mapper(sheet, None, col_to_index(cpart))
        return ("$" if cpart.startswith("$") else "") + index_to_col(new)

    def repl(m: re.Match) -> str:
        raw_sheet = m.group("sheet")
        sheet = None
        if raw_sheet:
            sheet, external = unquote_sheet(raw_sheet)
            if external:
                return m.group(0)
        prefix = (raw_sheet + "!") if raw_sheet else ""
        if sheet and rename and sheet.lower() in rename:
            prefix = quote_sheet(rename[sheet.lower()]) + "!"
        if m.group("c1"):
            s = fix_col(sheet, m.group("c1")) + fix_row(sheet, m.group("r1"))
            if m.group("c2"):
                s += ":" + fix_col(sheet, m.group("c2")) + fix_row(sheet, m.group("r2"))
        elif m.group("cc1"):
            s = fix_col(sheet, m.group("cc1")) + ":" + fix_col(sheet, m.group("cc2"))
        else:
            s = fix_row(sheet, m.group("rr1")) + ":" + fix_row(sheet, m.group("rr2"))
        return prefix + s

    return _sub_code(formula, lambda t: REF_RE.sub(repl, t))
