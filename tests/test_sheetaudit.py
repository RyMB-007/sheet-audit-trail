import subprocess
from pathlib import Path

import openpyxl
import pytest
from openpyxl.workbook.defined_name import DefinedName

from sheetaudit.cli import main
from sheetaudit.diff import Kind, diff_workbooks
from sheetaudit.model import Cell, Sheet, Workbook, load_workbook
from sheetaudit.refs import numeric_constants, parse_refs, remap_formula, to_r1c1
from sheetaudit.report import GITHUB_COMMENT_LIMIT, render, render_fitting, to_json

# --------------------------------------------------------------------------- helpers


def wb_from(sheets: dict, names=None, cached=False) -> Workbook:
    """Build a model directly: {"Sheet": {"A1": value_or_formula or (formula, cached)}}."""
    from openpyxl.utils.cell import coordinate_from_string, column_index_from_string

    wb = Workbook(names=names or {}, has_cached_values=cached)
    for name, cells in sheets.items():
        s = Sheet(name)
        for a1, v in cells.items():
            col, row = coordinate_from_string(a1)
            key = (row, column_index_from_string(col))
            if isinstance(v, tuple):
                s.cells[key] = Cell(v[0], v[1])
            elif isinstance(v, str) and v.startswith("="):
                s.cells[key] = Cell(v, None)
            else:
                s.cells[key] = Cell(None, v)
        wb.sheets[name] = s
    return wb


def kinds(d):
    return [ch.kind for s in d.sheets for ch in s.changes]


def titles(d):
    return [f.title for f in d.flags]


# --------------------------------------------------------------------------- refs


def test_parse_refs_handles_sheets_ranges_and_strings():
    refs = parse_refs("=SUM('My Sheet'!$A$1:B3)+Data!C5+LOG10(D2)+\"E5\"&A:A")
    got = [(r.sheet, r.r1, r.c1, r.r2, r.c2) for r in refs]
    assert ("My Sheet", 1, 1, 3, 2) in got
    assert ("Data", 5, 3, 5, 3) in got
    assert (None, 2, 4, 2, 4) in got  # D2 inside LOG10(...) but not LOG10 itself
    assert all(r.r1 != 5 or r.sheet for r in refs)  # "E5" in a string is ignored
    assert any(r.c1 == 1 and r.c2 == 1 and r.r2 == 1_048_576 for r in refs)  # A:A


def test_external_reference_detected():
    (ref,) = parse_refs("=[1]Prices!B2")
    assert ref.external and ref.sheet == "Prices"


def test_r1c1_is_position_independent():
    assert to_r1c1("=B3*$C$1", 3, 4) == to_r1c1("=C4*$C$1", 4, 5)
    assert to_r1c1("=B3*2", 3, 4) != to_r1c1("=B3*3", 3, 4)


def test_numeric_constants_ignores_refs_and_strings():
    assert numeric_constants('=B4*1.05+SUM(A1:A3)+"10"') == ["1.05"]
    assert numeric_constants("=Revenue2024*2") == ["2"]


def test_remap_formula_shifts_rows():
    mapper = lambda s, r, c: ((r + 1 if r and r >= 6 else r), c)  # noqa: E731
    assert remap_formula("=B5+B6+SUM(B6:B9)", mapper) == "=B5+B7+SUM(B7:B10)"
    assert remap_formula("='Old'!A1", lambda s, r, c: (r, c), {"old": "New Name"}) == "='New Name'!A1"


# --------------------------------------------------------------------------- diff


def test_formula_overwritten_with_hardcode_is_high_risk():
    old = wb_from({"S": {"A1": "WACC", "B1": "=B2*B3", "B2": 2, "B3": 3}})
    new = wb_from({"S": {"A1": "WACC", "B1": 6, "B2": 2, "B3": 3}})
    d = diff_workbooks(old, new, "m.xlsx")
    assert kinds(d) == [Kind.FORMULA_TO_VALUE]
    assert d.flags[0].severity == "high"
    assert d.sheets[0].changes[0].label == "WACC"


def test_unit_error_and_magic_number_flags():
    old = wb_from({"S": {"B1": 0.08, "B2": "=B1*A2"}})
    new = wb_from({"S": {"B1": 8, "B2": "=B1*1.05"}})
    t = titles(diff_workbooks(old, new, "m.xlsx"))
    assert "Order-of-magnitude change to input" in t
    assert "Hardcoded number inside formula" in t


def test_inconsistent_formula_in_row():
    row = {"A1": 1, "B1": 2, "C1": 3, "D1": 4, "A2": "=A1*2", "B2": "=B1*2", "C2": "=C1*2", "D2": "=D1*2"}
    new_row = dict(row, C2="=C1*2+A1")
    d = diff_workbooks(wb_from({"S": row}), wb_from({"S": new_row}), "m.xlsx")
    assert "Inconsistent formula" in titles(d)


def test_row_insertion_does_not_flag_shifted_formulas():
    old = {"A1": "Rev", "B1": 100, "A2": "Cost", "B2": 40, "A3": "Profit", "B3": "=B1-B2",
           "A4": "Margin", "B4": "=B3/B1"}
    new = {"A1": "Rev", "B1": 100, "A2": "Cost", "B2": 40, "A3": "Other", "B3": 5,
           "A4": "Profit", "B4": "=B1-B2", "A5": "Margin", "B5": "=B4/B1"}
    d = diff_workbooks(wb_from({"P&L": old}), wb_from({"P&L": new}), "m.xlsx")
    sd = d.sheets[0]
    assert sd.rows_inserted == [3]
    assert all(ch.kind == Kind.ADDED for ch in sd.changes), sd.changes


def test_sheet_rename_detected_and_refs_follow():
    a = {"A1": "x", "B1": 1, "B2": 2, "B3": 3, "B4": 4}
    old = wb_from({"Inputs": a, "Calc": {"A1": "=Inputs!B1*2"}})
    new = wb_from({"Drivers": a, "Calc": {"A1": "=Drivers!B1*2"}})
    d = diff_workbooks(old, new, "m.xlsx")
    assert {s.name: s.status for s in d.sheets} == {"Drivers": "renamed", "Calc": "modified"}
    assert d.total_changes == 0


def test_named_range_redefinition_flagged():
    old = wb_from({"S": {"B1": 1}}, names={"WACC": "S!$B$1"})
    new = wb_from({"S": {"B1": 1}}, names={"WACC": "S!$B$2"})
    assert "Named range redefined" in titles(diff_workbooks(old, new, "m.xlsx"))


def test_recalculated_values_and_new_errors():
    old = wb_from({"S": {"A1": 10, "A2": 0.5, "B1": ("=A1/A2", 20), "C1": ("=B1*2", 40)}}, cached=True)
    new = wb_from({"S": {"A1": 10, "A2": 0, "B1": ("=A1/A2", "#DIV/0!"), "C1": ("=B1*2", "#DIV/0!")}}, cached=True)
    d = diff_workbooks(old, new, "m.xlsx")
    assert kinds(d) == [Kind.INPUT_CHANGED]
    assert "Formula now evaluates to #DIV/0!" in titles(d)
    assert d.downstream_count == 2
    outputs = [rc for rc in d.key_outputs if rc.is_output]
    assert [(rc.sheet, rc.pos) for rc in outputs] == [("S", (1, 3))]


def test_copied_flags_are_merged_into_a_range():
    old = wb_from({"S": {"B1": 1, "C1": 1, "D1": 1}})
    new = wb_from({"S": {"B1": 1, "C1": 1, "D1": 1, "B2": "=B1*0.04", "C2": "=C1*0.04", "D2": "=D1*0.04"}})
    d = diff_workbooks(old, new, "m.xlsx")
    (flag,) = [f for f in d.flags if f.title == "Hardcoded number inside formula"]
    assert flag.address == "B2:D2"


def test_added_and_removed_workbooks():
    wb = wb_from({"S": {"A1": 1}})
    assert diff_workbooks(None, wb, "x.xlsx").status == "added"
    assert diff_workbooks(wb, None, "x.xlsx").status == "removed"


# --------------------------------------------------------------------------- files / report / git


def _save(path: Path, cells: dict, names=None, hidden=False):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Model"
    for a1, v in cells.items():
        ws[a1] = v
    if hidden:
        wb.create_sheet("Secret").sheet_state = "hidden"
    for n, t in (names or {}).items():
        wb.defined_names[n] = DefinedName(n, attr_text=t)
    wb.save(path)


def test_load_xlsx_roundtrip(tmp_path):
    p = tmp_path / "a.xlsx"
    _save(p, {"A1": "Rate", "B1": 0.05, "B2": "=B1*Rate"}, names={"Rate": "Model!$B$1"}, hidden=True)
    wb = load_workbook(p)
    assert wb.sheets["Model"].cells[(2, 2)].formula == "=B1*Rate"
    assert wb.names == {"RATE": "Model!$B$1"}
    assert wb.sheets["Secret"].state == "hidden"


def test_cli_diff_and_report_size(tmp_path, capsys):
    a, b = tmp_path / "a.xlsx", tmp_path / "b.xlsx"
    _save(a, {f"A{i}": i for i in range(1, 400)})
    _save(b, {f"A{i}": i * 3 for i in range(1, 400)})
    assert main(["diff", str(a), str(b), "--fail-on", "medium"]) == 0  # plain input edits raise no flags
    out = capsys.readouterr().out
    assert "Spreadsheet Audit Trail" in out
    d = diff_workbooks(load_workbook(a), load_workbook(b), "b.xlsx")
    assert len(render_fitting([d], limit=5_000)) <= 5_000
    assert len(render([d])) > 5_000
    assert to_json([d]).startswith("[")


def test_git_mode(tmp_path, monkeypatch, capsys):
    def git(*args):
        subprocess.run(["git", *args], cwd=tmp_path, check=True, capture_output=True)

    git("init", "-q", "-b", "main")
    git("config", "user.email", "t@example.com")
    git("config", "user.name", "t")
    (tmp_path / "models").mkdir()
    _save(tmp_path / "models" / "dcf.xlsx", {"A1": "Growth", "B1": 0.03, "B2": "=B1*2"})
    (tmp_path / "notes.txt").write_text("hi")
    git("add", ".")
    git("commit", "-qm", "v1")
    git("checkout", "-qb", "feature")
    _save(tmp_path / "models" / "dcf.xlsx", {"A1": "Growth", "B1": 0.03, "B2": 0.06})
    git("commit", "-qam", "v2")
    monkeypatch.chdir(tmp_path)
    rc = main(["git", "--base", "main", "--head", "feature", "--fail-on", "high", "--json", "r.json"])
    out = capsys.readouterr().out
    assert rc == 1
    assert "models/dcf.xlsx" in out and "Formula overwritten" in out
    assert "notes.txt" not in out


@pytest.mark.skipif(not Path("examples/demo/valuation_v1.xlsx").exists(), reason="demo files not built")
def test_demo_model():
    d = diff_workbooks(load_workbook("examples/demo/valuation_v1.xlsx"),
                       load_workbook("examples/demo/valuation_v2.xlsx"), "valuation.xlsx")
    dcf = next(s for s in d.sheets if s.name == "DCF")
    assert dcf.rows_inserted == [6]
    assert d.flag_counts()["high"] == 1
