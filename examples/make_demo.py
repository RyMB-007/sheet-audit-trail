"""Build a small DCF valuation model in two versions to demo the audit diff.

    python examples/make_demo.py examples/demo

v2 contains the kinds of edits that break real models:
  * a row inserted into the DCF (Capex), shifting every formula below it
  * the WACC formula overwritten with a hardcoded number
  * revenue growth typed as 8 instead of 8% (unit error)
  * one year's EBITDA formula replaced with a different pattern + magic number
  * ordinary input changes (tax rate, net debt)

If LibreOffice is installed the files are recalculated so they carry cached
values, exactly like a file saved from Excel.
"""
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName

YEARS = ["B", "C", "D", "E", "F"]


def build(v2: bool) -> Workbook:
    wb = Workbook()
    a = wb.active
    a.title = "Assumptions"
    rows = [
        ("Assumptions", None),
        (None, None),
        ("Revenue growth", 8 if v2 else 0.08),
        ("EBITDA margin", 0.25),
        ("Tax rate", 0.25 if v2 else 0.21),
        ("Cost of equity", 0.11),
        ("Cost of debt", 0.06),
        ("Equity weight", 0.7),
        ("WACC", 0.085 if v2 else "=B6*B8+B7*(1-B8)*(1-B5)"),
        ("Terminal growth", 0.025),
        (None, None),
        ("Base revenue", 1_000_000),
    ]
    for i, (label, val) in enumerate(rows, start=1):
        a.cell(i, 1, label)
        if val is not None:
            a.cell(i, 2, val)
    wb.defined_names["WACC"] = DefinedName("WACC", attr_text="Assumptions!$B$9")
    wb.defined_names["TGR"] = DefinedName("TGR", attr_text="Assumptions!$B$10")

    d = wb.create_sheet("DCF")
    d["A1"] = "DCF"
    for i, col in enumerate(YEARS):
        d[f"{col}2"] = 2026 + i
    lines = [("Revenue", None), ("EBITDA", "={c}3*Assumptions!$B$4"), ("Taxes", "=-{c}4*Assumptions!$B$5")]
    if v2:
        lines.append(("Capex", "=-{c}3*0.04"))
    fcf_row = 3 + len(lines)
    fcf = "={c}4+{c}5+{c}6" if v2 else "={c}4+{c}5"
    lines += [("Free cash flow", fcf), ("Discount factor", "=1/(1+WACC)^({c}2-2025)"),
              ("PV of FCF", f"={{c}}{fcf_row}*{{c}}{fcf_row + 1}")]
    for r, (label, tmpl) in enumerate(lines, start=3):
        d.cell(r, 1, label)
        for i, col in enumerate(YEARS):
            if label == "Revenue":
                prev = "Assumptions!B12" if i == 0 else f"{YEARS[i - 1]}3"
                d[f"{col}{r}"] = f"={prev}*(1+Assumptions!$B$3)"
            else:
                d[f"{col}{r}"] = tmpl.format(c=col)
    if v2:
        d["E4"] = "=E3*0.3"  # one-off override that breaks the row pattern
    pv_row, df_row = fcf_row + 2, fcf_row + 1
    s = pv_row + 2
    d.cell(s, 1, "Sum of PV"); d.cell(s, 2, f"=SUM(B{pv_row}:F{pv_row})")
    d.cell(s + 1, 1, "Terminal value"); d.cell(s + 1, 2, f"=F{fcf_row}*(1+TGR)/(WACC-TGR)")
    d.cell(s + 2, 1, "PV of terminal value"); d.cell(s + 2, 2, f"=B{s + 1}*F{df_row}")
    d.cell(s + 3, 1, "Enterprise value"); d.cell(s + 3, 2, f"=B{s}+B{s + 2}")

    m = wb.create_sheet("Summary")
    m["A1"] = "Summary"
    m["A3"], m["B3"] = "Enterprise value", f"=DCF!B{s + 3}"
    m["A4"], m["B4"] = "Net debt", 175_000 if v2 else 150_000
    m["A5"], m["B5"] = "Equity value", "=B3-B4"
    return wb


def recalc(path: Path) -> None:
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        print("LibreOffice not found: files will have no cached values")
        return
    with tempfile.TemporaryDirectory() as tmp:
        subprocess.run([soffice, "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML",
                        "--outdir", tmp, str(path)], check=True, capture_output=True)
        shutil.move(str(Path(tmp) / path.name), path)


if __name__ == "__main__":
    out = Path(sys.argv[1] if len(sys.argv) > 1 else "examples/demo")
    out.mkdir(parents=True, exist_ok=True)
    for name, v2 in (("valuation_v1.xlsx", False), ("valuation_v2.xlsx", True)):
        p = out / name
        build(v2).save(p)
        recalc(p)
        print("wrote", p)
