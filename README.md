# 📊 Sheet Audit Trail

[![GitHub Marketplace](https://img.shields.io/badge/Marketplace-Spreadsheet%20Audit%20Trail-blue?logo=github)](https://github.com/marketplace/actions/spreadsheet-audit-trail)
[![CI](https://github.com/RyMB-007/sheet-audit-trail/actions/workflows/ci.yml/badge.svg)](https://github.com/RyMB-007/sheet-audit-trail/actions/workflows/ci.yml)

**Git diffs for spreadsheets.** A GitHub Action (and CLI) that parses Excel workbooks
whenever they change in a pull request and posts an audit-ready Markdown diff of
formulas, hardcoded inputs and cell dependencies, instead of
`Binary files differ`.

Built for finance teams whose valuation models, forecasts and board packs live in Excel
or Google Sheets, where finding *which cell changed* after a model breaks is a forensic
exercise.

## What the PR comment shows

| Section | What it tells the reviewer |
|---|---|
| **Risk flags** 🚩 | The edits that break models (see below), ranked by severity |
| **Impact on calculated values** 🎯 | Cells whose formula did *not* change but whose result did, e.g. *Equity value 3.71M → 4.02M (+8.4%)*. Final outputs are starred |
| **Dependency impact** 🕸️ | Every cell downstream of the edits, traced through cross-sheet references and named ranges |
| **Changes by sheet** 🔍 | Each changed cell with its business label (*"Free cash flow · 2028"*), before/after formula **and** value. An edit copied across a row is collapsed into one line (`B7:F7 · 5 cells`) |
| **Structure** ↕️ | Inserted/deleted rows and columns, added/deleted/renamed/hidden sheets, named-range edits |

See [`examples/demo/sample-report.md`](examples/demo/sample-report.md) for a full report on the
included DCF model.

### Risk flags

| Severity | Flag |
|---|---|
| 🔴 high | **Formula overwritten with a hardcoded value**, the classic "plug" |
| 🔴 high | Formula now evaluates to an error (`#REF!`, `#DIV/0!`, …) or contains `#REF!` |
| 🟠 medium | **Hardcoded number inside a formula** (`=E3*0.3` instead of referencing an assumption) |
| 🟠 medium | **Inconsistent formula**: the cell breaks the pattern its row/column neighbours share |
| 🟠 medium | **Order-of-magnitude change** to an input (`0.08 → 8`: a likely % / unit error) |
| 🟠 medium | New link to an external workbook · named range redefined/deleted · sheet deleted or made *very hidden* |
| 🟡 low | Sheet added · sheet hidden |

### Smart enough not to cry wolf

* **Row/column insertion detection.** Inserting one row in a model shifts every formula
  below it. A naive diff reports hundreds of "changes". Sheet Audit Trail aligns rows the way
  `git diff` aligns lines: it compares position-independent (R1C1) formula signatures and row
  labels. It then translates old formulas into new coordinates, so only real edits are
  reported.
* **Sheet renames** are detected, and references to the renamed sheet aren't flagged.
* **Merged flags.** The same issue copied across `B6:F6` is one flag, not five.
* **Fits in a PR comment.** The comment shrinks to GitHub's 65k-character limit; the full
  report goes to the job summary, and a JSON report is written for downstream tooling.

## Quick start (GitHub Action)

```yaml
# .github/workflows/spreadsheet-audit.yml
name: Spreadsheet audit
on:
  pull_request:
    paths: ["**/*.xlsx", "**/*.xlsm"]
permissions:
  contents: read
  pull-requests: write
jobs:
  audit:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v7
        with:
          fetch-depth: 0   # compare against the PR base
          lfs: true        # if workbooks are in Git LFS
      - uses: RyMB-007/sheet-audit-trail@v1
        with:
          fail-on: high    # optional: block merges with high-risk edits
```

### Inputs

| Input | Default | |
|---|---|---|
| `include` | `*.xlsx` `*.xlsm` | Newline-separated globs of workbooks to audit |
| `fail-on` | `none` | Fail the check on flags at/above `low` / `medium` / `high` |
| `comment` | `true` | Post/update one sticky PR comment |
| `base`, `head` | PR base/head (or push before/after) | Revisions to compare |
| `github-token` | `github.token` | Token for commenting |

Outputs: `changed-files`, `cell-changes`, `high`, `medium`, `low`, `report-path`, `json-path`.

## Google Sheets

Google Sheets don't live in git, so the tool snapshots them as `.xlsx`:

1. List the sheets in a manifest, `gsheets.json`:
   `{"sheets": [{"id": "<id from the sheet URL>", "path": "models/forecast.xlsx"}]}`
2. Share each sheet (Viewer) with a Google Cloud service account and store its key as the
   secret `GOOGLE_SERVICE_ACCOUNT_JSON`.
3. Use [`examples/workflows/google-sheets-sync.yml`](examples/workflows/google-sheets-sync.yml):
   on a schedule it exports each sheet. If the *content* changed (re-exports with identical
   content are ignored), it opens a PR and posts the audit diff on it.

## CLI

```bash
pip install .                      # or: pip install ".[gsheets]"

sheetaudit diff old.xlsx new.xlsx              # Markdown to stdout
sheetaudit diff old.xlsx new.xlsx --json r.json --fail-on high
sheetaudit git --base main --head HEAD         # every workbook changed on this branch
sheetaudit textconv model.xlsx                 # plain-text dump for `git diff`
sheetaudit gsheet-sync --manifest gsheets.json
```

**Readable local `git diff`:** add `*.xlsx diff=sheetaudit` to `.gitattributes`, then run
`git config diff.sheetaudit.textconv "sheetaudit textconv"`.

## How it works

```
 .xlsx (base) ─┐                ┌─ formulas + cached values + names ─┐
               ├─ openpyxl ─────┤                                    ├─ align rows/cols ─ cell diff ─┐
 .xlsx (head) ─┘                └─ formulas + cached values + names ─┘                               │
                                                                                                     ▼
              PR comment / job summary / JSON  ◀── Markdown report ◀── risk rules + dependency graph
```

* `sheetaudit/refs.py`: formula tokenizer: references (quoted sheets, ranges, whole
  columns, external `[1]Book` links), numeric literals, R1C1 conversion, reference remapping
* `sheetaudit/model.py`: loads formulas *and* last-calculated values (Excel stores both)
* `sheetaudit/align.py`: row/column insertion and deletion detection
* `sheetaudit/diff.py`: cell classification and risk rules
* `sheetaudit/deps.py`: reverse dependency graph (cells, ranges, named ranges, cross-sheet)
* `sheetaudit/report.py`: Markdown/JSON rendering

## Limitations

* Value-impact analysis uses the **cached results saved in the file**. Files saved by Excel,
  Google Sheets or LibreOffice have them. Files generated by scripts (e.g. openpyxl) usually
  don't; the dependency graph still works for those.
* The tool doesn't evaluate formulas itself, and doesn't diff VBA macros, charts,
  formatting or comments.
* `INDIRECT`/`OFFSET` targets are dynamic, so they aren't traced in the dependency graph.
* Row and column insertions in the *same* sheet in one PR fall back to a positional diff
  for that sheet.

## Development

```bash
pip install -e ".[test]"
python examples/make_demo.py examples/demo   # builds the demo models (uses LibreOffice if present)
pytest -q
```
