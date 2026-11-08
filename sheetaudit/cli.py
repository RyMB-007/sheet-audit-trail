"""Command-line interface: `sheetaudit diff|git|comment|textconv|gsheet-sync`."""
from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

from . import __version__
from .diff import SEVERITY_ORDER, WorkbookDiff, diff_workbooks
from .model import load_workbook
from .report import render, render_fitting, to_json


def _load_or_none(path: str):
    if path in ("-", "/dev/null", "") or not Path(path).exists():
        return None
    return load_workbook(path)


def _write_outputs(diffs: list[WorkbookDiff], args) -> str:
    footer = os.environ.get("SHEETAUDIT_FOOTER", "")
    full = render(diffs, args.max_rows, footer)
    comment = render_fitting(diffs, footer=footer) if args.fit else full
    if args.output:
        Path(args.output).write_text(comment)
    else:
        sys.stdout.write(comment)
    if args.full_output:
        Path(args.full_output).write_text(full)
    if args.json:
        Path(args.json).write_text(to_json(diffs))
    _github_outputs(diffs, args)
    return comment


def _github_outputs(diffs: list[WorkbookDiff], args):
    out = os.environ.get("GITHUB_OUTPUT")
    if not out:
        return
    high = sum(d.flag_counts()["high"] for d in diffs)
    medium = sum(d.flag_counts()["medium"] for d in diffs)
    low = sum(d.flag_counts()["low"] for d in diffs)
    with open(out, "a") as f:
        f.write(f"changed-files={sum(1 for d in diffs if d.has_changes)}\n")
        f.write(f"cell-changes={sum(d.total_changes for d in diffs)}\n")
        f.write(f"high={high}\nmedium={medium}\nlow={low}\n")


def _exit_code(diffs: list[WorkbookDiff], fail_on: str) -> int:
    if fail_on == "none":
        return 0
    threshold = SEVERITY_ORDER[fail_on]
    worst = [f for d in diffs for f in d.flags if SEVERITY_ORDER[f.severity] <= threshold]
    if any(d.error for d in diffs):
        return 1
    if worst:
        print(f"::error::{len(worst)} risk flag(s) at or above '{fail_on}' severity", file=sys.stderr)
        return 1
    return 0


def cmd_diff(args) -> int:
    diffs = [diff_workbooks(_load_or_none(args.old), _load_or_none(args.new), args.label or args.new)]
    _write_outputs(diffs, args)
    return _exit_code(diffs, args.fail_on)


def cmd_git(args) -> int:
    from .gitio import DEFAULT_PATTERNS, changed_files, merge_base, read_blob

    patterns = [p.strip() for p in (args.include or []) if p.strip()] or list(DEFAULT_PATTERNS)
    base = merge_base(args.base, args.head) if not args.two_dot else args.base
    diffs: list[WorkbookDiff] = []
    for f in changed_files(base, args.head, patterns):
        try:
            old = None if f.status == "A" else load_workbook(read_blob(base, f.old_path or f.path))
            new = None if f.status == "D" else load_workbook(read_blob(args.head, f.path))
            diffs.append(diff_workbooks(old, new, f.path, f.old_path))
        except Exception as e:  # keep going: one corrupt file shouldn't hide the others
            wd = WorkbookDiff(path=f.path, status="modified", error=f"{type(e).__name__}: {e}")
            diffs.append(wd)
            print(f"::warning::{f.path}: {e}", file=sys.stderr)
    _write_outputs(diffs, args)
    return _exit_code(diffs, args.fail_on)


def cmd_comment(args) -> int:
    from .github import upsert_comment

    token = args.token or os.environ.get("GITHUB_TOKEN")
    if not token:
        print("GITHUB_TOKEN is required", file=sys.stderr)
        return 2
    print(upsert_comment(args.repo, args.pr, Path(args.body_file).read_text(), token, update_only=args.update_only))
    return 0


def cmd_textconv(args) -> int:
    """Plain-text dump for `git diff` (configure as a textconv driver)."""
    from .refs import addr

    wb = load_workbook(args.file)
    for sname, sheet in wb.sheets.items():
        state = "" if sheet.state == "visible" else f" ({sheet.state})"
        print(f"=== Sheet: {sname}{state}")
        for (r, c) in sorted(sheet.cells):
            cell = sheet.cells[(r, c)]
            if cell.is_formula:
                print(f"{sname}!{addr(r, c)}\t{cell.formula}\t= {cell.value!r}")
            else:
                print(f"{sname}!{addr(r, c)}\t{cell.value!r}")
    for name, target in sorted(wb.names.items()):
        print(f"=== Name: {name} -> {target}")
    return 0


def cmd_gsheet_sync(args) -> int:
    from .gsheets import sync

    changed = sync(args.manifest)
    out = os.environ.get("GITHUB_OUTPUT")
    if out:
        with open(out, "a") as f:
            f.write(f"changed={'true' if changed else 'false'}\n")
    return 0


def _common(p: argparse.ArgumentParser):
    p.add_argument("-o", "--output", help="write Markdown here instead of stdout")
    p.add_argument("--full-output", help="also write the untruncated report here")
    p.add_argument("--json", help="write a machine-readable JSON report here")
    p.add_argument("--max-rows", type=int, default=100, help="max table rows per section (default 100)")
    p.add_argument("--fit", action="store_true", help="shrink the report to fit a GitHub comment")
    p.add_argument("--fail-on", choices=["none", "high", "medium", "low"], default="none",
                   help="exit 1 if any risk flag at/above this severity is raised")


def main(argv=None) -> int:
    p = argparse.ArgumentParser(prog="sheetaudit", description="Audit-grade diffs for Excel workbooks.")
    p.add_argument("--version", action="version", version=__version__)
    sub = p.add_subparsers(dest="cmd", required=True)

    d = sub.add_parser("diff", help="diff two workbook files")
    d.add_argument("old", help="old .xlsx (use /dev/null for none)")
    d.add_argument("new", help="new .xlsx (use /dev/null for none)")
    d.add_argument("--label", help="display name for the workbook")
    _common(d)
    d.set_defaults(fn=cmd_diff)

    g = sub.add_parser("git", help="diff all spreadsheets changed between two git revisions")
    g.add_argument("--base", required=True)
    g.add_argument("--head", default="HEAD")
    g.add_argument("--include", action="append", help="glob of files to audit (repeatable)")
    g.add_argument("--two-dot", action="store_true", help="compare base..head directly, not merge-base")
    _common(g)
    g.set_defaults(fn=cmd_git)

    c = sub.add_parser("comment", help="create/update the sticky PR comment")
    c.add_argument("--repo", required=True, help="owner/name")
    c.add_argument("--pr", required=True, type=int)
    c.add_argument("--body-file", required=True)
    c.add_argument("--token")
    c.add_argument("--update-only", action="store_true", help="only edit an existing comment; never create one")
    c.set_defaults(fn=cmd_comment)

    t = sub.add_parser("textconv", help="print a workbook as text (git textconv driver)")
    t.add_argument("file")
    t.set_defaults(fn=cmd_textconv)

    s = sub.add_parser("gsheet-sync", help="export Google Sheets listed in a manifest to .xlsx")
    s.add_argument("--manifest", default="gsheets.json")
    s.set_defaults(fn=cmd_gsheet_sync)

    args = p.parse_args(argv)
    return args.fn(args)


if __name__ == "__main__":
    sys.exit(main())
