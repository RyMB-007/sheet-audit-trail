"""Find changed spreadsheet files between two git revisions and read their bytes."""
from __future__ import annotations

import fnmatch
import subprocess
from dataclasses import dataclass
from typing import Optional

DEFAULT_PATTERNS = ("*.xlsx", "*.xlsm")


@dataclass
class ChangedFile:
    status: str  # A, M, D, R
    path: str
    old_path: Optional[str] = None


def _git(*args: str, binary: bool = False):
    res = subprocess.run(["git", *args], capture_output=True, check=False)
    if res.returncode != 0:
        raise RuntimeError(f"git {' '.join(args)} failed: {res.stderr.decode(errors='replace').strip()}")
    return res.stdout if binary else res.stdout.decode()


def merge_base(base: str, head: str) -> str:
    try:
        return _git("merge-base", base, head).strip()
    except RuntimeError:
        return base  # unrelated histories / shallow clone: fall back to a two-dot diff


def _matches(path: str, patterns) -> bool:
    name = path.rsplit("/", 1)[-1]
    return any(fnmatch.fnmatch(path, p) or fnmatch.fnmatch(name, p) for p in patterns)


def changed_files(base: str, head: str, patterns=DEFAULT_PATTERNS) -> list[ChangedFile]:
    out = _git("diff", "--name-status", "-z", "-M", "--no-color", base, head)
    parts = out.split("\0")
    files: list[ChangedFile] = []
    i = 0
    while i < len(parts) - 1:
        status = parts[i]
        if not status:
            i += 1
            continue
        code = status[0]
        if code in "RC":
            old, new = parts[i + 1], parts[i + 2]
            i += 3
            if _matches(new, patterns) or _matches(old, patterns):
                files.append(ChangedFile("R" if code == "R" else "A", new, old if code == "R" else None))
        else:
            path = parts[i + 1]
            i += 2
            if _matches(path, patterns) and code in "AMDT":
                files.append(ChangedFile("M" if code == "T" else code, path))
    return files


LFS_POINTER = b"version https://git-lfs.github.com/spec"


def read_blob(rev: str, path: str) -> bytes:
    """Read a file at a revision, applying smudge filters (so Git LFS files resolve)."""
    try:
        data = _git("cat-file", "--filters", f"{rev}:{path}", binary=True)
    except RuntimeError:
        data = _git("show", f"{rev}:{path}", binary=True)
    if data.startswith(LFS_POINTER):
        raise RuntimeError(
            f"{path} is a Git LFS pointer at {rev[:12]}; check out with `lfs: true` "
            "(actions/checkout) so the real file can be read."
        )
    return data
