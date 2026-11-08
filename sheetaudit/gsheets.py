"""Google Sheets support: snapshot sheets into the repo as .xlsx files.

Google Sheets live outside git, so the workflow is:

1. A manifest (``gsheets.json``) maps Google Sheet IDs to paths in the repo.
2. ``sheetaudit gsheet-sync`` exports each sheet to .xlsx via the Drive API
   (using a service account) and only rewrites a file when its *content*
   changed, so re-exports don't create noise.
3. The snapshot is committed on a branch and a PR is opened; the normal
   audit action then posts the diff on that PR.

Requires the optional dependencies: ``pip install "sheet-audit-trail[gsheets]"``.
"""
from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Optional

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]


def _session(credentials_json: Optional[str] = None):
    try:
        from google.auth.transport.requests import AuthorizedSession
        from google.oauth2 import service_account
    except ImportError as e:  # pragma: no cover
        raise SystemExit('Google Sheets support needs: pip install "sheet-audit-trail[gsheets]"') from e
    raw = credentials_json or os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if not raw:
        raise SystemExit("Set GOOGLE_SERVICE_ACCOUNT_JSON to a service-account key (JSON text or file path).")
    info = json.loads(Path(raw).read_text() if os.path.isfile(raw) else raw)
    creds = service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
    return AuthorizedSession(creds)


def export_xlsx(sheet_id: str, session=None) -> bytes:
    session = session or _session()
    url = f"https://www.googleapis.com/drive/v3/files/{sheet_id}/export"
    resp = session.get(url, params={"mimeType": XLSX_MIME, "supportsAllDrives": "true"}, timeout=120)
    if resp.status_code != 200:
        raise RuntimeError(f"Export of {sheet_id} failed ({resp.status_code}): {resp.text[:300]}. "
                           "Is the sheet shared with the service account's email?")
    return resp.content


def sync(manifest_path: str) -> list[str]:
    """Export every sheet in the manifest; return the paths whose content changed."""
    from .diff import diff_workbooks
    from .model import load_workbook

    manifest = json.loads(Path(manifest_path).read_text())
    session = _session()
    changed = []
    for entry in manifest.get("sheets", []):
        target = Path(entry["path"])
        data = export_xlsx(entry["id"], session)
        if target.exists():
            d = diff_workbooks(load_workbook(target.read_bytes()), load_workbook(data), str(target))
            if not d.has_changes:
                print(f"unchanged: {target}")
                continue
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(data)
        changed.append(str(target))
        print(f"updated:   {target}")
    return changed
