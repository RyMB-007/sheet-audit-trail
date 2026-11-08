"""Create or update a single sticky PR comment containing the report."""
from __future__ import annotations

import json
import os
import urllib.request
from typing import Optional

from .report import MARKER


def _request(method: str, url: str, token: str, body: Optional[dict] = None):
    data = json.dumps(body).encode() if body is not None else None
    req = urllib.request.Request(url, data=data, method=method, headers={
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
        "User-Agent": "sheet-audit-trail",
        **({"Content-Type": "application/json"} if data else {}),
    })
    with urllib.request.urlopen(req, timeout=30) as resp:
        payload = resp.read()
        return json.loads(payload) if payload else None


def upsert_comment(repo: str, pr: int, body: str, token: str, api: Optional[str] = None,
                   update_only: bool = False) -> str:
    api = (api or os.environ.get("GITHUB_API_URL") or "https://api.github.com").rstrip("/")
    page = 1
    existing = None
    while existing is None:
        comments = _request("GET", f"{api}/repos/{repo}/issues/{pr}/comments?per_page=100&page={page}", token)
        if not comments:
            break
        existing = next((c for c in comments if MARKER in (c.get("body") or "")), None)
        page += 1
    if existing:
        res = _request("PATCH", f"{api}/repos/{repo}/issues/comments/{existing['id']}", token, {"body": body})
        return f"updated {res['html_url']}"
    if update_only:
        return "skipped (no spreadsheet changes and no existing comment)"
    res = _request("POST", f"{api}/repos/{repo}/issues/{pr}/comments", token, {"body": body})
    return f"created {res['html_url']}"
