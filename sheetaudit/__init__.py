"""Sheet Audit Trail: semantic, audit-focused diffs for Excel workbooks."""
__version__ = "1.0.0"

from .diff import diff_workbooks  # noqa: E402,F401
from .model import load_workbook  # noqa: E402,F401
