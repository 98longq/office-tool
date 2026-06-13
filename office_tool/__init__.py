"""Source-tree launcher shim for OfficeTool.

This package makes ``python -m office_tool`` work before the project is
installed in editable mode. Real implementation modules live in
``src/office_tool``.
"""

from pathlib import Path

_SRC_PACKAGE = Path(__file__).resolve().parents[1] / "src" / "office_tool"
if _SRC_PACKAGE.exists():
    __path__.append(str(_SRC_PACKAGE))

from .config import OfficeToolConfig
from .audit import OfficialDocumentAuditor
from .formatter import OfficialDocumentFormatter

__all__ = ["OfficeToolConfig", "OfficialDocumentAuditor", "OfficialDocumentFormatter"]

__version__ = "0.1.0"
