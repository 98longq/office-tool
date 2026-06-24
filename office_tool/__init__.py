"""Source-checkout import bridge for the package stored under ``src``."""

from pathlib import Path

_SOURCE_PACKAGE = Path(__file__).resolve().parents[1] / "src" / "office_tool"
if _SOURCE_PACKAGE.is_dir():
    __path__.append(str(_SOURCE_PACKAGE))

from .audit import OfficialDocumentAuditor
from .config import OfficeToolConfig
from .formatter import OfficialDocumentFormatter

__all__ = ["OfficeToolConfig", "OfficialDocumentAuditor", "OfficialDocumentFormatter"]
__version__ = "0.1.0"
