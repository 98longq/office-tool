"""OfficeTool package."""

from .config import OfficeToolConfig
from .audit import OfficialDocumentAuditor
from .formatter import OfficialDocumentFormatter

__all__ = ["OfficeToolConfig", "OfficialDocumentAuditor", "OfficialDocumentFormatter"]

__version__ = "0.1.0"
