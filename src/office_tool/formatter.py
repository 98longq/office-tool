"""Official document formatting engine."""

from __future__ import annotations

from pathlib import Path

from docx.document import Document as DocxDocument
from docx.shared import Cm
from docx.text.paragraph import Paragraph

from .audit import OfficialDocumentAuditor
from .config import OfficeToolConfig
from .docx_utils import add_bottom_border, add_page_number, apply_style_to_paragraph
from .io import load_document
from .models import AuditReport
from .patterns import heading_role


class OfficialDocumentFormatter:
    """Audit and format documents using configurable official-document rules."""

    def __init__(self, config: OfficeToolConfig | None = None):
        self.config = config or OfficeToolConfig()
        self.auditor = OfficialDocumentAuditor(self.config)

    def format_path(self, input_path: str | Path, output_path: str | Path) -> AuditReport:
        doc, _kind = load_document(input_path)
        report = self.format_document(doc)
        output = Path(output_path).expanduser().resolve()
        output.parent.mkdir(parents=True, exist_ok=True)
        doc.save(output)
        return report

    def format_document(self, doc: DocxDocument) -> AuditReport:
        report = self.auditor.audit_document(doc)
        if self.config.format.apply_page_setup:
            self._apply_page_setup(doc)
        if self.config.format.apply_styles:
            self._apply_styles(doc, report)
        if self.config.format.add_page_number:
            self._apply_page_number(doc)
        return report

    def _apply_page_setup(self, doc: DocxDocument) -> None:
        page = self.config.page
        for section in doc.sections:
            section.page_width = Cm(page.paper_width_cm)
            section.page_height = Cm(page.paper_height_cm)
            section.top_margin = Cm(page.margin_top_cm)
            section.bottom_margin = Cm(page.margin_bottom_cm)
            section.left_margin = Cm(page.margin_left_cm)
            section.right_margin = Cm(page.margin_right_cm)
            section.footer_distance = Cm(page.footer_distance_cm)

    def _apply_styles(self, doc: DocxDocument, report: AuditReport) -> None:
        role_by_index = self._roles_from_report(report)
        red_head_doc_number = None
        if report.is_red_head:
            red_head_doc_number = self._first_index_for(role_by_index, "document_number")

        for index, paragraph in enumerate(doc.paragraphs):
            if not paragraph.text.strip():
                continue
            role = role_by_index.get(index) or heading_role(paragraph.text) or "body"
            style_name = self._style_for_role(role)
            style = self.config.styles[style_name]
            apply_style_to_paragraph(
                paragraph,
                style,
                preserve_bold_italic=self.config.format.preserve_existing_bold_italic,
            )
            if (
                self.config.format.draw_red_separator
                and report.is_red_head
                and role == "document_number"
                and index == red_head_doc_number
            ):
                add_bottom_border(paragraph, self.config.styles["red_head"].color or "FF0000")

    @staticmethod
    def _roles_from_report(report: AuditReport) -> dict[int, str]:
        roles: dict[int, str] = {}
        for element in report.elements:
            roles[element.block_index] = element.name
        return roles

    @staticmethod
    def _first_index_for(role_by_index: dict[int, str], role: str) -> int | None:
        for index, value in sorted(role_by_index.items()):
            if value == role:
                return index
        return None

    @staticmethod
    def _style_for_role(role: str) -> str:
        mapping = {
            "red_head": "red_head",
            "document_number": "document_number",
            "signer": "signer",
            "title": "title",
            "main_send": "main_send",
            "body_start": "body",
            "attachment_note": "attachment",
            "signatory": "signatory",
            "date": "date",
            "copy_to": "copy_to",
            "print_org_date": "copy_to",
        }
        return mapping.get(role, role if role in {"h1", "h2", "h3", "h4"} else "body")

    def _apply_page_number(self, doc: DocxDocument) -> None:
        style = self.config.styles["page_number"]
        for section in doc.sections:
            footer = section.footer
            paragraph: Paragraph = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
            add_page_number(paragraph, style.font, style.size_pt)
