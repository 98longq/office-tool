"""Conservative optional content generation for red-head documents."""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from docx.document import Document as DocxDocument
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

from .audit import OfficialDocumentAuditor
from .config import OfficeToolConfig


@dataclass
class GenerationResult:
    generated: list[str] = field(default_factory=list)
    skipped: list[str] = field(default_factory=list)

    def to_stats(self) -> dict[str, object]:
        return {
            "generated_content": list(self.generated),
            "skipped_generation": list(self.skipped),
        }


class OfficialDocumentContentGenerator:
    """Generate only wholly absent structures whose values were supplied by the user."""

    RED_HEAD_PROFILES = {"red_head", "letter_head", "meeting_minutes"}

    def __init__(self, config: OfficeToolConfig):
        self.config = config
        self.auditor = OfficialDocumentAuditor(config)

    def apply(self, doc: DocxDocument) -> GenerationResult:
        result = GenerationResult()
        options = self.config.generation
        profile = self.config.audit.profile
        if profile not in self.RED_HEAD_PROFILES:
            return result

        structure = self.auditor.detect_structure(self.auditor._paragraph_items(doc.paragraphs))
        if options.add_red_head:
            self._add_red_head(doc, profile, structure, result)

        # Re-detect because newly generated paragraphs change all indices.
        structure = self.auditor.detect_structure(self.auditor._paragraph_items(doc.paragraphs))
        if options.add_imprint and profile != "letter_head":
            self._add_imprint(doc, profile, structure, result)
        return result

    def _add_red_head(self, doc, profile, structure, result: GenerationResult) -> None:
        existing = [
            structure.internal_notice,
            structure.red_head,
            structure.document_number,
            structure.meeting_number,
            structure.meeting_issue_line,
        ]
        if any(index is not None for index in existing):
            result.skipped.append("red_head_existing_or_partial")
            return

        options = self.config.generation
        if profile == "meeting_minutes":
            self._require(
                ("会议期号", options.meeting_number),
                ("编发单位", options.meeting_organization),
                ("编发日期", options.meeting_date),
            )
            paragraphs = [
                "内部资料不得外传",
                "会议纪要",
                self._meeting_number(options.meeting_number),
                f"{options.meeting_organization}    {options.meeting_date}",
            ]
        else:
            self._require(("红头名称", options.red_head_title), ("发文字号", options.document_number))
            paragraphs = [options.red_head_title.strip(), options.document_number.strip()]
            if profile == "red_head":
                paragraphs.insert(0, "内部资料不得外传")

        self._prepend_paragraphs(doc, paragraphs)
        if profile == "letter_head":
            self._append_letter_notes(doc)
        result.generated.append("red_head")

    def _add_imprint(self, doc, profile, structure, result: GenerationResult) -> None:
        if profile == "meeting_minutes":
            if structure.distribution is not None:
                result.skipped.append("distribution_existing")
                return
            value = re.sub(r"^分送\s*[:：]\s*", "", self.config.generation.distribution.strip())
            self._require(("分送内容", value))
            doc.add_paragraph(f"分送：{value}")
            result.generated.append("distribution")
            return

        if structure.copy_to is not None or structure.print_org_date is not None or structure.simple_imprint is not None:
            result.skipped.append("imprint_existing_or_partial")
            return
        options = self.config.generation
        copy_to = re.sub(r"^(?:抄送|抄报|发送)\s*[:：]\s*", "", options.copy_to.strip())
        print_date = re.sub(r"印发\s*$", "", options.print_date.strip())
        self._require(
            ("印发单位", options.print_organization),
            ("印发日期", print_date),
        )
        if copy_to:
            doc.add_paragraph(f"抄送：{copy_to}")
            doc.add_paragraph(f"{options.print_organization.strip()}    {print_date}印发")
            result.generated.append("imprint")
        else:
            paragraph = doc.add_paragraph(f"{options.print_organization.strip()}    {print_date}")
            self._set_paragraph_style_id(paragraph, "OfficeToolGeneratedSimpleImprint")
            result.generated.append("simple_imprint")

    def _append_letter_notes(self, doc: DocxDocument) -> None:
        texts = [paragraph.text.strip() for paragraph in doc.paragraphs]
        if not any("内部资料" in text and "不得外传" in text for text in texts):
            doc.add_paragraph("（内部资料　　不得外传）")

    @staticmethod
    def _set_paragraph_style_id(paragraph: Paragraph, style_id: str) -> None:
        p_pr = paragraph._p.get_or_add_pPr()
        p_style = p_pr.get_or_add_pStyle()
        p_style.set(qn("w:val"), style_id)

    @staticmethod
    def _meeting_number(value: str) -> str:
        stripped = value.strip().strip("（）()")
        return f"（{stripped}）"

    @staticmethod
    def _require(*values: tuple[str, str]) -> None:
        missing = [label for label, value in values if not str(value).strip()]
        if missing:
            raise ValueError("添加内容前请填写：" + "、".join(missing))

    @staticmethod
    def _prepend_paragraphs(doc: DocxDocument, texts: list[str]) -> None:
        body = doc._element.body
        for index, text in enumerate(texts):
            element = OxmlElement("w:p")
            body.insert(index, element)
            Paragraph(element, doc._body).add_run(text)
