"""Official document structure audit."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable

from docx.document import Document as DocxDocument
from docx.shared import Length
from docx.text.paragraph import Paragraph

from .config import OfficeToolConfig
from .models import AuditReport, DetectedStructure
from .patterns import (
    RE_ATTACHMENT_MARK,
    RE_ATTACHMENT_NOTE,
    RE_COPY_NUMBER,
    RE_COPY_TO,
    RE_DATE,
    RE_DOCUMENT_NUMBER,
    RE_EMPTY,
    RE_MAIN_SEND,
    RE_OBSOLETE_SUBJECT,
    RE_PRINT_ORG_DATE,
    RE_RED_HEAD_KEYWORDS,
    RE_SECRECY,
    RE_SIGNER,
    RE_TITLE_LIKE,
    RE_URGENCY,
    is_heading,
)


@dataclass
class ParagraphItem:
    index: int
    paragraph: Paragraph
    text: str


class OfficialDocumentAuditor:
    """Detect and audit Chinese official-document elements."""

    def __init__(self, config: OfficeToolConfig | None = None):
        self.config = config or OfficeToolConfig()

    def audit_document(self, doc: DocxDocument) -> AuditReport:
        paragraphs = self._paragraph_items(doc.paragraphs)
        structure = self.detect_structure(paragraphs)
        profile = self._resolve_profile(structure)
        report = AuditReport(profile=profile, is_red_head=profile == "red_head")
        report.stats["paragraphs"] = len(paragraphs)
        report.stats["tables"] = len(doc.tables)

        self._add_elements(report, structure, paragraphs)
        self._check_structure(report, structure, paragraphs)
        if self.config.audit.check_page_layout and doc.sections:
            self._check_page_layout(report, doc)
        self._check_obsolete_terms(report, paragraphs)
        return report

    def detect_structure(self, paragraphs: Iterable[ParagraphItem]) -> DetectedStructure:
        items = list(paragraphs)
        structure = DetectedStructure()
        scan_items = items[: max(1, self.config.audit.front_matter_scan_paragraphs)]

        for item in scan_items:
            text = item.text
            if structure.copy_number is None and RE_COPY_NUMBER.match(text):
                structure.copy_number = item.index
                continue
            if structure.secrecy is None and RE_SECRECY.match(text):
                structure.secrecy = item.index
                continue
            if structure.urgency is None and RE_URGENCY.match(text):
                structure.urgency = item.index
                continue
            if structure.document_number is None and RE_DOCUMENT_NUMBER.search(text):
                structure.document_number = item.index
                continue
            if structure.signer is None and RE_SIGNER.search(text):
                structure.signer = item.index
                continue
            if structure.red_head is None and self._looks_like_red_head(text):
                structure.red_head = item.index

        title = self._find_title(items, structure)
        structure.title = title.index if title else None
        if title:
            main_send = self._find_main_send(items, title.index)
            structure.main_send = main_send.index if main_send else None
            structure.body_start = self._find_body_start(items, title.index, structure.main_send)

        for item in items:
            if RE_ATTACHMENT_NOTE.match(item.text) or RE_ATTACHMENT_MARK.match(item.text):
                structure.attachment_notes.append(item.index)
            if RE_COPY_TO.match(item.text) and structure.copy_to is None:
                structure.copy_to = item.index
            if RE_PRINT_ORG_DATE.search(item.text) and structure.print_org_date is None:
                structure.print_org_date = item.index
            if RE_DATE.search(item.text):
                structure.date = item.index

        if structure.date is not None:
            structure.signatory = self._find_signatory_before_date(items, structure.date)

        return structure

    @staticmethod
    def _paragraph_items(paragraphs: Iterable[Paragraph]) -> list[ParagraphItem]:
        return [
            ParagraphItem(index, paragraph, paragraph.text.strip())
            for index, paragraph in enumerate(paragraphs)
            if paragraph.text and paragraph.text.strip()
        ]

    def _resolve_profile(self, structure: DetectedStructure) -> str:
        configured = self.config.audit.profile
        if configured != "auto":
            return configured
        if structure.red_head is not None:
            return "red_head"
        return "standard"

    @staticmethod
    def _looks_like_red_head(text: str) -> bool:
        if not text or len(text) > 42:
            return False
        if is_heading(text):
            return False
        if RE_DOCUMENT_NUMBER.search(text) or RE_SIGNER.search(text):
            return False
        return bool(RE_RED_HEAD_KEYWORDS.search(text))

    def _find_title(self, items: list[ParagraphItem], structure: DetectedStructure) -> ParagraphItem | None:
        start = structure.front_matter_end() + 1
        candidates = [item for item in items if item.index >= start]
        if not candidates and items:
            candidates = items

        preferred: list[ParagraphItem] = []
        fallback: list[ParagraphItem] = []
        for item in candidates[:12]:
            text = item.text
            if self._is_non_title_front_matter(text):
                continue
            if is_heading(text) or RE_MAIN_SEND.match(text):
                continue
            if RE_TITLE_LIKE.match(text):
                preferred.append(item)
            elif 4 <= len(text) <= 80:
                fallback.append(item)
        return (preferred or fallback or [None])[0]

    @staticmethod
    def _is_non_title_front_matter(text: str) -> bool:
        return bool(
            RE_COPY_NUMBER.match(text)
            or RE_SECRECY.match(text)
            or RE_URGENCY.match(text)
            or RE_DOCUMENT_NUMBER.search(text)
            or RE_SIGNER.search(text)
            or RE_ATTACHMENT_MARK.match(text)
        )

    @staticmethod
    def _find_main_send(items: list[ParagraphItem], title_index: int) -> ParagraphItem | None:
        for item in items:
            if item.index <= title_index:
                continue
            if RE_MAIN_SEND.match(item.text):
                return item
            if is_heading(item.text) or len(item.text) > 100:
                return None
        return None

    @staticmethod
    def _find_body_start(items: list[ParagraphItem], title_index: int, main_send_index: int | None) -> int | None:
        start_after = main_send_index if main_send_index is not None else title_index
        for item in items:
            if item.index > start_after:
                return item.index
        return None

    @staticmethod
    def _find_signatory_before_date(items: list[ParagraphItem], date_index: int) -> int | None:
        before = [item for item in items if item.index < date_index]
        for item in reversed(before[-6:]):
            if 2 <= len(item.text) <= 40 and not is_heading(item.text):
                if not RE_ATTACHMENT_NOTE.match(item.text) and not RE_MAIN_SEND.match(item.text):
                    return item.index
        return None

    def _add_elements(self, report: AuditReport, structure: DetectedStructure, paragraphs: list[ParagraphItem]) -> None:
        by_index = {item.index: item.text for item in paragraphs}
        names = [
            ("copy_number", "份号"),
            ("secrecy", "密级和保密期限"),
            ("urgency", "紧急程度"),
            ("red_head", "发文机关标志"),
            ("document_number", "发文字号"),
            ("signer", "签发人"),
            ("title", "标题"),
            ("main_send", "主送机关"),
            ("body_start", "正文起始"),
            ("signatory", "发文机关署名"),
            ("date", "成文日期"),
            ("copy_to", "抄送机关"),
            ("print_org_date", "印发机关和日期"),
        ]
        for field_name, role in names:
            index = getattr(structure, field_name)
            if index is not None and index in by_index:
                report.add_element(field_name, index, by_index[index], role)
        for index in structure.attachment_notes:
            if index in by_index:
                report.add_element("attachment_note", index, by_index[index], "附件说明")

    def _check_structure(
        self, report: AuditReport, structure: DetectedStructure, paragraphs: list[ParagraphItem]
    ) -> None:
        by_index = {item.index: item.text for item in paragraphs}
        if structure.title is None:
            report.add_finding(
                "missing_title",
                "error",
                "未识别到公文标题。",
                suggestion="在版头或发文字号之后补充独立标题段落。",
                can_fix=False,
            )

        if report.is_red_head and structure.document_number is None:
            severity = "warning" if self.config.audit.require_document_number_for_red_head else "info"
            report.add_finding(
                "missing_document_number",
                severity,
                "疑似红头文件未识别到发文字号。",
                block_index=structure.red_head,
                text=by_index.get(structure.red_head or -1, ""),
                expected="示例：某政发〔2026〕1号",
                suggestion="补充或修正发文字号行。",
                can_fix=False,
            )

        if report.is_red_head and self.config.audit.require_signer_for_red_head and structure.signer is None:
            report.add_finding(
                "missing_signer",
                "warning",
                "当前配置要求红头文件包含签发人，但未识别到签发人行。",
                suggestion="上行文通常需要在发文字号右侧或相邻行标注签发人。",
            )

        if self.config.audit.require_main_send and structure.main_send is None:
            report.add_finding(
                "missing_main_send",
                "warning",
                "未识别到主送机关。",
                suggestion="如该文种需要主送机关，请在标题下方添加以冒号结尾的主送机关段落。",
            )

        if self.config.audit.require_date and structure.date is None:
            report.add_finding(
                "missing_date",
                "warning",
                "未识别到成文日期。",
                expected="YYYY年M月D日",
                suggestion="在正文或发文机关署名之后补充成文日期。",
            )

        if structure.document_number is not None and structure.title is not None:
            if structure.document_number > structure.title:
                report.add_finding(
                    "document_number_after_title",
                    "warning",
                    "发文字号出现在标题之后，顺序可能不符合公文版头习惯。",
                    block_index=structure.document_number,
                    text=by_index.get(structure.document_number, ""),
                    can_fix=False,
                )

        for index in structure.attachment_notes:
            if structure.date is not None and index > structure.date:
                report.add_finding(
                    "attachment_note_after_date",
                    "warning",
                    "附件说明出现在成文日期之后，建议放在正文之后、署名日期之前。",
                    block_index=index,
                    text=by_index.get(index, ""),
                    can_fix=False,
                )

    def _check_page_layout(self, report: AuditReport, doc: DocxDocument) -> None:
        page = self.config.page
        section = doc.sections[0]
        checks = [
            ("paper_width_cm", "纸张宽度", section.page_width, page.paper_width_cm),
            ("paper_height_cm", "纸张高度", section.page_height, page.paper_height_cm),
            ("margin_top_cm", "上边距", section.top_margin, page.margin_top_cm),
            ("margin_bottom_cm", "下边距", section.bottom_margin, page.margin_bottom_cm),
            ("margin_left_cm", "左边距", section.left_margin, page.margin_left_cm),
            ("margin_right_cm", "右边距", section.right_margin, page.margin_right_cm),
        ]
        for code, label, actual_len, expected_cm in checks:
            actual_cm = _length_to_cm(actual_len)
            if abs(actual_cm - expected_cm) > self.config.audit.layout_tolerance_cm:
                report.add_finding(
                    f"layout_{code}",
                    "info",
                    f"{label}与默认公文版式不一致。",
                    expected=f"{expected_cm:.2f} cm",
                    actual=f"{actual_cm:.2f} cm",
                    suggestion="执行 format 命令可按默认配置修复页面设置。",
                    can_fix=True,
                )

    @staticmethod
    def _check_obsolete_terms(report: AuditReport, paragraphs: list[ParagraphItem]) -> None:
        for item in paragraphs:
            if RE_OBSOLETE_SUBJECT.match(item.text):
                report.add_finding(
                    "obsolete_subject_words",
                    "info",
                    "识别到“主题词”，2012 年后常规党政机关公文一般不再使用主题词。",
                    block_index=item.index,
                    text=item.text,
                    suggestion="如非历史文档或特殊模板，建议删除主题词栏。",
                    can_fix=False,
                )
            if RE_EMPTY.match(item.text):
                continue


def _length_to_cm(value: Length | None) -> float:
    if value is None:
        return 0.0
    return value.cm
