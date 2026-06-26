"""Official document structure audit."""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Iterable

from docx.document import Document as DocxDocument
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Length
from docx.text.paragraph import Paragraph

from .config import OfficeToolConfig
from .docx_utils import get_document_grid
from .models import AuditReport, DetectedStructure
from .patterns import (
    RE_ARABIC_DATE,
    RE_ATTACHMENT_MARK,
    RE_ATTACHMENT_END_PUNCT,
    RE_ATTACHMENT_NOTE,
    RE_BAD_EFFECTIVE_DATE,
    RE_COPY_NUMBER,
    RE_COPY_TO,
    RE_DATE,
    RE_DOCUMENT_NUMBER,
    RE_EMPTY,
    RE_HALFWIDTH_PUNCT,
    RE_INTERNAL_NOTICE,
    RE_LETTER_CONTACT,
    RE_LETTER_DOCUMENT_NUMBER,
    RE_MAIN_SEND,
    RE_MEETING_ATTENDEES,
    RE_DISTRIBUTION,
    RE_MEETING_ISSUE_LINE,
    RE_MEETING_NUMBER,
    RE_MEETING_RED_HEAD,
    RE_OBSOLETE_SUBJECT,
    RE_PRINT_ORG_DATE,
    RE_REGULATION_ARTICLE_PARTS,
    RE_REGULATION_CHAPTER,
    RE_REGULATION_CODE,
    RE_RED_HEAD_KEYWORDS,
    RE_SECRECY,
    RE_SIGNER,
    RE_TITLE_LIKE,
    RE_TITLE_DATE_LINE,
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
        report = AuditReport(
            profile=profile,
            is_red_head=profile in {"red_head", "letter_head", "meeting_minutes"},
            is_letter_head=profile == "letter_head",
            is_meeting_minutes=profile == "meeting_minutes",
        )
        report.stats["paragraphs"] = len(paragraphs)
        report.stats["tables"] = len(doc.tables)

        self._add_elements(report, structure, paragraphs)
        self._check_structure(report, structure, paragraphs)
        if self.config.audit.check_front_matter_order:
            self._check_front_matter_order(report, structure, paragraphs)
        if self.config.audit.check_document_number_format:
            self._check_document_number_format(report, structure, paragraphs)
        if self.config.audit.check_page_layout and doc.sections:
            self._check_page_layout(report, doc)
        if self.config.audit.check_document_grid and doc.sections:
            self._check_document_grid(report, doc)
        if self.config.audit.check_unit_typography:
            self._check_unit_typography(report, paragraphs)
        if self.config.audit.check_title_line_shape:
            self._check_title_line_shape(report, structure, paragraphs)
        if self.config.audit.check_imprint_rules:
            self._check_imprint_rules(report, structure, paragraphs)
        if self.config.audit.check_attachment_layout:
            self._check_attachment_layout(report, structure, paragraphs)
        if report.is_letter_head:
            self._check_letter_head_rules(report, doc)
        if report.is_meeting_minutes:
            self._check_meeting_minutes_rules(report, doc, structure)
        if structure.distribution is not None:
            self._check_distribution_rules(report, doc)
        self._check_unit_rules(report, structure, paragraphs, doc)
        self._check_obsolete_terms(report, paragraphs)
        return report

    def detect_structure(self, paragraphs: Iterable[ParagraphItem]) -> DetectedStructure:
        items = list(paragraphs)
        structure = DetectedStructure()
        scan_items = items[: max(1, self.config.audit.front_matter_scan_paragraphs)]

        for item in scan_items:
            text = item.text
            if self._paragraph_style_id(item.paragraph) == "OfficeToolGeneratedSimpleImprint":
                structure.simple_imprint = item.index
                continue
            if structure.copy_number is None and RE_COPY_NUMBER.match(text):
                structure.copy_number = item.index
                continue
            if structure.secrecy is None and RE_SECRECY.match(text):
                structure.secrecy = item.index
                continue
            if structure.urgency is None and RE_URGENCY.match(text):
                structure.urgency = item.index
                continue
            if structure.internal_notice is None and RE_INTERNAL_NOTICE.match(text):
                structure.internal_notice = item.index
                continue
            if RE_MEETING_RED_HEAD.match(text):
                structure.red_head = item.index
                structure.is_meeting_minutes = True
                continue
            if structure.meeting_number is None and RE_MEETING_NUMBER.match(text):
                structure.meeting_number = item.index
                continue
            if structure.meeting_issue_line is None and RE_MEETING_ISSUE_LINE.match(text):
                structure.meeting_issue_line = item.index
                continue
            if structure.document_number is None and (RE_DOCUMENT_NUMBER.search(text) or RE_LETTER_DOCUMENT_NUMBER.search(text)):
                structure.document_number = item.index
                structure.is_letter_head = bool(RE_LETTER_DOCUMENT_NUMBER.search(text))
                continue
            if structure.signer is None and RE_SIGNER.search(text):
                structure.signer = item.index
                continue
            if structure.red_head is None and self._looks_like_red_head(text):
                structure.red_head = item.index

        if structure.red_head is None and structure.document_number is not None and (
            structure.is_letter_head or self.config.audit.profile == "letter_head"
        ):
            candidates = [
                item
                for item in scan_items
                if item.index < structure.document_number
                and not RE_INTERNAL_NOTICE.match(item.text)
                and not RE_COPY_NUMBER.match(item.text)
                and not RE_SECRECY.match(item.text)
                and not RE_URGENCY.match(item.text)
            ]
            if candidates:
                structure.red_head = candidates[-1].index

        title = self._find_title(items, structure)
        structure.title = title.index if title else None
        structure._title_text = title.text if title else ""
        if title:
            main_send = self._find_main_send(items, title.index)
            structure.main_send = main_send.index if main_send else None
            structure.body_start = self._find_body_start(items, title.index, structure.main_send)

        for item in items:
            if self._paragraph_style_id(item.paragraph) == "OfficeToolGeneratedSimpleImprint":
                structure.simple_imprint = item.index
                continue
            if structure.internal_notice is None and RE_INTERNAL_NOTICE.match(item.text):
                structure.internal_notice = item.index
            if RE_ATTACHMENT_NOTE.match(item.text) or RE_ATTACHMENT_MARK.match(item.text):
                structure.attachment_notes.append(item.index)
            if RE_COPY_TO.match(item.text) and structure.copy_to is None:
                structure.copy_to = item.index
            if RE_PRINT_ORG_DATE.search(item.text) and structure.print_org_date is None:
                structure.print_org_date = item.index
            if structure.regulation_code is None and RE_REGULATION_CODE.match(item.text):
                structure.regulation_code = item.index
            if RE_REGULATION_CHAPTER.match(item.text):
                structure.regulation_chapters.append(item.index)
            if RE_REGULATION_ARTICLE_PARTS.match(item.text):
                structure.regulation_articles.append(item.index)
            if RE_LETTER_CONTACT.match(item.text):
                structure.letter_contacts.append(item.index)
            if RE_MEETING_ATTENDEES.match(item.text):
                structure.meeting_attendees.append(item.index)
            if RE_DISTRIBUTION.match(item.text) and structure.distribution is None:
                structure.distribution = item.index
            if (
                not structure.is_meeting_minutes
                and self._is_document_date_line(item.text)
                and not self._is_title_date_line(item, structure)
                and not RE_COPY_TO.match(item.text)
                and not RE_PRINT_ORG_DATE.search(item.text)
                and item.index != structure.meeting_issue_line
            ):
                structure.date = item.index

        if structure.date is not None:
            structure.signatory = self._find_signatory_before_date(items, structure.date)

        if structure.regulation_code is not None:
            structure.regulation_title = self._find_regulation_title(items, structure.regulation_code)
        else:
            structure.regulation_chapters.clear()
            structure.regulation_articles.clear()

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
        if structure.is_meeting_minutes:
            return "meeting_minutes"
        if structure.is_letter_head:
            return "letter_head"
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
            or RE_INTERNAL_NOTICE.match(text)
            or RE_DOCUMENT_NUMBER.search(text)
            or RE_SIGNER.search(text)
            or RE_ATTACHMENT_MARK.match(text)
        )

    @staticmethod
    def _is_title_date_line(item: ParagraphItem, structure: DetectedStructure) -> bool:
        return (
            structure.title is not None
            and item.index > structure.title
            and item.index <= structure.title + 3
            and bool(RE_TITLE_DATE_LINE.match(item.text))
        )

    @staticmethod
    def _is_document_date_line(text: str) -> bool:
        normalized = re.sub(r"\s+", "", text.strip())
        return bool(RE_DATE.fullmatch(normalized))

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
    def _find_regulation_title(items: list[ParagraphItem], code_index: int) -> int | None:
        for item in items:
            if item.index <= code_index:
                continue
            if RE_REGULATION_CHAPTER.match(item.text):
                return None
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
            ("internal_notice", "内部资料提示"),
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
            ("simple_imprint", "印发机关和日期简版版记"),
            ("regulation_code", "规章制度编号"),
            ("regulation_title", "规章制度标题"),
            ("meeting_number", "会议纪要期号"),
            ("meeting_issue_line", "会议纪要编发机关和日期"),
            ("distribution", "分送版记"),
        ]
        for field_name, role in names:
            index = getattr(structure, field_name)
            if index is not None and index in by_index:
                report.add_element(field_name, index, by_index[index], role)
        for index in structure.attachment_notes:
            if index in by_index:
                report.add_element("attachment_note", index, by_index[index], "附件说明")
        for index in structure.regulation_chapters:
            if index in by_index:
                report.add_element("regulation_chapter", index, by_index[index], "规章制度章标题")
        for index in structure.regulation_articles:
            if index in by_index:
                report.add_element("regulation_article", index, by_index[index], "规章制度条文")
        for index in structure.letter_contacts:
            if index in by_index:
                report.add_element("letter_contact", index, by_index[index], "函联系人")
        for index in structure.meeting_attendees:
            if index in by_index:
                report.add_element("meeting_attendees", index, by_index[index], "出席人员")

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

        if report.is_red_head and not report.is_meeting_minutes and structure.document_number is None:
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

        if self.config.audit.require_date and structure.date is None and not report.is_meeting_minutes:
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

    @staticmethod
    def _check_letter_head_rules(report: AuditReport, doc: DocxDocument) -> None:
        if not doc.sections:
            return
        section = doc.sections[0]
        vml_shape = "{urn:schemas-microsoft-com:vml}shape"
        vml_line = "{urn:schemas-microsoft-com:vml}line"
        body_shapes = {shape.get("id") for shape in doc._element.body.iter(vml_shape)}
        body_lines = {line.get("id") for line in doc._element.body.iter(vml_line)}
        first_footer_text = "".join(section.first_page_footer._element.itertext())
        default_footer_text = "".join(section.footer._element.itertext())

        checks = [
            (
                "letter_header_textbox_missing",
                "OfficeToolLetterHeadTextBox" in body_shapes,
                "函格式首页缺少正文锚定的版头文本框。",
            ),
            (
                "letter_top_rule_missing",
                "OfficeToolLetterRedTop" in body_lines,
                "函格式版头下方缺少粗到细复合红线。",
            ),
            (
                "letter_bottom_rule_missing",
                "OfficeToolLetterRedBottom" in body_lines,
                "函格式首页底部缺少细到粗复合红线。",
            ),
            (
                "letter_first_page_number_wrong",
                section.different_first_page_header_footer
                and "PAGE" not in first_footer_text
                and "PAGE" in default_footer_text,
                "函格式应隐藏首页页码，并从后续页面显示实际页序。",
            ),
        ]
        for code, passed, message in checks:
            if not passed:
                report.add_finding(
                    code,
                    "warning",
                    message,
                    can_fix=True,
                    suggestion="执行校对导出可按函格式补齐首页版头、红线和页码设置。",
                )

    @staticmethod
    def _check_meeting_minutes_rules(
        report: AuditReport,
        doc: DocxDocument,
        structure: DetectedStructure,
    ) -> None:
        required = [
            ("meeting_number_missing", structure.meeting_number is not None, "未识别到会议纪要期号。"),
            ("meeting_issue_line_missing", structure.meeting_issue_line is not None, "未识别到会议纪要编发部门和日期行。"),
        ]
        for code, passed, message in required:
            if not passed:
                report.add_finding(code, "warning", message, can_fix=False)
        line_ids = {
            line.get("id")
            for line in doc._element.body.iter("{urn:schemas-microsoft-com:vml}line")
        }
        if "OfficeToolRedSeparator" not in line_ids:
            report.add_finding(
                "meeting_red_rule_missing",
                "warning",
                "会议纪要编发部门和日期下方缺少红色分隔线。",
                can_fix=True,
                suggestion="执行校对导出可补齐会议纪要红线。",
            )

    @staticmethod
    def _check_distribution_rules(report: AuditReport, doc: DocxDocument) -> None:
        line_ids = {
            line.get("id")
            for line in doc._element.body.iter("{urn:schemas-microsoft-com:vml}line")
        }
        required = {"OfficeToolDistributionTopLine", "OfficeToolDistributionBottomLine"}
        if not required.issubset(line_ids):
            report.add_finding(
                "distribution_layout_missing",
                "warning",
                "分送版记应位于偶数末页底部，并置于两条 15.6 厘米分隔线之间。",
                can_fix=True,
                suggestion="执行校对导出可使用段落标记补齐到偶数页末并添加分隔线。",
            )

    def _check_front_matter_order(
        self, report: AuditReport, structure: DetectedStructure, paragraphs: list[ParagraphItem]
    ) -> None:
        by_index = {item.index: item.text for item in paragraphs}
        ordered_roles = [
            ("copy_number", "份号"),
            ("secrecy", "密级和保密期限"),
            ("urgency", "紧急程度"),
            ("red_head", "发文机关标志"),
            ("document_number", "发文字号"),
            ("signer", "签发人"),
        ]
        if report.is_meeting_minutes:
            ordered_roles = [
                ("internal_notice", "内部资料提示"),
                ("red_head", "会议纪要版头"),
                ("meeting_number", "期号"),
                ("meeting_issue_line", "编发机关和日期"),
            ]
        if not report.is_letter_head and not report.is_meeting_minutes:
            ordered_roles.insert(3, ("internal_notice", "内部资料提示"))
        present: list[tuple[str, str, int]] = []
        for field_name, label in ordered_roles:
            index = getattr(structure, field_name)
            if index is not None:
                present.append((field_name, label, index))
        for previous, current in zip(present, present[1:]):
            if previous[2] > current[2]:
                report.add_finding(
                    "front_matter_order",
                    "warning",
                    f"{current[1]}出现在{previous[1]}之前，版头要素顺序不符合常规公文格式。",
                    block_index=current[2],
                    text=by_index.get(current[2], ""),
                    expected="份号、密级和保密期限、紧急程度、发文机关标志、发文字号、签发人",
                    can_fix=False,
                )
                return

    def _check_document_number_format(
        self, report: AuditReport, structure: DetectedStructure, paragraphs: list[ParagraphItem]
    ) -> None:
        if structure.document_number is None:
            return
        item = next((item for item in paragraphs if item.index == structure.document_number), None)
        if item is None:
            return
        text = item.text.strip()
        if not re.match(r"^[\u4e00-\u9fffA-Za-z0-9]{1,24}〔\d{4}〕\d+号$", text):
            report.add_finding(
                "document_number_format_irregular",
                "info",
                "发文字号建议使用“机关代字〔年份〕序号号”的规范格式。",
                block_index=item.index,
                text=text,
                expected="示例：某政发〔2026〕8号",
                suggestion="请核对六角括号、年份、序号和“号”字，中间通常不留空格。",
                can_fix=False,
            )

    def _check_page_layout(self, report: AuditReport, doc: DocxDocument) -> None:
        page = self.config.page
        for section_index, section in enumerate(doc.sections, start=1):
            checks = [
                ("paper_width_cm", "纸张宽度", section.page_width, page.paper_width_cm),
                ("paper_height_cm", "纸张高度", section.page_height, page.paper_height_cm),
                ("margin_top_cm", "上边距", section.top_margin, page.margin_top_cm),
                (
                    "margin_bottom_cm",
                    "下边距",
                    section.bottom_margin,
                    2.5 if report.is_letter_head else page.margin_bottom_cm,
                ),
                ("margin_left_cm", "左边距", section.left_margin, page.margin_left_cm),
                ("margin_right_cm", "右边距", section.right_margin, page.margin_right_cm),
                ("footer_distance_cm", "页脚距", section.footer_distance, page.footer_distance_cm),
            ]
            for code, label, actual_len, expected_cm in checks:
                actual_cm = _length_to_cm(actual_len)
                if abs(actual_cm - expected_cm) > self.config.audit.layout_tolerance_cm:
                    suffix = "" if section_index == 1 else f"_section_{section_index}"
                    report.add_finding(
                        f"layout_{code}{suffix}",
                        "info",
                        f"第 {section_index} 节{label}与默认公文版式不一致。",
                        expected=f"{expected_cm:.2f} cm",
                        actual=f"{actual_cm:.2f} cm",
                        suggestion="执行校对导出可按默认配置修复所有节的页面设置。",
                        can_fix=True,
                    )

    def _check_document_grid(self, report: AuditReport, doc: DocxDocument) -> None:
        page = self.config.page
        normal_size_pt = self.config.styles["body"].size_pt
        for section_index, section in enumerate(doc.sections, start=1):
            grid = get_document_grid(section)
            suffix = "" if section_index == 1 else f"_section_{section_index}"
            if not grid:
                report.add_finding(
                    f"layout_document_grid_missing{suffix}",
                    "info",
                    f"第 {section_index} 节未检测到文档网格设置，单位要求通常按每行 28 字、每页 22 行控制版心。",
                    expected=f"每行 {page.chars_per_line} 字，每页 {page.lines_per_page} 行",
                    suggestion="执行校对导出可为所有节写入文档网格设置。",
                    can_fix=True,
                )
                continue
            available_twips = section.page_width.twips - section.left_margin.twips - section.right_margin.twips
            max_pitch_pt = (available_twips - 2) / (page.chars_per_line * 20)
            effective_pitch_pt = min(page.grid_char_space_pt, max_pitch_pt)
            expected = {
                "type": "lines",
                "charsPerLine": str(page.chars_per_line),
                "linesPerPage": str(page.lines_per_page),
                "charSpace": str(int(round((effective_pitch_pt - normal_size_pt) * 4096))),
                "linePitch": str(int(round(page.grid_line_pitch_pt * 20))),
            }
            for key, expected_value in expected.items():
                if grid.get(key) != expected_value:
                    report.add_finding(
                        f"layout_document_grid_{key}{suffix}",
                        "info",
                        f"第 {section_index} 节文档网格与单位公文格式要求不一致。",
                        expected=f"{key}={expected_value}",
                        actual=f"{key}={grid.get(key, '')}",
                        suggestion="执行校对导出可按默认单位规则修复所有节的文档网格。",
                        can_fix=True,
                    )

    def _check_unit_typography(self, report: AuditReport, paragraphs: list[ParagraphItem]) -> None:
        role_by_index = {element.block_index: element.name for element in report.elements}
        expected_by_role = {
            "title": ("华文中宋", 22, True, "主标题应使用华文中宋 2 号加粗，行距 30 磅。"),
            "h1": ("黑体", 16, False, "一级标题应使用 3 号黑体，通常不加粗。"),
            "h2": ("楷体_GB2312", 16, False, "二级标题应使用 3 号楷体_GB2312，通常不加粗。"),
            "h3": ("仿宋_GB2312", 16, False, "三级标题应使用 3 号仿宋_GB2312，通常不加粗。"),
        }
        latin_checked = False
        for item in paragraphs:
            role = self._typography_role_for_text(report.profile, item.text) or role_by_index.get(item.index)
            if role in expected_by_role and item.paragraph.runs:
                font, size_pt, bold, message = expected_by_role[role]
                run = next((run for run in item.paragraph.runs if run.text.strip()), item.paragraph.runs[0])
                actual_font = self._run_east_asia_font(run)
                actual_size = run.font.size.pt if run.font.size else None
                actual_bold = run.bold
                mismatches = []
                if actual_font and actual_font != font:
                    mismatches.append(f"字体 {actual_font}")
                if actual_size is not None and abs(actual_size - size_pt) > 0.1:
                    mismatches.append(f"字号 {actual_size:g}pt")
                if actual_bold is not None and actual_bold != bold:
                    mismatches.append("加粗状态不符")
                if mismatches:
                    report.add_finding(
                        f"typography_{role}",
                        "info",
                        message,
                        block_index=item.index,
                        text=item.text,
                        expected=f"{font} {size_pt:g}pt {'加粗' if bold else '不加粗'}",
                        actual="，".join(mismatches),
                        suggestion="执行校对导出可按单位默认样式修复。",
                        can_fix=True,
                    )
            if not latin_checked and self._paragraph_has_non_times_latin(item.paragraph):
                latin_checked = True
                report.add_finding(
                    "typography_latin_font",
                    "info",
                    "文中包含数字或字母，单位要求数字、字母使用 Times New Roman 字体。",
                    block_index=item.index,
                    text=item.text,
                    suggestion="执行校对导出可将数字、字母字体修复为 Times New Roman。",
                    can_fix=True,
                )

    @staticmethod
    def _heading_role_for_text(text: str) -> str | None:
        from .patterns import heading_role

        return heading_role(text)

    def _typography_role_for_text(self, profile: str, text: str) -> str | None:
        return self._heading_role_for_text(text)

    @staticmethod
    def _paragraph_has_non_times_latin(paragraph: Paragraph) -> bool:
        for run in paragraph.runs:
            if not any(ch.isascii() and ch.isalnum() for ch in run.text):
                continue
            r_pr = run._element.rPr
            r_fonts = r_pr.rFonts if r_pr is not None else None
            ascii_font = ""
            if r_fonts is not None:
                ascii_font = r_fonts.get(qn("w:ascii"), "") or r_fonts.get(qn("w:hAnsi"), "")
            ascii_font = ascii_font or run.font.name or ""
            if ascii_font and ascii_font != "Times New Roman":
                return True
            if not ascii_font:
                return True
        return False

    @staticmethod
    def _run_east_asia_font(run) -> str:
        r_pr = run._element.rPr
        r_fonts = r_pr.rFonts if r_pr is not None else None
        if r_fonts is not None:
            return r_fonts.get(qn("w:eastAsia"), "") or run.font.name or ""
        return run.font.name or ""

    def _check_title_line_shape(
        self, report: AuditReport, structure: DetectedStructure, paragraphs: list[ParagraphItem]
    ) -> None:
        if structure.title is None:
            return
        title_item = next((item for item in paragraphs if item.index == structure.title), None)
        if title_item is None:
            return
        compact_title = title_item.text.replace(" ", "")
        if len(compact_title) <= self.config.page.chars_per_line + 4:
            return
        if "\n" in title_item.paragraph.text or "\r" in title_item.paragraph.text:
            return
        report.add_finding(
            "title_line_shape_hint",
            "info",
            "标题较长，定稿时应检查断行后的梯形、菱形或沙漏形排布。",
            block_index=title_item.index,
            text=title_item.text,
            expected="标题多行排列匀称，避免上下同宽或单字悬挂。",
            suggestion="建议人工检查标题换行位置；自动格式化仅统一字体、字号和居中。",
            can_fix=False,
        )

    def _check_imprint_rules(
        self, report: AuditReport, structure: DetectedStructure, paragraphs: list[ParagraphItem]
    ) -> None:
        indexes = [item.index for item in paragraphs]
        by_index = {item.index: item for item in paragraphs}
        if structure.copy_to is not None and structure.print_org_date is not None:
            if structure.copy_to > structure.print_org_date:
                copy_item = by_index.get(structure.copy_to)
                report.add_finding(
                    "imprint_order_wrong",
                    "warning",
                    "版记中抄送机关应位于印发机关和印发日期之前。",
                    block_index=structure.copy_to,
                    text=copy_item.text if copy_item else "",
                    can_fix=False,
                )
            last_imprint = max(structure.copy_to, structure.print_org_date)
            if any(index > last_imprint for index in indexes):
                last_item = by_index.get(last_imprint)
                report.add_finding(
                    "imprint_not_at_end",
                    "info",
                    "版记通常应置于文档末尾，当前版记之后仍有正文段落。",
                    block_index=last_imprint,
                    text=last_item.text if last_item else "",
                    suggestion="请核对版记之后的内容是否应移动到版记之前。",
                    can_fix=False,
                )
            copy_item = by_index.get(structure.copy_to)
            print_item = by_index.get(structure.print_org_date)
            if copy_item and print_item and not self._has_complete_imprint_lines(copy_item.paragraph):
                report.add_finding(
                    "imprint_lines_missing",
                    "info",
                    "版记应使用分隔线区分抄送机关、印发机关和印发日期。",
                    block_index=structure.copy_to,
                    text=copy_item.text,
                    suggestion="执行校对导出可为版记段落添加基础分隔线。",
                    can_fix=True,
                )
        elif structure.copy_to is not None:
            copy_item = by_index.get(structure.copy_to)
            report.add_finding(
                "imprint_print_org_missing",
                "info",
                "识别到抄送机关，但未识别到印发机关和印发日期。",
                block_index=structure.copy_to,
                text=copy_item.text if copy_item else "",
                suggestion="如为正式定稿，请补充印发机关和印发日期。",
                can_fix=False,
            )

    @staticmethod
    def _has_complete_imprint_lines(paragraph: Paragraph) -> bool:
        parent = paragraph._p.getparent()
        tag = "{urn:schemas-microsoft-com:vml}line"
        shape_ids = {line.get("id") for line in parent.iter(tag)}
        return {
            "OfficeToolImprintTopLine",
            "OfficeToolImprintMiddleLine",
            "OfficeToolImprintBottomLine",
        }.issubset(shape_ids)

    def _check_attachment_layout(
        self, report: AuditReport, structure: DetectedStructure, paragraphs: list[ParagraphItem]
    ) -> None:
        by_index = {item.index: item for item in paragraphs}
        for index in structure.attachment_notes:
            item = by_index.get(index)
            if item is None:
                continue
            left_chars = self._paragraph_left_indent_chars(item.paragraph)
            hanging_chars = self._paragraph_hanging_indent_chars(item.paragraph)
            if not (
                left_chars is not None
                and abs(left_chars - 2) <= 0.25
                and hanging_chars is not None
                and abs(hanging_chars - 4.5) <= 0.25
                and item.paragraph.alignment == WD_ALIGN_PARAGRAPH.JUSTIFY
            ):
                report.add_finding(
                    "attachment_note_indent",
                    "info",
                    "附件首段应左缩进 2 字符、悬挂缩进 4.5 字符并两端对齐。",
                    block_index=index,
                    text=item.text,
                    expected="左缩进 2 字符，悬挂缩进 4.5 字符，两端对齐",
                    suggestion="执行校对导出可按附件首段样式修复。",
                    can_fix=True,
                )
            following = [candidate for candidate in paragraphs if candidate.index > index]
            for candidate in following:
                if not re.match(r"^\s*\d{1,2}[．.、]", candidate.text):
                    break
                left_chars = self._paragraph_left_indent_chars(candidate.paragraph)
                hanging_chars = self._paragraph_hanging_indent_chars(candidate.paragraph)
                if (
                    left_chars is not None
                    and abs(left_chars - 5) <= 0.25
                    and hanging_chars is not None
                    and abs(hanging_chars - 1.5) <= 0.25
                    and candidate.paragraph.alignment == WD_ALIGN_PARAGRAPH.JUSTIFY
                ):
                    continue
                report.add_finding(
                    "attachment_item_indent",
                    "info",
                    "后续附件段落应左缩进 5 字符、悬挂缩进 1.5 字符并两端对齐。",
                    block_index=candidate.index,
                    text=candidate.text,
                    expected="左缩进 5 字符，悬挂缩进 1.5 字符，两端对齐",
                    suggestion="执行校对导出可按后续附件样式修复。",
                    can_fix=True,
                )

    @staticmethod
    def _paragraph_left_indent_chars(paragraph: Paragraph) -> float | None:
        p_pr = paragraph._p.pPr
        ind = p_pr.ind if p_pr is not None else None
        if ind is not None:
            left_chars = ind.get(qn("w:leftChars"))
            if left_chars:
                try:
                    return int(left_chars) / 100
                except ValueError:
                    return None
        left_indent = paragraph.paragraph_format.left_indent
        size = next((run.font.size.pt for run in paragraph.runs if run.font.size), None)
        if left_indent is not None and size:
            return left_indent.pt / size
        return None

    @staticmethod
    def _paragraph_style_id(paragraph: Paragraph) -> str:
        p_pr = paragraph._p.pPr
        p_style = p_pr.pStyle if p_pr is not None else None
        return p_style.get(qn("w:val"), "") if p_style is not None else ""

    @staticmethod
    def _paragraph_hanging_indent_chars(paragraph: Paragraph) -> float | None:
        p_pr = paragraph._p.pPr
        ind = p_pr.ind if p_pr is not None else None
        if ind is not None:
            hanging_chars = ind.get(qn("w:hangingChars"))
            if hanging_chars:
                try:
                    return int(hanging_chars) / 100
                except ValueError:
                    return None
        first_line_indent = paragraph.paragraph_format.first_line_indent
        size = next((run.font.size.pt for run in paragraph.runs if run.font.size), None)
        if first_line_indent is not None and first_line_indent.pt < 0 and size:
            return -first_line_indent.pt / size
        return None

    def _check_unit_rules(self, report: AuditReport, structure: DetectedStructure, paragraphs: list[ParagraphItem], doc: DocxDocument) -> None:
        by_index = {item.index: item.text for item in paragraphs}
        if self.config.audit.check_date_format and structure.date is not None:
            date_text = by_index.get(structure.date, "")
            match = RE_DATE.search(date_text)
            if match and not RE_ARABIC_DATE.search(match.group(0)):
                report.add_finding(
                    "date_not_arabic",
                    "warning",
                    "成文日期应使用阿拉伯数字，年份写全称，月日不编虚位。",
                    block_index=structure.date,
                    text=date_text,
                    expected="2026年6月13日",
                    actual=match.group(0),
                    can_fix=False,
                )

        if self.config.audit.check_attachment_format:
            for index in structure.attachment_notes:
                text = by_index.get(index, "")
                if "：" not in text and RE_ATTACHMENT_NOTE.match(text):
                    report.add_finding(
                        "attachment_colon_not_fullwidth",
                        "warning",
                        "附件说明应使用全角冒号。",
                        block_index=index,
                        text=text,
                        expected="附件：附件名称",
                        suggestion="将英文冒号改为全角冒号。",
                        can_fix=False,
                    )
                name = text.split("：", 1)[-1].split(":", 1)[-1].strip()
                if name and RE_ATTACHMENT_END_PUNCT.search(name):
                    report.add_finding(
                        "attachment_name_has_punctuation",
                        "warning",
                        "附件名称后不应加标点符号。",
                        block_index=index,
                        text=text,
                        suggestion="删除附件名称末尾标点。",
                        can_fix=False,
                    )

        if self.config.audit.check_finalization_terms:
            for item in paragraphs:
                if RE_BAD_EFFECTIVE_DATE.search(item.text):
                    report.add_finding(
                        "bad_effective_date_wording",
                        "warning",
                        "识别到“自发布之日起执行/施行”等不当表述，单位要求考虑 5-10 天宣贯期。",
                        block_index=item.index,
                        text=item.text,
                        suggestion="改为明确施行日期，或预留宣贯期后执行。",
                        can_fix=False,
                    )
                if RE_HALFWIDTH_PUNCT.search(item.text):
                    report.add_finding(
                        "halfwidth_punctuation",
                        "info",
                        "识别到半角标点，单位要求中文公文标点符号使用全角。",
                        block_index=item.index,
                        text=item.text,
                        suggestion="将英文逗号、冒号、分号、问号、感叹号等改为中文全角标点。",
                        can_fix=False,
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
