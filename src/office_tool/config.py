"""Configuration defaults for OfficeTool."""

from __future__ import annotations

from copy import deepcopy
from dataclasses import asdict, dataclass, field
from typing import Any


@dataclass
class PageSpec:
    paper_width_cm: float = 21.0
    paper_height_cm: float = 29.7
    margin_top_cm: float = 3.7
    margin_bottom_cm: float = 3.5
    margin_left_cm: float = 2.8
    margin_right_cm: float = 2.6
    footer_distance_cm: float = 2.5
    line_spacing_pt: float = 28.0
    title_line_spacing_pt: float = 33.0
    red_head_line_spacing_pt: float = 48.0
    chars_per_line: int = 28
    lines_per_page: int = 22
    grid_char_space_pt: float = 15.8
    grid_line_pitch_pt: float = 28.95


@dataclass
class StyleSpec:
    font: str
    size_pt: float
    alignment: str = "left"
    bold: bool | None = None
    color: str | None = None
    first_line_chars: int = 0
    line_spacing_pt: float | None = None
    space_before_pt: float = 0.0
    space_after_pt: float = 0.0
    latin_font: str = "Times New Roman"
    left_indent_chars: int = 0
    right_indent_chars: int = 0


@dataclass
class AuditOptions:
    profile: str = "auto"
    front_matter_scan_paragraphs: int = 14
    require_document_number_for_red_head: bool = True
    require_signer_for_red_head: bool = False
    require_main_send: bool = False
    require_date: bool = True
    check_page_layout: bool = True
    check_document_grid: bool = True
    check_unit_typography: bool = True
    check_date_format: bool = True
    check_attachment_format: bool = True
    check_finalization_terms: bool = True
    check_title_line_shape: bool = False
    check_imprint_rules: bool = True
    check_document_number_format: bool = True
    check_attachment_layout: bool = True
    check_front_matter_order: bool = True
    layout_tolerance_cm: float = 0.2


@dataclass
class FormatOptions:
    apply_page_setup: bool = True
    apply_document_grid: bool = True
    apply_styles: bool = True
    add_page_number: bool = True
    page_number_odd_even: bool = True
    draw_red_separator: bool = True
    draw_imprint_lines: bool = True
    preserve_existing_bold_italic: bool = True


@dataclass
class AIReviewOptions:
    enabled: bool = False
    provider: str = "deepseek"
    base_url: str = ""
    endpoint_path: str = "/chat/completions"
    model: str = "deepseek-chat"
    api_key: str = ""
    api_key_env: str = "DEEPSEEK_API_KEY"
    auth_prefix: str = "Bearer"
    stream: bool = False
    strip_reasoning: bool = True
    timeout_seconds: int = 300
    temperature: float = 0.1
    max_tokens: int = 2048
    max_input_chars: int = 12000


@dataclass
class ContentGenerationOptions:
    """Optional business-content generation used only during file export."""

    add_red_head: bool = False
    add_imprint: bool = False
    red_head_title: str = ""
    document_number: str = ""
    copy_to: str = ""
    print_organization: str = ""
    print_date: str = ""
    meeting_number: str = ""
    meeting_organization: str = ""
    meeting_date: str = ""
    distribution: str = ""


def _default_styles() -> dict[str, StyleSpec]:
    return {
        "internal_notice": StyleSpec("\u9ed1\u4f53", 14, "right", bold=False),
        "red_head": StyleSpec("\u534e\u6587\u4e2d\u5b8b", 42, "center", bold=False, color="FF0000", line_spacing_pt=48),
        "document_number": StyleSpec("\u4eff\u5b8b_GB2312", 16, "center", bold=False, line_spacing_pt=30),
        "signer": StyleSpec("\u4eff\u5b8b_GB2312", 16, "right"),
        "title": StyleSpec("\u534e\u6587\u4e2d\u5b8b", 22, "center", bold=True, line_spacing_pt=30),
        "title_date": StyleSpec("\u4eff\u5b8b_GB2312", 16, "center"),
        "subtitle": StyleSpec("\u6977\u4f53_GB2312", 16, "center", line_spacing_pt=33),
        "main_send": StyleSpec("\u4eff\u5b8b_GB2312", 16, "left"),
        "body": StyleSpec("\u4eff\u5b8b_GB2312", 16, "justify", first_line_chars=2),
        "h1": StyleSpec("\u9ed1\u4f53", 16, "justify", bold=False, first_line_chars=2),
        "h2": StyleSpec("\u6977\u4f53_GB2312", 16, "justify", bold=False, first_line_chars=2),
        "h3": StyleSpec("\u4eff\u5b8b_GB2312", 16, "justify", bold=False, first_line_chars=2),
        "h4": StyleSpec("\u4eff\u5b8b_GB2312", 16, "justify", bold=False, first_line_chars=2),
        "attachment": StyleSpec("\u4eff\u5b8b_GB2312", 16, "justify", left_indent_chars=2),
        "attachment_item": StyleSpec("\u4eff\u5b8b_GB2312", 16, "justify", left_indent_chars=5),
        "signatory": StyleSpec("\u4eff\u5b8b_GB2312", 16, "right"),
        "date": StyleSpec("\u4eff\u5b8b_GB2312", 16, "right"),
        "regulation_code": StyleSpec("Times New Roman", 16, "justify", bold=False, line_spacing_pt=30),
        "regulation_title": StyleSpec("\u534e\u6587\u4e2d\u5b8b", 22, "center", bold=True, line_spacing_pt=30),
        "regulation_chapter": StyleSpec("\u9ed1\u4f53", 16, "center", bold=False),
        "regulation_article": StyleSpec("\u4eff\u5b8b_GB2312", 16, "justify", first_line_chars=2),
        "copy_to": StyleSpec("\u4eff\u5b8b_GB2312", 14, "justify", bold=False, left_indent_chars=1, right_indent_chars=1),
        "page_number": StyleSpec("\u5b8b\u4f53", 14, "center"),
        "letter_red_head": StyleSpec("\u534e\u6587\u4e2d\u5b8b", 41, "center", bold=False, color="FF0000", line_spacing_pt=40),
        "letter_document_number": StyleSpec("\u4eff\u5b8b_GB2312", 16, "right"),
        "letter_internal_notice": StyleSpec("\u9ed1\u4f53", 16, "justify", bold=False, first_line_chars=2),
        "letter_contact": StyleSpec("\u4eff\u5b8b_GB2312", 16, "justify", first_line_chars=2, line_spacing_pt=None),
        "meeting_red_head": StyleSpec("\u534e\u6587\u4e2d\u5b8b", 49, "center", bold=False, color="FF0000", line_spacing_pt=48),
        "meeting_number": StyleSpec("\u4eff\u5b8b_GB2312", 16, "center", bold=False),
        "meeting_issue_line": StyleSpec("\u4eff\u5b8b_GB2312", 16, "center", bold=False, left_indent_chars=1, right_indent_chars=1),
        "meeting_attendees": StyleSpec("\u4eff\u5b8b_GB2312", 16, "justify", bold=False),
        "meeting_distribution": StyleSpec("\u4eff\u5b8b_GB2312", 14, "justify", bold=False, left_indent_chars=1, right_indent_chars=1),
    }


@dataclass
class OfficeToolConfig:
    page: PageSpec = field(default_factory=PageSpec)
    audit: AuditOptions = field(default_factory=AuditOptions)
    format: FormatOptions = field(default_factory=FormatOptions)
    ai_review: AIReviewOptions = field(default_factory=AIReviewOptions)
    generation: ContentGenerationOptions = field(default_factory=ContentGenerationOptions)
    styles: dict[str, StyleSpec] = field(default_factory=_default_styles)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, raw: dict[str, Any] | None) -> "OfficeToolConfig":
        config = cls()
        if not raw:
            return config

        raw = deepcopy(raw)
        if "page" in raw:
            config.page = _merge_dataclass(config.page, raw["page"])
        if "audit" in raw:
            config.audit = _merge_dataclass(config.audit, raw["audit"])
        if "format" in raw:
            config.format = _merge_dataclass(config.format, raw["format"])
        if "ai_review" in raw:
            config.ai_review = _merge_dataclass(config.ai_review, raw["ai_review"])
        if "generation" in raw:
            config.generation = _merge_dataclass(config.generation, raw["generation"])
        if "styles" in raw:
            styles = dict(config.styles)
            for name, value in raw["styles"].items():
                if name in styles:
                    styles[name] = _merge_dataclass(styles[name], value)
                else:
                    styles[name] = StyleSpec(**value)
            config.styles = styles
        return config

    def set_path(self, path: str, value: Any) -> None:
        parts = path.split(".")
        if not parts:
            raise ValueError("\u914d\u7f6e\u8def\u5f84\u4e0d\u80fd\u4e3a\u7a7a")
        target: Any = self
        for part in parts[:-1]:
            if isinstance(target, dict):
                target = target[part]
            else:
                target = getattr(target, part)
        last = parts[-1]
        if isinstance(target, dict):
            if last not in target:
                raise KeyError(path)
            target[last] = value
        else:
            if not hasattr(target, last):
                raise KeyError(path)
            setattr(target, last, value)


def _merge_dataclass(instance: Any, raw: dict[str, Any]) -> Any:
    values = asdict(instance)
    values.update({key: value for key, value in raw.items() if key in values})
    return type(instance)(**values)
