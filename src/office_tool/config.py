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
    red_head_line_spacing_pt: float = 42.0
    chars_per_line: int = 28
    lines_per_page: int = 22


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


@dataclass
class AuditOptions:
    profile: str = "auto"
    front_matter_scan_paragraphs: int = 14
    require_document_number_for_red_head: bool = True
    require_signer_for_red_head: bool = False
    require_main_send: bool = False
    require_date: bool = True
    check_page_layout: bool = True
    layout_tolerance_cm: float = 0.2


@dataclass
class FormatOptions:
    apply_page_setup: bool = True
    apply_styles: bool = True
    add_page_number: bool = True
    draw_red_separator: bool = True
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
    timeout_seconds: int = 120
    temperature: float = 0.1
    max_tokens: int = 2048
    max_input_chars: int = 12000


def _default_styles() -> dict[str, StyleSpec]:
    return {
        "red_head": StyleSpec("\u65b9\u6b63\u5c0f\u6807\u5b8b\u7b80\u4f53", 32, "center", color="FF0000"),
        "document_number": StyleSpec("\u4eff\u5b8b_GB2312", 16, "center"),
        "signer": StyleSpec("\u4eff\u5b8b_GB2312", 16, "right"),
        "title": StyleSpec("\u65b9\u6b63\u5c0f\u6807\u5b8b\u7b80\u4f53", 22, "center", line_spacing_pt=33),
        "subtitle": StyleSpec("\u6977\u4f53_GB2312", 16, "center", line_spacing_pt=33),
        "main_send": StyleSpec("\u4eff\u5b8b_GB2312", 16, "left", line_spacing_pt=28),
        "body": StyleSpec("\u4eff\u5b8b_GB2312", 16, "justify", first_line_chars=2, line_spacing_pt=28),
        "h1": StyleSpec("\u9ed1\u4f53", 16, "justify", first_line_chars=2, line_spacing_pt=28),
        "h2": StyleSpec("\u6977\u4f53_GB2312", 16, "justify", first_line_chars=2, line_spacing_pt=28),
        "h3": StyleSpec("\u4eff\u5b8b_GB2312", 16, "justify", first_line_chars=2, line_spacing_pt=28),
        "h4": StyleSpec("\u4eff\u5b8b_GB2312", 16, "justify", first_line_chars=2, line_spacing_pt=28),
        "attachment": StyleSpec("\u4eff\u5b8b_GB2312", 16, "left", line_spacing_pt=28),
        "signatory": StyleSpec("\u4eff\u5b8b_GB2312", 16, "right", line_spacing_pt=28),
        "date": StyleSpec("\u4eff\u5b8b_GB2312", 16, "right", line_spacing_pt=28),
        "copy_to": StyleSpec("\u4eff\u5b8b_GB2312", 14, "left", line_spacing_pt=24),
        "page_number": StyleSpec("\u5b8b\u4f53", 14, "center"),
    }


@dataclass
class OfficeToolConfig:
    page: PageSpec = field(default_factory=PageSpec)
    audit: AuditOptions = field(default_factory=AuditOptions)
    format: FormatOptions = field(default_factory=FormatOptions)
    ai_review: AIReviewOptions = field(default_factory=AIReviewOptions)
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
    values.update(raw)
    return type(instance)(**values)
