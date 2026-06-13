"""Command line interface."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

from . import __version__
from .ai import DeepSeekTextReviewer
from .audit import OfficialDocumentAuditor
from .config import OfficeToolConfig
from .formatter import OfficialDocumentFormatter
from .io import load_document
from .reports import write_json_report, write_markdown_report


def main(argv: list[str] | None = None) -> int:
    _configure_stdio()
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        return args.func(args)
    except Exception as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 1


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="OfficeTool 公文审计与格式处理")
    parser.add_argument("--version", action="version", version=f"office-tool {__version__}")
    subparsers = parser.add_subparsers(dest="command", required=True)

    audit = subparsers.add_parser("audit", help="审计文档并输出报告摘要")
    audit.add_argument("input", help="输入 .docx/.txt/.md")
    add_config_args(audit)
    audit.add_argument("--json", dest="json_report", help="写出 JSON 审计报告")
    audit.add_argument("--markdown", dest="markdown_report", help="写出 Markdown 审计报告")
    audit.set_defaults(func=cmd_audit)

    fmt = subparsers.add_parser("format", help="审计并生成格式化后的 .docx")
    fmt.add_argument("input", help="输入 .docx/.txt/.md")
    fmt.add_argument("-o", "--output", required=True, help="输出 .docx")
    add_config_args(fmt)
    fmt.add_argument("--audit-json", help="写出 JSON 审计报告")
    fmt.add_argument("--audit-markdown", help="写出 Markdown 审计报告")
    fmt.set_defaults(func=cmd_format)

    show = subparsers.add_parser("show-config", help="显示默认配置")
    show.set_defaults(func=cmd_show_config)

    init = subparsers.add_parser("init-config", help="生成默认配置 JSON")
    init.add_argument("-o", "--output", default="office_tool_config.json", help="输出配置文件路径")
    init.set_defaults(func=cmd_init_config)
    return parser


def add_config_args(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--config", help="JSON 配置文件")
    parser.add_argument("--set", action="append", default=[], help="覆盖配置，格式 path=value，例如 page.margin_top_cm=3.7")
    parser.add_argument("--ai-review", action="store_true", help="启用 DeepSeek 公文文本 AI 审查")
    parser.add_argument("--ai-base-url", help="内网 DeepSeek OpenAI-compatible base URL，例如 http://host:8000/v1")
    parser.add_argument("--ai-model", help="DeepSeek 模型名，默认 deepseek-chat")
    parser.add_argument("--ai-api-key-env", help="读取 API Key 的环境变量名，默认 DEEPSEEK_API_KEY")


def cmd_audit(args: argparse.Namespace) -> int:
    config = load_config_from_args(args)
    doc, _kind = load_document(args.input)
    report = OfficialDocumentAuditor(config).audit_document(doc)
    maybe_run_ai_review(config, doc, report)
    _write_reports(report, args)
    print(report.summary())
    for finding in report.findings:
        block = "" if finding.block_index is None else f"第 {finding.block_index + 1} 段"
        location = f"{block} " if block else ""
        print(f"[{finding.severity}] {location}{finding.message}")
    return 1 if report.count("error") else 0


def cmd_format(args: argparse.Namespace) -> int:
    config = load_config_from_args(args)
    doc, _kind = load_document(args.input)
    formatter = OfficialDocumentFormatter(config)
    report = formatter.format_document(doc)
    maybe_run_ai_review(config, doc, report)
    output = Path(args.output).expanduser().resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    doc.save(output)
    if args.audit_json:
        write_json_report(report, args.audit_json)
    if args.audit_markdown:
        write_markdown_report(report, args.audit_markdown)
    print(str(output))
    print(report.summary(), file=sys.stderr)
    return 1 if report.count("error") else 0


def cmd_show_config(_args: argparse.Namespace) -> int:
    print(json.dumps(OfficeToolConfig().to_dict(), ensure_ascii=False, indent=2))
    return 0


def cmd_init_config(args: argparse.Namespace) -> int:
    output = Path(args.output).expanduser().resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(OfficeToolConfig().to_dict(), ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(str(output))
    return 0


def load_config_from_args(args: argparse.Namespace) -> OfficeToolConfig:
    raw: dict[str, Any] = {}
    if getattr(args, "config", None):
        raw = json.loads(Path(args.config).expanduser().read_text(encoding="utf-8"))
    config = OfficeToolConfig.from_dict(raw)
    for item in getattr(args, "set", []) or []:
        if "=" not in item:
            raise ValueError(f"--set 必须是 path=value：{item}")
        path, raw_value = item.split("=", 1)
        config.set_path(path.strip(), parse_value(raw_value))
    if getattr(args, "ai_review", False):
        config.ai_review.enabled = True
    if getattr(args, "ai_base_url", None):
        config.ai_review.base_url = args.ai_base_url
        config.ai_review.enabled = True
    if getattr(args, "ai_model", None):
        config.ai_review.model = args.ai_model
    if getattr(args, "ai_api_key_env", None):
        config.ai_review.api_key_env = args.ai_api_key_env
    return config


def maybe_run_ai_review(config: OfficeToolConfig, doc, report) -> None:
    if not config.ai_review.enabled:
        return
    text = "\n".join(paragraph.text for paragraph in doc.paragraphs if paragraph.text.strip())
    DeepSeekTextReviewer(config.ai_review).review_into_report(text, report)


def parse_value(raw: str) -> Any:
    text = raw.strip()
    lowered = text.lower()
    if lowered in {"true", "yes", "on"}:
        return True
    if lowered in {"false", "no", "off"}:
        return False
    if lowered in {"null", "none"}:
        return None
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        return raw


def _write_reports(report, args: argparse.Namespace) -> None:
    if getattr(args, "json_report", None):
        write_json_report(report, args.json_report)
    if getattr(args, "markdown_report", None):
        write_markdown_report(report, args.markdown_report)


def _configure_stdio() -> None:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8")
