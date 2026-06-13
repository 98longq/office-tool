"""Command line interface for OfficeTool."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

from . import __version__
from .config import OfficeToolConfig
from .excel import clean_workbook, inspect_workbook
from .services import (
    audit_document_path,
    audit_many,
    format_document_path,
    format_many,
    summarize_results,
)


def main(argv: list[str] | None = None) -> int:
    configure_stdio()
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        return args.func(args)
    except Exception as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 1


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="OfficeTool 公文审计、格式处理和 Excel 小工具")
    parser.add_argument("--version", action="version", version=f"office-tool {__version__}")
    subparsers = parser.add_subparsers(dest="command", required=True)

    audit = subparsers.add_parser("audit", help="审计单个文档")
    audit.add_argument("input", help="输入 .docx/.txt/.md")
    add_config_args(audit)
    audit.add_argument("--json", dest="json_report", help="写出 JSON 审计报告")
    audit.add_argument("--markdown", dest="markdown_report", help="写出 Markdown 审计报告")
    audit.set_defaults(func=cmd_audit)

    batch_audit = subparsers.add_parser("batch-audit", help="批量审计文件或目录")
    batch_audit.add_argument("inputs", nargs="+", help="输入文件或目录")
    add_config_args(batch_audit)
    batch_audit.add_argument("-r", "--report-dir", required=True, help="报告输出目录")
    batch_audit.add_argument("--markdown", action="store_true", help="同时输出 Markdown 报告")
    batch_audit.set_defaults(func=cmd_batch_audit)

    fmt = subparsers.add_parser("format", help="审计并生成格式化后的 .docx")
    fmt.add_argument("input", help="输入 .docx/.txt/.md")
    fmt.add_argument("-o", "--output", required=True, help="输出 .docx")
    add_config_args(fmt)
    fmt.add_argument("--audit-json", help="写出 JSON 审计报告")
    fmt.add_argument("--audit-markdown", help="写出 Markdown 审计报告")
    fmt.set_defaults(func=cmd_format)

    batch_format = subparsers.add_parser("batch-format", help="批量格式化文件或目录")
    batch_format.add_argument("inputs", nargs="+", help="输入文件或目录")
    batch_format.add_argument("-o", "--output", required=True, help="输出目录")
    add_config_args(batch_format)
    batch_format.add_argument("-r", "--report-dir", help="报告输出目录")
    batch_format.add_argument("--markdown", action="store_true", help="同时输出 Markdown 报告")
    batch_format.set_defaults(func=cmd_batch_format)

    excel = subparsers.add_parser("excel", help="Excel 小工具")
    excel_sub = excel.add_subparsers(dest="excel_command", required=True)

    excel_inspect = excel_sub.add_parser("inspect", help="检查工作簿概况")
    excel_inspect.add_argument("input", help="输入 .xlsx")
    excel_inspect.add_argument("--json", dest="json_report", help="写出 JSON 报告")
    excel_inspect.set_defaults(func=cmd_excel_inspect)

    excel_clean = excel_sub.add_parser("clean", help="清洗工作簿文本空白和空行")
    excel_clean.add_argument("input", help="输入 .xlsx")
    excel_clean.add_argument("-o", "--output", required=True, help="输出 .xlsx")
    excel_clean.add_argument("--keep-empty-rows", action="store_true", help="保留空行")
    excel_clean.set_defaults(func=cmd_excel_clean)

    show = subparsers.add_parser("show-config", help="显示默认配置")
    show.set_defaults(func=cmd_show_config)

    init = subparsers.add_parser("init-config", help="生成默认配置 JSON")
    init.add_argument("-o", "--output", default="office_tool_config.json", help="输出配置文件路径")
    init.set_defaults(func=cmd_init_config)

    gui = subparsers.add_parser("gui", help="启动桌面界面")
    gui.set_defaults(func=cmd_gui)
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
    result = audit_document_path(args.input, config, args.json_report, args.markdown_report)
    print_result(result)
    return 1 if not result.ok or (result.report and result.report.count("error")) else 0


def cmd_batch_audit(args: argparse.Namespace) -> int:
    config = load_config_from_args(args)
    results = audit_many(args.inputs, config, args.report_dir, markdown=args.markdown, log=lambda msg: print(msg, file=sys.stderr))
    for result in results:
        print_result(result)
    print(summarize_results(results), file=sys.stderr)
    return 1 if any(not item.ok for item in results) else 0


def cmd_format(args: argparse.Namespace) -> int:
    config = load_config_from_args(args)
    result = format_document_path(args.input, args.output, config, args.audit_json, args.audit_markdown)
    print_result(result)
    if result.output:
        print(str(result.output))
    return 1 if not result.ok or (result.report and result.report.count("error")) else 0


def cmd_batch_format(args: argparse.Namespace) -> int:
    config = load_config_from_args(args)
    results = format_many(
        args.inputs,
        args.output,
        config,
        report_dir=args.report_dir,
        markdown=args.markdown,
        log=lambda msg: print(msg, file=sys.stderr),
    )
    for result in results:
        print_result(result)
        if result.output:
            print(str(result.output))
    print(summarize_results(results), file=sys.stderr)
    return 1 if any(not item.ok for item in results) else 0


def cmd_excel_inspect(args: argparse.Namespace) -> int:
    summary = inspect_workbook(args.input)
    payload = summary.to_dict()
    if args.json_report:
        output = Path(args.json_report).expanduser().resolve()
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(payload, ensure_ascii=False, indent=2))
    return 0


def cmd_excel_clean(args: argparse.Namespace) -> int:
    summary = clean_workbook(args.input, args.output, remove_empty_rows=not args.keep_empty_rows)
    print(json.dumps(summary.to_dict(), ensure_ascii=False, indent=2))
    print(str(Path(args.output).expanduser().resolve()))
    return 0


def cmd_show_config(_args: argparse.Namespace) -> int:
    print(json.dumps(OfficeToolConfig().to_dict(), ensure_ascii=False, indent=2))
    return 0


def cmd_init_config(args: argparse.Namespace) -> int:
    output = Path(args.output).expanduser().resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(OfficeToolConfig().to_dict(), ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(str(output))
    return 0


def cmd_gui(_args: argparse.Namespace) -> int:
    from .gui import main as gui_main

    return gui_main()


def print_result(result) -> None:
    if result.error:
        print(f"[失败] {result.source}: {result.error}", file=sys.stderr)
        return
    if result.report:
        print(f"{result.source}: {result.report.summary()}")
        for finding in result.report.findings:
            block = "" if finding.block_index is None else f"第 {finding.block_index + 1} 段 "
            print(f"  [{finding.severity}] {block}{finding.message}")


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


def configure_stdio() -> None:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8")
