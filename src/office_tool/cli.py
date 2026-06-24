"""Command line interface for OfficeTool."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

from . import __version__
from .config import OfficeToolConfig
from .services import audit_document_path, audit_many, format_document_path, format_many, summarize_results


def main(argv: list[str] | None = None) -> int:
    configure_stdio()
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        return args.func(args)
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="OfficeTool document proofreading and formatting tools.")
    parser.add_argument("--version", action="version", version=f"office-tool {__version__}")
    subparsers = parser.add_subparsers(dest="command", required=True)

    proofread = subparsers.add_parser("proofread", help="Proofread one .doc/.docx/.txt/.md file.")
    proofread.add_argument("input", help="Input .doc/.docx/.txt/.md file.")
    add_config_args(proofread)
    proofread.add_argument("--json", dest="json_report", help="Write JSON proofreading report.")
    proofread.add_argument("--markdown", dest="markdown_report", help="Write Markdown proofreading report.")
    proofread.set_defaults(func=cmd_proofread)

    batch_proofread = subparsers.add_parser("batch-proofread", help="Proofread files or folders.")
    batch_proofread.add_argument("inputs", nargs="+", help="Input files or folders.")
    add_config_args(batch_proofread)
    batch_proofread.add_argument("-r", "--report-dir", required=True, help="Report output folder.")
    batch_proofread.add_argument("--markdown", action="store_true", help="Also write Markdown reports.")
    batch_proofread.set_defaults(func=cmd_batch_proofread)

    fmt = subparsers.add_parser("format", help="Proofread and generate a formatted .docx file.")
    fmt.add_argument("input", help="Input .doc/.docx/.txt/.md file.")
    fmt.add_argument("-o", "--output", required=True, help="Output .docx file.")
    add_config_args(fmt)
    fmt.add_argument("--proofreading-json", help="Write JSON proofreading report.")
    fmt.add_argument("--proofreading-markdown", help="Write Markdown proofreading report.")
    fmt.set_defaults(func=cmd_format)

    batch_format = subparsers.add_parser("batch-format", help="Format files or folders.")
    batch_format.add_argument("inputs", nargs="+", help="Input files or folders.")
    batch_format.add_argument("-o", "--output", required=True, help="Output folder.")
    add_config_args(batch_format)
    batch_format.add_argument("-r", "--report-dir", help="Report output folder.")
    batch_format.add_argument("--markdown", action="store_true", help="Also write Markdown reports.")
    batch_format.set_defaults(func=cmd_batch_format)

    show = subparsers.add_parser("show-config", help="Print default config.")
    show.set_defaults(func=cmd_show_config)

    init = subparsers.add_parser("init-config", help="Write default config JSON.")
    init.add_argument("-o", "--output", default="office_tool_config.json", help="Output config path.")
    init.set_defaults(func=cmd_init_config)

    gui = subparsers.add_parser("gui", help="Start desktop GUI.")
    gui.set_defaults(func=cmd_gui)
    return parser


def add_config_args(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--config", help="JSON config file.")
    parser.add_argument("--set", action="append", default=[], help="Override config path=value, e.g. page.margin_top_cm=3.7")
    parser.add_argument("--ai-review", action="store_true", help="Enable DeepSeek text review.")
    parser.add_argument("--ai-base-url", help="DeepSeek base URL or full /chat/completions URL.")
    parser.add_argument("--ai-model", help="DeepSeek model name, e.g. DeepSeek-R1.")
    parser.add_argument("--ai-api-key", help="Authorization token. Prefer env vars for shared machines.")
    parser.add_argument("--ai-api-key-env", help="Environment variable for API key, default DEEPSEEK_API_KEY.")
    parser.add_argument("--ai-auth-prefix", help="Authorization prefix. Use empty string for raw token.")
    parser.add_argument("--ai-stream", action="store_true", help="Parse line-delimited streaming response.")


def cmd_proofread(args: argparse.Namespace) -> int:
    config = load_config_from_args(args)
    result = audit_document_path(args.input, config, args.json_report, args.markdown_report)
    print_result(result)
    return 1 if not result.ok or (result.report and result.report.count("error")) else 0


def cmd_batch_proofread(args: argparse.Namespace) -> int:
    config = load_config_from_args(args)
    results = audit_many(args.inputs, config, args.report_dir, markdown=args.markdown, log=lambda msg: print(msg, file=sys.stderr))
    for result in results:
        print_result(result)
    print(summarize_results(results), file=sys.stderr)
    return 1 if any(not item.ok for item in results) else 0


def cmd_format(args: argparse.Namespace) -> int:
    config = load_config_from_args(args)
    result = format_document_path(args.input, args.output, config, args.proofreading_json, args.proofreading_markdown)
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
        print(f"[failed] {result.source}: {result.error}", file=sys.stderr)
        return
    if result.report:
        print(f"{result.source}: {result.report.summary()}")
        for finding in result.report.findings:
            block = "" if finding.block_index is None else f"block {finding.block_index + 1} "
            print(f"  [{finding.severity}] {block}{finding.message}")


def load_config_from_args(args: argparse.Namespace) -> OfficeToolConfig:
    raw: dict[str, Any] = {}
    if getattr(args, "config", None):
        raw = json.loads(Path(args.config).expanduser().read_text(encoding="utf-8"))
    config = OfficeToolConfig.from_dict(raw)
    for item in getattr(args, "set", []) or []:
        if "=" not in item:
            raise ValueError(f"--set must be path=value: {item}")
        path, raw_value = item.split("=", 1)
        config.set_path(path.strip(), parse_value(raw_value))
    if getattr(args, "ai_review", False):
        config.ai_review.enabled = True
    if getattr(args, "ai_base_url", None):
        config.ai_review.base_url = args.ai_base_url
        config.ai_review.enabled = True
    if getattr(args, "ai_model", None):
        config.ai_review.model = args.ai_model
    if getattr(args, "ai_api_key", None):
        config.ai_review.api_key = args.ai_api_key
    if getattr(args, "ai_api_key_env", None):
        config.ai_review.api_key_env = args.ai_api_key_env
    if getattr(args, "ai_auth_prefix", None) is not None:
        config.ai_review.auth_prefix = args.ai_auth_prefix
    if getattr(args, "ai_stream", False):
        config.ai_review.stream = True
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
