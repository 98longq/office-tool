"""Application services for batch document processing."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Iterable

from .ai import DeepSeekTextReviewer
from .audit import OfficialDocumentAuditor
from .config import OfficeToolConfig
from .formatter import OfficialDocumentFormatter
from .io import SUPPORTED_INPUTS, load_document
from .models import AuditReport
from .reports import write_json_report, write_markdown_report


LogCallback = Callable[[str], None]


@dataclass
class DocumentJobResult:
    source: Path
    output: Path | None = None
    json_report: Path | None = None
    markdown_report: Path | None = None
    report: AuditReport | None = None
    error: str = ""

    @property
    def ok(self) -> bool:
        return not self.error


def is_supported_document(path: Path) -> bool:
    return path.is_file() and not path.name.startswith("~$") and path.suffix.lower() in SUPPORTED_INPUTS


def collect_document_inputs(paths: Iterable[str | Path], recursive: bool = True) -> list[Path]:
    inputs: list[Path] = []
    for raw in paths:
        path = Path(raw).expanduser().resolve()
        if path.is_file():
            if not is_supported_document(path):
                raise ValueError(f"不支持的文档格式: {path}")
            inputs.append(path)
            continue
        if path.is_dir():
            iterator = path.rglob("*") if recursive else path.glob("*")
            inputs.extend(sorted(file for file in iterator if is_supported_document(file)))
            continue
        raise FileNotFoundError(f"输入路径不存在: {path}")
    if not inputs:
        raise FileNotFoundError("未找到可处理的文档。支持 .docx/.txt/.md")
    return inputs


def default_output_for(source: Path, output_root: Path | None, multiple: bool) -> Path:
    if output_root is None:
        return source.with_name(f"{source.stem}_formatted.docx")
    if multiple or output_root.suffix.lower() != ".docx":
        return output_root / f"{source.stem}_formatted.docx"
    return output_root


def default_report_for(source: Path, report_dir: Path | None, suffix: str) -> Path | None:
    if report_dir is None:
        return None
    return report_dir / f"{source.stem}_audit.{suffix}"


def run_ai_review_if_enabled(config: OfficeToolConfig, doc, report: AuditReport) -> None:
    if not config.ai_review.enabled:
        return
    text = "\n".join(paragraph.text for paragraph in doc.paragraphs if paragraph.text.strip())
    DeepSeekTextReviewer(config.ai_review).review_into_report(text, report)


def audit_document_path(
    source: str | Path,
    config: OfficeToolConfig,
    json_report: str | Path | None = None,
    markdown_report: str | Path | None = None,
) -> DocumentJobResult:
    source_path = Path(source).expanduser().resolve()
    result = DocumentJobResult(source=source_path)
    try:
        doc, _kind = load_document(source_path)
        report = OfficialDocumentAuditor(config).audit_document(doc)
        run_ai_review_if_enabled(config, doc, report)
        result.report = report
        if json_report:
            result.json_report = write_json_report(report, json_report)
        if markdown_report:
            result.markdown_report = write_markdown_report(report, markdown_report)
    except Exception as exc:
        result.error = str(exc)
    return result


def format_document_path(
    source: str | Path,
    output: str | Path,
    config: OfficeToolConfig,
    json_report: str | Path | None = None,
    markdown_report: str | Path | None = None,
) -> DocumentJobResult:
    source_path = Path(source).expanduser().resolve()
    output_path = Path(output).expanduser().resolve()
    result = DocumentJobResult(source=source_path, output=output_path)
    try:
        doc, _kind = load_document(source_path)
        formatter = OfficialDocumentFormatter(config)
        report = formatter.format_document(doc)
        run_ai_review_if_enabled(config, doc, report)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        doc.save(output_path)
        result.report = report
        if json_report:
            result.json_report = write_json_report(report, json_report)
        if markdown_report:
            result.markdown_report = write_markdown_report(report, markdown_report)
    except Exception as exc:
        result.error = str(exc)
    return result


def audit_many(
    sources: Iterable[str | Path],
    config: OfficeToolConfig,
    report_dir: str | Path | None = None,
    markdown: bool = False,
    log: LogCallback | None = None,
) -> list[DocumentJobResult]:
    source_paths = collect_document_inputs(sources)
    report_root = Path(report_dir).expanduser().resolve() if report_dir else None
    if report_root:
        report_root.mkdir(parents=True, exist_ok=True)

    results: list[DocumentJobResult] = []
    for index, source in enumerate(source_paths, start=1):
        if log:
            log(f"审计 {index}/{len(source_paths)}: {source.name}")
        result = audit_document_path(
            source,
            config,
            json_report=default_report_for(source, report_root, "json"),
            markdown_report=default_report_for(source, report_root, "md") if markdown else None,
        )
        results.append(result)
    return results


def format_many(
    sources: Iterable[str | Path],
    output: str | Path | None,
    config: OfficeToolConfig,
    report_dir: str | Path | None = None,
    markdown: bool = False,
    log: LogCallback | None = None,
) -> list[DocumentJobResult]:
    source_paths = collect_document_inputs(sources)
    output_root = Path(output).expanduser().resolve() if output else None
    report_root = Path(report_dir).expanduser().resolve() if report_dir else None
    multiple = len(source_paths) > 1 or (output_root is not None and output_root.suffix.lower() != ".docx")
    if output_root and multiple:
        output_root.mkdir(parents=True, exist_ok=True)
    if report_root:
        report_root.mkdir(parents=True, exist_ok=True)

    results: list[DocumentJobResult] = []
    for index, source in enumerate(source_paths, start=1):
        if log:
            log(f"格式化 {index}/{len(source_paths)}: {source.name}")
        result = format_document_path(
            source,
            default_output_for(source, output_root, multiple),
            config,
            json_report=default_report_for(source, report_root, "json"),
            markdown_report=default_report_for(source, report_root, "md") if markdown else None,
        )
        results.append(result)
    return results


def summarize_results(results: Iterable[DocumentJobResult]) -> str:
    items = list(results)
    ok = sum(1 for item in items if item.ok)
    failed = len(items) - ok
    return f"完成 {len(items)} 个任务：成功 {ok}，失败 {failed}。"
