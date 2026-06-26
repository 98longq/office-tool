"""Report serialization helpers."""

from __future__ import annotations

import json
from pathlib import Path

from .models import AuditReport


def write_json_report(report: AuditReport, path: str | Path) -> Path:
    output = Path(path).expanduser().resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(report.to_dict(), ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return output


def write_markdown_report(report: AuditReport, path: str | Path) -> Path:
    output = Path(path).expanduser().resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    lines = [f"# 公文校对报告", "", report.summary(), ""]
    lines.extend(["## 识别到的要素", ""])
    if report.elements:
        lines.append("| 要素 | 段落 | 内容 |")
        lines.append("|---|---:|---|")
        for element in report.elements:
            text = element.text.replace("|", "\\|")
            lines.append(f"| {element.role or element.name} | {element.block_index + 1} | {text} |")
    else:
        lines.append("未识别到结构要素。")
    lines.extend(["", "## 问题", ""])
    if report.findings:
        lines.append("| 级别 | 编码 | 段落 | 说明 | 建议 |")
        lines.append("|---|---|---:|---|---|")
        for finding in report.findings:
            block = "" if finding.block_index is None else str(finding.block_index + 1)
            msg = finding.message.replace("|", "\\|")
            suggestion = finding.suggestion.replace("|", "\\|")
            lines.append(f"| {finding.severity} | {finding.code} | {block} | {msg} | {suggestion} |")
    else:
        lines.append("未发现问题。")
    output.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return output
