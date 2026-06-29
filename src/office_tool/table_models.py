"""Data models for spreadsheet inspection and merging."""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any


@dataclass
class TableFinding:
    code: str
    severity: str
    message: str
    workbook: str = ""
    sheet: str = ""
    row: int | None = None
    column: int | None = None
    actual: str = ""
    suggestion: str = ""

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class SheetInfo:
    workbook: str
    sheet: str
    max_row: int
    max_column: int
    header_row: int
    headers: list[str] = field(default_factory=list)
    merged_ranges: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class TableReport:
    sheets: list[SheetInfo] = field(default_factory=list)
    findings: list[TableFinding] = field(default_factory=list)
    stats: dict[str, Any] = field(default_factory=dict)

    def add_finding(
        self,
        code: str,
        severity: str,
        message: str,
        workbook: str | Path = "",
        sheet: str = "",
        row: int | None = None,
        column: int | None = None,
        actual: str = "",
        suggestion: str = "",
    ) -> None:
        self.findings.append(
            TableFinding(
                code=code,
                severity=severity,
                message=message,
                workbook=str(workbook),
                sheet=sheet,
                row=row,
                column=column,
                actual=actual[:200],
                suggestion=suggestion,
            )
        )

    def count(self, severity: str) -> int:
        return sum(1 for finding in self.findings if finding.severity == severity)

    def summary(self) -> str:
        if not self.findings:
            return f"表格处理完成：{len(self.sheets)} 个工作表，未发现需要提示的问题。"
        return (
            f"表格处理完成：{len(self.sheets)} 个工作表，"
            f"错误 {self.count('error')}，警告 {self.count('warning')}，提示 {self.count('info')}。"
        )

    def to_dict(self) -> dict[str, Any]:
        return {
            "summary": self.summary(),
            "stats": dict(self.stats),
            "sheets": [sheet.to_dict() for sheet in self.sheets],
            "findings": [finding.to_dict() for finding in self.findings],
        }


@dataclass(frozen=True)
class SourceColumnMapping:
    path: Path
    sheet: str
    key_column: str | int
    value_column: str | int
    header_row: int | None = None


@dataclass(frozen=True)
class TableMergeOptions:
    master_path: Path
    output_path: Path
    master_sheet: str
    master_key_column: str | int
    master_target_column: str | int
    sources: list[SourceColumnMapping]
    master_header_row: int | None = None
    normalize_keys: bool = True
    fuzzy_match: bool = False
    fuzzy_threshold: int = 90
    separator: str = "\n"
