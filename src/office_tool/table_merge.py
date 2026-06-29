"""Column-based spreadsheet merge engine."""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass
from difflib import SequenceMatcher
from pathlib import Path
from typing import TYPE_CHECKING

from .table_audit import _load_workbook, detect_header_row, headers_for_row
from .table_models import SourceColumnMapping, TableMergeOptions, TableReport

if TYPE_CHECKING:
    from openpyxl.worksheet.worksheet import Worksheet


@dataclass(frozen=True)
class SourceValue:
    value: str
    workbook: Path
    sheet: str
    row: int


def merge_by_columns(options: TableMergeOptions) -> TableReport:
    """Merge source column values into a master sheet by matching key columns."""
    master_path = Path(options.master_path).expanduser().resolve()
    output_path = Path(options.output_path).expanduser().resolve()
    master_wb = _load_workbook(master_path, data_only=False)
    if options.master_sheet not in master_wb.sheetnames:
        raise KeyError(f"主表不存在工作表: {options.master_sheet}")
    master_ws = master_wb[options.master_sheet]
    master_header_row = options.master_header_row or detect_header_row(master_ws)
    master_headers = headers_for_row(master_ws, master_header_row)
    master_key_col = resolve_column(master_ws, master_header_row, options.master_key_column)
    master_target_col = resolve_column(master_ws, master_header_row, options.master_target_column)

    report = TableReport()
    report.stats["master"] = str(master_path)
    report.stats["output"] = str(output_path)
    report.stats["sources"] = len(options.sources)
    report.stats["updated_cells"] = 0
    report.stats["appended_values"] = 0
    report.sheets.append(
        _sheet_info(str(master_path), master_ws, master_header_row, master_headers)
    )

    source_values = _collect_source_values(options, report)
    match_plan = _build_key_match_plan(master_ws, master_header_row, master_key_col, source_values, options, report)
    if report.count("error"):
        raise ValueError("模糊匹配存在一对多或多对一歧义，请调整匹配度或关闭模糊匹配后重试。")
    used_source_keys: set[tuple[str, str, int]] = set()
    for row in range(master_header_row + 1, master_ws.max_row + 1):
        key = normalized_cell_text(master_ws.cell(row, master_key_col).value, options.normalize_keys)
        if not key:
            report.add_finding(
                "empty_master_key",
                "warning",
                "主表校验列为空，已跳过该行。",
                workbook=master_path,
                sheet=master_ws.title,
                row=row,
                column=master_key_col,
            )
            continue
        matches = match_plan.get(row, [])
        if not matches:
            report.add_finding(
                "unmatched_master_key",
                "info",
                "主表校验值未在副表中找到匹配数据。",
                workbook=master_path,
                sheet=master_ws.title,
                row=row,
                column=master_key_col,
                actual=master_ws.cell(row, master_key_col).value or "",
            )
            continue
        target_cell = master_ws.cell(row, master_target_col)
        existing = split_cell_values(target_cell.value, options.separator)
        additions: list[str] = []
        for match in matches:
            used_source_keys.add((str(match.workbook), match.sheet, match.row))
            if match.value and match.value not in existing and match.value not in additions:
                additions.append(match.value)
        if not additions:
            continue
        target_cell.value = options.separator.join(existing + additions) if existing else options.separator.join(additions)
        report.stats["updated_cells"] += 1
        report.stats["appended_values"] += len(additions)
        if len(additions) > 1 or existing:
            report.add_finding(
                "merged_multiple_values",
                "info",
                "同一主表单元格合并了多条不同内容。",
                workbook=master_path,
                sheet=master_ws.title,
                row=row,
                column=master_target_col,
                actual=target_cell.value or "",
            )

    for key, values in source_values.items():
        for value in values:
            marker = (str(value.workbook), value.sheet, value.row)
            if marker not in used_source_keys:
                report.add_finding(
                    "unused_source_value",
                    "info",
                    "副表数据未匹配到主表行。",
                    workbook=value.workbook,
                    sheet=value.sheet,
                    row=value.row,
                    actual=f"{key}: {value.value}",
                )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    master_wb.save(output_path)
    return report


def _build_key_match_plan(
    master_ws: "Worksheet",
    master_header_row: int,
    master_key_col: int,
    source_values: dict[str, list[SourceValue]],
    options: TableMergeOptions,
    report: TableReport,
) -> dict[int, list[SourceValue]]:
    if not options.fuzzy_match:
        plan: dict[int, list[SourceValue]] = {}
        for row in range(master_header_row + 1, master_ws.max_row + 1):
            key = normalized_cell_text(master_ws.cell(row, master_key_col).value, options.normalize_keys)
            if key:
                plan[row] = source_values.get(key, [])
        return plan

    threshold = max(0, min(100, int(options.fuzzy_threshold)))
    master_keys: dict[int, str] = {}
    best_by_row: dict[int, tuple[str, float]] = {}
    rows_by_source_key: dict[str, list[int]] = {}
    for row in range(master_header_row + 1, master_ws.max_row + 1):
        key = normalized_cell_text(master_ws.cell(row, master_key_col).value, options.normalize_keys)
        if not key:
            continue
        master_keys[row] = key
        candidates: list[tuple[str, float]] = []
        for source_key in source_values:
            score = _similarity_score(key, source_key)
            if score >= threshold:
                candidates.append((source_key, score))
        if not candidates:
            continue
        candidates.sort(key=lambda item: item[1], reverse=True)
        top_score = candidates[0][1]
        top_keys = [item[0] for item in candidates if item[1] == top_score]
        if len(top_keys) > 1:
            report.add_finding(
                "ambiguous_fuzzy_match",
                "error",
                "模糊匹配出现一对多候选，已停止汇总。",
                sheet=master_ws.title,
                row=row,
                column=master_key_col,
                actual=key,
                suggestion="提高匹配度，或手动统一主表和副表校验文字。",
            )
            continue
        source_key = top_keys[0]
        if len(source_values.get(source_key, [])) > 1:
            report.add_finding(
                "ambiguous_fuzzy_match",
                "error",
                "模糊匹配命中多条副表数据，已停止汇总。",
                sheet=master_ws.title,
                row=row,
                column=master_key_col,
                actual=f"{key} ≈ {source_key}",
                suggestion="提高匹配度，或先处理副表重复校验值。",
            )
            continue
        best_by_row[row] = (source_key, top_score)
        rows_by_source_key.setdefault(source_key, []).append(row)

    for source_key, rows in rows_by_source_key.items():
        if len(rows) <= 1:
            continue
        report.add_finding(
            "ambiguous_fuzzy_match",
            "error",
            "模糊匹配出现多对一候选，已停止汇总。",
            sheet=master_ws.title,
            actual=source_key,
            suggestion="提高匹配度，或手动统一主表和副表校验文字。",
        )

    if report.count("error"):
        return {}

    plan: dict[int, list[SourceValue]] = {}
    for row, (source_key, score) in best_by_row.items():
        plan[row] = source_values.get(source_key, [])
        if master_keys[row] != source_key:
            report.add_finding(
                "fuzzy_key_match",
                "info",
                "已使用模糊匹配关联主表和副表校验值。",
                sheet=master_ws.title,
                row=row,
                column=master_key_col,
                actual=f"{master_keys[row]} ≈ {source_key} ({score:.0f}%)",
            )
    return plan


def _similarity_score(left: str, right: str) -> float:
    if not left or not right:
        return 0.0
    ratio = SequenceMatcher(None, left, right).ratio()
    if left in right or right in left:
        ratio = max(ratio, min(len(left), len(right)) / max(len(left), len(right)))
    return ratio * 100


def merge_same_layout(
    master_path: str | Path,
    source_paths: list[str | Path],
    output_path: str | Path,
    *,
    master_sheet: str | None = None,
    source_sheet: str | None = None,
    separator: str = "\n",
) -> TableReport:
    """Merge same-layout workbooks by cell position, preserving the master workbook."""
    master_path = Path(master_path).expanduser().resolve()
    output_path = Path(output_path).expanduser().resolve()
    master_wb = _load_workbook(master_path, data_only=False)
    master_ws = master_wb[master_sheet] if master_sheet else master_wb.worksheets[0]

    report = TableReport()
    report.stats["master"] = str(master_path)
    report.stats["output"] = str(output_path)
    report.stats["sources"] = len(source_paths)
    report.stats["updated_cells"] = 0
    report.stats["appended_values"] = 0
    header_row = detect_header_row(master_ws)
    report.sheets.append(_sheet_info(str(master_path), master_ws, header_row, headers_for_row(master_ws, header_row)))

    for raw_source in source_paths:
        source_path = Path(raw_source).expanduser().resolve()
        workbook = _load_workbook(source_path, data_only=False)
        if source_sheet:
            if source_sheet not in workbook.sheetnames:
                report.add_finding(
                    "missing_source_sheet",
                    "warning",
                    "副表不存在指定工作表，已使用第一个工作表。",
                    workbook=source_path,
                    sheet=source_sheet,
                    suggestion="如该副表格式特殊，请在高级设置中单独配置。",
                )
                sheet = workbook.worksheets[0]
            else:
                sheet = workbook[source_sheet]
        else:
            sheet = workbook.worksheets[0]
        source_header_row = detect_header_row(sheet)
        report.sheets.append(_sheet_info(str(source_path), sheet, source_header_row, headers_for_row(sheet, source_header_row)))
        max_row = min(master_ws.max_row, sheet.max_row)
        max_col = min(master_ws.max_column, sheet.max_column)
        for row in range(1, max_row + 1):
            for column in range(1, max_col + 1):
                value = normalized_cell_text(sheet.cell(row, column).value, normalize=False)
                if not value:
                    continue
                target_cell = master_ws.cell(row, column)
                existing = split_cell_values(target_cell.value, separator)
                if value in existing:
                    continue
                target_cell.value = separator.join(existing + [value]) if existing else value
                report.stats["updated_cells"] += 1
                report.stats["appended_values"] += 1

    output_path.parent.mkdir(parents=True, exist_ok=True)
    master_wb.save(output_path)
    return report


def resolve_column(sheet: "Worksheet", header_row: int, column: str | int) -> int:
    if isinstance(column, int):
        if column < 1:
            raise ValueError("列序号必须从 1 开始。")
        return column
    text = str(column).strip()
    if not text:
        raise ValueError("列不能为空。")
    if re.fullmatch(r"[A-Za-z]{1,3}", text):
        try:
            from openpyxl.utils.cell import column_index_from_string
        except ImportError as exc:
            raise RuntimeError("表格功能需要安装 openpyxl。") from exc
        return column_index_from_string(text.upper())
    headers = headers_for_row(sheet, header_row)
    for index, header in enumerate(headers, start=1):
        if header == text:
            return index
    raise KeyError(f"未找到列标题: {text}")


def normalized_cell_text(value, normalize: bool = True) -> str:
    if value is None:
        return ""
    text = str(value)
    if not normalize:
        return text.strip()
    text = unicodedata.normalize("NFKC", text)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def split_cell_values(value, separator: str) -> list[str]:
    text = "" if value is None else str(value)
    if not text.strip():
        return []
    return [part.strip() for part in text.split(separator) if part.strip()]


def _collect_source_values(options: TableMergeOptions, report: TableReport) -> dict[str, list[SourceValue]]:
    values_by_key: dict[str, list[SourceValue]] = {}
    for source in options.sources:
        source_path = Path(source.path).expanduser().resolve()
        workbook = _load_workbook(source_path, data_only=False)
        if source.sheet not in workbook.sheetnames:
            raise KeyError(f"副表不存在工作表: {source.sheet}")
        sheet = workbook[source.sheet]
        header_row = source.header_row or detect_header_row(sheet)
        report.sheets.append(_sheet_info(str(source_path), sheet, header_row, headers_for_row(sheet, header_row)))
        key_col = resolve_column(sheet, header_row, source.key_column)
        value_col = resolve_column(sheet, header_row, source.value_column)
        seen_keys: set[str] = set()
        for row in range(header_row + 1, sheet.max_row + 1):
            key = normalized_cell_text(sheet.cell(row, key_col).value, options.normalize_keys)
            value = normalized_cell_text(sheet.cell(row, value_col).value, normalize=False)
            if not key:
                continue
            if key in seen_keys:
                report.add_finding(
                    "duplicate_source_key",
                    "warning",
                    "同一副表内校验值重复，后续会按多条候选合并。",
                    workbook=source_path,
                    sheet=sheet.title,
                    row=row,
                    column=key_col,
                    actual=key,
                )
            seen_keys.add(key)
            if not value:
                continue
            bucket = values_by_key.setdefault(key, [])
            if value not in {item.value for item in bucket}:
                bucket.append(SourceValue(value=value, workbook=source_path, sheet=sheet.title, row=row))
    return values_by_key


def _sheet_info(workbook_path: str, sheet: "Worksheet", header_row: int, headers: list[str]):
    from .table_models import SheetInfo

    return SheetInfo(
        workbook=workbook_path,
        sheet=sheet.title,
        max_row=sheet.max_row,
        max_column=sheet.max_column,
        header_row=header_row,
        headers=headers,
        merged_ranges=[str(item) for item in sheet.merged_cells.ranges],
    )
