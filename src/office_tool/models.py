"""Shared data models for audit reports."""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
from typing import Any


@dataclass
class DocumentElement:
    name: str
    block_index: int
    text: str
    role: str = ""

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class AuditFinding:
    code: str
    severity: str
    message: str
    block_index: int | None = None
    text: str = ""
    expected: str = ""
    actual: str = ""
    suggestion: str = ""
    can_fix: bool = False

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class AuditReport:
    profile: str = "standard"
    is_red_head: bool = False
    elements: list[DocumentElement] = field(default_factory=list)
    findings: list[AuditFinding] = field(default_factory=list)
    stats: dict[str, Any] = field(default_factory=dict)

    def add_element(self, name: str, block_index: int, text: str, role: str = "") -> None:
        self.elements.append(DocumentElement(name, block_index, text[:120], role))

    def add_finding(
        self,
        code: str,
        severity: str,
        message: str,
        block_index: int | None = None,
        text: str = "",
        expected: str = "",
        actual: str = "",
        suggestion: str = "",
        can_fix: bool = False,
    ) -> None:
        self.findings.append(
            AuditFinding(
                code=code,
                severity=severity,
                message=message,
                block_index=block_index,
                text=text[:120],
                expected=expected,
                actual=actual,
                suggestion=suggestion,
                can_fix=can_fix,
            )
        )

    def count(self, severity: str) -> int:
        return sum(1 for finding in self.findings if finding.severity == severity)

    def summary(self) -> str:
        if not self.findings:
            return f"审计完成：{self.profile}，未发现问题。"
        return (
            f"审计完成：{self.profile}，"
            f"错误 {self.count('error')}，警告 {self.count('warning')}，提示 {self.count('info')}。"
        )

    def to_dict(self) -> dict[str, Any]:
        return {
            "profile": self.profile,
            "is_red_head": self.is_red_head,
            "summary": self.summary(),
            "stats": dict(self.stats),
            "elements": [element.to_dict() for element in self.elements],
            "findings": [finding.to_dict() for finding in self.findings],
        }


@dataclass
class DetectedStructure:
    copy_number: int | None = None
    secrecy: int | None = None
    urgency: int | None = None
    red_head: int | None = None
    document_number: int | None = None
    signer: int | None = None
    title: int | None = None
    main_send: int | None = None
    body_start: int | None = None
    attachment_notes: list[int] = field(default_factory=list)
    signatory: int | None = None
    date: int | None = None
    copy_to: int | None = None
    print_org_date: int | None = None

    def front_matter_end(self) -> int:
        values = [
            self.copy_number,
            self.secrecy,
            self.urgency,
            self.red_head,
            self.document_number,
            self.signer,
        ]
        present = [value for value in values if value is not None]
        return max(present) if present else -1
