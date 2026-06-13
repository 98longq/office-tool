"""DeepSeek-compatible AI text review.

The client intentionally uses the OpenAI-compatible chat-completions shape
because most private DeepSeek deployments expose that contract. It does not run
unless explicitly enabled by config or CLI flags.
"""

from __future__ import annotations

import json
import os
import re
import urllib.error
import urllib.request
from dataclasses import dataclass
from typing import Any

from ..config import AIReviewOptions
from ..models import AuditReport


@dataclass
class AITextFinding:
    category: str
    severity: str
    message: str
    quote: str = ""
    suggestion: str = ""


class DeepSeekTextReviewer:
    """Review official-document text with an internal DeepSeek endpoint."""

    def __init__(self, options: AIReviewOptions):
        self.options = options

    def review_into_report(self, text: str, report: AuditReport) -> None:
        if not self.options.enabled:
            return
        findings = self.review_text(text)
        for item in findings:
            report.add_finding(
                code=f"ai_{safe_code(item.category)}",
                severity=item.severity,
                message=item.message,
                text=item.quote,
                suggestion=item.suggestion,
                can_fix=False,
            )
        report.stats["ai_review_provider"] = "deepseek"
        report.stats["ai_review_findings"] = len(findings)

    def review_text(self, text: str) -> list[AITextFinding]:
        payload = self._build_payload(text)
        response = self._post(payload)
        content = self._extract_content(response)
        return self.parse_findings(content)

    def _build_payload(self, text: str) -> dict[str, Any]:
        clipped = text[: max(1, self.options.max_input_chars)]
        return {
            "model": self.options.model,
            "temperature": self.options.temperature,
            "max_tokens": self.options.max_tokens,
            "messages": [
                {
                    "role": "system",
                    "content": (
                        "你是严谨的中文公文审核助手。只审查用户提供的公文文本，"
                        "重点关注文种是否匹配、标题是否准确、主送/附件/日期是否前后一致、"
                        "表述是否符合党政机关公文风格、是否存在语病歧义、责任边界不清、"
                        "数字日期前后矛盾、过度承诺或不适宜措辞。"
                    ),
                },
                {
                    "role": "user",
                    "content": (
                        "请以 JSON 数组返回审查结果，不要输出额外说明。每项字段为："
                        "category、severity、message、quote、suggestion。"
                        "severity 只能是 error、warning、info。"
                        "没有问题时返回 []。\n\n公文文本：\n"
                        f"{clipped}"
                    ),
                },
            ],
        }

    def _post(self, payload: dict[str, Any]) -> dict[str, Any]:
        base_url = self.options.base_url.rstrip("/")
        if not base_url:
            raise ValueError("AI 审查已启用，但未配置 DeepSeek base_url。")
        url = base_url + self.options.endpoint_path
        headers = {"Content-Type": "application/json"}
        api_key = self.options.api_key or os.environ.get(self.options.api_key_env, "")
        if api_key:
            headers["Authorization"] = f"Bearer {api_key}"

        request = urllib.request.Request(
            url,
            data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
            headers=headers,
            method="POST",
        )
        try:
            with urllib.request.urlopen(request, timeout=self.options.timeout_seconds) as response:
                return json.loads(response.read().decode("utf-8"))
        except urllib.error.HTTPError as exc:
            detail = exc.read().decode("utf-8", errors="ignore")
            raise RuntimeError(f"DeepSeek 审查请求失败：HTTP {exc.code} {detail}") from exc
        except urllib.error.URLError as exc:
            raise RuntimeError(f"DeepSeek 审查请求失败：{exc.reason}") from exc

    @staticmethod
    def _extract_content(response: dict[str, Any]) -> str:
        try:
            return response["choices"][0]["message"]["content"]
        except (KeyError, IndexError, TypeError) as exc:
            raise RuntimeError("DeepSeek 返回结构不符合 chat-completions 格式。") from exc

    @staticmethod
    def parse_findings(content: str) -> list[AITextFinding]:
        raw = extract_json_text(content)
        try:
            data = json.loads(raw)
        except json.JSONDecodeError:
            return [
                AITextFinding(
                    category="raw_review",
                    severity="info",
                    message="AI 返回了非 JSON 审查意见，请人工查看原文。",
                    quote="",
                    suggestion=content.strip()[:500],
                )
            ]
        if not isinstance(data, list):
            data = [data]
        findings: list[AITextFinding] = []
        for item in data:
            if not isinstance(item, dict):
                continue
            severity = str(item.get("severity", "info")).lower()
            if severity not in {"error", "warning", "info"}:
                severity = "info"
            findings.append(
                AITextFinding(
                    category=str(item.get("category", "review")),
                    severity=severity,
                    message=str(item.get("message", "")).strip() or "AI 审查提示。",
                    quote=str(item.get("quote", "")).strip(),
                    suggestion=str(item.get("suggestion", "")).strip(),
                )
            )
        return findings


def extract_json_text(content: str) -> str:
    text = content.strip()
    fenced = re.search(r"```(?:json)?\s*(.*?)```", text, re.DOTALL | re.IGNORECASE)
    if fenced:
        return fenced.group(1).strip()
    start_array = text.find("[")
    end_array = text.rfind("]")
    if start_array != -1 and end_array > start_array:
        return text[start_array : end_array + 1]
    start_obj = text.find("{")
    end_obj = text.rfind("}")
    if start_obj != -1 and end_obj > start_obj:
        return text[start_obj : end_obj + 1]
    return text


def safe_code(value: str) -> str:
    code = re.sub(r"[^0-9A-Za-z_\u4e00-\u9fff]+", "_", value.strip()).strip("_")
    return code or "review"
