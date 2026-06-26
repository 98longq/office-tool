"""DeepSeek-compatible AI text review."""

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


SYSTEM_PROMPT = (
    "你是严谨的中文公文校对助手。只校对用户提供的公文文本，"
    "重点关注文种是否匹配、标题是否准确、主送机关、附件、日期是否前后一致，"
    "表述是否符合党政机关公文风格，是否存在语病歧义、责任边界不清、"
    "数字日期前后矛盾、过度承诺或不适宜措辞。"
)

USER_PROMPT_TEMPLATE = (
    "请以 JSON 数组返回校对结果，不要输出额外说明。"
    "category \u8bf7\u4f7f\u7528\u4e2d\u6587\u9519\u8bef\u7c7b\u578b\uff0c\u4f18\u5148\u4ece"
    "\u683c\u5f0f\u9519\u8bef\u3001\u6f0f\u5b57\u9519\u5b57\u3001\u8bed\u53e5\u4e0d\u5f53\u3001\u8bed\u4e49\u6b67\u4e49\u3001"
    "\u524d\u540e\u4e0d\u4e00\u81f4\u3001\u98ce\u9669\u8868\u8ff0\u3001\u5176\u4ed6\u5efa\u8bae\u4e2d\u9009\u62e9\u3002"
    "\u6bcf\u9879\u5b57\u6bb5\u4e3a category\u3001severity\u3001message\u3001quote\u3001suggestion\u3002"
    "severity \u53ea\u80fd\u662f error\u3001warning\u3001info\u3002"
    "\u6ca1\u6709\u95ee\u9898\u65f6\u8fd4\u56de []\u3002\n\n"
    "\u516c\u6587\u6587\u672c\uff1a\n{text}"
)


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
        if self.options.strip_reasoning:
            content = strip_reasoning_text(content)
        return self.parse_findings(content)

    def _build_payload(self, text: str) -> dict[str, Any]:
        clipped = text[: max(1, self.options.max_input_chars)]
        return {
            "model": self.options.model,
            "temperature": self.options.temperature,
            "max_tokens": self.options.max_tokens,
            "stream": self.options.stream,
            "messages": [
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": USER_PROMPT_TEMPLATE.format(text=clipped)},
            ],
        }

    def _post(self, payload: dict[str, Any]) -> dict[str, Any]:
        url = self._resolve_url()
        request = urllib.request.Request(
            url,
            data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
            headers=self._headers(),
            method="POST",
        )
        try:
            with urllib.request.urlopen(request, timeout=self.options.timeout_seconds) as response:
                raw = response.read().decode("utf-8", errors="replace")
                if payload.get("stream"):
                    return stream_text_to_chat_completion(raw)
                return json.loads(raw)
        except urllib.error.HTTPError as exc:
            detail = exc.read().decode("utf-8", errors="ignore")
            raise RuntimeError(f"DeepSeek 校对请求失败：HTTP {exc.code} {detail}") from exc
        except urllib.error.URLError as exc:
            raise RuntimeError(f"DeepSeek 校对请求失败：{exc.reason}") from exc

    def _resolve_url(self) -> str:
        base_url = self.options.base_url.strip().rstrip("/")
        if not base_url:
            raise ValueError("AI 校对已启用，但 DeepSeek 服务地址为空。")
        endpoint_path = self.options.endpoint_path.strip()
        if not endpoint_path:
            return base_url
        normalized_endpoint = endpoint_path if endpoint_path.startswith("/") else f"/{endpoint_path}"
        if base_url.endswith(normalized_endpoint.rstrip("/")) or base_url.endswith("/chat/completions"):
            return base_url
        return base_url + normalized_endpoint

    def _headers(self) -> dict[str, str]:
        headers = {
            "Content-Type": "application/json",
            "Accept": "application/json",
            "Connection": "keep-alive",
        }
        api_key = self.options.api_key or os.environ.get(self.options.api_key_env, "")
        if api_key:
            prefix = self.options.auth_prefix.strip()
            headers["Authorization"] = f"{prefix} {api_key}" if prefix else api_key
        return headers

    @staticmethod
    def _extract_content(response: dict[str, Any]) -> str:
        try:
            choice = response["choices"][0]
            if "message" in choice:
                return choice["message"]["content"]
            if "delta" in choice:
                return choice["delta"].get("content", "")
        except (KeyError, IndexError, TypeError) as exc:
            raise RuntimeError("DeepSeek response is not a supported chat-completions shape.") from exc
        raise RuntimeError("DeepSeek response does not contain assistant content.")

    @staticmethod
    def parse_findings(content: str) -> list[AITextFinding]:
        raw = extract_json_text(content)
        try:
            data = json.loads(raw)
        except json.JSONDecodeError:
            return [
                AITextFinding(
                    category="原始建议",
                    severity="info",
                    message="AI 未返回标准 JSON，请查看原始建议。",
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
                    category=str(item.get("category", "其他建议")),
                    severity=severity,
                    message=str(item.get("message", "")).strip() or "AI 校对建议。",
                    quote=str(item.get("quote", "")).strip(),
                    suggestion=str(item.get("suggestion", "")).strip(),
                )
            )
        return findings


def stream_text_to_chat_completion(raw: str) -> dict[str, Any]:
    pieces: list[str] = []
    for line in raw.splitlines():
        text = line.strip()
        if not text:
            continue
        if text.startswith("data:"):
            text = text[5:].strip()
        if text == "[DONE]":
            continue
        try:
            chunk = json.loads(text)
        except json.JSONDecodeError:
            continue
        for choice in chunk.get("choices", []):
            delta = choice.get("delta") or {}
            message = choice.get("message") or {}
            pieces.append(delta.get("content") or message.get("content") or "")
    return {"choices": [{"message": {"role": "assistant", "content": "".join(pieces)}}]}


def strip_reasoning_text(content: str) -> str:
    return re.sub(r"<think>.*?</think>", "", content, flags=re.DOTALL | re.IGNORECASE).strip()


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
    return code or "proofreading"
