"""Minimal desktop GUI for OfficeTool."""

from __future__ import annotations

import json
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from .audit import OfficialDocumentAuditor
from .cli import load_config_from_args
from .config import OfficeToolConfig
from .formatter import OfficialDocumentFormatter
from .io import load_document
from .reports import write_json_report, write_markdown_report


class OfficeToolGUI:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("OfficeTool 公文审计助手")
        self.root.geometry("920x640")
        self.root.minsize(820, 560)

        self.input_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.json_report_var = tk.StringVar()
        self.markdown_report_var = tk.StringVar()
        self.ai_enabled_var = tk.BooleanVar(value=False)
        self.ai_base_url_var = tk.StringVar()
        self.ai_model_var = tk.StringVar(value="deepseek-chat")

        self._build()

    def _build(self) -> None:
        outer = ttk.Frame(self.root, padding=12)
        outer.pack(fill=tk.BOTH, expand=True)
        outer.columnconfigure(1, weight=1)
        outer.rowconfigure(6, weight=1)

        self._path_row(outer, 0, "输入文件", self.input_var, self._choose_input)
        self._path_row(outer, 1, "输出文件", self.output_var, self._choose_output)
        self._path_row(outer, 2, "JSON报告", self.json_report_var, self._choose_json_report)
        self._path_row(outer, 3, "Markdown报告", self.markdown_report_var, self._choose_markdown_report)

        ai_frame = ttk.LabelFrame(outer, text="DeepSeek AI 审查（可选）", padding=8)
        ai_frame.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(8, 8))
        ai_frame.columnconfigure(1, weight=1)
        ttk.Checkbutton(ai_frame, text="启用 AI 审查", variable=self.ai_enabled_var).grid(row=0, column=0, sticky="w")
        ttk.Label(ai_frame, text="Base URL").grid(row=1, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(ai_frame, textvariable=self.ai_base_url_var).grid(row=1, column=1, sticky="ew", padx=8, pady=(6, 0))
        ttk.Label(ai_frame, text="模型").grid(row=1, column=2, sticky="w", pady=(6, 0))
        ttk.Entry(ai_frame, width=20, textvariable=self.ai_model_var).grid(row=1, column=3, sticky="w", pady=(6, 0))

        actions = ttk.Frame(outer)
        actions.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(0, 8))
        ttk.Button(actions, text="审计", command=self.audit).pack(side=tk.LEFT)
        ttk.Button(actions, text="格式化并审计", command=self.format).pack(side=tk.LEFT, padx=8)
        ttk.Button(actions, text="清空日志", command=lambda: self.log.delete("1.0", tk.END)).pack(side=tk.LEFT)

        self.log = tk.Text(outer, wrap="word", height=18)
        self.log.grid(row=6, column=0, columnspan=3, sticky="nsew")
        scroll = ttk.Scrollbar(outer, orient=tk.VERTICAL, command=self.log.yview)
        scroll.grid(row=6, column=3, sticky="ns")
        self.log.configure(yscrollcommand=scroll.set)

    def _path_row(self, parent, row: int, label: str, var: tk.StringVar, command) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=4)
        ttk.Entry(parent, textvariable=var).grid(row=row, column=1, sticky="ew", padx=8, pady=4)
        ttk.Button(parent, text="选择", command=command).grid(row=row, column=2, sticky="e", pady=4)

    def _choose_input(self) -> None:
        path = filedialog.askopenfilename(
            title="选择输入文件",
            filetypes=[("OfficeTool 支持文件", "*.docx *.txt *.md"), ("所有文件", "*.*")],
        )
        if path:
            self.input_var.set(path)
            source = Path(path)
            self.output_var.set(str(source.with_name(f"{source.stem}_formatted.docx")))
            self.json_report_var.set(str(source.with_name(f"{source.stem}_audit.json")))

    def _choose_output(self) -> None:
        path = filedialog.asksaveasfilename(title="保存格式化文件", defaultextension=".docx", filetypes=[("Word 文档", "*.docx")])
        if path:
            self.output_var.set(path)

    def _choose_json_report(self) -> None:
        path = filedialog.asksaveasfilename(title="保存 JSON 报告", defaultextension=".json", filetypes=[("JSON", "*.json")])
        if path:
            self.json_report_var.set(path)

    def _choose_markdown_report(self) -> None:
        path = filedialog.asksaveasfilename(title="保存 Markdown 报告", defaultextension=".md", filetypes=[("Markdown", "*.md")])
        if path:
            self.markdown_report_var.set(path)

    def _config(self) -> OfficeToolConfig:
        config = OfficeToolConfig()
        config.ai_review.enabled = self.ai_enabled_var.get()
        config.ai_review.base_url = self.ai_base_url_var.get().strip()
        config.ai_review.model = self.ai_model_var.get().strip() or "deepseek-chat"
        return config

    def audit(self) -> None:
        try:
            source = self._require_input()
            config = self._config()
            doc, _kind = load_document(source)
            report = OfficialDocumentAuditor(config).audit_document(doc)
            if config.ai_review.enabled:
                from .cli import maybe_run_ai_review

                maybe_run_ai_review(config, doc, report)
            self._write_reports(report)
            self._show_report(report)
        except Exception as exc:
            messagebox.showerror("审计失败", str(exc))
            self._append(f"审计失败：{exc}\n")

    def format(self) -> None:
        try:
            source = self._require_input()
            output = self.output_var.get().strip()
            if not output:
                raise ValueError("请选择输出文件。")
            config = self._config()
            report = OfficialDocumentFormatter(config).format_path(source, output)
            self._write_reports(report)
            self._show_report(report)
            self._append(f"已生成：{Path(output).resolve()}\n")
        except Exception as exc:
            messagebox.showerror("格式化失败", str(exc))
            self._append(f"格式化失败：{exc}\n")

    def _require_input(self) -> str:
        source = self.input_var.get().strip()
        if not source:
            raise ValueError("请选择输入文件。")
        if not Path(source).exists():
            raise FileNotFoundError(source)
        return source

    def _write_reports(self, report) -> None:
        if self.json_report_var.get().strip():
            write_json_report(report, self.json_report_var.get().strip())
        if self.markdown_report_var.get().strip():
            write_markdown_report(report, self.markdown_report_var.get().strip())

    def _show_report(self, report) -> None:
        self._append(report.summary() + "\n")
        for finding in report.findings:
            location = "" if finding.block_index is None else f"第 {finding.block_index + 1} 段 "
            self._append(f"[{finding.severity}] {location}{finding.message}\n")
            if finding.suggestion:
                self._append(f"  建议：{finding.suggestion}\n")
        self._append("\n")

    def _append(self, text: str) -> None:
        self.log.insert(tk.END, text)
        self.log.see(tk.END)


def main() -> int:
    root = tk.Tk()
    OfficeToolGUI(root)
    root.mainloop()
    return 0
