"""Tkinter desktop GUI for OfficeTool."""

from __future__ import annotations

import json
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from .config import OfficeToolConfig
from .excel import clean_workbook, inspect_workbook
from .services import audit_many, format_many, summarize_results


class OfficeToolGUI:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("OfficeTool 办公助手")
        self.root.geometry("1080x720")
        self.root.minsize(920, 620)
        self.config = OfficeToolConfig()

        self.document_paths: list[Path] = []
        self.doc_output_dir = tk.StringVar()
        self.doc_report_dir = tk.StringVar()
        self.ai_enabled = tk.BooleanVar(value=False)
        self.ai_base_url = tk.StringVar()
        self.ai_model = tk.StringVar(value="deepseek-chat")

        self.excel_input = tk.StringVar()
        self.excel_output = tk.StringVar()

        self._build()

    def _build(self) -> None:
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill=tk.BOTH, expand=True)
        self.doc_log = tk.Text()
        self.excel_log = tk.Text()
        notebook.add(self._document_tab(notebook), text="公文审计")
        notebook.add(self._excel_tab(notebook), text="Excel 工具")

    def _document_tab(self, parent) -> ttk.Frame:
        frame = ttk.Frame(parent, padding=12)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(1, weight=1)

        toolbar = ttk.Frame(frame)
        toolbar.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        ttk.Button(toolbar, text="添加文件", command=self.add_document_files).pack(side=tk.LEFT)
        ttk.Button(toolbar, text="添加文件夹", command=self.add_document_folder).pack(side=tk.LEFT, padx=6)
        ttk.Button(toolbar, text="移除选中", command=self.remove_selected_documents).pack(side=tk.LEFT)
        ttk.Button(toolbar, text="清空", command=self.clear_documents).pack(side=tk.LEFT, padx=6)
        ttk.Button(toolbar, text="加载配置", command=self.load_config).pack(side=tk.RIGHT)
        ttk.Button(toolbar, text="保存配置", command=self.save_config).pack(side=tk.RIGHT, padx=6)

        self.doc_list = tk.Listbox(frame, height=8)
        self.doc_list.grid(row=1, column=0, sticky="nsew")

        options = ttk.LabelFrame(frame, text="输出与 AI 审查", padding=8)
        options.grid(row=2, column=0, sticky="ew", pady=8)
        options.columnconfigure(1, weight=1)
        options.columnconfigure(3, weight=1)
        self._path_entry(options, 0, "输出目录", self.doc_output_dir, self.choose_doc_output_dir)
        self._path_entry(options, 1, "报告目录", self.doc_report_dir, self.choose_doc_report_dir)
        ttk.Checkbutton(options, text="启用 DeepSeek AI 审查", variable=self.ai_enabled).grid(row=2, column=0, sticky="w", pady=(8, 0))
        ttk.Label(options, text="Base URL").grid(row=3, column=0, sticky="w", pady=4)
        ttk.Entry(options, textvariable=self.ai_base_url).grid(row=3, column=1, sticky="ew", padx=8, pady=4)
        ttk.Label(options, text="模型").grid(row=3, column=2, sticky="w", pady=4)
        ttk.Entry(options, textvariable=self.ai_model).grid(row=3, column=3, sticky="ew", padx=8, pady=4)

        actions = ttk.Frame(frame)
        actions.grid(row=3, column=0, sticky="ew", pady=(0, 8))
        ttk.Button(actions, text="批量审计", command=self.audit_documents).pack(side=tk.LEFT)
        ttk.Button(actions, text="批量格式化并审计", command=self.format_documents).pack(side=tk.LEFT, padx=8)
        ttk.Button(actions, text="清空日志", command=lambda: self._clear_log(self.doc_log)).pack(side=tk.LEFT)

        self.doc_log = tk.Text(frame, wrap="word", height=12)
        self.doc_log.grid(row=4, column=0, sticky="nsew")
        frame.rowconfigure(4, weight=1)
        return frame

    def _excel_tab(self, parent) -> ttk.Frame:
        frame = ttk.Frame(parent, padding=12)
        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(4, weight=1)

        self._path_entry(frame, 0, "输入工作簿", self.excel_input, self.choose_excel_input)
        self._path_entry(frame, 1, "输出工作簿", self.excel_output, self.choose_excel_output)

        actions = ttk.Frame(frame)
        actions.grid(row=2, column=0, columnspan=3, sticky="ew", pady=8)
        ttk.Button(actions, text="检查工作簿", command=self.inspect_excel).pack(side=tk.LEFT)
        ttk.Button(actions, text="清洗并另存", command=self.clean_excel).pack(side=tk.LEFT, padx=8)
        ttk.Button(actions, text="清空日志", command=lambda: self._clear_log(self.excel_log)).pack(side=tk.LEFT)

        note = ttk.Label(frame, text="清洗会去除文本前后空白，并删除完全空白的行。原文件不会被覆盖，除非你手动选择同名输出。")
        note.grid(row=3, column=0, columnspan=3, sticky="w")

        self.excel_log = tk.Text(frame, wrap="word", height=18)
        self.excel_log.grid(row=4, column=0, columnspan=3, sticky="nsew", pady=(8, 0))
        return frame

    def _path_entry(self, parent, row: int, label: str, var: tk.StringVar, command) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=4)
        ttk.Entry(parent, textvariable=var).grid(row=row, column=1, sticky="ew", padx=8, pady=4)
        ttk.Button(parent, text="选择", command=command).grid(row=row, column=2, sticky="e", pady=4)

    def add_document_files(self) -> None:
        paths = filedialog.askopenfilenames(filetypes=[("支持文档", "*.docx *.txt *.md"), ("所有文件", "*.*")])
        self._add_document_paths(paths)

    def add_document_folder(self) -> None:
        folder = filedialog.askdirectory()
        if folder:
            self._add_document_paths([folder])

    def _add_document_paths(self, paths) -> None:
        for raw in paths:
            path = Path(raw).resolve()
            if path not in self.document_paths:
                self.document_paths.append(path)
                self.doc_list.insert(tk.END, str(path))
        if paths and not self.doc_output_dir.get():
            first = Path(paths[0]).resolve()
            base = first if first.is_dir() else first.parent
            self.doc_output_dir.set(str(base / "formatted_output"))
            self.doc_report_dir.set(str(base / "audit_reports"))

    def remove_selected_documents(self) -> None:
        selected = list(self.doc_list.curselection())
        for index in reversed(selected):
            self.doc_list.delete(index)
            del self.document_paths[index]

    def clear_documents(self) -> None:
        self.doc_list.delete(0, tk.END)
        self.document_paths.clear()

    def choose_doc_output_dir(self) -> None:
        folder = filedialog.askdirectory()
        if folder:
            self.doc_output_dir.set(folder)

    def choose_doc_report_dir(self) -> None:
        folder = filedialog.askdirectory()
        if folder:
            self.doc_report_dir.set(folder)

    def choose_excel_input(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Excel 工作簿", "*.xlsx"), ("所有文件", "*.*")])
        if path:
            self.excel_input.set(path)
            source = Path(path)
            self.excel_output.set(str(source.with_name(f"{source.stem}_cleaned.xlsx")))

    def choose_excel_output(self) -> None:
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel 工作簿", "*.xlsx")])
        if path:
            self.excel_output.set(path)

    def load_config(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("JSON 配置", "*.json"), ("所有文件", "*.*")])
        if not path:
            return
        self.config = OfficeToolConfig.from_dict(json.loads(Path(path).read_text(encoding="utf-8")))
        self.ai_enabled.set(self.config.ai_review.enabled)
        self.ai_base_url.set(self.config.ai_review.base_url)
        self.ai_model.set(self.config.ai_review.model)
        self._append(self.doc_log, f"已加载配置：{path}\n")

    def save_config(self) -> None:
        path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON 配置", "*.json")])
        if not path:
            return
        self._sync_config_from_form()
        Path(path).write_text(json.dumps(self.config.to_dict(), ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        self._append(self.doc_log, f"已保存配置：{path}\n")

    def audit_documents(self) -> None:
        try:
            self._sync_config_from_form()
            paths = self._require_documents()
            results = audit_many(paths, self.config, self.doc_report_dir.get() or None, markdown=True, log=lambda msg: self._append(self.doc_log, msg + "\n"))
            self._show_document_results(results)
        except Exception as exc:
            messagebox.showerror("审计失败", str(exc))

    def format_documents(self) -> None:
        try:
            self._sync_config_from_form()
            paths = self._require_documents()
            if not self.doc_output_dir.get():
                raise ValueError("请选择输出目录。")
            results = format_many(
                paths,
                self.doc_output_dir.get(),
                self.config,
                report_dir=self.doc_report_dir.get() or None,
                markdown=True,
                log=lambda msg: self._append(self.doc_log, msg + "\n"),
            )
            self._show_document_results(results)
        except Exception as exc:
            messagebox.showerror("格式化失败", str(exc))

    def inspect_excel(self) -> None:
        try:
            summary = inspect_workbook(self._require_excel_input())
            self._append(self.excel_log, json.dumps(summary.to_dict(), ensure_ascii=False, indent=2) + "\n")
        except Exception as exc:
            messagebox.showerror("检查失败", str(exc))

    def clean_excel(self) -> None:
        try:
            if not self.excel_output.get():
                raise ValueError("请选择输出工作簿。")
            summary = clean_workbook(self._require_excel_input(), self.excel_output.get())
            self._append(self.excel_log, f"已生成：{Path(self.excel_output.get()).resolve()}\n")
            self._append(self.excel_log, json.dumps(summary.to_dict(), ensure_ascii=False, indent=2) + "\n")
        except Exception as exc:
            messagebox.showerror("清洗失败", str(exc))

    def _sync_config_from_form(self) -> None:
        self.config.ai_review.enabled = self.ai_enabled.get()
        self.config.ai_review.base_url = self.ai_base_url.get().strip()
        self.config.ai_review.model = self.ai_model.get().strip() or "deepseek-chat"

    def _require_documents(self) -> list[Path]:
        if not self.document_paths:
            raise ValueError("请先添加文件或文件夹。")
        return list(self.document_paths)

    def _require_excel_input(self) -> str:
        path = self.excel_input.get().strip()
        if not path:
            raise ValueError("请选择输入工作簿。")
        if not Path(path).exists():
            raise FileNotFoundError(path)
        return path

    def _show_document_results(self, results) -> None:
        self._append(self.doc_log, summarize_results(results) + "\n")
        for result in results:
            if result.error:
                self._append(self.doc_log, f"[失败] {result.source}: {result.error}\n")
                continue
            if result.output:
                self._append(self.doc_log, f"输出：{result.output}\n")
            if result.json_report:
                self._append(self.doc_log, f"报告：{result.json_report}\n")
            if result.report:
                self._append(self.doc_log, result.report.summary() + "\n")
        self._append(self.doc_log, "\n")

    @staticmethod
    def _append(widget: tk.Text, text: str) -> None:
        widget.insert(tk.END, text)
        widget.see(tk.END)

    @staticmethod
    def _clear_log(widget: tk.Text) -> None:
        widget.delete("1.0", tk.END)


def main() -> int:
    root = tk.Tk()
    OfficeToolGUI(root)
    root.mainloop()
    return 0
