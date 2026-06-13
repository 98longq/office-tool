"""Tkinter desktop GUI for OfficeTool."""

from __future__ import annotations

import json
import os
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
        self.root.geometry("1180x780")
        self.root.minsize(980, 680)
        self.config = OfficeToolConfig()

        self.document_paths: list[Path] = []
        self.doc_output_dir = tk.StringVar()
        self.doc_report_dir = tk.StringVar()
        self.ai_enabled = tk.BooleanVar(value=False)
        self.ai_base_url = tk.StringVar()
        self.ai_model = tk.StringVar(value="deepseek-chat")
        self.page_vars: dict[str, tk.StringVar] = {}
        self.audit_vars: dict[str, tk.BooleanVar] = {}
        self.format_vars: dict[str, tk.BooleanVar] = {}
        self.style_vars: dict[str, tuple[tk.StringVar, tk.StringVar]] = {}

        self.excel_input = tk.StringVar()
        self.excel_output = tk.StringVar()

        self._init_config_vars()
        self._build()

    def _init_config_vars(self) -> None:
        page_fields = [
            "margin_top_cm",
            "margin_bottom_cm",
            "margin_left_cm",
            "margin_right_cm",
            "footer_distance_cm",
            "line_spacing_pt",
            "title_line_spacing_pt",
            "red_head_line_spacing_pt",
            "chars_per_line",
            "lines_per_page",
        ]
        self.page_vars = {name: tk.StringVar() for name in page_fields}

        self.format_vars = {
            "apply_page_setup": tk.BooleanVar(),
            "apply_styles": tk.BooleanVar(),
            "add_page_number": tk.BooleanVar(),
            "draw_red_separator": tk.BooleanVar(),
            "preserve_existing_bold_italic": tk.BooleanVar(),
        }
        self.audit_vars = {
            "require_document_number_for_red_head": tk.BooleanVar(),
            "require_signer_for_red_head": tk.BooleanVar(),
            "require_main_send": tk.BooleanVar(),
            "require_date": tk.BooleanVar(),
            "check_page_layout": tk.BooleanVar(),
        }
        self.style_vars = {
            name: (tk.StringVar(), tk.StringVar())
            for name in ["red_head", "title", "body", "h1", "h2", "h3", "copy_to", "page_number"]
        }
        self._sync_form_from_config()

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
        frame.rowconfigure(4, weight=1)

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

        settings = ttk.Notebook(frame)
        settings.grid(row=2, column=0, sticky="ew", pady=8)
        settings.add(self._output_settings_tab(settings), text="输出与 AI")
        settings.add(self._page_settings_tab(settings), text="页面版式")
        settings.add(self._style_settings_tab(settings), text="字体与格式")
        settings.add(self._audit_settings_tab(settings), text="审计规则")

        actions = ttk.Frame(frame)
        actions.grid(row=3, column=0, sticky="ew", pady=(0, 8))
        ttk.Button(actions, text="批量审计（只生成报告）", command=self.audit_documents).pack(side=tk.LEFT)
        ttk.Button(actions, text="批量格式化并审计（生成修改后文件）", command=self.format_documents).pack(side=tk.LEFT, padx=8)
        ttk.Button(actions, text="打开输出目录", command=lambda: self.open_folder(self.doc_output_dir.get())).pack(side=tk.LEFT)
        ttk.Button(actions, text="打开报告目录", command=lambda: self.open_folder(self.doc_report_dir.get())).pack(side=tk.LEFT, padx=8)
        ttk.Button(actions, text="清空日志", command=lambda: self._clear_log(self.doc_log)).pack(side=tk.LEFT)

        self.doc_log = tk.Text(frame, wrap="word", height=12)
        self.doc_log.grid(row=4, column=0, sticky="nsew")
        return frame

    def _output_settings_tab(self, parent) -> ttk.Frame:
        frame = ttk.Frame(parent, padding=8)
        frame.columnconfigure(1, weight=1)
        frame.columnconfigure(3, weight=1)
        self._path_entry(frame, 0, "格式化输出目录", self.doc_output_dir, self.choose_doc_output_dir)
        self._path_entry(frame, 1, "审计报告目录", self.doc_report_dir, self.choose_doc_report_dir)
        ttk.Checkbutton(frame, text="启用 DeepSeek AI 审查", variable=self.ai_enabled).grid(row=2, column=0, sticky="w", pady=(8, 0))
        ttk.Label(frame, text="Base URL").grid(row=3, column=0, sticky="w", pady=4)
        ttk.Entry(frame, textvariable=self.ai_base_url).grid(row=3, column=1, sticky="ew", padx=8, pady=4)
        ttk.Label(frame, text="模型").grid(row=3, column=2, sticky="w", pady=4)
        ttk.Entry(frame, textvariable=self.ai_model).grid(row=3, column=3, sticky="ew", padx=8, pady=4)
        ttk.Label(frame, text="说明：审计只生成报告；格式化会在输出目录生成修改后的 .docx。").grid(row=4, column=0, columnspan=4, sticky="w", pady=(8, 0))
        return frame

    def _page_settings_tab(self, parent) -> ttk.Frame:
        frame = ttk.Frame(parent, padding=8)
        labels = [
            ("上边距 cm", "margin_top_cm"),
            ("下边距 cm", "margin_bottom_cm"),
            ("左边距 cm", "margin_left_cm"),
            ("右边距 cm", "margin_right_cm"),
            ("页脚距 cm", "footer_distance_cm"),
            ("正文行距 pt", "line_spacing_pt"),
            ("标题行距 pt", "title_line_spacing_pt"),
            ("红头行距 pt", "red_head_line_spacing_pt"),
            ("每行字数", "chars_per_line"),
            ("每页行数", "lines_per_page"),
        ]
        for index, (label, key) in enumerate(labels):
            row = index // 2
            col = (index % 2) * 2
            ttk.Label(frame, text=label).grid(row=row, column=col, sticky="w", padx=(0, 6), pady=4)
            ttk.Entry(frame, width=12, textvariable=self.page_vars[key]).grid(row=row, column=col + 1, sticky="w", padx=(0, 24), pady=4)
        return frame

    def _style_settings_tab(self, parent) -> ttk.Frame:
        frame = ttk.Frame(parent, padding=8)
        for row, (key, label) in enumerate([
            ("red_head", "红头版头"),
            ("title", "标题"),
            ("body", "正文"),
            ("h1", "一级标题"),
            ("h2", "二级标题"),
            ("h3", "三级标题"),
            ("copy_to", "版记/抄送"),
            ("page_number", "页码"),
        ]):
            font_var, size_var = self.style_vars[key]
            ttk.Label(frame, text=label).grid(row=row, column=0, sticky="w", pady=4)
            ttk.Label(frame, text="字体").grid(row=row, column=1, sticky="e", padx=(12, 4), pady=4)
            ttk.Entry(frame, width=18, textvariable=font_var).grid(row=row, column=2, sticky="w", pady=4)
            ttk.Label(frame, text="字号 pt").grid(row=row, column=3, sticky="e", padx=(12, 4), pady=4)
            ttk.Entry(frame, width=8, textvariable=size_var).grid(row=row, column=4, sticky="w", pady=4)

        checks = ttk.Frame(frame)
        checks.grid(row=8, column=0, columnspan=5, sticky="w", pady=(8, 0))
        ttk.Checkbutton(checks, text="应用页面设置", variable=self.format_vars["apply_page_setup"]).pack(side=tk.LEFT)
        ttk.Checkbutton(checks, text="应用字体段落样式", variable=self.format_vars["apply_styles"]).pack(side=tk.LEFT, padx=10)
        ttk.Checkbutton(checks, text="添加页码", variable=self.format_vars["add_page_number"]).pack(side=tk.LEFT)
        ttk.Checkbutton(checks, text="绘制红头分隔线", variable=self.format_vars["draw_red_separator"]).pack(side=tk.LEFT, padx=10)
        ttk.Checkbutton(checks, text="保留原有加粗/斜体", variable=self.format_vars["preserve_existing_bold_italic"]).pack(side=tk.LEFT)
        return frame

    def _audit_settings_tab(self, parent) -> ttk.Frame:
        frame = ttk.Frame(parent, padding=8)
        ttk.Checkbutton(frame, text="红头文件必须有发文字号", variable=self.audit_vars["require_document_number_for_red_head"]).grid(row=0, column=0, sticky="w", pady=4)
        ttk.Checkbutton(frame, text="红头文件必须有签发人", variable=self.audit_vars["require_signer_for_red_head"]).grid(row=0, column=1, sticky="w", padx=16, pady=4)
        ttk.Checkbutton(frame, text="要求主送机关", variable=self.audit_vars["require_main_send"]).grid(row=1, column=0, sticky="w", pady=4)
        ttk.Checkbutton(frame, text="要求成文日期", variable=self.audit_vars["require_date"]).grid(row=1, column=1, sticky="w", padx=16, pady=4)
        ttk.Checkbutton(frame, text="检查页面版式", variable=self.audit_vars["check_page_layout"]).grid(row=2, column=0, sticky="w", pady=4)
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

        note = ttk.Label(frame, text="清洗会去除文本前后空白，并删除完全空白的行。原文件不会被覆盖，除非手动选择同名输出。")
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
        self._sync_form_from_config()
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
            self._append(self.doc_log, "正在审计：此操作只生成审计报告，不生成修改后的文件。\n")
            results = audit_many(paths, self.config, self.doc_report_dir.get() or None, markdown=True, log=lambda msg: self._append(self.doc_log, msg + "\n"))
            self._show_document_results(results)
        except Exception as exc:
            messagebox.showerror("审计失败", str(exc))

    def format_documents(self) -> None:
        try:
            self._sync_config_from_form()
            paths = self._require_documents()
            if not self.doc_output_dir.get():
                raise ValueError("请选择格式化输出目录。")
            self._append(self.doc_log, f"正在格式化：修改后的文件将写入 {Path(self.doc_output_dir.get()).resolve()}\n")
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

    def _sync_form_from_config(self) -> None:
        self.ai_enabled.set(self.config.ai_review.enabled)
        self.ai_base_url.set(self.config.ai_review.base_url)
        self.ai_model.set(self.config.ai_review.model)
        for key, var in self.page_vars.items():
            var.set(str(getattr(self.config.page, key)))
        for key, var in self.format_vars.items():
            var.set(bool(getattr(self.config.format, key)))
        for key, var in self.audit_vars.items():
            var.set(bool(getattr(self.config.audit, key)))
        for key, (font_var, size_var) in self.style_vars.items():
            style = self.config.styles[key]
            font_var.set(style.font)
            size_var.set(str(style.size_pt))

    def _sync_config_from_form(self) -> None:
        self.config.ai_review.enabled = self.ai_enabled.get()
        self.config.ai_review.base_url = self.ai_base_url.get().strip()
        self.config.ai_review.model = self.ai_model.get().strip() or "deepseek-chat"
        for key, var in self.page_vars.items():
            current = getattr(self.config.page, key)
            raw = var.get().strip()
            setattr(self.config.page, key, int(raw) if isinstance(current, int) else float(raw))
        for key, var in self.format_vars.items():
            setattr(self.config.format, key, var.get())
        for key, var in self.audit_vars.items():
            setattr(self.config.audit, key, var.get())
        for key, (font_var, size_var) in self.style_vars.items():
            self.config.styles[key].font = font_var.get().strip()
            self.config.styles[key].size_pt = float(size_var.get().strip())

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
            if result.markdown_report:
                self._append(self.doc_log, f"Markdown 报告：{result.markdown_report}\n")
            if result.report:
                self._append(self.doc_log, result.report.summary() + "\n")
        self._append(self.doc_log, "\n")

    def open_folder(self, folder: str) -> None:
        if not folder:
            messagebox.showinfo("未选择目录", "请先选择目录。")
            return
        path = Path(folder).expanduser().resolve()
        path.mkdir(parents=True, exist_ok=True)
        os.startfile(path)

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
