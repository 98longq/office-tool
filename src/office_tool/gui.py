"""Tkinter desktop GUI for OfficeTool."""

from __future__ import annotations

import json
import os
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from .ai import DeepSeekTextReviewer
from .config import OfficeToolConfig
from .excel import clean_workbook, inspect_workbook
from .services import audit_many, format_many, summarize_results


class OfficeToolGUI:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("OfficeTool")
        self.root.geometry("1240x820")
        self.root.minsize(1040, 700)
        self.config = OfficeToolConfig()

        self.document_paths: list[Path] = []
        self.doc_output_dir = tk.StringVar()
        self.doc_report_dir = tk.StringVar()
        self.ai_enabled = tk.BooleanVar(value=False)
        self.ai_base_url = tk.StringVar()
        self.ai_model = tk.StringVar(value="DeepSeek-R1")
        self.ai_api_key = tk.StringVar()
        self.ai_key_env = tk.StringVar(value="DEEPSEEK_API_KEY")
        self.ai_auth_prefix = tk.StringVar(value="Bearer")
        self.ai_stream = tk.BooleanVar(value=False)
        self.page_vars: dict[str, tk.StringVar] = {}
        self.audit_vars: dict[str, tk.BooleanVar] = {}
        self.format_vars: dict[str, tk.BooleanVar] = {}
        self.style_vars: dict[str, tuple[tk.StringVar, tk.StringVar]] = {}

        self.excel_input = tk.StringVar()
        self.excel_output = tk.StringVar()

        self._init_config_vars()
        self._configure_theme()
        self._build()

    def _configure_theme(self) -> None:
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        self.root.configure(bg="#f4f6f8")
        style.configure(".", font=("Microsoft YaHei UI", 10), background="#f4f6f8", foreground="#17202a")
        style.configure("Panel.TFrame", background="#ffffff", relief="flat")
        style.configure("Side.TFrame", background="#152238")
        style.configure("SideTitle.TLabel", background="#152238", foreground="#ffffff", font=("Microsoft YaHei UI", 18, "bold"))
        style.configure("SideText.TLabel", background="#152238", foreground="#c9d4e5")
        style.configure("Section.TLabel", background="#ffffff", foreground="#17202a", font=("Microsoft YaHei UI", 13, "bold"))
        style.configure("Hint.TLabel", background="#ffffff", foreground="#5b677a")
        style.configure("Primary.TButton", padding=(14, 8), background="#2563eb", foreground="#ffffff")
        style.map("Primary.TButton", background=[("active", "#1d4ed8")])
        style.configure("TButton", padding=(10, 6))
        style.configure("TNotebook", background="#f4f6f8", borderwidth=0)
        style.configure("TNotebook.Tab", padding=(16, 8))

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
        shell = ttk.Frame(self.root)
        shell.pack(fill=tk.BOTH, expand=True)
        shell.columnconfigure(1, weight=1)
        shell.rowconfigure(0, weight=1)

        side = ttk.Frame(shell, style="Side.TFrame", padding=18)
        side.grid(row=0, column=0, sticky="ns")
        ttk.Label(side, text="OfficeTool", style="SideTitle.TLabel").pack(anchor="w")
        ttk.Label(side, text="Document audit and office utilities", style="SideText.TLabel", wraplength=210).pack(anchor="w", pady=(8, 24))
        ttk.Button(side, text="Add files", command=self.add_document_files, style="Primary.TButton").pack(fill=tk.X, pady=4)
        ttk.Button(side, text="Add folder", command=self.add_document_folder).pack(fill=tk.X, pady=4)
        ttk.Button(side, text="Remove selected", command=self.remove_selected_documents).pack(fill=tk.X, pady=4)
        ttk.Button(side, text="Clear list", command=self.clear_documents).pack(fill=tk.X, pady=4)
        ttk.Separator(side).pack(fill=tk.X, pady=16)
        ttk.Button(side, text="Run audit only", command=self.audit_documents).pack(fill=tk.X, pady=4)
        ttk.Button(side, text="Format and audit", command=self.format_documents, style="Primary.TButton").pack(fill=tk.X, pady=4)
        ttk.Button(side, text="Open output", command=lambda: self.open_folder(self.doc_output_dir.get())).pack(fill=tk.X, pady=4)
        ttk.Button(side, text="Open reports", command=lambda: self.open_folder(self.doc_report_dir.get())).pack(fill=tk.X, pady=4)
        ttk.Separator(side).pack(fill=tk.X, pady=16)
        ttk.Button(side, text="Load config", command=self.load_config).pack(fill=tk.X, pady=4)
        ttk.Button(side, text="Save config", command=self.save_config).pack(fill=tk.X, pady=4)

        main = ttk.Frame(shell, padding=18)
        main.grid(row=0, column=1, sticky="nsew")
        main.columnconfigure(0, weight=1)
        main.rowconfigure(1, weight=1)

        header = ttk.Frame(main, style="Panel.TFrame", padding=16)
        header.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        ttk.Label(header, text="Official Document Workspace", style="Section.TLabel").pack(anchor="w")
        ttk.Label(
            header,
            text="Audit creates reports only. Format and audit writes corrected .docx files to the output folder.",
            style="Hint.TLabel",
        ).pack(anchor="w", pady=(4, 0))

        body = ttk.PanedWindow(main, orient=tk.HORIZONTAL)
        body.grid(row=1, column=0, sticky="nsew")

        left = ttk.Frame(body, style="Panel.TFrame", padding=12)
        left.columnconfigure(0, weight=1)
        left.rowconfigure(1, weight=1)
        ttk.Label(left, text="Document Queue", style="Section.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 8))
        self.doc_list = tk.Listbox(left, height=18, borderwidth=0, highlightthickness=1, highlightbackground="#d6dde8")
        self.doc_list.grid(row=1, column=0, sticky="nsew")
        body.add(left, weight=1)

        right = ttk.Frame(body, style="Panel.TFrame", padding=12)
        right.columnconfigure(0, weight=1)
        right.rowconfigure(0, weight=1)
        notebook = ttk.Notebook(right)
        notebook.grid(row=0, column=0, sticky="nsew")
        notebook.add(self._io_ai_tab(notebook), text="Output + AI")
        notebook.add(self._layout_tab(notebook), text="Layout")
        notebook.add(self._style_tab(notebook), text="Styles")
        notebook.add(self._audit_tab(notebook), text="Rules")
        notebook.add(self._excel_tab(notebook), text="Excel")
        notebook.add(self._log_tab(notebook), text="Run Log")
        body.add(right, weight=2)

    def _io_ai_tab(self, parent) -> ttk.Frame:
        frame = ttk.Frame(parent, padding=14)
        frame.columnconfigure(1, weight=1)
        self._path_entry(frame, 0, "Formatted output folder", self.doc_output_dir, self.choose_doc_output_dir)
        self._path_entry(frame, 1, "Audit report folder", self.doc_report_dir, self.choose_doc_report_dir)
        ttk.Separator(frame).grid(row=2, column=0, columnspan=3, sticky="ew", pady=12)
        ttk.Checkbutton(frame, text="Enable DeepSeek text review", variable=self.ai_enabled).grid(row=3, column=0, sticky="w", pady=4)
        self._labeled_entry(frame, 4, "DeepSeek URL", self.ai_base_url, "http://ai.crc.cr/deepseek/v1/chat/completions")
        self._labeled_entry(frame, 5, "Model", self.ai_model, "DeepSeek-R1")
        self._labeled_entry(frame, 6, "API key", self.ai_api_key, "Optional; stored only if you save config", show="*")
        self._labeled_entry(frame, 7, "API key env", self.ai_key_env, "DEEPSEEK_API_KEY")
        self._labeled_entry(frame, 8, "Auth prefix", self.ai_auth_prefix, "Bearer; leave empty if Authorization must be raw token")
        ttk.Checkbutton(frame, text="Use streaming response parser", variable=self.ai_stream).grid(row=9, column=1, sticky="w", pady=4)
        ttk.Button(frame, text="Test AI connection", command=self.test_ai_connection).grid(row=10, column=1, sticky="w", pady=10)
        return frame

    def _layout_tab(self, parent) -> ttk.Frame:
        frame = ttk.Frame(parent, padding=14)
        labels = [
            ("Top margin cm", "margin_top_cm"),
            ("Bottom margin cm", "margin_bottom_cm"),
            ("Left margin cm", "margin_left_cm"),
            ("Right margin cm", "margin_right_cm"),
            ("Footer distance cm", "footer_distance_cm"),
            ("Body line spacing pt", "line_spacing_pt"),
            ("Title line spacing pt", "title_line_spacing_pt"),
            ("Red head spacing pt", "red_head_line_spacing_pt"),
            ("Chars per line", "chars_per_line"),
            ("Lines per page", "lines_per_page"),
        ]
        for index, (label, key) in enumerate(labels):
            row = index // 2
            col = (index % 2) * 2
            ttk.Label(frame, text=label).grid(row=row, column=col, sticky="w", padx=(0, 8), pady=6)
            ttk.Entry(frame, width=12, textvariable=self.page_vars[key]).grid(row=row, column=col + 1, sticky="w", padx=(0, 26), pady=6)
        return frame

    def _style_tab(self, parent) -> ttk.Frame:
        frame = ttk.Frame(parent, padding=14)
        rows = [
            ("red_head", "Red head"),
            ("title", "Title"),
            ("body", "Body"),
            ("h1", "Heading 1"),
            ("h2", "Heading 2"),
            ("h3", "Heading 3"),
            ("copy_to", "Copy/print note"),
            ("page_number", "Page number"),
        ]
        for row, (key, label) in enumerate(rows):
            font_var, size_var = self.style_vars[key]
            ttk.Label(frame, text=label).grid(row=row, column=0, sticky="w", pady=5)
            ttk.Label(frame, text="Font").grid(row=row, column=1, sticky="e", padx=(16, 4), pady=5)
            ttk.Entry(frame, width=20, textvariable=font_var).grid(row=row, column=2, sticky="w", pady=5)
            ttk.Label(frame, text="pt").grid(row=row, column=3, sticky="e", padx=(16, 4), pady=5)
            ttk.Entry(frame, width=8, textvariable=size_var).grid(row=row, column=4, sticky="w", pady=5)
        checks = ttk.Frame(frame)
        checks.grid(row=9, column=0, columnspan=5, sticky="w", pady=(12, 0))
        ttk.Checkbutton(checks, text="Page setup", variable=self.format_vars["apply_page_setup"]).pack(side=tk.LEFT)
        ttk.Checkbutton(checks, text="Paragraph styles", variable=self.format_vars["apply_styles"]).pack(side=tk.LEFT, padx=10)
        ttk.Checkbutton(checks, text="Page number", variable=self.format_vars["add_page_number"]).pack(side=tk.LEFT)
        ttk.Checkbutton(checks, text="Red separator", variable=self.format_vars["draw_red_separator"]).pack(side=tk.LEFT, padx=10)
        ttk.Checkbutton(checks, text="Preserve bold/italic", variable=self.format_vars["preserve_existing_bold_italic"]).pack(side=tk.LEFT)
        return frame

    def _audit_tab(self, parent) -> ttk.Frame:
        frame = ttk.Frame(parent, padding=14)
        ttk.Checkbutton(frame, text="Red-head document must include document number", variable=self.audit_vars["require_document_number_for_red_head"]).grid(row=0, column=0, sticky="w", pady=6)
        ttk.Checkbutton(frame, text="Red-head document must include signer", variable=self.audit_vars["require_signer_for_red_head"]).grid(row=1, column=0, sticky="w", pady=6)
        ttk.Checkbutton(frame, text="Require main recipient", variable=self.audit_vars["require_main_send"]).grid(row=2, column=0, sticky="w", pady=6)
        ttk.Checkbutton(frame, text="Require written date", variable=self.audit_vars["require_date"]).grid(row=3, column=0, sticky="w", pady=6)
        ttk.Checkbutton(frame, text="Check page layout", variable=self.audit_vars["check_page_layout"]).grid(row=4, column=0, sticky="w", pady=6)
        return frame

    def _excel_tab(self, parent) -> ttk.Frame:
        frame = ttk.Frame(parent, padding=14)
        frame.columnconfigure(1, weight=1)
        self._path_entry(frame, 0, "Input workbook", self.excel_input, self.choose_excel_input)
        self._path_entry(frame, 1, "Output workbook", self.excel_output, self.choose_excel_output)
        ttk.Button(frame, text="Inspect workbook", command=self.inspect_excel).grid(row=2, column=1, sticky="w", pady=10)
        ttk.Button(frame, text="Clean and save as", command=self.clean_excel).grid(row=2, column=1, sticky="w", padx=(140, 0), pady=10)
        return frame

    def _log_tab(self, parent) -> ttk.Frame:
        frame = ttk.Frame(parent, padding=8)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)
        self.doc_log = tk.Text(frame, wrap="word", borderwidth=0, highlightthickness=1, highlightbackground="#d6dde8")
        self.doc_log.grid(row=0, column=0, sticky="nsew")
        ttk.Button(frame, text="Clear log", command=lambda: self._clear_log(self.doc_log)).grid(row=1, column=0, sticky="w", pady=(8, 0))
        return frame

    def _path_entry(self, parent, row: int, label: str, var: tk.StringVar, command) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=6)
        ttk.Entry(parent, textvariable=var).grid(row=row, column=1, sticky="ew", padx=8, pady=6)
        ttk.Button(parent, text="Browse", command=command).grid(row=row, column=2, sticky="e", pady=6)

    def _labeled_entry(self, parent, row: int, label: str, var: tk.StringVar, hint: str = "", show: str | None = None) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=6)
        ttk.Entry(parent, textvariable=var, show=show or "").grid(row=row, column=1, sticky="ew", padx=8, pady=6)
        if hint:
            ttk.Label(parent, text=hint, style="Hint.TLabel").grid(row=row, column=2, sticky="w", pady=6)

    def add_document_files(self) -> None:
        paths = filedialog.askopenfilenames(filetypes=[("Supported documents", "*.docx *.txt *.md"), ("All files", "*.*")])
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
        path = filedialog.askopenfilename(filetypes=[("Excel workbook", "*.xlsx"), ("All files", "*.*")])
        if path:
            self.excel_input.set(path)
            source = Path(path)
            self.excel_output.set(str(source.with_name(f"{source.stem}_cleaned.xlsx")))

    def choose_excel_output(self) -> None:
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel workbook", "*.xlsx")])
        if path:
            self.excel_output.set(path)

    def load_config(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("JSON config", "*.json"), ("All files", "*.*")])
        if not path:
            return
        self.config = OfficeToolConfig.from_dict(json.loads(Path(path).read_text(encoding="utf-8")))
        self._sync_form_from_config()
        self._append(self.doc_log, f"Loaded config: {path}\n")

    def save_config(self) -> None:
        path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON config", "*.json")])
        if not path:
            return
        self._sync_config_from_form()
        Path(path).write_text(json.dumps(self.config.to_dict(), ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        self._append(self.doc_log, f"Saved config: {path}\n")

    def audit_documents(self) -> None:
        try:
            self._sync_config_from_form()
            paths = self._require_documents()
            self._append(self.doc_log, "Audit only: reports will be generated, no formatted file will be written.\n")
            results = audit_many(paths, self.config, self.doc_report_dir.get() or None, markdown=True, log=lambda msg: self._append(self.doc_log, msg + "\n"))
            self._show_document_results(results)
        except Exception as exc:
            messagebox.showerror("Audit failed", str(exc))

    def format_documents(self) -> None:
        try:
            self._sync_config_from_form()
            paths = self._require_documents()
            if not self.doc_output_dir.get():
                raise ValueError("Please choose a formatted output folder.")
            self._append(self.doc_log, f"Format and audit: formatted files will be written to {Path(self.doc_output_dir.get()).resolve()}\n")
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
            messagebox.showerror("Format failed", str(exc))

    def test_ai_connection(self) -> None:
        try:
            self._sync_config_from_form()
            if not self.config.ai_review.base_url:
                raise ValueError("Please fill DeepSeek URL first.")
            reviewer = DeepSeekTextReviewer(self.config.ai_review)
            findings = reviewer.review_text("Test document title\nMain recipient:\nThis is a short connectivity test.")
            self._append(self.doc_log, f"AI connection succeeded. Parsed findings: {len(findings)}\n")
            messagebox.showinfo("AI connection", "DeepSeek connection succeeded.")
        except Exception as exc:
            self._append(self.doc_log, f"AI connection failed: {exc}\n")
            messagebox.showerror("AI connection failed", str(exc))

    def inspect_excel(self) -> None:
        try:
            summary = inspect_workbook(self._require_excel_input())
            self._append(self.doc_log, json.dumps(summary.to_dict(), ensure_ascii=False, indent=2) + "\n")
        except Exception as exc:
            messagebox.showerror("Inspect failed", str(exc))

    def clean_excel(self) -> None:
        try:
            if not self.excel_output.get():
                raise ValueError("Please choose an output workbook.")
            summary = clean_workbook(self._require_excel_input(), self.excel_output.get())
            self._append(self.doc_log, f"Excel saved: {Path(self.excel_output.get()).resolve()}\n")
            self._append(self.doc_log, json.dumps(summary.to_dict(), ensure_ascii=False, indent=2) + "\n")
        except Exception as exc:
            messagebox.showerror("Clean failed", str(exc))

    def _sync_form_from_config(self) -> None:
        self.ai_enabled.set(self.config.ai_review.enabled)
        self.ai_base_url.set(self.config.ai_review.base_url)
        self.ai_model.set(self.config.ai_review.model)
        self.ai_api_key.set(self.config.ai_review.api_key)
        self.ai_key_env.set(self.config.ai_review.api_key_env)
        self.ai_auth_prefix.set(self.config.ai_review.auth_prefix)
        self.ai_stream.set(self.config.ai_review.stream)
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
        self.config.ai_review.model = self.ai_model.get().strip() or "DeepSeek-R1"
        self.config.ai_review.api_key = self.ai_api_key.get().strip()
        self.config.ai_review.api_key_env = self.ai_key_env.get().strip() or "DEEPSEEK_API_KEY"
        self.config.ai_review.auth_prefix = self.ai_auth_prefix.get().strip()
        self.config.ai_review.stream = self.ai_stream.get()
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
            raise ValueError("Please add files or folders first.")
        return list(self.document_paths)

    def _require_excel_input(self) -> str:
        path = self.excel_input.get().strip()
        if not path:
            raise ValueError("Please choose an input workbook.")
        if not Path(path).exists():
            raise FileNotFoundError(path)
        return path

    def _show_document_results(self, results) -> None:
        self._append(self.doc_log, summarize_results(results) + "\n")
        for result in results:
            if result.error:
                self._append(self.doc_log, f"[failed] {result.source}: {result.error}\n")
                continue
            if result.output:
                self._append(self.doc_log, f"Output: {result.output}\n")
            if result.json_report:
                self._append(self.doc_log, f"Report: {result.json_report}\n")
            if result.markdown_report:
                self._append(self.doc_log, f"Markdown report: {result.markdown_report}\n")
            if result.report:
                self._append(self.doc_log, result.report.summary() + "\n")
        self._append(self.doc_log, "\n")

    def open_folder(self, folder: str) -> None:
        if not folder:
            messagebox.showinfo("No folder", "Please choose a folder first.")
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
