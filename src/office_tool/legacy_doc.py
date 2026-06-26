"""Offline conversion of legacy .doc files through installed office software."""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Any


DOCX_FILE_FORMAT = 16
AUTOMATION_SECURITY_FORCE_DISABLE = 3
COM_BACKENDS = (
    ("Microsoft Word", "Word.Application"),
    ("WPS 文字", "kwps.Application"),
    ("WPS 文字", "wps.Application"),
)


class LegacyDocConversionError(RuntimeError):
    """Raised when no installed office application can convert a .doc file."""


def convert_legacy_doc(source: str | Path, destination: str | Path) -> Path:
    """Convert one legacy Word document to DOCX without network access."""
    source_path = Path(source).expanduser().resolve()
    destination_path = Path(destination).expanduser().resolve()
    if source_path.suffix.lower() != ".doc":
        raise ValueError(f"旧版文档转换只接受 .doc 文件: {source_path}")
    if destination_path.suffix.lower() != ".docx":
        raise ValueError(f"旧版文档转换目标必须是 .docx: {destination_path}")
    if not source_path.is_file():
        raise FileNotFoundError(source_path)
    if sys.platform != "win32":
        raise LegacyDocConversionError(".doc 离线转换仅支持 Windows，并且需要安装 Microsoft Word 或 WPS。")

    pythoncom, client = _load_com_modules()
    destination_path.parent.mkdir(parents=True, exist_ok=True)
    errors: list[str] = []
    pythoncom.CoInitialize()
    try:
        for display_name, prog_id in COM_BACKENDS:
            try:
                if destination_path.exists():
                    destination_path.unlink()
                _convert_with_application(client, prog_id, source_path, destination_path)
                if destination_path.is_file() and destination_path.stat().st_size > 0:
                    return destination_path
                raise RuntimeError("转换程序未生成 DOCX 文件")
            except Exception as exc:
                errors.append(f"{display_name}（{prog_id}）：{_friendly_error(exc)}")
    finally:
        pythoncom.CoUninitialize()

    if destination_path.exists():
        destination_path.unlink()
    detail = "；".join(errors)
    raise LegacyDocConversionError(
        "无法转换 .doc 文件。请确认电脑已安装可正常打开该文件的 Microsoft Word 或 WPS。"
        + (f" 详细信息：{detail}" if detail else "")
    )


def _load_com_modules() -> tuple[Any, Any]:
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise LegacyDocConversionError(
            "缺少 Windows COM 组件支持，请重新安装完整版本的 OfficeTool 或安装 pywin32。"
        ) from exc
    return pythoncom, win32com.client


def _convert_with_application(client: Any, prog_id: str, source: Path, destination: Path) -> None:
    application = None
    document = None
    try:
        try:
            application = client.DispatchEx(prog_id)
        except Exception as exc:
            raise LegacyDocConversionError("启动程序失败") from exc
        _set_if_supported(application, "Visible", False)
        _set_if_supported(application, "DisplayAlerts", 0)
        _set_if_supported(application, "AutomationSecurity", AUTOMATION_SECURITY_FORCE_DISABLE)
        try:
            document = application.Documents.Open(str(source), False, True, False)
        except Exception as exc:
            raise LegacyDocConversionError("打开 .doc 文件失败，请确认文件未损坏且未被占用") from exc
        try:
            save_as = getattr(document, "SaveAs2")
        except Exception:
            save_as = document.SaveAs
        try:
            save_as(str(destination), DOCX_FILE_FORMAT)
        except Exception as exc:
            raise LegacyDocConversionError("另存为 .docx 失败") from exc
    finally:
        if document is not None:
            try:
                document.Close(False)
            except Exception:
                pass
        if application is not None:
            try:
                application.Quit()
            except Exception:
                pass


def _set_if_supported(target: Any, name: str, value: Any) -> None:
    try:
        setattr(target, name, value)
    except Exception:
        pass


def _friendly_error(exc: Exception) -> str:
    chain: list[str] = []
    current: BaseException | None = exc
    while current is not None:
        message = str(current).strip()
        if message:
            chain.append(message)
        current = current.__cause__
    return "；".join(chain) if chain else exc.__class__.__name__
