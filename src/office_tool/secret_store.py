"""Windows user-scoped secret protection for local AI profiles."""

from __future__ import annotations

import base64
import ctypes
import sys
from ctypes import wintypes


class _DataBlob(ctypes.Structure):
    _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_byte))]


def protect_secret(value: str) -> str:
    if not value or sys.platform != "win32":
        return ""
    raw = value.encode("utf-8")
    buffer = ctypes.create_string_buffer(raw)
    source = _DataBlob(len(raw), ctypes.cast(buffer, ctypes.POINTER(ctypes.c_byte)))
    result = _DataBlob()
    if not ctypes.windll.crypt32.CryptProtectData(
        ctypes.byref(source),
        "OfficeTool AI Key",
        None,
        None,
        None,
        0x1,
        ctypes.byref(result),
    ):
        raise ctypes.WinError()
    try:
        protected = ctypes.string_at(result.pbData, result.cbData)
        return base64.b64encode(protected).decode("ascii")
    finally:
        ctypes.windll.kernel32.LocalFree(result.pbData)


def unprotect_secret(value: str) -> str:
    if not value or sys.platform != "win32":
        return ""
    try:
        raw = base64.b64decode(value, validate=True)
        buffer = ctypes.create_string_buffer(raw)
        source = _DataBlob(len(raw), ctypes.cast(buffer, ctypes.POINTER(ctypes.c_byte)))
        result = _DataBlob()
        if not ctypes.windll.crypt32.CryptUnprotectData(
            ctypes.byref(source),
            None,
            None,
            None,
            None,
            0x1,
            ctypes.byref(result),
        ):
            return ""
        try:
            return ctypes.string_at(result.pbData, result.cbData).decode("utf-8")
        finally:
            ctypes.windll.kernel32.LocalFree(result.pbData)
    except (ValueError, OSError):
        return ""
