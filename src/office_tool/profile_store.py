"""Persistent user-defined OfficeTool configuration profiles."""

from __future__ import annotations

import json
import uuid
from copy import deepcopy
from pathlib import Path

from .config import OfficeToolConfig


class ConfigProfileStore:
    def __init__(self, directory: str | Path):
        self.directory = Path(directory)

    def load_all(self) -> dict[str, OfficeToolConfig]:
        profiles: dict[str, OfficeToolConfig] = {}
        if not self.directory.exists():
            return profiles
        for path in sorted(self.directory.glob("*.json")):
            try:
                raw = json.loads(path.read_text(encoding="utf-8"))
                name = str(raw.get("name", "")).strip()
                config = raw.get("config")
                if name and isinstance(config, dict):
                    profiles[name] = OfficeToolConfig.from_dict(config)
            except (OSError, ValueError, TypeError, json.JSONDecodeError):
                continue
        return profiles

    def save(self, name: str, config: OfficeToolConfig) -> Path:
        profile_name = name.strip()
        if not profile_name:
            raise ValueError("配置名称不能为空。")
        self.directory.mkdir(parents=True, exist_ok=True)
        current = self._find_path(profile_name)
        path = current or self.directory / f"{uuid.uuid4().hex}.json"
        raw = deepcopy(config.to_dict())
        raw["ai_review"]["api_key"] = ""
        payload = {"name": profile_name, "config": raw}
        path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        return path

    def rename(self, old_name: str, new_name: str) -> Path:
        old_path = self._find_path(old_name)
        if old_path is None:
            raise KeyError(old_name)
        profile_name = new_name.strip()
        if not profile_name:
            raise ValueError("配置名称不能为空。")
        target = self._find_path(profile_name)
        if target is not None and target != old_path:
            target.unlink()
        raw = json.loads(old_path.read_text(encoding="utf-8"))
        raw["name"] = profile_name
        old_path.write_text(json.dumps(raw, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        return old_path

    def delete(self, name: str) -> bool:
        path = self._find_path(name)
        if path is None:
            return False
        path.unlink()
        return True

    def _find_path(self, name: str) -> Path | None:
        if not self.directory.exists():
            return None
        for path in self.directory.glob("*.json"):
            try:
                raw = json.loads(path.read_text(encoding="utf-8"))
            except (OSError, ValueError, TypeError, json.JSONDecodeError):
                continue
            if str(raw.get("name", "")).strip() == name.strip():
                return path
        return None
