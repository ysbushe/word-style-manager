"""Application configuration and user-library paths."""

from __future__ import annotations

import os
import shutil
from pathlib import Path

from src.ewt.config import APP_NAME, CONFIG_FILE, TEMPLATE_DIR


def config_path(application_dir):
    """Keep portable settings beside the application, independent of cwd."""
    return Path(application_dir) / CONFIG_FILE


def user_documents_dir():
    """Return the Windows Documents known folder, with a portable fallback."""
    if os.name == "nt":
        try:
            import winreg

            key_path = r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path) as key:
                value, _ = winreg.QueryValueEx(key, "Personal")
            expanded = os.path.expandvars(value)
            if expanded:
                return Path(expanded)
        except OSError:
            pass
    return Path.home() / "Documents"


def default_template_library_dir():
    return user_documents_dir() / "模板库"


def is_legacy_bundled_template_dir(path, application_dir):
    """Identify the old default that stored templates beside the executable."""
    candidate = Path(path)
    application_dir = Path(application_dir)
    try:
        if candidate.resolve() == (application_dir / TEMPLATE_DIR).resolve():
            return True
    except OSError:
        pass
    return candidate.name.casefold() == TEMPLATE_DIR.casefold() and candidate.parent.name == APP_NAME


def migrate_legacy_templates(source_dir, target_dir):
    """Copy old bundled templates without replacing files already in Documents."""
    source_dir = Path(source_dir)
    target_dir = Path(target_dir)
    if not source_dir.is_dir():
        return 0
    target_dir.mkdir(parents=True, exist_ok=True)
    copied = 0
    for source in source_dir.glob("*.dotx"):
        target = target_dir / source.name
        if target.exists():
            continue
        shutil.copy2(source, target)
        copied += 1
    return copied


def resolve_template_library_dir(saved_path, application_dir):
    """Use Documents by default and migrate only the former bundled default."""
    default_dir = default_template_library_dir()
    if not saved_path:
        return default_dir, False
    saved_dir = Path(saved_path)
    if not is_legacy_bundled_template_dir(saved_dir, application_dir):
        return saved_dir, False
    migrate_legacy_templates(saved_dir, default_dir)
    return default_dir, True
