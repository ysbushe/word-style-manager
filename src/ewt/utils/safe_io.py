"""Safe file-writing helpers for user configuration and generated artifacts."""

from __future__ import annotations

import json
import os
import tempfile
from datetime import datetime
from pathlib import Path


def atomic_write_bytes(path, data):
    """Write bytes beside the target, then atomically replace the target."""
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    fd, temp_name = tempfile.mkstemp(prefix=f".{target.name}.", suffix=".tmp", dir=target.parent)
    temp = Path(temp_name)
    try:
        with os.fdopen(fd, "wb") as handle:
            handle.write(data)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temp, target)
    except Exception:
        try:
            temp.unlink(missing_ok=True)
        except OSError:
            pass
        raise


def atomic_write_text(path, text, encoding="utf-8"):
    atomic_write_bytes(path, text.encode(encoding))


def atomic_write_json(path, value):
    atomic_write_text(path, json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")


def preserve_corrupt_file(path):
    """Move an unreadable data file aside so a later save cannot silently erase it."""
    source = Path(path)
    if not source.exists():
        return None
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = source.with_name(f"{source.name}.corrupt_{stamp}.bak")
    index = 2
    while backup.exists():
        backup = source.with_name(f"{source.name}.corrupt_{stamp}_{index}.bak")
        index += 1
    os.replace(source, backup)
    return backup
