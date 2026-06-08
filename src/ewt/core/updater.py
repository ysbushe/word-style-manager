"""GitHub Release update checks and portable-app replacement."""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import tempfile
import urllib.request
import webbrowser
import zipfile
from pathlib import Path

from src.ewt.config import APP_NAME, APP_VERSION, GITHUB_RELEASES_URL, GITHUB_REPOSITORY


API_URL = f"https://api.github.com/repos/{GITHUB_REPOSITORY}/releases/latest"


def _version_tuple(value):
    value = str(value or "").strip().lstrip("vV")
    parts = []
    for part in value.split("."):
        digits = "".join(char for char in part if char.isdigit())
        parts.append(int(digits or 0))
    return tuple((parts + [0, 0, 0])[:3])


def is_newer_version(latest, current=APP_VERSION):
    return _version_tuple(latest) > _version_tuple(current)


def _select_zip_asset(assets):
    zip_assets = [asset for asset in assets if asset.get("name", "").lower().endswith(".zip")]
    for asset in zip_assets:
        name = asset.get("name", "").lower()
        if "绿色版" in name or "portable" in name:
            return asset
    return zip_assets[0] if zip_assets else None


def check_for_updates(timeout=12):
    request = urllib.request.Request(
        API_URL,
        headers={
            "Accept": "application/vnd.github+json",
            "User-Agent": f"word-style-manager/{APP_VERSION}",
            "X-GitHub-Api-Version": "2022-11-28",
        },
    )
    with urllib.request.urlopen(request, timeout=timeout) as response:
        data = json.load(response)
    tag = data.get("tag_name", "")
    asset = _select_zip_asset(data.get("assets", []))
    return {
        "success": True,
        "current_version": APP_VERSION,
        "latest_version": tag.lstrip("vV"),
        "has_update": is_newer_version(tag),
        "release_name": data.get("name") or tag,
        "release_notes": data.get("body") or "",
        "release_url": data.get("html_url") or GITHUB_RELEASES_URL,
        "published_at": data.get("published_at") or "",
        "asset_name": asset.get("name") if asset else "",
        "download_url": asset.get("browser_download_url") if asset else "",
        "asset_size": asset.get("size", 0) if asset else 0,
    }


def open_repository(url):
    webbrowser.open(url)


def download_update(download_url, progress=None):
    if not download_url:
        raise RuntimeError("该版本没有可下载的绿色版 ZIP 附件。")
    folder = Path(tempfile.mkdtemp(prefix="word_style_manager_update_"))
    archive = folder / "update.zip"
    request = urllib.request.Request(download_url, headers={"User-Agent": f"word-style-manager/{APP_VERSION}"})
    with urllib.request.urlopen(request, timeout=60) as response, archive.open("wb") as output:
        total = int(response.headers.get("Content-Length") or 0)
        received = 0
        while True:
            chunk = response.read(1024 * 256)
            if not chunk:
                break
            output.write(chunk)
            received += len(chunk)
            if progress:
                progress(received, total)
    extract_dir = folder / "extracted"
    extract_dir.mkdir()
    with zipfile.ZipFile(archive, "r") as zf:
        zf.extractall(extract_dir)
    return folder, _find_payload_dir(extract_dir)


def _find_payload_dir(extract_dir):
    extract_dir = Path(extract_dir)
    direct_exe = extract_dir / f"{APP_NAME}.exe"
    if direct_exe.exists():
        return extract_dir
    candidates = list(extract_dir.rglob(f"{APP_NAME}.exe"))
    if not candidates:
        raise RuntimeError("更新包中未找到 Word 样式管理器.exe。")
    return candidates[0].parent


def install_downloaded_update(download_root, payload_dir):
    if not getattr(sys, "frozen", False):
        raise RuntimeError("源码运行模式不会自动替换文件，请前往 GitHub Releases 下载。")
    app_directory = Path(sys.executable).parent
    executable = Path(sys.executable)
    script = Path(download_root) / "apply_update.ps1"
    source_ps = str(Path(payload_dir)).replace("'", "''")
    target_ps = str(app_directory).replace("'", "''")
    executable_ps = str(executable).replace("'", "''")
    download_root_ps = str(Path(download_root)).replace("'", "''")
    script.write_text(
        "\n".join(
            [
                "$ErrorActionPreference = 'Stop'",
                f"$pidToWait = {os.getpid()}",
                f"$source = '{source_ps}'",
                f"$target = '{target_ps}'",
                f"$exe = '{executable_ps}'",
                "while (Get-Process -Id $pidToWait -ErrorAction SilentlyContinue) { Start-Sleep -Milliseconds 300 }",
                "robocopy $source $target /E /R:3 /W:1 /XD templates /XF cleaner_config.json | Out-Null",
                "if ($LASTEXITCODE -gt 7) { exit $LASTEXITCODE }",
                "Start-Process -FilePath $exe",
                f"Start-Sleep -Seconds 2; Remove-Item -LiteralPath '{download_root_ps}' -Recurse -Force",
            ]
        ),
        encoding="utf-8-sig",
    )
    subprocess.Popen(
        [
            "powershell.exe",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(script),
        ],
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )


def cleanup_download(path):
    shutil.rmtree(path, ignore_errors=True)
