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
from uuid import uuid4

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


def _safe_extract_zip(archive, extract_dir):
    """Reject path traversal, suspicious links, and unexpectedly large update packages."""
    extract_dir = Path(extract_dir).resolve()
    max_uncompressed_size = 1_500_000_000
    total_size = 0
    with zipfile.ZipFile(archive, "r") as zf:
        bad_member = zf.testzip()
        if bad_member:
            raise RuntimeError(f"更新包校验失败，压缩文件已损坏：{bad_member}")
        for item in zf.infolist():
            name = item.filename.replace("\\", "/")
            if not name or name.startswith("/") or ":" in name.split("/")[0]:
                raise RuntimeError(f"更新包包含不安全路径：{item.filename}")
            destination = (extract_dir / name).resolve()
            try:
                destination.relative_to(extract_dir)
            except ValueError as exc:
                raise RuntimeError(f"更新包包含越界路径：{item.filename}") from exc
            unix_mode = item.external_attr >> 16
            if unix_mode & 0o170000 == 0o120000:
                raise RuntimeError(f"更新包包含不支持的符号链接：{item.filename}")
            total_size += item.file_size
            if total_size > max_uncompressed_size:
                raise RuntimeError("更新包解压后体积异常，已停止更新。")
        zf.extractall(extract_dir)


def download_update(download_url, progress=None, expected_size=None):
    if not download_url:
        raise RuntimeError("该版本没有可下载的绿色版 ZIP 附件。")
    folder = Path(tempfile.mkdtemp(prefix="word_style_manager_update_"))
    try:
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
        if expected_size and received != int(expected_size):
            raise RuntimeError(f"更新包大小不完整：应为 {expected_size} 字节，实际为 {received} 字节。")
        if not zipfile.is_zipfile(archive):
            raise RuntimeError("下载内容不是有效的 ZIP 更新包。")
        extract_dir = folder / "extracted"
        extract_dir.mkdir()
        _safe_extract_zip(archive, extract_dir)
        payload = _find_payload_dir(extract_dir)
        executable = payload / f"{APP_NAME}.exe"
        if not executable.is_file() or executable.stat().st_size == 0:
            raise RuntimeError("更新包中的主程序无效。")
        return folder, payload
    except Exception:
        cleanup_download(folder)
        raise


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
    download_root = Path(download_root).resolve()
    payload_dir = Path(payload_dir).resolve()
    try:
        payload_dir.relative_to(download_root)
    except ValueError as exc:
        raise RuntimeError("更新文件不在受控的临时目录内，已停止安装。") from exc
    if not (payload_dir / f"{APP_NAME}.exe").is_file():
        raise RuntimeError("更新目录中缺少主程序，已停止安装。")
    script = Path(download_root) / "apply_update.ps1"
    backup = download_root / f"backup_{uuid4().hex}"
    source_ps = str(payload_dir).replace("'", "''")
    target_ps = str(app_directory).replace("'", "''")
    executable_ps = str(executable).replace("'", "''")
    download_root_ps = str(download_root).replace("'", "''")
    backup_ps = str(backup).replace("'", "''")
    script.write_text(
        "\n".join(
            [
                "$ErrorActionPreference = 'Stop'",
                f"$pidToWait = {os.getpid()}",
                f"$source = '{source_ps}'",
                f"$target = '{target_ps}'",
                f"$exe = '{executable_ps}'",
                f"$backup = '{backup_ps}'",
                "while (Get-Process -Id $pidToWait -ErrorAction SilentlyContinue) { Start-Sleep -Milliseconds 300 }",
                "$existing = @{}",
                "Get-ChildItem -LiteralPath $target -Recurse -File | ForEach-Object {",
                "  $relative = $_.FullName.Substring($target.Length).TrimStart('\\')",
                "  $existing[$relative] = $true",
                "}",
                "New-Item -ItemType Directory -Path $backup -Force | Out-Null",
                "robocopy $target $backup /E /R:2 /W:1 /XD templates /XF cleaner_config.json | Out-Null",
                "if ($LASTEXITCODE -gt 7) { throw \"无法创建更新备份，代码 $LASTEXITCODE\" }",
                "try {",
                "  robocopy $source $target /E /R:3 /W:1 /XD templates /XF cleaner_config.json | Out-Null",
                "  if ($LASTEXITCODE -gt 7) { throw \"复制新版文件失败，代码 $LASTEXITCODE\" }",
                "  if (-not (Test-Path -LiteralPath $exe)) { throw '更新后主程序不存在' }",
                "  Start-Process -FilePath $exe -ErrorAction Stop",
                "  Start-Sleep -Seconds 2",
                f"  Remove-Item -LiteralPath '{download_root_ps}' -Recurse -Force",
                "} catch {",
                "  Get-ChildItem -LiteralPath $source -Recurse -File | ForEach-Object {",
                "    $relative = $_.FullName.Substring($source.Length).TrimStart('\\')",
                "    if (-not $existing.ContainsKey($relative)) {",
                "      $newFile = Join-Path $target $relative",
                "      if (Test-Path -LiteralPath $newFile) { Remove-Item -LiteralPath $newFile -Force }",
                "    }",
                "  }",
                "  robocopy $backup $target /E /R:3 /W:1 /XD templates /XF cleaner_config.json | Out-Null",
                "  $message = $_.Exception.Message",
                "  Set-Content -LiteralPath (Join-Path $target 'update_error.txt') -Value $message -Encoding UTF8",
                "  if (Test-Path -LiteralPath $exe) { Start-Process -FilePath $exe }",
                "  exit 1",
                "}",
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
