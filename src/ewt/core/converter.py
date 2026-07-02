"""Batch Office/WPS format conversion core.

This module is intentionally independent from the Tkinter UI.  It converts
legacy Office formats into modern Open XML formats and returns structured
results that a future UI can display directly.
"""

from __future__ import annotations

import argparse
import gc
import json
import os
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

try:
    import pythoncom
    import win32com.client

    _HAS_WIN32COM = True
except ImportError:
    pythoncom = None
    win32com = None
    _HAS_WIN32COM = False


DEFAULT_ENGINE = "office"
ENGINE_LABELS = {"office": "Microsoft Office", "wps": "WPS"}
KIND_LABELS = {"word": "Word", "excel": "Excel"}
WORD_OPENXML_FORMAT = 16  # wdFormatDocumentDefault
EXCEL_OPENXML_FORMAT = 51  # xlOpenXMLWorkbook

ENGINE_PROGIDS = {
    "office": {
        "word": ("Word.Application.16", "Word.Application.15", "Word.Application.14", "Word.Application"),
        "excel": ("Excel.Application.16", "Excel.Application.15", "Excel.Application.14", "Excel.Application"),
    },
    "wps": {
        "word": ("KWPS.Application", "Word.Application"),
        "excel": ("KET.Application", "Excel.Application"),
    },
}

SUPPORTED_EXTENSIONS = {
    ".doc": ("word", ".docx"),
    ".xls": ("excel", ".xlsx"),
}


class ConversionError(RuntimeError):
    """Base error for conversion failures with user-facing messages."""


class EngineUnavailableError(ConversionError):
    """Raised when the requested conversion engine is unavailable or mismatched."""


@dataclass
class ConversionTask:
    source: Path
    output: Path
    kind: str
    target_ext: str


@dataclass
class EngineBinding:
    requested_engine: str
    kind: str
    progid: str
    product: str
    name: str
    version: str
    path: str

    def as_dict(self):
        return {
            "requested_engine": self.requested_engine,
            "kind": self.kind,
            "progid": self.progid,
            "product": self.product,
            "name": self.name,
            "version": self.version,
            "path": self.path,
        }


def _normalize_engine(engine):
    value = (engine or DEFAULT_ENGINE).strip().lower()
    if value not in ENGINE_LABELS:
        raise ValueError("转换引擎只能选择 office 或 wps。")
    return value


def _read_com_attr(application, attr):
    try:
        value = getattr(application, attr, "")
    except Exception:
        return ""
    if value is None:
        return ""
    return str(value)


def _application_identity(application):
    name = _read_com_attr(application, "Name")
    caption = _read_com_attr(application, "Caption")
    version = _read_com_attr(application, "Version")
    build = _read_com_attr(application, "Build")
    path = _read_com_attr(application, "Path")
    combined = " ".join([name, caption, version, build, path]).lower()

    if any(token in combined for token in ("kingsoft", "wps", "kwps", "ket")):
        product = "wps"
    elif "microsoft" in combined:
        product = "office"
    else:
        product = "unknown"

    return {
        "product": product,
        "name": name or caption,
        "version": version or build,
        "path": path,
    }


def _safe_quit(application):
    if application is None:
        return
    try:
        application.Quit()
    except Exception:
        pass


def _configure_quiet_application(application):
    for attr, value in (("Visible", False), ("DisplayAlerts", False)):
        try:
            setattr(application, attr, value)
        except Exception:
            pass


def _require_win32com():
    if not _HAS_WIN32COM:
        raise EngineUnavailableError(
            "当前 Python 环境缺少 pywin32，无法调用 Microsoft Office 或 WPS 进行格式转换。"
            "请先在本机可用环境中安装依赖，或使用已打包好的程序环境。"
        )


def _probe_engine_available(engine, kind):
    if not _HAS_WIN32COM:
        return False
    for progid in ENGINE_PROGIDS[engine][kind]:
        application = None
        try:
            application = win32com.client.DispatchEx(progid)
            identity = _application_identity(application)
            if identity["product"] == engine:
                return True
        except Exception:
            pass
        finally:
            _safe_quit(application)
            application = None
            gc.collect()
    return False


def _engine_unavailable_message(engine, kind, attempts):
    requested = ENGINE_LABELS[engine]
    opposite = "wps" if engine == "office" else "office"
    opposite_label = ENGINE_LABELS[opposite]
    kind_label = KIND_LABELS[kind]
    hint = ""
    if _probe_engine_available(opposite, kind):
        if engine == "office":
            hint = (
                f"检测到 {opposite_label} 可以处理 {kind_label} 文件，但默认不会自动改用 WPS。"
                "如需使用 WPS，请明确选择 --engine wps。"
            )
        else:
            hint = f"检测到 {opposite_label} 可以处理 {kind_label} 文件，可改用 --engine office。"
    else:
        hint = f"也没有检测到可用于 {kind_label} 文件的另一种转换引擎。"

    details = "\n".join(f"- {item}" for item in attempts)
    return (
        f"未能启动并核验 {requested} 的 {kind_label} 转换组件。{hint}"
        f"\n尝试记录：\n{details}"
    )


def _launch_application(engine, kind):
    _require_win32com()
    attempts = []
    for progid in ENGINE_PROGIDS[engine][kind]:
        application = None
        try:
            application = win32com.client.DispatchEx(progid)
            identity = _application_identity(application)
            product = identity["product"]
            if product != engine:
                actual = ENGINE_LABELS.get(product, "无法识别的 Office/WPS 程序")
                attempts.append(f"{progid}: 启动后识别为 {actual}，与所选引擎不一致")
                _safe_quit(application)
                continue
            _configure_quiet_application(application)
            return application, EngineBinding(
                requested_engine=engine,
                kind=kind,
                progid=progid,
                product=product,
                name=identity["name"],
                version=identity["version"],
                path=identity["path"],
            )
        except Exception as exc:
            attempts.append(f"{progid}: {exc}")
            _safe_quit(application)
        finally:
            application = None
            gc.collect()
    raise EngineUnavailableError(_engine_unavailable_message(engine, kind, attempts))


def _unique_output_path(source, target_ext, output_dir=None):
    directory = Path(output_dir).expanduser() if output_dir else source.parent
    stem = source.stem
    candidate = directory / f"{stem}{target_ext}"
    if not candidate.exists():
        return candidate

    candidate = directory / f"{stem}_转换版{target_ext}"
    index = 2
    while candidate.exists():
        candidate = directory / f"{stem}_转换版_{index}{target_ext}"
        index += 1
    return candidate


def _task_for_path(path, output_dir=None):
    source = Path(path).expanduser()
    if not source.exists():
        raise FileNotFoundError(f"找不到文件：{source}")
    if not source.is_file():
        raise ValueError(f"不是可转换的文件：{source}")

    suffix = source.suffix.lower()
    if suffix not in SUPPORTED_EXTENSIONS:
        raise ValueError("暂只支持 .doc 转 .docx，以及 .xls 转 .xlsx。")

    kind, target_ext = SUPPORTED_EXTENSIONS[suffix]
    output = _unique_output_path(source, target_ext, output_dir)
    return ConversionTask(source=source.resolve(), output=output.resolve(), kind=kind, target_ext=target_ext)


def _open_word_document(word, source):
    try:
        return word.Documents.Open(str(source), False, True, False)
    except Exception:
        return word.Documents.Open(str(source))


def _save_word_as_docx(word, task):
    document = None
    try:
        document = _open_word_document(word, task.source)
        save_as = getattr(document, "SaveAs2", None) or document.SaveAs
        _save_as_with_format(save_as, task.output, WORD_OPENXML_FORMAT)
    finally:
        if document is not None:
            try:
                document.Close(False)
            except Exception:
                pass


def _open_excel_workbook(excel, source):
    try:
        return excel.Workbooks.Open(str(source), UpdateLinks=0, ReadOnly=True, IgnoreReadOnlyRecommended=True)
    except Exception:
        return excel.Workbooks.Open(str(source))


def _save_excel_as_xlsx(excel, task):
    workbook = None
    try:
        workbook = _open_excel_workbook(excel, task.source)
        _save_as_with_format(workbook.SaveAs, task.output, EXCEL_OPENXML_FORMAT)
    finally:
        if workbook is not None:
            try:
                workbook.Close(False)
            except Exception:
                pass


def _save_as_with_format(save_as, output, file_format):
    try:
        save_as(str(output), FileFormat=file_format)
    except Exception:
        save_as(str(output), file_format)


def _convert_task(application, task):
    task.output.parent.mkdir(parents=True, exist_ok=True)
    if task.output.exists():
        raise FileExistsError(f"输出文件已存在，已停止以避免覆盖：{task.output}")
    try:
        if task.kind == "word":
            _save_word_as_docx(application, task)
        elif task.kind == "excel":
            _save_excel_as_xlsx(application, task)
        else:
            raise ValueError(f"不支持的转换类型：{task.kind}")
        if not task.output.exists():
            raise ConversionError("转换程序没有生成输出文件，可能是源文件受保护或已损坏。")
    except Exception:
        try:
            task.output.unlink(missing_ok=True)
        except OSError:
            pass
        raise


def _success_record(task, binding):
    return {
        "source": str(task.source),
        "output": str(task.output),
        "kind": task.kind,
        "engine": binding.requested_engine,
        "application": binding.as_dict(),
    }


def _failure_record(source, error, kind="", engine=""):
    return {
        "source": str(source),
        "kind": kind,
        "engine": engine,
        "error": str(error),
    }


def _empty_result(engine):
    return {
        "success": False,
        "engine": engine,
        "total": 0,
        "succeeded": 0,
        "failed_count": 0,
        "converted": [],
        "failed": [],
    }


def convert_files(paths: Iterable[str | os.PathLike], output_dir=None, engine=DEFAULT_ENGINE):
    """Convert .doc/.xls files and continue after per-file failures.

    Args:
        paths: Files to convert.
        output_dir: Optional directory for all converted files.  When omitted,
            each output is saved beside its source file.
        engine: ``"office"`` by default.  Use ``"wps"`` only when explicitly
            requested by the caller.

    Returns:
        A dictionary with ``converted`` and ``failed`` lists.  Existing outputs
        are never overwritten; name conflicts use ``_转换版``, then
        ``_转换版_2``, ``_转换版_3`` and so on.
    """

    selected_engine = _normalize_engine(engine)
    result = _empty_result(selected_engine)
    if isinstance(paths, (str, os.PathLike)):
        raw_paths = [paths]
    else:
        raw_paths = list(paths)
    tasks = []

    for raw_path in raw_paths:
        try:
            tasks.append(_task_for_path(raw_path, output_dir))
        except Exception as exc:
            result["failed"].append(_failure_record(raw_path, exc, engine=selected_engine))

    if not tasks:
        result["failed_count"] = len(result["failed"])
        result["total"] = len(raw_paths)
        return result

    if not _HAS_WIN32COM:
        message = (
            "当前 Python 环境缺少 pywin32，无法调用 Microsoft Office 或 WPS 进行格式转换。"
            "请先在本机可用环境中安装依赖，或使用已打包好的程序环境。"
        )
        for task in tasks:
            result["failed"].append(_failure_record(task.source, message, task.kind, selected_engine))
        result["total"] = len(raw_paths)
        result["failed_count"] = len(result["failed"])
        return result

    pythoncom.CoInitialize()
    try:
        for kind in ("word", "excel"):
            group = [task for task in tasks if task.kind == kind]
            if not group:
                continue
            application = None
            binding = None
            try:
                application, binding = _launch_application(selected_engine, kind)
                for task in group:
                    try:
                        _convert_task(application, task)
                        result["converted"].append(_success_record(task, binding))
                    except Exception as exc:
                        result["failed"].append(_failure_record(task.source, exc, task.kind, selected_engine))
            except Exception as exc:
                for task in group:
                    result["failed"].append(_failure_record(task.source, exc, task.kind, selected_engine))
            finally:
                _safe_quit(application)
                application = None
                binding = None
                gc.collect()
    finally:
        pythoncom.CoUninitialize()

    result["succeeded"] = len(result["converted"])
    result["failed_count"] = len(result["failed"])
    result["total"] = len(raw_paths)
    result["success"] = result["failed_count"] == 0 and result["succeeded"] > 0
    return result


def main(argv=None):
    parser = argparse.ArgumentParser(description="批量将 .doc/.xls 转换为 .docx/.xlsx。")
    parser.add_argument("files", nargs="+", help="要转换的 .doc 或 .xls 文件")
    parser.add_argument(
        "--engine",
        choices=sorted(ENGINE_LABELS),
        default=DEFAULT_ENGINE,
        help="转换引擎。默认 office；只有明确选择 wps 时才使用 WPS。",
    )
    parser.add_argument("--output-dir", help="输出目录。不填写则保存到源文件同目录。")
    args = parser.parse_args(argv)

    result = convert_files(args.files, output_dir=args.output_dir, engine=args.engine)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    if result["failed_count"] and result["succeeded"]:
        return 2
    if result["failed_count"]:
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
