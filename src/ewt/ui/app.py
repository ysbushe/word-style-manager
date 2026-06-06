"""主窗口 — Word 样式管理器 GUI"""

from __future__ import annotations

import json
import os
import sys
import threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

import ttkbootstrap as tb
from ttkbootstrap.constants import *
from tkinterdnd2 import DND_FILES, TkinterDnD

from src.ewt.config import APP_NAME, CONFIG_FILE, TEMPLATE_DIR
from src.ewt.ui.dialogs import FinishDialog, open_folder
from src.ewt.ui.widgets import PathRow
from src.ewt.core.style_engine import analyze_document, auto_clean_docx
from src.ewt.core.templates import (
    export_style_template,
    import_styles_to_document,
    batch_import_styles,
    list_template_library,
)


def app_dir():
    """获取应用根目录（源码或 PyInstaller 打包）"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    # src/ewt/ui/app.py → 项目根目录
    return Path(__file__).resolve().parent.parent.parent.parent


class StyleManagerApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.geometry(self._center_geometry(1040, 720))
        self.minsize(940, 640)

        self.style = tb.Style("cosmo")
        self.option_add("*Font", ("Microsoft YaHei UI", 10))

        self.config_data = self.load_config()
        self.output_dir = tk.StringVar(value=self.config_data.get("save_dir", "跟随源文件目录"))
        self.template_dir = app_dir() / TEMPLATE_DIR
        self.template_dir.mkdir(exist_ok=True)

        self.clean_files = []
        self.import_targets = []
        self.result_paths = []

        self._build_ui()
        self.refresh_templates()
        self.drop_target_register(DND_FILES)
        self.dnd_bind("<<Drop>>", self.on_drop)

    # ── 窗口 & 配置 ──────────────────────

    def _center_geometry(self, width, height):
        x = (self.winfo_screenwidth() - width) // 2
        y = (self.winfo_screenheight() - height) // 2
        return f"{width}x{height}+{x}+{y}"

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                return {}
        return {}

    def save_config(self):
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(self.config_data, f, ensure_ascii=False, indent=2)

    # ── UI 搭建 ──────────────────────────

    def _build_ui(self):
        shell = tb.Frame(self, padding=18)
        shell.pack(fill=BOTH, expand=True)

        header = tb.Frame(shell)
        header.pack(fill=X)
        tb.Label(header, text=APP_NAME, font=("Microsoft YaHei UI", 22, "bold")).pack(side=LEFT)
        tb.Label(header, text="清理、提取、模板化、批量导入 Word 样式", bootstyle=SECONDARY).pack(side=LEFT, padx=14, pady=(8, 0))
        tb.Button(header, text="打开结果位置", bootstyle=SECONDARY, command=self.open_last_result).pack(side=RIGHT)

        top = tb.Labelframe(shell, text="输出位置", padding=12)
        top.pack(fill=X, pady=(16, 12))
        tb.Label(top, textvariable=self.output_dir, bootstyle=INFO).pack(side=LEFT, fill=X, expand=True)
        tb.Button(top, text="更改目录", bootstyle=SECONDARY, command=self.change_output_dir).pack(side=RIGHT)
        tb.Button(top, text="跟随源文件", bootstyle=LIGHT, command=self.reset_output_dir).pack(side=RIGHT, padx=8)

        self.notebook = tb.Notebook(shell)
        self.notebook.pack(fill=BOTH, expand=True)
        self.clean_tab = tb.Frame(self.notebook, padding=14)
        self.export_tab = tb.Frame(self.notebook, padding=14)
        self.import_tab = tb.Frame(self.notebook, padding=14)
        self.library_tab = tb.Frame(self.notebook, padding=14)
        self.notebook.add(self.clean_tab, text="样式清理")
        self.notebook.add(self.export_tab, text="提取模板")
        self.notebook.add(self.import_tab, text="导入样式")
        self.notebook.add(self.library_tab, text="模板库")

        self._build_clean_tab()
        self._build_export_tab()
        self._build_import_tab()
        self._build_library_tab()

        self.status = tk.StringVar(value="就绪。可以拖拽 .doc / .docx 文件到窗口。")
        status_bar = tb.Frame(shell)
        status_bar.pack(fill=X, pady=(12, 0))
        tb.Label(status_bar, textvariable=self.status, bootstyle=SECONDARY).pack(side=LEFT)
        self.progress = tb.Progressbar(status_bar, mode="indeterminate", length=180)
        self.progress.pack(side=RIGHT)

    def _build_clean_tab(self):
        left = tb.Frame(self.clean_tab)
        left.pack(side=LEFT, fill=BOTH, expand=True)
        right = tb.Frame(self.clean_tab)
        right.pack(side=RIGHT, fill=BOTH, expand=True, padx=(14, 0))

        tb.Label(left, text="拖拽或选择一个/多个 Word 文件", font=("Microsoft YaHei UI", 13, "bold")).pack(anchor=W)
        tb.Button(left, text="添加文件", bootstyle=PRIMARY, command=self.add_clean_files).pack(anchor=W, pady=10)
        self.clean_list = tk.Listbox(left, height=11, activestyle="none")
        self.clean_list.pack(fill=BOTH, expand=True)
        btns = tb.Frame(left)
        btns.pack(fill=X, pady=10)
        tb.Button(btns, text="清空列表", bootstyle=LIGHT, command=self.clear_clean_files).pack(side=LEFT)
        tb.Button(btns, text="执行清理", bootstyle=SUCCESS, command=self.start_clean).pack(side=RIGHT)

        tb.Label(right, text="样式预览", font=("Microsoft YaHei UI", 13, "bold")).pack(anchor=W)
        self.clean_tree = tb.Treeview(right, columns=("name", "count", "type", "id"), show=HEADINGS, height=14)
        for col, text, width in (
            ("name", "样式名称", 180),
            ("count", "使用次数", 80),
            ("type", "类型", 80),
            ("id", "样式ID", 130),
        ):
            self.clean_tree.heading(col, text=text)
            self.clean_tree.column(col, width=width, anchor=W)
        self.clean_tree.pack(fill=BOTH, expand=True, pady=(10, 0))
        tb.Button(right, text="分析第一个文件", bootstyle=SECONDARY, command=self.preview_clean_styles).pack(anchor=E, pady=10)

    def _build_export_tab(self):
        self.export_source = tk.StringVar()
        self.template_name = tk.StringVar()
        PathRow(self.export_tab, "源文档", self.export_source, [("Word 文件", "*.doc *.docx")]).pack(fill=X, pady=(0, 10))

        name_row = tb.Frame(self.export_tab)
        name_row.pack(fill=X, pady=(0, 10))
        tb.Label(name_row, text="模板名称", width=12, anchor=W).pack(side=LEFT)
        tb.Entry(name_row, textvariable=self.template_name).pack(side=LEFT, fill=X, expand=True, padx=8)

        options = tb.Labelframe(self.export_tab, text="保存格式", padding=12)
        options.pack(fill=X, pady=8)
        self.export_dotx = tk.BooleanVar(value=True)
        self.export_docx = tk.BooleanVar(value=True)
        self.export_library = tk.BooleanVar(value=True)
        tb.Checkbutton(options, text="保存 .dotx Word 模板", variable=self.export_dotx, bootstyle=SUCCESS).pack(side=LEFT, padx=(0, 18))
        tb.Checkbutton(options, text="保存 .docx 样式文档", variable=self.export_docx, bootstyle=SUCCESS).pack(side=LEFT, padx=(0, 18))
        tb.Checkbutton(options, text="加入程序模板库", variable=self.export_library, bootstyle=SUCCESS).pack(side=LEFT)

        tb.Label(self.export_tab, text="导出时默认清理源文档中未被使用的样式，并保留必要依赖与编号。", bootstyle=SECONDARY).pack(anchor=W, pady=8)
        tb.Button(self.export_tab, text="提取并保存模板", bootstyle=SUCCESS, command=self.start_export).pack(anchor=E, pady=12)

    def _build_import_tab(self):
        self.import_source = tk.StringVar()
        self.conflict_mode = tk.StringVar(value="overwrite")
        self.clean_unused = tk.BooleanVar(value=True)

        PathRow(self.import_tab, "样式来源", self.import_source, [("Word/模板文件", "*.doc *.docx *.dotx")]).pack(fill=X, pady=(0, 10))

        row = tb.Frame(self.import_tab)
        row.pack(fill=BOTH, expand=True)
        left = tb.Frame(row)
        left.pack(side=LEFT, fill=BOTH, expand=True)
        right = tb.Labelframe(row, text="导入设置", padding=12)
        right.pack(side=RIGHT, fill=Y, padx=(14, 0))

        tb.Label(left, text="目标文件，可批量", font=("Microsoft YaHei UI", 13, "bold")).pack(anchor=W)
        self.import_list = tk.Listbox(left, height=12, activestyle="none")
        self.import_list.pack(fill=BOTH, expand=True, pady=10)
        target_btns = tb.Frame(left)
        target_btns.pack(fill=X)
        tb.Button(target_btns, text="添加目标文件", bootstyle=PRIMARY, command=self.add_import_targets).pack(side=LEFT)
        tb.Button(target_btns, text="清空", bootstyle=LIGHT, command=self.clear_import_targets).pack(side=LEFT, padx=8)

        tb.Label(right, text="同名样式").pack(anchor=W)
        tb.Radiobutton(right, text="覆盖目标样式", value="overwrite", variable=self.conflict_mode).pack(anchor=W, pady=(6, 0))
        tb.Radiobutton(right, text="跳过已有样式", value="skip", variable=self.conflict_mode).pack(anchor=W)
        tb.Radiobutton(right, text="自动重命名", value="rename", variable=self.conflict_mode).pack(anchor=W)
        tb.Separator(right).pack(fill=X, pady=12)
        tb.Checkbutton(right, text="默认清理未使用样式", variable=self.clean_unused, bootstyle=SUCCESS).pack(anchor=W)
        tb.Label(right, text="导入会同步编号/多级列表，并始终生成新文件。", wraplength=220, bootstyle=SECONDARY).pack(anchor=W, pady=12)
        tb.Button(right, text="开始导入", bootstyle=SUCCESS, command=self.start_import).pack(fill=X, pady=(12, 0))

    def _build_library_tab(self):
        top = tb.Frame(self.library_tab)
        top.pack(fill=X)
        tb.Label(top, text="程序模板库", font=("Microsoft YaHei UI", 13, "bold")).pack(side=LEFT)
        tb.Button(top, text="刷新", bootstyle=SECONDARY, command=self.refresh_templates).pack(side=RIGHT)
        tb.Button(top, text="打开模板目录", bootstyle=LIGHT, command=lambda: open_folder(str(self.template_dir))).pack(side=RIGHT, padx=8)

        self.library_tree = tb.Treeview(self.library_tab, columns=("name", "created", "source", "path"), show=HEADINGS)
        for col, text, width in (
            ("name", "模板名称", 170),
            ("created", "保存时间", 150),
            ("source", "来源", 260),
            ("path", "模板文件", 320),
        ):
            self.library_tree.heading(col, text=text)
            self.library_tree.column(col, width=width, anchor=W)
        self.library_tree.pack(fill=BOTH, expand=True, pady=12)
        tb.Button(self.library_tab, text="选中模板作为导入来源", bootstyle=PRIMARY, command=self.use_selected_template).pack(anchor=E)

    # ── 输出路径 ─────────────────────────

    def change_output_dir(self):
        folder = filedialog.askdirectory(title="选择输出目录")
        if folder:
            self.output_dir.set(folder)
            self.config_data["save_dir"] = folder
            self.save_config()

    def reset_output_dir(self):
        self.output_dir.set("跟随源文件目录")
        self.config_data["save_dir"] = "跟随源文件目录"
        self.save_config()

    def get_output_dir(self, source_path):
        selected = self.output_dir.get()
        if selected == "跟随源文件目录" or not os.path.exists(selected):
            return os.path.dirname(source_path)
        return selected

    def output_path(self, source_path, suffix):
        source = Path(source_path)
        folder = Path(self.get_output_dir(source_path))
        candidate = folder / f"{source.stem}{suffix}.docx"
        index = 2
        while candidate.exists():
            candidate = folder / f"{source.stem}{suffix}_{index}.docx"
            index += 1
        return str(candidate)

    # ── 拖拽 ────────────────────────────

    def on_drop(self, event):
        files = [f for f in self.tk.splitlist(event.data) if f.lower().endswith((".doc", ".docx", ".dotx"))]
        if not files:
            return
        tab = self.notebook.index(self.notebook.select())
        if tab == 0:
            self.clean_files.extend([f for f in files if f.lower().endswith((".doc", ".docx"))])
            self.refresh_listbox(self.clean_list, self.clean_files)
            self.status.set(f"已加入 {len(files)} 个文件到清理列表。")
        elif tab == 1:
            self.export_source.set(files[0])
            self.status.set("已设置模板提取源文档。")
        elif tab == 2:
            if not self.import_source.get():
                self.import_source.set(files[0])
            else:
                self.import_targets.extend([f for f in files if f.lower().endswith((".doc", ".docx"))])
                self.refresh_listbox(self.import_list, self.import_targets)
            self.status.set("已更新导入任务文件。")

    def refresh_listbox(self, widget, values):
        widget.delete(0, END)
        for item in values:
            widget.insert(END, item)

    # ── 样式清理 ────────────────────────

    def add_clean_files(self):
        files = filedialog.askopenfilenames(filetypes=[("Word 文件", "*.doc *.docx")])
        self.clean_files.extend(files)
        self.refresh_listbox(self.clean_list, self.clean_files)

    def clear_clean_files(self):
        self.clean_files = []
        self.clean_list.delete(0, END)
        self.clean_tree.delete(*self.clean_tree.get_children())

    def preview_clean_styles(self):
        if not self.clean_files:
            messagebox.showinfo(APP_NAME, "请先添加需要分析的 Word 文件。")
            return
        self.run_task(lambda: analyze_document(self.clean_files[0]), self.show_preview)

    def show_preview(self, result):
        self.clean_tree.delete(*self.clean_tree.get_children())
        success, builtin, custom = result
        if not success:
            self.status.set(f"分析失败：{builtin}")
            return
        for row in builtin + custom:
            name, count, style_type, sid = row
            self.clean_tree.insert("", END, values=(name, count, style_type, sid))
        self.status.set(f"分析完成：系统样式 {len(builtin)} 个，自定义样式 {len(custom)} 个。")

    def start_clean(self):
        if not self.clean_files:
            messagebox.showinfo(APP_NAME, "请先添加需要清理的 Word 文件。")
            return

        def work():
            outputs = []
            ok = 0
            for file_path in self.clean_files:
                out = self.output_path(file_path, "_样式清理版")
                success, _, _ = auto_clean_docx(file_path, out)
                if success:
                    ok += 1
                    outputs.append(out)
            return ok, outputs

        self.run_task(work, self.finish_clean)

    def finish_clean(self, result):
        ok, outputs = result
        self.result_paths = outputs
        self.status.set(f"清理完成：成功 {ok}/{len(self.clean_files)} 个文件。")
        FinishDialog(self, "清理完成", f"已生成 {ok} 个样式清理版文档。", outputs)

    # ── 模板导出 ────────────────────────

    def start_export(self):
        source = self.export_source.get().strip()
        if not source:
            messagebox.showinfo(APP_NAME, "请先选择源文档。")
            return
        formats = []
        if self.export_dotx.get():
            formats.append("dotx")
        if self.export_docx.get():
            formats.append("docx")
        if self.export_library.get():
            formats.append("library")
        if not formats:
            messagebox.showinfo(APP_NAME, "请至少选择一种保存格式。")
            return
        output_dir = self.get_output_dir(source)
        name = self.template_name.get().strip() or None

        def work():
            return export_style_template(source, output_dir, name, formats, str(self.template_dir))

        self.run_task(work, self.finish_export)

    def finish_export(self, result):
        if not result.get("success"):
            self.status.set(f"模板提取失败：{result.get('error')}")
            messagebox.showerror(APP_NAME, result.get("error", "模板提取失败"))
            return
        outputs = list(result["outputs"].values())
        self.result_paths = outputs
        self.refresh_templates()
        self.status.set(f"模板已保存，已清理未使用样式 {result.get('removed_unused', 0)} 个。")
        FinishDialog(self, "模板已保存", "样式模板已生成，并已按设置保存到模板库/输出目录。", outputs)

    # ── 样式导入 ────────────────────────

    def start_import(self):
        source = self.import_source.get().strip()
        if not source or not self.import_targets:
            messagebox.showinfo(APP_NAME, "请先选择样式来源和目标文件。")
            return
        out_dir = self.output_dir.get()
        output_dir = None if out_dir == "跟随源文件目录" else out_dir

        def work():
            if len(self.import_targets) == 1 and output_dir is None:
                return [import_styles_to_document(
                    source,
                    self.import_targets[0],
                    None,
                    self.conflict_mode.get(),
                    self.clean_unused.get(),
                )]
            return batch_import_styles(
                source,
                self.import_targets,
                output_dir,
                self.conflict_mode.get(),
                self.clean_unused.get(),
            )

        self.run_task(work, self.finish_import)

    def finish_import(self, results):
        ok_results = [item for item in results if item.get("success")]
        outputs = [item["output"] for item in ok_results]
        self.result_paths = outputs
        added = sum(item.get("added", 0) for item in ok_results)
        overwritten = sum(item.get("overwritten", 0) for item in ok_results)
        skipped = sum(item.get("skipped", 0) for item in ok_results)
        renamed = sum(item.get("renamed", 0) for item in ok_results)
        numbering = sum(item.get("numbering", 0) for item in ok_results)
        self.status.set(f"导入完成：成功 {len(ok_results)}/{len(results)}。新增 {added}，覆盖 {overwritten}，跳过 {skipped}，重命名 {renamed}。")
        msg = f"已生成 {len(outputs)} 个新文档。\n新增 {added} 个样式，覆盖 {overwritten} 个，跳过 {skipped} 个，重命名 {renamed} 个，同步编号 {numbering} 组。"
        FinishDialog(self, "导入完成", msg, outputs)

    # ── 模板库 ──────────────────────────

    def refresh_templates(self):
        if not hasattr(self, "library_tree"):
            return
        self.library_tree.delete(*self.library_tree.get_children())
        for item in list_template_library(self.template_dir):
            self.library_tree.insert("", END, values=(item.get("name"), item.get("created_at"), item.get("source"), item.get("path")))

    def use_selected_template(self):
        selected = self.library_tree.selection()
        if not selected:
            return
        values = self.library_tree.item(selected[0], "values")
        if len(values) >= 4:
            self.import_source.set(values[3])
            self.notebook.select(self.import_tab)
            self.status.set(f"已选择模板：{values[0]}")

    def open_last_result(self):
        if self.result_paths:
            open_folder(self.result_paths[0], select_file=True)
        else:
            open_folder(str(self.template_dir))

    # ── 后台任务 ────────────────────────

    def run_task(self, worker, callback):
        self.progress.start(12)
        self.status.set("正在处理，请稍候...")

        def runner():
            try:
                result = worker()
                self.after(0, lambda: callback(result))
            except Exception as e:
                self.after(0, lambda: messagebox.showerror(APP_NAME, str(e)))
                self.after(0, lambda: self.status.set(f"处理失败：{e}"))
            finally:
                self.after(0, self.progress.stop)

        threading.Thread(target=runner, daemon=True).start()
