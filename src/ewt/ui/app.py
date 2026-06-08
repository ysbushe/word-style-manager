"""Main GUI for Word Style Manager."""

from __future__ import annotations

import json
import os
import sys
import threading
import time
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog

import ttkbootstrap as tb
from ttkbootstrap.constants import *
from tkinterdnd2 import DND_FILES, TkinterDnD

from src.ewt.config import (
    APP_NAME,
    APP_VERSION,
    CONFIG_FILE,
    GITHUB_RELEASES_URL,
    GITHUB_URL,
    TEMPLATE_DIR,
)
from src.ewt.core.updater import (
    check_for_updates,
    cleanup_download,
    download_update,
    install_downloaded_update,
    open_repository,
)
from src.ewt.core.style_engine import auto_clean_docx, inspect_document
from src.ewt.core.templates import (
    batch_import_styles,
    export_style_template,
    list_task_schemes,
    list_template_library,
    save_task_scheme,
)
from src.ewt.ui.dialogs import FinishDialog, open_folder
from src.ewt.ui.template_editor import TemplateEditor
from src.ewt.ui.widgets import PathRow, StylePreviewPanel


def app_dir():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parent.parent.parent.parent


class StyleManagerApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.geometry(self._center_geometry(1240, 800))
        self.minsize(1080, 700)
        self.style = tb.Style("cosmo")
        self.option_add("*Font", ("Microsoft YaHei UI", 10))

        self.config_data = self.load_config()
        self.output_dir = tk.StringVar(value=self.config_data.get("save_dir", "跟随源文件目录"))
        self.template_dir = Path(self.config_data.get("template_dir", str(app_dir() / TEMPLATE_DIR)))
        self.template_dir.mkdir(parents=True, exist_ok=True)
        self.template_dir_var = tk.StringVar(value=str(self.template_dir))
        self.clean_files = list(self.config_data.get("clean_files", []))
        self.import_targets = list(self.config_data.get("import_targets", []))
        self.result_paths = []
        self.report_paths = []

        self.export_source = tk.StringVar(value=self.config_data.get("export_source", ""))
        self.import_source = tk.StringVar(value=self.config_data.get("import_source", ""))
        self.template_name = tk.StringVar(value=self.config_data.get("template_name", ""))
        self.conflict_mode = tk.StringVar(value=self.config_data.get("conflict_mode", "overwrite"))
        self.clean_unused = tk.BooleanVar(value=self.config_data.get("clean_unused", True))
        self.include_dependencies = tk.BooleanVar(value=self.config_data.get("include_dependencies", True))
        self.generate_report = tk.BooleanVar(value=self.config_data.get("generate_report", True))
        self.auto_update = tk.BooleanVar(value=self.config_data.get("auto_update", True))
        self.recursive_folder = tk.BooleanVar(value=self.config_data.get("recursive_folder", False))
        self.name_include = tk.StringVar(value=self.config_data.get("name_include", ""))

        self._build_ui()
        self.refresh_all_lists()
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.drop_target_register(DND_FILES)
        self.dnd_bind("<<Drop>>", self.on_drop)
        self.after(1800, self.check_update_on_start)

    def _center_geometry(self, width, height):
        return f"{width}x{height}+{(self.winfo_screenwidth()-width)//2}+{(self.winfo_screenheight()-height)//2}"

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                return {}
        return {}

    def save_config(self):
        self.config_data.update(
            {
                "save_dir": self.output_dir.get(),
                "template_dir": str(self.template_dir),
                "clean_files": self.clean_files,
                "import_targets": self.import_targets,
                "export_source": self.export_source.get(),
                "import_source": self.import_source.get(),
                "template_name": self.template_name.get(),
                "conflict_mode": self.conflict_mode.get(),
                "clean_unused": self.clean_unused.get(),
                "include_dependencies": self.include_dependencies.get(),
                "generate_report": self.generate_report.get(),
                "auto_update": self.auto_update.get(),
                "recursive_folder": self.recursive_folder.get(),
                "name_include": self.name_include.get(),
            }
        )
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(self.config_data, f, ensure_ascii=False, indent=2)

    def _build_ui(self):
        shell = tb.Frame(self, padding=16)
        shell.pack(fill=BOTH, expand=True)
        header = tb.Frame(shell)
        header.pack(fill=X)
        title = tb.Frame(header)
        title.pack(side=LEFT)
        title_line = tb.Frame(title)
        title_line.pack(anchor=W)
        tb.Label(title_line, text=APP_NAME, font=("Microsoft YaHei UI", 22, "bold"), bootstyle=PRIMARY).pack(side=LEFT)
        tb.Label(title_line, text=f"v{APP_VERSION}", bootstyle=SECONDARY).pack(side=LEFT, padx=10, pady=(7, 0))
        tb.Label(title, text="样式清理、模板提取、批量套用、多级编号编辑", bootstyle=SECONDARY).pack(anchor=W)
        actions = tb.Frame(header)
        actions.pack(side=RIGHT)
        tb.Button(actions, text="GitHub", bootstyle="outline-primary", command=lambda: open_repository(GITHUB_URL)).pack(side=LEFT, padx=4)
        tb.Button(actions, text="检查更新", bootstyle="outline-primary", command=lambda: self.check_updates(manual=True)).pack(side=LEFT, padx=4)
        tb.Button(actions, text="打开结果", bootstyle="outline-secondary", command=self.open_last_result).pack(side=LEFT, padx=4)
        tb.Button(actions, text="打开报告", bootstyle="outline-secondary", command=self.open_last_report).pack(side=LEFT, padx=4)

        top = tb.Frame(shell)
        top.pack(fill=X, pady=(14, 12))
        out_box = tb.Labelframe(top, text="输出位置", padding=10)
        out_box.pack(side=LEFT, fill=X, expand=True, padx=(0, 8))
        tb.Label(out_box, textvariable=self.output_dir, bootstyle=INFO).pack(side=LEFT, fill=X, expand=True)
        tb.Button(out_box, text="更改", bootstyle="outline-secondary", command=self.change_output_dir).pack(side=RIGHT)
        tb.Button(out_box, text="跟随源文件", bootstyle="link", command=self.reset_output_dir).pack(side=RIGHT, padx=4)
        tpl_box = tb.Labelframe(top, text="模板库", padding=10)
        tpl_box.pack(side=RIGHT, fill=X, expand=True)
        tb.Label(tpl_box, textvariable=self.template_dir_var, bootstyle=INFO).pack(side=LEFT, fill=X, expand=True)
        tb.Button(tpl_box, text="更改", bootstyle="outline-secondary", command=self.change_template_dir).pack(side=RIGHT)

        self.notebook = tb.Notebook(shell)
        self.notebook.pack(fill=BOTH, expand=True)
        self.clean_tab = tb.Frame(self.notebook, padding=12)
        self.export_tab = tb.Frame(self.notebook, padding=12)
        self.import_tab = tb.Frame(self.notebook, padding=12)
        self.library_tab = tb.Frame(self.notebook, padding=12)
        self.notebook.add(self.clean_tab, text="样式清理")
        self.notebook.add(self.export_tab, text="提取模板")
        self.notebook.add(self.import_tab, text="导入样式")
        self.notebook.add(self.library_tab, text="模板库")
        self.notebook.bind("<<NotebookTabChanged>>", lambda _e: self.save_config())

        self._build_clean_tab()
        self._build_export_tab()
        self._build_import_tab()
        self._build_library_tab()

        bottom = tb.Frame(shell)
        bottom.pack(fill=X, pady=(10, 0))
        self.status = tk.StringVar(value="就绪。文件和选项会在切换选项卡时保留。")
        tb.Label(bottom, textvariable=self.status, bootstyle=SECONDARY).pack(side=LEFT)
        self.progress = tb.Progressbar(bottom, length=180, mode="indeterminate")
        self.progress.pack(side=RIGHT)
        tb.Checkbutton(
            bottom,
            text="自动检查更新",
            variable=self.auto_update,
            bootstyle=SUCCESS,
            command=self.save_config,
        ).pack(side=RIGHT, padx=12)

    def _build_clean_tab(self):
        paned = tb.Panedwindow(self.clean_tab, orient=HORIZONTAL)
        paned.pack(fill=BOTH, expand=True)
        left = tb.Frame(paned, padding=(0, 0, 10, 0))
        right = tb.Frame(paned)
        paned.add(left, weight=1)
        paned.add(right, weight=2)
        tb.Label(left, text="待清理文件", font=("Microsoft YaHei UI", 13, "bold")).pack(anchor=W)
        self.clean_list = tk.Listbox(left, height=16, activestyle="none")
        self.clean_list.pack(fill=BOTH, expand=True, pady=8)
        row = tb.Frame(left)
        row.pack(fill=X)
        tb.Button(row, text="添加文件", bootstyle=PRIMARY, command=self.add_clean_files).pack(side=LEFT)
        tb.Button(row, text="添加文件夹", bootstyle="outline-primary", command=self.add_clean_folder).pack(side=LEFT, padx=6)
        tb.Button(row, text="清空", bootstyle="outline-secondary", command=self.clear_clean_files).pack(side=RIGHT)
        opts = tb.Labelframe(left, text="清理选项", padding=10)
        opts.pack(fill=X, pady=10)
        tb.Label(opts, text="默认清理未使用样式、项目符号和多级列表。", bootstyle=SECONDARY).pack(anchor=W)
        tb.Button(left, text="预览第一个文件", bootstyle="outline-secondary", command=self.preview_clean).pack(fill=X, pady=(8, 4))
        tb.Button(left, text="生成清理版文档", bootstyle=SUCCESS, command=self.start_clean).pack(fill=X)
        self.clean_preview = StylePreviewPanel(right, selectable=False, title="样式与编号预览")
        self.clean_preview.pack(fill=BOTH, expand=True)

    def _build_export_tab(self):
        PathRow(self.export_tab, "源文档", self.export_source, [("Word 文件", "*.doc *.docx")]).pack(fill=X)
        top = tb.Frame(self.export_tab)
        top.pack(fill=X, pady=8)
        tb.Label(top, text="模板名称", width=11).pack(side=LEFT)
        tb.Entry(top, textvariable=self.template_name).pack(side=LEFT, fill=X, expand=True, padx=8)
        tb.Button(top, text="读取并预览", bootstyle=PRIMARY, command=self.preview_export).pack(side=LEFT)
        options = tb.Frame(self.export_tab)
        options.pack(fill=X, pady=(0, 8))
        self.export_dotx = tk.BooleanVar(value=True)
        self.export_docx = tk.BooleanVar(value=True)
        self.export_library = tk.BooleanVar(value=True)
        tb.Checkbutton(options, text="保存 dotx", variable=self.export_dotx, bootstyle=SUCCESS).pack(side=LEFT)
        tb.Checkbutton(options, text="保存 docx", variable=self.export_docx, bootstyle=SUCCESS).pack(side=LEFT, padx=12)
        tb.Checkbutton(options, text="加入模板库", variable=self.export_library, bootstyle=SUCCESS).pack(side=LEFT)
        tb.Checkbutton(options, text="自动补选依赖", variable=self.include_dependencies, bootstyle=SUCCESS).pack(side=LEFT, padx=12)
        self.export_preview = StylePreviewPanel(self.export_tab, selectable=True, title="选择要保存的样式")
        self.export_preview.pack(fill=BOTH, expand=True)
        tb.Button(self.export_tab, text="保存为样式模板", bootstyle=SUCCESS, command=self.start_export).pack(anchor=E, pady=(8, 0))

    def _build_import_tab(self):
        paned = tb.Panedwindow(self.import_tab, orient=HORIZONTAL)
        paned.pack(fill=BOTH, expand=True)
        left = tb.Frame(paned, padding=(0, 0, 8, 0))
        middle = tb.Frame(paned, padding=(8, 0))
        right = tb.Frame(paned, padding=(8, 0, 0, 0))
        paned.add(left, weight=1)
        paned.add(middle, weight=2)
        paned.add(right, weight=1)

        tb.Label(left, text="模板库列表", font=("Microsoft YaHei UI", 12, "bold")).pack(anchor=W)
        self.import_template_list = tk.Listbox(left, height=8, activestyle="none")
        self.import_template_list.pack(fill=X, pady=7)
        self.import_template_list.bind("<<ListboxSelect>>", self.choose_import_template)
        PathRow(left, "外部来源", self.import_source, [("Word/模板文件", "*.doc *.docx *.dotx")]).pack(fill=X, pady=5)
        tb.Button(left, text="预览来源", bootstyle=PRIMARY, command=self.preview_import_source).pack(fill=X, pady=5)
        schemes = tb.Labelframe(left, text="任务方案", padding=8)
        schemes.pack(fill=X, pady=8)
        self.scheme_var = tk.StringVar()
        self.scheme_combo = tb.Combobox(schemes, textvariable=self.scheme_var, state="readonly")
        self.scheme_combo.pack(fill=X)
        row = tb.Frame(schemes)
        row.pack(fill=X, pady=(7, 0))
        tb.Button(row, text="保存当前", bootstyle="outline-primary", command=self.save_scheme).pack(side=LEFT)
        tb.Button(row, text="载入", bootstyle="outline-secondary", command=self.load_scheme).pack(side=RIGHT)

        self.import_preview = StylePreviewPanel(middle, selectable=True, title="来源样式预览")
        self.import_preview.pack(fill=BOTH, expand=True)

        tb.Label(right, text="目标文件", font=("Microsoft YaHei UI", 12, "bold")).pack(anchor=W)
        self.import_list = tk.Listbox(right, height=10, activestyle="none")
        self.import_list.pack(fill=BOTH, expand=True, pady=7)
        row = tb.Frame(right)
        row.pack(fill=X)
        tb.Button(row, text="添加文件", bootstyle=PRIMARY, command=self.add_import_targets).pack(side=LEFT)
        tb.Button(row, text="添加文件夹", bootstyle="outline-primary", command=self.add_import_folder).pack(side=LEFT, padx=5)
        tb.Button(row, text="清空", bootstyle="outline-secondary", command=self.clear_import_targets).pack(side=RIGHT)
        folder_opts = tb.Frame(right)
        folder_opts.pack(fill=X, pady=6)
        tb.Checkbutton(folder_opts, text="含子目录", variable=self.recursive_folder, bootstyle=SUCCESS).pack(side=LEFT)
        tb.Entry(folder_opts, textvariable=self.name_include, width=13).pack(side=RIGHT)
        tb.Label(folder_opts, text="文件名包含").pack(side=RIGHT, padx=4)
        settings = tb.Labelframe(right, text="导入设置", padding=10)
        settings.pack(fill=X, pady=8)
        for text, value in (("覆盖", "overwrite"), ("跳过", "skip"), ("重命名", "rename")):
            tb.Radiobutton(settings, text=text, value=value, variable=self.conflict_mode).pack(side=LEFT)
        tb.Checkbutton(settings, text="清理未用", variable=self.clean_unused, bootstyle=SUCCESS).pack(anchor=W, pady=(8, 0))
        tb.Checkbutton(settings, text="自动补依赖", variable=self.include_dependencies, bootstyle=SUCCESS).pack(anchor=W)
        tb.Checkbutton(settings, text="HTML报告", variable=self.generate_report, bootstyle=SUCCESS).pack(anchor=W)
        target_actions = tb.Frame(right)
        target_actions.pack(fill=X, pady=(6, 3))
        tb.Button(target_actions, text="预览目标", bootstyle="outline-secondary", command=self.preview_first_target).pack(side=LEFT, fill=X, expand=True)
        tb.Button(target_actions, text="导入预检", bootstyle="outline-warning", command=self.preflight_import).pack(side=LEFT, fill=X, expand=True, padx=(5, 0))
        tb.Button(right, text="批量生成导入版文档", bootstyle=SUCCESS, command=self.start_import).pack(fill=X)

    def _build_library_tab(self):
        top = tb.Frame(self.library_tab)
        top.pack(fill=X)
        tb.Button(top, text="刷新", bootstyle="outline-secondary", command=self.refresh_templates).pack(side=RIGHT)
        tb.Button(top, text="打开目录", bootstyle="outline-secondary", command=lambda: open_folder(str(self.template_dir))).pack(side=RIGHT, padx=6)
        tb.Button(top, text="编辑模板", bootstyle=PRIMARY, command=self.edit_selected_template).pack(side=RIGHT)
        tb.Button(top, text="应用到导入页", bootstyle=SUCCESS, command=self.use_selected_template).pack(side=RIGHT, padx=6)
        tb.Label(top, text="模板库", font=("Microsoft YaHei UI", 13, "bold")).pack(side=LEFT)
        paned = tb.Panedwindow(self.library_tab, orient=HORIZONTAL)
        paned.pack(fill=BOTH, expand=True, pady=8)
        left = tb.Frame(paned)
        right = tb.Frame(paned)
        paned.add(left, weight=1)
        paned.add(right, weight=2)
        self.library_tree = tb.Treeview(left, columns=("name", "created", "path"), show=HEADINGS)
        for col, text, width in (("name", "模板名称", 150), ("created", "保存时间", 135), ("path", "路径", 250)):
            self.library_tree.heading(col, text=text)
            self.library_tree.column(col, width=width)
        self.library_tree.pack(fill=BOTH, expand=True)
        self.library_tree.bind("<<TreeviewSelect>>", lambda _e: self.preview_selected_library_template())
        self.library_preview = StylePreviewPanel(right, selectable=False, title="模板样式与示例预览")
        self.library_preview.pack(fill=BOTH, expand=True)

    # paths and state
    def change_output_dir(self):
        folder = filedialog.askdirectory(title="选择输出目录")
        if folder:
            self.output_dir.set(folder)
            self.save_config()

    def reset_output_dir(self):
        self.output_dir.set("跟随源文件目录")
        self.save_config()

    def change_template_dir(self):
        folder = filedialog.askdirectory(title="选择模板库目录")
        if folder:
            self.template_dir = Path(folder)
            self.template_dir.mkdir(parents=True, exist_ok=True)
            self.template_dir_var.set(str(self.template_dir))
            self.save_config()
            self.refresh_all_lists()

    def get_output_dir(self, source_path):
        selected = self.output_dir.get()
        if selected == "跟随源文件目录" or not os.path.exists(selected):
            return os.path.dirname(source_path)
        return selected

    def output_path(self, source_path, suffix):
        folder = Path(self.get_output_dir(source_path))
        source = Path(source_path)
        candidate = folder / f"{source.stem}{suffix}.docx"
        index = 2
        while candidate.exists():
            candidate = folder / f"{source.stem}{suffix}_{index}.docx"
            index += 1
        return str(candidate)

    def report_dir(self):
        base = self.output_dir.get()
        if base == "跟随源文件目录" or not os.path.exists(base):
            base = str(self.template_dir)
        path = Path(base) / "reports"
        path.mkdir(parents=True, exist_ok=True)
        return str(path)

    # list helpers
    def refresh_all_lists(self):
        self.refresh_file_list(self.clean_list, self.clean_files)
        self.refresh_file_list(self.import_list, self.import_targets)
        self.refresh_templates()
        self.refresh_schemes()

    def refresh_file_list(self, widget, values):
        if not hasattr(self, "clean_list"):
            return
        widget.delete(0, END)
        for path in values:
            widget.insert(END, path)

    def refresh_templates(self):
        items = list_template_library(self.template_dir)
        self.template_items = items
        if hasattr(self, "library_tree"):
            self.library_tree.delete(*self.library_tree.get_children())
            for idx, item in enumerate(items):
                self.library_tree.insert("", END, iid=str(idx), values=(item.get("name"), item.get("created_at"), item.get("path")))
        if hasattr(self, "import_template_list"):
            self.import_template_list.delete(0, END)
            for item in items:
                self.import_template_list.insert(END, item.get("name"))

    def refresh_schemes(self):
        self.task_schemes = list_task_schemes(self.template_dir)
        if hasattr(self, "scheme_combo"):
            self.scheme_combo.configure(values=[item.get("name") for item in self.task_schemes])

    def add_files(self, target, widget, filetypes):
        files = filedialog.askopenfilenames(filetypes=filetypes)
        target.extend(files)
        self.refresh_file_list(widget, target)
        self.save_config()

    def add_folder_files(self, target, widget):
        folder = filedialog.askdirectory()
        if not folder:
            return
        pattern = "**/*" if self.recursive_folder.get() else "*"
        include = self.name_include.get().strip()
        for path in Path(folder).glob(pattern):
            if path.is_file() and path.suffix.lower() in {".doc", ".docx"}:
                if include and include not in path.name:
                    continue
                target.append(str(path))
        self.refresh_file_list(widget, target)
        self.save_config()

    def add_clean_files(self):
        self.add_files(self.clean_files, self.clean_list, [("Word 文件", "*.doc *.docx")])

    def add_clean_folder(self):
        self.add_folder_files(self.clean_files, self.clean_list)

    def clear_clean_files(self):
        self.clean_files.clear()
        self.refresh_file_list(self.clean_list, self.clean_files)
        self.save_config()

    def add_import_targets(self):
        self.add_files(self.import_targets, self.import_list, [("Word 文件", "*.doc *.docx")])

    def add_import_folder(self):
        self.add_folder_files(self.import_targets, self.import_list)

    def clear_import_targets(self):
        self.import_targets.clear()
        self.refresh_file_list(self.import_list, self.import_targets)
        self.save_config()

    # previews
    def preview_path_into(self, path, panel):
        if not path:
            messagebox.showinfo(APP_NAME, "请先选择文件。")
            return
        self.run_task(lambda: inspect_document(path), lambda result: self.finish_preview(result, panel))

    def finish_preview(self, result, panel):
        if not result.get("success"):
            messagebox.showerror(APP_NAME, result.get("error", "读取失败"))
            return
        panel.set_data(result.get("styles", []))
        self.status.set(f"预览完成：样式 {len(result.get('styles', []))} 个，编号 {len(result.get('numbering', []))} 组。")

    def preview_clean(self):
        if self.clean_files:
            self.preview_path_into(self.clean_files[0], self.clean_preview)

    def preview_export(self):
        self.preview_path_into(self.export_source.get(), self.export_preview)

    def preview_import_source(self):
        self.preview_path_into(self.import_source.get(), self.import_preview)

    def preview_first_target(self):
        if self.import_targets:
            self.run_task(
                lambda: inspect_document(self.import_targets[0]),
                lambda result: self.show_preview_window(result, f"目标预览 - {Path(self.import_targets[0]).name}"),
            )

    def show_preview_window(self, result, title):
        if not result.get("success"):
            messagebox.showerror(APP_NAME, result.get("error", "读取失败"))
            return
        window = tb.Toplevel(self)
        window.title(title)
        window.geometry("880x650")
        panel = StylePreviewPanel(window, selectable=False, title=title)
        panel.pack(fill=BOTH, expand=True, padx=14, pady=14)
        panel.set_data(result.get("styles", []))

    def preflight_import(self):
        source = self.import_source.get().strip()
        if not source or not self.import_targets:
            messagebox.showinfo(APP_NAME, "请先选择来源和至少一个目标文件。")
            return
        def work():
            src = inspect_document(source)
            tgt = inspect_document(self.import_targets[0])
            if not src.get("success") or not tgt.get("success"):
                return {"success": False, "error": src.get("error") or tgt.get("error")}
            selected = set(self.import_preview.get_selected_ids()) or {item["style_id"] for item in src["styles"]}
            target_ids = {item["style_id"] for item in tgt["styles"]}
            return {
                "success": True,
                "selected": len(selected),
                "conflicts": len(selected & target_ids),
                "new": len(selected - target_ids),
                "targets": len(self.import_targets),
            }
        self.run_task(work, self.finish_preflight)

    def finish_preflight(self, result):
        if not result.get("success"):
            messagebox.showerror(APP_NAME, result.get("error", "预检失败"))
            return
        messagebox.showinfo(
            "导入预检",
            f"目标文件：{result['targets']} 个\n"
            f"准备导入：{result['selected']} 个样式\n"
            f"同名冲突：{result['conflicts']} 个\n"
            f"新增样式：{result['new']} 个\n\n"
            "实际结果还会受到自动补选依赖和编号关联影响。",
        )

    def preview_selected_library_template(self):
        item = self.selected_template_item()
        if item:
            self.preview_path_into(item.get("path"), self.library_preview)

    # actions
    def start_clean(self):
        if not self.clean_files:
            messagebox.showinfo(APP_NAME, "请先添加文件。")
            return
        def work():
            outputs = []
            ok = 0
            numbering_removed = 0
            for file in self.clean_files:
                out = self.output_path(file, "_样式清理版")
                result = auto_clean_docx(file, out)
                if result[0]:
                    ok += 1
                    numbering_removed += result[2]
                    outputs.append(out)
            return ok, outputs, numbering_removed
        self.run_task(work, self.finish_clean)

    def finish_clean(self, result):
        ok, outputs, numbering_removed = result
        self.result_paths = outputs
        self.status.set(f"清理完成：{ok}/{len(self.clean_files)}，清理编号/多级列表 {numbering_removed} 项。")
        FinishDialog(self, "清理完成", f"已生成 {ok} 个清理版文档。", outputs)

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
        selected = self.export_preview.get_selected_ids()
        def work():
            return export_style_template(
                source,
                self.get_output_dir(source),
                self.template_name.get().strip() or None,
                formats,
                str(self.template_dir),
                selected or None,
                self.include_dependencies.get(),
            )
        self.run_task(work, self.finish_export)

    def finish_export(self, result):
        if not result.get("success"):
            messagebox.showerror(APP_NAME, result.get("error", "导出失败"))
            return
        self.result_paths = list(result.get("outputs", {}).values())
        self.status.set(f"模板已保存。清理样式 {result.get('removed_unused', 0)}，过滤未选 {result.get('filtered', 0)}。")
        self.refresh_templates()
        FinishDialog(self, "模板已保存", "模板已保存到指定位置，并按设置加入模板库。", self.result_paths)

    def start_import(self):
        if not self.import_source.get().strip() or not self.import_targets:
            messagebox.showinfo(APP_NAME, "请先选择来源和目标文件。")
            return
        output_dir = None if self.output_dir.get() == "跟随源文件目录" else self.output_dir.get()
        selected = self.import_preview.get_selected_ids()
        report_dir = self.report_dir() if self.generate_report.get() else None
        def work():
            return batch_import_styles(
                self.import_source.get().strip(),
                self.import_targets,
                output_dir,
                self.conflict_mode.get(),
                self.clean_unused.get(),
                selected or None,
                self.include_dependencies.get(),
                report_dir,
            )
        self.run_task(work, self.finish_import)

    def finish_import(self, results):
        ok = [item for item in results if item.get("success")]
        self.result_paths = [item["output"] for item in ok]
        self.report_paths = [item.get("report") for item in ok if item.get("report")]
        self.status.set(f"导入完成：成功 {len(ok)}/{len(results)}。报告 {len(self.report_paths)} 份。")
        FinishDialog(self, "导入完成", f"已生成 {len(ok)} 个导入版文档。", self.result_paths)

    # template library
    def selected_template_item(self):
        if not hasattr(self, "library_tree"):
            return None
        selection = self.library_tree.selection()
        if not selection:
            return None
        return self.template_items[int(selection[0])]

    def choose_import_template(self, event=None):
        sel = self.import_template_list.curselection()
        if not sel:
            return
        item = self.template_items[sel[0]]
        self.import_source.set(item.get("path"))
        self.preview_import_source()
        self.save_config()

    def use_selected_template(self):
        item = self.selected_template_item()
        if item:
            self.import_source.set(item.get("path"))
            self.notebook.select(self.import_tab)
            self.preview_import_source()
            self.save_config()

    def edit_selected_template(self):
        item = self.selected_template_item()
        if not item:
            messagebox.showinfo(APP_NAME, "请先选择模板。")
            return
        TemplateEditor(self, item.get("path"), self.template_dir, on_saved=lambda _p: self.refresh_templates())

    # task schemes
    def current_scheme_data(self):
        return {
            "source": self.import_source.get(),
            "targets": self.import_targets,
            "output_dir": self.output_dir.get(),
            "conflict": self.conflict_mode.get(),
            "clean_unused": self.clean_unused.get(),
            "include_dependencies": self.include_dependencies.get(),
            "generate_report": self.generate_report.get(),
        }

    def save_scheme(self):
        name = simpledialog.askstring(APP_NAME, "任务方案名称：", parent=self)
        if not name:
            return
        save_task_scheme(self.template_dir, name, self.current_scheme_data())
        self.refresh_schemes()

    def load_scheme(self):
        name = self.scheme_var.get()
        scheme = next((item for item in self.task_schemes if item.get("name") == name), None)
        if not scheme:
            return
        self.import_source.set(scheme.get("source", ""))
        self.import_targets = list(scheme.get("targets", []))
        self.output_dir.set(scheme.get("output_dir", "跟随源文件目录"))
        self.conflict_mode.set(scheme.get("conflict", "overwrite"))
        self.clean_unused.set(scheme.get("clean_unused", True))
        self.include_dependencies.set(scheme.get("include_dependencies", True))
        self.generate_report.set(scheme.get("generate_report", True))
        self.refresh_file_list(self.import_list, self.import_targets)
        self.save_config()

    # drag/drop
    def on_drop(self, event):
        files = [f for f in self.tk.splitlist(event.data) if f.lower().endswith((".doc", ".docx", ".dotx"))]
        if not files:
            return
        tab = self.notebook.index(self.notebook.select())
        if tab == 0:
            self.clean_files.extend([f for f in files if f.lower().endswith((".doc", ".docx"))])
            self.refresh_file_list(self.clean_list, self.clean_files)
        elif tab == 1:
            self.export_source.set(files[0])
        else:
            if not self.import_source.get():
                self.import_source.set(files[0])
            else:
                self.import_targets.extend([f for f in files if f.lower().endswith((".doc", ".docx"))])
                self.refresh_file_list(self.import_list, self.import_targets)
        self.save_config()

    # misc
    def open_last_result(self):
        if self.result_paths:
            open_folder(self.result_paths[0], select_file=True)
        else:
            open_folder(str(self.template_dir))

    def open_last_report(self):
        if self.report_paths:
            open_folder(self.report_paths[0], select_file=True)
        else:
            open_folder(self.report_dir())

    # GitHub and updates
    def check_update_on_start(self):
        if not self.auto_update.get():
            return
        last_check = float(self.config_data.get("last_update_check", 0) or 0)
        if time.time() - last_check < 24 * 60 * 60:
            return
        self.check_updates(manual=False)

    def check_updates(self, manual=False):
        if manual:
            self.status.set("正在连接 GitHub 检查更新...")
        def worker():
            try:
                return check_for_updates()
            except Exception as exc:
                return {"success": False, "error": str(exc)}
        self.run_task(worker, lambda result: self.finish_update_check(result, manual))

    def finish_update_check(self, result, manual):
        self.config_data["last_update_check"] = time.time()
        self.save_config()
        if not result.get("success"):
            if manual:
                messagebox.showerror("检查更新", f"无法连接 GitHub：\n{result.get('error', '未知错误')}")
            else:
                self.status.set("自动更新检查失败，不影响本地功能。")
            return
        if not result.get("has_update"):
            self.status.set(f"当前已是最新版本 v{APP_VERSION}。")
            if manual:
                messagebox.showinfo("检查更新", f"当前版本 v{APP_VERSION} 已是最新版本。")
            return
        notes = result.get("release_notes", "").strip()
        if len(notes) > 700:
            notes = notes[:700] + "\n..."
        prompt = (
            f"发现新版本 v{result.get('latest_version')}\n"
            f"当前版本 v{APP_VERSION}\n\n"
            f"{notes or '该版本未提供更新说明。'}\n\n"
        )
        if getattr(sys, "frozen", False) and result.get("download_url"):
            prompt += "是否立即下载并自动更新？程序会在更新后重新启动。"
            if messagebox.askyesno("发现新版本", prompt):
                self.download_and_install_update(result)
        else:
            prompt += "当前为源码模式或该版本没有绿色版附件，是否打开 GitHub Release 页面？"
            if messagebox.askyesno("发现新版本", prompt):
                open_repository(result.get("release_url") or GITHUB_RELEASES_URL)

    def download_and_install_update(self, release):
        self.status.set(f"正在下载 v{release.get('latest_version')}...")
        self.progress.start(10)
        def worker():
            try:
                root, payload = download_update(release["download_url"])
                return {"success": True, "root": root, "payload": payload}
            except Exception as exc:
                return {"success": False, "error": str(exc)}
        def finish(result):
            self.progress.stop()
            if not result.get("success"):
                messagebox.showerror("自动更新", f"更新下载失败：\n{result.get('error')}")
                return
            try:
                install_downloaded_update(result["root"], result["payload"])
            except Exception as exc:
                cleanup_download(result["root"])
                messagebox.showerror("自动更新", f"无法启动更新程序：\n{exc}")
                return
            self.status.set("更新已下载，程序即将关闭并自动重启。")
            self.after(500, self.destroy)
        threading.Thread(target=lambda: self.after(0, finish, worker()), daemon=True).start()

    def run_task(self, worker, callback):
        self.progress.start(12)
        self.status.set("正在处理，请稍候...")
        def runner():
            try:
                result = worker()
                self.after(0, lambda: callback(result))
            except Exception as exc:
                self.after(0, lambda: messagebox.showerror(APP_NAME, str(exc)))
                self.after(0, lambda: self.status.set(f"处理失败：{exc}"))
            finally:
                self.after(0, self.progress.stop)
        threading.Thread(target=runner, daemon=True).start()

    def on_close(self):
        self.save_config()
        self.destroy()


def main():
    app = StyleManagerApp()
    app.mainloop()
