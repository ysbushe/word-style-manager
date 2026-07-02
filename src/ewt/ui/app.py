"""Main GUI for Word Style Manager."""

from __future__ import annotations

import json
import os
import queue
import sys
import threading
import time
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

import ttkbootstrap as tb
from ttkbootstrap.constants import *
from tkinterdnd2 import DND_FILES, TkinterDnD

from src.ewt.config import (
    APP_NAME,
    APP_VERSION,
    GITHUB_RELEASES_URL,
    GITHUB_URL,
    USER_GUIDE_FILE,
)
from src.ewt.core import converter
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
    delete_template_from_library,
    export_style_template,
    list_template_library,
)
from src.ewt.ui.dialogs import FinishDialog, open_folder
from src.ewt.ui.template_editor import TemplateEditor
from src.ewt.ui.theme import COLORS, configure_listbox, configure_office_theme, configure_text
from src.ewt.ui.widgets import PathRow, StylePreviewPanel
from src.ewt.utils.paths import config_path, resolve_template_library_dir
from src.ewt.utils.safe_io import atomic_write_json, preserve_corrupt_file


CONVERTER_EXTENSIONS = {".doc", ".xls"}


def app_dir():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parent.parent.parent.parent


def merge_document_paths(*groups):
    """Merge Word document paths while preserving order and removing duplicates."""
    merged = []
    seen = set()
    for group in groups:
        if isinstance(group, (str, Path)):
            group = [group]
        for path in group or []:
            value = str(path).strip()
            if not value or Path(value).suffix.lower() not in {".doc", ".docx"}:
                continue
            key = os.path.normcase(os.path.abspath(value))
            if key not in seen:
                seen.add(key)
                merged.append(value)
    return merged


def merge_converter_paths(*groups):
    """Merge legacy Office paths while preserving order and removing duplicates."""
    merged = []
    seen = set()
    for group in groups:
        if isinstance(group, (str, Path)):
            group = [group]
        for path in group or []:
            value = str(path).strip()
            if not value or Path(value).suffix.lower() not in CONVERTER_EXTENSIONS:
                continue
            key = os.path.normcase(os.path.abspath(value))
            if key not in seen:
                seen.add(key)
                merged.append(value)
    return merged


def collect_converter_paths(entries):
    """Collect .doc/.xls files from files or folders and count ignored entries."""
    found = []
    ignored = 0
    for entry in entries or []:
        path = Path(str(entry).strip())
        if not str(path):
            continue
        if path.is_dir():
            for child in sorted(path.rglob("*")):
                if not child.is_file():
                    continue
                if child.suffix.lower() in CONVERTER_EXTENSIONS:
                    found.append(str(child))
                else:
                    ignored += 1
        elif path.is_file() and path.suffix.lower() in CONVERTER_EXTENSIONS:
            found.append(str(path))
        else:
            ignored += 1
    return found, ignored


class StyleManagerApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.geometry(self._center_geometry(1440, 900))
        self.minsize(1180, 760)
        self.style = tb.Style("cosmo")
        configure_office_theme(self.style)
        self.configure(background=COLORS["background"])
        self.option_add("*Font", ("Microsoft YaHei UI", 10))

        self.config_path = config_path(app_dir())
        self.config_data = self.load_config()
        self.output_dir = tk.StringVar(value=self.config_data.get("save_dir", "跟随源文件目录"))
        self.template_dir, template_dir_migrated = resolve_template_library_dir(
            self.config_data.get("template_dir"),
            app_dir(),
        )
        self.template_dir.mkdir(parents=True, exist_ok=True)
        if template_dir_migrated:
            self.config_data["template_dir"] = str(self.template_dir)
        self.template_dir_var = tk.StringVar(value=str(self.template_dir))
        saved_export_source = ""
        self.document_files = []
        self.clean_files = self.document_files
        self.import_targets = self.document_files
        self.converter_files = []
        self.result_paths = []
        self.report_paths = []
        self.preview_cache = {}
        self.preview_loading = set()
        self.preview_waiters = {}
        self.preview_results = queue.Queue()
        self.preview_executor = ThreadPoolExecutor(max_workers=2, thread_name_prefix="style-preview")
        self.closing = False
        self.settings_window = None
        self.converter_engine_state = {
            "office": {"word": None, "excel": None},
            "wps": {"word": None, "excel": None},
        }
        self.converter_engine_checking = False
        self.converter_engine_checked = False
        self.converter_running = False
        self.output_mode = tk.StringVar(
            value="source" if self.output_dir.get() == "跟随源文件目录" else "folder"
        )

        self.export_source = tk.StringVar(value=saved_export_source or (self.document_files[0] if self.document_files else ""))
        self.import_source = tk.StringVar(value=self.config_data.get("import_source", ""))
        self.converter_engine = tk.StringVar(value=self.config_data.get("converter_engine", "office"))
        if self.converter_engine.get() not in {"office", "wps"}:
            self.converter_engine.set("office")
        self.converter_office_selected = tk.BooleanVar(value=self.converter_engine.get() == "office")
        self.converter_wps_selected = tk.BooleanVar(value=self.converter_engine.get() == "wps")
        self.converter_output_mode = tk.StringVar(value=self.config_data.get("converter_output_mode", "source"))
        if self.converter_output_mode.get() not in {"source", "folder"}:
            self.converter_output_mode.set("source")
        self.converter_output_dir = tk.StringVar(value=self.config_data.get("converter_output_dir", ""))
        self.converter_count_text = tk.StringVar(value="0 个待转换文件")
        self.converter_engine_text = tk.StringVar(value="进入格式转换页后会自动检测本机可用的转换引擎。")
        self.converter_result_text = tk.StringVar(value="尚未开始转换。")
        self.template_name = tk.StringVar(value=self.config_data.get("template_name", ""))
        self.conflict_mode = tk.StringVar(value=self.config_data.get("conflict_mode", "overwrite"))
        self.clean_unused = tk.BooleanVar(value=self.config_data.get("clean_unused", True))
        self.clean_mode = tk.StringVar(value=self.config_data.get("clean_mode", "safe"))
        self.clean_mode_description = tk.StringVar()
        self.document_count_text = tk.StringVar(value="0 个文档")
        self.clean_unused_styles = tk.BooleanVar(value=self.config_data.get("clean_unused_styles", True))
        self.clean_unused_numbering = tk.BooleanVar(value=self.config_data.get("clean_unused_numbering", True))
        self.clean_unused_multilevel = tk.BooleanVar(value=self.config_data.get("clean_unused_multilevel", True))
        self.clean_duplicate_numbering = tk.BooleanVar(value=self.config_data.get("clean_duplicate_numbering", True))
        self.clean_repair_links = tk.BooleanVar(value=self.config_data.get("clean_repair_links", True))
        self.include_dependencies = tk.BooleanVar(value=self.config_data.get("include_dependencies", True))
        self.export_auto_clean = tk.BooleanVar(value=self.config_data.get("export_auto_clean", True))
        self.generate_report = tk.BooleanVar(value=self.config_data.get("generate_report", True))
        self.recursive_folder = tk.BooleanVar(value=self.config_data.get("recursive_folder", False))

        if template_dir_migrated:
            self.save_config()
        self._build_ui()
        if getattr(self, "config_recovery_path", None):
            self.after(
                100,
                lambda: messagebox.showwarning(
                    APP_NAME,
                    "原配置文件无法读取，程序已使用默认设置启动。\n\n"
                    f"损坏文件已保留为：\n{self.config_recovery_path}",
                    parent=self,
                ),
            )
        self.refresh_all_lists()
        self.after(25, self.process_preview_results)
        if self.document_files:
            self.document_list.selection_set(0)
            self.document_list.activate(0)
            self.on_document_select()
            self.prefetch_documents(self.document_files[1:])
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.bind("<Destroy>", self.on_root_destroy, add="+")
        self.drop_target_register(DND_FILES)
        self.dnd_bind("<<Drop>>", self.on_drop)

    def _center_geometry(self, width, height):
        return f"{width}x{height}+{(self.winfo_screenwidth()-width)//2}+{(self.winfo_screenheight()-height)//2}"

    def load_config(self):
        if self.config_path.exists():
            try:
                with self.config_path.open("r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                try:
                    self.config_recovery_path = preserve_corrupt_file(self.config_path)
                except OSError:
                    self.config_recovery_path = None
                return {}
        return {}

    def save_config(self):
        self.config_data.pop("clean_hide_builtin", None)
        self.config_data.update(
            {
                "save_dir": self.output_dir.get(),
                "template_dir": str(self.template_dir),
                "document_files": [],
                "clean_files": [],
                "import_targets": [],
                "export_source": "",
                "import_source": self.import_source.get(),
                "converter_engine": self.converter_engine.get(),
                "converter_output_mode": self.converter_output_mode.get(),
                "converter_output_dir": self.converter_output_dir.get(),
                "template_name": self.template_name.get(),
                "conflict_mode": self.conflict_mode.get(),
                "clean_unused": self.clean_unused.get(),
                "clean_mode": self.clean_mode.get(),
                "clean_unused_styles": self.clean_unused_styles.get(),
                "clean_unused_numbering": self.clean_unused_numbering.get(),
                "clean_unused_multilevel": self.clean_unused_multilevel.get(),
                "clean_duplicate_numbering": self.clean_duplicate_numbering.get(),
                "clean_repair_links": self.clean_repair_links.get(),
                "include_dependencies": self.include_dependencies.get(),
                "export_auto_clean": self.export_auto_clean.get(),
                "generate_report": self.generate_report.get(),
                "auto_update": False,
                "recursive_folder": self.recursive_folder.get(),
            }
        )
        try:
            atomic_write_json(self.config_path, self.config_data)
            return True
        except Exception as exc:
            try:
                messagebox.showerror(
                    APP_NAME,
                    f"设置保存失败，程序不会覆盖原配置。\n\n请检查目录权限或磁盘空间：\n{exc}",
                    parent=self,
                )
            except tk.TclError:
                pass
            return False

    def _build_ui(self):
        shell = tb.Frame(self, style="Office.TFrame")
        shell.pack(fill=BOTH, expand=True)

        sidebar = tb.Frame(shell, width=228, padding=(16, 20), style="Sidebar.TFrame")
        sidebar.pack(side=LEFT, fill=Y)
        sidebar.pack_propagate(False)
        brand = tb.Frame(sidebar, style="Sidebar.TFrame")
        brand.pack(fill=X, pady=(0, 24))
        mark = tk.Label(
            brand,
            text="W",
            width=3,
            height=1,
            background=COLORS["primary"],
            foreground="#FFFFFF",
            font=("Microsoft YaHei UI", 13, "bold"),
            relief="flat",
        )
        mark.pack(side=LEFT, padx=(0, 10))
        brand_text = tb.Frame(brand, style="Sidebar.TFrame")
        brand_text.pack(side=LEFT)
        tb.Label(brand_text, text="Word 样式管理器", style="SidebarTitle.TLabel").pack(anchor=W)
        tb.Label(brand_text, text=f"桌面版  v{APP_VERSION}", style="SidebarText.TLabel").pack(anchor=W)
        tb.Label(sidebar, text="工作区", style="SidebarText.TLabel").pack(anchor=W, padx=10, pady=(0, 6))
        self.nav_buttons = []
        for index, text in enumerate(("样式清理", "提取模板", "格式转换", "模板库")):
            button = tb.Button(
                sidebar,
                text=text,
                style="Nav.TButton",
                command=lambda value=index: self.select_workspace(value),
            )
            button.pack(fill=X, pady=1)
            self.nav_buttons.append(button)
        tb.Label(
            sidebar,
            text="安全原则\n所有处理均生成新文件，源文档保持不变。",
            wraplength=180,
            justify=LEFT,
            style="SidebarText.TLabel",
        ).pack(side=BOTTOM, anchor=W, padx=10, pady=(14, 0))

        separator = tk.Frame(shell, width=1, background=COLORS["border"])
        separator.pack(side=LEFT, fill=Y)

        content = tb.Frame(shell, padding=(22, 14, 22, 10), style="Content.TFrame")
        content.pack(side=LEFT, fill=BOTH, expand=True)
        header = tb.Frame(content, style="Content.TFrame")
        header.pack(fill=X)
        title = tb.Frame(header, style="Content.TFrame")
        title.pack(side=LEFT)
        self.page_title = tk.StringVar(value="样式清理")
        self.page_hint = tk.StringVar(value="检查并整理 Word/WPS 文档中的样式与编号")
        tb.Label(title, text="WORD WORKSPACE", style="Eyebrow.TLabel").pack(anchor=W)
        tb.Label(title, textvariable=self.page_title, style="PageTitle.TLabel").pack(anchor=W, pady=(2, 0))
        tb.Label(title, textvariable=self.page_hint, style="PageHint.TLabel").pack(anchor=W, pady=(2, 0))
        actions = tb.Frame(header, style="Content.TFrame")
        actions.pack(side=RIGHT)
        tb.Button(actions, text="GitHub", style="Ghost.TButton", command=lambda: open_repository(GITHUB_URL)).pack(side=LEFT, padx=2)
        tb.Button(actions, text="检查更新", style="Ghost.TButton", command=lambda: self.check_updates(manual=True)).pack(side=LEFT, padx=2)
        tb.Button(actions, text="使用说明", style="Ghost.TButton", command=self.open_user_guide).pack(side=LEFT, padx=2)
        tb.Button(actions, text="设置", style="Secondary.TButton", command=self.open_settings).pack(side=LEFT, padx=(8, 0))

        self.documents_panel = tb.Labelframe(content, text="文档任务", padding=10, style="Card.TLabelframe")
        self.documents_panel.pack(fill=X, pady=(10, 8))
        task_meta = tb.Frame(self.documents_panel, style="Surface.TFrame")
        task_meta.pack(fill=X, pady=(0, 8))
        tb.Label(
            task_meta,
            text="添加一个或多个 Word 文档，列表选择会同步到各工作区。",
            style="Muted.TLabel",
        ).pack(side=LEFT)
        tb.Label(task_meta, textvariable=self.document_count_text, style="Badge.TLabel").pack(side=RIGHT)
        document_list_frame = tb.Frame(self.documents_panel)
        document_list_frame.pack(side=LEFT, fill=BOTH, expand=True)
        self.document_list = tk.Listbox(document_list_frame, height=2, activestyle="none", exportselection=False)
        configure_listbox(self.document_list)
        document_scroll = tb.Scrollbar(document_list_frame, orient=VERTICAL, command=self.document_list.yview)
        self.document_list.configure(yscrollcommand=document_scroll.set)
        self.document_list.pack(side=LEFT, fill=BOTH, expand=True)
        document_scroll.pack(side=RIGHT, fill=Y)
        self.document_list.bind("<<ListboxSelect>>", self.on_document_select)
        document_actions = tb.Frame(self.documents_panel)
        document_actions.pack(side=RIGHT, padx=(10, 0))
        tb.Button(document_actions, text="添加文件", width=11, style="Primary.TButton", command=self.add_clean_files).grid(row=0, column=0, padx=3, pady=2)
        tb.Button(document_actions, text="添加文件夹", width=11, style="Secondary.TButton", command=self.add_clean_folder).grid(row=0, column=1, padx=3, pady=2)
        tb.Button(document_actions, text="移除选中", width=11, style="Secondary.TButton", command=self.remove_selected_document).grid(row=1, column=0, padx=3, pady=2)
        tb.Button(document_actions, text="清空列表", width=11, style="Danger.TButton", command=self.clear_clean_files).grid(row=1, column=1, padx=3, pady=2)

        self.notebook = tb.Notebook(content, style="Workspace.TNotebook")
        self.notebook.pack(fill=BOTH, expand=True)
        self.clean_tab = tb.Frame(self.notebook, padding=4, style="Content.TFrame")
        self.export_tab = tb.Frame(self.notebook, padding=4, style="Content.TFrame")
        self.converter_tab = tb.Frame(self.notebook, padding=4, style="Content.TFrame")
        self.import_tab = tb.Frame(self.notebook, padding=4, style="Content.TFrame")
        self.library_tab = tb.Frame(self.notebook, padding=4, style="Content.TFrame")
        self.notebook.add(self.clean_tab, text="样式清理")
        self.notebook.add(self.export_tab, text="提取模板")
        self.notebook.add(self.converter_tab, text="格式转换")
        self.notebook.add(self.library_tab, text="模板库")
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)

        self._build_clean_tab()
        self._build_export_tab()
        self._build_converter_tab()
        self._build_import_tab()
        self._build_library_tab()

        bottom = tb.Frame(content, padding=(12, 8), style="Surface.TFrame")
        bottom.pack(fill=X, pady=(10, 0))
        self.status = tk.StringVar(value="就绪。文件和选项会在切换选项卡时保留。")
        tb.Label(bottom, textvariable=self.status, bootstyle=SECONDARY).pack(side=LEFT)
        self.progress = tb.Progressbar(bottom, length=180, mode="indeterminate")
        self.progress.pack(side=RIGHT)
        self.select_workspace(0)

    def select_workspace(self, index):
        titles = (
            ("样式清理", "检查并整理 Word/WPS 文档中的样式与编号"),
            ("提取模板", "从现有文档提取可重复使用的轻量 Word 模板"),
            ("格式转换", "将旧版 Office 文件转换为更稳定的新格式"),
            ("模板库", "集中预览、管理和编辑已保存的模板"),
        )
        self.notebook.select(index)
        self.page_title.set(titles[index][0])
        self.page_hint.set(titles[index][1])
        for button_index, button in enumerate(self.nav_buttons):
            button.configure(style="NavActive.TButton" if button_index == index else "Nav.TButton")

    def _build_clean_tab(self):
        paned = tb.Panedwindow(self.clean_tab, orient=HORIZONTAL)
        paned.pack(fill=BOTH, expand=True)
        preview = tb.Frame(paned, padding=(0, 0, 10, 0))
        controls = tb.Frame(paned, width=300)
        controls.pack_propagate(False)
        paned.add(preview, weight=5)
        paned.add(controls, weight=1)
        self.clean_preview = StylePreviewPanel(preview, selectable=False, title="2、样式与编号预览")
        self.clean_preview.pack(fill=BOTH, expand=True)
        mode_box = tb.Labelframe(controls, text="清理模式", padding=12, style="Card.TLabelframe")
        mode_box.pack(fill=X)
        for text, value in (("安全清理（推荐）", "safe"), ("深度清理", "deep")):
            tb.Radiobutton(
                mode_box,
                text=text,
                value=value,
                variable=self.clean_mode,
                command=self.on_clean_mode_changed,
            ).pack(anchor=W, pady=2)
        tb.Label(
            mode_box,
            textvariable=self.clean_mode_description,
            wraplength=280,
            justify=LEFT,
            style="Muted.TLabel",
        ).pack(fill=X, pady=(8, 0))

        opts = tb.Labelframe(controls, text="清理范围", padding=12, style="Card.TLabelframe")
        opts.pack(fill=X, pady=(8, 0))
        for text, variable in (
            ("清理未使用样式", self.clean_unused_styles),
            ("清理未使用项目符号和编号", self.clean_unused_numbering),
            ("清理未使用多级列表", self.clean_unused_multilevel),
            ("合并重复编号定义", self.clean_duplicate_numbering),
            ("修复断开的样式与编号关联", self.clean_repair_links),
        ):
            tb.Checkbutton(
                opts,
                text=text,
                variable=variable,
                command=self.save_config,
            ).pack(anchor=W)
        tb.Label(
            controls,
            text="源文档永不覆盖。只有通过压缩包、XML、样式引用和编号引用复检的新副本才会保留。",
            wraplength=280,
            justify=LEFT,
            style="PageHint.TLabel",
        ).pack(fill=X, pady=(10, 2))
        tb.Button(
            controls,
            text="清理选中文档",
            style="Primary.TButton",
            command=self.start_clean,
        ).pack(fill=X, pady=(10, 5))
        tb.Button(
            controls,
            text="批量清理全部文档",
            style="Secondary.TButton",
            command=self.start_batch_clean,
        ).pack(fill=X)
        self.on_clean_mode_changed(save=False)

    def on_clean_mode_changed(self, save=True):
        if self.clean_mode.get() == "deep":
            self.clean_mode_description.set(
                "在保护正文、默认样式、继承链和有效编号的前提下，"
                "删除更多未使用内置样式，并解除不影响现有排版的后续/链接关系。"
            )
        else:
            self.clean_mode.set("safe")
            self.clean_mode_description.set(
                "保留内置样式及完整编辑关系；未使用内置样式仅隐藏。"
                "兼容性最高，适合日常文档。"
            )
        if save:
            self.save_config()

    def _build_export_tab(self):
        source = tb.Labelframe(self.export_tab, text="当前源文档", padding=10, style="Card.TLabelframe")
        source.pack(fill=X, pady=(0, 6))
        tb.Label(source, textvariable=self.export_source, style="Muted.TLabel").pack(side=LEFT, fill=X, expand=True)
        tb.Button(source, text="添加文档", style="Secondary.TButton", command=self.pick_export_source).pack(side=RIGHT)
        top = tb.Labelframe(self.export_tab, text="模板设置", padding=10, style="Card.TLabelframe")
        top.pack(fill=X, pady=(0, 6))
        top.columnconfigure(1, weight=1)
        tb.Label(top, text="模板名称", width=11).grid(row=0, column=0, sticky=W)
        tb.Entry(top, textvariable=self.template_name).grid(row=0, column=1, sticky=EW, padx=8)
        tb.Button(
            top,
            text="保存当前文档模板",
            style="Primary.TButton",
            command=self.start_export,
        ).grid(row=0, column=3, sticky=E)
        tb.Button(
            top,
            text="批量保存全部模板",
            style="Secondary.TButton",
            command=self.start_batch_export,
        ).grid(row=0, column=2, sticky=E, padx=(0, 8))
        options = tb.Frame(top, style="Card.TFrame")
        options.grid(row=1, column=0, columnspan=4, sticky=EW, pady=(10, 0))
        self.export_dotx = tk.BooleanVar(value=True)
        self.export_library = tk.BooleanVar(value=True)
        tb.Checkbutton(
            options,
            text="保存前清理未使用内容",
            variable=self.export_auto_clean,
            command=self.save_config,
        ).pack(side=LEFT)
        tb.Checkbutton(options, text="自动补全关联样式", variable=self.include_dependencies).pack(side=LEFT, padx=(18, 0))
        tb.Label(options, text="输出为轻量 .dotx 并保存到模板库", style="Muted.TLabel").pack(side=RIGHT)
        self.export_preview = StylePreviewPanel(
            self.export_tab,
            selectable=True,
            title="4、选择要保存的样式",
            show_details=False,
        )
        self.export_preview.pack(fill=BOTH, expand=True)

    def _build_converter_tab(self):
        intro = tb.Frame(self.converter_tab, padding=12, style="Card.TFrame")
        intro.pack(fill=X, pady=(0, 8))
        tb.Label(intro, text="批量格式转换", style="CardTitle.TLabel").pack(anchor=W)
        tb.Label(
            intro,
            text=(
                "将旧版 .doc / .xls 转换为 .docx / .xlsx，便于 agent 和其他工具稳定读取。"
                "添加文件，确认转换引擎与输出位置，然后开始转换；源文件不会被覆盖。"
            ),
            wraplength=920,
            justify=LEFT,
            style="Muted.TLabel",
        ).pack(anchor=W, pady=(4, 0))

        paned = tb.Panedwindow(self.converter_tab, orient=HORIZONTAL)
        paned.pack(fill=BOTH, expand=True)
        left = tb.Frame(paned, padding=(0, 0, 10, 0), style="Content.TFrame")
        right = tb.Frame(paned, width=360, style="Content.TFrame")
        right.pack_propagate(False)
        paned.add(left, weight=5)
        paned.add(right, weight=2)

        file_box = tb.Labelframe(left, text="待转换文件", padding=10, style="Card.TLabelframe")
        file_box.pack(fill=BOTH, expand=True)
        file_top = tb.Frame(file_box, style="Surface.TFrame")
        file_top.pack(fill=X, pady=(0, 8))
        tb.Label(file_top, text="支持 .doc 和 .xls；拖入文件或文件夹也会自动筛选。", style="Muted.TLabel").pack(side=LEFT)
        tb.Label(file_top, textvariable=self.converter_count_text, style="Badge.TLabel").pack(side=RIGHT)

        list_frame = tb.Frame(file_box, style="Surface.TFrame")
        list_frame.pack(fill=BOTH, expand=True)
        self.converter_tree = tb.Treeview(
            list_frame,
            columns=("name", "type", "folder"),
            show=HEADINGS,
            selectmode="extended",
        )
        for col, text, width, anchor in (
            ("name", "文件名", 260, W),
            ("type", "转换", 95, CENTER),
            ("folder", "所在位置", 520, W),
        ):
            self.converter_tree.heading(col, text=text)
            self.converter_tree.column(col, width=width, minwidth=80, stretch=col == "folder", anchor=anchor)
        y_scroll = tb.Scrollbar(list_frame, orient=VERTICAL, command=self.converter_tree.yview)
        x_scroll = tb.Scrollbar(list_frame, orient=HORIZONTAL, command=self.converter_tree.xview)
        self.converter_tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        self.converter_tree.grid(row=0, column=0, sticky=NSEW)
        y_scroll.grid(row=0, column=1, sticky=NS)
        x_scroll.grid(row=1, column=0, sticky=EW)
        list_frame.rowconfigure(0, weight=1)
        list_frame.columnconfigure(0, weight=1)

        file_actions = tb.Frame(file_box, style="Surface.TFrame")
        file_actions.pack(fill=X, pady=(8, 0))
        tb.Button(file_actions, text="添加文件", style="Primary.TButton", command=self.pick_converter_files).pack(side=LEFT)
        tb.Button(file_actions, text="添加文件夹", style="Secondary.TButton", command=self.pick_converter_folder).pack(side=LEFT, padx=6)
        tb.Button(file_actions, text="移除选中", style="Secondary.TButton", command=self.remove_selected_converter_files).pack(side=RIGHT)
        tb.Button(file_actions, text="清空列表", style="Danger.TButton", command=self.clear_converter_files).pack(side=RIGHT, padx=6)

        engine_box = tb.Labelframe(right, text="转换引擎", padding=12, style="Card.TLabelframe")
        engine_box.pack(fill=X)
        self.converter_office_check = tb.Checkbutton(
            engine_box,
            text="使用 Microsoft Office",
            variable=self.converter_office_selected,
            command=lambda: self.select_converter_engine("office"),
        )
        self.converter_office_check.pack(anchor=W)
        self.converter_wps_check = tb.Checkbutton(
            engine_box,
            text="使用 WPS",
            variable=self.converter_wps_selected,
            command=lambda: self.select_converter_engine("wps"),
        )
        self.converter_wps_check.pack(anchor=W, pady=(3, 0))
        tb.Label(
            engine_box,
            textvariable=self.converter_engine_text,
            wraplength=320,
            justify=LEFT,
            style="Muted.TLabel",
        ).pack(fill=X, pady=(8, 0))
        tb.Button(
            engine_box,
            text="重新检测",
            style="Ghost.TButton",
            command=lambda: self.ensure_converter_engine_check(force=True),
        ).pack(anchor=E, pady=(6, 0))

        output_box = tb.Labelframe(right, text="输出位置", padding=12, style="Card.TLabelframe")
        output_box.pack(fill=X, pady=(8, 0))
        tb.Radiobutton(
            output_box,
            text="输出到原文件同目录",
            value="source",
            variable=self.converter_output_mode,
            command=self.save_config,
        ).pack(anchor=W)
        tb.Radiobutton(
            output_box,
            text="输出到指定文件夹",
            value="folder",
            variable=self.converter_output_mode,
            command=self.save_config,
        ).pack(anchor=W, pady=(3, 0))
        output_row = tb.Frame(output_box, style="Surface.TFrame")
        output_row.pack(fill=X, pady=(8, 0))
        tb.Entry(output_row, textvariable=self.converter_output_dir).pack(side=LEFT, fill=X, expand=True)
        tb.Button(output_row, text="选择", style="Secondary.TButton", command=self.pick_converter_output_dir).pack(side=RIGHT, padx=(6, 0))

        action_box = tb.Labelframe(right, text="执行与结果", padding=12, style="Card.TLabelframe")
        action_box.pack(fill=BOTH, expand=True, pady=(8, 0))
        self.converter_start_button = tb.Button(
            action_box,
            text="开始转换",
            style="Primary.TButton",
            command=self.start_converter,
        )
        self.converter_start_button.pack(fill=X)
        self.converter_open_button = tb.Button(
            action_box,
            text="打开结果位置",
            style="Secondary.TButton",
            command=self.open_converter_result_location,
            state=DISABLED,
        )
        self.converter_open_button.pack(fill=X, pady=(6, 0))
        tb.Label(
            action_box,
            textvariable=self.converter_result_text,
            wraplength=320,
            justify=LEFT,
            style="Muted.TLabel",
        ).pack(fill=X, pady=(8, 4))
        self.converter_result_detail = tk.Text(action_box, height=9, wrap="word", relief="flat")
        configure_text(self.converter_result_detail, COLORS["surface_alt"])
        self.converter_result_detail.pack(fill=BOTH, expand=True)
        self.converter_result_detail.configure(state=DISABLED)

    def _build_import_tab(self):
        paned = tb.Panedwindow(self.import_tab, orient=HORIZONTAL)
        paned.pack(fill=BOTH, expand=True)
        left = tb.Frame(paned, padding=12, width=430, style="Card.TFrame")
        workspace = tb.Frame(paned, style="Content.TFrame")
        paned.add(left, weight=1)
        paned.add(workspace, weight=4)

        tb.Label(left, text="选择样式来源", style="CardTitle.TLabel").pack(anchor=W)
        template_list_frame = tb.Frame(left, style="Card.TFrame")
        template_list_frame.pack(fill=BOTH, expand=True, pady=7)
        self.import_template_list = tb.Treeview(
            template_list_frame,
            columns=("name", "created"),
            show=HEADINGS,
            height=12,
            selectmode="browse",
        )
        self.import_template_list.heading("name", text="模板名称")
        self.import_template_list.heading("created", text="保存时间")
        self.import_template_list.column("name", width=360, minwidth=240, stretch=False, anchor=W)
        self.import_template_list.column("created", width=145, minwidth=135, stretch=False, anchor=CENTER)
        template_scroll = tb.Scrollbar(template_list_frame, orient=VERTICAL, command=self.import_template_list.yview)
        template_scroll_x = tb.Scrollbar(template_list_frame, orient=HORIZONTAL, command=self.import_template_list.xview)
        self.import_template_list.configure(
            yscrollcommand=template_scroll.set,
            xscrollcommand=template_scroll_x.set,
        )
        self.import_template_list.grid(row=0, column=0, sticky=NSEW)
        template_scroll.grid(row=0, column=1, sticky=NS)
        template_scroll_x.grid(row=1, column=0, sticky=EW)
        template_list_frame.rowconfigure(0, weight=1)
        template_list_frame.columnconfigure(0, weight=1)
        self.import_template_list.bind("<<TreeviewSelect>>", self.choose_import_template)
        external = tb.Labelframe(left, text="外部来源", padding=8, style="Card.TLabelframe")
        external.pack(fill=X, pady=5)
        tb.Entry(external, textvariable=self.import_source).pack(fill=X)
        tb.Button(
            external,
            text="选择来源文件",
            style="Secondary.TButton",
            command=self.pick_import_source,
        ).pack(fill=X, pady=(6, 0))

        action_bar = tb.Labelframe(workspace, text="导入设置", padding=10, style="Card.TLabelframe")
        action_bar.pack(fill=X, pady=(0, 7))
        conflict_box = tb.Frame(action_bar)
        conflict_box.pack(side=LEFT)
        tb.Label(conflict_box, text="同名样式：").pack(side=LEFT)
        for text, value in (("覆盖", "overwrite"), ("跳过", "skip"), ("重命名", "rename")):
            tb.Radiobutton(conflict_box, text=text, value=value, variable=self.conflict_mode).pack(side=LEFT, padx=(0, 8))
        option_box = tb.Frame(action_bar)
        option_box.pack(side=LEFT, padx=10)
        tb.Checkbutton(option_box, text="清理未使用内容", variable=self.clean_unused).pack(side=LEFT)
        tb.Checkbutton(option_box, text="补全关联样式", variable=self.include_dependencies).pack(side=LEFT, padx=8)
        tb.Checkbutton(option_box, text="生成报告", variable=self.generate_report).pack(side=LEFT)
        tb.Button(action_bar, text="开始导入", width=11, style="Primary.TButton", command=self.start_import).pack(side=RIGHT)
        tb.Button(action_bar, text="导入预检", width=10, style="Secondary.TButton", command=self.preflight_import).pack(side=RIGHT, padx=6)

        preview_tabs = tb.Notebook(workspace)
        preview_tabs.pack(fill=BOTH, expand=True)
        source_preview_tab = tb.Frame(preview_tabs, padding=4)
        target_preview_tab = tb.Frame(preview_tabs, padding=4)
        preview_tabs.add(source_preview_tab, text="来源模板")
        preview_tabs.add(target_preview_tab, text="当前目标文档")
        self.import_preview = StylePreviewPanel(source_preview_tab, selectable=True, title="4、来源样式预览")
        self.import_preview.pack(fill=BOTH, expand=True)
        self.import_target_preview = StylePreviewPanel(target_preview_tab, selectable=False, title="4、目标文档样式预览")
        self.import_target_preview.pack(fill=BOTH, expand=True)

    def _build_library_tab(self):
        top = tb.Frame(self.library_tab, padding=12, style="Card.TFrame")
        top.pack(fill=X, pady=(0, 8))
        tb.Button(top, text="刷新", style="Ghost.TButton", command=self.refresh_templates).pack(side=RIGHT)
        tb.Button(top, text="打开目录", style="Secondary.TButton", command=lambda: open_folder(str(self.template_dir))).pack(side=RIGHT, padx=6)
        tb.Button(top, text="删除模板", style="Danger.TButton", command=self.delete_selected_template).pack(side=RIGHT, padx=(0, 6))
        tb.Button(top, text="编辑模板", style="Primary.TButton", command=self.edit_selected_template).pack(side=RIGHT)
        tb.Label(top, text="模板库管理", style="CardTitle.TLabel").pack(side=LEFT)
        tb.Label(
            self.library_tab,
            textvariable=self.template_dir_var,
            style="PageHint.TLabel",
            anchor=W,
        ).pack(fill=X, pady=(0, 4))
        paned = tb.Panedwindow(self.library_tab, orient=HORIZONTAL)
        paned.pack(fill=BOTH, expand=True, pady=8)
        left = tb.Frame(paned, width=620)
        right = tb.Frame(paned)
        paned.add(left, weight=2)
        paned.add(right, weight=3)
        self.library_tree = tb.Treeview(left, columns=("name", "created"), show=HEADINGS)
        for col, text, width in (("name", "模板名称", 420), ("created", "保存时间", 165)):
            self.library_tree.heading(col, text=text)
            self.library_tree.column(
                col,
                width=width,
                minwidth=240 if col == "name" else 145,
                stretch=False,
                anchor=W if col == "name" else CENTER,
            )
        library_scroll = tb.Scrollbar(left, orient=VERTICAL, command=self.library_tree.yview)
        library_scroll_x = tb.Scrollbar(left, orient=HORIZONTAL, command=self.library_tree.xview)
        self.library_tree.configure(
            yscrollcommand=library_scroll.set,
            xscrollcommand=library_scroll_x.set,
        )
        self.library_tree.grid(row=0, column=0, sticky=NSEW)
        library_scroll.grid(row=0, column=1, sticky=NS)
        library_scroll_x.grid(row=1, column=0, sticky=EW)
        left.rowconfigure(0, weight=1)
        left.columnconfigure(0, weight=1)
        self.library_tree.bind("<<TreeviewSelect>>", lambda _e: self.preview_selected_library_template())
        self.library_preview = StylePreviewPanel(right, selectable=False, title="模板样式与排版预览")
        self.library_preview.pack(fill=BOTH, expand=True)

    # paths and state
    def open_user_guide(self):
        guide = app_dir() / USER_GUIDE_FILE
        if guide.exists():
            os.startfile(guide)
        else:
            messagebox.showerror(APP_NAME, f"找不到使用说明：\n{guide}")

    def open_settings(self):
        if self.settings_window is not None and self.settings_window.winfo_exists():
            self.settings_window.deiconify()
            self.settings_window.lift()
            self.settings_window.focus_force()
            return
        window = tb.Toplevel(self)
        self.settings_window = window
        window.title("设置")
        window.geometry("720x410")
        window.resizable(False, False)
        window.transient(self)
        window.protocol("WM_DELETE_WINDOW", self.close_settings)
        window.configure(background=COLORS["background"])
        body = tb.Frame(window, padding=22, style="Office.TFrame")
        body.pack(fill=BOTH, expand=True)
        tb.Label(body, text="PREFERENCES", style="Eyebrow.TLabel").pack(anchor=W)
        tb.Label(body, text="设置", style="PageTitle.TLabel").pack(anchor=W, pady=(2, 0))
        tb.Label(body, text="管理输出位置、模板库和文件夹导入方式", style="PageHint.TLabel").pack(anchor=W, pady=(2, 0))

        output = tb.Labelframe(body, text="输出位置", padding=10, style="Card.TLabelframe")
        output.pack(fill=X, pady=(16, 8))
        tb.Label(output, textvariable=self.output_dir, bootstyle=INFO).pack(side=LEFT, fill=X, expand=True)
        self.output_folder_button = tb.Button(output, text="选择目录", command=self.change_output_dir)
        self.output_folder_button.pack(side=RIGHT)
        self.output_source_button = tb.Button(output, text="跟随源文件", command=self.reset_output_dir)
        self.output_source_button.pack(side=RIGHT, padx=6)
        self.refresh_output_mode_buttons()

        library = tb.Labelframe(body, text="模板库存放位置", padding=10, style="Card.TLabelframe")
        library.pack(fill=X, pady=8)
        tb.Label(library, textvariable=self.template_dir_var, bootstyle=INFO).pack(side=LEFT, fill=X, expand=True)
        tb.Button(library, text="选择目录", style="Secondary.TButton", command=self.change_template_dir).pack(side=RIGHT)

        options = tb.Labelframe(body, text="文件夹导入", padding=10, style="Card.TLabelframe")
        options.pack(fill=X, pady=8)
        tb.Checkbutton(
            options,
            text="添加文件夹时包含子目录",
            variable=self.recursive_folder,
            command=self.save_config,
        ).pack(anchor=W)
        tb.Button(body, text="完成", style="Primary.TButton", command=self.close_settings).pack(anchor=E, pady=(8, 0))
        window.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() - window.winfo_width()) // 2
        y = self.winfo_y() + (self.winfo_height() - window.winfo_height()) // 2
        window.geometry(f"+{x}+{y}")
        window.lift()

    def close_settings(self):
        if self.settings_window is not None:
            try:
                if self.settings_window.winfo_exists():
                    self.settings_window.destroy()
            finally:
                self.settings_window = None

    def refresh_output_mode_buttons(self):
        if not hasattr(self, "output_source_button"):
            return
        source_active = self.output_mode.get() == "source"
        self.output_source_button.configure(
            bootstyle="primary" if source_active else "outline-secondary"
        )
        self.output_folder_button.configure(
            bootstyle="primary" if not source_active else "outline-secondary"
        )

    def change_output_dir(self):
        folder = filedialog.askdirectory(title="选择输出目录")
        if folder:
            self.output_dir.set(folder)
            self.output_mode.set("folder")
            self.refresh_output_mode_buttons()
            self.save_config()

    def reset_output_dir(self):
        self.output_dir.set("跟随源文件目录")
        self.output_mode.set("source")
        self.refresh_output_mode_buttons()
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
        self.refresh_shared_document_lists()
        self.refresh_converter_file_list()
        self.refresh_templates()

    def refresh_file_list(self, widget, values):
        if not widget:
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
                self.library_tree.insert("", END, iid=str(idx), values=(item.get("name"), item.get("created_at")))
        if hasattr(self, "import_template_list"):
            self.import_template_list.delete(*self.import_template_list.get_children())
            for idx, item in enumerate(items):
                self.import_template_list.insert(
                    "",
                    END,
                    iid=str(idx),
                    values=(item.get("name"), item.get("created_at")),
                )
        current_source = self.import_source.get().strip()
        if current_source and not Path(current_source).is_file():
            self.import_source.set("")
            if hasattr(self, "import_preview"):
                self.import_preview.set_data([])

    def add_files(self, target, widget, filetypes):
        files = filedialog.askopenfilenames(filetypes=filetypes)
        if files:
            self.add_shared_documents(files, preferred_source=files[0])

    def add_folder_files(self, target, widget):
        folder = filedialog.askdirectory()
        if not folder:
            return
        pattern = "**/*" if self.recursive_folder.get() else "*"
        files = []
        for path in Path(folder).glob(pattern):
            if path.is_file() and path.suffix.lower() in {".doc", ".docx"}:
                files.append(str(path))
        self.add_shared_documents(files, preferred_source=files[0] if files else None)

    def add_shared_documents(self, files, preferred_source=None):
        previous = set(self.document_files)
        merged = merge_document_paths(self.document_files, files)
        self.document_files[:] = merged
        self.refresh_shared_document_lists()
        selected = str(preferred_source) if preferred_source else next(
            (path for path in self.document_files if path not in previous),
            self.document_files[0] if self.document_files else "",
        )
        if selected and selected in self.document_files:
            index = self.document_files.index(selected)
            self.document_list.selection_clear(0, END)
            self.document_list.selection_set(index)
            self.document_list.activate(index)
            self.document_list.see(index)
            self.on_document_select()
        self.prefetch_documents(path for path in self.document_files if path != selected)
        self.save_config()

    def refresh_shared_document_lists(self):
        self.document_count_text.set(f"{len(self.document_files)} 个文档")
        if hasattr(self, "document_list"):
            selection = self.document_list.curselection()
            selected_path = self.document_files[selection[0]] if selection and selection[0] < len(self.document_files) else None
            self.refresh_file_list(self.document_list, self.document_files)
            if selected_path in self.document_files:
                index = self.document_files.index(selected_path)
                self.document_list.selection_set(index)

    def pick_export_source(self):
        path = filedialog.askopenfilename(filetypes=[("Word 文件", "*.doc *.docx")])
        if path:
            self.add_shared_documents([path], preferred_source=path)

    def pick_import_source(self):
        path = filedialog.askopenfilename(filetypes=[("Word/模板文件", "*.doc *.docx *.dotx")])
        if path:
            self.import_source.set(path)
            self.preview_import_source()
            self.save_config()

    # format conversion
    def converter_kind_label(self, path):
        suffix = Path(path).suffix.lower()
        if suffix == ".doc":
            return ".doc → .docx"
        if suffix == ".xls":
            return ".xls → .xlsx"
        return ""

    def refresh_converter_file_list(self):
        self.converter_count_text.set(f"{len(self.converter_files)} 个待转换文件")
        if not hasattr(self, "converter_tree"):
            return
        selected = {
            self.converter_files[int(iid)]
            for iid in self.converter_tree.selection()
            if iid.isdigit() and int(iid) < len(self.converter_files)
        }
        self.converter_tree.delete(*self.converter_tree.get_children())
        for index, path in enumerate(self.converter_files):
            source = Path(path)
            self.converter_tree.insert(
                "",
                END,
                iid=str(index),
                values=(source.name, self.converter_kind_label(path), str(source.parent)),
            )
        for index, path in enumerate(self.converter_files):
            if path in selected:
                self.converter_tree.selection_add(str(index))

    def add_converter_paths(self, entries):
        files, ignored = collect_converter_paths(entries)
        previous_count = len(self.converter_files)
        self.converter_files[:] = merge_converter_paths(self.converter_files, files)
        added = len(self.converter_files) - previous_count
        self.refresh_converter_file_list()
        if added:
            self.status.set(f"已添加 {added} 个待转换文件。")
        if ignored:
            messagebox.showinfo(
                APP_NAME,
                f"已忽略 {ignored} 个不支持的项目。\n\n这里只接收旧版 .doc 和 .xls 文件。",
                parent=self,
            )
        if not added and not ignored:
            messagebox.showinfo(APP_NAME, "没有找到可转换的 .doc 或 .xls 文件。", parent=self)

    def pick_converter_files(self):
        files = filedialog.askopenfilenames(filetypes=[("旧版 Office 文件", "*.doc *.xls")])
        if files:
            self.add_converter_paths(files)

    def pick_converter_folder(self):
        folder = filedialog.askdirectory(title="选择包含旧版 Office 文件的文件夹")
        if folder:
            self.add_converter_paths([folder])

    def remove_selected_converter_files(self):
        selection = self.converter_tree.selection() if hasattr(self, "converter_tree") else ()
        if not selection:
            messagebox.showinfo(APP_NAME, "请先在列表中选择要移除的文件。", parent=self)
            return
        indexes = sorted((int(iid) for iid in selection if iid.isdigit()), reverse=True)
        for index in indexes:
            if index < len(self.converter_files):
                self.converter_files.pop(index)
        self.refresh_converter_file_list()
        self.status.set("已从待转换列表移除选中文件。")

    def clear_converter_files(self):
        self.converter_files.clear()
        self.refresh_converter_file_list()
        self.set_converter_result("尚未开始转换。", "")
        self.converter_open_button.configure(state=DISABLED)
        self.status.set("已清空格式转换列表。")

    def pick_converter_output_dir(self):
        folder = filedialog.askdirectory(title="选择转换输出目录")
        if folder:
            self.converter_output_dir.set(folder)
            self.converter_output_mode.set("folder")
            self.save_config()

    def sync_converter_engine_checks(self):
        engine = self.converter_engine.get()
        self.converter_office_selected.set(engine == "office")
        self.converter_wps_selected.set(engine == "wps")

    def engine_any_available(self, engine):
        state = self.converter_engine_state.get(engine, {})
        return any(value is True for value in state.values())

    def engine_required_missing(self, engine, required):
        state = self.converter_engine_state.get(engine, {})
        return [kind for kind in required if state.get(kind) is False]

    def converter_required_kinds(self):
        required = set()
        for path in self.converter_files:
            suffix = Path(path).suffix.lower()
            if suffix == ".doc":
                required.add("word")
            elif suffix == ".xls":
                required.add("excel")
        return required

    def select_converter_engine(self, engine, notify=True):
        other = "wps" if engine == "office" else "office"
        if self.converter_engine_checked and not self.engine_any_available(engine):
            self.converter_engine.set(other if self.engine_any_available(other) else engine)
            self.sync_converter_engine_checks()
            if notify:
                messagebox.showinfo(
                    APP_NAME,
                    f"未检测到 {converter.ENGINE_LABELS[engine]} 的转换组件。"
                    + (
                        f"\n\n可以改用 {converter.ENGINE_LABELS[other]}。"
                        if self.engine_any_available(other)
                        else "\n\n请先确认本机已安装 Microsoft Office 或 WPS。"
                    ),
                    parent=self,
                )
            self.save_config()
            return
        self.converter_engine.set(engine)
        self.sync_converter_engine_checks()
        self.save_config()

    def ensure_converter_engine_check(self, force=False):
        if self.converter_engine_checking:
            return
        if self.converter_engine_checked and not force:
            return
        self.converter_engine_checking = True
        self.converter_engine_text.set("正在检测 Microsoft Office 和 WPS，请稍候...")
        self.status.set("正在检测格式转换引擎...")

        def worker():
            result = {
                "office": {"word": False, "excel": False},
                "wps": {"word": False, "excel": False},
            }
            if not converter._HAS_WIN32COM:
                return result
            if converter.pythoncom:
                converter.pythoncom.CoInitialize()
            try:
                for engine in ("office", "wps"):
                    for kind in ("word", "excel"):
                        result[engine][kind] = converter._probe_engine_available(engine, kind)
            finally:
                if converter.pythoncom:
                    converter.pythoncom.CoUninitialize()
            return result

        def runner():
            try:
                result = worker()
                self.after(0, lambda: self.finish_converter_engine_check(result))
            except Exception as exc:
                self.after(0, lambda: self.finish_converter_engine_check(None, str(exc)))

        threading.Thread(target=runner, daemon=True).start()

    def finish_converter_engine_check(self, result, error=None):
        self.converter_engine_checking = False
        self.converter_engine_checked = True
        if result:
            self.converter_engine_state = result
        if error:
            self.converter_engine_text.set(f"引擎检测未完成：{error}")
            self.status.set("格式转换引擎检测失败。")
            return
        self.converter_engine_text.set(self.converter_engine_summary())
        if not self.engine_any_available("office") and self.engine_any_available("wps"):
            self.converter_engine.set("wps")
            self.sync_converter_engine_checks()
        elif not self.engine_any_available("office") and not self.engine_any_available("wps"):
            self.converter_engine.set("office")
            self.sync_converter_engine_checks()
        self.status.set("格式转换引擎检测完成。")
        self.save_config()

    def converter_engine_summary(self):
        def label(engine):
            state = self.converter_engine_state[engine]
            word = "Word 可用" if state["word"] else "Word 未检测到"
            excel = "Excel 可用" if state["excel"] else "Excel 未检测到"
            return f"{converter.ENGINE_LABELS[engine]}：{word}，{excel}"
        if not converter._HAS_WIN32COM:
            return "当前环境缺少 pywin32，暂时无法调用 Office 或 WPS。"
        return f"{label('office')}\n{label('wps')}"

    def validate_converter_engine_for_current_files(self):
        if not self.converter_engine_checked:
            self.ensure_converter_engine_check()
            messagebox.showinfo(APP_NAME, "正在检测本机可用的转换引擎，请稍后再开始转换。", parent=self)
            return False
        required = self.converter_required_kinds()
        engine = self.converter_engine.get()
        missing = self.engine_required_missing(engine, required)
        if not missing:
            return True
        other = "wps" if engine == "office" else "office"
        other_missing = self.engine_required_missing(other, required)
        missing_labels = "、".join(converter.KIND_LABELS[item] for item in missing)
        if self.engine_any_available(other) and not other_missing:
            self.converter_engine.set(other)
            self.sync_converter_engine_checks()
            self.save_config()
            messagebox.showinfo(
                APP_NAME,
                f"未检测到 {converter.ENGINE_LABELS[engine]} 的 {missing_labels} 转换组件。\n\n"
                f"已为你切换到 {converter.ENGINE_LABELS[other]}，请再次点击“开始转换”。",
                parent=self,
            )
        else:
            messagebox.showerror(
                APP_NAME,
                f"未检测到可处理当前文件的转换组件：{missing_labels}。\n\n"
                "请确认 Microsoft Office 或 WPS 已安装，并关闭可能占用文件的窗口后再试。",
                parent=self,
            )
        return False

    def converter_output_folder(self):
        if self.converter_output_mode.get() != "folder":
            return None
        folder = self.converter_output_dir.get().strip()
        if not folder:
            messagebox.showinfo(APP_NAME, "请先选择输出文件夹，或改为输出到原文件同目录。", parent=self)
            return False
        try:
            Path(folder).mkdir(parents=True, exist_ok=True)
        except Exception as exc:
            messagebox.showerror(APP_NAME, f"输出文件夹无法使用：\n{exc}", parent=self)
            return False
        return folder

    def set_converter_result(self, summary, detail):
        self.converter_result_text.set(summary)
        if hasattr(self, "converter_result_detail"):
            self.converter_result_detail.configure(state=NORMAL)
            self.converter_result_detail.delete("1.0", END)
            self.converter_result_detail.insert("1.0", detail or "暂无详细结果。")
            self.converter_result_detail.configure(state=DISABLED)

    def start_converter(self):
        if self.converter_running:
            return
        if not self.converter_files:
            messagebox.showinfo(APP_NAME, "请先添加要转换的 .doc 或 .xls 文件。", parent=self)
            return
        if not self.validate_converter_engine_for_current_files():
            return
        output_dir = self.converter_output_folder()
        if output_dir is False:
            return

        files = list(self.converter_files)
        engine = self.converter_engine.get()
        self.converter_running = True
        self.converter_start_button.configure(state=DISABLED)
        self.converter_open_button.configure(state=DISABLED)
        self.set_converter_result("正在转换，请稍候...", "")

        def work():
            try:
                return {"ok": True, "result": converter.convert_files(files, output_dir=output_dir, engine=engine)}
            except Exception as exc:
                return {"ok": False, "error": str(exc)}

        self.run_task(work, self.finish_converter)

    def finish_converter(self, payload):
        self.converter_running = False
        self.converter_start_button.configure(state=NORMAL)
        if not payload.get("ok"):
            self.set_converter_result("转换未完成。", payload.get("error", "未知错误"))
            messagebox.showerror(APP_NAME, f"转换未完成：\n{payload.get('error', '未知错误')}", parent=self)
            return
        result = payload["result"]
        converted = result.get("converted", [])
        failed = result.get("failed", [])
        outputs = [item.get("output") for item in converted if item.get("output")]
        self.result_paths = outputs
        if outputs:
            self.converter_open_button.configure(state=NORMAL)
        summary = f"转换完成：成功 {len(converted)} 个，失败 {len(failed)} 个。"
        details = []
        if converted:
            details.append("成功：")
            details.extend(f"- {Path(item['source']).name} → {item['output']}" for item in converted)
        if failed:
            details.append("\n失败：")
            details.extend(f"- {Path(item.get('source', '')).name}: {item.get('error', '未知错误')}" for item in failed)
        self.set_converter_result(summary, "\n".join(details))
        self.status.set(summary)
        FinishDialog(self, "格式转换完成", summary, outputs)

    def open_converter_result_location(self):
        if self.result_paths:
            open_folder(self.result_paths[0], select_file=True)

    def on_tab_changed(self, _event=None):
        tab = self.notebook.index(self.notebook.select())
        titles = (
            ("样式清理", "检查并整理 Word/WPS 文档中的样式与编号"),
            ("提取模板", "从现有文档提取可重复使用的轻量 Word 模板"),
            ("格式转换", "将旧版 Office 文件转换为更稳定的新格式"),
            ("模板库", "集中预览、管理和编辑已保存的模板"),
        )
        self.page_title.set(titles[tab][0])
        self.page_hint.set(titles[tab][1])
        for button_index, button in enumerate(self.nav_buttons):
            button.configure(style="NavActive.TButton" if button_index == tab else "Nav.TButton")
        if hasattr(self, "documents_panel"):
            if tab in {2, 3}:
                self.documents_panel.pack_forget()
            elif not self.documents_panel.winfo_manager():
                self.documents_panel.pack(fill=X, pady=(10, 8), before=self.notebook)
        self.refresh_shared_document_lists()
        if tab == 2:
            self.ensure_converter_engine_check()
        if tab == 3:
            self.refresh_templates()
        if tab in {0, 1} and hasattr(self, "import_target_preview"):
            self.on_document_select()
        self.save_config()

    def add_clean_files(self):
        self.add_files(self.clean_files, self.document_list, [("Word 文件", "*.doc *.docx")])

    def add_clean_folder(self):
        self.add_folder_files(self.clean_files, self.document_list)

    def clear_clean_files(self):
        self.document_files.clear()
        self.export_source.set("")
        self.preview_cache.clear()
        self.clean_preview.set_data([])
        self.export_preview.set_data([])
        self.import_target_preview.set_data([])
        self.refresh_shared_document_lists()
        self.save_config()

    def remove_selected_document(self):
        selection = self.document_list.curselection()
        if not selection:
            messagebox.showinfo(APP_NAME, "请先在文档列表中选择要移除的文档。")
            return
        index = selection[0]
        removed = self.document_files.pop(index)
        self.preview_cache.pop(os.path.normcase(os.path.abspath(removed)), None)
        self.refresh_shared_document_lists()
        if self.document_files:
            next_index = min(index, len(self.document_files) - 1)
            self.document_list.selection_set(next_index)
            self.document_list.activate(next_index)
            self.on_document_select()
        else:
            self.export_source.set("")
            self.clean_preview.set_data([])
            self.export_preview.set_data([])
            self.import_target_preview.set_data([])
        self.save_config()

    def add_import_targets(self):
        self.add_files(self.import_targets, self.document_list, [("Word 文件", "*.doc *.docx")])

    def add_import_folder(self):
        self.add_folder_files(self.import_targets, self.document_list)

    def clear_import_targets(self):
        self.clear_clean_files()

    # previews
    def selected_document(self):
        selection = self.document_list.curselection()
        if selection and selection[0] < len(self.document_files):
            return self.document_files[selection[0]]
        return self.document_files[0] if self.document_files else ""

    def preview_signature(self, path):
        try:
            stat = os.stat(path)
            return stat.st_mtime_ns, stat.st_size
        except OSError:
            return None

    def request_preview(self, path, callback=None):
        if not path:
            return
        key = os.path.normcase(os.path.abspath(path))
        signature = self.preview_signature(path)
        cached = self.preview_cache.get(key)
        if cached and cached["signature"] == signature:
            if callback:
                callback(cached["result"])
            return
        if callback:
            self.preview_waiters.setdefault(key, []).append(callback)
        if key in self.preview_loading:
            return
        self.preview_loading.add(key)

        def worker():
            self.preview_results.put((key, signature, inspect_document(path)))

        self.preview_executor.submit(worker)

    def process_preview_results(self):
        while True:
            try:
                key, signature, result = self.preview_results.get_nowait()
            except queue.Empty:
                break
            self.preview_loading.discard(key)
            self.preview_cache[key] = {"signature": signature, "result": result}
            callbacks = self.preview_waiters.pop(key, [])
            for waiter in callbacks:
                waiter(result)
        if not self.closing:
            self.after(25, self.process_preview_results)

    def prefetch_documents(self, paths):
        for path in paths:
            self.request_preview(path)

    def on_document_select(self, _event=None):
        path = self.selected_document()
        if not path:
            return
        self.export_source.set(path)
        self.status.set(f"正在准备预览：{Path(path).name}")
        tab = self.notebook.index(self.notebook.select())
        if tab == 0:
            self.clean_preview.set_loading(True)
        elif tab == 1:
            self.export_preview.set_loading(True)
        self.request_preview(path, lambda result, current=path: self.finish_document_preview(current, result))

    def finish_document_preview(self, path, result):
        if path != self.selected_document():
            return
        if not result.get("success"):
            tab = self.notebook.index(self.notebook.select())
            panel = {0: self.clean_preview, 1: self.export_preview}.get(tab)
            if panel:
                panel.set_loading(False)
            self.status.set(f"预览失败：{result.get('error', '读取失败')}")
            return
        styles = result.get("styles", [])
        tab = self.notebook.index(self.notebook.select())
        if tab == 0:
            self.clean_preview.set_data(styles)
        elif tab == 1:
            self.export_preview.set_data(styles)
        self.status.set(
            f"当前文档：{Path(path).name}；样式 {len(styles)} 个，编号 {len(result.get('numbering', []))} 组。"
        )

    def preview_path_into(self, path, panel):
        if not path:
            messagebox.showinfo(APP_NAME, "请先选择文件。")
            return
        panel.set_loading(True)
        self.request_preview(path, lambda result: self.finish_preview(result, panel))

    def finish_preview(self, result, panel):
        if not result.get("success"):
            panel.set_loading(False)
            messagebox.showerror(APP_NAME, result.get("error", "读取失败"))
            return
        panel.set_data(result.get("styles", []))
        self.status.set(f"预览完成：样式 {len(result.get('styles', []))} 个，编号 {len(result.get('numbering', []))} 组。")

    def preview_export(self):
        self.preview_path_into(self.export_source.get(), self.export_preview)

    def preview_import_source(self):
        self.preview_path_into(self.import_source.get(), self.import_preview)

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
        if item and self.ensure_template_exists(item):
            self.preview_path_into(item.get("path"), self.library_preview)

    # actions
    def clean_options(self):
        return {
            "mode": self.clean_mode.get(),
            "unused_styles": self.clean_unused_styles.get(),
            "hide_unused_builtin": self.clean_mode.get() == "safe",
            "unused_numbering": self.clean_unused_numbering.get(),
            "unused_multilevel": self.clean_unused_multilevel.get(),
            "deduplicate_numbering": self.clean_duplicate_numbering.get(),
            "repair_links": self.clean_repair_links.get(),
        }

    def start_clean(self):
        selected = self.selected_document()
        if not selected:
            messagebox.showinfo(APP_NAME, "请先添加文件。")
            return
        self.start_clean_files([selected])

    def start_batch_clean(self):
        if not self.clean_files:
            messagebox.showinfo(APP_NAME, "请先添加文件。")
            return
        self.start_clean_files(list(self.clean_files))

    def start_clean_files(self, files):
        options = self.clean_options()
        if not any(value for key, value in options.items() if key != "mode"):
            messagebox.showinfo(APP_NAME, "请至少选择一个清理选项。")
            return
        if options["mode"] == "deep" and not messagebox.askyesno(
            "深度清理确认",
            "深度清理会删除更多未使用内置样式，并解除不影响现有排版的后续/链接关系。\n\n"
            "程序仍只生成新文件，不覆盖源文档，并会在完成后自动复检。\n\n是否继续？",
            parent=self,
        ):
            return

        def work():
            outputs = []
            results = []
            for file in files:
                out = self.output_path(file, "_样式清理版")
                result = auto_clean_docx(file, out, options=options)
                result["source"] = file
                results.append(result)
                if result.get("success"):
                    outputs.append(out)
            return results, outputs
        self.run_task(work, self.finish_clean)

    def finish_clean(self, result):
        results, outputs = result
        successful = [item for item in results if item.get("success")]
        failed = [item for item in results if not item.get("success")]
        ok = len(successful)
        attempted = len(results)
        styles_removed = sum(item.get("removed_styles", 0) for item in successful)
        styles_hidden = sum(item.get("hidden_builtin", 0) for item in successful)
        numbering_removed = sum(
            item.get("removed_numbering", 0) + item.get("deduplicated_numbering", 0)
            for item in successful
        )
        links_repaired = sum(
            item.get("pruned_links", 0) + item.get("repaired_links", 0)
            for item in successful
        )
        effects_removed = sum(item.get("effects_removed", 0) for item in successful)
        self.result_paths = outputs
        self.status.set(
            f"清理完成：{ok}/{attempted}，删除样式 {styles_removed}，"
            f"隐藏内置样式 {styles_hidden}，清理列表/编号 {numbering_removed}，"
            f"修复关联 {links_repaired}。"
        )
        message = (
            f"已生成并通过复检 {ok} 个清理版文档。\n"
            f"删除样式 {styles_removed} 个，隐藏内置样式 {styles_hidden} 个，"
            f"清理列表/编号 {numbering_removed} 项，修复或解除关联 {links_repaired} 项，"
            f"同步清理辅助样式表 {effects_removed} 项。"
        )
        if failed:
            details = "\n".join(
                f"- {Path(item.get('source', '')).name}: {item.get('error', '未知错误')}"
                for item in failed[:5]
            )
            message += f"\n\n有 {len(failed)} 个文件未通过复检，未保留输出：\n{details}"
        FinishDialog(self, "清理完成", message, outputs)

    def start_export(self):
        source = self.export_source.get().strip()
        if not source:
            messagebox.showinfo(APP_NAME, "请先选择源文档。")
            return
        formats = ["dotx", "library"]
        selected = self.export_preview.get_selected_ids()
        def work():
            return export_style_template(
                source,
                str(self.template_dir),
                self.template_name.get().strip() or None,
                formats,
                str(self.template_dir),
                selected or None,
                self.include_dependencies.get(),
                self.export_auto_clean.get(),
            )
        self.run_task(work, self.finish_export)

    def start_batch_export(self):
        if not self.document_files:
            messagebox.showinfo(APP_NAME, "请先添加至少一个源文档。")
            return
        formats = ["dotx", "library"]

        files = list(self.document_files)

        def work():
            results = []
            for source in files:
                results.append(
                    export_style_template(
                        source,
                        str(self.template_dir),
                        None,
                        formats,
                        str(self.template_dir),
                        None,
                        self.include_dependencies.get(),
                        self.export_auto_clean.get(),
                    )
                )
            return results

        self.run_task(work, self.finish_batch_export)

    def finish_batch_export(self, results):
        ok = [item for item in results if item.get("success")]
        self.result_paths = list(
            dict.fromkeys(
                path
                for item in ok
                for path in item.get("outputs", {}).values()
            )
        )
        self.refresh_templates()
        self.status.set(f"批量提取完成：成功 {len(ok)}/{len(results)} 个文档。")
        FinishDialog(
            self,
            "批量提取完成",
            f"已从 {len(ok)} 个文档生成模板，保存到模板库目录。",
            self.result_paths,
        )

    def finish_export(self, result):
        if not result.get("success"):
            messagebox.showerror(APP_NAME, result.get("error", "导出失败"))
            return
        self.result_paths = list(dict.fromkeys(result.get("outputs", {}).values()))
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
        index = int(selection[0])
        return self.template_items[index] if index < len(self.template_items) else None

    def ensure_template_exists(self, item):
        path = Path(item.get("path", ""))
        if path.is_file():
            return True
        missing_path = str(path)
        self.refresh_templates()
        if hasattr(self, "library_preview"):
            self.library_preview.set_data([])
        self.status.set(f"模板文件已不存在，已从模板库列表移除：{path.name}")
        messagebox.showinfo(
            APP_NAME,
            f"模板文件已被删除或移动，失效记录已自动清理：\n{missing_path}",
        )
        return False

    def choose_import_template(self, event=None):
        selection = self.import_template_list.selection()
        if not selection:
            return
        index = int(selection[0])
        if index >= len(self.template_items):
            return
        item = self.template_items[index]
        if not self.ensure_template_exists(item):
            return
        self.import_source.set(item.get("path"))
        self.preview_import_source()
        self.save_config()

    def edit_selected_template(self):
        item = self.selected_template_item()
        if not item:
            messagebox.showinfo(APP_NAME, "请先选择模板。")
            return
        if not self.ensure_template_exists(item):
            return
        TemplateEditor(self, item.get("path"), self.template_dir, on_saved=lambda _p: self.refresh_templates())

    def delete_selected_template(self):
        item = self.selected_template_item()
        if not item:
            messagebox.showinfo(APP_NAME, "请先选择要删除的模板。")
            return
        path = Path(item.get("path", ""))
        if not self.ensure_template_exists(item):
            return
        if not messagebox.askyesno(
            "删除模板",
            f"将永久删除模板文件：\n\n{path.name}\n\n"
            "此操作不会删除由该模板生成的文档，且无法撤销。是否继续？",
            parent=self,
        ):
            return
        try:
            delete_template_from_library(self.template_dir, path)
        except Exception as exc:
            messagebox.showerror(APP_NAME, f"删除模板失败：\n{exc}", parent=self)
            return
        if os.path.normcase(os.path.abspath(self.import_source.get().strip())) == os.path.normcase(str(path.resolve())):
            self.import_source.set("")
            self.import_preview.set_data([])
        self.library_preview.set_data([])
        self.refresh_templates()
        self.status.set(f"已删除模板：{path.name}")

    # drag/drop
    def on_drop(self, event):
        tab = self.notebook.index(self.notebook.select())
        dropped = list(self.tk.splitlist(event.data))
        if tab == 2:
            self.add_converter_paths(dropped)
            self.save_config()
            return
        files = [f for f in dropped if f.lower().endswith((".doc", ".docx", ".dotx"))]
        if not files:
            return
        if tab == 0:
            self.add_shared_documents(files, preferred_source=files[0])
        elif tab == 1:
            self.add_shared_documents(files, preferred_source=files[0])
        else:
            self.status.set("模板库页面不接收拖入文件，请使用“打开目录”管理模板。")
        self.save_config()

    # GitHub and updates
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
                root, payload = download_update(
                    release["download_url"],
                    expected_size=release.get("asset_size") or None,
                )
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
        self.closing = True
        self.close_settings()
        self.preview_executor.shutdown(wait=False, cancel_futures=True)
        self.save_config()
        self.destroy()

    def on_root_destroy(self, event):
        if event.widget is self:
            self.close_settings()


def main():
    app = StyleManagerApp()
    app.mainloop()
