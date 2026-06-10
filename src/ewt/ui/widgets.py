"""Reusable UI widgets."""

from __future__ import annotations

import tkinter as tk
from tkinter import filedialog

import ttkbootstrap as tb
from ttkbootstrap.constants import *

from src.ewt.utils.display import chinese_style_name
from src.ewt.ui.theme import COLORS, configure_text


class PathRow(tb.Frame):
    def __init__(self, parent, label, variable, filetypes=None, mode="open", command=None):
        super().__init__(parent)
        self.variable = variable
        self.filetypes = filetypes or []
        self.mode = mode
        self.command = command or self.pick
        tb.Label(self, text=label, width=11, anchor=W).pack(side=LEFT)
        tb.Entry(self, textvariable=variable).pack(side=LEFT, fill=X, expand=True, padx=8)
        tb.Button(self, text="选择", style="Secondary.TButton", command=self.command).pack(side=LEFT)

    def pick(self):
        if self.mode == "dir":
            path = filedialog.askdirectory()
        elif self.mode == "save":
            path = filedialog.asksaveasfilename(filetypes=self.filetypes)
        else:
            path = filedialog.askopenfilename(filetypes=self.filetypes)
        if path:
            self.variable.set(path)


class StylePreviewPanel(tb.Frame):
    """Selectable style inventory with property and formatted sample previews."""

    TYPE_NAMES = {"paragraph": "段落", "character": "字符", "table": "表格", "numbering": "编号"}

    @staticmethod
    def _points(value):
        if value in (None, ""):
            return "-"
        try:
            return f"{float(value) / 20:g} 磅"
        except (TypeError, ValueError):
            return str(value)

    @staticmethod
    def _size(value):
        if value in (None, ""):
            return ""
        return f"{value} 磅"

    @staticmethod
    def _line(value):
        labels = {"240": "单倍", "276": "1.15 倍", "300": "1.25 倍", "360": "1.5 倍", "480": "2 倍"}
        return labels.get(str(value), StylePreviewPanel._points(value))

    def __init__(self, parent, selectable=True, title="样式预览", show_details=True):
        super().__init__(parent, padding=8, style="Card.TFrame")
        self.selectable = selectable
        self.show_details = show_details
        self.styles = []
        self.style_by_id = {}
        self.selected_ids = set()
        self.filter_var = tk.StringVar(value="全部")
        self.show_hidden_builtin = tk.BooleanVar(value=False)
        self.search_var = tk.StringVar()

        head = tb.Frame(self, style="Card.TFrame")
        head.pack(fill=X)
        tb.Label(head, text=title, style="CardTitle.TLabel").pack(side=LEFT)
        self.loading_label = tb.Label(head, text="", bootstyle=WARNING)
        self.loading_label.pack(side=RIGHT)
        self.summary_label = tb.Label(self, text="", style="Muted.TLabel", anchor=W)
        self.summary_label.pack(fill=X, pady=(3, 6))

        tools = tb.Frame(self, style="Card.TFrame")
        tools.pack(fill=X, pady=(0, 6))
        filters = tb.Frame(tools, style="Card.TFrame")
        filters.pack(fill=X)
        for text in ("全部", "未使用", "可清理", "系统可隐藏", "使用中", "自定义", "带编号"):
            tb.Radiobutton(
                filters,
                text=text,
                value=text,
                variable=self.filter_var,
                bootstyle="secondary-toolbutton",
            ).pack(side=LEFT, padx=(0, 3))
        secondary_tools = tb.Frame(tools, style="Card.TFrame")
        secondary_tools.pack(fill=X, pady=(5, 0))
        tb.Checkbutton(
            secondary_tools,
            text="显示已隐藏系统样式",
            variable=self.show_hidden_builtin,
            bootstyle="round-toggle",
            command=self.refresh,
        ).pack(side=LEFT)
        self.filter_var.trace_add("write", lambda *_: self.refresh())
        if selectable:
            tb.Button(secondary_tools, text="全选", style="Ghost.TButton", command=self.select_all).pack(side=RIGHT)
            tb.Button(secondary_tools, text="只选已用", style="Ghost.TButton", command=self.select_used).pack(side=RIGHT)
            tb.Button(secondary_tools, text="取消选择", style="Ghost.TButton", command=self.select_none).pack(side=RIGHT)
        self.search_entry = tb.Entry(secondary_tools, textvariable=self.search_var, width=22)
        self.search_entry.pack(side=RIGHT, padx=(6, 12))
        tb.Label(secondary_tools, text="搜索样式", style="Muted.TLabel").pack(side=RIGHT)
        self.search_var.trace_add("write", lambda *_: self.refresh())

        paned = tb.Panedwindow(self, orient=VERTICAL)
        paned.pack(fill=BOTH, expand=True)
        tree_frame = tb.Frame(paned, style="Surface.TFrame")
        preview_frame = tb.Frame(paned, style="Surface.TFrame")
        paned.add(tree_frame, weight=3)
        if self.show_details:
            paned.add(preview_frame, weight=2)

        columns = ("name", "status", "type", "count", "font", "size", "number", "note")
        self.tree = tb.Treeview(tree_frame, columns=columns, show=HEADINGS, height=10, selectmode="browse")
        headings = {
            "name": "样式名称",
            "status": "状态",
            "type": "类型",
            "count": "使用次数",
            "font": "字体",
            "size": "字号",
            "number": "编号",
            "note": "备注",
        }
        widths = {
            "name": 175,
            "status": 68,
            "type": 58,
            "count": 72,
            "font": 105,
            "size": 48,
            "number": 125,
            "note": 330,
        }
        anchors = {
            "name": W,
            "status": CENTER,
            "type": CENTER,
            "count": CENTER,
            "font": W,
            "size": CENTER,
            "number": CENTER,
            "note": W,
        }
        for col in columns:
            self.tree.heading(col, text=headings[col], anchor=anchors[col])
            self.tree.column(
                col,
                width=widths[col],
                stretch=col in {"name", "font", "note"},
                anchor=anchors[col],
            )
        self.tree.tag_configure("unused", foreground=COLORS["danger"], background=COLORS["surface"])
        self.tree.tag_configure("builtin_unused", foreground=COLORS["warning"], background=COLORS["surface"])
        self.tree.tag_configure("used", foreground=COLORS["text"], background=COLORS["surface"])
        scroll_y = tb.Scrollbar(tree_frame, orient=VERTICAL, command=self.tree.yview)
        scroll_x = tb.Scrollbar(tree_frame, orient=HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        self.tree.grid(row=0, column=0, sticky=NSEW)
        scroll_y.grid(row=0, column=1, sticky=NS)
        scroll_x.grid(row=1, column=0, sticky=EW)
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)
        self.tree.bind("<<TreeviewSelect>>", self._show_selected)
        self.tree.bind("<Double-1>", self._toggle_selected)

        tabs = tb.Notebook(preview_frame)
        tabs.pack(fill=BOTH, expand=True, pady=(5, 0))
        attr_tab = tb.Frame(tabs, padding=6)
        sample_tab = tb.Frame(tabs, padding=6)
        tabs.add(attr_tab, text="属性")
        tabs.add(sample_tab, text="示例排版")
        self.attr_text = tk.Text(attr_tab, height=8, wrap="word", relief="flat")
        configure_text(self.attr_text, COLORS["surface_alt"])
        self.attr_text.pack(fill=BOTH, expand=True)
        self.sample = tk.Text(sample_tab, height=8, wrap="word", relief="flat", padx=16, pady=10)
        configure_text(self.sample)
        self.sample.pack(fill=BOTH, expand=True)

    def set_data(self, styles):
        self.set_loading(False)
        self.styles = list(styles or [])
        self.style_by_id = {item["style_id"]: item for item in self.styles}
        if self.selectable:
            self.selected_ids = {item["style_id"] for item in self.styles}
        self.refresh()

    def set_loading(self, loading=True, text="文档样式加载中..."):
        self.loading_label.configure(text=text if loading else "", bootstyle=WARNING if loading else SECONDARY)
        if loading:
            self.styles = []
            self.style_by_id = {}
            self.selected_ids.clear()
            if hasattr(self, "tree"):
                self.tree.delete(*self.tree.get_children())
            self.summary_label.configure(text="正在读取样式与编号")

    def _matches(self, item):
        query = self.search_var.get().strip().casefold()
        if query:
            searchable = " ".join(
                str(item.get(key, ""))
                for key in ("display_name", "name", "style_id")
            ).casefold()
            if query not in searchable:
                return False
        is_unused_builtin = item.get(
            "hideable",
            item.get("count", 0) == 0
            and item.get("builtin", False)
            and not item.get("dependency_required", False),
        )
        if is_unused_builtin and item.get("hidden", False) and not self.show_hidden_builtin.get():
            return False
        current = self.filter_var.get()
        if current == "使用中":
            return item.get("count", 0) > 0
        if current == "未使用":
            return item.get("count", 0) == 0
        if current == "可清理":
            return item.get(
                "cleanable",
                item.get("count", 0) == 0
                and not item.get("builtin", False)
                and not item.get("dependency_required", False),
            )
        if current == "系统可隐藏":
            return is_unused_builtin
        if current == "自定义":
            return not item.get("builtin", False)
        if current == "带编号":
            return bool(item.get("num_id"))
        if current == "段落":
            return item.get("type") == "paragraph"
        if current == "字符":
            return item.get("type") == "character"
        if current == "表格":
            return item.get("type") == "table"
        return True

    def refresh(self):
        if not hasattr(self, "tree"):
            return
        self.tree.delete(*self.tree.get_children())
        visible = 0
        if not self.styles:
            self.summary_label.configure(text="尚未加载样式。请先添加或选择一个 Word 文档。")
            self.attr_text.configure(state=NORMAL)
            self.attr_text.delete("1.0", END)
            self.attr_text.insert("1.0", "选择文档后，这里会显示样式属性、使用次数和编号关系。")
            self.attr_text.configure(state=DISABLED)
            self.sample.configure(state=NORMAL)
            self.sample.delete("1.0", END)
            self.sample.insert("1.0", "暂无排版预览")
            self.sample.configure(state=DISABLED)
            return
        for item in self.styles:
            if not self._matches(item):
                continue
            visible += 1
            sid = item["style_id"]
            is_used = item.get("count", 0) > 0
            is_builtin = bool(item.get("builtin"))
            is_required = bool(item.get("dependency_required"))
            is_cleanable = item.get(
                "cleanable",
                not is_used and not is_builtin and not is_required,
            )
            picked = "✓" if sid in self.selected_ids else ""
            localized_name = item.get("display_name") or item.get("name")
            display_name = f"✓  {localized_name}" if self.selectable and picked else localized_name
            number = (
                f"定义 {item.get('num_id')}·第 {int(item.get('level') or 0) + 1} 级"
                if item.get("num_id")
                else ""
            )
            if is_required:
                note = "核心样式，或被正在使用的样式通过基于、后续、链接关系引用，清理时必须保留"
            elif is_builtin and not is_used:
                note = (
                    "已隐藏的系统样式；启用后 WPS/Word 仍可自动重新显示"
                    if item.get("hidden")
                    else "系统样式无法永久删除；可隐藏，使用时由 WPS/Word 自动重建"
                )
            else:
                note = ""
            tag = "used" if is_used or is_required else ("builtin_unused" if is_builtin else "unused")
            status = (
                "● 使用中"
                if is_used
                else "● 必要保留"
                if is_required
                else "● 可清理"
                if is_cleanable
                else "● 可隐藏"
            )
            self.tree.insert(
                "",
                END,
                iid=sid,
                values=(
                    display_name,
                    status,
                    self.TYPE_NAMES.get(item.get("type"), item.get("type")),
                    item.get("count") if is_used else "0",
                    item.get("font"),
                    self._size(item.get("size")),
                    number,
                    note,
                ),
                tags=(tag,),
            )
        unused = sum(1 for item in self.styles if item.get("count", 0) == 0)
        cleanable = sum(
            1 for item in self.styles
            if item.get(
                "cleanable",
                item.get("count", 0) == 0
                and not item.get("builtin", False)
                and not item.get("dependency_required", False),
            )
        )
        rebuildable = sum(
            1 for item in self.styles
            if item.get(
                "hideable",
                item.get("count", 0) == 0
                and item.get("builtin", False)
                and not item.get("dependency_required", False),
            )
        )
        required = sum(1 for item in self.styles if item.get("dependency_required", False))
        self.summary_label.configure(
            text=(
                f"共 {len(self.styles)}，未使用 {unused}，"
                f"可清理 {cleanable}，必要保留 {required}，"
                f"系统可隐藏 {rebuildable}（当前显示 {visible}）"
            )
        )

    def _toggle_selected(self, event=None):
        if not self.selectable:
            return
        sid = self.tree.identify_row(event.y) if event else ""
        if not sid:
            return
        if sid in self.selected_ids:
            self.selected_ids.remove(sid)
        else:
            self.selected_ids.add(sid)
        self.refresh()
        if self.tree.exists(sid):
            self.tree.selection_set(sid)

    def select_all(self):
        self.selected_ids = {item["style_id"] for item in self.styles}
        self.refresh()

    def select_used(self):
        self.selected_ids = {item["style_id"] for item in self.styles if item.get("count", 0) > 0}
        self.refresh()

    def select_none(self):
        self.selected_ids.clear()
        self.refresh()

    def get_selected_ids(self):
        return sorted(self.selected_ids)

    def _show_selected(self, event=None):
        selection = self.tree.selection()
        if not selection:
            return
        item = self.style_by_id.get(selection[0])
        if not item:
            return
        lines = [
            f"名称：{item.get('display_name') or item.get('name')}",
            f"类型：{self.TYPE_NAMES.get(item.get('type'), item.get('type'))}    使用次数：{item.get('count')}",
            f"字体：{item.get('font') or '继承'}    字号：{self._size(item.get('size')) or '继承'}    颜色：{item.get('color') or '自动'}",
            f"加粗：{'是' if item.get('bold') else '否'}    斜体：{'是' if item.get('italic') else '否'}",
            f"段前/段后：{self._points(item.get('before'))} / {self._points(item.get('after'))}    行距：{self._line(item.get('line'))}",
            f"基于：{chinese_style_name('', item.get('based_on')) or '-'}    后续：{chinese_style_name('', item.get('next')) or '-'}",
            f"编号定义：{item.get('num_id') or '-'}    列表级别：{item.get('level') or '-'}",
        ]
        self.attr_text.configure(state=NORMAL)
        self.attr_text.delete("1.0", END)
        self.attr_text.insert("1.0", "\n".join(lines))
        self.attr_text.configure(state=DISABLED)
        self._render_sample(item)

    def _render_sample(self, item):
        level = int(item.get("level") or 0)
        samples = [
            "第一章  一级标题",
            "1.1  二级标题",
            "1.1.1  三级标题",
            "1.1.1.1  四级标题",
            "正文段落示例：这是用于观察字体、字号、行距和段落间距的固定示例文字。",
        ]
        text = samples[min(level, 3)] if item.get("num_id") else samples[4]
        font = item.get("font") or "Microsoft YaHei UI"
        try:
            size = max(8, min(28, int(float(item.get("size") or 11))))
        except ValueError:
            size = 11
        color = item.get("color") or "202a35"
        color = f"#{color}" if len(color) == 6 else COLORS["text"]
        self.sample.configure(state=NORMAL)
        self.sample.delete("1.0", END)
        style_bits = []
        if item.get("bold"):
            style_bits.append("bold")
        if item.get("italic"):
            style_bits.append("italic")
        self.sample.tag_configure(
            "style",
            font=(font, size, " ".join(style_bits) or "normal"),
            foreground=color,
            lmargin1=min(level, 4) * 26,
            lmargin2=min(level, 4) * 26,
            spacing1=8,
            spacing3=8,
        )
        self.sample.insert("1.0", text, "style")
        self.sample.configure(state=DISABLED)
