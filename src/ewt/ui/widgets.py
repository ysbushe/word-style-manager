"""Reusable UI widgets."""

from __future__ import annotations

import tkinter as tk
from tkinter import filedialog

import ttkbootstrap as tb
from ttkbootstrap.constants import *


class PathRow(tb.Frame):
    def __init__(self, parent, label, variable, filetypes=None, mode="open", command=None):
        super().__init__(parent)
        self.variable = variable
        self.filetypes = filetypes or []
        self.mode = mode
        self.command = command or self.pick
        tb.Label(self, text=label, width=11, anchor=W).pack(side=LEFT)
        tb.Entry(self, textvariable=variable).pack(side=LEFT, fill=X, expand=True, padx=8)
        tb.Button(self, text="选择", bootstyle="outline-secondary", command=self.command).pack(side=LEFT)

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

    def __init__(self, parent, selectable=True, title="样式预览"):
        super().__init__(parent)
        self.selectable = selectable
        self.styles = []
        self.style_by_id = {}
        self.selected_ids = set()
        self.search_var = tk.StringVar()
        self.filter_var = tk.StringVar(value="全部")

        head = tb.Frame(self)
        head.pack(fill=X)
        tb.Label(head, text=title, font=("Microsoft YaHei UI", 12, "bold")).pack(side=LEFT)
        tb.Entry(head, textvariable=self.search_var, width=18).pack(side=RIGHT)
        tb.Label(head, text="搜索").pack(side=RIGHT, padx=6)
        self.search_var.trace_add("write", lambda *_: self.refresh())

        tools = tb.Frame(self)
        tools.pack(fill=X, pady=(7, 7))
        tb.Combobox(
            tools,
            textvariable=self.filter_var,
            values=["全部", "正在使用", "自定义", "带编号", "段落", "字符", "表格"],
            state="readonly",
            width=10,
        ).pack(side=LEFT)
        self.filter_var.trace_add("write", lambda *_: self.refresh())
        if selectable:
            tb.Button(tools, text="全选", bootstyle="link", command=self.select_all).pack(side=RIGHT)
            tb.Button(tools, text="只选已用", bootstyle="link", command=self.select_used).pack(side=RIGHT)
            tb.Button(tools, text="清空", bootstyle="link", command=self.select_none).pack(side=RIGHT)

        paned = tb.Panedwindow(self, orient=VERTICAL)
        paned.pack(fill=BOTH, expand=True)
        tree_frame = tb.Frame(paned)
        preview_frame = tb.Frame(paned)
        paned.add(tree_frame, weight=3)
        paned.add(preview_frame, weight=2)

        columns = ("pick", "name", "type", "count", "font", "size", "number")
        self.tree = tb.Treeview(tree_frame, columns=columns, show=HEADINGS, height=10, selectmode="browse")
        headings = {
            "pick": "选",
            "name": "样式名称",
            "type": "类型",
            "count": "次数",
            "font": "字体",
            "size": "字号",
            "number": "编号",
        }
        widths = {"pick": 36, "name": 150, "type": 58, "count": 52, "font": 105, "size": 48, "number": 65}
        for col in columns:
            self.tree.heading(col, text=headings[col])
            self.tree.column(col, width=widths[col], stretch=col in {"name", "font"})
        scroll = tb.Scrollbar(tree_frame, orient=VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side=LEFT, fill=BOTH, expand=True)
        scroll.pack(side=RIGHT, fill=Y)
        self.tree.bind("<<TreeviewSelect>>", self._show_selected)
        self.tree.bind("<Double-1>", self._toggle_selected)

        tabs = tb.Notebook(preview_frame)
        tabs.pack(fill=BOTH, expand=True, pady=(8, 0))
        attr_tab = tb.Frame(tabs, padding=10)
        sample_tab = tb.Frame(tabs, padding=10)
        tabs.add(attr_tab, text="属性")
        tabs.add(sample_tab, text="示例排版")
        self.attr_text = tk.Text(attr_tab, height=6, wrap="word", relief="flat", bg="#f8fafc")
        self.attr_text.pack(fill=BOTH, expand=True)
        self.sample = tk.Text(sample_tab, height=6, wrap="word", relief="flat", bg="white", padx=16, pady=12)
        self.sample.pack(fill=BOTH, expand=True)

    def set_data(self, styles):
        self.styles = list(styles or [])
        self.style_by_id = {item["style_id"]: item for item in self.styles}
        if self.selectable:
            self.selected_ids = {item["style_id"] for item in self.styles}
        self.refresh()

    def _matches(self, item):
        query = self.search_var.get().strip().lower()
        if query and query not in item.get("name", "").lower() and query not in item.get("style_id", "").lower():
            return False
        current = self.filter_var.get()
        if current == "正在使用":
            return item.get("count", 0) > 0
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
        for item in self.styles:
            if not self._matches(item):
                continue
            sid = item["style_id"]
            picked = "✓" if sid in self.selected_ids else ""
            number = f"{item.get('num_id', '')}/{item.get('level', '')}" if item.get("num_id") else ""
            self.tree.insert(
                "",
                END,
                iid=sid,
                values=(
                    picked,
                    item.get("name"),
                    self.TYPE_NAMES.get(item.get("type"), item.get("type")),
                    item.get("count"),
                    item.get("font"),
                    item.get("size"),
                    number,
                ),
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
            f"名称：{item.get('name')}    ID：{item.get('style_id')}",
            f"类型：{self.TYPE_NAMES.get(item.get('type'), item.get('type'))}    使用次数：{item.get('count')}",
            f"字体：{item.get('font') or '继承'}    字号：{item.get('size') or '继承'}    颜色：{item.get('color') or '自动'}",
            f"加粗：{'是' if item.get('bold') else '否'}    斜体：{'是' if item.get('italic') else '否'}",
            f"段前/段后：{item.get('before') or '-'} / {item.get('after') or '-'}    行距：{item.get('line') or '-'}",
            f"基于：{item.get('based_on') or '-'}    后续：{item.get('next') or '-'}",
            f"编号：{item.get('num_id') or '-'}    级别：{item.get('level') or '-'}",
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
        color = f"#{color}" if len(color) == 6 else "#202a35"
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
