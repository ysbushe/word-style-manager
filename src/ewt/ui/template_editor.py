"""Standalone template editor window."""

from __future__ import annotations

import os
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, simpledialog

import ttkbootstrap as tb
from ttkbootstrap.constants import *

from src.ewt.config import NUMBER_FORMATS
from src.ewt.core.editor import (
    clean_old_template_versions,
    load_numbering_presets,
    save_custom_preset,
    save_template_edits,
    template_editor_data,
)


class TemplateEditor(tb.Toplevel):
    def __init__(self, parent, template_path, library_dir, on_saved=None):
        super().__init__(parent)
        self.template_path = str(template_path)
        self.library_dir = str(library_dir)
        self.on_saved = on_saved
        self.title(f"模板编辑器 - {Path(template_path).name}")
        self.geometry("1180x760")
        self.minsize(1040, 680)
        self.transient(parent)

        self.data = template_editor_data(template_path)
        if not self.data.get("success"):
            messagebox.showerror("模板编辑器", self.data.get("error", "无法读取模板"), parent=self)
            self.destroy()
            return
        self.styles = self.data.get("styles", [])
        self.style_by_id = {item["style_id"]: item for item in self.styles}
        self.style_updates = {}
        self.level_rows = []
        self.presets = load_numbering_presets(library_dir)

        self._build()
        self._load_styles()
        self._load_presets()

    def _build(self):
        shell = tb.Frame(self, padding=16)
        shell.pack(fill=BOTH, expand=True)
        header = tb.Frame(shell)
        header.pack(fill=X, pady=(0, 12))
        tb.Label(header, text="模板编辑器", font=("Microsoft YaHei UI", 20, "bold")).pack(side=LEFT)
        tb.Label(header, text=Path(self.template_path).name, bootstyle=SECONDARY).pack(side=LEFT, padx=12)
        tb.Button(header, text="保存", bootstyle=SUCCESS, command=self.save).pack(side=RIGHT)

        paned = tb.Panedwindow(shell, orient=HORIZONTAL)
        paned.pack(fill=BOTH, expand=True)
        left = tb.Frame(paned, padding=(0, 0, 10, 0))
        right = tb.Frame(paned)
        paned.add(left, weight=1)
        paned.add(right, weight=3)

        tb.Label(left, text="样式", font=("Microsoft YaHei UI", 12, "bold")).pack(anchor=W)
        self.style_search = tk.StringVar()
        entry = tb.Entry(left, textvariable=self.style_search)
        entry.pack(fill=X, pady=8)
        entry.bind("<KeyRelease>", lambda _e: self._load_styles())
        self.style_list = tk.Listbox(left, activestyle="none")
        self.style_list.pack(fill=BOTH, expand=True)
        self.style_list.bind("<<ListboxSelect>>", self.load_style_form)

        tabs = tb.Notebook(right)
        tabs.pack(fill=BOTH, expand=True)
        style_tab = tb.Frame(tabs, padding=16)
        numbering_tab = tb.Frame(tabs, padding=16)
        tabs.add(style_tab, text="样式属性")
        tabs.add(numbering_tab, text="多级列表")
        self._build_style_form(style_tab)
        self._build_numbering_form(numbering_tab)

        footer = tb.Frame(shell)
        footer.pack(fill=X, pady=(12, 0))
        self.overwrite = tk.BooleanVar(value=False)
        tb.Checkbutton(footer, text="覆盖原模板（默认生成新版本）", variable=self.overwrite, bootstyle=WARNING).pack(side=LEFT)
        tb.Button(footer, text="清理旧版本", bootstyle="outline-danger", command=self.clean_versions).pack(side=LEFT, padx=12)
        tb.Button(footer, text="关闭", bootstyle=SECONDARY, command=self.destroy).pack(side=RIGHT)

    def _build_style_form(self, parent):
        self.style_vars = {
            "name": tk.StringVar(),
            "font": tk.StringVar(),
            "size": tk.StringVar(),
            "color": tk.StringVar(),
            "before": tk.StringVar(),
            "after": tk.StringVar(),
            "line": tk.StringVar(),
            "left": tk.StringVar(),
            "hanging": tk.StringVar(),
            "based_on": tk.StringVar(),
            "next": tk.StringVar(),
            "link": tk.StringVar(),
        }
        self.bold_var = tk.BooleanVar()
        self.italic_var = tk.BooleanVar()
        grid = tb.Frame(parent)
        grid.pack(fill=X)
        fields = [
            ("样式名称", "name"), ("字体", "font"), ("字号", "size"), ("颜色（HEX）", "color"),
            ("段前（twip）", "before"), ("段后（twip）", "after"), ("行距（twip）", "line"),
            ("左缩进", "left"), ("悬挂缩进", "hanging"), ("基于样式ID", "based_on"),
            ("后续样式ID", "next"), ("链接样式ID", "link"),
        ]
        for index, (label, key) in enumerate(fields):
            row, col = divmod(index, 2)
            frame = tb.Frame(grid)
            frame.grid(row=row, column=col, sticky=EW, padx=(0, 14), pady=6)
            tb.Label(frame, text=label, width=13).pack(side=LEFT)
            tb.Entry(frame, textvariable=self.style_vars[key]).pack(side=LEFT, fill=X, expand=True)
        grid.columnconfigure(0, weight=1)
        grid.columnconfigure(1, weight=1)
        flags = tb.Frame(parent)
        flags.pack(fill=X, pady=10)
        tb.Checkbutton(flags, text="加粗", variable=self.bold_var, bootstyle=SUCCESS).pack(side=LEFT)
        tb.Checkbutton(flags, text="斜体", variable=self.italic_var, bootstyle=SUCCESS).pack(side=LEFT, padx=18)
        tb.Button(flags, text="暂存当前样式修改", bootstyle=PRIMARY, command=self.store_style).pack(side=RIGHT)
        self.style_sample = tk.Text(parent, height=10, wrap="word", relief="flat", bg="#f8fafc", padx=20, pady=16)
        self.style_sample.pack(fill=BOTH, expand=True, pady=(10, 0))
        self.style_vars["name"].trace_add("write", lambda *_: self.update_style_sample())
        self.style_vars["font"].trace_add("write", lambda *_: self.update_style_sample())
        self.style_vars["size"].trace_add("write", lambda *_: self.update_style_sample())

    def _build_numbering_form(self, parent):
        top = tb.Frame(parent)
        top.pack(fill=X)
        tb.Label(top, text="常用编号方案").pack(side=LEFT)
        self.preset_var = tk.StringVar()
        self.preset_list = tk.Listbox(top, height=4, width=38, activestyle="none")
        self.preset_list.pack(side=LEFT, padx=8)
        self.preset_list.bind("<<ListboxSelect>>", self._select_preset)
        tb.Button(top, text="应用方案", bootstyle=PRIMARY, command=self.apply_preset).pack(side=LEFT)
        tb.Button(top, text="保存为自定义方案", bootstyle="outline-primary", command=self.save_preset).pack(side=LEFT, padx=8)

        switches = tb.Frame(parent)
        switches.pack(fill=X, pady=12)
        self.auto_link = tk.BooleanVar(value=True)
        self.auto_restart = tk.BooleanVar(value=True)
        tb.Checkbutton(switches, text="自动绑定 标题1~标题9", variable=self.auto_link, bootstyle=SUCCESS).pack(side=LEFT)
        tb.Checkbutton(switches, text="自动设置逐级重启", variable=self.auto_restart, bootstyle=SUCCESS).pack(side=LEFT, padx=18)

        self.level_container = tb.Frame(parent)
        self.level_container.pack(fill=BOTH, expand=True)
        for index in range(9):
            self._add_level_row(index)
        self.show_more = tk.BooleanVar(value=False)
        self.more_button = tb.Button(parent, text="显示 5-9 级", bootstyle="link", command=self.toggle_more)
        self.more_button.pack(anchor=W, pady=6)
        self._update_level_visibility()

        tb.Label(parent, text="编号预览", font=("Microsoft YaHei UI", 11, "bold")).pack(anchor=W, pady=(8, 4))
        self.number_preview = tk.Text(parent, height=7, relief="flat", bg="#f8fafc", padx=16, pady=10)
        self.number_preview.pack(fill=X)
        self.refresh_number_preview()

    def _add_level_row(self, index):
        frame = tb.Frame(self.level_container)
        vars_ = {
            "format": tk.StringVar(value=".".join(f"%{i}" for i in range(1, index + 2))),
            "numFmt": tk.StringVar(value="decimal"),
            "suffix": tk.StringVar(value="space"),
            "linked_style": tk.StringVar(value=f"Heading{index + 1}"),
            "start": tk.StringVar(value="1"),
            "left": tk.StringVar(value=str((index + 1) * 720)),
            "hanging": tk.StringVar(value="360"),
        }
        tb.Label(frame, text=f"{index + 1}级", width=5).pack(side=LEFT)
        tb.Entry(frame, textvariable=vars_["format"], width=16).pack(side=LEFT, padx=3)
        tb.Combobox(frame, textvariable=vars_["numFmt"], values=list(NUMBER_FORMATS), width=17).pack(side=LEFT, padx=3)
        tb.Combobox(frame, textvariable=vars_["suffix"], values=["space", "tab", "nothing"], width=9).pack(side=LEFT, padx=3)
        tb.Entry(frame, textvariable=vars_["linked_style"], width=13).pack(side=LEFT, padx=3)
        tb.Entry(frame, textvariable=vars_["start"], width=5).pack(side=LEFT, padx=3)
        tb.Entry(frame, textvariable=vars_["left"], width=7).pack(side=LEFT, padx=3)
        tb.Entry(frame, textvariable=vars_["hanging"], width=7).pack(side=LEFT, padx=3)
        for variable in vars_.values():
            variable.trace_add("write", lambda *_: self.refresh_number_preview())
        self.level_rows.append((frame, vars_))

    def _load_styles(self):
        query = self.style_search.get().lower() if hasattr(self, "style_search") else ""
        self.style_list.delete(0, END)
        self.visible_style_ids = []
        for item in self.styles:
            if query and query not in item.get("name", "").lower() and query not in item.get("style_id", "").lower():
                continue
            self.visible_style_ids.append(item["style_id"])
            self.style_list.insert(END, f"{item.get('name')}  [{item.get('style_id')}]")

    def _load_presets(self):
        names = [item.get("name") for item in self.presets]
        self.preset_list.delete(0, END)
        for name in names:
            self.preset_list.insert(END, name)
        if names:
            self.preset_var.set(names[0])
            self.preset_list.selection_set(0)
            self.apply_preset()

    def _select_preset(self, event=None):
        selection = self.preset_list.curselection()
        if selection:
            self.preset_var.set(self.preset_list.get(selection[0]))
            self.apply_preset()

    def load_style_form(self, event=None):
        selection = self.style_list.curselection()
        if not selection:
            return
        sid = self.visible_style_ids[selection[0]]
        self.current_style_id = sid
        item = self.style_by_id[sid]
        for key, variable in self.style_vars.items():
            variable.set(item.get(key, ""))
        self.bold_var.set(bool(item.get("bold")))
        self.italic_var.set(bool(item.get("italic")))
        self.update_style_sample()

    def store_style(self):
        if not getattr(self, "current_style_id", ""):
            return
        values = {key: variable.get() for key, variable in self.style_vars.items()}
        values["bold"] = self.bold_var.get()
        values["italic"] = self.italic_var.get()
        self.style_updates[self.current_style_id] = values
        messagebox.showinfo("模板编辑器", "当前样式修改已暂存，点击“保存”写入模板。", parent=self)

    def update_style_sample(self):
        if not hasattr(self, "style_sample"):
            return
        font = self.style_vars["font"].get() or "Microsoft YaHei UI"
        try:
            size = max(8, min(30, int(float(self.style_vars["size"].get() or 12))))
        except ValueError:
            size = 12
        self.style_sample.configure(state=NORMAL)
        self.style_sample.delete("1.0", END)
        style_bits = []
        if self.bold_var.get():
            style_bits.append("bold")
        if self.italic_var.get():
            style_bits.append("italic")
        self.style_sample.tag_configure(
            "sample",
            font=(font, size, " ".join(style_bits) or "normal"),
            spacing1=12,
            spacing3=12,
        )
        self.style_sample.insert("1.0", "第一章  一级标题\n1.1  二级标题\n正文段落示例文字", "sample")
        self.style_sample.configure(state=DISABLED)

    def toggle_more(self):
        self.show_more.set(not self.show_more.get())
        self._update_level_visibility()

    def _update_level_visibility(self):
        for index, (frame, _vars) in enumerate(self.level_rows):
            if index < 4 or self.show_more.get():
                frame.pack(fill=X, pady=3)
            else:
                frame.pack_forget()
        self.more_button.configure(text="收起 5-9 级" if self.show_more.get() else "显示 5-9 级")

    def level_data(self):
        return [{key: variable.get() for key, variable in vars_.items()} for _frame, vars_ in self.level_rows]

    def apply_preset(self):
        preset = next((item for item in self.presets if item.get("name") == self.preset_var.get()), None)
        if not preset:
            return
        levels = preset.get("levels", [])
        for index, (_frame, vars_) in enumerate(self.level_rows):
            if index >= len(levels):
                continue
            for key, value in levels[index].items():
                if key in vars_:
                    vars_[key].set(value)
        self.refresh_number_preview()

    def save_preset(self):
        name = simpledialog.askstring("自定义编号方案", "方案名称：", parent=self)
        if not name:
            return
        preset = {"name": name, "levels": self.level_data()}
        save_custom_preset(self.library_dir, preset)
        self.presets = load_numbering_presets(self.library_dir)
        self._load_presets()
        for index in range(self.preset_list.size()):
            if self.preset_list.get(index) == name:
                self.preset_list.selection_clear(0, END)
                self.preset_list.selection_set(index)
                self.preset_var.set(name)
                break

    def refresh_number_preview(self):
        if not hasattr(self, "number_preview"):
            return
        samples = ["一级标题", "二级标题", "三级标题", "四级标题"]
        lines = []
        for index in range(4):
            fmt = self.level_rows[index][1]["format"].get()
            for number in range(1, index + 2):
                fmt = fmt.replace(f"%{number}", "1")
            suffix = {"space": " ", "tab": "    ", "nothing": ""}.get(self.level_rows[index][1]["suffix"].get(), " ")
            lines.append(f"{'  ' * index}{fmt}{suffix}{samples[index]}")
        self.number_preview.configure(state=NORMAL)
        self.number_preview.delete("1.0", END)
        self.number_preview.insert("1.0", "\n".join(lines))
        self.number_preview.configure(state=DISABLED)

    def save(self):
        if getattr(self, "current_style_id", ""):
            values = {key: variable.get() for key, variable in self.style_vars.items()}
            values["bold"] = self.bold_var.get()
            values["italic"] = self.italic_var.get()
            self.style_updates[self.current_style_id] = values
        try:
            result = save_template_edits(
                self.template_path,
                self.style_updates,
                self.level_data(),
                self.auto_link.get(),
                self.auto_restart.get(),
                self.overwrite.get(),
            )
        except Exception as exc:
            messagebox.showerror("模板编辑器", str(exc), parent=self)
            return
        messagebox.showinfo("模板编辑器", f"已保存：\n{result['output']}", parent=self)
        if self.on_saved:
            self.on_saved(result["output"])
        self.destroy()

    def clean_versions(self):
        keep = simpledialog.askinteger("清理旧版本", "保留最新几个版本？", initialvalue=3, minvalue=1, parent=self)
        if keep is None:
            return
        base = Path(self.template_path).stem.split("_v20")[0]
        removed = clean_old_template_versions(self.library_dir, base, keep)
        messagebox.showinfo("清理旧版本", f"已清理 {len(removed)} 个旧版本。", parent=self)
