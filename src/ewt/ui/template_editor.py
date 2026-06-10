"""Standalone template editor window."""

from __future__ import annotations

import os
import re
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, simpledialog

import ttkbootstrap as tb
from ttkbootstrap.constants import *

from src.ewt.core.editor import (
    clean_old_template_versions,
    load_numbering_presets,
    save_template_edits,
    template_editor_data,
)
from src.ewt.utils.display import CHINESE_LEVELS, chinese_style_name, make_choice_maps
from src.ewt.ui.theme import COLORS, configure_listbox, configure_office_theme, configure_text

NUMBER_FORMAT_LABELS = {
    "decimal": "阿拉伯数字（1、2、3）",
    "decimalZero": "补零数字（01、02、03）",
    "chineseCounting": "中文数字（一、二、三）",
    "chineseCountingThousand": "中文计数（一、二、三）",
    "upperLetter": "大写字母（A、B、C）",
    "lowerLetter": "小写字母（a、b、c）",
    "upperRoman": "大写罗马数字（I、II、III）",
    "lowerRoman": "小写罗马数字（i、ii、iii）",
}
NUMBER_FORMAT_VALUES = {label: key for key, label in NUMBER_FORMAT_LABELS.items()}
SUFFIX_LABELS = {"space": "空格", "tab": "制表符", "nothing": "无"}
SUFFIX_VALUES = {label: key for key, label in SUFFIX_LABELS.items()}

SIZE_CHOICES = {
    "八号（5 磅）": "5",
    "七号（5.5 磅）": "5.5",
    "小六（6.5 磅）": "6.5",
    "六号（7.5 磅）": "7.5",
    "小五（9 磅）": "9",
    "五号（10.5 磅）": "10.5",
    "小四（12 磅）": "12",
    "四号（14 磅）": "14",
    "小三（15 磅）": "15",
    "三号（16 磅）": "16",
    "小二（18 磅）": "18",
    "二号（22 磅）": "22",
    "小一（24 磅）": "24",
    "一号（26 磅）": "26",
    "小初（36 磅）": "36",
}
SPACING_CHOICES = {
    "0 磅": "0",
    "3 磅": "60",
    "6 磅": "120",
    "9 磅": "180",
    "12 磅": "240",
    "18 磅": "360",
    "24 磅": "480",
}
LINE_CHOICES = {
    "单倍行距": "240",
    "1.15 倍行距": "276",
    "1.25 倍行距": "300",
    "1.5 倍行距": "360",
    "2 倍行距": "480",
}
INDENT_CHOICES = {
    "0 厘米": "0",
    "0.32 厘米": "180",
    "0.63 厘米": "360",
    "0.95 厘米": "540",
    "1.27 厘米": "720",
    "1.9 厘米": "1080",
    "2.54 厘米": "1440",
    "3.17 厘米": "1800",
    "3.81 厘米": "2160",
}
COLOR_CHOICES = {
    "黑色（#000000）": "000000",
    "深蓝（#1F4E79）": "1F4E79",
    "蓝色（#2F5597）": "2F5597",
    "红色（#C00000）": "C00000",
    "紫色（#7030A0）": "7030A0",
    "灰色（#808080）": "808080",
}
FORMAT_CHOICES = {
    "本级编号（%1）": "%1",
    "章标题（第%1章）": "第%1章",
    "条标题（第%1条）": "第%1条",
    "两级编号（%1.%2）": "%1.%2",
    "三级编号（%1.%2.%3）": "%1.%2.%3",
    "四级编号（%1.%2.%3.%4）": "%1.%2.%3.%4",
    "五级编号（%1.%2.%3.%4.%5）": "%1.%2.%3.%4.%5",
    "六级编号（%1.%2.%3.%4.%5.%6）": "%1.%2.%3.%4.%5.%6",
    "七级编号（%1.%2.%3.%4.%5.%6.%7）": "%1.%2.%3.%4.%5.%6.%7",
    "八级编号（%1.%2.%3.%4.%5.%6.%7.%8）": "%1.%2.%3.%4.%5.%6.%7.%8",
    "九级编号（%1.%2.%3.%4.%5.%6.%7.%8.%9）": "%1.%2.%3.%4.%5.%6.%7.%8.%9",
    "括号编号（（%1））": "（%1）",
    "顿号编号（%1、）": "%1、",
}
FORMAT_VALUES = {value: label for label, value in FORMAT_CHOICES.items()}
STYLE_UNIT_OPTIONS = {
    "before": ("磅", "行"),
    "after": ("磅", "行"),
    "line": ("倍", "行", "磅"),
    "left": ("厘米", "毫米", "字符", "磅"),
    "hanging": ("厘米", "毫米", "字符", "磅"),
}
STYLE_VALUE_OPTIONS = {
    ("before", "磅"): ("0", "3", "6", "9", "12", "18", "24"),
    ("before", "行"): ("0", "0.5", "1", "1.5", "2"),
    ("after", "磅"): ("0", "3", "6", "9", "12", "18", "24"),
    ("after", "行"): ("0", "0.5", "1", "1.5", "2"),
    ("line", "倍"): ("1", "1.15", "1.25", "1.5", "2"),
    ("line", "行"): ("1", "1.15", "1.25", "1.5", "2"),
    ("line", "磅"): ("9", "10.5", "12", "14", "16", "18", "20", "24", "28", "36"),
    ("left", "厘米"): ("0", "0.32", "0.63", "0.95", "1.27", "1.9", "2.54", "3.17", "3.81"),
    ("left", "毫米"): ("0", "3.2", "6.3", "9.5", "12.7", "19", "25.4", "31.7", "38.1"),
    ("left", "字符"): ("0", "0.5", "1", "1.5", "2", "3", "4"),
    ("left", "磅"): ("0", "9", "18", "27", "36", "54", "72", "90", "108"),
    ("hanging", "厘米"): ("0", "0.32", "0.63", "0.95", "1.27", "1.9", "2.54"),
    ("hanging", "毫米"): ("0", "3.2", "6.3", "9.5", "12.7", "19", "25.4"),
    ("hanging", "字符"): ("0", "0.5", "1", "1.5", "2", "3", "4"),
    ("hanging", "磅"): ("0", "9", "18", "27", "36", "54", "72"),
}


def _format_number(value, number_format):
    try:
        number = max(1, int(value))
    except (TypeError, ValueError):
        number = 1
    if number_format == "decimalZero":
        return f"{number:02d}"
    if number_format in {"upperLetter", "lowerLetter"}:
        result = ""
        current = number
        while current:
            current, remainder = divmod(current - 1, 26)
            result = chr(ord("A") + remainder) + result
        return result if number_format == "upperLetter" else result.lower()
    if number_format in {"upperRoman", "lowerRoman"}:
        values = (
            (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"),
            (100, "C"), (90, "XC"), (50, "L"), (40, "XL"),
            (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I"),
        )
        current = number
        result = ""
        for amount, symbol in values:
            while current >= amount:
                result += symbol
                current -= amount
        return result if number_format == "upperRoman" else result.lower()
    if number_format in {"chineseCounting", "chineseCountingThousand"}:
        digits = "零一二三四五六七八九"
        if number < 10:
            return digits[number]
        if number < 20:
            return "十" + (digits[number % 10] if number % 10 else "")
        if number < 100:
            return digits[number // 10] + "十" + (digits[number % 10] if number % 10 else "")
    return str(number)


def _display_value(raw, choices, unit, divisor=1):
    raw = str(raw or "")
    for label, value in choices.items():
        if raw == value:
            return label
    if not raw:
        return ""
    try:
        return f"{float(raw) / divisor:g} {unit}"
    except ValueError:
        return raw


def _raw_value(display, choices, multiplier=1):
    display = str(display or "").strip()
    if display in choices:
        return choices[display]
    match = re.search(r"-?\d+(?:\.\d+)?", display)
    if not match:
        return display
    return str(round(float(match.group()) * multiplier))


def _open_combobox_from_body(event):
    """让下拉输入框点击文字区域时也展开选项。"""
    widget = event.widget
    if widget.instate(["disabled"]):
        return
    if "downarrow" in str(widget.identify(event.x, event.y)):
        return
    widget.after_idle(lambda: widget.event_generate("<Down>"))


def _bind_combobox_body(control):
    if getattr(control, "_body_click_bound", False):
        return control
    control.bind("<Button-1>", _open_combobox_from_body, add="+")
    control._body_click_bound = True
    return control


def _bind_all_comboboxes(parent):
    for child in parent.winfo_children():
        if isinstance(child, tb.Combobox):
            _bind_combobox_body(child)
        _bind_all_comboboxes(child)


class TemplateEditor(tb.Toplevel):
    def __init__(self, parent, template_path, library_dir, on_saved=None):
        super().__init__(parent)
        self.template_path = str(template_path)
        self.library_dir = str(library_dir)
        self.on_saved = on_saved
        self.title(f"模板编辑器 - {Path(template_path).name}")
        self.geometry("1360x860")
        self.minsize(1120, 700)
        self.resizable(True, True)
        configure_office_theme(getattr(parent, "style", None) or tb.Style())
        self.configure(background=COLORS["background"])

        self.data = template_editor_data(template_path)
        if not self.data.get("success"):
            messagebox.showerror("模板编辑器", self.data.get("error", "无法读取模板"), parent=self)
            self.destroy()
            return
        self.styles = self.data.get("styles", [])
        self.style_by_id = {item["style_id"]: item for item in self.styles}
        self.style_label_to_id, self.style_id_to_label = make_choice_maps(self.styles)
        for level, chinese_level in enumerate(CHINESE_LEVELS, start=1):
            label = f"{chinese_level}级标题"
            style_id = f"Heading{level}"
            self.style_label_to_id.setdefault(label, style_id)
            self.style_id_to_label.setdefault(style_id, label)
        self.style_updates = {}
        self.level_rows = []
        self.numbering_editor_enabled = False
        self.presets = load_numbering_presets(library_dir)

        self._build()
        _bind_all_comboboxes(self)
        self._load_styles()
        self._load_presets()
        self._capture_initial_state()
        self.after_idle(self._center_on_parent)

    def _center_on_parent(self):
        self.update_idletasks()
        width = max(self.winfo_width(), 1120)
        height = max(self.winfo_height(), 700)
        parent = self.master
        x = parent.winfo_rootx() + max(0, (parent.winfo_width() - width) // 2)
        y = parent.winfo_rooty() + max(0, (parent.winfo_height() - height) // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")

    def _capture_initial_state(self):
        self.initial_level_state = [
            {key: variable.get() for key, variable in variables.items()}
            for _frame, variables in self.level_rows
        ]
        self.initial_auto_link = self.auto_link.get()
        self.initial_auto_restart = self.auto_restart.get()
        self.initial_overwrite = self.overwrite.get()

    def _build(self):
        shell = tb.Frame(self, padding=14, style="Office.TFrame")
        shell.pack(fill=BOTH, expand=True)
        header = tb.Frame(shell, style="Office.TFrame")
        header.pack(fill=X, pady=(0, 8))
        heading = tb.Frame(header, style="Office.TFrame")
        heading.pack(side=LEFT, fill=X, expand=True)
        tb.Label(heading, text="模板编辑器", style="DialogTitle.TLabel").pack(side=LEFT)
        tb.Label(
            heading,
            text=Path(self.template_path).name,
            style="PageHint.TLabel",
        ).pack(side=LEFT, padx=(12, 0))

        paned = tb.Panedwindow(shell, orient=HORIZONTAL)
        paned.pack(fill=BOTH, expand=True)
        left = tb.Frame(paned, padding=12, width=270, style="Card.TFrame")
        right = tb.Frame(paned)
        paned.add(left, weight=1)
        paned.add(right, weight=3)

        tb.Label(left, text="模板中的样式", style="CardTitle.TLabel").pack(anchor=W)
        tb.Label(left, text="选择样式后在右侧调整属性", style="Muted.TLabel").pack(anchor=W, pady=(3, 8))
        style_list_frame = tb.Frame(left, style="Card.TFrame")
        style_list_frame.pack(fill=BOTH, expand=True, pady=(8, 0))
        self.style_list = tk.Listbox(style_list_frame, activestyle="none", exportselection=False)
        configure_listbox(self.style_list)
        style_scroll = tb.Scrollbar(style_list_frame, orient=VERTICAL, command=self.style_list.yview)
        self.style_list.configure(yscrollcommand=style_scroll.set)
        self.style_list.pack(side=LEFT, fill=BOTH, expand=True)
        style_scroll.pack(side=RIGHT, fill=Y)
        self.style_list.bind("<<ListboxSelect>>", self.load_style_form)

        tabs = tb.Notebook(right)
        tabs.pack(fill=BOTH, expand=True)
        style_tab = tb.Frame(tabs, padding=8)
        numbering_tab = tb.Frame(tabs, padding=8)
        tabs.add(style_tab, text="样式属性")
        self._build_style_form(style_tab)
        self._build_numbering_form(numbering_tab)

        footer = tb.Frame(shell, padding=(0, 7, 0, 0), style="Office.TFrame")
        footer.pack(fill=X)
        self.overwrite = tk.BooleanVar(value=False)
        tb.Checkbutton(footer, text="覆盖原模板（默认生成新版本）", variable=self.overwrite).pack(side=LEFT)
        tb.Button(footer, text="清理旧版本", style="Danger.TButton", command=self.clean_versions).pack(side=LEFT, padx=12)
        tb.Button(footer, text="关闭", style="Ghost.TButton", command=self.destroy).pack(side=RIGHT)
        tb.Button(
            footer,
            text="重置全部修改",
            style="Ghost.TButton",
            command=self.reset_all_edits,
        ).pack(side=RIGHT, padx=(0, 8))
        tb.Button(
            footer,
            text="保存模板修改",
            width=16,
            style="Primary.TButton",
            command=self.save,
        ).pack(side=RIGHT, padx=8)

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
        self.style_unit_vars = {
            "before": tk.StringVar(value="磅"),
            "after": tk.StringVar(value="磅"),
            "line": tk.StringVar(value="倍"),
            "left": tk.StringVar(value="厘米"),
            "hanging": tk.StringVar(value="厘米"),
        }
        self._last_style_units = {
            key: variable.get() for key, variable in self.style_unit_vars.items()
        }
        self.style_value_controls = {}
        property_group = tb.Labelframe(
            parent,
            text="样式属性",
            padding=8,
            style="Card.TLabelframe",
        )
        property_group.pack(fill=X)
        choices = {
            "font": ["宋体", "仿宋", "黑体", "楷体", "微软雅黑", "等线", "Arial", "Times New Roman"],
            "size": list(SIZE_CHOICES),
            "color": list(COLOR_CHOICES),
            "based_on": ["无"] + list(self.style_label_to_id),
            "next": ["无"] + list(self.style_label_to_id),
            "link": ["无"] + list(self.style_label_to_id),
        }
        fields = (
            ("样式名称", "name"),
            ("中文字体", "font"),
            ("字号", "size"),
            ("文字颜色", "color"),
            ("段前间距", "before"),
            ("段后间距", "after"),
            ("行距", "line"),
            ("左缩进", "left"),
            ("悬挂缩进", "hanging"),
            ("基于样式", "based_on"),
            ("后续样式", "next"),
            ("链接样式", "link"),
        )
        for index, (label, key) in enumerate(fields):
            row, col = divmod(index, 4)
            field = tb.Frame(property_group, style="Card.TFrame")
            field.grid(
                row=row,
                column=col,
                sticky=EW,
                padx=(0 if col == 0 else 5, 5 if col < 3 else 0),
                pady=2,
            )
            tb.Label(field, text=label, style="Muted.TLabel").pack(anchor=W)
            if key in self.style_unit_vars:
                value_row = tb.Frame(field, style="Card.TFrame")
                value_row.pack(fill=X, expand=True, pady=(1, 0))
                control = _bind_combobox_body(
                    tb.Combobox(
                        value_row,
                        textvariable=self.style_vars[key],
                        values=STYLE_VALUE_OPTIONS.get(
                            (key, self.style_unit_vars[key].get()),
                            (),
                        ),
                        state="normal",
                        width=12,
                    )
                )
                control.pack(side=LEFT, fill=X, expand=True)
                self.style_value_controls[key] = control
                unit_control = _bind_combobox_body(
                    tb.Combobox(
                        value_row,
                        textvariable=self.style_unit_vars[key],
                        values=STYLE_UNIT_OPTIONS[key],
                        state="readonly",
                        width=7,
                    )
                )
                unit_control.pack(side=LEFT, padx=(6, 0))
                unit_control.bind(
                    "<<ComboboxSelected>>",
                    lambda _event, field_key=key: self._change_style_unit(field_key),
                )
            elif key == "color":
                color_row = tb.Frame(field, style="Card.TFrame")
                color_row.pack(fill=X, expand=True, pady=(1, 0))
                self.color_swatch = tk.Label(
                    color_row,
                    width=3,
                    relief="solid",
                    borderwidth=1,
                    background="#000000",
                )
                self.color_swatch.pack(side=LEFT, fill=Y, padx=(0, 4))
                control = tk.Menubutton(
                    color_row,
                    textvariable=self.style_vars[key],
                    anchor="w",
                    relief="solid",
                    borderwidth=1,
                    background="#FFFFFF",
                    foreground=COLORS["text"],
                    font=("Microsoft YaHei UI", 10),
                    padx=6,
                )
                color_menu = tk.Menu(control, tearoff=False)
                control.configure(menu=color_menu)
                for color_label, color_value in COLOR_CHOICES.items():
                    color_menu.add_radiobutton(
                        label=f"■  {color_label}",
                        value=color_label,
                        variable=self.style_vars[key],
                        foreground=f"#{color_value}",
                        activeforeground=f"#{color_value}",
                    )
                control.pack(side=LEFT, fill=X, expand=True)
            else:
                control = (
                    tb.Entry(field, textvariable=self.style_vars[key])
                    if key == "name"
                    else _bind_combobox_body(
                        tb.Combobox(
                            field,
                            textvariable=self.style_vars[key],
                            values=choices.get(key, []),
                            state="normal",
                        )
                    )
                )
                control.pack(fill=X, expand=True, pady=(1, 0))
        for column in range(4):
            property_group.columnconfigure(column, weight=1, uniform="style_fields")

        flags = tb.Frame(property_group, style="Card.TFrame")
        flags.grid(row=3, column=0, columnspan=4, sticky=EW, pady=(5, 0))
        tb.Checkbutton(flags, text="加粗", variable=self.bold_var).pack(side=LEFT)
        tb.Checkbutton(flags, text="斜体", variable=self.italic_var).pack(side=LEFT, padx=12)
        tb.Button(
            flags,
            text="保存当前样式修改",
            style="Primary.TButton",
            command=self.save_current_style,
        ).pack(side=RIGHT)
        tb.Button(
            flags,
            text="撤销样式修改",
            style="Ghost.TButton",
            command=self.reset_current_style,
        ).pack(side=RIGHT, padx=(0, 6))

        sample_box = tb.Labelframe(parent, text="排版预览", padding=8, style="Card.TLabelframe")
        sample_box.pack(fill=BOTH, expand=True, pady=(7, 0))
        self.style_preview_info = tk.StringVar()
        tb.Label(
            sample_box,
            textvariable=self.style_preview_info,
            style="Muted.TLabel",
        ).pack(fill=X, pady=(0, 4))
        self.style_sample = tk.Text(sample_box, height=14, wrap="word", relief="flat", padx=20, pady=12)
        configure_text(self.style_sample, COLORS["surface_alt"])
        self.style_sample.pack(fill=BOTH, expand=True)
        for variable in self.style_vars.values():
            variable.trace_add("write", lambda *_: self.update_style_sample())
        self.style_vars["color"].trace_add("write", lambda *_: self._update_color_swatch())
        self.bold_var.trace_add("write", lambda *_: self.update_style_sample())
        self.italic_var.trace_add("write", lambda *_: self.update_style_sample())

    def _build_numbering_form(self, parent):
        top = tb.Labelframe(parent, text="常用编号方案", padding=8, style="Card.TLabelframe")
        top.pack(fill=X)
        self.preset_var = tk.StringVar()
        self.preset_list = tk.Listbox(top, height=2, width=44, activestyle="none", exportselection=False)
        configure_listbox(self.preset_list)
        self.preset_list.pack(side=LEFT, fill=X, expand=True)
        self.preset_list.bind("<<ListboxSelect>>", self._select_preset)
        preset_actions = tb.Frame(top)
        preset_actions.pack(side=RIGHT, padx=(8, 0))
        tb.Button(
            preset_actions,
            text="载入选中方案",
            width=16,
            style="Secondary.TButton",
            command=self.apply_preset,
        ).pack(fill=X)

        switches = tb.Frame(parent)
        switches.pack(fill=X, pady=6)
        self.auto_link = tk.BooleanVar(value=True)
        self.auto_restart = tk.BooleanVar(value=True)
        tb.Checkbutton(switches, text="自动绑定 一级标题～九级标题", variable=self.auto_link).pack(side=LEFT)
        tb.Checkbutton(switches, text="自动设置逐级重启", variable=self.auto_restart).pack(side=LEFT, padx=18)
        tb.Button(
            switches,
            text="保存修改",
            style="Primary.TButton",
            command=self.save,
        ).pack(side=RIGHT)
        tb.Button(
            switches,
            text="撤销编号修改",
            style="Ghost.TButton",
            command=self.reset_numbering_edits,
        ).pack(side=RIGHT, padx=(0, 6))

        levels_box = tb.Labelframe(parent, text="各级编号设置", padding=6, style="Card.TLabelframe")
        levels_box.pack(fill=X)
        self.level_canvas = tk.Canvas(
            levels_box,
            height=154,
            highlightthickness=0,
            background=COLORS["surface"],
        )
        level_scroll_x = tb.Scrollbar(levels_box, orient=HORIZONTAL, command=self.level_canvas.xview)
        level_scroll_y = tb.Scrollbar(levels_box, orient=VERTICAL, command=self.level_canvas.yview)
        self.level_canvas.configure(
            xscrollcommand=level_scroll_x.set,
            yscrollcommand=level_scroll_y.set,
        )
        self.level_canvas.grid(row=0, column=0, sticky=EW)
        level_scroll_y.grid(row=0, column=1, sticky=NS)
        level_scroll_x.grid(row=1, column=0, sticky=EW, pady=(4, 0))
        levels_box.columnconfigure(0, weight=1)
        self.level_container = tb.Frame(self.level_canvas)
        self.level_window = self.level_canvas.create_window((0, 0), window=self.level_container, anchor=NW)
        self.level_container.bind(
            "<Configure>",
            lambda _event: self.level_canvas.configure(scrollregion=self.level_canvas.bbox("all")),
        )
        self.level_canvas.bind(
            "<Configure>",
            lambda event: self.level_canvas.itemconfigure(
                self.level_window,
                width=max(event.width, self.level_container.winfo_reqwidth()),
            ),
        )
        self.level_canvas.bind("<MouseWheel>", self._scroll_levels)
        self.level_container.bind("<MouseWheel>", self._scroll_levels)
        self.level_canvas.bind("<Shift-MouseWheel>", self._scroll_levels_horizontal)
        self.level_container.bind("<Shift-MouseWheel>", self._scroll_levels_horizontal)
        headers = tb.Frame(self.level_container)
        headers.pack(fill=X, pady=(0, 4))
        for text, width in (
            ("级别", 6),
            ("编号格式", 18),
            ("编号样式", 24),
            ("编号后跟随", 11),
            ("关联样式", 16),
            ("起始值", 7),
            ("左缩进（厘米）", 12),
            ("悬挂（厘米）", 12),
        ):
            tb.Label(headers, text=text, width=width, anchor=CENTER, style="Muted.TLabel").pack(side=LEFT, padx=2)
        for index in range(9):
            self._add_level_row(index)
        self.show_more = tk.BooleanVar(value=False)
        self.more_button = tb.Button(
            parent,
            text="展开并编辑 5-9 级 ▼",
            style="Secondary.TButton",
            command=self.toggle_more,
        )
        self.more_button.pack(fill=X, pady=4)
        self._update_level_visibility()

        preview_box = tb.Labelframe(parent, text="编号预览", padding=6, style="Card.TLabelframe")
        preview_box.pack(fill=BOTH, expand=True, pady=(2, 0))
        tb.Label(
            preview_box,
            text="每一级的首行显示编号位置，换行后显示正文对齐位置。",
            style="Muted.TLabel",
        ).pack(fill=X, pady=(0, 4))
        self.number_preview = tk.Text(preview_box, height=14, wrap="word", relief="flat", padx=16, pady=8)
        configure_text(self.number_preview, COLORS["surface_alt"])
        self.number_preview.pack(fill=BOTH, expand=True)
        self.refresh_number_preview()

    def _add_level_row(self, index):
        frame = tb.Frame(self.level_container)
        vars_ = {
            "format": tk.StringVar(
                value=FORMAT_VALUES.get(
                    ".".join(f"%{i}" for i in range(1, index + 2)),
                    ".".join(f"%{i}" for i in range(1, index + 2)),
                )
            ),
            "numFmt": tk.StringVar(value=NUMBER_FORMAT_LABELS["decimal"]),
            "suffix": tk.StringVar(value=SUFFIX_LABELS["space"]),
            "linked_style": tk.StringVar(
                value=self.style_id_to_label.get(
                    f"Heading{index + 1}",
                    f"{CHINESE_LEVELS[index]}级标题",
                )
            ),
            "start": tk.StringVar(value="1"),
            "left": tk.StringVar(value=_display_value(str((index + 1) * 720), INDENT_CHOICES, "厘米", 567)),
            "hanging": tk.StringVar(value=_display_value("360", INDENT_CHOICES, "厘米", 567)),
        }
        tb.Label(frame, text=f"第 {index + 1} 级", width=6, anchor=CENTER).pack(side=LEFT, padx=2)
        tb.Combobox(
            frame,
            textvariable=vars_["format"],
            values=list(FORMAT_CHOICES),
            state="normal",
            width=16,
        ).pack(side=LEFT, padx=2)
        tb.Combobox(
            frame,
            textvariable=vars_["numFmt"],
            values=list(NUMBER_FORMAT_VALUES),
            state="readonly",
            width=22,
        ).pack(side=LEFT, padx=2)
        tb.Combobox(
            frame,
            textvariable=vars_["suffix"],
            values=list(SUFFIX_VALUES),
            state="readonly",
            width=9,
        ).pack(side=LEFT, padx=2)
        tb.Combobox(
            frame,
            textvariable=vars_["linked_style"],
            values=list(self.style_label_to_id),
            state="normal",
            width=22,
        ).pack(side=LEFT, padx=2)
        tb.Combobox(frame, textvariable=vars_["start"], values=["1（从 1 开始）", "2（从 2 开始）", "3（从 3 开始）"], state="normal", width=11).pack(side=LEFT, padx=2)
        tb.Combobox(
            frame,
            textvariable=vars_["left"],
            values=list(INDENT_CHOICES),
            state="normal",
            width=10,
        ).pack(side=LEFT, padx=2)
        tb.Combobox(
            frame,
            textvariable=vars_["hanging"],
            values=list(INDENT_CHOICES),
            state="normal",
            width=10,
        ).pack(side=LEFT, padx=2)
        for variable in vars_.values():
            variable.trace_add("write", lambda *_: self.refresh_number_preview())
        for widget in (frame, *frame.winfo_children()):
            widget.bind("<MouseWheel>", self._scroll_levels, add="+")
            widget.bind("<Shift-MouseWheel>", self._scroll_levels_horizontal, add="+")
        self.level_rows.append((frame, vars_))

    def _load_styles(self):
        self.style_list.delete(0, END)
        self.visible_style_ids = []
        for item in self.styles:
            self.visible_style_ids.append(item["style_id"])
            self.style_list.insert(END, item.get("display_name") or chinese_style_name(item.get("name"), item.get("style_id")))

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

    def _font_size_points(self):
        try:
            return max(1.0, float(_raw_value(self.style_vars["size"].get(), SIZE_CHOICES) or 12))
        except (TypeError, ValueError):
            return 12.0

    def _style_value_to_twips(self, key, value=None, unit=None):
        try:
            number = float(self.style_vars[key].get() if value is None else value)
        except (TypeError, ValueError):
            return 0.0
        unit = unit or self.style_unit_vars[key].get()
        if unit == "厘米":
            return number * 567
        if unit == "毫米":
            return number * 56.7
        if unit == "磅":
            return number * 20
        if unit == "字符":
            return number * self._font_size_points() * 20
        if unit in {"倍", "行"}:
            return number * 240
        return number

    def _twips_to_style_value(self, key, twips, unit=None):
        unit = unit or self.style_unit_vars[key].get()
        divisor = {
            "厘米": 567,
            "毫米": 56.7,
            "磅": 20,
            "字符": self._font_size_points() * 20,
            "倍": 240,
            "行": 240,
        }.get(unit, 1)
        try:
            return f"{float(twips) / divisor:g}"
        except (TypeError, ValueError, ZeroDivisionError):
            return ""

    def _change_style_unit(self, key):
        old_unit = self._last_style_units.get(key, self.style_unit_vars[key].get())
        new_unit = self.style_unit_vars[key].get()
        twips = self._style_value_to_twips(key, unit=old_unit)
        self._last_style_units[key] = new_unit
        control = self.style_value_controls.get(key)
        if control is not None:
            control.configure(values=STYLE_VALUE_OPTIONS.get((key, new_unit), ()))
        self.style_vars[key].set(self._twips_to_style_value(key, twips, new_unit))

    def _set_style_measurement(self, key, raw_value, unit):
        self.style_unit_vars[key].set(unit)
        self._last_style_units[key] = unit
        control = self.style_value_controls.get(key)
        if control is not None:
            control.configure(values=STYLE_VALUE_OPTIONS.get((key, unit), ()))
        self.style_vars[key].set(
            "" if raw_value in (None, "") else self._twips_to_style_value(key, raw_value, unit)
        )

    def _update_color_swatch(self):
        if not hasattr(self, "color_swatch"):
            return
        raw = COLOR_CHOICES.get(
            self.style_vars["color"].get(),
            self.style_vars["color"].get(),
        ).strip().replace("#", "")
        color = f"#{raw}" if re.fullmatch(r"[0-9A-Fa-f]{6}", raw) else "#FFFFFF"
        self.color_swatch.configure(background=color)

    def load_style_form(self, event=None):
        selection = self.style_list.curselection()
        if not selection:
            return
        sid = self.visible_style_ids[selection[0]]
        self.current_style_id = sid
        item = self.style_by_id[sid]
        self.loaded_style_name = item.get("name", "")
        self.loaded_display_name = item.get("display_name") or chinese_style_name(item.get("name"), sid)
        self.style_vars["name"].set(self.loaded_display_name)
        self.style_vars["font"].set(item.get("font", ""))
        self.style_vars["size"].set(_display_value(item.get("size"), SIZE_CHOICES, "磅"))
        color_value = str(item.get("color") or "").strip().replace("#", "")
        self.style_vars["color"].set(
            next(
                (label for label, raw in COLOR_CHOICES.items() if raw == color_value),
                f"#{color_value}" if color_value else "",
            )
        )
        for key in ("before", "after"):
            self._set_style_measurement(key, item.get(key), "磅")
        line_unit = "磅" if item.get("line_rule") in {"exact", "atLeast"} else "倍"
        self._set_style_measurement("line", item.get("line"), line_unit)
        for key in ("left", "hanging"):
            self._set_style_measurement(key, item.get(key), "厘米")
        for key in ("based_on", "next", "link"):
            value = item.get(key, "")
            self.style_vars[key].set(self.style_id_to_label.get(value, "无" if not value else chinese_style_name("", value)))
        self.bold_var.set(bool(item.get("bold")))
        self.italic_var.set(bool(item.get("italic")))
        self.update_style_sample()

    def save_current_style(self):
        if not getattr(self, "current_style_id", ""):
            messagebox.showinfo("模板编辑器", "请先在左侧选择一个样式。", parent=self)
            return
        values = self._style_form_values()
        values["bold"] = self.bold_var.get()
        values["italic"] = self.italic_var.get()
        try:
            result = save_template_edits(
                self.template_path,
                {self.current_style_id: values},
                numbering_levels=None,
                overwrite=self.overwrite.get(),
            )
        except Exception as exc:
            messagebox.showerror("模板编辑器", str(exc), parent=self)
            return
        messagebox.showinfo(
            "模板编辑器",
            f"当前样式修改已保存：\n{result['output']}",
            parent=self,
        )
        if self.on_saved:
            self.on_saved(result["output"])
        self.destroy()

    def reset_current_style(self):
        if not getattr(self, "current_style_id", ""):
            messagebox.showinfo("模板编辑器", "请先在左侧选择一个样式。", parent=self)
            return
        self.style_updates.pop(self.current_style_id, None)
        self.load_style_form()

    def reset_numbering_edits(self):
        for (_frame, variables), initial in zip(
            self.level_rows,
            self.initial_level_state,
        ):
            for key, value in initial.items():
                variables[key].set(value)
        self.auto_link.set(self.initial_auto_link)
        self.auto_restart.set(self.initial_auto_restart)
        self.refresh_number_preview()

    def reset_all_edits(self):
        if not messagebox.askyesno(
            "重置全部修改",
            "将撤销本次打开模板编辑器后所做的全部样式和多级列表修改。是否继续？",
            parent=self,
        ):
            return
        self.style_updates.clear()
        for (_frame, variables), initial in zip(
            self.level_rows,
            self.initial_level_state,
        ):
            for key, value in initial.items():
                variables[key].set(value)
        self.auto_link.set(self.initial_auto_link)
        self.auto_restart.set(self.initial_auto_restart)
        self.overwrite.set(self.initial_overwrite)
        if getattr(self, "current_style_id", ""):
            self.load_style_form()
        self.refresh_number_preview()

    def _style_form_values(self):
        values = {key: variable.get() for key, variable in self.style_vars.items()}
        if values["name"] == getattr(self, "loaded_display_name", ""):
            values["name"] = getattr(self, "loaded_style_name", values["name"])
        values["size"] = _raw_value(values["size"], SIZE_CHOICES)
        values["color"] = COLOR_CHOICES.get(values["color"], values["color"]).replace("#", "")
        for key in ("before", "after", "line", "left", "hanging"):
            values[key] = (
                str(round(self._style_value_to_twips(key)))
                if self.style_vars[key].get().strip()
                else ""
            )
        values["line_rule"] = (
            "exact" if self.style_unit_vars["line"].get() == "磅" else "auto"
        ) if values["line"] else ""
        for key in ("based_on", "next", "link"):
            display = values[key]
            values[key] = "" if display == "无" else self.style_label_to_id.get(display, display)
        return values

    def update_style_sample(self):
        if not hasattr(self, "style_sample"):
            return
        font = self.style_vars["font"].get() or "Microsoft YaHei UI"
        try:
            raw_size = _raw_value(self.style_vars["size"].get(), SIZE_CHOICES)
            size = max(8, min(30, int(float(raw_size or 12))))
        except ValueError:
            size = 12
        color = COLOR_CHOICES.get(
            self.style_vars["color"].get(),
            self.style_vars["color"].get(),
        ).strip().replace("#", "")
        foreground = f"#{color}" if re.fullmatch(r"[0-9A-Fa-f]{6}", color) else COLORS["text"]

        def preview_pixels(key):
            return max(0, min(120, round(self._style_value_to_twips(key) / 20 * 4 / 3)))

        before = preview_pixels("before")
        after = preview_pixels("after")
        left = preview_pixels("left")
        hanging = preview_pixels("hanging")
        try:
            line_raw = self._style_value_to_twips("line") or 240
            line_spacing = max(0, min(24, round((line_raw / 240 - 1) * size * 4 / 3)))
        except (TypeError, ValueError):
            line_spacing = 0
        self.style_sample.configure(state=NORMAL)
        self.style_sample.delete("1.0", END)
        style_bits = []
        if self.bold_var.get():
            style_bits.append("bold")
        if self.italic_var.get():
            style_bits.append("italic")
        first_margin = max(0, left - hanging)
        self.style_preview_info.set(
            f"左缩进：{self.style_vars['left'].get() or '0'} {self.style_unit_vars['left'].get()}    "
            f"悬挂：{self.style_vars['hanging'].get() or '0'} {self.style_unit_vars['hanging'].get()}    "
            f"段前/段后：{self.style_vars['before'].get() or '0'} {self.style_unit_vars['before'].get()} / "
            f"{self.style_vars['after'].get() or '0'} {self.style_unit_vars['after'].get()}"
        )
        self.style_sample.tag_configure(
            "sample_first",
            font=(font, size, " ".join(style_bits) or "normal"),
            foreground=foreground,
            spacing1=before,
            spacing2=line_spacing,
            spacing3=after,
            lmargin1=first_margin,
            lmargin2=left,
        )
        self.style_sample.tag_configure(
            "sample_second",
            font=(font, size, " ".join(style_bits) or "normal"),
            foreground=foreground,
            spacing1=before,
            spacing2=line_spacing,
            spacing3=after,
            lmargin1=left,
            lmargin2=left,
        )
        name = self.style_vars["name"].get().strip() or "当前样式"
        self.style_sample.insert(
            "1.0",
            f"{name}：这是应用当前样式的长段落预览文字。"
            "请观察首行起点与自动换行后的正文起点差异，悬挂缩进越大，首行越靠左；"
            "左缩进越大，整段正文越靠右。",
            "sample_first",
        )
        self.style_sample.insert(
            END,
            "\n第二段用于观察段前、段后和行距变化。调整相关数值后，"
            "两段之间的垂直距离以及段内文字的疏密会立即变化。",
            "sample_second",
        )
        self.style_sample.configure(state=DISABLED)

    def toggle_more(self):
        self.show_more.set(not self.show_more.get())
        self._update_level_visibility()

    def _update_level_visibility(self):
        for index, (frame, _vars) in enumerate(self.level_rows):
            if index < 4 or self.show_more.get():
                frame.pack(fill=X, pady=1)
            else:
                frame.pack_forget()
        self.level_canvas.configure(height=210 if self.show_more.get() else 140)
        self.more_button.configure(
            text="收起 5-9 级 ▲" if self.show_more.get() else "展开并编辑 5-9 级 ▼"
        )

    def _scroll_levels(self, event):
        self.level_canvas.yview_scroll(-1 if event.delta > 0 else 1, "units")
        return "break"

    def _scroll_levels_horizontal(self, event):
        self.level_canvas.xview_scroll(-1 if event.delta > 0 else 1, "units")
        return "break"

    def level_data(self):
        result = []
        for _frame, vars_ in self.level_rows:
            item = {key: variable.get() for key, variable in vars_.items()}
            item["format"] = FORMAT_CHOICES.get(item["format"], item["format"])
            item["numFmt"] = NUMBER_FORMAT_VALUES.get(item["numFmt"], item["numFmt"])
            item["suffix"] = SUFFIX_VALUES.get(item["suffix"], item["suffix"])
            item["linked_style"] = self.style_label_to_id.get(item["linked_style"], item["linked_style"])
            item["start"] = _raw_value(item["start"], {})
            item["left"] = _raw_value(item["left"], INDENT_CHOICES, 567)
            item["hanging"] = _raw_value(item["hanging"], INDENT_CHOICES, 567)
            result.append(item)
        return result

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
                    if key == "format":
                        value = FORMAT_VALUES.get(value, value)
                    elif key == "numFmt":
                        value = NUMBER_FORMAT_LABELS.get(value, value)
                    elif key == "suffix":
                        value = SUFFIX_LABELS.get(value, value)
                    elif key == "linked_style":
                        value = self.style_id_to_label.get(value, chinese_style_name("", value))
                    elif key in {"left", "hanging"}:
                        value = _display_value(value, INDENT_CHOICES, "厘米", 567)
                    elif key == "start":
                        value = f"{value}（从 {value} 开始）"
                    vars_[key].set(value)
        self.refresh_number_preview()

    def refresh_number_preview(self):
        if not hasattr(self, "number_preview"):
            return
        samples = ["一级标题", "二级标题", "三级标题", "四级标题"]
        self.number_preview.configure(state=NORMAL)
        self.number_preview.delete("1.0", END)
        for index in range(4):
            variables = self.level_rows[index][1]
            fmt_display = variables["format"].get()
            fmt = FORMAT_CHOICES.get(fmt_display, fmt_display)
            for number in range(1, index + 2):
                level_variables = self.level_rows[number - 1][1]
                level_format = NUMBER_FORMAT_VALUES.get(
                    level_variables["numFmt"].get(),
                    level_variables["numFmt"].get(),
                )
                level_start = _raw_value(level_variables["start"].get(), {}) or "1"
                fmt = fmt.replace(
                    f"%{number}",
                    _format_number(level_start, level_format),
                )
            suffix_value = SUFFIX_VALUES.get(variables["suffix"].get(), variables["suffix"].get())
            suffix = {"space": " ", "tab": "    ", "nothing": ""}.get(suffix_value, " ")
            try:
                left = max(0, round(float(_raw_value(variables["left"].get(), INDENT_CHOICES, 567)) / 15))
                hanging = max(0, round(float(_raw_value(variables["hanging"].get(), INDENT_CHOICES, 567)) / 15))
            except (TypeError, ValueError):
                left = hanging = 0
            tag = f"level_{index}"
            self.number_preview.tag_configure(
                tag,
                lmargin1=max(0, left - hanging),
                lmargin2=left,
                spacing1=5,
                spacing3=7,
            )
            self.number_preview.insert(
                END,
                f"{fmt}{suffix}{samples[index]}的较长示例文字，用于观察换行后正文与编号之间的对齐位置。"
                f"（左缩进 {variables['left'].get()}，悬挂 {variables['hanging'].get()}）"
                + ("\n" if index < 3 else ""),
                tag,
            )
        self.number_preview.configure(state=DISABLED)

    def save(self):
        if getattr(self, "current_style_id", ""):
            values = self._style_form_values()
            values["bold"] = self.bold_var.get()
            values["italic"] = self.italic_var.get()
            self.style_updates[self.current_style_id] = values
        try:
            result = save_template_edits(
                self.template_path,
                self.style_updates,
                self.level_data() if self.numbering_editor_enabled else None,
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
        candidates = clean_old_template_versions(self.library_dir, base, keep, dry_run=True)
        if not candidates:
            messagebox.showinfo("清理旧版本", "没有需要清理的旧版本。原始模板不会被删除。", parent=self)
            return
        names = "\n".join(f"- {Path(path).name}" for path in candidates)
        if not messagebox.askyesno(
            "确认清理旧版本",
            f"将永久删除以下 {len(candidates)} 个版本文件：\n\n{names}\n\n原始模板不会被删除。是否继续？",
            parent=self,
        ):
            return
        removed = clean_old_template_versions(self.library_dir, base, keep)
        messagebox.showinfo("清理旧版本", f"已清理 {len(removed)} 个旧版本。", parent=self)
