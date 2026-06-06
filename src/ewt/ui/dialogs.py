"""弹窗组件"""

import os
import tkinter as tk
from tkinter import LEFT, RIGHT, BOTH, X, W

import ttkbootstrap as tb
from ttkbootstrap.constants import *


def open_folder(path, select_file=False):
    """在资源管理器中打开目录或选中文件"""
    if not path:
        return
    path = os.path.normpath(path)
    if select_file and os.path.exists(path):
        import subprocess
        subprocess.Popen(f'explorer /select,"{path}"')
    elif os.path.exists(path):
        os.startfile(path)


class FinishDialog(tb.Toplevel):
    """处理完成弹窗，含打开结果/打开位置/关闭按钮"""

    def __init__(self, parent, title, message, paths=None):
        super().__init__(parent)
        self.title(title)
        self.geometry("460x230")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        self.paths = paths or []

        body = tb.Frame(self, padding=22)
        body.pack(fill=BOTH, expand=True)
        tb.Label(body, text=title, font=("Microsoft YaHei UI", 16, "bold")).pack(anchor=W)
        tb.Label(body, text=message, wraplength=400, justify=LEFT, bootstyle=SECONDARY).pack(anchor=W, pady=(12, 18))

        btns = tb.Frame(body)
        btns.pack(fill=X, side=BOTTOM)
        if self.paths:
            tb.Button(btns, text="打开第一个结果", bootstyle=SUCCESS, command=self.open_first).pack(side=LEFT)
            tb.Button(btns, text="打开所在位置", bootstyle=SECONDARY, command=self.open_location).pack(side=LEFT, padx=8)
        tb.Button(btns, text="关闭", bootstyle=LIGHT, command=self.destroy).pack(side=RIGHT)

        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - self.winfo_width()) // 2
        y = parent.winfo_y() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{x}+{y}")

    def open_first(self):
        if self.paths and os.path.exists(self.paths[0]):
            os.startfile(self.paths[0])
        self.destroy()

    def open_location(self):
        if self.paths:
            open_folder(self.paths[0], select_file=True)
        self.destroy()
