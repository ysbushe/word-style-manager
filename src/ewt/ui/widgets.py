"""复用 UI 组件"""

import tkinter as tk
from tkinter import LEFT, X, filedialog

import ttkbootstrap as tb
from ttkbootstrap.constants import *


class PathRow(tb.Frame):
    """文件/目录选择行：标签 + 输入框 + 选择按钮"""

    def __init__(self, parent, label, variable, filetypes, mode="open", command=None):
        super().__init__(parent)
        self.variable = variable
        self.filetypes = filetypes
        self.mode = mode
        self.command = command or self.pick
        tb.Label(self, text=label, width=12, anchor=W).pack(side=LEFT)
        tb.Entry(self, textvariable=variable).pack(side=LEFT, fill=X, expand=True, padx=8)
        tb.Button(self, text="选择", bootstyle=SECONDARY, command=self.command).pack(side=LEFT)

    def pick(self):
        if self.mode == "dir":
            path = filedialog.askdirectory()
        elif self.mode == "save":
            path = filedialog.asksaveasfilename(filetypes=self.filetypes)
        else:
            path = filedialog.askopenfilename(filetypes=self.filetypes)
        if path:
            self.variable.set(path)
