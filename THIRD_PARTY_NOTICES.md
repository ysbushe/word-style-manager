# 第三方开源软件说明

本程序自身采用 MIT License，详见 `LICENSE`。

本程序使用或随打包结果分发以下开源组件。下表是便于用户识别的摘要，具体权利与义务以各组件随安装包或源代码发布的原始许可证文件为准。

| 组件 | 用途 | 许可证 |
| --- | --- | --- |
| Python | 程序运行环境 | Python Software Foundation License |
| python-docx | 读取和处理 Word 文档 | MIT |
| lxml | 解析和写入 OOXML/XML | BSD-3-Clause；发行包还包含其依赖库许可证 |
| ttkbootstrap | Tkinter 界面主题和控件 | MIT AND (Apache-2.0 OR BSD-2-Clause) |
| tkinterdnd2 / tkDnD | 文件拖放支持 | Python 包为 MIT；底层 tkDnD 以其随包许可证为准 |
| pywin32 | Word/WPS COM 调用 | 包含多种许可证，必须保留其发行包内许可证和版权声明 |
| Pillow | ttkbootstrap 的图像处理依赖 | MIT-CMU；发行包可能包含第三方库许可证 |
| typing_extensions | python-docx 的兼容性依赖 | PSF-2.0 |
| PyInstaller | Windows 绿色版构建工具 | GPL-2.0-or-later，带 Bootloader Exception |

## 发布要求

1. 发布源码时保留本项目 `LICENSE`。
2. 发布绿色版时同时携带本文件和 `LICENSE`。
3. 打包依赖升级后，应重新核对实际安装版本及其许可证文件。
4. 对 `pywin32`、`lxml`、Pillow、tkDnD 等包含附加组件的依赖，应以发行包中附带的许可证文件为准，不应笼统标注为 MIT。

## 上游项目

- https://www.python.org/
- https://github.com/python-openxml/python-docx
- https://github.com/lxml/lxml
- https://github.com/israel-dryer/ttkbootstrap
- https://github.com/Eliav2/tkinterdnd2
- https://github.com/mhammond/pywin32
- https://github.com/python-pillow/Pillow
- https://github.com/pyinstaller/pyinstaller
