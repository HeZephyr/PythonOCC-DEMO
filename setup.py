# setup.py
import sys
import os
from cx_Freeze import setup, Executable

# 根据不同平台设置基本参数
base = None
if sys.platform == "win32":
    base = "Win32GUI"  # 如果是GUI应用，去掉控制台窗口
    # 如果是控制台应用，使用 "Console"

executables = [
    Executable(
        "visualize_xml.py",        # 要打包的Python脚本
        base=base,
        target_name="visualize_xml",  # 输出的可执行文件名
        # icon="icon.ico",            # 可选：应用图标(如果有)
    )
]

# 设置构建选项
build_options = {
    "build_exe": "dist/visualize_xml",                   # 输出目录
}

setup(
    name="VisualizeXML",
    version="1.0",
    description="Excel Visualization Tool",
    options={"build_exe": build_options},
    executables=executables
)