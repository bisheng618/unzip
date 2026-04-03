#!/bin/bash

# 启动递归解压GUI程序

# 设置Python环境
PYTHON="python3"

# 检查Python是否安装
if ! command -v $PYTHON &> /dev/null; then
    echo "错误: 未找到Python3，请先安装Python3"
    exit 1
fi

# 检查PyQt5是否安装
if ! $PYTHON -c "import PyQt5" &> /dev/null; then
    echo "警告: 未找到PyQt5模块，正在尝试安装..."
    $PYTHON -m pip install PyQt5
    if [ $? -ne 0 ]; then
        echo "错误: 安装PyQt5失败，请手动安装: pip install PyQt5"
        exit 1
    fi
fi

# 检查openpyxl是否安装
if ! $PYTHON -c "import openpyxl" &> /dev/null; then
    echo "警告: 未找到openpyxl模块，正在尝试安装..."
    $PYTHON -m pip install openpyxl
    if [ $? -ne 0 ]; then
        echo "错误: 安装openpyxl失败，请手动安装: pip install openpyxl"
        exit 1
    fi
fi

# 检查unar是否安装（用于解压RAR文件）
if ! command -v unar &> /dev/null; then
    echo "警告: 未找到unar命令，RAR文件解压功能可能无法使用"
    echo "macOS用户请使用: brew install unar"
    echo "Ubuntu/Debian用户请使用: sudo apt-get install unar"
fi

# 检查LibreOffice是否安装（用于转换.doc文件）
if [ ! -f "/Applications/LibreOffice.app/Contents/MacOS/soffice" ] && ! command -v libreoffice &> /dev/null; then
    echo "警告: 未找到LibreOffice，.doc文件转换功能可能无法使用"
    echo "请安装LibreOffice以启用.doc转.docx功能"
fi

# 启动程序
echo "启动递归解压GUI程序..."
$PYTHON "$(dirname "$0")/recursive_unzip_gui.py"
