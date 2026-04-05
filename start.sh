#!/bin/bash
echo "====================================="
echo "办公文件HTML可视化转换工具 - 启动中"
echo "====================================="

# 检查操作系统
OS=$(uname)
echo "检测操作系统: $OS"

# -----------------------
# 安装 Python 依赖
# -----------------------
echo "1. 安装 Python 依赖..."
pip3 install --upgrade pip > /dev/null 2>&1
pip3 install flask pandas pdf2docx python-docx weasyprint markdown > /dev/null 2>&1

# -----------------------
# 安装 LibreOffice (仅 Linux)
# -----------------------
if [[ "$OS" == "Linux" ]]; then
    echo "2. 安装 LibreOffice..."
    sudo apt update -y > /dev/null 2>&1
    sudo apt install libreoffice -y > /dev/null 2>&1
    echo "LibreOffice 安装完成"
else
    echo "⚠️ Windows 系统请确保已安装 LibreOffice，并将路径添加到系统环境变量"
fi

# -----------------------
# 创建文件夹
# -----------------------
mkdir -p uploads outputs

# -----------------------
# 启动 Flask 服务
# -----------------------
echo "3. 启动后端服务..."
echo "请在浏览器访问: http://127.0.0.1:5000"
python3 app.py
