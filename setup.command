#!/bin/bash
cd "$(dirname "$0")"
echo "====================================="
echo "  自动评论机 — 初次安装"
echo "====================================="
echo ""
echo "正在检查 Python 环境..."
python3 --version || { echo "请先安装 Python 3"; exit 1; }
echo ""
echo "正在安装依赖..."
pip3 install openpyxl requests playwright --quiet
echo ""
echo "正在安装 Playwright 浏览器..."
python3 -m playwright install chromium
echo ""
echo "====================================="
echo "  安装完成！"
echo "====================================="
echo ""
echo "下次双击「自动评论机.command」即可启动。"
echo ""
read -p "按回车键退出..."
