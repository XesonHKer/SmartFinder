#!/bin/bash

# 打包 SmartFinder v8 成 macOS app

echo "开始打包 SmartFinder_v8..."

PYTHON_PATH="/Library/Frameworks/Python.framework/Versions/3.11/bin/python3"

SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$SCRIPT_DIR"

echo "清理之前的打包文件..."
rm -rf build dist

echo "运行 PyInstaller 打包 (使用 SmartFinder.spec)..."
$PYTHON_PATH -m PyInstaller --noconfirm SmartFinder.spec

if [ -d "dist/SmartFinder.app" ]; then
    echo "✓ 打包成功！"
    echo "应用位置: $SCRIPT_DIR/dist/SmartFinder.app"
    touch "dist/SmartFinder.app"
    echo ""
    echo "打包完成！"
    echo "- 应用文件: $SCRIPT_DIR/dist/SmartFinder.app"
else
    echo "✗ 打包失败，请检查错误信息"
    exit 1
fi
