#!/bin/bash
# ============================================================
#  巡检数据处理工具 - 打包脚本
#  用法: bash build.sh
# ============================================================

set -e  # 出错即停

# ============ 配置区 ============
SCRIPT_FILE="mergeExcel.py"
OUTPUT_NAME="巡检数据处理工具"
ICON_FILE="images/favicon.ico"
BUILD_DIR="venv"

# 颜色定义
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

# ============ 函数 ============
print_step() {
    echo -e "\n${GREEN}[✓] $1${NC}"
}

print_warn() {
    echo -e "${YELLOW}[!] $1${NC}"
}

print_error() {
    echo -e "${RED}[✗] $1${NC}"
}

# ============ 主流程 ============
echo "=================================================="
echo "  巡检数据处理工具 - 打包脚本"
echo "=================================================="

# 1. 检查 Python 版本
print_step "检查 Python 环境"
python3 --version
if [ $? -ne 0 ]; then
    print_error "未找到 Python3，请先安装"
    exit 1
fi

# 2. 检查入口文件
if [ ! -f "$SCRIPT_FILE" ]; then
    print_error "找不到入口文件: $SCRIPT_FILE"
    exit 1
fi

# 3. 检查图标文件
if [ ! -f "$ICON_FILE" ]; then
    print_warn "找不到图标文件: $ICON_FILE，将使用默认图标"
    ICON_PARAM=""
else
    ICON_PARAM="-i $ICON_FILE"
fi

# 4. 创建虚拟环境
if [ ! -d "$BUILD_DIR" ]; then
    print_step "创建虚拟环境: $BUILD_DIR"
    python3 -m venv $BUILD_DIR
else
    print_warn "虚拟环境已存在，跳过创建"
fi

# 5. 激活虚拟环境
print_step "激活虚拟环境"
source $BUILD_DIR/bin/activate

# 6. 安装依赖
print_step "安装必要依赖"
pip install --upgrade pip -q
pip install -q \
    openpyxl \
    paramiko \
    pythonping \
    PyYAML \
    alive-progress \
    pyinstaller

# 7. 清理旧的构建文件
print_step "清理旧的构建文件"
rm -rf build/ dist/ __pycache__/
rm -f *.spec

# 8. 打包（参考你的命令，添加排除大包）
print_step "开始打包"
pyinstaller \
    -F \
    $ICON_PARAM \
    -n "$OUTPUT_NAME" \
    --collect-all grapheme \
    --exclude-module pandas \
    --exclude-module numpy \
    --exclude-module matplotlib \
    --exclude-module scipy \
    --exclude-module PyQt5 \
    --exclude-module PyQt6 \
    --exclude-module tkinter \
    --exclude-module test \
    --exclude-module unittest \
    --exclude-module email \
    --exclude-module xml \
    --exclude-module pydoc \
    --exclude-module doctest \
    --clean \
    $SCRIPT_FILE

# 9. 检查结果
if [ -f "dist/$OUTPUT_NAME" ] || [ -f "dist/$OUTPUT_NAME.exe" ]; then
    print_step "打包成功！"
    echo ""
    echo "  输出文件: dist/"
    ls -lh dist/
    echo ""
    
    # 计算大小
    if [ -f "dist/$OUTPUT_NAME" ]; then
        SIZE=$(du -h "dist/$OUTPUT_NAME" | cut -f1)
    else
        SIZE=$(du -h "dist/$OUTPUT_NAME.exe" | cut -f1)
    fi
    echo "  文件大小: $SIZE"
    echo ""
    echo "=================================================="
    echo "  打包完成！"
    echo "=================================================="
else
    print_error "打包失败，请检查错误信息"
    exit 1
fi
