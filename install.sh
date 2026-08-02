#!/bin/sh
# ============================================================
# code-scenario-testcases · 安装脚本
# 将本仓库（skill 源码）部署到 Claude Code 运行时目录，
# 使最新版本立即生效。
#
# 触发方式：
#   1. 手动：./install.sh
#   2. 自动：每次 git commit 后由 .git/hooks/post-commit 调用
#
# 部署内容：SKILL.md / README.md / references/*.md / scripts/*.py
# 保留内容：运行时目录下的 testcase/（历史用例产出，不入库、不覆盖）
# ============================================================

set -e

# 源目录 = 本脚本所在目录（仓库根）
SRC_DIR=$(cd "$(dirname "$0")" && pwd)

# 目标目录 = Claude Code 运行时 skill 目录（可用环境变量 SKILL_DST 覆盖）
if [ -n "$SKILL_DST" ]; then
    DST_DIR="$SKILL_DST"
else
    DST_DIR="$HOME/.claude/skills/code-scenario-testcases"
fi

# 校验源目录（防止在错误位置运行）
if [ ! -f "$SRC_DIR/SKILL.md" ]; then
    echo "install: 未找到 SKILL.md，请在仓库根目录运行（当前: $SRC_DIR）" >&2
    exit 1
fi

mkdir -p "$DST_DIR/references" "$DST_DIR/scripts"

cp "$SRC_DIR/SKILL.md"        "$DST_DIR/SKILL.md"
cp "$SRC_DIR/README.md"       "$DST_DIR/README.md"
cp "$SRC_DIR/references/"*.md "$DST_DIR/references/"
cp "$SRC_DIR/scripts/"*.py    "$DST_DIR/scripts/"

echo "install: skill 已部署到 $DST_DIR"
echo "         （SKILL.md / README.md / references/*.md / scripts/*.py；testcase/ 保留）"
