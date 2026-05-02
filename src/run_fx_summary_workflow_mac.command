#!/bin/zsh
set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"
VENV_PYTHON="$REPO_ROOT/venv/bin/python3"

if [ -x "$VENV_PYTHON" ]; then
  PYTHON_CMD="$VENV_PYTHON"
elif command -v python3 >/dev/null 2>&1; then
  PYTHON_CMD="python3"
else
  echo "未找到可用的 Python 3。"
  echo "请先准备项目 venv，或安装 Python 3 后重试。"
  read -k 1 "?按任意键退出..."
  echo
  exit 1
fi

"$PYTHON_CMD" "$REPO_ROOT/src/fx_summary_workflow/fx_summary_workflow_app.py"

if [ $? -ne 0 ]; then
  echo
  echo "程序未成功完成，请查看界面提示或 result 文件夹中的日志。"
  read -k 1 "?按任意键退出..."
  echo
fi
