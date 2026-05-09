#!/usr/bin/env bash
# obsidian-finance-bookkeeping-workflow — 一键安装 Hermes Agent Skill
set -euo pipefail

REPO_OWNER="${REPO_OWNER:-chycici}"
REPO_NAME="${REPO_NAME:-obsidian-finance-bookkeeping-workflow}"
BRANCH="${BRANCH:-main}"
SKILL_NAME="obsidian-finance-bookkeeping-workflow"
HERMES_HOME="${HERMES_HOME:-$HOME/.hermes}"
SKILL_DIR="$HERMES_HOME/skills/$SKILL_NAME"

echo "📦 安装 $SKILL_NAME ..."
mkdir -p "$SKILL_DIR/scripts"

BASE_URL="https://raw.githubusercontent.com/$REPO_OWNER/$REPO_NAME/$BRANCH/skills/$SKILL_NAME"

echo "  下载 SKILL.md ..."
curl -fsSL "$BASE_URL/SKILL.md" -o "$SKILL_DIR/SKILL.md"

echo "  下载 DESIGN.md ..."
curl -fsSL "$BASE_URL/DESIGN.md" -o "$SKILL_DIR/DESIGN.md" 2>/dev/null || true

echo "  下载脚本 ..."
for s in create_entry export_excel update_monthly validate_links; do
  curl -fsSL "$BASE_URL/scripts/$s.py" -o "$SKILL_DIR/scripts/$s.py"
  chmod +x "$SKILL_DIR/scripts/$s.py"
done

echo "  安装 Python 依赖 ..."
pip3 install openpyxl python-frontmatter --quiet 2>/dev/null || true

# 完整性校验
if [ -f "$SKILL_DIR/SKILL.md" ] && [ -f "$SKILL_DIR/scripts/create_entry.py" ]; then
  echo "✅ $SKILL_NAME 安装完成 → $SKILL_DIR"
  echo ""
  echo "⚠️  请配置环境变量:"
  echo "   export OBSIDIAN_VAULT_ROOT=\"/path/to/your-obsidian-vault\""
else
  echo "❌ 下载失败，请检查网络连接"
  exit 1
fi
