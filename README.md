# obsidian-finance-bookkeeping-workflow — 中文财务记账/报销知识库工作流

在 Obsidian 中按月、按任务整理财务信息，AI 辅助录入、汇总、导出。保留原始凭证，归档防重复报销。

> **⚠️ 这是 Hermes Agent 的记账 skill，不是独立的 Obsidian 插件。**  
> 需要配合 [Hermes Agent](https://github.com/chycici/hermes-agent) 使用。

## 功能

- 📄 **单笔 PDF 入账** — 解析发票/收据 PDF，自动提取日期、供应商、金额、抬头
- 📊 **批量 Excel 导入** — 从原始 Excel 按模板导入多笔记录
- 📑 **按抬头导出** — 按公司/个人抬头分组导出 Excel 汇总表
- 🔗 **凭证校验** — 自动检查 frontmatter 与正文图片引用一致性
- 📈 **月度统计** — 自动更新月度汇总页（有发票/收据/无发票/待确认）
- 🔄 **口径变更** — 移动条目抬头分组，同步更新全部关联文件

## 安装

```bash
# 一键安装
bash <(curl -fsSL https://raw.githubusercontent.com/chycici/obsidian-finance-bookkeeping-workflow/main/install.sh)

# 或手动复制
git clone --depth=1 https://github.com/chycici/obsidian-finance-bookkeeping-workflow.git /tmp/repo
cp -r /tmp/repo/skills/obsidian-finance-bookkeeping-workflow ~/.hermes/skills/
rm -rf /tmp/repo
```

安装后需要设置环境变量（指向你的 Obsidian 保险库路径）：
```bash
export OBSIDIAN_VAULT_ROOT="/path/to/your-obsidian-vault"
```

重启 Hermes Agent 或在会话中使用 `/refresh` 即可加载。

## 前置依赖

- Python 3.9+
- Obsidian 保险库
- pip: `openpyxl`, `python-frontmatter`

## 文件结构

```
skills/obsidian-finance-bookkeeping-workflow/
├── SKILL.md                  # 技能入口（核心工作流）
├── DESIGN.md                 # 设计文档（业务规则权威定义）
└── scripts/
    ├── create_entry.py       # 单笔入账
    ├── export_excel.py       # 按抬头导出Excel
    ├── update_monthly.py     # 月度统计更新
    └── validate_links.py     # 凭证链接校验
```

## 数据来源

- 用户提供的发票/收据 PDF
- 财务 Excel 原始数据
- 实时汇率（助手自动补查）

## License

MIT
