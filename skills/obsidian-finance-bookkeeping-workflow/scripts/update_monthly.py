#!/usr/bin/env python3
"""读取指定月份的所有条目文件，重新计算统计数据，更新月度页的汇总区块和明细区块。

用法：
  python update_monthly.py --vault /path/to/vault --month 2026-05
"""

import argparse
import json
import os
import re
import sys
from pathlib import Path
from collections import defaultdict

try:
    import frontmatter
except ImportError:
    print(json.dumps({
        "status": "error",
        "error_code": "MISSING_DEPENDENCY",
        "message": "缺少 python-frontmatter 依赖，请执行: pip install python-frontmatter",
        "recoverable": False,
    }, ensure_ascii=False))
    sys.exit(1)

PAYER_GROUPS_ORDER = [
    "北京盛通佳达科技有限公司",
    "成都汇智迅析科技有限公司",
    "海南云创软通科技有限公司",
    "个人",
]
VALID_PAYERS = set(PAYER_GROUPS_ORDER)
INVOICE_STATUSES = ["有发票", "有收据", "无发票", "待确认"]


def error(errcode, message, recoverable=False):
    print(json.dumps({
        "status": "error",
        "error_code": errcode,
        "message": message,
        "recoverable": recoverable,
    }, ensure_ascii=False))
    sys.exit(0 if recoverable else 1)


def success(data):
    data["status"] = "success"
    print(json.dumps(data, ensure_ascii=False))


def format_amount(value):
    """数值千分位格式化，保留两位小数。"""
    if value is None:
        return "0.00"
    return f"{float(value):,.2f}"


def parse_entry(filepath):
    """解析单个条目文件，返回 dict 或 None（有警告时跳过）。"""
    try:
        post = frontmatter.load(filepath)
    except Exception:
        print(f"# WARNING: 无法解析文件 {filepath}", file=sys.stderr)
        return None

    fm = post.metadata

    try:
        amount_cny = float(fm.get("amount_cny", 0) or 0)
    except (ValueError, TypeError):
        print(f"# WARNING: amount_cny 字段为空或非数字，跳过：{filepath}", file=sys.stderr)
        return None

    payer = fm.get("payer", "个人")
    if payer not in VALID_PAYERS:
        payer = "个人"

    return {
        "path": str(filepath.relative_to(filepath.parents[3])),  # relative to vault root
        "date": str(fm.get("date", "")),
        "title": str(fm.get("title", "")),
        "category": str(fm.get("category", "其他")),
        "reimbursement_type": str(fm.get("reimbursement_type", "日常")),
        "amount": float(fm.get("amount", 0) or 0),
        "currency": str(fm.get("currency", "CNY")),
        "amount_cny": amount_cny,
        "invoice_status": str(fm.get("invoice_status", "")),
        "pay_method": str(fm.get("pay_method", "")),
        "payer": payer,
        "payee": str(fm.get("payee", "")),
        "location": str(fm.get("location", "")),
        "project": str(fm.get("project", "")),
        "trip_name": str(fm.get("trip_name", "")),
        "status": str(fm.get("status", "已记录")),
        "voucher": str(fm.get("voucher") or ""),
    }


def group_entries(entries):
    """按 payer 分组，每组内按 date 升序排列。"""
    groups = defaultdict(list)
    for e in entries:
        groups[e["payer"]].append(e)
    for g in groups:
        groups[g].sort(key=lambda x: x["date"])
    return groups


def build_summary_table(stats):
    """构造 ## 汇总 区块内容。"""
    lines = [
        "## 汇总",
        "",
        "| 指标 | 数值 |",
        "|------|------|",
        f"| 条目总数 | {stats['total_entries']} |",
        f"| 总金额（CNY） | {format_amount(stats['total_cny'])} |",
    ]
    for status in INVOICE_STATUSES:
        label = f"{status}金额"
        amount = stats["by_invoice_status"].get(status, {}).get("amount", 0)
        lines.append(f"| {label} | {format_amount(amount)} |")
    return "\n".join(lines)


def build_pending_section(pending_entries):
    """构造 ## 待确认条目 区块内容。"""
    lines = ["## 待确认条目", ""]
    if pending_entries:
        for path in pending_entries:
            lines.append(f"- [[{path}]]")
    else:
        lines.append("（无）")
    return "\n".join(lines)


def build_detail_table(group_name, entries):
    """构造单个抬头组的明细表。"""
    lines = [
        f"### {group_name}",
        "",
        "| 序号 | 日期 | 摘要 | 类别 | 金额（CNY） | 发票状态 | 凭证 |",
        "|------|------|------|------|------------|---------|------|",
    ]

    total_cny = 0
    for i, e in enumerate(entries, 1):
        nn = f"{i:02d}"
        date_str = e["date"]
        # 摘要从 title 提取（去掉日期前缀）
        summary = e["title"]
        if re.match(r"^\d{4}-\d{2}-\d{2}\s+", summary):
            summary = re.sub(r"^\d{4}-\d{2}-\d{2}\s+", "", summary)
        category = e["category"]
        amount_str = format_amount(e["amount_cny"])
        invoice = e["invoice_status"]
        entry_link = e["path"][:-3] if e["path"].endswith(".md") else e["path"]  # 去掉 .md 后缀用于 Wiki 链接
        total_cny += e["amount_cny"]

        lines.append(
            f"| {nn} | {date_str} | {summary} | {category} | {amount_str} | {invoice} | [[{entry_link}]] |"
        )

    lines.append("")
    lines.append(f"**小计：{format_amount(total_cny)}**")
    return "\n".join(lines), total_cny


def find_block_boundaries(lines, start_marker):
    """在 Markdown 行列表中查找某个 ## 标题区块的起止行号。
    返回 (start_line, end_line)，其中 start_line 是标题行，end_line 是下一个 ## 标题的前一行（或文件末尾）。
    """
    heading_pattern = re.compile(r"^##\s+")
    start_idx = None
    for i, line in enumerate(lines):
        stripped = line.strip()
        if stripped == start_marker:
            start_idx = i
        elif heading_pattern.match(stripped) and start_idx is not None:
            return start_idx, i - 1
    if start_idx is not None:
        return start_idx, len(lines) - 1
    return None, None


def replace_section(full_lines, heading_text, new_content):
    """替换 Markdown 中以 heading_text 开头的区块。如果不存在则追加到末尾。"""
    start, end = find_block_boundaries(full_lines, heading_text)
    if start is not None:
        # 保留标题前的空行
        while start > 0 and full_lines[start - 1].strip() == "":
            start -= 1
        # 删除旧区块
        del full_lines[start:end + 1]
        # 插入新内容
        new_lines = new_content.split("\n")
        full_lines[start:start] = new_lines + [""]
    else:
        # 追加到文件末尾
        full_lines.append("")
        full_lines.extend(new_content.split("\n"))
    return full_lines


def build_monthly_page(vault_root, month, entries):
    """构造完整的月度页内容。"""
    stats = compute_stats(entries)
    groups = group_entries(entries)
    pending_entries = [e["path"] for e in entries if e["status"] == "待确认"]

    summary_block = build_summary_table(stats)
    pending_block = build_pending_section(pending_entries)

    # 构建明细区块
    detail_lines = ["## 明细", ""]
    by_payer_stats = {}
    for payer_name in PAYER_GROUPS_ORDER:
        group_entries_list = groups.get(payer_name, [])
        if group_entries_list:
            table, total = build_detail_table(payer_name, group_entries_list)
            detail_lines.append(table)
            detail_lines.append("")
        else:
            detail_lines.append(f"### {payer_name}")
            detail_lines.append("")
            detail_lines.append("（本月无此抬头条目）")
            detail_lines.append("")
            total = 0
        by_payer_stats[payer_name] = {"count": len(group_entries_list), "total_cny": total}

    detail_block = "\n".join(detail_lines)

    # 检查是否有出差条目
    has_trip = any(e["reimbursement_type"] == "出差" for e in entries)

    # 读取现有月度页或创建新的
    monthly_path = vault_root / "财务记账" / "按月" / f"{month}.md"
    if monthly_path.exists():
        original = monthly_path.read_text(encoding="utf-8")
        full_lines = original.split("\n")
    else:
        # 创建新月度页
        full_lines = [f"# {month[:4]}年{month[5:]}月 财务记录", ""]

    # 替换各区块
    full_lines = replace_section(full_lines, "## 汇总", summary_block)
    full_lines = replace_section(full_lines, "## 待确认条目", pending_block)
    full_lines = replace_section(full_lines, "## 明细", detail_block)

    # 处理出差费用区块
    trip_heading = "## 出差费用"
    start, end = find_block_boundaries(full_lines, trip_heading)
    if start is not None:
        while start > 0 and full_lines[start - 1].strip() == "":
            start -= 1
        del full_lines[start:end + 1]

    if has_trip:
        trip_paths = set()
        for e in entries:
            if e["reimbursement_type"] == "出差" and e["trip_name"]:
                trip_paths.add(f"财务记账/出差任务/{month[:7]}-{e['trip_name']}")
        trip_block = "\n".join([
            trip_heading,
            "",
        ] + [f'出差费用详见 [[{p}]]' for p in sorted(trip_paths)])
        full_lines.append("")
        full_lines.extend(trip_block.split("\n"))

    content = "\n".join(full_lines).strip() + "\n"
    monthly_path.parent.mkdir(parents=True, exist_ok=True)
    monthly_path.write_text(content, encoding="utf-8")

    return stats, by_payer_stats, pending_entries


def compute_stats(entries):
    """计算汇总统计数据。"""
    total_entries = len(entries)
    total_cny = sum(e["amount_cny"] for e in entries)

    by_invoice_status = {}
    checksum = 0
    for status in INVOICE_STATUSES:
        matching = [e for e in entries if e["invoice_status"] == status]
        amount = sum(e["amount_cny"] for e in matching)
        by_invoice_status[status] = {"count": len(matching), "amount": amount}
        checksum += amount

    checksum_ok = abs(checksum - total_cny) < 0.015

    return {
        "total_entries": total_entries,
        "total_cny": total_cny,
        "by_invoice_status": by_invoice_status,
        "checksum_ok": checksum_ok,
    }


def main():
    parser = argparse.ArgumentParser(description="更新月度页统计")
    parser.add_argument("--vault", default=os.environ.get("OBSIDIAN_VAULT_ROOT"),
                        help="Vault 根目录路径")
    parser.add_argument("--month", required=True, help="月份，格式 YYYY-MM")
    args = parser.parse_args()

    if not args.vault:
        error("VAULT_NOT_SET", "未指定 vault 路径")

    vault_root = Path(args.vault).expanduser().resolve()
    if not vault_root.exists():
        error("VAULT_NOT_FOUND", f"Vault 目录不存在：{vault_root}")

    month = args.month
    entries_dir = vault_root / "财务记账" / "条目" / month

    if not entries_dir.exists():
        error("ENTRIES_DIR_NOT_FOUND",
              f"条目目录不存在：{entries_dir}。请先创建条目或确认月份是否正确。",
              recoverable=True)

    # 扫描所有条目文件
    entry_files = sorted(entries_dir.glob("*.md"))
    if not entry_files:
        print(json.dumps({
            "month": month,
            "total_entries": 0,
            "total_cny": 0,
            "by_invoice_status": {s: {"count": 0, "amount": 0} for s in INVOICE_STATUSES},
            "pending_entries": [],
            "checksum_ok": True,
            "status": "success",
            "note": "当月无条目文件",
        }, ensure_ascii=False))
        return

    entries = []
    warnings = []
    for fp in entry_files:
        entry = parse_entry(fp)
        if entry:
            entries.append(entry)
        else:
            warnings.append(str(fp))

    stats = compute_stats(entries)
    if not stats["checksum_ok"]:
        error("CHECKSUM_FAILED",
              f"分类金额之和 ({sum(stats['by_invoice_status'][s]['amount'] for s in INVOICE_STATUSES)}) "
              f"≠ 总金额 ({stats['total_cny']})，误差超过 0.01，不写入月度页。")

    stats_out, by_payer, pending = build_monthly_page(vault_root, month, entries)

    success({
        "month": month,
        "total_entries": stats_out["total_entries"],
        "total_cny": round(stats_out["total_cny"], 2),
        "by_invoice_status": {
            s: {"count": stats_out["by_invoice_status"][s]["count"],
                "amount": round(stats_out["by_invoice_status"][s]["amount"], 2)}
            for s in INVOICE_STATUSES
        },
        "by_payer": {
            p: {"count": by_payer[p]["count"], "total_cny": round(by_payer[p]["total_cny"], 2)}
            for p in PAYER_GROUPS_ORDER
        },
        "pending_entries": pending,
        "checksum_ok": stats_out["checksum_ok"],
        "warnings": warnings if warnings else None,
    })


if __name__ == "__main__":
    main()
