#!/usr/bin/env python3
"""创建一条完整的条目笔记（frontmatter + 正文），并将凭证文件复制到正确位置。

用法：
  python create_entry.py --vault /path/to/vault --json '{"date": "...", ...}'
  python create_entry.py --vault /path/to/vault --date 2026-05-03 --title-summary "..." --category "域名" ...
"""

import argparse
import json
import os
import re
import shutil
import sys
from pathlib import Path

VALID_INVOICE_STATUSES = {"有发票", "有收据", "无发票", "待确认"}
VALID_PAYERS = {"北京盛通佳达科技有限公司", "成都汇智迅析科技有限公司", "海南云创软通科技有限公司", "个人"}
VALID_CATEGORIES = {"域名", "云计算", "餐饮", "交通", "住宿", "办公", "其他"}
VALID_REIMBURSEMENT_TYPES = {"日常", "出差"}
VALID_STATUSES = {"已记录", "待确认"}


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


def sanitize_keywords(text):
    """从摘要或 payee 中提取 2-4 个有辨识度的关键词，用下划线连接。"""
    cleaned = re.sub(r'[^\w一-鿿]', '_', text)
    cleaned = re.sub(r'_+', '_', cleaned).strip('_')
    parts = cleaned.split('_')
    parts = [p for p in parts if p]
    if not parts:
        return "未命名"
    return '_'.join(parts[:4])


def compute_sequence(entries_dir, year_month):
    """计算当月最大序号+1。"""
    target_dir = entries_dir / year_month
    if not target_dir.exists():
        return 1
    max_nn = 0
    for fname in target_dir.iterdir():
        if fname.suffix != ".md":
            continue
        m = re.match(r"^\d{4}-\d{2}-\d{2}-(\d{2})-", fname.name)
        if m:
            max_nn = max(max_nn, int(m.group(1)))
    return max_nn + 1


def safe_filename(summary, max_len=40):
    """将摘要转换为安全文件名。"""
    name = re.sub(r'[\\/:*?"<>|]', '', summary)
    name = name.strip().replace(' ', '_')
    return name[:max_len]


def build_entry_content(params):
    """构造完整的条目 Markdown 内容。"""
    frontmatter_lines = [
        "---",
        f"title: {params['date']} {params['title_summary']}",
        f"date: {params['date']}",
        "type: expense",
        f"category: {params['category']}",
        f"reimbursement_type: {params['reimbursement_type']}",
        f"amount: {params['amount']}",
        f"currency: {params['currency']}",
        f"amount_cny: {params['amount_cny']}",
        f"invoice_status: {params['invoice_status']}",
        f"pay_method: {params['pay_method']}",
        f"payer: {params['payer']}",
        f"payee: {params['payee']}",
        f"location: {params.get('location', '')}",
        f"project: {params.get('project', '')}",
        f"trip_name: {params.get('trip_name', '')}",
        f"status: {params['status']}",
        f"voucher: {params['voucher_path']}",
        "---",
    ]

    sections = [
        "",
        "## 说明",
        "",
        params.get("description", ""),
        "",
        "## 凭证",
        "",
        f"![[{params['voucher_path']}]]",
    ]

    if params.get("pending_note"):
        sections.extend([
            "",
            "> ⚠️ 此路径必须与 frontmatter `voucher` 字段完全一致。",
            "",
            "## 待确认事项",
            "",
            f"> ⚠️ 待确认：{params['pending_note']}",
        ])

    return "\n".join(frontmatter_lines + sections) + "\n"


def main():
    parser = argparse.ArgumentParser(description="创建财务记账条目笔记")
    parser.add_argument("--vault", default=os.environ.get("OBSIDIAN_VAULT_ROOT"),
                        help="Vault 根目录路径")
    parser.add_argument("--json", help="JSON 格式的完整参数")
    parser.add_argument("--date", help="日期 YYYY-MM-DD")
    parser.add_argument("--title-summary", help="摘要标题")
    parser.add_argument("--category", help="类别")
    parser.add_argument("--reimbursement-type", help="报销类型")
    parser.add_argument("--amount", type=float, help="原币金额")
    parser.add_argument("--currency", help="币种")
    parser.add_argument("--amount-cny", type=float, help="CNY金额")
    parser.add_argument("--invoice-status", help="发票状态")
    parser.add_argument("--pay-method", help="支付方式")
    parser.add_argument("--payer", help="付款方/抬头")
    parser.add_argument("--payee", help="收款方")
    parser.add_argument("--location", default="", help="地点")
    parser.add_argument("--project", default="", help="项目")
    parser.add_argument("--trip-name", default="", help="出差任务名")
    parser.add_argument("--status", default="已记录", help="状态")
    parser.add_argument("--description", default="", help="费用说明")
    parser.add_argument("--pending-note", default="", help="待确认说明")
    parser.add_argument("--source-voucher-path", default="", help="凭证源文件路径")
    args = parser.parse_args()

    if not args.vault:
        error("VAULT_NOT_SET", "未指定 vault 路径，请设置 --vault 或 OBSIDIAN_VAULT_ROOT")

    vault_root = Path(args.vault).expanduser().resolve()
    if not vault_root.exists():
        error("VAULT_NOT_FOUND", f"Vault 目录不存在：{vault_root}")

    # 解析参数
    if args.json:
        params = json.loads(args.json)
        # 为 JSON 输入中缺失的可选字段设置默认值
        params.setdefault("status", "已记录")
        params.setdefault("reimbursement_type", "日常")
        params.setdefault("location", "")
        params.setdefault("project", "")
        params.setdefault("trip_name", "")
        params.setdefault("description", "")
        params.setdefault("pending_note", "")
        params.setdefault("source_voucher_path", "")
        params.setdefault("pay_method", "")
        params.setdefault("invoice_status", "待确认")
        params.setdefault("amount", 0)
        params.setdefault("amount_cny", 0)
    else:
        if not all([args.date, args.title_summary, args.category, args.currency, args.payer, args.payee]):
            error("MISSING_FIELDS", "缺少必填字段（date, title-summary, category, currency, payer, payee）")
        params = {
            "date": args.date,
            "title_summary": args.title_summary,
            "category": args.category,
            "reimbursement_type": args.reimbursement_type or "日常",
            "amount": args.amount if args.amount else 0,
            "currency": args.currency,
            "amount_cny": args.amount_cny if args.amount_cny else 0,
            "invoice_status": args.invoice_status or "待确认",
            "pay_method": args.pay_method or "",
            "payer": args.payer,
            "payee": args.payee,
            "location": args.location,
            "project": args.project,
            "trip_name": args.trip_name,
            "status": args.status or "已记录",
            "description": args.description,
            "pending_note": args.pending_note,
            "source_voucher_path": args.source_voucher_path,
        }

    date = params["date"]
    year_month = date[:7]

    # 验证字段
    if params["invoice_status"] not in VALID_INVOICE_STATUSES:
        error("INVALID_INVOICE_STATUS",
              f"invoice_status 不在有效枚举值内：{params['invoice_status']}。有效值：{VALID_INVOICE_STATUSES}")

    if params["status"] not in VALID_STATUSES:
        error("INVALID_STATUS",
              f"status 不在有效枚举值内：{params['status']}。有效值：{VALID_STATUSES}")

    if params["category"] not in VALID_CATEGORIES:
        params["category"] = "其他"

    if params["reimbursement_type"] not in VALID_REIMBURSEMENT_TYPES:
        params["reimbursement_type"] = "日常"

    payer_input = params["payer"]
    if payer_input not in VALID_PAYERS:
        params["payer"] = "个人"
        print(f"# WARNING: payer '{payer_input}' 不在合法抬头列表中，已自动归入'个人'", file=sys.stderr)

    # 计算序号
    entries_dir = vault_root / "财务记账" / "条目"
    seq = compute_sequence(entries_dir, year_month)
    nn = f"{seq:02d}"

    # 构造凭证路径
    keywords = sanitize_keywords(params.get("title_summary", params.get("payee", "未命名")))
    source_voucher = params.get("source_voucher_path", "")

    if source_voucher:
        source_path = Path(source_voucher).expanduser().resolve()
        if not source_path.exists():
            error("VOUCHER_NOT_FOUND", f"凭证文件不存在：{source_path}")
        voucher_filename = source_path.name
        voucher_rel_dir = f"财务记账/原始凭证/{year_month}/{keywords}"
        voucher_rel_path = f"{voucher_rel_dir}/{voucher_filename}"
        voucher_abs_dir = vault_root / voucher_rel_dir
        voucher_abs_path = vault_root / voucher_rel_path
        voucher_abs_dir.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source_path, voucher_abs_path)
        if not voucher_abs_path.exists():
            error("VOUCHER_COPY_FAILED", f"凭证文件复制失败：{voucher_abs_path}")
    else:
        voucher_rel_path = ""

    params["voucher_path"] = voucher_rel_path

    # 构造条目文件名
    safe_summary = safe_filename(params["title_summary"])
    entry_filename = f"{date}-{nn}-{safe_summary}.md"
    entry_rel_dir = f"财务记账/条目/{year_month}"
    entry_rel_path = f"{entry_rel_dir}/{entry_filename}"
    entry_abs_dir = vault_root / entry_rel_dir
    entry_abs_dir.mkdir(parents=True, exist_ok=True)
    entry_abs_path = vault_root / entry_rel_path

    if entry_abs_path.exists():
        error("ENTRY_FILE_EXISTS",
              f"条目文件已存在：{entry_rel_path}（序号 {nn} 冲突，自动递增后重试）",
              recoverable=True)

    # 写入条目文件
    content = build_entry_content(params)
    entry_abs_path.write_text(content, encoding="utf-8")

    success({
        "entry_path": entry_rel_path,
        "voucher_path": voucher_rel_path,
        "sequence_number": seq,
    })


if __name__ == "__main__":
    main()
