#!/usr/bin/env python3
"""读取指定月份的所有条目，按抬头分组，生成一个包含多 Sheet 的美化 xlsx 文件。

用法：
  python export_excel.py --vault /path/to/vault --month 2026-05
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

try:
    import openpyxl
    from openpyxl.styles import (
        Font, PatternFill, Alignment, Border, Side, numbers
    )
    from openpyxl.utils import get_column_letter
except ImportError:
    print(json.dumps({
        "status": "error",
        "error_code": "MISSING_DEPENDENCY",
        "message": "缺少 openpyxl 依赖，请执行: pip install openpyxl",
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

# 样式定义
DARK_BLUE_FILL = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
WHITE_BOLD_FONT = Font(name="微软雅黑", bold=True, color="FFFFFF", size=14)
WHITE_BOLD_FONT_HEADER = Font(name="微软雅黑", bold=True, color="FFFFFF", size=11)
LIGHT_BLUE_FILL = PatternFill(start_color="DEEAF1", end_color="DEEAF1", fill_type="solid")
LIGHT_YELLOW_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
LIGHT_RED_FILL = PatternFill(start_color="FFE0E0", end_color="FFE0E0", fill_type="solid")
LIGHT_GRAY_FILL = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
BOLD_FONT = Font(name="微软雅黑", bold=True, size=11)
NORMAL_FONT = Font(name="微软雅黑", size=11)
LINK_FONT = Font(name="微软雅黑", size=11, color="0563C1", underline="single")
CENTER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT_ALIGN = Alignment(horizontal="left", vertical="center", wrap_text=True)
RIGHT_ALIGN = Alignment(horizontal="right", vertical="center")
THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)
MEDIUM_BOTTOM_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="medium"),
)

COLUMN_WIDTHS = {
    "A": 6, "B": 12, "C": 30, "D": 10, "E": 10,
    "F": 12, "G": 6, "H": 8, "I": 14, "J": 10,
    "K": 18, "L": 20, "M": 8,
}


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


def parse_entry(filepath):
    """解析单个条目文件，返回 dict 或 None。"""
    try:
        post = frontmatter.load(filepath)
    except Exception:
        return None

    fm = post.metadata

    try:
        amount_cny_val = float(fm.get("amount_cny", 0) or 0)
    except (ValueError, TypeError):
        amount_cny_val = 0

    payer = fm.get("payer", "个人")
    if payer not in VALID_PAYERS:
        payer = "个人"

    return {
        "path": str(filepath.relative_to(filepath.parents[3])),
        "date": str(fm.get("date", "")),
        "title": str(fm.get("title", "")),
        "category": str(fm.get("category", "其他")),
        "reimbursement_type": str(fm.get("reimbursement_type", "日常")),
        "amount": float(fm.get("amount", 0) or 0),
        "currency": str(fm.get("currency", "CNY")),
        "amount_cny": amount_cny_val,
        "invoice_status": str(fm.get("invoice_status", "")),
        "pay_method": str(fm.get("pay_method", "")),
        "payer": payer,
        "payee": str(fm.get("payee", "")),
        "status": str(fm.get("status", "已记录")),
        "voucher": str(fm.get("voucher") or ""),
    }


def apply_cell_style(cell, font=None, fill=None, alignment=None, border=None, number_format=None):
    """便捷设置单元格样式。"""
    if font:
        cell.font = font
    if fill:
        cell.fill = fill
    if alignment:
        cell.alignment = alignment
    if border:
        cell.border = border
    if number_format:
        cell.number_format = number_format


def build_payer_sheet(wb, sheet_name, entries, vault_abs_path, month_year, currencies_used):
    """构建一个抬头的完整 Sheet。"""

    ws = wb.create_sheet(title=sheet_name)
    month_label = f"{month_year[:4]}年{month_year[5:]}月"

    # 设置列宽
    for col_letter, width in COLUMN_WIDTHS.items():
        ws.column_dimensions[col_letter].width = width

    current_row = 1

    # === 行1：标题行 ===
    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=13)
    title_cell = ws.cell(row=current_row, column=1,
                         value=f"{month_label} {sheet_name} 费用明细")
    apply_cell_style(title_cell, font=WHITE_BOLD_FONT, fill=DARK_BLUE_FILL, alignment=CENTER_ALIGN)
    current_row += 1

    # === 行2：空行 ===
    current_row += 1

    # === 汇率参考行 ===
    if currencies_used and not (len(currencies_used) == 1 and "CNY" in currencies_used):
        rate_lines = []
        for curr, rate in currencies_used.items():
            if curr != "CNY":
                rate_lines.append(f"1 {curr} = {rate:.4f} CNY")
        if rate_lines:
            rate_text = f"汇率参考：{' | '.join(rate_lines)}"
            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=13)
            rate_cell = ws.cell(row=current_row, column=1, value=rate_text)
            apply_cell_style(rate_cell, font=NORMAL_FONT, fill=LIGHT_YELLOW_FILL, alignment=LEFT_ALIGN)
            current_row += 1

    # === 发票状态汇总行（先占位，后面用公式填充）===
    summary_row = current_row
    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=13)
    summary_cell = ws.cell(row=current_row, column=1, value="发票状态汇总")
    apply_cell_style(summary_cell, font=BOLD_FONT, fill=LIGHT_BLUE_FILL, alignment=LEFT_ALIGN)
    current_row += 1

    # === 表头行 ===
    header_row = current_row
    headers = ["序号", "日期", "摘要", "类别", "报销类型", "原币金额", "币种",
               "汇率", "CNY金额", "发票状态", "支付方式", "收款方", "凭证"]
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=header_row, column=col_idx, value=h)
        apply_cell_style(cell, font=WHITE_BOLD_FONT_HEADER, fill=DARK_BLUE_FILL,
                         alignment=CENTER_ALIGN, border=THIN_BORDER)

    # 冻结表头行
    ws.freeze_panes = f"A{header_row + 1}"
    current_row += 1

    # === 明细行 ===
    data_start_row = current_row
    entries_sorted = sorted(entries, key=lambda x: (x["date"], x["path"]))

    for i, e in enumerate(entries_sorted, 1):
        row = current_row
        is_pending = e["status"] == "待确认"
        is_zero_cny = e["amount_cny"] == 0

        # 汇率计算
        if e["currency"] == "CNY":
            rate = 1.0
        else:
            rate = currencies_used.get(e["currency"], 1.0)
        if e["amount"] != 0 and rate != 0:
            rate_effective = e["amount_cny"] / e["amount"]
        else:
            rate_effective = rate

        values = [
            i,                          # A 序号
            e["date"],                  # B 日期
            e["title"],                 # C 摘要
            e["category"],              # D 类别
            e["reimbursement_type"],    # E 报销类型
            e["amount"],                # F 原币金额
            e["currency"],              # G 币种
            rate_effective,             # H 汇率
            None,                       # I CNY金额（公式）
            e["invoice_status"],        # J 发票状态
            e["pay_method"],            # K 支付方式
            e["payee"],                 # L 收款方
            e["voucher"],               # M 凭证（超链接）
        ]

        for col_idx, val in enumerate(values, 1):
            cell = ws.cell(row=row, column=col_idx, value=val)

            # 列样式
            if col_idx == 1:  # 序号
                align = CENTER_ALIGN
            elif col_idx == 2:  # 日期
                align = CENTER_ALIGN
            elif col_idx in (3, 11, 12):  # 摘要、支付方式、收款方
                align = LEFT_ALIGN
            elif col_idx in (4, 5, 7, 10):  # 类别、报销类型、币种、发票状态
                align = CENTER_ALIGN
            elif col_idx in (6, 8, 9):  # 金额列
                align = RIGHT_ALIGN
            elif col_idx == 13:  # 凭证列
                align = CENTER_ALIGN
            else:
                align = CENTER_ALIGN

            # 特殊样式
            row_fill = None
            if is_pending:
                row_fill = LIGHT_RED_FILL
            elif is_zero_cny:
                row_fill = LIGHT_YELLOW_FILL

            apply_cell_style(cell, font=NORMAL_FONT, alignment=align, border=THIN_BORDER, fill=row_fill)

            # 金额格式
            if col_idx == 6:  # F 原币金额
                cell.number_format = '#,##0.00'
            elif col_idx == 8:  # H 汇率
                cell.number_format = '0.0000'
            elif col_idx == 9:  # I CNY金额 公式
                cell.value = f"=F{row}*H{row}"
                cell.number_format = '#,##0.00'

        # M 列凭证超链接
        col_m = ws.cell(row=row, column=13)
        voucher_path = e["voucher"]
        if voucher_path:
            abs_voucher = vault_abs_path / voucher_path
            if abs_voucher.exists():
                col_m.value = "查看"
                col_m.hyperlink = f"file:///{abs_voucher.as_posix()}"
                apply_cell_style(col_m, font=LINK_FONT, alignment=CENTER_ALIGN, border=THIN_BORDER,
                                 fill=LIGHT_RED_FILL if is_pending else None)
            else:
                col_m.value = "缺失"
                apply_cell_style(col_m, font=Font(name="微软雅黑", size=11, color="FF0000"),
                                 alignment=CENTER_ALIGN, border=THIN_BORDER,
                                 fill=LIGHT_RED_FILL)
        else:
            col_m.value = ""
            apply_cell_style(col_m, font=NORMAL_FONT, alignment=CENTER_ALIGN, border=THIN_BORDER,
                             fill=LIGHT_RED_FILL if is_pending else None)

        current_row += 1

    data_end_row = current_row - 1

    # === 小计行 ===
    subtotal_row = current_row
    for col_idx in range(1, 14):
        cell = ws.cell(row=subtotal_row, column=col_idx)
        border = MEDIUM_BOTTOM_BORDER
        fill = LIGHT_GRAY_FILL
        font = BOLD_FONT

        if col_idx == 1:
            cell.value = "小计"
            apply_cell_style(cell, font=font, fill=fill, alignment=CENTER_ALIGN, border=border)
        elif col_idx == 6:  # F 原币金额小计
            if data_start_row <= data_end_row:
                cell.value = f"=SUM(F{data_start_row}:F{data_end_row})"
            else:
                cell.value = 0
            cell.number_format = '#,##0.00'
            apply_cell_style(cell, font=font, fill=fill, alignment=RIGHT_ALIGN, border=border)
        elif col_idx == 9:  # I CNY金额小计
            if data_start_row <= data_end_row:
                cell.value = f"=SUM(I{data_start_row}:I{data_end_row})"
            else:
                cell.value = 0
            cell.number_format = '#,##0.00'
            apply_cell_style(cell, font=font, fill=fill, alignment=RIGHT_ALIGN, border=border)
        else:
            apply_cell_style(cell, font=font, fill=fill, alignment=CENTER_ALIGN, border=border)

    # === 更新发票状态汇总行 ===
    summary_parts = []
    for status in INVOICE_STATUSES:
        count_formula = f'=COUNTIF(J{data_start_row}:J{data_end_row},"{status}")'
        amount_formula = f'=SUMIF(J{data_start_row}:J{data_end_row},"{status}",I{data_start_row}:I{data_end_row})'
        summary_parts.append(f'{status}: {count_formula}条 / {amount_formula}')

    summary_text = " | ".join(summary_parts)
    ws.cell(row=summary_row, column=1).value = summary_text

    return {
        "data_start_row": data_start_row,
        "data_end_row": data_end_row,
        "summary_row": summary_row,
        "subtotal_row": subtotal_row,
        "total_cny": sum(e["amount_cny"] for e in entries_sorted),
        "count": len(entries_sorted),
        "by_invoice_status": {
            s: {
                "count": sum(1 for e in entries_sorted if e["invoice_status"] == s),
            }
            for s in INVOICE_STATUSES
        },
    }


def build_overview_sheet(wb, payer_sheet_info, month_label):
    """构建总览 Sheet。"""
    ws = wb.create_sheet(title="总览", index=0)

    # 格式参考
    ws.column_dimensions["A"].width = 28
    for col_letter in "BCDEFG":
        ws.column_dimensions[col_letter].width = 16

    # 标题
    current_row = 1
    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=7)
    title_cell = ws.cell(row=current_row, column=1,
                         value=f"{month_label} 费用总览")
    apply_cell_style(title_cell, font=WHITE_BOLD_FONT, fill=DARK_BLUE_FILL, alignment=CENTER_ALIGN)
    current_row += 2

    # 表头
    overview_headers = ["抬头", "条目数", "总金额CNY", "有发票", "有收据", "无发票", "待确认"]
    for col_idx, h in enumerate(overview_headers, 1):
        cell = ws.cell(row=current_row, column=col_idx, value=h)
        apply_cell_style(cell, font=WHITE_BOLD_FONT_HEADER, fill=DARK_BLUE_FILL,
                         alignment=CENTER_ALIGN, border=THIN_BORDER)
    current_row += 1

    # 数据行
    data_start_row = current_row
    for payer_name in PAYER_GROUPS_ORDER:
        info = payer_sheet_info.get(payer_name)
        if info and info["count"] > 0:
            sheet_ref = f"'{payer_name}'"
            by_status = info.get("by_invoice_status", {})
            values = [
                payer_name,
                info["count"],                                    # B: 条目数
                f"={sheet_ref}!I{info['subtotal_row']}",           # C: 总金额（跨Sheet公式）
                by_status.get("有发票", {}).get("count", 0),       # D: 有发票条数
                by_status.get("有收据", {}).get("count", 0),       # E: 有收据条数
                by_status.get("无发票", {}).get("count", 0),       # F: 无发票条数
                by_status.get("待确认", {}).get("count", 0),       # G: 待确认条数
            ]
        else:
            values = [payer_name, 0, 0, 0, 0, 0, 0]

        for col_idx, val in enumerate(values, 1):
            cell = ws.cell(row=current_row, column=col_idx, value=val)
            apply_cell_style(cell, font=NORMAL_FONT, alignment=CENTER_ALIGN if col_idx > 1 else LEFT_ALIGN,
                             border=THIN_BORDER)
            if col_idx == 3:  # C 列总金额
                cell.number_format = '#,##0.00'

        current_row += 1

    data_end_row = current_row - 1

    # 合计行
    total_row = current_row
    total_vals = ["合计",
                  f"=SUM(B{data_start_row}:B{data_end_row})",
                  f"=SUM(C{data_start_row}:C{data_end_row})",
                  f"=SUM(D{data_start_row}:D{data_end_row})",
                  f"=SUM(E{data_start_row}:E{data_end_row})",
                  f"=SUM(F{data_start_row}:F{data_end_row})",
                  f"=SUM(G{data_start_row}:G{data_end_row})"]
    for col_idx, val in enumerate(total_vals, 1):
        cell = ws.cell(row=total_row, column=col_idx, value=val)
        apply_cell_style(cell, font=BOLD_FONT, fill=LIGHT_GRAY_FILL,
                         alignment=CENTER_ALIGN if col_idx > 1 else LEFT_ALIGN,
                         border=MEDIUM_BOTTOM_BORDER)
        if col_idx == 3:
            cell.number_format = '#,##0.00'


def main():
    parser = argparse.ArgumentParser(description="导出当月费用 Excel")
    parser.add_argument("--vault", default=os.environ.get("OBSIDIAN_VAULT_ROOT"),
                        help="Vault 根目录路径")
    parser.add_argument("--month", required=True, help="月份 YYYY-MM")
    parser.add_argument("--output", help="输出文件路径（可选，默认在 vault 导出目录下）")
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
              f"条目目录不存在：{entries_dir}。请先创建条目再导出。",
              recoverable=True)

    # 解析所有条目
    entry_files = sorted(entries_dir.glob("*.md"))
    entries = []
    for fp in entry_files:
        entry = parse_entry(fp)
        if entry:
            entries.append(entry)

    if not entries:
        print(json.dumps({
            "month": month,
            "total_entries": 0,
            "note": "当月无有效条目，跳过导出",
            "status": "success",
        }, ensure_ascii=False))
        return

    # 按 payer 分组
    groups = defaultdict(list)
    for e in entries:
        groups[e["payer"]].append(e)

    # 收集汇率信息
    currencies = {}
    for e in entries:
        if e["currency"] != "CNY" and e["currency"] not in currencies:
            if e["amount"] != 0:
                currencies[e["currency"]] = e["amount_cny"] / e["amount"]
            else:
                currencies[e["currency"]] = 1.0

    month_label = f"{month[:4]}年{month[5:]}月"

    # 创建 workbook
    wb = openpyxl.Workbook()
    # 删除默认 Sheet
    wb.remove(wb.active)

    # 先创建各抬头 Sheet
    payer_sheet_info = {}
    for payer_name in PAYER_GROUPS_ORDER:
        group_entries = groups.get(payer_name, [])
        if group_entries:
            info = build_payer_sheet(wb, payer_name, group_entries, vault_root, month, currencies)
        else:
            # 空 Sheet
            ws = wb.create_sheet(title=payer_name)
            for col_letter, width in COLUMN_WIDTHS.items():
                ws.column_dimensions[col_letter].width = width
            ws.merge_cells("A1:M1")
            title_cell = ws.cell(row=1, column=1,
                                 value=f"{month_label} {payer_name} 费用明细")
            apply_cell_style(title_cell, font=WHITE_BOLD_FONT, fill=DARK_BLUE_FILL, alignment=CENTER_ALIGN)
            ws.merge_cells("A3:M3")
            ws.cell(row=3, column=1, value="本月无此抬头条目")
            ws.freeze_panes = "A4"

            headers = ["序号", "日期", "摘要", "类别", "报销类型", "原币金额", "币种",
                       "汇率", "CNY金额", "发票状态", "支付方式", "收款方", "凭证"]
            for col_idx, h in enumerate(headers, 1):
                cell = ws.cell(row=4, column=col_idx, value=h)
                apply_cell_style(cell, font=WHITE_BOLD_FONT_HEADER, fill=DARK_BLUE_FILL,
                                 alignment=CENTER_ALIGN, border=THIN_BORDER)

            info = {"data_start_row": 5, "data_end_row": 4, "summary_row": 2,
                    "subtotal_row": 5, "total_cny": 0, "count": 0,
                    "by_invoice_status": {s: {"count": 0} for s in INVOICE_STATUSES}}
        payer_sheet_info[payer_name] = info

    # 创建总览 Sheet（作为第一个 Sheet）
    build_overview_sheet(wb, payer_sheet_info, month_label)

    # 确保输出目录存在
    output_dir = vault_root / "财务记账" / "导出" / month
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / f"{month}-费用导出.xlsx"

    wb.save(output_path)

    broken_vouchers = []
    for e in entries:
        if e["voucher"] and not (vault_root / e["voucher"]).exists():
            broken_vouchers.append({
                "source_file": e["path"],
                "voucher_path": e["voucher"],
            })

    success({
        "output_file": str(output_path.relative_to(vault_root)),
        "total_entries": len(entries),
        "by_payer": {
            p: {
                "count": payer_sheet_info[p]["count"],
                "total_cny": round(payer_sheet_info[p]["total_cny"], 2),
            }
            for p in PAYER_GROUPS_ORDER
        },
        "broken_vouchers": broken_vouchers if broken_vouchers else None,
    })


if __name__ == "__main__":
    main()
