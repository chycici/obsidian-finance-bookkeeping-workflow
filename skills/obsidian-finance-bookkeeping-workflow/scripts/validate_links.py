#!/usr/bin/env python3
"""扫描指定范围的所有 Markdown 文件，验证 [[...]] 链接和 voucher 字段路径的可达性。

用法：
  python validate_links.py --vault /path/to/vault --month 2026-05
  python validate_links.py --vault /path/to/vault --all
"""

import argparse
import json
import os
import re
import sys
from pathlib import Path

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

WIKI_LINK_PATTERN = re.compile(r"\[\[([^\]|#]+)(?:[|#][^\]]*)?\]\]")


def error(errcode, message, recoverable=False):
    print(json.dumps({
        "status": "error",
        "error_code": errcode,
        "message": message,
        "recoverable": recoverable,
    }, ensure_ascii=False))
    sys.exit(0 if recoverable else 1)


def success(data):
    data["status"] = "有问题" if (data.get("broken_links") or data.get("broken_vouchers")) else "通过"
    print(json.dumps(data, ensure_ascii=False))


def collect_files(vault_root, month=None):
    """收集要扫描的文件列表。"""
    if month:
        # 扫描当月相关的所有文件：条目、月度页，以及涉及的模板和说明
        paths = [
            vault_root / "财务记账" / "条目" / month,
            vault_root / "财务记账" / "按月",
        ]
        files = []
        for p in paths:
            if p.exists():
                if p.is_dir():
                    files.extend(p.rglob("*.md"))
                else:
                    files.append(p)
        return files
    else:
        # 全量扫描
        finance_dir = vault_root / "财务记账"
        if not finance_dir.exists():
            return []
        return list(finance_dir.rglob("*.md"))


def check_link_exists(vault_root, link_path, source_file_rel):
    """检查一个 [[...]] 链接目标是否存在。"""
    # 尝试多种可能
    candidates = [
        vault_root / link_path,
        vault_root / f"{link_path}.md",
        # 如果链接是相对于 source_file 的
        (Path(source_file_rel).parent / link_path).resolve(),
    ]
    for c in candidates:
        try:
            c = c.resolve()
            if c.exists() and vault_root in c.parents:
                return True
        except Exception:
            pass

    # 尝试用 pathlib 的 glob 放宽匹配
    try:
        target = vault_root / link_path
        parent = target.parent
        name = target.name
        if parent.exists():
            for child in parent.iterdir():
                if child.stem == name or child.name == name:
                    return True
                if child.name.startswith(name):
                    return True
    except Exception:
        pass

    return False


def validate_file(vault_root, filepath):
    """验证单个文件的所有链接和凭证路径。返回 broken 列表。"""
    broken_links = []
    broken_vouchers = []
    rel_path = str(filepath.relative_to(vault_root))

    try:
        content = filepath.read_text(encoding="utf-8")
    except Exception:
        return broken_links, broken_vouchers

    # 检查 frontmatter voucher 字段
    try:
        post = frontmatter.load(filepath)
        voucher = post.metadata.get("voucher") or ""
        if voucher and str(voucher).strip():
            voucher_path = vault_root / voucher.strip()
            if not voucher_path.exists():
                broken_vouchers.append({
                    "source_file": rel_path,
                    "voucher_path": voucher.strip(),
                    "type": "voucher_not_found",
                })
    except Exception:
        pass

    # 检查正文所有 [[...]] 链接
    for match in WIKI_LINK_PATTERN.finditer(content):
        link_target = match.group(1).strip()
        if not link_target:
            continue
        if not check_link_exists(vault_root, link_target, rel_path):
            broken_links.append({
                "source_file": rel_path,
                "broken_link": link_target,
                "type": "internal_link",
            })

    return broken_links, broken_vouchers


def main():
    parser = argparse.ArgumentParser(description="验证 Markdown 链接和凭证路径")
    parser.add_argument("--vault", default=os.environ.get("OBSIDIAN_VAULT_ROOT"),
                        help="Vault 根目录路径")
    parser.add_argument("--month", help="限定月份，格式 YYYY-MM")
    parser.add_argument("--all", action="store_true", help="全量扫描整个财务记账目录")
    parser.add_argument("--json-output", help="输出报告到指定 JSON 文件")
    args = parser.parse_args()

    if not args.vault:
        error("VAULT_NOT_SET", "未指定 vault 路径")

    if not args.month and not args.all:
        error("MISSING_SCOPE", "请指定 --month 或 --all")

    vault_root = Path(args.vault).expanduser().resolve()
    if not vault_root.exists():
        error("VAULT_NOT_FOUND", f"Vault 目录不存在：{vault_root}")

    files = collect_files(vault_root, month=args.month)

    all_broken_links = []
    all_broken_vouchers = []

    for fp in files:
        broken_links, broken_vouchers = validate_file(vault_root, fp)
        all_broken_links.extend(broken_links)
        all_broken_vouchers.extend(broken_vouchers)

    report = {
        "scanned_files": len(files),
        "broken_links": all_broken_links,
        "broken_vouchers": all_broken_vouchers,
        "total_broken": len(all_broken_links) + len(all_broken_vouchers),
    }

    if args.json_output:
        Path(args.json_output).write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

    success(report)


if __name__ == "__main__":
    main()
