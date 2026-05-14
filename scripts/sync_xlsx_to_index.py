#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""从「幻想少女公会招募图鉴.xlsx」生成 index.html 内 const allData 数据块。"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

try:
    import openpyxl
except ImportError as e:  # pragma: no cover
    print("需要 openpyxl: pip install openpyxl", file=sys.stderr)
    raise SystemExit(1) from e

BEGIN = "            /* >>> SYNC_XLSX_DATA_BEGIN */"
END = "            /* <<< SYNC_XLSX_DATA_END */"

# 第一行须与之一致（顺序可任意，列名须完全匹配）
XLSX_HEADERS = ("位阶", "角色", "职业", "种族", "系别", "地形")

RARITY_TEXT = {
    "传说": 3,
    "史诗": 2,
    "稀有": 1,
    "普通": 0,
}


def _norm(s: object) -> str:
    if s is None:
        return ""
    return str(s).strip()


def _parse_rarity(v: object) -> int:
    s = _norm(v)
    if not s:
        return 0
    if s.isdigit():
        return max(0, min(3, int(s)))
    return RARITY_TEXT.get(s, 0)


def _header_col_index(header_row: tuple) -> dict[str, int]:
    headers = [_norm(x) for x in header_row]
    idx: dict[str, int] = {}
    for name in XLSX_HEADERS:
        try:
            idx[name] = headers.index(name)
        except ValueError:
            raise SystemExit(f"表头缺少「{name}」，当前: {headers}") from None
    return idx


def _row_to_obj(row: tuple, col: dict[str, int]) -> dict | None:
    name = _norm(row[col["角色"]])
    if not name:
        return None
    return {
        "角色名": name,
        "职业": _norm(row[col["职业"]]),
        "种族": _norm(row[col["种族"]]),
        "属性": _norm(row[col["系别"]]),
        "地区": _norm(row[col["地形"]]),
        "稀有度": _parse_rarity(row[col["位阶"]]),
    }


def read_rows(xlsx: Path) -> list[dict]:
    wb = openpyxl.load_workbook(xlsx, read_only=True, data_only=True)
    ws = wb.active
    it = ws.iter_rows(values_only=True)
    first = next(it, None)
    if not first:
        return []
    col = _header_col_index(first)
    out: list[dict] = []
    for row in it:
        if not row:
            continue
        obj = _row_to_obj(row, col)
        if obj:
            out.append(obj)
    return out


def format_js_lines(rows: list[dict]) -> str:
    lines = [BEGIN]
    pad = "            "
    for i, obj in enumerate(rows):
        s = json.dumps(obj, ensure_ascii=False, separators=(", ", ": "))
        line = pad + s
        if i < len(rows) - 1:
            line += ","
        lines.append(line)
    lines.append(END)
    return "\n".join(lines) + "\n"


def inject(html: str, block: str) -> str:
    if BEGIN not in html or END not in html:
        raise SystemExit(
            "index.html 中未找到 SYNC 标记。请保留 "
            "/* >>> SYNC_XLSX_DATA_BEGIN */ 与 /* <<< SYNC_XLSX_DATA_END */ 两行。"
        )
    pat = re.compile(
        re.escape(BEGIN) + r".*?" + re.escape(END),
        re.DOTALL,
    )
    m = pat.search(html)
    if not m:
        raise SystemExit("无法匹配 SYNC 块（标记顺序或内容异常）。")
    return pat.sub(block.rstrip("\n"), html, count=1)


def main() -> None:
    script_dir = Path(__file__).resolve().parent
    default_root = script_dir.parent
    ap = argparse.ArgumentParser(description="用 xlsx 覆盖 index.html 内 allData")
    ap.add_argument("--xlsx", type=Path, default=default_root / "幻想少女公会招募图鉴.xlsx")
    ap.add_argument("--index", type=Path, default=default_root / "index.html")
    ap.add_argument("--dry-run", action="store_true", help="只打印行数不写文件")
    args = ap.parse_args()

    if not args.xlsx.is_file():
        raise SystemExit(f"找不到表格: {args.xlsx}")
    if not args.index.is_file():
        raise SystemExit(f"找不到页面: {args.index}")

    rows = read_rows(args.xlsx)
    if not rows:
        raise SystemExit("表格无有效数据行")

    block = format_js_lines(rows)
    text = args.index.read_text(encoding="utf-8")
    new_text = inject(text, block)

    if args.dry_run:
        print(f"将写入 {len(rows)} 条角色记录（dry-run，未写文件）")
        return

    args.index.write_text(new_text, encoding="utf-8", newline="\n")
    print(f"已更新 {args.index}，共 {len(rows)} 条")


if __name__ == "__main__":
    main()
