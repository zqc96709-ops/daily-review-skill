#!/usr/bin/env python3
"""Append one time-log entry to the Feishu/Lark online workbook."""

from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo


DEFAULT_SPREADSHEET_TOKEN = "CRtSsJXa7hz7X9t7LRCcpDZcnFh"
TIME_SHEET_ID = "k2Hjnb"
TIME_SHEET_NAME = "时间记录"
MAX_TIME_ROW = 600

CATEGORIES = [
    "选品/市场调研",
    "上架/Listing优化",
    "广告/数据分析",
    "内容素材",
    "供应链/采购",
    "客服/售后",
    "订单/物流",
    "财务/记账",
    "学习/资料整理",
    "沟通/会议",
    "行政/工具/杂事",
    "休息/生活",
    "干扰/刷信息",
]

DISTRACTIONS = ["手机/短视频", "平台消息", "客户消息", "临时想法", "工具问题", "家务/生活", "疲劳拖延", "其他", ""]

CATEGORY_KEYWORDS = [
    ("选品/市场调研", ["选品", "竞品", "市场", "调研", "利润", "关键词调研", "需求"]),
    ("上架/Listing优化", ["上架", "listing", "标题", "五点", "关键词", "优化链接"]),
    ("广告/数据分析", ["广告", "投放", "acos", "roi", "数据", "报表", "转化", "点击"]),
    ("内容素材", ["素材", "图片", "视频", "文案", "短视频", "主图", "买家秀", "设计工具", "创业工具", "详情页设计"]),
    ("供应链/采购", ["供应链", "采购", "工厂", "1688", "打样", "报价", "交期", "供应商"]),
    ("客服/售后", ["客服", "客户", "售后", "纠纷", "评价", "消息"]),
    ("订单/物流", ["订单", "物流", "发货", "跟踪", "仓库", "运单"]),
    ("财务/记账", ["财务", "记账", "利润", "账单", "税", "成本"]),
    ("学习/资料整理", ["学习", "课程", "资料", "笔记", "整理方法"]),
    ("沟通/会议", ["会议", "沟通", "电话", "对接"]),
    ("行政/工具/杂事", ["工具", "账号", "插件", "表格", "设置", "杂事"]),
    ("休息/生活", ["休息", "午休", "吃饭", "零食", "运动", "家务", "生活", "睡觉", "洗漱"]),
    ("干扰/刷信息", ["刷", "短视频", "信息流", "摸鱼", "分心", "拖延"]),
]


def run_lark(args: list[str], stdin: str | None = None) -> dict:
    proc = subprocess.run(args, input=stdin, text=True, capture_output=True)
    if proc.returncode != 0:
        raise RuntimeError(
            "lark-cli failed\n"
            f"command: {' '.join(args)}\n"
            f"stdout: {proc.stdout}\n"
            f"stderr: {proc.stderr}"
        )
    try:
        return json.loads(proc.stdout)
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"lark-cli returned non-JSON output: {proc.stdout}") from exc


def today() -> str:
    return datetime.now(ZoneInfo("Asia/Shanghai")).date().isoformat()


def normalize_time(value: str) -> str:
    value = value.strip()
    if not value:
        return value
    for fmt in ("%H:%M", "%H点%M", "%H点"):
        try:
            parsed = datetime.strptime(value, fmt)
            return parsed.strftime("%H:%M")
        except ValueError:
            pass
    if value.isdigit() and 0 <= int(value) <= 23:
        return f"{int(value):02d}:00"
    return value


def infer_category(text: str) -> str:
    lower = text.lower()
    for category, keywords in CATEGORY_KEYWORDS:
        if any(keyword.lower() in lower for keyword in keywords):
            return category
    return "行政/工具/杂事"


def normalize_category(category: str | None, text: str) -> str:
    if not category:
        return infer_category(text)
    category = category.strip()
    if category in CATEGORIES:
        return category
    for item in CATEGORIES:
        if category in item or item in category:
            return item
    raise SystemExit(f"Unknown category: {category}. Valid categories: {', '.join(CATEGORIES)}")


def normalize_score(value: str | None, name: str) -> str:
    if value in (None, ""):
        return ""
    value = str(value).strip()
    if value not in {"1", "2", "3", "4", "5"}:
        raise SystemExit(f"{name} must be 1-5, got {value}")
    return value


def normalize_yes_no(value: str | None) -> str:
    if not value:
        return "否"
    value = value.strip().lower()
    if value in {"是", "yes", "y", "true", "1", "主线"}:
        return "是"
    return "否"


def next_empty_row(token: str) -> int:
    payload = run_lark([
        "lark-cli", "sheets", "+cells-get",
        "--as", "user",
        "--spreadsheet-token", token,
        "--sheet-id", TIME_SHEET_ID,
        "--range", f"A2:A{MAX_TIME_ROW}",
        "--include", "value",
        "--json",
    ])
    ranges = payload.get("data", {}).get("ranges", [])
    if not ranges:
        return 2
    cells = ranges[0].get("cells", [])
    row_indices = ranges[0].get("row_indices", [])
    for idx, row in enumerate(cells):
        row_number = row_indices[idx] if idx < len(row_indices) else idx + 2
        if not row or not row[0].get("value"):
            return int(row_number)
    raise SystemExit(f"No empty row found in {TIME_SHEET_NAME}!A2:A{MAX_TIME_ROW}. Extend the sheet first.")


def append_entry(args: argparse.Namespace) -> int:
    token = args.spreadsheet_token or os.environ.get("CBTR_SPREADSHEET_TOKEN") or DEFAULT_SPREADSHEET_TOKEN
    text_for_category = " ".join([args.task or "", args.output or "", args.note or ""])
    category = normalize_category(args.category, text_for_category)
    distraction = args.distraction or ""
    if distraction and distraction not in DISTRACTIONS:
        distraction = "其他"
    row = args.row or next_empty_row(token)
    date_value = args.date or today()
    start = normalize_time(args.start)
    end = normalize_time(args.end)
    values = [[
        {"value": date_value},
        {},
        {"value": start},
        {"value": end},
        {},
        {"value": category},
        {},
        {"value": normalize_yes_no(args.main)},
        {"value": args.task},
        {"value": args.output or ""},
        {"value": normalize_score(args.energy, "energy")},
        {"value": normalize_score(args.focus, "focus")},
        {"value": distraction},
        {"value": args.note or ""},
    ]]
    if args.dry_run:
        print(json.dumps({
            "spreadsheet_token": token,
            "sheet": TIME_SHEET_NAME,
            "row": row,
            "range": f"A{row}:N{row}",
            "values": values,
        }, ensure_ascii=False, indent=2))
        return row
    run_lark([
        "lark-cli", "sheets", "+cells-set",
        "--as", "user",
        "--spreadsheet-token", token,
        "--sheet-id", TIME_SHEET_ID,
        "--range", f"A{row}:N{row}",
        "--cells", "-",
        "--json",
    ], json.dumps(values, ensure_ascii=False))
    print(json.dumps({
        "ok": True,
        "spreadsheet_url": f"https://my.feishu.cn/sheets/{token}",
        "sheet": TIME_SHEET_NAME,
        "row": row,
        "date": date_value,
        "start": start,
        "end": end,
        "category": category,
        "main": normalize_yes_no(args.main),
        "task": args.task,
        "output": args.output or "",
    }, ensure_ascii=False, indent=2))
    return row


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--spreadsheet-token", default=None)
    parser.add_argument("--date", default=None, help="YYYY-MM-DD. Defaults to today in Asia/Shanghai.")
    parser.add_argument("--start", required=True, help="Start time, e.g. 09:00.")
    parser.add_argument("--end", required=True, help="End time, e.g. 10:00.")
    parser.add_argument("--category", default=None, choices=CATEGORIES)
    parser.add_argument("--main", default="否", help="是/否")
    parser.add_argument("--task", required=True, help="What was done.")
    parser.add_argument("--output", default="", help="Concrete result/output.")
    parser.add_argument("--energy", default="", help="1-5")
    parser.add_argument("--focus", default="", help="1-5")
    parser.add_argument("--distraction", default="", help="Optional distraction source.")
    parser.add_argument("--note", default="")
    parser.add_argument("--row", type=int, default=None, help="Override target row for repair/testing.")
    parser.add_argument("--dry-run", action="store_true")
    return parser.parse_args()


def main() -> None:
    try:
        append_entry(parse_args())
    except Exception as exc:
        print(str(exc), file=sys.stderr)
        raise SystemExit(1) from exc


if __name__ == "__main__":
    main()
