#!/usr/bin/env python3
"""Write one daily review entry to the Feishu/Lark online workbook."""

from __future__ import annotations

import argparse
import json
import os
import subprocess
from datetime import datetime
from zoneinfo import ZoneInfo


DEFAULT_SPREADSHEET_TOKEN = "CRtSsJXa7hz7X9t7LRCcpDZcnFh"
REVIEW_SHEET_ID = "l4b0oN"
REVIEW_SHEET_NAME = "每日复盘"
MAX_REVIEW_ROW = 120


def run_lark(args: list[str], stdin: str | None = None) -> dict:
    proc = subprocess.run(args, input=stdin, text=True, capture_output=True)
    if proc.returncode != 0:
        raise RuntimeError(f"lark-cli failed: {' '.join(args)}\n{proc.stdout}\n{proc.stderr}")
    return json.loads(proc.stdout)


def today() -> str:
    return datetime.now(ZoneInfo("Asia/Shanghai")).date().isoformat()


def normalize_score(value: str | None, field: str) -> str:
    if value in (None, ""):
        return ""
    value = str(value).strip()
    if value not in {"1", "2", "3", "4", "5"}:
        raise SystemExit(f"{field} must be 1-5, got {value}")
    return value


def normalize_completion(value: str | None) -> str:
    if not value:
        return ""
    value = str(value).strip()
    mapping = {"0": "0%", "25": "25%", "50": "50%", "75": "75%", "100": "100%"}
    if value in mapping:
        return mapping[value]
    if value in {"0%", "25%", "50%", "75%", "100%"}:
        return value
    raise SystemExit("completion must be one of 0%, 25%, 50%, 75%, 100%")


def row_for_date(token: str, target_date: str) -> int:
    payload = run_lark([
        "lark-cli", "sheets", "+cells-get",
        "--as", "user",
        "--spreadsheet-token", token,
        "--sheet-id", REVIEW_SHEET_ID,
        "--range", f"A2:A{MAX_REVIEW_ROW}",
        "--include", "value",
        "--json",
    ])
    rng = payload.get("data", {}).get("ranges", [{}])[0]
    cells = rng.get("cells", [])
    rows = rng.get("row_indices", [])
    for idx, row in enumerate(cells):
        row_num = rows[idx] if idx < len(rows) else idx + 2
        value = str(row[0].get("value", "")) if row else ""
        if value[:10] == target_date:
            return int(row_num)
    raise SystemExit(f"Date {target_date} was not found in {REVIEW_SHEET_NAME}!A2:A{MAX_REVIEW_ROW}.")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--spreadsheet-token", default=None)
    parser.add_argument("--date", default=None)
    parser.add_argument("--main-task", required=True)
    parser.add_argument("--completion", default="")
    parser.add_argument("--outputs", default="")
    parser.add_argument("--time-sink", default="")
    parser.add_argument("--distraction", default="")
    parser.add_argument("--tomorrow-main", default="")
    parser.add_argument("--tomorrow-first-step", default="")
    parser.add_argument("--focus", default="")
    parser.add_argument("--energy", default="")
    parser.add_argument("--satisfaction", default="")
    parser.add_argument("--conclusion", default="")
    parser.add_argument("--dry-run", action="store_true")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    token = args.spreadsheet_token or os.environ.get("CBTR_SPREADSHEET_TOKEN") or DEFAULT_SPREADSHEET_TOKEN
    date_value = args.date or today()
    row = row_for_date(token, date_value)
    cells = [[
        {},
        {"value": args.main_task},
        {"value": normalize_completion(args.completion)},
        {"value": args.outputs},
        {"value": args.time_sink},
        {"value": args.distraction},
        {"value": args.tomorrow_main},
        {"value": args.tomorrow_first_step},
        {"value": normalize_score(args.focus, "focus")},
        {"value": normalize_score(args.energy, "energy")},
        {"value": normalize_score(args.satisfaction, "satisfaction")},
        {"value": args.conclusion},
    ]]
    if args.dry_run:
        print(json.dumps({"row": row, "range": f"A{row}:L{row}", "cells": cells}, ensure_ascii=False, indent=2))
        return
    run_lark([
        "lark-cli", "sheets", "+cells-set",
        "--as", "user",
        "--spreadsheet-token", token,
        "--sheet-id", REVIEW_SHEET_ID,
        "--range", f"A{row}:L{row}",
        "--cells", "-",
        "--json",
    ], json.dumps(cells, ensure_ascii=False))
    print(json.dumps({
        "ok": True,
        "spreadsheet_url": f"https://my.feishu.cn/sheets/{token}",
        "sheet": REVIEW_SHEET_NAME,
        "row": row,
        "date": date_value,
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
