# -*- coding: utf-8 -*-
"""非必填列偏好记忆与询问。

规则（用户要求）：
  首次遇到某非必填列时询问「是否要填」：
    - 不需要 → 记入本地偏好（该列=跳过），下次不再问
    - 需要   → 问「统一填什么值」，该列所有行统一填此值，记忆该偏好，下次自动填
  输入里自带的列则直接用输入值，不问。
全部本地完成。
"""
import json
import os
from typing import Dict, Optional

import config

OPTIONAL_PREFS_FILE = "optional_cols_prefs.json"


def load_prefs(path: Optional[str] = None) -> Dict[str, str]:
    """读取非必填列偏好。返回 {列名: 'skip' 或 统一值}"""
    path = path or os.path.join(config.DATA_DIR, OPTIONAL_PREFS_FILE)
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_prefs(prefs: Dict[str, str], path: Optional[str] = None) -> None:
    path = path or os.path.join(config.DATA_DIR, OPTIONAL_PREFS_FILE)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(prefs, f, ensure_ascii=False, indent=2)


def resolve_optional_cols(
    records,
    columns,
    interactive: bool = True,
    prefs_path: Optional[str] = None,
) -> Dict[str, str]:
    """对非必填列 columns 逐一解析统一值。

    返回 {列名: 值}，其中值 '' 表示「跳过」（整列留空）。
    """
    prefs = load_prefs(prefs_path)
    result: Dict[str, str] = {}

    for col in columns:
        # 已有偏好
        if col in prefs:
            result[col] = prefs[col]
            continue

        # 输入自带该列且全表值一致 → 采用并记忆？这里遵循用户需求：自带列直接用不问
        vals = set()
        for rec in records:
            v = rec.get(col)
            if v not in (None, ""):
                vals.add(str(v).strip())
        if len(vals) == 1:
            solo = vals.pop()
            result[col] = solo
            prefs[col] = solo  # 也记忆，保持一致性
            continue

        if not interactive:
            result[col] = ""
            continue

        # 首次询问
        while True:
            ans = input(f"非必填列「{col}」是否要填写？(Y/N) [N]: ").strip().upper()
            if not ans:
                ans = "N"
            if ans in ("Y", "N"):
                break
            print("  输入无效，请输入 Y 或 N。")
        if ans == "N":
            result[col] = ""
            prefs[col] = ""  # skip
        else:
            val = input(f"  请填写「{col}」统一值（该列所有行统一填此值）：").strip()
            result[col] = val
            prefs[col] = val

    save_prefs(prefs, prefs_path)
    return result