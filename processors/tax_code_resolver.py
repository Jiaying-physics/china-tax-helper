# -*- coding: utf-8 -*-
"""税收分类编码解析与询问。

三级解析：
  1. 输入自带「商品和服务税收分类编码」列
  2. 记忆库 data/tax_code_map.json（精确+子串匹配）
  3. 命令行交互询问用户，并记忆入库，下次自动命中
也支持 --tax-code 为所有未命中项提供统一兜底编码。全部本地完成。
"""
import json
import os
import re
from typing import Dict, Optional

import config

# 19 位编码正则
_CODE_RE = re.compile(r"^\d{19}$")


def load_tax_map(path: Optional[str] = None) -> Dict[str, str]:
    """读取记忆库 {项目名/子串: 编码}"""
    path = path or os.path.join(config.DATA_DIR, "tax_code_map.json")
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_tax_map(mapping: Dict[str, str], path: Optional[str] = None) -> None:
    """保存记忆库到 data/tax_code_map.json"""
    path = path or os.path.join(config.DATA_DIR, "tax_code_map.json")
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(mapping, f, ensure_ascii=False, indent=2)


def _lookup(name: str, mapping: Dict[str, str]) -> Optional[str]:
    """精确/子串匹配记忆库。返回编码；命中则返回，否则 None。"""
    if name in mapping:
        return mapping[name]
    # 子串匹配：某个 key 是项目名的子串（如「螺母」命中「320大孔径锁紧螺母」）
    for key, code in mapping.items():
        if key and key in name:
            return code
    return None


def resolve_tax_included(
    records,
    default_value: Optional[str] = None,
    interactive: bool = True,
) -> str:
    """解析整张表统一的「是否含税」标识（Y=含税 / N=不含税）。

    通常一张发票要么全部含税，要么全部不含税，故统一询问一次。
    优先级：
      1. 输入自带「是否含税」列且全表值一致 → 采用
      2. default_value 传入（--tax-included）→ 采用
      3. 交互询问一次（回车默认 N）
    """
    # 1. 输入自带列且全表一致
    vals = set()
    for rec in records:
        v = rec.get("是否含税")
        if v not in (None, ""):
            vals.add(str(v).strip().upper())
    if len(vals) == 1:
        solo = vals.pop()
        if solo in ("Y", "N"):
            return solo

    # 2. 默认值（--tax-included）
    if default_value:
        dv = default_value.strip().upper()
        if dv in ("Y", "N"):
            return dv

    # 3. 交互询问一次
    if interactive:
        while True:
            ans = input("本张发票是否含税？输入 Y（含税）/ N（不含税）[回车=N]: ").strip().upper()
            if not ans:
                ans = "N"
            if ans in ("Y", "N"):
                return ans
            print("  输入无效，请输入 Y 或 N。")

    # 非交互且无有效输入时默认不含税
    return "N"


def resolve_codes(
    records,
    map_path: Optional[str] = None,
    default_code: Optional[str] = None,
    interactive: bool = True,
) -> Dict[str, str]:
    """为每条记录解析税收分类编码；返回 {项目名称: 编码}。"""
    mapping = load_tax_map(map_path)
    result: Dict[str, str] = {}
    unresolved = {}  # name -> 记录中的值或 None

    for rec in records:
        name = str(rec.get("项目名称", "")).strip() or f"项目{len(result)+1}"
        # 1. 输入自带编码
        raw = rec.get("商品和服务税收分类编码")
        if raw not in (None, ""):
            code = str(raw).strip()
            if _CODE_RE.match(code):
                result[name] = code
                continue
        # 2. 记忆库
        found = _lookup(name, mapping)
        if found:
            result[name] = found
            continue
        unresolved[name] = raw

    # 3. 兜底（--tax-code）
    if default_code:
        for name in unresolved:
            result[name] = default_code
        return result

    # 4. 交互：一次性清单 + 单一编码应用到全部（一般一张表就 1~2 个编码）
    if interactive and unresolved:
        print("\n以下项目未能自动匹配到有效的 19 位税收分类编码：")
        for idx, name in enumerate(unresolved, 1):
            print(f"  {idx}. {name}")
        while True:
            ans = input("请输入一个 19 位税收分类编码应用到以上全部项目（回车=跳过，留空）: ").strip()
            if not ans:
                print("  未提供编码，以上项目编码留空（可后续在 inputs 中补充后重跑）。")
                for name in unresolved:
                    result[name] = ""
                break
            if _CODE_RE.match(ans):
                for name in unresolved:
                    result[name] = ans
                    mapping[name] = ans  # 记忆，供下次自动匹配
                break
            print("  编码格式错误，需为 19 位纯数字。")
        save_tax_map(mapping, map_path)

    return result