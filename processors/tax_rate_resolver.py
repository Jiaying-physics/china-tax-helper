# -*- coding: utf-8 -*-
"""税率询问与默认记忆。

优先级：
  1. 输入自带「税率」列且全表一致 → 直接用
  2. 本地默认税率 data/default_tax_rate.json → 自动填（不再问）
  3. 询问用户，并问「是否设为默认税率」，Y 则记忆下次自动填
也支持 --tax-rate 直接指定。全部本地完成。
"""
import json
import os
from typing import Optional

import config

DEFAULT_TAX_RATE_FILE = "default_tax_rate.json"


def load_default_rate(path: Optional[str] = None) -> Optional[str]:
    """读取本地默认税率；无则返回 None"""
    path = path or os.path.join(config.DATA_DIR, DEFAULT_TAX_RATE_FILE)
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            return json.load(f).get("tax_rate")
    return None


def save_default_rate(rate: str, path: Optional[str] = None) -> None:
    """保存默认税率到本地"""
    path = path or os.path.join(config.DATA_DIR, DEFAULT_TAX_RATE_FILE)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump({"tax_rate": rate}, f, ensure_ascii=False, indent=2)


def _norm_rate(v) -> Optional[str]:
    """规范化税率：输入 0.13 或 13 或 '13%' 都归一为 '0.13'；无效返回 None"""
    if v is None or v == "":
        return None
    s = str(v).strip().replace("%", "")
    try:
        f = float(s)
    except ValueError:
        return None
    # 如果 >1 视为百分数（13 → 0.13）
    if f > 1:
        f = f / 100.0
    if f < 0 or f > 1:
        return None
    return f"{f:.6f}".rstrip("0").rstrip(".") if f != 0 else "0"


def resolve_tax_rate(
    records,
    default_value: Optional[str] = None,
    interactive: bool = True,
    rate_path: Optional[str] = None,
) -> str:
    """解析整张表统一的税率。返回形如 '0.13' 的字符串。"""
    # 1. 输入自带列且全表一致
    vals = set()
    for rec in records:
        v = rec.get("税率")
        if v not in (None, ""):
            norm = _norm_rate(v)
            if norm is not None:
                vals.add(norm)
    if len(vals) == 1:
        return vals.pop()

    # 2. 命令行指定
    if default_value:
        norm = _norm_rate(default_value)
        if norm is not None:
            return norm

    # 3. 本地默认税率
    saved = load_default_rate(rate_path)
    if saved is not None:
        return saved

    # 4. 交互询问 + 是否设为默认
    if interactive:
        while True:
            ans = input("请输入本张发票的税率（如 0.13，或 13）：").strip()
            norm = _norm_rate(ans)
            if norm is None:
                print("  税率格式无效，请输入 0~1 之间的小数或百分数。")
                continue
            while True:
                mem = input("是否将该税率设为默认，下次自动填充？(Y/N) [Y]: ").strip().upper()
                if not mem:
                    mem = "Y"
                if mem in ("Y", "N"):
                    break
                print("  输入无效，请输入 Y 或 N。")
            if mem == "Y":
                save_default_rate(norm, rate_path)
            return norm

    # 非交互且无可用值时，报错而非猜测税率
    raise ValueError(
        "无法确定税率：请用 --tax-rate 0.13 指定，"
        "或在交互模式下运行以询问税率。"
    )