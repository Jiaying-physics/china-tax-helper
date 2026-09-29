# -*- coding: utf-8 -*-
"""Excel 读取与表头智能识别。

自动定位输入文件的表头行，将列名按别名映射到模板列名，
返回规范化后的数据记录列表。全部本地处理。
"""
import os
from typing import Dict, List, Optional

import openpyxl

import config


def _norm(s) -> str:
    """归一化：去空白、去换行、去空格"""
    return "".join(str(s).split()) if s is not None else ""


def _to_num(v):
    """尽力把值转成 float；失败返回 None"""
    if v is None or v == "":
        return None
    s = str(v).strip().replace(",", "").replace("元", "").replace("%", "")
    try:
        return float(s)
    except ValueError:
        return None


def match_column_candidates(header: str, columns=None) -> list:
    """返回表头命中的所有候选 (模板列, 权重)，按权重降序。

    columns: 可选，限定只匹配这些模板列。
    """
    h = _norm(header)
    if not h:
        return []
    keywords = config.HEADER_KEYWORDS
    if columns:
        keywords = {k: v for k, v in keywords.items() if k in columns}

    candidates = []
    for template_col, kw_list in keywords.items():
        for kw in kw_list:
            nkw = _norm(kw)
            if not nkw:
                continue
            if h == nkw:               # 精确匹配
                candidates.append((template_col, 10000 + len(nkw)))
            elif nkw in h:             # 包含匹配
                candidates.append((template_col, len(nkw)))
    candidates.sort(key=lambda x: -x[1])
    return candidates


def match_column(header: str, columns=None) -> Optional[str]:
    """根据表头文本，返回映射到的模板列名；识别不到返回 None。"""
    cands = match_column_candidates(header, columns)
    return cands[0][0] if cands else None


# 数值型模板列（其数据应为数字，用于与「项目名称」等文本列消歧）
_NUMERIC_COLS = {"商品数量", "商品单价", "金额", "单价", "税率", "折扣金额"}


def _column_all_numeric(ws, header_row, col_idx, probe_rows=5) -> bool:
    """判断某列数据是否「全部为数字」（用于与项目名称文本列消歧）。

    关键点：项目名称（商品名称）虽常混有数字（如「320偏心环」），但**一定含文字**。
    因此：只要采样若干行中任一含文字 → 视为文本列（项目名称）；
          全部为纯数字（或数字+金额单位，如 '38.9元'）→ 视为数值列。
    """
    seen = 0
    for rr in range(header_row + 1, min(header_row + 1 + probe_rows, ws.max_row + 1)):
        v = ws.cell(row=rr, column=col_idx + 1).value
        if v is None or _norm(v) == "":
            continue
        seen += 1
        s = _norm(v)
        # 纯数字
        if isinstance(v, (int, float)):
            continue
        # 数字（含金额单位 元/个 等）→ 算数字
        s2 = s.replace("元", "").replace("个", "").replace("件", "")
        if _to_num(s2) is not None:
            continue
        # 含汉字 → 判定为文本列（项目名称）
        return False
    # 有数据且全部数字 → True；无数据 → False（保守，回到项目名称）
    return seen > 0


def _resolve_mapping(ws, header_row, columns) -> dict:
    """逐列解析表头 → 模板列，并对歧义作内容判断。

    歧义场景：表头如「商品数量」「商品单价」同时命中「项目名称」（因含「商品」）和
    数值列（因含「数量/单价」）。此时按该列数据内容判断：
      - 数据是数字 → 数值列（商品数量/商品单价）
      - 数据是文本 → 项目名称
    """
    mapping = {}
    for col in range(ws.max_column):
        cell = ws.cell(row=header_row, column=col + 1).value
        cands = match_column_candidates(cell, columns)
        if not cands:
            continue
        best = cands[0][0]
        # 是否存在「项目名称 vs 数值列」歧义
        has_name = any(t == "项目名称" for t, _ in cands)
        has_num = any(t in _NUMERIC_COLS for t, _ in cands)
        if has_name and has_num and best == "项目名称":
            # 若该列数据全部为数字 → 判定为数值列（如「商品数量」）
            numeric_col = next((t for t, _ in cands if t in _NUMERIC_COLS), None)
            if numeric_col and _column_all_numeric(ws, header_row, col):
                best = numeric_col
        mapping[col] = best
    return mapping


def _disambiguate_amount(records: List[Dict[str, object]]) -> List[Dict[str, object]]:
    """用「单价×数量=金额」的恒等式消歧。

    当一行中多个候选列都命中了「金额」或「商品单价」（如「总价」vs「含税金额」、
    「含税单价」vs「未税单价」），用第一行数据做计算校验：
      候选金额 ≈ 候选单价 × 候选数量 → 保留吻合的那对，其余降权/丢弃。
    """
    if not records:
        return records

    # 收集所有命中「金额」和「商品单价」的候选列
    first = records[0]
    amount_cols = [k for k in first.keys() if k == "金额"]
    price_cols = [k for k in first.keys() if k == "商品单价"]

    # 已有数量列
    qty_col = "商品数量"

    # 如果只有一列金额和一列单价，无需消歧
    if len(amount_cols) <= 1 and len(price_cols) <= 1:
        return records

    # 取第一行有完整数据的记录做校验
    test_rec = None
    for rec in records:
        q = _to_num(rec.get(qty_col))
        prices = {k: _to_num(rec.get(k)) for k in price_cols if rec.get(k) not in (None, "")}
        amounts = {k: _to_num(rec.get(k)) for k in amount_cols if rec.get(k) not in (None, "")}
        if q and prices and amounts:
            test_rec = (q, prices, amounts)
            break
    if test_rec is None:
        return records

    q, prices, amounts = test_rec
    # 计算：哪个价格 × 数量 约等于哪个金额 → 吻合的组合保留
    best_price, best_amount = None, None
    best_err = float("inf")
    for pk, pv in prices.items():
        for ak, av in amounts.items():
            if pv is None or av is None:
                continue
            err = abs(q * pv - av)
            if err < best_err:
                best_err = err
                best_price, best_amount = pk, av
    # 误差阈值：金额的 5%（含四舍五入误差）
    if best_price is None or best_amount is None:
        return records
    if best_err > max(0.01, abs(best_amount) * 0.05):
        return records  # 无法确认，保留原样

    # 保留吻合的列，把其他「金额」「单价」候选列的值清空
    for rec in records:
        for k in amount_cols:
            if k != best_amount_key(amount_cols, prices, amounts, best_price, best_amount):
                rec[k] = ""
    return records


def best_amount_key(amount_cols, prices, amounts, best_price, best_amount):
    """返回应该保留的金额列 key。"""
    for pk, pv in prices.items():
        if pk == best_price:
            for ak, av in amounts.items():
                if av == best_amount:
                    return ak
    return amount_cols[0]


def _disambiguate_mapping(ws, header_row, mapping, price_choice=None):
    """表头映射消歧：多个输入列命中同一模板列时，用「数量×单价=金额」校验选留。

    mapping: {输入列号(0基): 模板列名}。会修改 mapping。
    price_choice: 可选 '含税'/'不含税'。当同时有含税/不含税单价时，
                  若未指定则交互询问用户选一个；每次询问，不记忆。
    场景示例：
      - 「含税单价」「未税单价」都命中「商品单价」+「含税金额」「税前金额」都命中「金额」
        → 询问用户用哪个口径，保留对应单价+金额列
      - 「总价」命中「金额」但同时有「单价」「数量」→ 校验后正确归位
    """
    # 找出每个模板列命中的输入列
    by_tpl = {}
    for col_idx, tpl in mapping.items():
        by_tpl.setdefault(tpl, []).append(col_idx)

    qty_cols = by_tpl.get("商品数量", [])
    price_cols = by_tpl.get("商品单价", [])
    amount_cols = by_tpl.get("金额", [])
    if not qty_cols or not price_cols:
        return

    # 识别每个单价列的「口径」：含税 / 不含税 / 中性
    def _price_type(col_idx):
        h = _norm(ws.cell(row=header_row, column=col_idx + 1).value)
        if "含税" in h:
            return "含税"
        if any(k in h for k in ("未税", "不含税", "税前", "无税")):
            return "不含税"
        return "中性"

    price_types = {col: _price_type(col) for col in price_cols}
    has_taxed = any(t == "含税" for t in price_types.values())
    has_untaxed = any(t == "不含税" for t in price_types.values())

    # 决定选用哪个单价列
    if has_taxed and has_untaxed:
        # 含税+不含税并存 → 询问（或按 price_choice 指定）
        choice = price_choice
        if choice is None:
            # 展示两个选项的值，让用户判断
            data_row = header_row + 1
            desc = []
            for col in price_cols:
                v = _to_num(ws.cell(row=data_row, column=col + 1).value)
                t = price_types[col]
                if t != "中性":
                    desc.append(f"{t}价({v})")
            print("表格中同时存在含税与不含税单价，请选择使用哪个：")
            if has_taxed:
                print("  1. 含税价")
            if has_untaxed:
                print("  2. 不含税价")
            while True:
                ans = input("请输入编号 (1/2)：").strip()
                if ans == "1" and has_taxed:
                    choice = "含税"
                    break
                if ans == "2" and has_untaxed:
                    choice = "不含税"
                    break
                print("  输入无效，请重新输入。")
        if choice == "含税":
            best_price_col = next(c for c in price_cols if price_types[c] == "含税")
        else:
            best_price_col = next(c for c in price_cols if price_types[c] == "不含税")
    else:
        # 无含税/不含税并存：用第一行数据配对校验
        qty_col = qty_cols[0]
        data_row = header_row + 1
        q = _to_num(ws.cell(row=data_row, column=qty_col + 1).value)
        if not q:
            return
        price_vals = []
        for col in price_cols:
            v = _to_num(ws.cell(row=data_row, column=col + 1).value)
            if v is not None:
                price_vals.append((col, v))
        if not price_vals:
            return
        # 单价列多个但无含税/不含税区分 → 用第一个有值的（通常就一个）
        best_price_col = price_vals[0][0]

    # 确定金额列：保留「数量×选中单价 = 金额」吻合的那一列
    qty_col = qty_cols[0]
    data_row = header_row + 1
    q = _to_num(ws.cell(row=data_row, column=qty_col + 1).value)
    pv = _to_num(ws.cell(row=data_row, column=best_price_col + 1).value)
    if q and pv and amount_cols:
        best_amt_col, best_err = None, float("inf")
        for acol in amount_cols:
            av = _to_num(ws.cell(row=data_row, column=acol + 1).value)
            if av is None:
                continue
            err = abs(q * pv - av)
            if err < best_err:
                best_err = err
                best_amt_col = acol
        if best_amt_col is not None:
            ref = abs(q * pv)
            if best_err <= max(0.01, ref * 0.05):
                # 保留吻合的金额列，移除其他
                for col in amount_cols:
                    if col != best_amt_col:
                        mapping.pop(col, None)

    # 移除其他单价列，保留选定的
    for col in price_cols:
        if col != best_price_col:
            mapping.pop(col, None)


class _SheetGrid:
    """统一封装一个工作表，屏蔽 openpyxl(.xlsx) 与 xlrd(.xls) 的 API 差异。

    暴露：max_row、max_col、cell_value(row, col)   —— 均基于 1-based 行列。
    """

    def __init__(self, values, name=""):
        self.values = values          # 二维数组：values[r][c]（0-based）
        self.max_row = len(values)
        self.max_col = max((len(row) for row in values), default=0)
        self.name = name

    @property
    def max_column(self):
        """兼容 openpyxl 的 ws.max_column 命名。"""
        return self.max_col

    def cell_value(self, row, col):
        """取第 row 行、第 col 列（1-based）的值；越界返回 None。"""
        if row < 1 or col < 1 or row > self.max_row or col > self.max_col:
            return None
        return self.values[row - 1][col - 1]

    def cell(self, row, column):
        """兼容 openpyxl 的 ws.cell(row=, column=).value 用法。"""
        return _Cell(self.cell_value(row, column))


class _Cell:
    """包装单个值，提供 .value 属性以兼容 ws.cell(...).value。"""
    __slots__ = ("value",)

    def __init__(self, value):
        self.value = value


def _read_sheets(path: str) -> List[_SheetGrid]:
    """按扩展名读取文件，返回所有非空 sheet 的网格列表。

    支持 .xlsx（openpyxl）与 .xls（xlrd）。原始数据（含公式值取缓存）。
    """
    ext = os.path.splitext(path)[1].lower()
    grids = []
    if ext in (".xlsx", ".xlsm", ".xltx", ".xltm"):
        import openpyxl
        wb = openpyxl.load_workbook(path, data_only=True)
        for ws in wb.worksheets:
            vals = []
            if ws.max_row > 0 and ws.max_column > 0:
                for r in range(1, ws.max_row + 1):
                    row = []
                    for c in range(1, ws.max_column + 1):
                        row.append(ws.cell(row=r, column=c).value)
                    vals.append(row)
            grids.append(_SheetGrid(vals, ws.title))
    elif ext in (".xls", ".xlsm"):
        import xlrd
        wb = xlrd.open_workbook(path)
        for name in wb.sheet_names():
            sh = wb.sheet_by_name(name)
            vals = []
            for r in range(sh.nrows):
                vals.append([sh.cell_value(r, c) for c in range(sh.ncols)])
            grids.append(_SheetGrid(vals, name))
    else:
        raise ValueError(f"不支持的文件格式：{ext}，仅支持 .xlsx / .xls")
    return grids


def read_input(path: str, columns=None, sheet_name: Optional[str] = None,
               price_choice: Optional[str] = None) -> List[Dict[str, object]]:
    """读取输入 Excel（.xlsx/.xls），返回 [{模板列名: 值}, ...]。

    columns:     可选模板列集合（如明细导入传 config.MINGXI_COLUMNS）。
    sheet_name:  可选，指定要读取的工作表名。若文件有多个 sheet 且未指定，
                 则交互询问用户选择（暂只支持单 sheet，不做合并）。
    price_choice: 可选 '含税'/'不含税'。当表格同时有含税/不含税单价时，
                  用它直接指定；缺省则交互询问（每次询问，不记忆）。
    智能定位表头行（前 10 行内识别最多列的行）。
    """
    if not os.path.exists(path):
        raise FileNotFoundError(f"输入文件不存在: {path}")

    grids = _read_sheets(path)
    if not grids:
        raise ValueError(f"未在文件 {path} 中找到有效数据表")

    # 单 sheet 无需询问
    if len(grids) == 1:
        grid = grids[0]
    elif sheet_name:
        grid = next((g for g in grids if g.name == sheet_name), None)
        if grid is None:
            raise ValueError(f"文件中不存在工作表：{sheet_name}，可用：{ [g.name for g in grids] }")
    else:
        # 多 sheet：列出并询问
        print(f"文件含 {len(grids)} 个工作表，请选择要处理的（一次只处理一个）:")
        for i, g in enumerate(grids, 1):
            print(f"  {i}. {g.name}")
        while True:
            choice = input(f"请输入工作表编号 (1-{len(grids)})：").strip()
            if choice.isdigit() and 1 <= int(choice) <= len(grids):
                grid = grids[int(choice) - 1]
                break
            print("  输入无效，请重新输入。")

    if grid.max_row < 2 or grid.max_col < 2:
        raise ValueError(f"工作表 {grid.name} 数据过少，无法识别表头")

    return _extract_records(grid, columns, price_choice)


def _extract_records(grid, columns, price_choice=None) -> List[Dict[str, object]]:
    """从单个 sheet 网格提取数据记录。"""
    best_row, best_hits, best_mapping = None, -1, None
    for r in range(1, min(grid.max_row, 10) + 1):
        mapping = _resolve_mapping(grid, r, columns)
        hits = len(mapping)
        # 消歧（多列命中同一模板列时，用数量×单价=金额 校验 + 含税/不含税询问）
        if "商品数量" in mapping.values() and "商品单价" in mapping.values() and "金额" in mapping.values():
            _disambiguate_mapping(grid, r, mapping, price_choice)
        if hits > best_hits:
            best_hits, best_row, best_mapping = hits, r, mapping

    if best_row is None or not best_mapping:
        raise ValueError(f"无法在 {grid.name} 中识别表头")

    records = []
    for r in range(best_row + 1, grid.max_row + 1):
        rec = {}
        any_val = False
        for col, tpl_col in best_mapping.items():
            v = grid.cell_value(r, col + 1)
            if v is not None and _norm(v) != "":
                any_val = True
            rec[tpl_col] = v
        if any_val and rec.get("项目名称"):
            records.append(rec)
    return records


def normalize_records(records: List[Dict[str, object]], columns=None) -> List[Dict[str, object]]:
    """将识别的记录补齐指定列（默认开票信息维护 11 列），确保列名齐全、类型规范。"""
    columns = columns or config.KAIPIAO_COLUMNS
    out = []
    for rec in records:
        nrec = {}
        for col in columns:
            v = rec.get(col)
            if col == "商品和服务税收分类编码" and v is not None:
                v = str(int(v)) if isinstance(v, float) else str(v).strip()
            if col in ("单价", "商品单价", "金额", "商品数量") and v is not None:
                try:
                    v = round(float(v), 2)
                except (TypeError, ValueError):
                    pass
            nrec[col] = v if v is not None else ""
        out.append(nrec)
    return out