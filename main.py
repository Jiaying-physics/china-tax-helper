# -*- coding: utf-8 -*-
"""china-tax-helper 命令行入口。

用法示例：
  python main.py 开票信息维护 inputs/jiayin_9_month.xlsx
  python main.py 开票信息维护 inputs/jiayin_9_month.xlsx --tax-code 1099900000000000000
  python main.py 明细导入 inputs/jiayin_9_month.xlsx --tax-rate 0.13
"""
import argparse
import os
import sys

from processors.kaipiao_info import build_kaipiao
from processors.mingxi_import import build_mingxi


def _resolve_input(args) -> str:
    """定位输入文件：传入路径，或自动取 inputs/ 下第一个 xlsx。"""
    if args.input:
        return os.path.abspath(args.input)
    inp_dir = os.path.abspath("inputs")
    candidates = []
    if os.path.isdir(inp_dir):
        for fn in sorted(os.listdir(inp_dir)):
            if fn.lower().endswith((".xlsx", ".xls")):
                candidates.append(os.path.join(inp_dir, fn))
    if not candidates:
        print("⚠ 未指定输入文件，且 inputs/ 目录下没有 Excel 文件。")
        print("   用法：python main.py [开票信息维护|明细导入] [输入文件]")
        sys.exit(1)
    if len(candidates) > 1:
        print("inputs/ 下存在多个文件，请显式指定：")
        for c in candidates:
            print("  ", c)
        sys.exit(1)
    return candidates[0]


def main():
    parser = argparse.ArgumentParser(description="税务局标准模板转换工具（本地处理，不上传）")
    parser.add_argument("task", nargs="?", default="开票信息维护",
                        help="任务：开票信息维护 / 明细导入")
    parser.add_argument("input", nargs="?", help="输入 Excel 路径（缺省自动找 inputs/）")
    parser.add_argument("--tax-code", default=None,
                        help="为所有未提供编码的项目统一填入此 19 位税收分类编码")
    parser.add_argument("--tax-included", default=None, choices=["Y", "N"],
                        help="整张发票是否含税：Y 含税 / N 不含税（缺省则交互询问一次）")
    parser.add_argument("--tax-rate", default=None,
                        help="明细导入：统一税率，如 0.13 或 13（缺省则询问并可选记忆为默认）")
    parser.add_argument("--sheet", default=None,
                        help="指定要处理的工作表名（多 sheet 文件用；缺省则交互询问）")
    parser.add_argument("--price-type", default=None, choices=["含税", "不含税"],
                        help="表格同时含含税/不含税单价时选用哪个（缺省则每次交互询问）")
    args = parser.parse_args()

    if args.task not in ("开票信息维护", "明细导入"):
        print(f"暂不支持的任务：{args.task}，支持：开票信息维护 / 明细导入")
        sys.exit(1)

    input_path = _resolve_input(args)

    if args.task == "开票信息维护":
        build_kaipiao(input_path, default_code=args.tax_code, tax_included=args.tax_included,
                      sheet_name=args.sheet, price_choice=args.price_type)
    else:
        build_mingxi(input_path, tax_rate=args.tax_rate, code_override=args.tax_code,
                     sheet_name=args.sheet, price_choice=args.price_type)


if __name__ == "__main__":
    main()

