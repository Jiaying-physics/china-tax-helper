# -*- coding: utf-8 -*-
"""全局配置：路径、模板列名、列别名映射。所有数据处理均在本地完成。"""
import os
from datetime import date

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")
INPUTS_DIR = os.path.join(BASE_DIR, "inputs")
OUTPUTS_DIR = os.path.join(BASE_DIR, "outputs")
DATA_DIR = os.path.join(BASE_DIR, "data")

# 开票信息维护模板文件名（只读基准，永不修改）
KAIPIAO_TEMPLATE_FILE = "开票信息维护.xlsx"
KAIPIAO_OUTPUT_PREFIX = "开票信息维护_"

# 明细导入模板文件名（只读基准，永不修改）
MINGXI_TEMPLATE_FILE = "明细导入.xlsx"
MINGXI_OUTPUT_PREFIX = "明细导入_"

# 明细导入主 sheet 名
MINGXI_SHEET = "1-明细模板"

# 开票信息维护主 sheet 名
KAIPIAO_SHEET = "项目信息导入"

# 模板三层结构位置：行1说明、行2填写提示、行3表头、行4起数据
KAIPIAO_TITLE_ROW = 1
KAIPIAO_HINT_ROW = 2
KAIPIAO_HEADER_ROW = 3
KAIPIAO_DATA_START = 4
MINGXI_DATA_START = 4

# 开票信息维护 11 列（按模板表头顺序）
KAIPIAO_COLUMNS = [
    "项目名称",
    "项目分类",
    "商品和服务税收分类编码",
    "规格型号",
    "单位",
    "单价",
    "是否含税",
    "简码",
    "税率",
    "是否使用优惠政策",
    "优惠政策类型",
]

# 明细导入 11 列（A~K 有内容，L~Q 为空，R/S 涉税专业服务）
MINGXI_COLUMNS = [
    "项目名称",          # A
    "商品和服务税收分类编码",  # B
    "规格型号",          # C
    "单位",            # D
    "商品数量",         # E
    "商品单价",         # F
    "金额",            # G
    "税率",            # H
    "折扣金额",         # I
    "优惠政策类型",      # J
    "煤炭种类",         # K
]

# 明细导入必填列
MINGXI_REQUIRED = ["项目名称", "金额", "税率"]

# 明细导入非必填列（首次询问是否要填，要填则询问统一值并记忆）
MINGXI_OPTIONAL_COLS = ["规格型号", "折扣金额", "优惠政策类型", "煤炭种类"]

# 输入文件「列关键词」→ 模板列名 的映射，用于表头智能识别。
# 匹配规则：表头「包含」某个关键词即命中（如「入库数量」含「数量」→商品数量）。
# 因此换词/加前后缀也能识别，不必穷举所有写法。
HEADER_KEYWORDS = {
    "项目名称": ["项目名称", "品名", "商品名称", "货名", "货物名称", "名称", "商品"],
    "项目分类": ["项目分类", "分类"],
    "商品和服务税收分类编码": ["税收分类编码", "商品编码", "分类编码", "税收编码", "商编", "商品和服务税收分类编码", "编码"],
    "规格型号": ["规格型号", "规格", "型号"],
    "单位": ["单位", "计量"],
    "单价": ["单价"],
    "商品单价": ["单价", "价格", "售价"],
    "商品数量": ["数量", "个数", "件数", "数量数"],
    "金额": ["金额", "合计", "价税", "总价", "总额", "价额", "货款"],
    "税率": ["税率", "税点"],
    "折扣金额": ["折扣", "优惠金额", "让利"],
    "是否含税": ["含税", "是否含税", "含税标志"],
    "简码": ["简码"],
    "是否使用优惠政策": ["是否使用优惠政策", "是否优惠", "优惠政策"],
    "优惠政策类型": ["优惠政策类型", "优惠类型"],
}


def datetime_stamp() -> str:
    """返回当前日期字符串 YYYYMMDD"""
    return date.today().strftime("%Y%m%d")


def output_file_name(input_file: str, prefix: str = KAIPIAO_OUTPUT_PREFIX) -> str:
    """由输入文件名生成输出文件名：前缀 + 输入文件名 + 日期"""
    base = os.path.splitext(os.path.basename(input_file))[0]
    return f"{prefix}{base}_{datetime_stamp()}.xlsx"