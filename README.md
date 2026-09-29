# china-tax-helper

支持批量解析非标准化 Excel，一键生成全国统一规范电子税务局导入模板。
**所有数据处理均在本地完成，绝不上传任何数据。**

## 功能
- **开票信息维护**：将原始项目清单（品名/单位/单价等）转换为
  `templates/开票信息维护.xlsx` 标准格式，表头/样式与模板完全一致。
- **明细导入**：将项目明细（数量/单价/金额）转换为 `templates/明细导入.xlsx`
  标准格式，含隐藏版本标识 sheet（excelVersion/xzqhdm）原样保留。
- **表头智能识别**：自动识别输入文件列名（品名/商品名称/数量/单价/金额等均可），
  无需手动指定列映射，也支持不同表头所在行。
- **税收分类编码三级解析**：
  1. 输入自带「税收分类编码」列则直接使用；
  2. 命中本地记忆库（`data/tax_code_map.json`，精确/子串匹配）自动填入；
  3. 仍未命中则**先把未命中项目列成清单，你输入一个编码应用到全部**，
     输入后自动记忆，下次自动命中。也可用 `--tax-code` 统一兜底。
- **税率询问 + 默认记忆**（明细导入必填列）：
  首次询问本张发票税率 → 询问「是否设为默认」→ 是则存
  `data/default_tax_rate.json`，下次自动填；否则下次再问。
  也可用 `--tax-rate 0.13` 预设。
- **非必填列首次询问 + 偏好记忆**（规格型号/折扣金额/优惠政策类型/煤炭种类）：
  首次询问「是否要填」→ 不需要则记忆跳过，下次不再问；
  需要则询问统一值并记忆，下次自动填。输入里自带的列直接用不问。
- **是否含税整表统一询问**：先询问本张发票含税与否（Y/N，回车默认 N），
  将统一值应用到整表。也可用 `--tax-included Y/N` 预设免询问。

## 使用

```bash
# 开票信息维护（自动处理 inputs/ 下唯一的 Excel，交互询问缺失编码/含税）
python main.py 开票信息维护

# 开票信息维护：预设编码和不含税（不交互）
python main.py 开票信息维护 inputs/jiayin_9_month.xlsx --tax-code 1099900000000000000 --tax-included N

# 明细导入（首次交互询问税率+设为默认，询问非必填列）
python main.py 明细导入 inputs/jiayin_9_month.xlsx

# 明细导入：预设税率和编码（不交互）
python main.py 明细导入 inputs/jiayin_9_month.xlsx --tax-rate 0.13 --tax-code 1099900000000000000

# 明细导入：.xls 旧格式 + 多工作表文件，指定要处理的 sheet
python main.py 明细导入 inputs/为公.xls --sheet '2025年7月' --tax-rate 0.13 --tax-code 1099900000000000000

# 表格同时含「含税单价/未税单价」时，直接指定用含税价（否则每次交互询问）
python main.py 明细导入 inputs/佳音20260727.xlsx --sheet 'Sheet2' --price-type 含税 --tax-rate 0.13 --tax-code 1099900000000000000
```

> **格式支持**：输入支持 `.xlsx`（openpyxl）和旧格式 `.xls`（xlrd）。
> **多工作表**：文件含多个 sheet 时，默认交互询问你选择处理哪一个（一次只处理一个，
> 可用 `--sheet '工作表名'` 直接指定免询问）。
> **含税/不含税并存**：表格同时有含税与不含税单价时，默认交互询问你选用哪个
> （每次询问、不记忆），可用 `--price-type 含税|不含税` 直接指定免询问。

输出文件写入 `outputs/`，格式与对应模板一致
（表头、颜色、字体、边框、隐藏 sheet、版本标识全部保留）。

## # 税收编码数据库说明

本项目包含多个税收编码数据文件，用途不同：

## 📁 编码文件说明

### 1. `data/tax_code_official.json` - **官方标准编码库**
- **来源**：国家税务总局《商品和服务税收分类编码表》
- **用途**：给其他用户参考，提供完整的标准编码映射
- **包含**：轴承、齿轮、液压、气动、标准件等完整分类
- **状态**：仅供参考，不影响项目运行

### 2. ~~`data/tax_code_catalog.json` - 项目内置编码表~~
~~**用途**：项目的自动匹配规则（关键词匹配）~~
~~**当前状态**：简化的编码映射，主要使用 `10999` 和 `20105`~~
~~**维护者**：项目当前用户~~

### 3. `data/tax_code_map.json` - **用户记忆库**
- **用途**：用户明确指定的项目名 → 编码映射
- **优先级**：最高（优先于内置表匹配）
- **维护者**：项目当前用户

## 🎯 当前用户的使用策略

项目当前用户主要使用两个编码：
- `1099900000000000000` - 其他未列明货物（通用机械配件）
- `2010500000000000000` - 金属制品（加工件、电机座等）

**匹配优先级**：
1. 记忆库 `tax_code_map.json`（精确项目名）
2. 手动输入

## 📚 给其他用户的建议

如果你是本项目的新用户，希望使用完整的标准税收编码：

1. **直接使用官方编码表**：
   ```bash
   cp data/tax_code_official.json data/tax_code_catalog.json
   ```

2. **根据需要调整**：
   - 官方编码表可能需要根据你的业务场景调整
   - 可以将常用项目名加入记忆库 `tax_code_map.json`

## ⚠️ 重要提醒

- **不要删除** `data/tax_code_map.json`，包含用户的历史匹配记录
- 官方编码表基于国家税务总局标准，但实际使用中可能需要根据企业具体情况调整
- 当前用户主要使用两个编码：`1099900000000000000` 和 `2010500000000000000`

## 项目结构

```
china-tax-helper/
├── main.py                    # 主程序入口
├── processors/                # 核心处理模块
│   ├── kaipiao_info.py       # 开票信息处理
│   ├── mingxi_import.py      # 明细导入处理
│   ├── tax_rate_resolver.py  # 税率解析
│   └── optional_col_resolver.py # 非必填列解析
├── utils/                     # 工具模块
│   └── excel_reader.py       # Excel读取工具
├── data/                      # 数据文件
│   ├── tax_code_official.json    # 官方标准编码库（供其他用户参考）
│   ├── tax_code_official_backup.json # 备份文件
│   ├── tax_code_map.json         # 用户记忆库
│   ├── default_tax_rate.json     # 默认税率
│   └── optional_cols_prefs.json  # 非必填列偏好
├── inputs/                    # 输入文件目录
├── outputs/                   # 输出文件目录
├── tests/                     # 单元测试
├── templates/                 # Excel模板文件
├── 商品和服务税收分类编码表.xls # 原始官方编码表
└── requirements.txt           # Python依赖
```
```
config.py          # 路径、模板列名、列别名配置
main.py            # 命令行入口
utils/excel_reader.py        # 表头智能识别、读取输入
processors/kaipiao_info.py   # 开票信息维护处理器
processors/mingxi_import.py  # 明细导入处理器
processors/tax_code_resolver.py  # 税收分类编码解析+询问
processors/tax_rate_resolver.py  # 税率询问+默认记忆
processors/optional_col_resolver.py # 非必填列偏好记忆
tests/             # 单元测试
templates/         # 税务局标准模板（只读基准）
inputs/            # 原始数据（git 忽略）
outputs/           # 输出标准化 Excel（git 忽略）
data/              # 本地记忆/偏好/内置编码表（git 忽略）
```

## 隐私说明
`inputs/`、`outputs/`、`*.xlsx`、`data/`（税收编码记忆库、默认税率、非必填偏好）
均已通过 `.gitignore` 排除，不会进入版本库。所有数据处理均在本地完成。
