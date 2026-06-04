# 工业称重业务 — 宏观经济与政策一页纸报告

一站式宏观经济数据看板，用于支撑工业称重业务的行业洞察与决策。

## 在线报告

👉 **[查看在线报告](https://tzusu_chiao.github.io/weigh-report/)**

## 数据维度

| 维度 | 数据源 | 说明 |
|------|--------|------|
| PMI | 采购经理指数 | 制造业景气度 |
| PPI | 工业生产者出厂价格指数 | 重点行业价格走势 |
| 钢铁价格 | Steel Price | 钢材市场行情 |
| 固定资产投资 | FAI | 分行业投资增速 |
| 房地产 | Real Estate | 地产核心指标 |
| 工业用电 | Electricity | 工业能耗趋势 |
| 工业中长期贷款 | Industrial Loans | 行业资金面 |
| 工业主要产品产量 | Output | 重点产品产量 |
| 货物进出口 | Import/Export | 外贸数据 |
| 社会消费品零售 | Retail Sales | 消费端数据 |
| 物流 | Logistics | 物流景气度 |
| 行业增加值 | Value Added | 行业增速 |
| CPI | 居民消费价格 | 通胀指标 |
| 线索机遇数量 | Leads | 业务机会追踪 |

## 项目结构

```
├── index.html              # 在线报告主页（Chart.js 可视化）
├── one_page_report.html    # 报告源文件
├── generate_ppt.py         # PPT 生成脚本
├── extract_data.py         # 数据提取脚本
├── read_data.py            # 数据预览脚本
├── MT.JSON                 # 主题配色方案
├── data/                   # 原始 Excel 数据
│   ├── PMI.xlsx
│   ├── PPI.xlsx
│   ├── CPI.xlsx
│   ├── Steel Price.xlsx
│   └── ... (更多数据文件)
└── 工业称重业务_一页纸报告_2026年6月.pptx  # 已生成的 PPT 报告
```

## 本地运行

```bash
# 安装依赖
pip install pandas openpyxl python-pptx

# 预览数据
python read_data.py

# 提取分析数据
python extract_data.py

# 生成 PPT 报告
python generate_ppt.py
```

直接在浏览器中打开 `index.html` 即可查看完整报告。
