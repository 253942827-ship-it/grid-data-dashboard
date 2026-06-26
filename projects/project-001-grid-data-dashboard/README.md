# project-001-grid-data-dashboard

**项目编号：** 001
**创建日期：** 2026-06-20
**项目名称：** 网格积分数据看板

**项目说明：**
对沙太网格单元的运营数据（2026年4月~6月）进行多维度分析，生成交互式数据看板。
包含月度标签切换（4月/5月/6月）、每日明细、MTD排名、环比分析、优化建议等功能。

**数据来源：**
| 文件 | 说明 |
|------|------|
| data/网格单元运营数据_202604.xlsx | 4月每日明细（30天，105个网格） |
| data/网格单元运营数据_202605.xlsx | 5月每日明细（31天，103个网格） |
| data/网格单元运营数据_202606.xlsx | 6月每日明细（截至16日，101个网格） |
| data/网格单元业务变动统计_202606.xlsx | 6月业务变动明细 |
| data/沙太网格单元净增积分统计表202604.xlsx | 4月汇总（单行单网格） |
| data/网格单元运营数据202605.xlsx | 5月汇总（单行单网格） |

**交付物：**
| 文件 | 说明 |
|------|------|
| docs/dashboard.html | 综合数据看板（4月/5月/6月 + 环比分析 + 优化建议） |
| docs/2026-05_分析报告.html | 5月分析报告（旧版） |
| docs/2026-06-15_网格日报.html | 6月15日单日报告 |
| docs/分析报告.html | 旧版分析报告（5月） |
| data/数据分析结果.csv | CSV格式排名结果 |
| data/recommendations.json | 优化建议JSON数据 |

**分析脚本：**
| 文件 | 说明 |
|------|------|
| src/generate_dashboard.py | 综合看板生成脚本（读取3个月Excel，输出dashboard.html） |
| src/daily_report.py | 日报生成脚本（支持单日数据） |
| src/data_analysis.py | 数据分析脚本（含汇总统计、排名） |

**生成方法：**
```bash
# 生成综合看板
python3 src/generate_dashboard.py

# 每日日报
python3 src/daily_report.py data/网格单元运营数据_202606.xlsx
```

**目录结构：**
```
project-001-grid-data-dashboard/
├── README.md              # 本文件
├── src/                   # 源代码
│   ├── generate_dashboard.py
│   ├── daily_report.py
│   └── data_analysis.py
├── data/                  # 数据文件
│   ├── 网格单元运营数据_202604.xlsx
│   ├── 网格单元运营数据_202605.xlsx
│   ├── 网格单元运营数据_202606.xlsx
│   ├── 网格单元业务变动统计_202606.xlsx
│   ├── 沙太网格单元净增积分统计表202604.xlsx
│   ├── 网格单元运营数据202605.xlsx
│   ├── 数据分析结果.csv
│   └── recommendations.json
├── docs/                  # 看板和报告
│   ├── dashboard.html
│   ├── 2026-05_分析报告.html
│   ├── 2026-06-15_网格日报.html
│   └── 分析报告.html
└── outputs/               # 交付物
