# fin-analysis · 财务分析 Skill

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python](https://img.shields.io/badge/Python-3.8%2B-blue.svg)](https://www.python.org/)
[![Dependencies](https://img.shields.io/badge/dependencies-none-brightgreen.svg)](requirements.txt)

一个 Skill 名字，四个 action，覆盖财报分析完整链路：**算指标 → 找异常 → 做归因 → 出报告**。

采用「单 Skill 多 Action」打包模式，符合豆包 / 火山 AgentKit 规范。零第三方依赖，任意 Python 3 直接运行。

---

## 它解决什么问题

传统财务分析里最耗时的三件事：把数据抠进 Excel 手拉公式、盯着一堆指标找哪里不对劲、把结论写成报告。这个 Skill 把前两件完全自动化，第三件把骨架和数字全部准备好，只留判断和表达给人（或模型）。

| 环节 | 原来 | 用了之后 |
|---|---|---|
| 算 30 个指标 | 手拉公式，容易错 | 一次跑完，口径永远一致 |
| 找异常 | 靠眼睛看，容易漏 | 30 条规则自动扫描，按严重度排序 |
| 分析原因 | 翻年报找说法 | 给出深挖路径 + 四段式输出格式 |
| 写报告 | 从空白页开始 | 九段骨架 + 数字全部预填 |

---

## 四个 Action

| action | 脚本做什么 | 模型做什么 |
|---|---|---|
| `calc_metrics` | 算 27 个指标 + 同比，输出 CSV / Markdown / JSON | 确认年份与科目识别是否正确 |
| `detect_anomalies` | 按 30 条阈值规则输出异常清单（含严重度、触发规则） | 判断是否需按行业回调阈值 |
| `diagnose` | 组装诊断任务包：异常 + 深挖路径 + 四段式模板 | **补全四段式内容（核心判断环节）** |
| `generate_report` | 生成九段骨架，数字全部预填，免责声明自动生成 | 补写标注了 `【待补充】` 的定性文字 |

### 设计原则：算归算，说归说

**数字只在 `calc_metrics` 产生一次，后面三个环节只搬运、不计算。**

财务指标是确定性的，容不得大模型心算；而读文字、找关联、下判断恰恰是模型擅长的。所以脚本承担所有确定性工作（算数、比对、校验、预填骨架），模型只负责判断与表达。

---

## 快速开始

### 1. 准备数据

参考 `references/sample_input.csv`，格式是：第一行表头（第一列科目名，其余列为年份），金额单位统一。

```csv
项目,2023,2024,2025
营业收入,245000,280000,315000
营业成本,185000,216000,249000
净利润,32000,35000,37000
...
```

三个高频坑：**不要合并单元格**、**科目名不要重复**（"应收账款"和"应收账款净额"只留一个）、**年度列要能被识别出年份**。

### 2. 调用

```json
{"name":"fin-analysis","params":{"action":"calc_metrics","data_file":"data.xlsx"}}
{"name":"fin-analysis","params":{"action":"detect_anomalies","data_file":"data.xlsx"}}
{"name":"fin-analysis","params":{"action":"diagnose","data_file":"data.xlsx","anomalies":"<上一步结果>"}}
{"name":"fin-analysis","params":{"action":"generate_report","data_file":"data.xlsx","company":"示例公司"}}
```

### 3. 本地调试

```bash
# 单个 action
echo '{"action":"calc_metrics","data_file":"references/sample_input.csv"}' | python scripts/main.py

# 四个 action 串联自测
python tests/run_tests.py
```

> 调用 `detect_anomalies` / `diagnose` / `generate_report` 时建议每次都带 `data_file`。只传 `metrics` 时，依赖原始科目的规则（O3/O5/C3/C5/G3/G4 等）会被跳过，返回结果里会带 `警告` 字段提示。

---

## 目录结构

```
fin-analysis/
├─ SKILL.md                         元数据 + 参数定义 + action 枚举
├─ scripts/
│  └─ main.py                       唯一入口，execute(params) 路由分发
├─ references/
│  ├─ threshold_rules.json          30 条异常阈值规则（可独立调整，不用改代码）
│  ├─ metric_definitions.md         指标口径定义表
│  ├─ diagnosis_framework.md        四段式归因框架 + 四类异常深挖路径
│  ├─ report_template.md            九段报告模板 + 好坏示例对照
│  └─ sample_input.csv              数据填空模板
├─ tests/run_tests.py               本地自测
├─ requirements.txt                 无第三方依赖声明
└─ .skillignore                     打包排除规则
```

---

## 打包成 zip

解压后第一层必须是 skill 同名文件夹，`SKILL.md` 在该文件夹根目录：

```bash
cd ..
zip -r fin-analysis.zip fin-analysis -x "*/__pycache__/*" "*/fin_output/*" "*.pyc"
```

| 自检项 | 要求 |
|---|---|
| 目录层级 | zip 第一层是 `fin-analysis/`，`SKILL.md` 在其根目录 |
| 命名一致 | `SKILL.md` 的 `name` = 文件夹名，小写 + 连字符 |
| action 枚举 | `parameters` 里 `action` 必须写 `enum` |
| 入口函数 | `main.py` 实现 `execute(params)`，返回可 JSON 序列化的 dict |
| 干净打包 | 排除 `__pycache__` / `.pyc` / 临时文件 |

---

## 扩展新的 Action

不用改 `execute`，三步搞定：

1. `scripts/main.py` 新增 `handle_xxx(params)` 函数
2. `ACTION_MAP` 字典加一行 `"xxx": handle_xxx`
3. `SKILL.md` 的 `enum` 数组加上 `"xxx"`

---

## 调整阈值

`references/threshold_rules.json` 里 30 条规则都是通用的默认起点，**必须按行业和公司特性回调**：

| 行业 | 建议 |
|---|---|
| 重资产（制造、能源、地产） | 放宽周转天数、资产负债率阈值 |
| 强周期（化工、有色、航运） | 放宽毛利率、净利率波动阈值 |
| 高周转零售 / 快消 | 收紧存货周转天数阈值 |
| 项目制 / 工程类 | 放宽应收周转天数，收紧现金流阈值 |
| 高研发（医药、半导体） | 放宽费用率阈值 |

回调完记得更新文件里的 `version` 和 `updated` 字段。

---

## 边界与免责

- 只覆盖通用财务指标，不做行业专属指标（EV/EBITDA、同店增速等）
- 周转率用期初期末平均；首年无上年数据退化为期末口径
- 不做合并报表调整与会计政策差异还原
- 无法识别财务造假，只能提示"数据与常识背离，需重点核实"
- **不提供任何投资建议**，不输出买入/卖出/目标价
- 生成报告的数据应使用上市公司**已公开披露**的信息，不得使用未公开内部数据

---

## License

MIT License
