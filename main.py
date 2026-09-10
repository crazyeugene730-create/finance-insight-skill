#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fin-analysis · 统一入口
================================================
平台调用约定：execute(params: dict) -> dict（必须可 JSON 序列化）

action 路由：
    calc_metrics       财务指标计算（纯计算，产出唯一可信数字）
    detect_anomalies   异常识别（纯规则，可解释、可审计）
    diagnose           生成归因诊断任务包（确定性部分由脚本完成，判断交给模型）
    generate_report    生成报告骨架（数字全部预填，定性文字留给模型）

设计原则：算归算、说归说。
    数字只在 calc_metrics 产生一次，后面三个 action 只搬运、不计算。

本地调试：
    echo '{"action":"calc_metrics","data_file":"data.csv"}' | python main.py
"""

import csv
import io
import json
import os
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

SKILL_DIR = Path(__file__).resolve().parent.parent
REFS = SKILL_DIR / "references"
NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


# ============================================================ 一、读表
def _col_idx(ref):
    m = re.match(r"([A-Z]+)", ref or "")
    if not m:
        return 0
    n = 0
    for ch in m.group(1):
        n = n * 26 + ord(ch) - 64
    return n - 1


def read_xlsx(path, sheet_index=0):
    z = zipfile.ZipFile(path)
    shared = []
    if "xl/sharedStrings.xml" in z.namelist():
        root = ET.fromstring(z.read("xl/sharedStrings.xml"))
        for si in root.findall(NS + "si"):
            shared.append("".join(t.text or "" for t in si.iter(NS + "t")))
    sheets = [n for n in z.namelist() if re.match(r"xl/worksheets/sheet\d+\.xml$", n)]
    sheets.sort(key=lambda s: int(re.search(r"(\d+)", s.split("/")[-1]).group(1)))
    if not sheets:
        raise ValueError("未找到工作表，请确认是标准 .xlsx 文件")
    root = ET.fromstring(z.read(sheets[sheet_index]))
    rows = []
    for row in root.iter(NS + "row"):
        cells = {}
        for c in row.findall(NS + "c"):
            idx = _col_idx(c.get("r"))
            t, v, isel = c.get("t"), c.find(NS + "v"), c.find(NS + "is")
            if t == "inlineStr" and isel is not None:
                val = "".join(x.text or "" for x in isel.iter(NS + "t"))
            elif v is None:
                val = ""
            elif t == "s":
                val = shared[int(v.text)]
            else:
                val = v.text or ""
            cells[idx] = val
        if cells:
            w = max(cells) + 1
            rows.append([cells.get(i, "") for i in range(w)])
    return rows


def read_csv_text(text):
    return [r for r in csv.reader(io.StringIO(text))]


def read_csv_file(path):
    for enc in ("utf-8-sig", "gbk", "utf-8"):
        try:
            with io.open(path, "r", encoding=enc, newline="") as f:
                return [r for r in csv.reader(f)]
        except UnicodeDecodeError:
            continue
    raise ValueError("CSV 编码无法识别，请另存为 UTF-8 或 GBK")


def load_rows(params):
    if params.get("data_file"):
        p = params["data_file"]
        pl = str(p).lower()
        if pl.endswith((".xlsx", ".xlsm")):
            return read_xlsx(p)
        if pl.endswith((".csv", ".txt")):
            return read_csv_file(p)
        raise ValueError("仅支持 .xlsx / .xlsm / .csv")
    if params.get("data"):
        return read_csv_text(params["data"])
    raise ValueError("缺少输入数据：请提供 data_file 或 data")


# ============================================================ 二、识别科目
ALIASES = [
    ("营业总收入", "revenue"), ("营业收入", "revenue"),
    ("营业成本", "cost"), ("税金及附加", "tax_surcharge"),
    ("销售费用", "sell_exp"), ("管理费用", "admin_exp"), ("研发费用", "rd_exp"),
    ("财务费用", "fin_exp"), ("利息费用", "interest_exp"), ("利息支出", "interest_exp"),
    ("营业利润", "op_profit"), ("利润总额", "pretax_profit"),
    ("所得税费用", "income_tax"),
    ("净利润", "net_profit"), ("归属于母公司股东的净利润", "net_profit"),
    ("归属于母公司所有者的净利润", "net_profit"),
    ("流动资产合计", "cur_assets"), ("流动负债合计", "cur_liab"),
    ("应收账款", "ar"), ("应收票据及应收账款", "ar"), ("应收账款净额", "ar"),
    ("存货", "inv"),
    ("资产总计", "total_assets"), ("资产合计", "total_assets"), ("总资产", "total_assets"),
    ("负债合计", "total_liab"), ("负债总计", "total_liab"), ("总负债", "total_liab"),
    ("所有者权益合计", "equity"), ("所有者权益", "equity"), ("股东权益合计", "equity"),
    ("净资产", "equity"), ("归属于母公司股东权益合计", "equity"),
    ("经营活动产生的现金流量净额", "op_cashflow"), ("经营活动现金流量净额", "op_cashflow"),
    ("购建固定资产、无形资产和其他长期资产支付的现金", "capex"),
]
ALIASES.sort(key=lambda x: -len(x[0]))

RAW_CN = {
    "营业收入": "revenue", "营业成本": "cost", "净利润": "net_profit",
    "应收账款": "ar", "存货": "inv", "总资产": "total_assets", "总负债": "total_liab",
    "净资产": "equity", "经营活动现金流量净额": "op_cashflow", "研发费用": "rd_exp",
}


def norm(s):
    return str(s or "").strip().lstrip("\ufeff").replace("（", "(").replace("）", ")") \
        .replace(" ", "").replace("\u3000", "").replace(":", "").replace("：", "")


def to_num(s):
    s = norm(s)
    if s in ("", "-", "--", "—", "－", "N/A", "NA", "nan"):
        return None
    neg = s.startswith("(") and s.endswith(")")
    if neg:
        s = s[1:-1]
    if s.endswith("%"):
        try:
            v = float(s[:-1].replace(",", ""))
            return -v / 100.0 if neg else v / 100.0
        except ValueError:
            return None
    s = s.replace(",", "").replace("元", "").replace("万", "")
    try:
        v = float(s)
    except ValueError:
        return None
    return -v if neg else v


def match_key(name):
    n = norm(name)
    if not n:
        return None
    for a, k in ALIASES:
        if norm(a) == n:
            return k
    for a, k in ALIASES:
        if norm(a) in n:
            return k
    return None


def find_year(s):
    m = re.search(r"(19|20)\d{2}", str(s or ""))
    return int(m.group(0)) if m else None


def parse(rows):
    """-> (years, data) ; data[year][key] = 数值"""
    for ri, row in enumerate(rows[:8]):
        ys = [find_year(c) for c in row]
        if len([y for y in ys if y]) >= 2:
            data = {y: {} for y in ys if y}
            for r in rows[ri + 1:]:
                if not r or not norm(r[0]):
                    continue
                key = match_key(norm(r[0]))
                if not key:
                    continue
                for j, y in enumerate(ys):
                    if y and j < len(r):
                        v = to_num(r[j])
                        if v is not None:
                            data[y][key] = v
            return sorted(data.keys()), data
    col_years = [find_year(r[0]) for r in rows if r]
    if len([y for y in col_years if y]) >= 2:
        header = rows[0]
        data = {}
        for r in rows[1:]:
            y = find_year(r[0]) if r else None
            if not y:
                continue
            data.setdefault(y, {})
            for j, cell in enumerate(r):
                if j == 0 or j >= len(header):
                    continue
                key = match_key(norm(header[j]))
                if key:
                    v = to_num(cell)
                    if v is not None:
                        data[y][key] = v
        return sorted(data.keys()), data
    raise ValueError("没识别到年份：请确认表头含年份（如 2023/2024/2025），或第一列是年份")


# ============================================================ 三、算指标
def _g(d, y, k):
    return d.get(y, {}).get(k)


def _avg(d, years, i, k):
    v = _g(d, years[i], k)
    if v is None:
        return None
    if i > 0:
        p = _g(d, years[i - 1], k)
        if p is not None:
            return (v + p) / 2.0
    return v


def _div(a, b):
    if a is None or b in (None, 0):
        return None
    return a / b


def compute(years, d):
    m = {}

    def put(name, i, val):
        m.setdefault(name, [None] * len(years))
        m[name][i] = val

    for i, y in enumerate(years):
        rev, cost = _g(d, y, "revenue"), _g(d, y, "cost")
        np_, ta, eq = _g(d, y, "net_profit"), _g(d, y, "total_assets"), _g(d, y, "equity")
        opcf = _g(d, y, "op_cashflow")

        if rev is not None and cost is not None:
            put("毛利率", i, (rev - cost) / rev)
        put("营业利润率", i, _div(_g(d, y, "op_profit"), rev))
        put("净利率", i, _div(np_, rev))
        put("ROE（期末净资产）", i, _div(np_, eq))
        put("ROE（平均净资产）", i, _div(np_, _avg(d, years, i, "equity")))
        put("ROA（平均总资产）", i, _div(np_, _avg(d, years, i, "total_assets")))
        put("研发费用率", i, _div(_g(d, y, "rd_exp"), rev))
        sx = _g(d, y, "sell_exp") or 0
        adm = _g(d, y, "admin_exp") or 0
        rd = _g(d, y, "rd_exp") or 0
        if rev:
            put("期间费用率", i, (sx + adm + rd) / rev)

        ar_avg, inv_avg = _avg(d, years, i, "ar"), _avg(d, years, i, "inv")
        ta_avg = _avg(d, years, i, "total_assets")
        put("应收账款周转天数", i, 365.0 / (rev / ar_avg) if rev and ar_avg else None)
        put("存货周转天数", i, 365.0 / (cost / inv_avg) if cost and inv_avg else None)
        put("总资产周转率（次）", i, _div(rev, ta_avg))
        put("总资产周转天数", i, 365.0 / (rev / ta_avg) if rev and ta_avg else None)

        put("资产负债率", i, _div(_g(d, y, "total_liab"), ta))
        ca, cl = _g(d, y, "cur_assets"), _g(d, y, "cur_liab")
        put("流动比率（倍）", i, _div(ca, cl))
        if ca is not None and cl:
            put("速动比率（倍）", i, (ca - (_g(d, y, "inv") or 0)) / cl)
        put("权益乘数", i, _div(ta, eq))
        pretax, ie = _g(d, y, "pretax_profit"), _g(d, y, "interest_exp")
        if pretax is not None and ie:
            put("利息保障倍数", i, (pretax + ie) / ie)

        put("经营现金流/净利润", i, _div(opcf, np_))
        put("经营现金流/营业收入", i, _div(opcf, rev))

        if i > 0:
            p = years[i - 1]
            for label, key in (("营业收入增长率", "revenue"), ("净利润增长率", "net_profit"),
                               ("营业成本增长率", "cost"), ("应收账款增长率", "ar"),
                               ("存货增长率", "inv"), ("总资产增长率", "total_assets"),
                               ("净资产增长率", "equity"), ("经营现金流增长率", "op_cashflow")):
                cur, prev = _g(d, y, key), _g(d, p, key)
                put(label, i, _div(cur - prev, prev) if (cur is not None and prev) else None)
    return m


PCT_NAMES = {"毛利率", "营业利润率", "净利率", "研发费用率", "期间费用率", "资产负债率",
             "总资产周转率（次）", "营业收入增长率", "净利润增长率", "营业成本增长率",
             "应收账款增长率", "存货增长率", "总资产增长率", "净资产增长率",
             "经营现金流增长率", "经营现金流/营业收入", "ROE（期末净资产）",
             "ROE（平均净资产）", "ROA（平均总资产）"}


def fmt(name, v):
    if v is None:
        return ""
    if "天数" in name:
        return "%.1f" % v
    if "（倍）" in name or "倍数" in name or "周转率（次）" in name or "经营现金流/净利润" in name:
        return "%.2f" % v
    if name in PCT_NAMES or name.endswith("率"):
        return "%.2f%%" % (v * 100)
    return "%.2f" % v


def yoy_fmt(name, v):
    if v is None:
        return ""
    if name in PCT_NAMES or name.endswith("率"):
        return "%+.2fpct" % (v * 100)
    return "%+.2f" % v


# ============================================================ 四、异常识别
def load_rules():
    p = REFS / "threshold_rules.json"
    if p.exists():
        with io.open(p, encoding="utf-8") as f:
            return json.load(f)
    return {"rules": []}


def growth_of(ctx, name, idx):
    years, metrics, data = ctx
    if name in metrics:
        v = metrics[name]
        return v[idx] if len(v) > abs(idx) or idx == -1 else None
    key = RAW_CN.get(name)
    if not key or idx == -1 and len(years) < 2:
        return None
    i = len(years) + idx if idx < 0 else idx
    if i <= 0 or i >= len(years):
        return None
    cur, prev = _g(data, years[i], key), _g(data, years[i - 1], key)
    return _div(cur - prev, prev) if (cur is not None and prev) else None


def raw_at(ctx, key, idx):
    years, metrics, data = ctx
    i = len(years) + idx if idx < 0 else idx
    if i < 0 or i >= len(years):
        return None
    return _g(data, years[i], key)


def detect(ctx, rules, industry=None):
    years, metrics, data = ctx
    if len(years) < 2:
        return [], {"error": "至少需要两年数据才能做趋势判断"}
    hits = []
    for r in rules.get("rules", []):
        cond, sev = r.get("condition"), r.get("severity")
        metric = r.get("metric")
        hit, cur, prev, delta = False, None, None, None
        try:
            if cond in ("yoy_drop_pp", "yoy_rise_pp", "yoy_rise_abs", "yoy_drop_abs"):
                m = metrics.get(metric)
                if not m:
                    continue
                cur, prev = m[-1], m[-2]
                if cur is None or prev is None:
                    continue
                delta = cur - prev
                if cond == "yoy_drop_pp":
                    hit = delta <= -(r["value"] / 100.0)
                elif cond == "yoy_rise_pp":
                    hit = delta >= (r["value"] / 100.0)
                elif cond == "yoy_rise_abs":
                    hit = delta >= r["value"]
                else:
                    hit = delta <= -r["value"]
            elif cond in ("consecutive_decrease", "consecutive_increase"):
                m = metrics.get(metric)
                k = r.get("years", 2)
                if not m or len(m) < k + 1:
                    continue
                seg = m[-(k + 1):]
                if any(v is None for v in seg):
                    continue
                hit = all(seg[i + 1] < seg[i] for i in range(k)) if cond == "consecutive_decrease" \
                    else all(seg[i + 1] > seg[i] for i in range(k))
                cur, prev = seg[-1], seg[0]
            elif cond == "consecutive_lt":
                m = metrics.get(metric)
                k = r.get("years", 2)
                if not m or len(m) < k:
                    continue
                seg = m[-k:]
                if any(v is None for v in seg):
                    continue
                hit = all(v < r["value"] for v in seg)
                cur = seg[-1]
            elif cond in ("lt", "gt"):
                m = metrics.get(metric)
                if not m:
                    continue
                cur = m[-1]
                if cur is None:
                    continue
                hit = cur < r["value"] if cond == "lt" else cur > r["value"]
            elif cond in ("growth_gap", "growth_excess"):
                a = growth_of(ctx, r["metric_a"], -1)
                b = growth_of(ctx, r["metric_b"], -1)
                if a is None or b is None:
                    continue
                diff = (b - a) if cond == "growth_gap" else (a - b)
                hit = diff >= r["value"] / 100.0
                cur, prev, delta = a, b, diff
            elif cond == "growth_mixed":
                up = growth_of(ctx, r["metric_up"], -1)
                dn = growth_of(ctx, r["metric_down"], -1)
                if up is None or dn is None:
                    continue
                hit = up > 0 and dn <= -(r["value"] / 100.0)
                cur, prev = dn, up
            elif cond == "negative_positive":
                a = raw_at(ctx, r["metric_neg"], -1)
                b = raw_at(ctx, r["metric_pos"], -1)
                if a is None or b is None:
                    continue
                hit = a < 0 and b > 0
                cur, prev = a, b
            elif cond == "consecutive_gt_ratio":
                k = r.get("years", 2)
                vals = []
                for j in range(1, k + 1):
                    a = raw_at(ctx, r["metric_a"], -j)
                    b = raw_at(ctx, r["metric_b"], -j)
                    if a is None or b is None or b <= 0:
                        vals = []
                        break
                    vals.append(a / b)
                hit = len(vals) == k and all(v > r["value"] for v in vals)
                cur = vals[0] if vals else None
            elif cond == "divergence":
                ma, mb = metrics.get(r["metric_a"]), metrics.get(r["metric_b"])
                if not ma or not mb:
                    continue
                da = (ma[-1] - ma[-2]) if (ma[-1] is not None and ma[-2] is not None) else None
                db = (mb[-1] - mb[-2]) if (mb[-1] is not None and mb[-2] is not None) else None
                if da is None or db is None:
                    continue
                hit = (da * db < 0) and abs(da) >= r["value"] / 100.0 and abs(db) >= r["value"] / 100.0
                cur, prev = da, db
        except (TypeError, KeyError, IndexError, ZeroDivisionError):
            continue
        if hit:
            hits.append({
                "规则编号": r.get("id"),
                "分类": r.get("category"),
                "指标": metric or (r.get("metric_a", "") + " vs " + r.get("metric_b", "")),
                "触发规则": r.get("desc"),
                "严重度": sev,
                "本期值": cur, "对比值": prev, "变动": delta,
                "本期值_展示": fmt(metric, cur) if metric else (("%.2f%%" % (cur * 100)) if isinstance(cur, float) else cur),
                "对比值_展示": fmt(metric, prev) if metric else (("%.2f%%" % (prev * 100)) if isinstance(prev, float) else prev),
            })
    order = {"高": 0, "中": 1, "低": 2}
    hits.sort(key=lambda x: order.get(x["严重度"], 9))
    summary = {"高": 0, "中": 0, "低": 0}
    for h in hits:
        summary[h["严重度"]] = summary.get(h["严重度"], 0) + 1
    meta = {"命中数": len(hits), "分级统计": summary, "行业": industry or "默认（未指定）"}
    if not data:
        meta["警告"] = ("未提供原始数据，依赖原始科目的规则（O3/O5/C3/C5/G3/G4 等）"
                        "已跳过。建议同时传入 data_file 以获得完整覆盖。")
    return hits, meta


# ============================================================ 五、诊断任务包
PATHS = {
    "盈利能力": ["营业成本增速 vs 营业收入增速（谁更快）", "拆成本构成：原材料 / 人工 / 制造费用 / 折旧",
                 "产品结构：高毛利业务占比是否变化", "费用端：销售/管理/研发费用率是否同步抬升",
                 "是否具备成本传导能力（提价空间）"],
    "现金流": ["差额：净利润 − 经营现金流 = 被占用部分", "定位占用项：应收增加 / 存货增加 / 应付减少",
               "应收增加 → 回款周期、客户结构、信用政策", "存货增加 → 备货策略、滞销、原材料囤积",
               "资本开支：经营现金流能否覆盖，缺口靠什么补"],
    "营运能力": ["增速剪刀差：应收/存货 vs 收入或成本", "账龄结构：1 年内占比是否下降",
                 "坏账准备 / 跌价准备计提是否充分", "客户集中度是否提高",
                 "年报中关于回款、信用政策的表述"],
    "偿债能力": ["有息负债 vs 经营性负债（谁在涨）", "债务期限结构：短债占比是否提高",
                 "利息保障倍数趋势", "货币资金能否覆盖短期有息负债", "是否存在受限资金"],
    "成长性": ["增长的质量：是否伴随现金流改善", "增长的来源：量增 / 价增 / 并购",
               "增长的持续性：订单、产能、行业景气", "与同业增速对比"],
}
FOUR_STEP = ["财务现象：（一句话 + 具体数字）", "可能原因：（数据关系 或 年报依据，必须注明是哪种）",
             "风险判断：（趋势延续会怎样）", "建议关注：（具体到科目、指标、动作）"]


def build_diagnosis_tasks(anomalies, ctx, top=None):
    years, metrics, data = ctx
    picked = [a for a in anomalies if a["严重度"] == "高"]
    picked += [a for a in anomalies if a["严重度"] == "中"]
    if top:
        picked = picked[:top]
    tasks = []
    for a in picked:
        cat = a.get("分类") or ""
        tasks.append({
            "异常编号": a.get("规则编号"),
            "分类": cat,
            "异常描述": a.get("触发规则"),
            "严重度": a.get("严重度"),
            "数据依据": {"本期值_展示": a.get("本期值_展示"), "对比值_展示": a.get("对比值_展示")},
            "深挖路径": PATHS.get(cat, ["按指标本身的驱动因素逐层拆解"]),
            "输出格式": FOUR_STEP,
        })
    return tasks


# ============================================================ 六、报告骨架
SECTIONS = [
    ("一、公司整体经营概况", ["营业收入", "净利润", "总资产"]),
    ("二、盈利能力分析", ["毛利率", "营业利润率", "净利率", "ROE（期末净资产）", "ROA（平均总资产）"]),
    ("三、营运能力分析", ["应收账款周转天数", "存货周转天数", "总资产周转率（次）"]),
    ("四、偿债能力分析", ["资产负债率", "流动比率（倍）", "速动比率（倍）", "利息保障倍数"]),
    ("五、现金流分析", ["经营现金流/净利润", "经营现金流/营业收入"]),
]


RAW_SECTION = [("营业收入", "revenue"), ("净利润", "net_profit"),
               ("总资产", "total_assets"), ("经营活动现金流量净额", "op_cashflow")]


def build_report(ctx, anomalies, diagnoses, company, years_label):
    years, metrics, data = ctx
    L = []
    L.append("# %s 财务分析报告" % (company or "（待填公司名称）"))
    L.append("")
    L.append("> 分析期间：%s　|　生成方式：fin-analysis · generate_report" % (years_label or "-".join(str(y) for y in years)))
    L.append("")

    for title, names in SECTIONS:
        L.append("## " + title)
        L.append("")
        L.append("| 指标 | " + " | ".join(str(y) for y in years) + " |")
        L.append("| ---" + "".join([" | ---"] * len(years)))
        if title.startswith("一、"):
            for cn, key in RAW_SECTION:
                vals = [_g(data, y, key) for y in years]
                if any(v is not None for v in vals):
                    L.append("| %s | %s |" % (cn, " | ".join(
                        ("%.2f" % v) if v is not None else "-" for v in vals)))
        else:
            for n in [x for x in names if x in metrics]:
                L.append("| %s | %s |" % (n, " | ".join(fmt(n, v) for v in metrics[n])))
        L.append("")
        L.append("**【待补充】** 定性判断：")
        L.append("")
        L.append("<!-- 由模型基于上表数字撰写：先给数字，再给判断 -->")
        L.append("")

    L.append("## 六、重点异常指标")
    L.append("")
    if anomalies:
        L.append("| 序号 | 异常 | 分类 | 严重度 | 触发规则 |")
        L.append("| --- | --- | --- | --- | --- |")
        for i, a in enumerate(anomalies, 1):
            L.append("| %d | %s | %s | %s | %s |" % (i, a.get("指标", ""), a.get("分类", ""),
                                                     a.get("严重度", ""), a.get("触发规则", "")))
    else:
        L.append("【待补充】未传入异常清单，请先执行 detect_anomalies。")
    L.append("")

    L.append("## 七、异常原因分析")
    L.append("")
    if diagnoses:
        for d in diagnoses:
            L.append("### 【%s】%s" % (d.get("异常编号", ""), d.get("指标", d.get("异常描述", ""))))
            L.append("")
            for k in ("财务现象", "可能原因", "风险判断", "建议关注"):
                L.append("- **%s**：%s" % (k, d.get(k, "【待补充】")))
            L.append("")
    else:
        L.append("【待补充】未传入诊断结论，请先执行 diagnose 并由模型补全四段式内容。")
    L.append("")

    L.append("## 八、经营风险提示")
    L.append("")
    L.append("**【待补充】** 分短期（1 年内）、中期（1–3 年）两层，各 2–4 条；每条写清风险 + 为什么现在提 + 观察什么信号。")
    L.append("")

    L.append("## 九、管理建议")
    L.append("")
    L.append("**【待补充】** 3–5 条，每条含：建议动作 + 依据 + 预期效果/观察指标。禁止写“加强管理”这类空话。")
    L.append("")

    L.append("---")
    L.append("")
    L.append("数据来源：%s 公开披露的年度报告。" % (company or "（待填公司名称）"))
    L.append("")
    L.append("本报告由 fin-analysis 自动生成，数据口径与计算过程可复核。")
    L.append("")
    L.append("本报告仅供研究参考，不构成任何投资建议。")
    L.append("")
    return "\n".join(L)


# ============================================================ 七、Action 实现
def _outdir(params):
    od = params.get("outdir") or os.path.join(os.getcwd(), "fin_output")
    os.makedirs(od, exist_ok=True)
    return od


def _write(name, text, od):
    p = os.path.join(od, name)
    with io.open(p, "w", encoding="utf-8") as f:
        f.write(text)
    return p


def _ctx_with_rows(params):
    """-> (ctx, rows)，rows 用于统计未识别科目"""
    rows = load_rows(params)
    years, data = parse(rows)
    return (years, compute(years, data), data), rows


def _ctx(params):
    return _ctx_with_rows(params)[0]


def _ctx_optional(params):
    """detect / diagnose / generate_report 允许只传 metrics，不重复读原始数据"""
    try:
        return _ctx(params)
    except ValueError:
        raw = params.get("metrics")
        if not raw:
            raise
        if isinstance(raw, str):
            raw = json.loads(raw)
        if isinstance(raw, dict) and "metrics" in raw:
            return (raw.get("years") or [], raw["metrics"], {})
        return (params.get("years") or [], raw, {})


def _load_metrics_json(params, ctx):
    """从 params 读取外部传入的指标表（由上一步产出）"""
    years, metrics, data = ctx
    raw = params.get("metrics")
    if raw:
        if isinstance(raw, str):
            raw = json.loads(raw)
        if isinstance(raw, dict) and "metrics" in raw:
            ys = raw.get("years") or years
            return (ys, raw["metrics"], data)
        return (params.get("years") or years, raw, data)
    return ctx


def _load_anomalies(params):
    raw = params.get("anomalies")
    if not raw:
        return []
    if isinstance(raw, str):
        raw = json.loads(raw)
    if isinstance(raw, dict):
        raw = raw.get("anomalies", [])
    return raw


def handle_calc_metrics(params):
    ctx, rows = _ctx_with_rows(params)
    years, metrics, data = ctx
    od = _outdir(params)

    lines = ["# 财务指标计算结果", "",
             "期间：%s" % " / ".join(str(y) for y in years), ""]
    lines += ["## 一、指标", ""]
    lines.append("指标 | " + " | ".join(str(y) for y in years))
    lines.append("---" + "".join([" | ---"] * len(years)))
    csv_rows = [["指标"] + [str(y) for y in years]]
    for n in metrics:
        vals = metrics[n]
        line = [n] + [fmt(n, v) for v in vals]
        csv_rows.append(line)
        lines.append(" | ".join(line))
    lines += ["", "## 二、同比变动", ""]
    lines.append("指标 | " + " | ".join(str(y) for y in years[1:]))
    lines.append("---" + "".join([" | ---"] * (len(years) - 1)))
    for n in metrics:
        vals = metrics[n]
        diffs = []
        for i in range(1, len(vals)):
            a, b = vals[i], vals[i - 1]
            diffs.append(yoy_fmt(n, a - b) if (a is not None and b is not None) else "")
        if any(diffs):
            lines.append(" | ".join([n] + diffs))
    md = "\n".join(lines)

    missed = sorted({norm(r[0]) for r in rows[1:] if r and norm(r[0]) and not match_key(norm(r[0]))})
    f1 = _write("metrics.md", md, od)
    with io.open(os.path.join(od, "metrics.csv"), "w", encoding="utf-8-sig", newline="") as f:
        csv.writer(f).writerows(csv_rows)
    _write("metrics.json", json.dumps({"years": years, "metrics": metrics}, ensure_ascii=False, indent=2), od)

    return {"success": True, "action": "calc_metrics", "years": years,
            "指标数量": len(metrics), "metrics": metrics,
            "未识别科目": missed,
            "files": [f1, os.path.join(od, "metrics.csv"), os.path.join(od, "metrics.json")],
            "next": "把 metrics.json 路径传给 detect_anomalies 继续做异常识别"}


def handle_detect_anomalies(params):
    ctx = _ctx_optional(params)
    ctx = _load_metrics_json(params, ctx)
    rules = load_rules()
    hits, summary = detect(ctx, rules, params.get("industry"))
    od = _outdir(params)
    p = _write("anomalies.json", json.dumps({"anomalies": hits, "summary": summary},
                                            ensure_ascii=False, indent=2), od)
    return {"success": True, "action": "detect_anomalies",
            "anomalies": hits, "summary": summary, "files": [p],
            "提示": "异常是否成立由阈值规则决定，不含主观判断；原因解释请用 diagnose",
            "next": "把 anomalies 原样传给 diagnose"}


def handle_diagnose(params):
    ctx = _ctx_optional(params)
    ctx = _load_metrics_json(params, ctx)
    anomalies = _load_anomalies(params)
    if not anomalies:
        rules = load_rules()
        anomalies, _ = detect(ctx, rules, params.get("industry"))
    tasks = build_diagnosis_tasks(anomalies, ctx, params.get("top"))
    instruction = (
        "以下是待完成的归因诊断任务。对每条任务按“输出格式”给出四段式内容。"
        "硬性要求：1) 不得自行计算任何指标，数字只能取自 data_evidence 或 metrics；"
        "2) 原因必须有依据——数据关系或年报原文（注明章节）；推断的必须显式标注；"
        "3) 禁止出现买入/卖出/目标价等投资建议。"
    )
    return {"success": True, "action": "diagnose",
            "instruction": instruction,
            "diagnosis_tasks": tasks,
            "年报上下文": (params.get("context") or "")[:200] + ("…（已截断）" if len(params.get("context") or "") > 200 else ""),
            "next": "模型补全四段式后，把结果数组作为 diagnoses 传给 generate_report"}


def handle_generate_report(params):
    ctx = _ctx_optional(params)
    ctx = _load_metrics_json(params, ctx)
    years, metrics, data = ctx
    anomalies = _load_anomalies(params)
    if not anomalies and params.get("auto_detect", True):
        anomalies, _ = detect(ctx, load_rules(), params.get("industry"))
    diagnoses = params.get("diagnoses") or []
    if isinstance(diagnoses, str):
        try:
            diagnoses = json.loads(diagnoses)
        except ValueError:
            diagnoses = []
    if isinstance(diagnoses, dict):
        diagnoses = diagnoses.get("diagnoses", [])

    md = build_report(ctx, anomalies, diagnoses, params.get("company"), params.get("years_label"))
    od = _outdir(params)
    p = _write("report.md", md, od)
    checklist = [
        "报告中的每个数字都能在 metrics 表中找到",
        "九段齐全且顺序正确",
        "所有高严重度异常都有四段式分析",
        "建议具体可执行，无“加强管理”类空话",
        "无投资建议类表述",
        "已注明数据来源与免责声明",
    ]
    return {"success": True, "action": "generate_report",
            "报告正文": md,
            "待补充标记数": md.count("【待补充】"),
            "自检清单": checklist,
            "files": [p],
            "next": "模型补全【待补充】部分后，按 output_format 导出最终文件"}


ACTION_MAP = {
    "calc_metrics": handle_calc_metrics,
    "detect_anomalies": handle_detect_anomalies,
    "diagnose": handle_diagnose,
    "generate_report": handle_generate_report,
}


def execute(params: dict):
    if not isinstance(params, dict):
        return {"success": False, "msg": "params 必须是 dict"}
    action = params.get("action")
    handler = ACTION_MAP.get(action)
    if not handler:
        return {"success": False, "msg": "不支持的 action：%s" % action,
                "支持的action": list(ACTION_MAP.keys())}
    try:
        return handler(params)
    except Exception as e:
        return {"success": False, "action": action, "msg": "%s: %s" % (type(e).__name__, e)}


if __name__ == "__main__":
    raw = sys.stdin.read()
    try:
        input_params = json.loads(raw) if raw.strip() else {}
    except ValueError:
        input_params = {}
    print(json.dumps(execute(input_params), ensure_ascii=False))
