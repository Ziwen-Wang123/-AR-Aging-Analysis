#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
应收账款账龄分析工具
Author: Ziwen Wang
Date: 2026-06
Description: 模拟三大区域150+客户应收账款数据，实现账龄自动分层、
             逾期风险识别及Excel可视化报告输出
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from datetime import datetime, timedelta
import random

# ── 中文字体设置（防止乱码）───────────────────────
matplotlib.rcParams['font.family'] = 'Microsoft YaHei'
matplotlib.rcParams['axes.unicode_minus'] = False

# ══════════════════════════════════════════════════
# 第一步：生成模拟数据
# ══════════════════════════════════════════════════
np.random.seed(42)
random.seed(42)

# 三大区域、各8个客户，共24个客户
regions = {
    "华北区": ["北京科技有限公司", "天津制造集团", "河北钢铁贸易", "山东机械设备",
               "内蒙古能源科技", "石家庄电子商务", "青岛海洋科技", "太原煤炭集团"],
    "华东区": ["上海贸易集团", "苏州工业制造", "杭州互联网科技", "南京电子有限公司",
               "宁波港口贸易", "无锡半导体", "合肥家电集团", "温州轻工业"],
    "华南区": ["广州进出口贸易", "深圳科技有限公司", "东莞制造集团", "佛山陶瓷贸易",
               "厦门港口物流", "福州电子科技", "珠海精密仪器", "汕头纺织集团"]
}

records = []
today = datetime(2026, 6, 3)

for region, clients in regions.items():
    for client in clients:
        # 每个客户生成6-8张发票
        num_invoices = random.randint(6, 8)
        for _ in range(num_invoices):
            days_ago = random.randint(1, 180)
            invoice_date = today - timedelta(days=days_ago)

            # 集团类客户金额更大，模拟真实业务场景
            if "集团" in client:
                amount = round(random.uniform(50000, 300000), 2)
            else:
                amount = round(random.uniform(10000, 80000), 2)

            # 账龄越长，已收款比例越高（时间越久催收效果越好）
            if days_ago <= 30:
                status = random.choices(["已收款", "未收款"], weights=[30, 70])[0]
            elif days_ago <= 60:
                status = random.choices(["已收款", "未收款"], weights=[55, 45])[0]
            elif days_ago <= 90:
                status = random.choices(["已收款", "未收款"], weights=[75, 25])[0]
            else:
                status = random.choices(["已收款", "未收款"], weights=[85, 15])[0]

            records.append({
                "区域": region, "客户名": client,
                "发票日期": invoice_date.strftime("%Y-%m-%d"),
                "金额": amount, "收款状态": status
            })

df = pd.DataFrame(records)
print(f"✓ 共生成 {len(df)} 条记录，涵盖 {df['客户名'].nunique()} 个客户")

# ══════════════════════════════════════════════════
# 第二步：账龄计算与风险分层
# ══════════════════════════════════════════════════
today_ts = pd.Timestamp("2026-06-03")
df["发票日期"] = pd.to_datetime(df["发票日期"])

# 计算每张发票距今天数
df["账龄天数"] = (today_ts - df["发票日期"]).dt.days

def aging_bucket(days):
    """将账龄天数映射为四个风险区间"""
    if days <= 30:   return "0-30天"
    elif days <= 60: return "31-60天"
    elif days <= 90: return "61-90天"
    else:            return "90天以上"

def risk_level(days):
    """根据账龄天数判断风险等级"""
    if days <= 30:   return "正常"
    elif days <= 60: return "关注"
    elif days <= 90: return "预警"
    else:            return "高风险"

df["账龄分层"] = df["账龄天数"].apply(aging_bucket)
df["风险等级"] = df["账龄天数"].apply(risk_level)

# 筛选未收款记录
df_unpaid = df[df["收款状态"] == "未收款"].copy()
df_unpaid["发票日期"] = df_unpaid["发票日期"].dt.strftime("%Y-%m-%d")
print(f"✓ 未收款记录：{len(df_unpaid)} 条，合计金额：¥{df_unpaid['金额'].sum():,.2f}")

# ══════════════════════════════════════════════════
# 第三步：可视化——账龄分布柱状图
# ══════════════════════════════════════════════════
pivot = df_unpaid.pivot_table(
    values="金额", index="账龄分层", aggfunc="sum"
).reindex(["0-30天", "31-60天", "61-90天", "90天以上"])

fig, ax = plt.subplots(figsize=(8, 5))

# 用绿→黄→橙→红表示风险递增
colors = ["#4CAF50", "#FFC107", "#FF9800", "#F44336"]
bars = ax.bar(pivot.index, pivot["金额"], color=colors, width=0.5)

# 柱子顶部显示金额
for bar, val in zip(bars, pivot["金额"]):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 2000,
            f"¥{val:,.0f}", ha="center", va="bottom", fontsize=10)

ax.set_title("应收账款账龄分析（未收款）", fontsize=14, fontweight="bold")
ax.set_xlabel("账龄分层", fontsize=12)
ax.set_ylabel("未收款金额（元）", fontsize=12)
ax.yaxis.set_major_formatter(
    matplotlib.ticker.FuncFormatter(lambda x, _: f"¥{x/10000:.0f}万"))
plt.tight_layout()
plt.savefig("账龄分析图.png", dpi=150, bbox_inches="tight")
plt.show()
print("✓ 图表已保存：账龄分析图.png")

# ══════════════════════════════════════════════════
# 第四步：输出Excel报告（带颜色风险预警）
# ══════════════════════════════════════════════════
wb = openpyxl.Workbook()

# ── Sheet1：未收款明细 ────────────────────────────
ws1 = wb.active
ws1.title = "未收款明细"

# 四种风险等级对应的背景色
fills = {
    "正常":   PatternFill("solid", fgColor="C8E6C9"),  # 绿
    "关注":   PatternFill("solid", fgColor="FFF9C4"),  # 黄
    "预警":   PatternFill("solid", fgColor="FFE0B2"),  # 橙
    "高风险": PatternFill("solid", fgColor="FFCDD2"),  # 红
}
header_fill = PatternFill("solid", fgColor="1565C0")
header_font = Font(color="FFFFFF", bold=True)

headers = ["区域", "客户名", "发票日期", "金额", "账龄天数", "账龄分层", "风险等级"]
for col, h in enumerate(headers, 1):
    cell = ws1.cell(row=1, column=col, value=h)
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center")

for row_idx, row in enumerate(df_unpaid[headers].values.tolist(), 2):
    risk = row[6]
    for col_idx, val in enumerate(row, 1):
        cell = ws1.cell(row=row_idx, column=col_idx, value=val)
        cell.fill = fills[risk]
        cell.alignment = Alignment(horizontal="center")

col_widths = [10, 20, 14, 14, 10, 12, 10]
for col, width in enumerate(col_widths, 1):
    ws1.column_dimensions[openpyxl.utils.get_column_letter(col)].width = width
ws1.freeze_panes = "A2"  # 冻结首行，方便查看

# ── Sheet2：账龄汇总 ──────────────────────────────
ws2 = wb.create_sheet("账龄汇总")

summary_data = df_unpaid.groupby("账龄分层").agg(
    笔数=("金额", "count"),
    未收款金额=("金额", "sum")
).reindex(["0-30天", "31-60天", "61-90天", "90天以上"]).reset_index()
summary_data["未收款金额"] = summary_data["未收款金额"].round(2)

for col, h in enumerate(summary_data.columns, 1):
    cell = ws2.cell(row=1, column=col, value=h)
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center")

summary_fills = ["C8E6C9", "FFF9C4", "FFE0B2", "FFCDD2"]
for row_idx, row in enumerate(summary_data.values.tolist(), 2):
    f = PatternFill("solid", fgColor=summary_fills[row_idx - 2])
    for col_idx, val in enumerate(row, 1):
        cell = ws2.cell(row=row_idx, column=col_idx, value=val)
        cell.fill = f
        cell.alignment = Alignment(horizontal="center")

# 合计行
total_row = len(summary_data) + 2
ws2.cell(row=total_row, column=1, value="合计").font = Font(bold=True)
ws2.cell(row=total_row, column=2, value=int(summary_data["笔数"].sum())).font = Font(bold=True)
ws2.cell(row=total_row, column=3, value=round(summary_data["未收款金额"].sum(), 2)).font = Font(bold=True)
for col in range(1, 4):
    ws2.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 16

wb.save("应收账款账龄分析报告.xlsx")
print("✓ Excel报告已生成：应收账款账龄分析报告.xlsx")
