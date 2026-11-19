# -*- coding: utf-8 -*-
"""生成模拟银行流水与账面日记账（全部为虚构数据，不含任何真实客户信息）。

刻意制造审计现场常见的四类"脏数据"：
1. 尾差      银行扣了手续费，账面记的是原始付款额（相差 0.01~2 元）
2. 时差      跨日入账，1~3 天不等
3. 摘要改写  账面把"杭州晨曦贸易有限公司"写成"晨曦贸易"、"货款"写成"收货款"
4. 单向记录  银行已收企业未收 / 企业已记银行未达
运行：python sample_data/generate_sample.py
"""
from __future__ import annotations

import random
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd

random.seed(20250101)

OUT_DIR = Path(__file__).resolve().parent

COMPANIES = [
    "杭州晨曦贸易有限公司", "宁波远舟物流有限公司", "苏州恒泰机械制造有限公司",
    "上海知微信息技术有限公司", "嘉兴南湖包装材料有限公司", "绍兴越诚纺织有限公司",
    "合肥中衡检测技术有限公司", "南京金鼎工程咨询有限公司", "温州港盛进出口有限公司",
    "湖州长兴electronic科技有限公司", "台州宏信模具有限公司", "金华万邦供应链有限公司",
    "义乌小商品城贸易有限公司", "杭州云栖数据服务有限公司", "常州华通电气有限公司",
]

INCOME_MEMO = ["货款", "销售回款", "服务费", "项目尾款", "预收货款", "租金收入"]
EXPENSE_MEMO = ["采购付款", "运费", "咨询服务费", "办公用品", "差旅费报销", "设备款"]


def _rand_date() -> datetime:
    start = datetime(2025, 12, 1)
    return start + timedelta(days=random.randint(0, 30), hours=random.randint(0, 0))


def _base_transactions(n: int) -> list[dict]:
    """生成 n 对基准交易（银行与账面共用同一业务事实）。"""
    items = []
    for _ in range(n):
        direction = "in" if random.random() < 0.55 else "out"
        company = random.choice(COMPANIES)
        memo = random.choice(INCOME_MEMO if direction == "in" else EXPENSE_MEMO)
        if direction == "in":
            amount = round(random.uniform(3000, 480000), 2)
            summary = f"{memo}-{company}"
        else:
            # 少量整数支出，制造金额重复，考验分配算法
            amount = round(random.uniform(800, 260000), 2) if random.random() > 0.25 else float(random.choice([10000, 20000, 50000, 100000]))
            summary = f"{memo}-{company}"
        items.append({
            "date": _rand_date(),
            "direction": direction,
            "amount": amount,
            "summary": summary,
            "company": company,
            "memo": memo,
        })
    return items


def build():
    base = _base_transactions(860)

    bank_rows, ledger_rows = [], []

    for idx, t in enumerate(base, start=1):
        # ---- 银行侧
        bank_amount = t["amount"]
        bank_date = t["date"]
        bank_rows.append({
            "交易日期": bank_date.strftime("%Y-%m-%d"),
            "摘要": t["summary"],
            "收入金额": f"{bank_amount:,.2f}" if t["direction"] == "in" else "",
            "支出金额": f"{bank_amount:,.2f}" if t["direction"] == "out" else "",
            "交易流水号": f"BK2025{idx:06d}",
        })

        # 有 4% 的银行交易企业当期漏记（银行已收/付，企业未记）
        if random.random() < 0.04:
            continue

        ledger_amount = bank_amount
        ledger_date = bank_date
        ledger_summary = t["summary"]

        # 尾差：付款被银行扣手续费
        if t["direction"] == "out" and random.random() < 0.12:
            ledger_amount = round(bank_amount - round(random.uniform(0.01, 2.0), 2), 2)

        # 时差：1~3 天入账
        if random.random() < 0.15:
            ledger_date = bank_date + timedelta(days=random.randint(1, 3))

        # 摘要改写
        r = random.random()
        if r < 0.10:
            ledger_summary = t["summary"].replace(t["company"], t["company"].replace("有限公司", "")[2:])
        elif r < 0.18:
            ledger_summary = f"{t['memo']}（{t['company']}）"
        elif r < 0.24:
            ledger_summary = f"{t['summary']} 已核销"

        ledger_rows.append({
            "记账日期": ledger_date.strftime("%Y-%m-%d"),
            "摘要": ledger_summary,
            "借方发生额": f"{ledger_amount:,.2f}" if t["direction"] == "in" else "",   # 企业账：借方=存款增加
            "贷方发生额": f"{ledger_amount:,.2f}" if t["direction"] == "out" else "",  # 企业账：贷方=减少
            "凭证号": f"记-{idx:04d}",
        })

    # 银行侧额外：6% 银行已发生、企业完全未见（银行已收企业未收）
    for k in range(1, 56):
        direction = "in" if random.random() < 0.4 else "out"
        amount = round(random.uniform(5000, 90000), 2)
        d = _rand_date()
        company = random.choice(COMPANIES)
        summary = f"{'利息收入' if direction == 'in' and random.random()<0.3 else ('货款' if direction=='in' else '付款')}-{company}"
        bank_rows.append({
            "交易日期": d.strftime("%Y-%m-%d"),
            "摘要": summary,
            "收入金额": f"{amount:,.2f}" if direction == "in" else "",
            "支出金额": f"{amount:,.2f}" if direction == "out" else "",
            "交易流水号": f"BK2025U{k:05d}",
        })

    # 账面侧额外：企业已记、银行未达（或记账错误）
    for k in range(1, 40):
        direction = "in" if random.random() < 0.5 else "out"
        amount = round(random.uniform(3000, 120000), 2)
        d = datetime(2025, 12, 29) + timedelta(days=random.randint(0, 2))
        company = random.choice(COMPANIES)
        summary = f"{'预收货款' if direction == 'in' else '预付账款'}-{company}"
        ledger_rows.append({
            "记账日期": d.strftime("%Y-%m-%d"),
            "摘要": summary,
            "借方发生额": f"{amount:,.2f}" if direction == "in" else "",
            "贷方发生额": f"{amount:,.2f}" if direction == "out" else "",
            "凭证号": f"记-U{k:04d}",
        })

    bank_df = pd.DataFrame(bank_rows).sort_values("交易日期").reset_index(drop=True)
    ledger_df = pd.DataFrame(ledger_rows).sort_values("记账日期").reset_index(drop=True)

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    bank_path = OUT_DIR / "银行流水.xlsx"
    ledger_path = OUT_DIR / "银行存款日记账.xlsx"
    bank_df.to_excel(bank_path, index=False)
    ledger_df.to_excel(ledger_path, index=False)
    print(f"已生成：{bank_path.name}（{len(bank_df)} 条）、{ledger_path.name}（{len(ledger_df)} 条）")
    return bank_path, ledger_path


if __name__ == "__main__":
    build()
