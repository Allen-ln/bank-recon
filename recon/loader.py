# -*- coding: utf-8 -*-
"""
数据读取与字段归一化。

设计要点（审计场景）：
1. 自动识别中英文列名，不要求客户提供固定模板；
2. "借方/贷方"在两个来源里含义相反，必须按数据来源分别解释：
   - 银行对账单：贷方 = 存款增加，借方 = 支出
   - 企业日记账：借方 = 银行存款增加，贷方 = 支出
3. 保留原始行号（source_row），任何一条匹配结果都能回溯到原表第几行，
   这是可复核性的最低要求。
"""
from __future__ import annotations

import re
from pathlib import Path

import pandas as pd

# ---------------------------------------------------------------- 列名别名表

_ALIASES: dict[str, list[str]] = {
    "date": ["交易日期", "记账日期", "日期", "发生日期", "业务日期", "入账日期",
             "date", "trade date", "value date"],
    "summary": ["摘要", "交易摘要", "对方户名", "对方单位", "用途", "备注", "说明",
                "summary", "description", "memo", "narrative"],
    "ref": ["流水号", "交易流水号", "凭证号", "凭证编号", "业务编号", "单号",
            "ref", "reference", "voucher no", "voucher"],
}

# 来源 -> (增加列别名, 减少列别名)
_SIGN_RULES = {
    "bank": (
        ["收入金额", "收入", "贷方发生额", "贷方金额", "存入金额", "进账金额", "credit", "deposit", "in"],
        ["支出金额", "支出", "借方发生额", "借方金额", "付出金额", "出账金额", "debit", "withdraw", "out"],
    ),
    "ledger": (
        ["收入金额", "收入", "借方发生额", "借方金额", "增加额", "credit", "in"],
        ["支出金额", "支出", "贷方发生额", "贷方金额", "减少额", "debit", "out"],
    ),
}

_AMOUNT_NOISE = re.compile(r"[^\d.\-]")


def _norm_header(name: str) -> str:
    """去掉空白、括号内容与全角符号，便于别名匹配。"""
    s = str(name).strip().lower()
    s = re.sub(r"[（(].*?[)）]", "", s)
    s = s.replace("　", "").replace(" ", "").replace("：", "").replace(":", "")
    return s


def _find_column(columns: list[str], aliases: list[str]) -> str | None:
    """按别名找列：先精确匹配，再包含匹配。"""
    norm = {_norm_header(c): c for c in columns}
    for alias in aliases:
        a = _norm_header(alias)
        if a in norm:
            return norm[a]
    for alias in aliases:
        a = _norm_header(alias)
        for key, original in norm.items():
            if a and a in key:
                return original
    return None


def clean_amount(value) -> float:
    """把 '¥1,234.00'、'(500)'、'1 234.50' 这类文本金额转成 float。"""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if text in {"", "-", "—", "/", "nan", "None"}:
        return 0.0
    negative = text.startswith("(") and text.endswith(")")
    text = _AMOUNT_NOISE.sub("", text)
    if text in {"", "-", ".", "-."}:
        return 0.0
    try:
        number = float(text)
    except ValueError:
        return 0.0
    return -number if negative else number


def _read_raw(path: str | Path) -> pd.DataFrame:
    """按扩展名读取 csv / xlsx，全部按文本读入，避免编号被转成数字丢前导零。"""
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"找不到文件：{path}")
    suffix = path.suffix.lower()
    if suffix in {".xlsx", ".xlsm"}:
        return pd.read_excel(path, dtype=str, sheet_name=0)
    if suffix == ".xls":
        return pd.read_excel(path, dtype=str, sheet_name=0)
    if suffix in {".csv", ".txt"}:
        for encoding in ("utf-8-sig", "gbk", "utf-8"):
            try:
                return pd.read_csv(path, dtype=str, encoding=encoding)
            except UnicodeDecodeError:
                continue
        return pd.read_csv(path, dtype=str, encoding="utf-8", errors="replace")
    raise ValueError(f"不支持的文件类型：{suffix}（请使用 .xlsx / .xls / .csv）")


def load_statement(path: str | Path, source: str = "bank") -> pd.DataFrame:
    """读取流水/日记账，返回统一结构的 DataFrame。

    Parameters
    ----------
    path : 文件路径
    source : "bank"（银行对账单）或 "ledger"（企业银行存款日记账）

    返回列：row_id, source_row, date, summary, amount, ref
    amount 为正表示资金流入（收款），为负表示流出（付款）。
    """
    if source not in _SIGN_RULES:
        raise ValueError('source 只能是 "bank" 或 "ledger"')

    raw = _read_raw(path)
    columns = [str(c) for c in raw.columns]

    date_col = _find_column(columns, _ALIASES["date"])
    summary_col = _find_column(columns, _ALIASES["summary"])
    ref_col = _find_column(columns, _ALIASES["ref"])
    in_aliases, out_aliases = _SIGN_RULES[source]
    in_col = _find_column(columns, in_aliases)
    out_col = _find_column(columns, out_aliases)
    amount_col = _find_column(columns, ["金额", "交易金额", "发生额", "amount"])

    if date_col is None:
        raise ValueError(
            f"[{Path(path).name}] 找不到日期列。当前列名：{columns}\n"
            "请在表头中使用「交易日期 / 记账日期 / 日期」等常见名称。"
        )
    if in_col is None and out_col is None and amount_col is None:
        raise ValueError(
            f"[{Path(path).name}] 找不到金额列。当前列名：{columns}\n"
            "请提供「收入金额/支出金额」两列，或「借方发生额/贷方发生额」，或单列「金额」。"
        )

    result = pd.DataFrame(index=raw.index)
    result["row_id"] = [f"R{i}" for i in range(1, len(raw) + 1)]
    # Excel 第 1 行是表头，数据从第 2 行开始，便于人工回溯
    result["source_row"] = list(range(2, len(raw) + 2))
    result["date"] = pd.to_datetime(
        raw[date_col].astype(str).str.strip().str.replace("/", "-", regex=False),
        errors="coerce",
    )
    result["summary"] = (
        raw[summary_col].astype(str).str.strip() if summary_col else ""
    )
    result["ref"] = raw[ref_col].astype(str).str.strip() if ref_col else ""

    if in_col or out_col:
        inflow = raw[in_col].map(clean_amount) if in_col else 0.0
        outflow = raw[out_col].map(clean_amount) if out_col else 0.0
        result["amount"] = (inflow - outflow).round(2)
    else:
        result["amount"] = raw[amount_col].map(clean_amount).round(2)

    # 丢掉完全空白的行（Excel 常带尾部空行）
    result = result[result["amount"].notna()].copy()
    result = result[result["date"].notna() | (result["amount"] != 0)].copy()
    result["source"] = source
    return result.reset_index(drop=True)


def describe(df: pd.DataFrame) -> str:
    """给命令行输出用的数据概况。"""
    if df.empty:
        return "0 条记录"
    start, end = df["date"].min(), df["date"].max()
    inflow = df.loc[df["amount"] > 0, "amount"].sum()
    outflow = -df.loc[df["amount"] < 0, "amount"].sum()
    return (
        f"{len(df)} 条记录，期间 {start:%Y-%m-%d} ~ {end:%Y-%m-%d}，"
        f"流入 {inflow:,.2f}，流出 {outflow:,.2f}"
    )
