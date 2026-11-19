# -*- coding: utf-8 -*-
"""
三阶段对账匹配引擎。

阶段 1 精确金额    金额分到分相同 → 组内做一对一最优分配（匈牙利算法）
阶段 2 容差金额    金额差异 ≤ 容差（默认 0.02 元，典型为银行扣手续费）
阶段 3 近似匹配    金额差异比例 ≤ 阈值、日期差 ≤ 窗口、摘要相似度 ≥ 阈值

每个阶段只在"上一阶段剩下的未匹配记录"之间进行，不会重复占用。

设计原则：宁可漏配，不可乱配。
低于 min_score 的候选直接放弃、退回"未匹配"，交由审计人员人工处理——
把两笔无关交易撮合成一对，对底稿的伤害远大于多留几条人工核对。
"""
from __future__ import annotations

import math
from dataclasses import dataclass, field

import numpy as np
import pandas as pd
from rapidfuzz import fuzz

try:  # scipy 缺失时退化为贪心，保证工具在任何环境都能跑
    from scipy.optimize import linear_sum_assignment

    _HAS_SCIPY = True
except Exception:  # pragma: no cover
    _HAS_SCIPY = False


@dataclass
class MatchConfig:
    """匹配参数，全部可在命令行覆盖。"""

    date_window: int = 3            # 允许的日期差（天）
    amount_tolerance: float = 0.02  # 阶段 2 金额容差（元）
    fuzzy_amount_ratio: float = 0.05  # 阶段 3 金额差异比例上限
    summary_threshold: float = 60.0   # 阶段 3 摘要相似度下限（0-100）
    high_confidence: float = 80.0     # 综合分达到多少才算"可直接确认"
    min_score: float = 45.0           # 低于此分一律不撮合，退回人工
    summary_for_high: float = 40.0    # 高置信还要求摘要相似度不低于此值


@dataclass
class MatchResult:
    matched: pd.DataFrame = field(default_factory=pd.DataFrame)
    bank_only: pd.DataFrame = field(default_factory=pd.DataFrame)
    ledger_only: pd.DataFrame = field(default_factory=pd.DataFrame)
    stats: dict = field(default_factory=dict)


def _summary_score(a: str, b: str) -> float:
    """摘要相似度：字符级 / 词序无关 / 局部匹配取较大值。"""
    a, b = (a or "").strip(), (b or "").strip()
    if not a or not b:
        return 0.0
    return max(
        fuzz.ratio(a, b),
        fuzz.token_set_ratio(a, b),
        fuzz.partial_ratio(a, b) * 0.85,
    )


def _pair_quality(bank: pd.Series, ledger: pd.Series, cfg: MatchConfig) -> dict:
    """单对记录的质量分与解释。

    评分：金额 40 + 日期 35 + 摘要 25。
    金额项用「相对误差」而非绝对容差——10 万元差 1 元是手续费，
    应当是高分；而 100 元差 1 元才是真正可疑的差异。
    """
    amount_gap = abs(float(bank["amount"]) - float(ledger["amount"]))
    max_abs = max(abs(float(bank["amount"])), abs(float(ledger["amount"])), 1.0)

    day_gap = 999
    if pd.notna(bank["date"]) and pd.notna(ledger["date"]):
        day_gap = abs((bank["date"] - ledger["date"]).days)

    sim = _summary_score(bank["summary"], ledger["summary"])

    rel_tol = max(cfg.amount_tolerance, 0.0015 * max_abs, 0.01)
    amount_part = 40.0 * max(0.0, 1.0 - amount_gap / rel_tol)

    if day_gap <= cfg.date_window:
        date_part = 35.0 * (1.0 - day_gap / (cfg.date_window + 1))
    elif day_gap <= 7:
        date_part = 10.0 * (1.0 - (day_gap - cfg.date_window) / 4.0)
    else:
        date_part = 0.0

    summary_part = 25.0 * (sim / 100.0)
    score = round(amount_part + date_part + summary_part, 1)

    # 匹配依据：给审计人员看懂为什么配在一起
    basis = []
    if amount_gap <= 0.001:
        basis.append("金额一致")
    elif amount_gap <= cfg.amount_tolerance:
        basis.append(f"尾差 {amount_gap:.2f} 元")
    else:
        basis.append(f"金额差 {amount_gap:,.2f} 元（{amount_gap / max_abs:.2%}）")
    basis.append("日期相同" if day_gap == 0 else (f"日期差 {day_gap} 天" if day_gap < 900 else "日期缺失"))
    if sim >= 90:
        basis.append("摘要一致")
    elif sim >= cfg.summary_threshold:
        basis.append("摘要近似")
    else:
        basis.append("摘要不符")

    return {
        "amount_gap": round(amount_gap, 2),
        "day_gap": day_gap,
        "summary_similarity": round(sim, 1),
        "score": score,
        "basis": "、".join(basis),
    }


def _assign(quality: np.ndarray) -> list[tuple[int, int]]:
    """在候选矩阵上做一对一最优分配。

    用匈牙利算法最大化总得分；scipy 不可用时退化为按分数降序贪心。
    """
    if quality.size == 0:
        return []
    if quality.shape[0] == 1 or quality.shape[1] == 1:
        flat = int(quality.argmax())
        return [(flat // quality.shape[1], flat % quality.shape[1])]

    if _HAS_SCIPY:
        rows, cols = linear_sum_assignment(-quality.astype(float))
        return [(int(r), int(c)) for r, c in zip(rows, cols)]

    pairs, used_r, used_c = [], set(), set()
    for idx in np.argsort(-quality, axis=None):
        r, c = divmod(int(idx), quality.shape[1])
        if r in used_r or c in used_c:
            continue
        pairs.append((r, c))
        used_r.add(r)
        used_c.add(c)
        if len(pairs) == min(quality.shape):
            break
    return pairs


def _bucket_keys(df: pd.DataFrame, kind: str) -> pd.Series:
    """分桶键，避免全表 O(n²) 比较。

    exact : 分到分
    tol   : 容差桶（用于阶段 2，配合 ±1 邻桶）
    log   : 对数桶，约 5% 相对精度（用于阶段 3）
    """
    if kind == "exact":
        return df["amount"].round(2)
    if kind == "tol":
        step = 0.02
        return (df["amount"] / step).round().astype("Int64")
    sign = np.sign(df["amount"].to_numpy())
    mag = np.log1p(np.abs(df["amount"].to_numpy())) * 20
    return pd.Series(np.round(sign * mag).astype(np.int64), index=df.index)


def _run_stage(
    bank_pool: pd.DataFrame,
    ledger_pool: pd.DataFrame,
    cfg: MatchConfig,
    *,
    kind: str,
    neighbors: tuple[int, ...],
    candidate_filter,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """在一个分桶集合里跑一轮最优分配。

    返回 (新增匹配, 剩余银行, 剩余账面)
    """
    if bank_pool.empty or ledger_pool.empty:
        return pd.DataFrame(), bank_pool, ledger_pool

    bank_keys = _bucket_keys(bank_pool, kind)
    ledger_keys = _bucket_keys(ledger_pool, kind)

    matched_rows: list[dict] = []
    consumed_bank: set[int] = set()
    consumed_ledger: set[int] = set()

    for key, b_group in bank_pool.groupby(bank_keys, sort=False):
        wanted = {key + d for d in neighbors}
        candidates = ledger_pool[ledger_keys.isin(wanted)]
        candidates = candidates.drop(index=[i for i in candidates.index if i in consumed_ledger])
        b_group = b_group.drop(index=[i for i in b_group.index if i in consumed_bank])
        if candidates.empty or b_group.empty:
            continue

        b_idx, l_idx = list(b_group.index), list(candidates.index)
        quality = np.full((len(b_idx), len(l_idx)), -1.0)
        details: dict[tuple[int, int], dict] = {}
        for i, bi in enumerate(b_idx):
            for j, lj in enumerate(l_idx):
                b_row, l_row = bank_pool.loc[bi], ledger_pool.loc[lj]
                if not candidate_filter(b_row, l_row):
                    continue
                info = _pair_quality(b_row, l_row, cfg)
                if info["score"] < cfg.min_score:
                    continue  # 护栏：低分不撮合
                quality[i, j] = info["score"]
                details[(i, j)] = info

        if quality.max() < 0:
            continue

        for i, j in _assign(quality):
            if quality[i, j] < 0:
                continue
            bi, lj = b_idx[i], l_idx[j]
            if bi in consumed_bank or lj in consumed_ledger:
                continue
            b_row, l_row = bank_pool.loc[bi], ledger_pool.loc[lj]
            info = details[(i, j)]
            confident = (info["score"] >= cfg.high_confidence
                         and info["summary_similarity"] >= cfg.summary_for_high)
            matched_rows.append({
                "tier": "high" if confident else "medium",
                "bank_row": int(b_row["source_row"]),
                "ledger_row": int(l_row["source_row"]),
                "bank_date": b_row["date"],
                "ledger_date": l_row["date"],
                "amount_bank": round(float(b_row["amount"]), 2),
                "amount_ledger": round(float(l_row["amount"]), 2),
                "amount_gap": info["amount_gap"],
                "day_gap": info["day_gap"],
                "bank_summary": b_row["summary"],
                "ledger_summary": l_row["summary"],
                "summary_similarity": info["summary_similarity"],
                "score": info["score"],
                "basis": info["basis"],
                "bank_ref": b_row["ref"],
                "ledger_ref": l_row["ref"],
            })
            consumed_bank.add(bi)
            consumed_ledger.add(lj)

    new_matches = pd.DataFrame(matched_rows)
    left_bank = bank_pool.drop(index=list(consumed_bank), errors="ignore")
    left_ledger = ledger_pool.drop(index=list(consumed_ledger), errors="ignore")
    return new_matches, left_bank, left_ledger


def match(bank: pd.DataFrame, ledger: pd.DataFrame, cfg: MatchConfig | None = None) -> MatchResult:
    """执行三阶段匹配，返回 MatchResult。"""
    cfg = cfg or MatchConfig()
    bank = bank.copy()
    ledger = ledger.copy()

    def date_gap_ok(b, l, window: int, hard_limit: int = 30) -> bool:
        if pd.isna(b["date"]) or pd.isna(l["date"]):
            return True  # 缺日期时交给摘要与金额判断
        return abs((b["date"] - l["date"]).days) <= hard_limit

    def amount_close(b, l, tol: float) -> bool:
        return abs(float(b["amount"]) - float(l["amount"])) <= tol

    stages: list[pd.DataFrame] = []

    # ---------------- 阶段 1：精确金额
    m1, bank_left, ledger_left = _run_stage(
        bank, ledger, cfg, kind="exact", neighbors=(0,),
        candidate_filter=lambda b, l: date_gap_ok(b, l, cfg.date_window),
    )
    stages.append(m1)

    # ---------------- 阶段 2：容差金额（手续费、尾差）
    m2, bank_left, ledger_left = _run_stage(
        bank_left, ledger_left, cfg, kind="tol", neighbors=(-1, 0, 1),
        candidate_filter=lambda b, l: amount_close(b, l, cfg.amount_tolerance)
        and date_gap_ok(b, l, cfg.date_window),
    )
    stages.append(m2)

    # ---------------- 阶段 3：近似匹配（金额比例 + 摘要）
    m3, bank_left, ledger_left = _run_stage(
        bank_left, ledger_left, cfg, kind="log", neighbors=(-2, -1, 0, 1, 2),
        candidate_filter=lambda b, l: abs(float(b["amount"]) - float(l["amount"]))
        <= max(cfg.fuzzy_amount_ratio * max(abs(float(b["amount"])), abs(float(l["amount"]))), 1.0)
        and date_gap_ok(b, l, cfg.date_window)
        and _summary_score(b["summary"], l["summary"]) >= cfg.summary_threshold,
    )
    stages.append(m3)

    has_any = any(not s.empty for s in stages)
    matched = (pd.concat([s for s in stages if not s.empty], ignore_index=True)
               if has_any else pd.DataFrame())
    if not matched.empty:
        matched = matched.sort_values(
            ["tier", "score"], ascending=[True, False]
        ).reset_index(drop=True)
        matched["confidence"] = matched["tier"].map({"high": "可直接确认", "medium": "需人工复核"})

    bank_only = bank_left.copy()
    ledger_only = ledger_left.copy()

    # ---------------- 统计
    total = len(bank) + len(ledger)
    hit = len(matched) * 2
    bank_total_amt = float(bank["amount"].abs().sum())
    bank_hit_amt = float(matched["amount_bank"].abs().sum()) if not matched.empty else 0.0
    stats = {
        "bank_count": len(bank),
        "ledger_count": len(ledger),
        "matched_pairs": len(matched),
        "high_confidence": int((matched["tier"] == "high").sum()) if not matched.empty else 0,
        "medium_confidence": int((matched["tier"] == "medium").sum()) if not matched.empty else 0,
        "bank_only": len(bank_only),
        "ledger_only": len(ledger_only),
        "record_match_rate": round(hit / total * 100, 1) if total else 0.0,
        "amount_coverage": round(bank_hit_amt / bank_total_amt * 100, 1) if bank_total_amt else 0.0,
        "bank_only_amount": round(float(bank_only["amount"].sum()), 2) if not bank_only.empty else 0.0,
        "ledger_only_amount": round(float(ledger_only["amount"].sum()), 2) if not ledger_only.empty else 0.0,
        "date_window": cfg.date_window,
        "amount_tolerance": cfg.amount_tolerance,
        "summary_threshold": cfg.summary_threshold,
        "min_score": cfg.min_score,
        "engine": "hungarian" if _HAS_SCIPY else "greedy",
    }

    return MatchResult(matched=matched, bank_only=bank_only, ledger_only=ledger_only, stats=stats)
