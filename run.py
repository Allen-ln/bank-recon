# -*- coding: utf-8 -*-
"""杨帆 - 银行存款对账工具入口。

用法：
    python run.py --bank sample_data/银行流水.xlsx \
                  --ledger sample_data/银行存款日记账.xlsx \
                  --out output/对账结果.xlsx
"""
from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path

from recon.loader import describe, load_statement
from recon.matcher import MatchConfig, match
from recon.report import write_report


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="银行存款流水与账面日记账自动对账（审计辅助工具）",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    p.add_argument("--bank", required=True, help="银行对账单文件（xlsx/xls/csv）")
    p.add_argument("--ledger", required=True, help="企业银行存款日记账文件（xlsx/xls/csv）")
    p.add_argument("--out", default="output/对账结果.xlsx", help="输出 xlsx 路径")
    p.add_argument("--date-window", type=int, default=3, help="允许的日期差（天）")
    p.add_argument("--amount-tolerance", type=float, default=0.02, help="金额容差（元）")
    p.add_argument("--summary-threshold", type=float, default=60.0, help="摘要相似度阈值 0-100")
    p.add_argument("--bank-balance", type=float, default=None, help="银行对账单期末余额（可选，用于调节表）")
    p.add_argument("--ledger-balance", type=float, default=None, help="企业日记账期末余额（可选，用于调节表）")
    return p


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)

    t0 = time.perf_counter()
    print("① 读取数据 …")
    bank = load_statement(args.bank, source="bank")
    ledger = load_statement(args.ledger, source="ledger")
    print(f"   银行对账单：{describe(bank)}")
    print(f"   账面日记账：{describe(ledger)}")

    cfg = MatchConfig(
        date_window=args.date_window,
        amount_tolerance=args.amount_tolerance,
        summary_threshold=args.summary_threshold,
    )

    print("② 三阶段匹配 …")
    result = match(bank, ledger, cfg)
    elapsed = time.perf_counter() - t0

    stats = result.stats
    print("③ 输出底稿 …")
    out = write_report(
        result, args.out,
        meta={
            "bank_file": args.bank,
            "ledger_file": args.ledger,
            "bank_balance": args.bank_balance,
            "ledger_balance": args.ledger_balance,
        },
    )

    print("\n" + "=" * 56)
    print("对账完成")
    print("=" * 56)
    print(f"银行 {stats['bank_count']} 条 / 账面 {stats['ledger_count']} 条")
    print(f"匹配 {stats['matched_pairs']} 对（可确认 {stats['high_confidence']}，需复核 {stats['medium_confidence']}）")
    print(f"银行有账面无 {stats['bank_only']} 条，账面有银行无 {stats['ledger_only']} 条")
    print(f"记录匹配率 {stats['record_match_rate']}%，金额覆盖率 {stats['amount_coverage']}%")
    print(f"耗时 {elapsed:.2f} 秒")
    print(f"输出：{Path(out).resolve()}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
