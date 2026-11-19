# -*- coding: utf-8 -*-
"""
对账结果输出：一份带格式的 xlsx 底稿。

6 个 sheet：
  01_对账汇总      指标 + 参数 + 结论说明
  02_已匹配明细    可直接确认的一对一匹配
  03_待复核        金额/日期有尾差、摘要不一致，需人工判断
  04_银行有账面无  银行已发生、企业未入账（含未达账项）
  05_账面无银行有  企业已记账、银行未见（含未达账项）
  06_未达账项调节表 草稿（提供余额时可自动勾平）
"""
from __future__ import annotations

from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties

from .matcher import MatchResult

# ------------------------------------------------------------------ 样式常量
_TITLE_FONT = Font(name="微软雅黑", size=14, bold=True, color="1F3864")
_HEAD_FONT = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
_HEAD_FILL = PatternFill("solid", fgColor="1F3864")
_BODY_FONT = Font(name="微软雅黑", size=10)
_BOLD_FONT = Font(name="微软雅黑", size=10, bold=True)
_GREY = PatternFill("solid", fgColor="F2F2F2")
_GREEN = PatternFill("solid", fgColor="E2EFDA")
_YELLOW = PatternFill("solid", fgColor="FFF2CC")
_RED = PatternFill("solid", fgColor="FCE4E4")
_THIN = Side(style="thin", color="D9D9D9")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_MONEY_FMT = "#,##0.00;[Red]-#,##0.00"
_DATE_FMT = "yyyy-mm-dd"


def _write_table(ws, df: pd.DataFrame, start_row: int, *, widths: dict[str, int] | None = None,
                 money_cols: tuple[str, ...] = (), date_cols: tuple[str, ...] = (),
                 fill_rule=None) -> int:
    """把 DataFrame 写成带表头样式的表，返回下一个可用行号。"""
    if df.empty:
        ws.cell(row=start_row, column=1, value="（无记录）").font = _BODY_FONT
        return start_row + 2

    headers = list(df.columns)
    for j, name in enumerate(headers, start=1):
        cell = ws.cell(row=start_row, column=j, value=name)
        cell.font, cell.fill, cell.border = _HEAD_FONT, _HEAD_FILL, _BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(j)].width = (widths or {}).get(name, 14)

    for i, row in enumerate(df.to_dict("records"), start=start_row + 1):
        row_fill = fill_rule(row) if fill_rule else None
        for j, name in enumerate(headers, start=1):
            value = row[name]
            if isinstance(value, pd.Timestamp):
                value = value.to_pydatetime() if pd.notna(value) else None
            if isinstance(value, float) and pd.isna(value):
                value = None
            cell = ws.cell(row=i, column=j, value=value)
            cell.font, cell.border = _BODY_FONT, _BORDER
            if name in money_cols and isinstance(value, (int, float)):
                cell.number_format = _MONEY_FMT
            elif name in date_cols:
                cell.number_format = _DATE_FMT
                cell.alignment = Alignment(horizontal="center")
            if row_fill:
                cell.fill = row_fill
    return start_row + len(df) + 2


def write_report(result: MatchResult, out_path: str | Path, *, meta: dict | None = None) -> Path:
    """生成 xlsx 底稿。"""
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    meta = meta or {}
    stats = result.stats

    wb = Workbook()

    # ---------------------------------------------------------- 01 汇总
    ws = wb.active
    ws.title = "01_对账汇总"
    ws.sheet_view.showGridLines = False
    ws.merge_cells("A1:E1")
    ws["A1"] = "银行存款对账结果汇总"
    ws["A1"].font = _TITLE_FONT
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 28
    ws["A2"] = f"生成时间：{datetime.now():%Y-%m-%d %H:%M}　文件：{Path(meta.get('bank_file', '')).name} / {Path(meta.get('ledger_file', '')).name}"
    ws["A2"].font = Font(name="微软雅黑", size=9, color="808080")
    ws.merge_cells("A2:E2")

    rows = [
        ("银行流水条数", stats["bank_count"], "条", "银行对账单原始记录数"),
        ("账面记录条数", stats["ledger_count"], "条", "企业银行存款日记账原始记录数"),
        ("自动匹配对数", stats["matched_pairs"], "对", "含高置信与需复核两部分"),
        ("　其中：可确认", stats["high_confidence"], "对", "金额一致、日期在窗口内"),
        ("　其中：需复核", stats["medium_confidence"], "对", "存在尾差或摘要不一致，见 sheet 03"),
        ("银行有·账面无", stats["bank_only"], "条", f"金额合计 {stats['bank_only_amount']:,.2f} 元"),
        ("账面有·银行无", stats["ledger_only"], "条", f"金额合计 {stats['ledger_only_amount']:,.2f} 元"),
        ("记录匹配率", stats["record_match_rate"], "%", "已匹配记录数 / 总记录数"),
        ("金额覆盖率", stats["amount_coverage"], "%", "已匹配的银行流水金额 / 银行流水总金额"),
    ]
    ws["A4"], ws["B4"], ws["C4"], ws["D4"] = "指标", "数值", "单位", "说明"
    for j in range(1, 5):
        cell = ws.cell(row=4, column=j)
        cell.font, cell.fill, cell.border = _HEAD_FONT, _HEAD_FILL, _BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for i, (label, value, unit, note) in enumerate(rows, start=5):
        ws.cell(row=i, column=1, value=label).font = _BOLD_FONT if not label.startswith("　") else _BODY_FONT
        v = ws.cell(row=i, column=2, value=value)
        v.font, v.alignment = _BOLD_FONT, Alignment(horizontal="right")
        ws.cell(row=i, column=3, value=unit).font = _BODY_FONT
        ws.cell(row=i, column=4, value=note).font = _BODY_FONT
        for j in range(1, 5):
            ws.cell(row=i, column=j).border = _BORDER
        if i in (5 + 6, 5 + 5):  # 两个"仅单向"指标标黄提示
            for j in range(1, 5):
                ws.cell(row=i, column=j).fill = _YELLOW
            ws.cell(row=i, column=1).fill = _YELLOW
            ws.cell(row=i, column=2).fill = _YELLOW
    # 覆盖掉上面被误覆盖的单元格底色（保持表头区干净）
    ws["B7"].fill = _GREEN
    ws["B6"].fill = _GREEN

    text_row = 16
    ws.cell(row=text_row, column=1, value="执行参数").font = _BOLD_FONT
    params = [
        ("日期匹配窗口", f"±{stats['date_window']} 天"),
        ("金额容差", f"±{stats['amount_tolerance']:.2f} 元"),
        ("摘要相似度阈值", f"{stats.get('summary_threshold', 60)} / 100"),
        ("最低撮合分", f"{stats.get('min_score', 45)} / 100"),
        ("分配算法", "匈牙利最优分配" if stats["engine"] == "hungarian" else "贪心分配"),
    ]
    for i, (k, v) in enumerate(params, start=text_row + 1):
        ws.cell(row=i, column=1, value=k).font = _BODY_FONT
        ws.cell(row=i, column=2, value=v).font = _BODY_FONT

    note_row = text_row + 6
    ws.cell(row=note_row, column=1, value="结论与说明").font = _BOLD_FONT
    conclusion = (
        f"本次共比对 {stats['bank_count']} 条银行流水与 {stats['ledger_count']} 条账面记录，"
        f"自动撮合 {stats['matched_pairs']} 对，金额覆盖率 {stats['amount_coverage']}%；"
        f"其中 {stats['medium_confidence']} 对存在尾差或摘要差异，建议逐笔复核。"
        f"另有 {stats['bank_only']} 条仅见于银行、{stats['ledger_only']} 条仅见于账面，"
        f"需区分「未达账项」与「记账差错」。"
    )
    ws.merge_cells(start_row=note_row + 1, start_column=1, end_row=note_row + 1, end_column=4)
    c = ws.cell(row=note_row + 1, column=1, value=conclusion)
    c.font, c.alignment = _BODY_FONT, Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[note_row + 1].height = 58

    ws.merge_cells(start_row=note_row + 3, start_column=1, end_row=note_row + 4, end_column=4)
    warn = ws.cell(
        row=note_row + 3, column=1,
        value="⚠ 免责说明：本表为机器初筛结果，仅供审计人员减少机械核对工作量。"
              "「需复核」与「未匹配」项目必须由审计人员逐笔核查原始凭证后，方可作为审计底稿结论。"
              "工具无法识别一借多贷、合并付款、跨期调整等复杂情形。",
    )
    warn.font = Font(name="微软雅黑", size=9, color="C00000")
    warn.alignment = Alignment(wrap_text=True, vertical="top")
    warn.fill = _RED

    for col, width in zip("ABCDE", (20, 14, 7, 46, 4)):
        ws.column_dimensions[col].width = width

    # ---------------------------------------------------------- 02/03 明细
    detail_head = ["bank_row", "ledger_row", "bank_date", "ledger_date", "amount_bank",
                   "amount_ledger", "amount_gap", "day_gap", "bank_summary", "ledger_summary",
                   "summary_similarity", "score", "basis", "confidence", "bank_ref", "ledger_ref"]
    renamed = {
        "bank_row": "银行原表行号", "ledger_row": "账面原表行号",
        "bank_date": "银行日期", "ledger_date": "账面日期",
        "amount_bank": "银行金额", "amount_ledger": "账面金额",
        "amount_gap": "金额差异", "day_gap": "日期差(天)",
        "bank_summary": "银行摘要", "ledger_summary": "账面摘要",
        "summary_similarity": "摘要相似度", "score": "综合得分", "basis": "匹配依据",
        "confidence": "置信判定", "bank_ref": "银行流水号", "ledger_ref": "凭证号",
    }
    widths = {"银行摘要": 30, "账面摘要": 30, "匹配依据": 34, "银行流水号": 18, "凭证号": 14,
              "银行日期": 12, "账面日期": 12, "置信判定": 12}
    matched = result.matched
    high = matched[matched["tier"] == "high"] if not matched.empty else matched
    medium = matched[matched["tier"] == "medium"] if not matched.empty else matched

    ws2 = wb.create_sheet("02_已匹配明细")
    ws2.sheet_view.showGridLines = False
    if not high.empty:
        high = high[detail_head].rename(columns=renamed)
    _write_table(ws2, high, 1, widths=widths,
                 money_cols=("银行金额", "账面金额", "金额差异"),
                 date_cols=("银行日期", "账面日期"))
    ws2.freeze_panes = "A2"
    ws2.auto_filter.ref = f"A1:{get_column_letter(len(high.columns)) if not high.empty else 1}{max(len(high) + 1, 1)}"

    ws3 = wb.create_sheet("03_待复核")
    ws3.sheet_view.showGridLines = False
    if not medium.empty:
        medium = medium[detail_head].rename(columns=renamed)
    _write_table(ws3, medium, 1, widths=widths,
                 money_cols=("银行金额", "账面金额", "金额差异"),
                 date_cols=("银行日期", "账面日期"))
    ws3.freeze_panes = "A2"
    ws3.auto_filter.ref = f"A1:{get_column_letter(len(medium.columns)) if not medium.empty else 1}{max(len(medium) + 1, 1)}"

    # ---------------------------------------------------------- 04/05 未匹配
    def _unmatched_frame(df: pd.DataFrame) -> pd.DataFrame:
        if df.empty:
            return df
        out = df[["source_row", "date", "summary", "amount", "ref"]].copy()
        out.columns = ["原表行号", "日期", "摘要", "金额", "流水号/凭证号"]
        out["方向"] = out["金额"].map(lambda v: "收款" if v > 0 else "付款")
        return out[["原表行号", "日期", "摘要", "方向", "金额", "流水号/凭证号"]]

    ws4 = wb.create_sheet("04_银行有账面无")
    ws4.sheet_view.showGridLines = False
    _write_table(ws4, _unmatched_frame(result.bank_only), 1,
                 widths={"摘要": 36, "流水号/凭证号": 20, "日期": 12},
                 money_cols=("金额",), date_cols=("日期",))
    ws4.freeze_panes = "A2"

    ws5 = wb.create_sheet("05_账面无银行有")
    ws5.sheet_view.showGridLines = False
    _write_table(ws5, _unmatched_frame(result.ledger_only), 1,
                 widths={"摘要": 36, "流水号/凭证号": 20, "日期": 12},
                 money_cols=("金额",), date_cols=("日期",))
    ws5.freeze_panes = "A2"

    # ---------------------------------------------------------- 06 调节表
    ws6 = wb.create_sheet("06_未达账项调节表")
    ws6.sheet_view.showGridLines = False
    ws6.merge_cells("A1:D1")
    ws6["A1"] = "银行存款余额调节表（草稿）"
    ws6["A1"].font = _TITLE_FONT
    ws6.row_dimensions[1].height = 26

    bank_only = result.bank_only
    ledger_only = result.ledger_only
    # 银行已收企业未收：银行有、账面无且为收款
    bank_received = round(float(bank_only.loc[bank_only["amount"] > 0, "amount"].sum()), 2) if not bank_only.empty else 0.0
    bank_paid = round(float(-bank_only.loc[bank_only["amount"] < 0, "amount"].sum()), 2) if not bank_only.empty else 0.0
    ledger_received = round(float(ledger_only.loc[ledger_only["amount"] > 0, "amount"].sum()), 2) if not ledger_only.empty else 0.0
    ledger_paid = round(float(-ledger_only.loc[ledger_only["amount"] < 0, "amount"].sum()), 2) if not ledger_only.empty else 0.0

    bank_balance = meta.get("bank_balance")
    ledger_balance = meta.get("ledger_balance")

    ws6["A3"], ws6["B3"] = "项目", "金额（元）"
    for cell in (ws6["A3"], ws6["B3"]):
        cell.font, cell.fill, cell.border = _HEAD_FONT, _HEAD_FILL, _BORDER
        cell.alignment = Alignment(horizontal="center")

    lines = [
        ("企业银行存款日记账余额", ledger_balance if ledger_balance is not None else "（未提供，请手工填列）"),
        ("加：银行已收，企业未收", bank_received),
        ("减：银行已付，企业未付", bank_paid),
        ("调节后余额 A", None),
        ("", ""),
        ("银行对账单余额", bank_balance if bank_balance is not None else "（未提供，请手工填列）"),
        ("加：企业已收，银行未收", ledger_received),
        ("减：企业已付，银行未付", ledger_paid),
        ("调节后余额 B", None),
        ("", ""),
        ("差额（A - B）", None),
    ]
    for i, (label, value) in enumerate(lines, start=4):
        a = ws6.cell(row=i, column=1, value=label)
        a.font, a.border = (_BOLD_FONT if label.startswith("调节后") or label.startswith("差额") else _BODY_FONT), _BORDER
        b = ws6.cell(row=i, column=2, value=value)
        b.font, b.border = _BODY_FONT, _BORDER
        if isinstance(value, (int, float)):
            b.number_format = _MONEY_FMT
        if label == "调节后余额 A" and ledger_balance is not None:
            b.value = round(ledger_balance + bank_received - bank_paid, 2)
            b.font = _BOLD_FONT
        if label == "调节后余额 B" and bank_balance is not None:
            b.value = round(bank_balance + ledger_received - ledger_paid, 2)
            b.font = _BOLD_FONT
        if label == "差额（A - B）" and ledger_balance is not None and bank_balance is not None:
            diff = round((ledger_balance + bank_received - bank_paid) - (bank_balance + ledger_received - ledger_paid), 2)
            b.value = diff
            b.font = _BOLD_FONT
            b.fill = _GREEN if abs(diff) < 0.01 else _RED

    ws6["A16"] = ("说明：未达账项按收付方向自动归集。调节后 A、B 应相等；"
                  "若差额不为零，说明存在记账错误或遗漏，需进一步核查。")
    ws6["A16"].font = Font(name="微软雅黑", size=9, color="808080")
    ws6["A16"].alignment = Alignment(wrap_text=True, vertical="top")
    ws6.merge_cells("A16:D16")
    ws6.row_dimensions[16].height = 34
    ws6.column_dimensions["A"].width = 30
    ws6.column_dimensions["B"].width = 26

    # 打印/导出 PDF 时横向、按页宽缩放，避免宽表被裁切
    for sheet in wb.worksheets:
        sheet.sheet_properties.tabColor = "1F3864"
        sheet.page_setup.orientation = "landscape"
        sheet.page_setup.fitToWidth = 1
        sheet.page_setup.fitToHeight = 0
        sheet.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
        sheet.print_options.horizontalCentered = True

    wb.save(out_path)
    return out_path
