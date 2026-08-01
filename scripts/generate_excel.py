#!/usr/bin/env python3
"""从结构化 JSON 用例数据生成研发自测用例 Excel（.xlsx）。

输入 JSON 的形态：
  1. {"cases": [ {用例对象}, ... ], "coverage": {覆盖核算对象(可选)} }
  2. [ {用例对象}, ... ]

用例对象字段（与 SKILL.md 的列一一对应）：
  case_id         用例编号，如 TC-ORD-001
  scenario        业务场景，如 订单管理
  function        功能，如 创建订单
  name            用例名称，如 创建订单-用户为空校验
  precondition    前置条件
  steps           操作步骤（list 或单行字符串）
  priority        优先级，P0/P1/P2
  expected_result 预期结果（不通过分支写功能展示的结果）
  test_result     测试结果（留空，列内下拉 PASS/FAIL）
  remark          备注
  code_location   代码位置，如 OrderService.java:3

coverage 对象（SKILL.md 第 4 步覆盖率核对产出，用于生成「覆盖说明」sheet）：
  stats        {"entry_covered":3, "entry_total":3, "dp_covered":22, "dp_total":22}
  inventory    [ {"function":"创建订单","decision_point":"if (...)","line":"OrderService.java:3","branches":["真→TC-ORD-001","假→TC-ORD-002"]}, ... ]
  uncovered    [ 缺口描述, ... ]（无缺口可为空数组）
  notes        [ 无法静态验证项, ... ]

用法：
  python generate_excel.py cases.json [输出路径.xlsx]
"""

import argparse
import json
import re
import sys
from datetime import date
from pathlib import Path

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation
except ImportError:
    sys.exit("缺少 openpyxl，请先执行: pip install openpyxl")

COLUMNS = [
    ("case_id", "用例编号"),
    ("scenario", "业务场景"),
    ("function", "功能"),
    ("name", "用例名称"),
    ("precondition", "前置条件"),
    ("steps", "操作步骤"),
    ("priority", "优先级"),
    ("expected_result", "预期结果"),
    ("test_result", "测试结果"),
    ("remark", "备注"),
    ("code_location", "代码位置"),
]

# 每列宽度（近似，按最长表头/内容预估）
COLUMN_WIDTHS = {
    "case_id": 14, "scenario": 16, "function": 16, "name": 28,
    "precondition": 34, "steps": 42, "priority": 8,
    "expected_result": 44, "test_result": 12, "remark": 30, "code_location": 18,
}


def load_data(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def build_coverage_sheet(wb, coverage):
    """在同一工作簿新增「覆盖说明」sheet：覆盖统计（含变更范围）+ 追溯矩阵 + 标注区（缺口/无法验证/规则缺口）。"""
    ws = wb.create_sheet("覆盖说明")
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    wrap = Alignment(wrap_text=True, vertical="top")
    bold = Font(bold=True)

    row = 1
    ws.cell(row=row, column=1, value="覆盖说明（代码可表达范围内）").font = Font(bold=True, size=12)
    row += 2

    stats = coverage.get("stats", {}) if isinstance(coverage, dict) else {}
    ws.cell(row=row, column=1, value="覆盖统计").font = bold
    row += 1
    for label, covered, total in (
        ("业务入口覆盖", stats.get("entry_covered"), stats.get("entry_total")),
        ("判定点覆盖", stats.get("dp_covered"), stats.get("dp_total")),
    ):
        ws.cell(row=row, column=1, value=label).border = border
        val = f"{covered}/{total}" if covered is not None and total is not None else ""
        ws.cell(row=row, column=2, value=val).border = border
        row += 1

    # 功能覆盖清单（研发视角：每个功能必须有 ≥1 条用例）
    per_function = (coverage or {}).get("per_function") if isinstance(coverage, dict) else None
    if per_function:
        row += 1
        ws.cell(row=row, column=1, value="功能覆盖清单（每个功能必须有 ≥1 条用例）").font = bold
        row += 1
        for col, h in enumerate(["功能", "判定点覆盖"], start=1):
            c = ws.cell(row=row, column=col, value=h)
            c.font = bold
            c.border = border
        row += 1
        for fn, st in per_function.items():
            ws.cell(row=row, column=1, value=fn).border = border
            ws.cell(row=row, column=2,
                    value=f"{st.get('dp_covered', 0)}/{st.get('dp_total', 0)}").border = border
            row += 1

    row += 1
    ws.cell(row=row, column=1, value="变更范围").font = bold
    row += 1
    change_scope = (coverage or {}).get("change_scope") if isinstance(coverage, dict) else None
    if change_scope:
        for key, label in (
            ("mode", "模式"),
            ("file_count", "变更文件数"),
            ("function_count", "受影响函数数"),
            ("affected_decision_points_total", "受影响函数内判定点总数"),
        ):
            if change_scope.get(key) is not None:
                ws.cell(row=row, column=1, value=label).border = border
                ws.cell(row=row, column=2, value=change_scope.get(key)).border = border
                row += 1
        for key, label in (
            ("changed_files", "变更文件"),
            ("affected_functions", "受影响函数"),
            ("new_decision_points", "新增判定点"),
        ):
            if change_scope.get(key):
                ws.cell(row=row, column=1, value=label).border = border
                c = ws.cell(row=row, column=2, value="\n".join(change_scope[key]))
                c.alignment = wrap
                c.border = border
                row += 1

    row += 1
    ws.cell(row=row, column=1, value="判定点 → 用例 追溯矩阵").font = bold
    row += 1
    for col, h in enumerate(["功能", "判定点", "代码位置", "分支 → 用例编号"], start=1):
        c = ws.cell(row=row, column=col, value=h)
        c.font = bold
        c.border = border
    row += 1
    for item in (coverage or {}).get("inventory") or []:
        branches = item.get("branches") or []
        if isinstance(branches, list):
            branches = "\n".join(str(b) for b in branches)
        for col, v in enumerate(
            [item.get("function", ""), item.get("decision_point", ""), item.get("line", ""), branches],
            start=1,
        ):
            c = ws.cell(row=row, column=col, value=v)
            c.alignment = wrap
            c.border = border
        row += 1

    # ---- 区 3：标注区（缺口 / 无法验证 / 规则缺口 合并） ----
    row += 1
    ws.cell(row=row, column=1, value="标注区（缺口 / 无法静态验证 / 规则缺口）").font = bold
    row += 1
    items = []
    for u in (coverage or {}).get("uncovered") or []:
        items.append(f"【缺口】{u}")
    for n in (coverage or {}).get("notes") or []:
        items.append(f"【无法验证】{n}")
    for g in (coverage or {}).get("rule_gaps") or []:
        if isinstance(g, dict):
            g = f"{g.get('category', '')}: {g.get('gap', '')}（{g.get('function', '')}，建议确认：{g.get('confirm_with', '')}）"
        items.append(f"【规则缺口】{g}")
    if not items:
        ws.cell(row=row, column=1, value="无（代码可表达范围内已全覆盖，无额外标注）").alignment = wrap
    else:
        for it in items:
            ws.cell(row=row, column=1, value=it).alignment = wrap
            row += 1

    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 32
    ws.column_dimensions["C"].width = 20
    ws.column_dimensions["D"].width = 44
    return ws


def to_cell_value(value):
    if isinstance(value, (list, tuple)):
        return "\n".join(str(v) for v in value)
    if value is None:
        return ""
    return value


def build_workbook(cases):
    wb = Workbook()
    ws = wb.active
    ws.title = "研发自测用例"

    header_fill = PatternFill("solid", fgColor="4472C4")
    header_font = Font(bold=True, color="FFFFFF")
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    wrap = Alignment(wrap_text=True, vertical="top")

    # 表头
    for col, (_, title) in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=1, column=col, value=title)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(vertical="center", horizontal="center")
        cell.border = border

    # 数据行
    for row, case in enumerate(cases, start=2):
        for col, (key, _) in enumerate(COLUMNS, start=1):
            value = to_cell_value(case.get(key))
            cell = ws.cell(row=row, column=col, value=value)
            cell.alignment = wrap
            cell.border = border

    # 列宽 + 冻结首行 + 筛选
    for col, (key, _) in enumerate(COLUMNS, start=1):
        ws.column_dimensions[get_column_letter(col)].width = COLUMN_WIDTHS.get(key, 20)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    # 测试结果列：PASS/FAIL 下拉（数据验证），允许留空，预留给后续新增行
    test_col = COLUMNS.index(("test_result", "测试结果")) + 1
    test_letter = get_column_letter(test_col)
    dv = DataValidation(
        type="list",
        formula1='"PASS,FAIL"',
        allow_blank=True,
        showDropDown=False,  # False = 显示下拉箭头（Excel 内部 XML 属性是反的）
    )
    dv.error = "只能选择 PASS 或 FAIL"
    dv.errorTitle = "无效输入"
    ws.add_data_validation(dv)
    dv.add(f"{test_letter}2:{test_letter}{max(ws.max_row, 200)}")

    return wb


def safe_name(name):
    """清理文件/目录名中的非法字符（Windows 不允许 \\/:*?\"<>|）。"""
    return re.sub(r'[\\/:*?"<>|]', "_", str(name)).strip() or "testcase"


def main():
    parser = argparse.ArgumentParser(description="Generate test-case Excel from JSON")
    parser.add_argument("input", help="Path to JSON file (array of cases or {\"cases\":[...]})")
    parser.add_argument("output", nargs="?", help="Output .xlsx path（指定后覆盖自动命名）")
    parser.add_argument("--project-name", default=None, help="项目名称，用于输出目录与文件名")
    parser.add_argument("--version", default="v1.0", help="版本号，默认 v1.0")
    parser.add_argument("--output-dir", default="testcase", help="输出根目录（默认 testcase），其下按项目名建子目录")
    args = parser.parse_args()

    data = load_data(args.input)
    if isinstance(data, dict) and "cases" in data:
        cases, coverage = data["cases"], data.get("coverage")
    else:
        cases, coverage = data, None

    if args.output:
        out = args.output
    else:
        project = safe_name(args.project_name or "研发自测用例")
        version = str(args.version)
        if not version.startswith("v"):
            version = f"v{version}"
        fname = f"{project}_{date.today():%Y%m%d}_{version}.xlsx"
        out_dir = Path(args.output_dir) / project
        out_dir.mkdir(parents=True, exist_ok=True)
        out = str(out_dir / fname)

    wb = build_workbook(cases)
    if coverage:
        build_coverage_sheet(wb, coverage)
    wb.save(out)
    extra = "（含「覆盖说明」sheet）" if coverage else ""
    print(f"已生成 {len(cases)} 条用例{extra} -> {out}")


if __name__ == "__main__":
    main()
