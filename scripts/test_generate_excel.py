# -*- coding: utf-8 -*-
"""generate_excel.py 场景类型列回归测试（无 pytest 依赖，python 直接运行）"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import generate_excel as ge

SAMPLE_CASES = [
    {
        "case_id": "TC-ORD-001", "scenario": "订单管理", "function": "创建订单",
        "name": "创建订单-正常流程", "scenario_type": "正常流程",
        "verify_point": "成功主路径", "precondition": "登录态有效，库存充足",
        "steps": "POST /order，body 传 {userId:1,amount:100}",
        "verify_method": "接口响应", "auto_level": "可单测", "priority": "P0",
        "expected_result": "HTTP 200，返回订单，库存-1", "test_result": "",
        "remark": "", "code_location": "OrderService.java:27 (createOrder)",
    },
    {
        "case_id": "TC-ORD-002", "scenario": "订单管理", "function": "创建订单",
        "name": "创建订单-用户为空校验", "scenario_type": "异常流程",
        "verify_point": "必填校验", "precondition": "创建订单接口可用",
        "steps": "POST /order，body 传 {userId:null}",
        "verify_method": "接口响应", "auto_level": "可单测", "priority": "P0",
        "expected_result": "返回错误\"用户不能为空\"", "test_result": "",
        "remark": "", "code_location": "OrderService.java:28 (createOrder)",
    },
]


def test_columns_have_scenario_type_at_position_5():
    assert len(ge.COLUMNS) == 15, f"期望 15 列，实际 {len(ge.COLUMNS)}"
    assert ge.COLUMNS[4] == ("scenario_type", "场景类型"), ge.COLUMNS[4]


def test_main_sheet_header_and_values():
    wb = ge.build_workbook(SAMPLE_CASES)
    ws = wb["研发自测用例"]
    header = [c.value for c in ws[1]]
    assert len(header) == 15, header
    assert header[4] == "场景类型", header
    assert ws.cell(row=2, column=5).value == "正常流程"
    assert ws.cell(row=3, column=5).value == "异常流程"
    dv_formulas = [dv.formula1 or "" for dv in ws.data_validations.dataValidation]
    assert any("正常流程" in f and "异常流程" in f for f in dv_formulas), dv_formulas


def _normalize_color(font_color):
    """openpyxl 读回的字体色可能为 'RRGGBB'/'00RRGGBB'/'FFRRGGBB'，统一取后 6 位大写比较；
    theme/indexed 等无 rgb 的颜色（默认黑）返回 None。"""
    if font_color is None:
        return None
    try:
        rgb = font_color.rgb
    except Exception:
        return None
    if not rgb:
        return None
    return str(rgb)[-6:].upper()


def test_scenario_type_row_font_colors():
    """v1.21+v1.22: 场景类型整行着色（正常绿/异常红）；优先级列 v1.22 起由优先级色覆盖。"""
    wb = ge.build_workbook(SAMPLE_CASES)
    ws = wb["研发自测用例"]
    ncols = len(ge.COLUMNS)
    pri_col = ge.COLUMNS.index(("priority", "优先级")) + 1
    for col in range(1, ncols + 1):
        if col == pri_col:
            continue  # 优先级列显示优先级色（见下方断言）
        color = _normalize_color(ws.cell(row=2, column=col).font.color)
        assert color == "008000", f"正常流程行第{col}列应为绿色(008000)，实际 {color}"
        color = _normalize_color(ws.cell(row=3, column=col).font.color)
        assert color == "FF0000", f"异常流程行第{col}列应为红色(FF0000)，实际 {color}"
    # 场景类型列本身颜色一致
    assert _normalize_color(ws.cell(row=2, column=5).font.color) == "008000"
    assert _normalize_color(ws.cell(row=3, column=5).font.color) == "FF0000"
    # 优先级列显示优先级色（样例两行都是 P0 → 红）
    assert _normalize_color(ws.cell(row=2, column=pri_col).font.color) == "C00000"
    assert _normalize_color(ws.cell(row=3, column=pri_col).font.color) == "C00000"


def test_data_row_heights_set():
    """v1.22: 表头/数据行行高已按内容自适应设置（保证换行内容完全显示）。"""
    wb = ge.build_workbook(SAMPLE_CASES)
    ws = wb["研发自测用例"]
    assert ws.row_dimensions[1].height == 28, "表头行高应为 28"
    for r in range(2, 2 + len(SAMPLE_CASES)):
        h = ws.row_dimensions[r].height
        assert h is not None and h >= 20, f"第{r}行行高未设置或过小：{h}"


def test_banding_fill():
    """v1.22: 数据行斑马纹——偶数行浅蓝、奇数行白，提升逐行可读性。"""
    wb = ge.build_workbook(SAMPLE_CASES)
    ws = wb["研发自测用例"]
    f2 = str(ws.cell(row=2, column=1).fill.fgColor.rgb)[-6:].upper()
    f3 = str(ws.cell(row=3, column=1).fill.fgColor.rgb)[-6:].upper()
    assert f2 == "F2F7FB", f"偶数行应为浅蓝 F2F7FB，实际 {f2}"
    assert f3 == "FFFFFF", f"奇数行应为白 FFFFFF，实际 {f3}"


def test_priority_text_colors():
    """v1.22: 优先级文字着色——P0 红、P1 琥珀、P2 灰，突出高风险。"""
    p_cases = [
        {"case_id": "T1", "scenario_type": "正常流程", "priority": "P0"},
        {"case_id": "T2", "scenario_type": "正常流程", "priority": "P1"},
        {"case_id": "T3", "scenario_type": "正常流程", "priority": "P2"},
    ]
    wb = ge.build_workbook(p_cases)
    ws = wb["研发自测用例"]
    pri_col = ge.COLUMNS.index(("priority", "优先级")) + 1
    for i, exp in zip(range(2, 5), ("C00000", "BF8F00", "7F7F7F")):
        color = _normalize_color(ws.cell(row=i, column=pri_col).font.color)
        assert color == exp, f"优先级 {ws.cell(row=i, column=pri_col).value} 应为 {exp}，实际 {color}"


def test_coverage_sheet_row_heights():
    """v1.22: 覆盖说明 sheet 各行行高已按内容自适应设置。"""
    coverage = {
        "stats": {"entry_covered": 1, "entry_total": 1, "dp_covered": 2, "dp_total": 2},
        "inventory": [{"function": "创建订单", "decision_point": "if (amount <= 0) throw",
                       "line": "OrderService.java:3",
                       "branches": ["真→TC-ORD-001 校验失败拒绝", "假→TC-ORD-002 继续执行"]}],
        "uncovered": [], "notes": ["orderDao 为外部依赖，行为无法静态验证"],
        "verification_gaps": [], "rule_gaps": [],
    }
    wb = ge.build_workbook(SAMPLE_CASES)
    ge.build_coverage_sheet(wb, coverage, SAMPLE_CASES)
    ws = wb["覆盖说明"]
    heights = [rd.height for rd in ws.row_dimensions.values() if rd.height is not None]
    assert heights, "覆盖说明 sheet 未设置任何行高"
    assert all(h >= 20 for h in heights), heights


def test_coverage_sheet_scenario_type_distribution():
    coverage = {
        "stats": {"entry_covered": 1, "entry_total": 1, "dp_covered": 2, "dp_total": 2},
        "inventory": [], "uncovered": [], "notes": [],
        "verification_gaps": [], "rule_gaps": [],
    }
    wb = ge.build_workbook(SAMPLE_CASES)
    ge.build_coverage_sheet(wb, coverage, SAMPLE_CASES)
    ws = wb["覆盖说明"]
    hdr_row = None
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        for cell in row:
            if cell.value == "场景类型分布（按功能统计正常/异常流程用例数）":
                hdr_row = cell.row
                break
        if hdr_row:
            break
    assert hdr_row is not None, "覆盖说明未找到场景类型分布表头"
    assert ws.cell(row=hdr_row + 1, column=1).value == "功能"
    assert ws.cell(row=hdr_row + 1, column=2).value == "正常流程"
    assert ws.cell(row=hdr_row + 1, column=3).value == "异常流程"
    assert ws.cell(row=hdr_row + 2, column=1).value == "创建订单"
    assert ws.cell(row=hdr_row + 2, column=2).value == 1  # 正常流程计数
    assert ws.cell(row=hdr_row + 2, column=3).value == 1  # 异常流程计数


if __name__ == "__main__":
    failures = 0
    for name, fn in list(globals().items()):
        if name.startswith("test_") and callable(fn):
            try:
                fn()
                print(f"PASS {name}")
            except AssertionError as e:
                failures += 1
                print(f"FAIL {name}: {e}")
    if failures:
        print(f"\n{failures} 项失败")
        sys.exit(1)
    print("\n全部通过")
