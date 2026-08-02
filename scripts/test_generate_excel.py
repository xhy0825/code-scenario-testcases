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
