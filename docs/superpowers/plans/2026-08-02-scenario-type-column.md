# v1.19 新增「场景类型」列 实施计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 让 skill 明确体现「新增用例场景（正常场景+异常场景）并显示在 Excel 中」——主 sheet 新增「场景类型」列（正常流程/异常流程），覆盖说明新增「场景类型分布」。

**Architecture:** 三处改动——① `scripts/generate_excel.py`：`COLUMNS` 在 `name` 后插入 `("scenario_type","场景类型")` 使 Excel 主表 14→15 列，加下拉与覆盖说明分布统计；② `SKILL.md`：15 列字段表加「场景类型」行（含判定规则）、自检 9→10 条、输出示例与措辞同步；③ `README.md`：§1.1/§1.2/§2.1/§2.3/§2.4/§2.5/更新日志同步。

**Tech Stack:** Python 3 + openpyxl；Markdown。无 pytest 依赖（测试脚本可用 `python` 直接运行）。

**Spec:** `docs/superpowers/specs/2026-08-02-scenario-type-column-design.md`

## Global Constraints

- 只改 `scripts/generate_excel.py`、`SKILL.md`、`README.md`；`references/*.md`、`evals/` 一律不动
- 「场景类型」列位于主 sheet **第 5 列**（紧邻「用例名称」之后），全表 **15 列**
- 值枚举：`正常流程` / `异常流程`；判定按**路径最终现象**分类（正常=成功完成功能；异常=被拒绝/失败），**看现象不看分支方向**
- 每个功能**至少 1 条正常流程用例**（成功主路径）
- `cases.json` 字段名：`scenario_type`（与 `scripts/generate_excel.py` 的 `COLUMNS` 键一致）
- 覆盖说明 sheet 新增「场景类型分布」：按功能统计 正常流程/异常流程 用例数
- 版本 v1.19；commit 信息格式 `feat:`/`docs:`；每次 commit 触发 post-commit 自动部署属预期行为

---

### Task 1: generate_excel.py 新增场景类型列 + 场景类型分布

**Files:**
- Modify: `scripts/generate_excel.py`（`COLUMNS`、`COLUMN_WIDTHS`、`build_workbook` 下拉、`build_coverage_sheet` 分布）
- Create: `scripts/test_generate_excel.py`（无 pytest 依赖的回归测试）

**Interfaces:**
- Produces: `ge.COLUMNS` 长度 15、`COLUMNS[4] == ("scenario_type","场景类型")`；`build_workbook(cases)` 主表第 5 列「场景类型」含 正常流程/异常流程 下拉；`build_coverage_sheet(wb, coverage, cases)` 新增「场景类型分布（按功能统计正常/异常流程用例数）」区块

- [ ] **Step 1: 写失败测试**

创建 `scripts/test_generate_excel.py`：

```python
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
```

- [ ] **Step 2: 运行测试，确认失败**

Run: `cd /d/AgentDev/code-scenario-testcases && python scripts/test_generate_excel.py`
Expected: FAIL——`len(ge.COLUMNS) == 15` 断言失败（当前 COLUMNS 14 列）

- [ ] **Step 3: 实现 COLUMNS 与 COLUMN_WIDTHS**

在 `scripts/generate_excel.py` 的 `COLUMNS`（第 46 行起）中，`("name", "用例名称")` 之后插入：

```python
    ("scenario_type", "场景类型"),
```

在 `COLUMN_WIDTHS`（第 64 行起）中加：

```python
    "scenario_type": 12,
```

- [ ] **Step 4: 实现主表场景类型下拉**

在 `build_workbook` 中「测试结果列」DataValidation 之后（约第 281 行 `dv.add(...)` 之后）追加：

```python
    # 场景类型列：正常流程/异常流程 下拉（数据验证）
    st_col = COLUMNS.index(("scenario_type", "场景类型")) + 1
    st_letter = get_column_letter(st_col)
    st_dv = DataValidation(
        type="list",
        formula1='"正常流程,异常流程"',
        allow_blank=True,
        showDropDown=False,
    )
    st_dv.error = "只能选择 正常流程 或 异常流程"
    st_dv.errorTitle = "无效输入"
    ws.add_data_validation(st_dv)
    st_dv.add(f"{st_letter}2:{st_letter}{max(ws.max_row, 200)}")
```

- [ ] **Step 5: 实现覆盖说明「场景类型分布」**

在 `build_coverage_sheet` 的「可自动化程度分布」`if auto_dist:` 块之后、「功能覆盖清单」`per_function` 块之前插入：

```python
    # 场景类型分布（按功能统计正常/异常流程用例数）
    if cases:
        by_type = {}
        for c in cases:
            fn = c.get("function") or "?"
            by_type.setdefault(fn, {"正常流程": 0, "异常流程": 0})
            st = c.get("scenario_type")
            if st in by_type[fn]:
                by_type[fn][st] += 1
        if by_type:
            row += 1
            ws.cell(row=row, column=1, value="场景类型分布（按功能统计正常/异常流程用例数）").font = bold
            row += 1
            for col, h in enumerate(["功能", "正常流程", "异常流程"], start=1):
                c = ws.cell(row=row, column=col, value=h)
                c.font = bold
                c.border = border
            row += 1
            for fn in sorted(by_type):
                ws.cell(row=row, column=1, value=fn).border = border
                ws.cell(row=row, column=2, value=by_type[fn]["正常流程"]).border = border
                ws.cell(row=row, column=3, value=by_type[fn]["异常流程"]).border = border
                row += 1
```

- [ ] **Step 6: 运行测试，确认通过**

Run: `cd /d/AgentDev/code-scenario-testcases && python scripts/test_generate_excel.py`
Expected: 3 项 PASS，输出「全部通过」

- [ ] **Step 7: 冒烟：CLI 生成真实 xlsx 并抽查**

Run:
```bash
cd /d/AgentDev/code-scenario-testcases && python - <<'PY'
import json, subprocess, sys
sample = {
  "cases": [
    {"case_id":"TC-ORD-001","scenario":"订单管理","function":"创建订单","name":"创建订单-正常流程","scenario_type":"正常流程","verify_point":"成功主路径","precondition":"登录态有效，库存充足","steps":"POST /order","verify_method":"接口响应","auto_level":"可单测","priority":"P0","expected_result":"HTTP 200，返回订单","test_result":"","remark":"","code_location":"OrderService.java:27"},
    {"case_id":"TC-ORD-002","scenario":"订单管理","function":"创建订单","name":"创建订单-用户为空校验","scenario_type":"异常流程","verify_point":"必填校验","precondition":"接口可用","steps":"POST /order","verify_method":"接口响应","auto_level":"可单测","priority":"P0","expected_result":"返回错误\"用户不能为空\"","test_result":"","remark":"","code_location":"OrderService.java:28"}
  ],
  "coverage": {"stats":{"entry_covered":1,"entry_total":1,"dp_covered":2,"dp_total":2},"inventory":[],"uncovered":[],"notes":[],"verification_gaps":[],"rule_gaps":[]}
}
import tempfile, os
p = os.path.join(tempfile.gettempdir(), "v119_smoke.json")
json.dump(sample, open(p, "w", encoding="utf-8"), ensure_ascii=False)
out = os.path.join(tempfile.gettempdir(), "v119_smoke.xlsx")
subprocess.run([sys.executable, "scripts/generate_excel.py", p, out], check=True)
from openpyxl import load_workbook
wb = load_workbook(out)
ws = wb["研发自测用例"]
hdr = [c.value for c in ws[1]]
assert len(hdr) == 15 and hdr[4] == "场景类型", hdr
assert "覆盖说明" in wb.sheetnames
cov = wb["覆盖说明"]
vals = [[c.value for c in row] for row in cov.iter_rows(values_only=False)]
txt = " ".join(str(c.value) for row in vals for c in row if c.value)
assert "场景类型分布" in txt, txt
print("SMOKE OK: 15 列, 第5列=场景类型, 覆盖说明含场景类型分布")
PY
```
Expected: 输出 `SMOKE OK: 15 列, 第5列=场景类型, 覆盖说明含场景类型分布`

- [ ] **Step 8: Commit**

```bash
cd /d/AgentDev/code-scenario-testcases && git add scripts/generate_excel.py scripts/test_generate_excel.py && git commit -m "feat: v1.19 主表新增场景类型列(14→15) + 覆盖说明场景类型分布"
```

---

### Task 2: SKILL.md 15 列 + 场景类型判定规则

**Files:**
- Modify: `SKILL.md`（第 51、100-103、105-119、121-130、171、210-217 行区域）

**Interfaces:**
- Consumes: Task 1 的 `COLUMNS` 键名 `scenario_type`、列序（第 5 列）
- Produces: skill 运行时会为每条用例填写 `scenario_type`，主表 15 列

- [ ] **Step 1: 流程总览与措辞 14→15 列**

把 `SKILL.md` 第 51 行：
`| 阶段1 验证设计 | 3 设计验证用例 | 逐验证点可执行化（前置+步骤+预期），套 13 类展开模板 | 14 列用例 |`
改为：
`| 阶段1 验证设计 | 3 设计验证用例 | 逐验证点可执行化（前置+步骤+预期），套 13 类展开模板 | 15 列用例 |`

把第 100、101 行的「见 14 列字段表」改为「见 15 列字段表」。

把第 103 行：
`每条用例填 14 列（\`cases.json\` 键名与 \`scripts/generate_excel.py\` 的 \`COLUMNS\` 一致：\`case_id / scenario / function / name / verify_point / ...\`）：`
改为：
`每条用例填 15 列（\`cases.json\` 键名与 \`scripts/generate_excel.py\` 的 \`COLUMNS\` 一致：\`case_id / scenario / function / name / scenario_type / verify_point / precondition / steps / verify_method / auto_level / priority / expected_result / test_result / remark / code_location\`）：`

- [ ] **Step 2: 14 列字段表加「场景类型」行**

在 14 列字段表中「用例名称」行之后插入：

```markdown
| 场景类型 | **正常流程** / **异常流程**：按该路径**最终现象**分类——正常流程=成功完成功能（校验通过/成功返回/状态变更成功/正常展示）；异常流程=被拒绝或失败（抛异常/错误码/状态拒绝/外部依赖失败）。**看现象不看分支方向**（`if (stockOk){保存}` 真分支=正常；`if (coupon==null) throw` 真分支=异常）。复合条件"成功1条"=正常、"原子不满足"=异常；switch 各 case=正常、default=异常；外部契约成功态=正常、异常/超时/降级=异常。**每个功能至少 1 条正常流程用例（成功主路径）** |
```

- [ ] **Step 3: 字段质量自检 9→10 条**

在字段质量自检第 9 条之后追加：

```markdown
10. **场景类型标注正确**——正常流程用例预期为成功完成（返回成功/状态变更成功），异常流程用例预期为拒绝/失败（抛异常/错误码），两者不得互换
```

- [ ] **Step 4: 第 7 步主 sheet 描述 14→15 列**

把第 171 行：
`- 主 sheet 14 列（含 PASS/FAIL 下拉）+「覆盖说明」sheet（统计/追溯矩阵/验证点覆盖/可自动化分布/缺口/标注/规则缺口）`
改为：
`- 主 sheet 15 列（含 场景类型/测试结果 下拉）+「覆盖说明」sheet（统计/追溯矩阵/验证点覆盖/可自动化分布/场景类型分布/缺口/标注/规则缺口）`

- [ ] **Step 5: 输出示例表加「场景类型」列**

把第 210-217 行示例表改为（表头加一列，并给每行补场景类型值）：

```markdown
| 用例编号 | 功能 | 用例名称 | 场景类型 | 验证点 | 前置条件 | 操作步骤 | 验证方法 | 可自动化程度 | 优先级 | 预期结果 | 代码位置 |
|---|---|---|---|---|---|---|---|---|---|---|---|
| TC-ORD-001 | 创建订单 | 用户为空校验 | 异常流程 | 必填：userId 非空（业务规则） | 创建订单接口可用，无历史订单依赖 | POST /order，body 传 {userId:null} | 接口响应 | 可单测 | P0 | 返回错误"用户不能为空"，不生成订单 | OrderService.java:3 |
| TC-ORD-002 | 创建订单 | 库存不足校验 | 异常流程 | 库存规则（外部契约） | userId 有效；mock 库存校验返回 false | POST /order，body 传 {stock:false} | 接口响应 | 可单测 | P0 | 返回错误"库存不足" | OrderService.java:5 |
| TC-ORD-003 | 创建订单 | 微信支付 | 正常流程 | 支付方式选择（判定点） | userId 有效，库存充足 | POST /order，传 payType=1 | 数据状态 | 可单测 | P1 | 订单 payChannel 置为 wechat（查库确认） | OrderService.java:8 |
| TC-ORD-004 | 创建订单 | 支付方式不支持 | 异常流程 | 兜底分支（判定点） | userId 有效 | POST /order，传 payType=3 | 接口响应 | 可单测 | P2 | 返回错误"支付方式不支持" | OrderService.java:11 |
```

- [ ] **Step 6: 验证 SKILL.md 改动一致**

Run: `cd /d/AgentDev/code-scenario-testcases && grep -n "14 列" SKILL.md`
Expected: 无输出（`14 列` 已全部改为 15 列；「15 列」出现 ≥4 次）

Run: `grep -c "场景类型" SKILL.md`
Expected: ≥5（字段表行 + 自检 10 + 示例表 4 行 + 第 7 步描述）

- [ ] **Step 7: Commit**

```bash
cd /d/AgentDev/code-scenario-testcases && git add SKILL.md && git commit -m "feat: v1.19 SKILL.md 15 列 + 场景类型判定规则 + 自检 10 条"
```

---

### Task 3: README.md 同步

**Files:**
- Modify: `README.md`（§1.1 能力表、§1.2 主表列说明+覆盖说明列表、§2.1、§2.3、§2.4、§2.5、更新日志）

**Interfaces:**
- Consumes: Task 1/2 的 15 列与 `scenario_type` 字段名
- Produces: README 与实现一致的 15 列描述 + v1.19 更新日志

- [ ] **Step 1: §1.1 能力表加「场景类型标注」行**

在 §1.1 能力表「验证点设计」行（第 18 行）之后追加：

```markdown
| **场景类型标注** | 每条用例标注 **正常流程 / 异常流程**（按路径最终现象分类，看现象不看分支方向），每个功能保证至少 1 条正常流程用例，Excel 可按场景类型筛选 |
```

- [ ] **Step 2: §1.2 主 sheet 列说明表加「场景类型」行**

在「用例名称」行之后追加：

```markdown
| 场景类型 | 正常流程（成功完成功能）/ 异常流程（被拒绝/失败），按路径最终现象分类；每个功能至少 1 条正常流程 |
```

- [ ] **Step 3: §1.2 覆盖说明列表加「场景类型分布」**

在「可自动化程度分布」行之后追加：

```markdown
- **场景类型分布**：按功能统计 正常流程/异常流程 用例数（每个场景正常+异常覆盖可核算）
```

- [ ] **Step 4: §2.1 目录结构 14→15 列**

把第 168 行 `│   │   ├── 主 sheet（14 列 + 样式 + 筛选 + 冻结首行）` 改为：
`│   │   ├── 主 sheet（15 列 + 样式 + 筛选 + 冻结首行）`

- [ ] **Step 5: §2.3 cases.json 示例加 scenario_type**

在 §2.3 用例对象示例的 `"name": "创建订单-用户为空校验",` 之后追加：

```jsonc
      "scenario_type": "异常流程",                    // 正常流程 / 异常流程
```

- [ ] **Step 6: §2.4 关键设计决策加一行**

在 §2.4 表格末尾追加：

```markdown
| **场景类型显式标注** | 成功路径与拒绝/失败路径用「场景类型」列显式区分，Excel 可直接筛选"正常流程"，确保每个功能正常主路径不被漏测 |
```

- [ ] **Step 7: §2.5 第 3 步 14→15 列**

把第 264 行 `...产出 14 列用例 + 字段质量自检` 改为 `...产出 15 列用例 + 字段质量自检`。

- [ ] **Step 8: 更新日志加 v1.19**

在 v1.18 行之后追加：

```markdown
| **v1.19** | 场景类型标注 | 主 sheet 新增「场景类型」列（正常流程/异常流程，按路径最终现象分类），14→15 列；每个功能保证至少 1 条正常流程用例；覆盖说明新增「场景类型分布」（按功能统计正常/异常用例数）；字段质量自检 9→10 条 |
```

- [ ] **Step 9: 验证 README 改动一致**

Run: `cd /d/AgentDev/code-scenario-testcases && grep -n "15 列" README.md`
Expected: 命中 §2.1、§2.5 两处（v1.19 更新日志行含"14→15 列"亦属预期）

Run: `grep -c "场景类型" README.md`
Expected: ≥6（能力表/主表列说明/覆盖说明分布/更新日志/§2.3/§2.4）

- [ ] **Step 10: Commit**

```bash
cd /d/AgentDev/code-scenario-testcases && git add README.md && git commit -m "docs: v1.19 README 同步 15 列 + 场景类型标注 + 更新日志"
```

---

### Task 4: 终验（回归 + 部署一致性）

**Files:**
- Verify: `scripts/generate_excel.py`、`scripts/test_generate_excel.py`、`SKILL.md`、`README.md`、运行时 `C:\Users\46018\.claude\skills\code-scenario-testcases\SKILL.md`

**Interfaces:**
- Consumes: Task 1-3 全部改动

- [ ] **Step 1: 全量回归**

Run: `cd /d/AgentDev/code-scenario-testcases && python scripts/test_generate_excel.py`
Expected: 3 项 PASS，输出「全部通过」

- [ ] **Step 2: SKILL/README 列数一致性**

Run: `cd /d/AgentDev/code-scenario-testcases && grep -n "14 列" SKILL.md README.md`
Expected: 仅命中更新日志历史条目（v1.0「11 列用例表」、v1.16「11→14 列」、v1.17「13→14」、v1.18「14 列表权威定义」）——均为历史记载，保留

Run: `grep -n "scenario_type" SKILL.md README.md scripts/generate_excel.py`
Expected: 三处均含 `scenario_type`（SKILL 键名清单、README §2.3 示例、generate_excel.py COLUMNS）

- [ ] **Step 3: 确认运行时已部署**

Run: `diff <(sed -n '1,999p' /d/AgentDev/code-scenario-testcases/SKILL.md) /c/Users/46018/.claude/skills/code-scenario-testcases/SKILL.md && echo SAME`
Expected: `SAME`（若 Task 2/3 的 commit 均已触发 install.sh）

- [ ] **Step 4: 若 Step 3 不一致，手动部署**

Run: `cd /d/AgentDev/code-scenario-testcases && ./install.sh`
Expected: 输出 `install: skill 已部署到 /c/Users/46018/.claude/skills/code-scenario-testcases`，重跑 Step 3 确认 SAME

- [ ] **Step 5: （可选人工）用真实 skill 跑 order-java 抽查**

在 Claude Code 中调用 `code-scenario-testcases -all`（目录 `D:/AgentDev/fixtures/order-java`），抽查生成 Excel：第 5 列为「场景类型」、每个功能 ≥1 条正常流程、覆盖说明含「场景类型分布」。此步为人工冒烟，不阻塞验收。

---

## 自审结论

- **Spec 覆盖**：主表 15 列场景类型列 → Task 1 Step 3-4 + Task 2；判定规则（按现象/每功能≥1正常）→ Task 2 Step 2；覆盖说明场景类型分布 → Task 1 Step 5；自检 10 条 → Task 2 Step 3；README 同步 → Task 3；验证标准 → Task 1 Step 6-7 + Task 4。
- **占位符扫描**：无 TBD/TODO；每步含完整可复制代码。
- **类型一致性**：`scenario_type` 在 COLUMNS（Task 1）、SKILL 键名清单（Task 2）、README §2.3（Task 3）三处键名一致；列序均为第 5 列。
