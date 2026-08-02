# v1.21 场景类型文字着色 实施计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:test-driven-development（先写失败测试再实现），配合 writing-skills 的 RED-GREEN-REFACTOR。Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 主 sheet 按场景类型给整行文字着色——正常流程绿色、异常流程红色，让成功/拒绝路径在 Excel 中一眼可辨。

**Architecture:** 三处改动——① `scripts/generate_excel.py`：新增 `SCENARIO_TYPE_COLORS` 常量（正常 `008000` / 异常 `FF0000`），`build_workbook` 数据行循环内按场景类型给整行设字体色；② `scripts/test_generate_excel.py`：新增着色回归测试（先写、先红后绿）；③ `SKILL.md` + `README.md`：场景类型字段规范补充着色说明 + 更新日志。

**Tech Stack:** Python 3 + openpyxl；Markdown。无 pytest 依赖（`python scripts/test_generate_excel.py` 直接运行）。

**Spec:** `docs/superpowers/specs/2026-08-02-scenario-type-color-design.md`

## Global Constraints

- 只改 `scripts/generate_excel.py`、`scripts/test_generate_excel.py`、`SKILL.md`、`README.md`；`references/*.md`、`evals/`、`enumerate_decision_points.py` 一律不动
- 着色对象 = 数据行**全部 15 列**的字体色（font.color），不做底色填充
- 颜色：正常流程 `#008000` 绿、异常流程 `#FF0000` 红；未知/空场景类型**不设色**（默认黑）
- 场景类型判定规则沿用 v1.19，不变更；既有列数/下拉/覆盖说明分布行为不回归
- 版本 v1.21；commit 信息格式 `feat:`/`docs:`；每次 commit 触发 post-commit 自动部署属预期行为

---

### Task 1: generate_excel.py 场景类型着色 + 回归测试（RED-GREEN）

**Files:**
- Modify: `scripts/generate_excel.py`（`SCENARIO_TYPE_COLORS` 常量 + `build_workbook` 数据行着色）
- Modify: `scripts/test_generate_excel.py`（新增 `test_scenario_type_row_font_colors` + `_normalize_color` 辅助）

**Interfaces:**
- Produces: `ge.SCENARIO_TYPE_COLORS == {"正常流程":"008000","异常流程":"FF0000"}`；`build_workbook(cases)` 中正常流程行所有列 `font.color` 归一化为 `008000`、异常流程行为 `FF0000`、未知/空行保持默认（无显式字体色）

- [ ] **Step 1: 写失败测试**

在 `scripts/test_generate_excel.py` 加辅助函数与测试：

```python
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
    """v1.21: 场景类型文字着色——正常流程整行绿色，异常流程整行红色。"""
    wb = ge.build_workbook(SAMPLE_CASES)
    ws = wb["研发自测用例"]
    ncols = len(ge.COLUMNS)
    for col in range(1, ncols + 1):
        color = _normalize_color(ws.cell(row=2, column=col).font.color)
        assert color == "008000", f"正常流程行第{col}列应为绿色(008000)，实际 {color}"
    for col in range(1, ncols + 1):
        color = _normalize_color(ws.cell(row=3, column=col).font.color)
        assert color == "FF0000", f"异常流程行第{col}列应为红色(FF0000)，实际 {color}"
    assert _normalize_color(ws.cell(row=2, column=5).font.color) == "008000"
    assert _normalize_color(ws.cell(row=3, column=5).font.color) == "FF0000"
```

- [ ] **Step 2: 运行测试，确认失败**

Run: `cd /d/AgentDev/code-scenario-testcases && python scripts/test_generate_excel.py`
Expected: `test_scenario_type_row_font_colors` FAIL（当前未着色，字体色为默认 theme 色）

- [ ] **Step 3: 实现着色**

在 `scripts/generate_excel.py` 的 `COLUMN_WIDTHS` 之前加：

```python
# 场景类型文字着色（v1.21）：正常流程整行绿色、异常流程整行红色，便于目视区分成功/拒绝路径
SCENARIO_TYPE_COLORS = {
    "正常流程": "008000",  # 绿
    "异常流程": "FF0000",  # 红
}
```

把 `build_workbook` 数据行循环改为：

```python
    # 数据行（按场景类型给整行字体着色：正常流程绿 / 异常流程红，未知或空保持默认黑）
    for row, case in enumerate(cases, start=2):
        font_color = SCENARIO_TYPE_COLORS.get(case.get("scenario_type"))
        for col, (key, _) in enumerate(COLUMNS, start=1):
            value = to_cell_value(case.get(key))
            cell = ws.cell(row=row, column=col, value=value)
            cell.alignment = wrap
            cell.border = border
            if font_color:
                cell.font = Font(color=font_color)
```

- [ ] **Step 4: 运行测试，确认通过**

Run: `cd /d/AgentDev/code-scenario-testcases && python scripts/test_generate_excel.py`
Expected: 4 项 PASS，输出「全部通过」

- [ ] **Step 5: 冒烟：三种场景类型抽查**

Run:
```bash
cd /d/AgentDev/code-scenario-testcases && python - <<'PY'
import generate_excel as ge
cases = [
    {"case_id":"TC-1","scenario_type":"正常流程","name":"下单-正常"},
    {"case_id":"TC-2","scenario_type":"异常流程","name":"下单-库存不足"},
    {"case_id":"TC-3","scenario_type":"","name":"未标注"},
]
wb = ge.build_workbook(cases)
ws = wb["研发自测用例"]
for r in range(2, 5):
    c = ws.cell(row=r, column=1).font.color
    try:
        rgb = c.rgb
    except Exception:
        rgb = "default(black)"
    print(f"row{r} scenario_type={ws.cell(row=r, column=5).value} color={rgb}")
PY
```
Expected: row2=正常流程 绿色（`00` 前缀 ARGB 属 openpyxl 正常表现）、row3=异常流程 红色、row4=未标注 默认黑

- [ ] **Step 6: Commit**

```bash
cd /d/AgentDev/code-scenario-testcases && git add scripts/generate_excel.py scripts/test_generate_excel.py && git commit -m "feat: v1.21 主表场景类型整行文字着色(正常绿/异常红) + 回归测试"
```

---

### Task 2: SKILL.md + README.md 同步着色说明

**Files:**
- Modify: `SKILL.md`（15 列字段表「场景类型」行）
- Modify: `README.md`（§1.1 能力表「场景类型标注」行、§1.2 主 sheet 列说明「场景类型」行、更新日志加 v1.21）

**Interfaces:**
- Consumes: Task 1 的 `SCENARIO_TYPE_COLORS` 颜色映射与"整行着色、未知保持默认"行为
- Produces: skill 文档与实现一致的着色说明

- [ ] **Step 1: SKILL.md 场景类型字段行补充着色**

在 `SKILL.md` 15 列字段表「场景类型」行末尾追加：
`**生成 Excel 时整行文字着色：正常流程绿色、异常流程红色**（见 \`scripts/generate_excel.py\`），目视即可区分成功/拒绝路径`

- [ ] **Step 2: README §1.1 能力表 + §1.2 列说明**

§1.1「场景类型标注」行末尾追加：`；**整行文字着色**（正常流程绿色 / 异常流程红色），目视即区分成功/拒绝路径`
§1.2「场景类型」行末尾追加：`；**整行文字着色：正常流程绿色、异常流程红色**`

- [ ] **Step 3: 更新日志加 v1.21**

在 v1.20 行之后追加：
`| **v1.21** | 场景类型文字着色 | 主 sheet 按场景类型给整行字体着色：**正常流程绿色 / 异常流程红色**（\`scripts/generate_excel.py\` 内置 \`SCENARIO_TYPE_COLORS\`），未知/空场景类型保持默认黑色，目视即可区分成功/拒绝路径；新增回归测试 \`test_scenario_type_row_font_colors\` |`

- [ ] **Step 4: 验证文档一致**

Run: `cd /d/AgentDev/code-scenario-testcases && grep -c "正常流程绿色" SKILL.md README.md`
Expected: SKILL.md ≥1、README.md ≥2（能力表 + 更新日志；列说明行用词为"正常流程绿色"）

Run: `grep -c "SCENARIO_TYPE_COLORS" README.md`
Expected: ≥1（更新日志）

- [ ] **Step 5: Commit**

```bash
cd /d/AgentDev/code-scenario-testcases && git add SKILL.md README.md && git commit -m "docs: v1.21 README 同步场景类型文字着色 + 更新日志"
```

---

### Task 3: 终验（回归 + 部署一致性）

- [ ] **Step 1: 全量回归**

Run: `cd /d/AgentDev/code-scenario-testcases && python scripts/test_generate_excel.py`
Expected: 4 项 PASS，输出「全部通过」

- [ ] **Step 2: 确认运行时已部署**

Run: `diff <(sed -n '1,999p' /d/AgentDev/code-scenario-testcases/SKILL.md) /c/Users/46018/.claude/skills/code-scenario-testcases/SKILL.md && echo SAME`
Expected: `SAME`（若 Task 1/2 的 commit 均已触发 install.sh）；不一致则 `./install.sh` 手动部署

- [ ] **Step 3: （可选人工）用真实 skill 跑 order-java 抽查**

在 Claude Code 中调用 `code-scenario-testcases -all`（目录 `D:/AgentDev/fixtures/order-java`），抽查生成 Excel：正常流程行绿色、异常流程行红色。此步为人工冒烟，不阻塞验收。

---

## 自审结论

- **Spec 覆盖**：整行字体着色 → Task 1 Step 3；颜色映射（正常绿/异常红/未知默认黑）→ Task 1 Step 3 + Task 2 Step 1；测试 → Task 1 Step 1-2, 4-5；文档同步 → Task 2；验证标准 → Task 1 Step 4-5 + Task 3。
- **占位符扫描**：无 TBD/TODO；每步含完整可复制代码。
- **类型一致性**：颜色常量 `SCENARIO_TYPE_COLORS` 在 generate_excel.py（Task 1）与 README 更新日志（Task 2）表述一致；测试断言颜色值与常量一致。
