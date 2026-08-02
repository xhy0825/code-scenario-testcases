# v1.22 Excel 美化与完整显示 实施计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:test-driven-development（先写失败测试再实现），配合 writing-skills 的 RED-GREEN-REFACTOR。Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 让生成的 Excel 打开即读——内容换行且完全显示（行高自适应）、整体美化（斑马纹/优先级着色/统一字体/分节标题），提升阅读体验。

**Architecture:** 三处改动——① `scripts/generate_excel.py`：新增样式常量与辅助函数（`_text_width/_text_lines/fit_row_height/write_section_title/write_table_header`），重写 `build_workbook` 与 `build_coverage_sheet` 的样式与行高；② `scripts/test_generate_excel.py`：新增 4 项样式回归测试，更新 v1.21 场景着色测试排除优先级列；③ `SKILL.md` + `README.md`：补充美化说明 + 更新日志。

**Tech Stack:** Python 3 + openpyxl；Markdown。无 pytest 依赖（`python scripts/test_generate_excel.py` 直接运行）。

**Spec:** `docs/superpowers/specs/2026-08-02-excel-beautify-design.md`

## Global Constraints

- 只改 `scripts/generate_excel.py`、`scripts/test_generate_excel.py`、`SKILL.md`、`README.md`；`references/*.md`、`evals/`、`enumerate_decision_points.py` 一律不动
- 行高估算用显示宽度模型：中文/全角=2 单位、ASCII=1 单位；列宽预算留 1 单位余量，宁高勿裁
- 优先级列显示优先级色（P0 `C00000` / P1 `BF8F00` / P2 `7F7F7F`），**覆盖**场景色；其余列仍按场景色（正常绿/异常红）
- 斑马纹：偶数行 `F2F7FB`、奇数行 `FFFFFF`；表头 `4472C4` 白字加粗、行高 28
- 不改列数 / 下拉 / 覆盖说明结构；版本 v1.22；commit 格式 `feat:`/`docs:`；post-commit 自动部署属预期行为

---

### Task 1: generate_excel.py 样式重构 + 行高自适应 + 回归测试（RED-GREEN）

**Files:**
- Modify: `scripts/generate_excel.py`（样式常量 + 辅助函数 + `build_workbook` + `build_coverage_sheet`）
- Modify: `scripts/test_generate_excel.py`（新增 4 项 + 更新 v1.21 测试）

**Interfaces:**
- Produces: `ge.fit_row_height(ws, row, width_map)` 设置行高 `>= MIN_ROW_HEIGHT`；`build_workbook` 数据行斑马纹 + 优先级列优先级色 + 其余列场景色 + 表头行高 28；`build_coverage_sheet` 全表行高已设置、分节标题浅蓝、子表表头蓝底白字

- [ ] **Step 1: 写失败测试**

在 `scripts/test_generate_excel.py` 新增：

```python
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
```

同时更新 `test_scenario_type_row_font_colors`：断言时跳过优先级列（该列 v1.22 起显示优先级色），并断言优先级列 = `C00000`。

- [ ] **Step 2: 运行测试，确认失败**

Run: `cd /d/AgentDev/code-scenario-testcases && python scripts/test_generate_excel.py`
Expected: 4 项新测试 FAIL（行高未设/斑马纹无/优先级列仍为场景色/覆盖说明无行高），v1.21 场景着色测试因优先级列被覆盖而 FAIL

- [ ] **Step 3: 新增样式常量与辅助函数**

在 `COLUMN_WIDTHS` 之后插入（完整代码见 spec）：

```python
FONT_NAME = "微软雅黑"
FONT_SIZE = 10
HEADER_FONT_SIZE = 11
HEADER_FILL = PatternFill("solid", fgColor="4472C4")
SECTION_FILL = PatternFill("solid", fgColor="DDEBF7")
BAND_FILL_EVEN = PatternFill("solid", fgColor="F2F7FB")
BAND_FILL_ODD = PatternFill("solid", fgColor="FFFFFF")
PRIORITY_COLORS = {"P0": "C00000", "P1": "BF8F00", "P2": "7F7F7F"}
SECTION_TITLE_COLOR = "1F4E79"
HEADER_ROW_HEIGHT = 28
LINE_HEIGHT = 16
MIN_ROW_HEIGHT = 22
_THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
WRAP_TOP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(vertical="center", horizontal="center")
MAIN_WIDTH_MAP = {i: COLUMN_WIDTHS.get(key, 20) for i, (key, _) in enumerate(COLUMNS, start=1)}
COVERAGE_WIDTH_MAP = {1: 28, 2: 32, 3: 20, 4: 50}


def _text_width(text):
    return sum(2 if ord(ch) > 0x2E7F else 1 for ch in text)


def _chars_per_line(width):
    return max(int(width) - 1, 4)


def _text_lines(text, width):
    text = "" if text is None else str(text)
    if not text:
        return 1
    cpl = _chars_per_line(width)
    return sum(max(1, -(-_text_width(seg) // cpl)) for seg in text.split("\n"))


def fit_row_height(ws, row, width_map):
    max_lines = 1
    for col_idx, width in width_map.items():
        cell = ws.cell(row=row, column=col_idx)
        if cell.value is None:
            continue
        max_lines = max(max_lines, _text_lines(cell.value, width))
    ws.row_dimensions[row].height = max(MIN_ROW_HEIGHT, max_lines * LINE_HEIGHT + 4)
```

- [ ] **Step 4: 重写 build_workbook 样式**

表头：`HEADER_FILL` + `Font(FONT_NAME, HEADER_FONT_SIZE, bold, white)` + `CENTER` + `BORDER`，`ws.row_dimensions[1].height = HEADER_ROW_HEIGHT`。
数据行：`band_fill = BAND_FILL_EVEN if row % 2 == 0 else BAND_FILL_ODD`；每格 `fill=band_fill, alignment=WRAP_TOP, border=BORDER`；优先级列 `Font(bold=True, color=pri_color)`，其余列按场景色 `font_color`；行末 `fit_row_height(ws, row, MAIN_WIDTH_MAP)`。

- [ ] **Step 5: 重写 build_coverage_sheet**

新增 `write_section_title`（浅蓝底 + 深蓝加粗 + 换行）与 `write_table_header`（蓝底白字加粗居中）辅助函数；分节标题统一走 `write_section_title`，子表表头走 `write_table_header`，数据单元格统一 `cset`（字体+边框+换行）；末尾设列宽 `COVERAGE_WIDTH_MAP` 并对 `1..max_row` 逐个 `fit_row_height`。

- [ ] **Step 6: 运行测试，确认通过**

Run: `cd /d/AgentDev/code-scenario-testcases && python scripts/test_generate_excel.py`
Expected: 8 项 PASS，输出「全部通过」

- [ ] **Step 7: 冒烟：真实 xlsx 抽查样式**

用含 正常/异常 × P0/P1 的样例 + 覆盖说明生成 xlsx，抽查：表头行高 28、偶数行 `F2F7FB`/奇数行 `FFFFFF`、优先级列 P0 `C00000`/P1 `BF8F00`、数据行行高 ≥ 20、覆盖说明全表行高已设置。若发现纯 ASCII 长列被过度撑高（>48pt），改用显示宽度模型（Step 3 已实现）。

- [ ] **Step 8: Commit**

```bash
cd /d/AgentDev/code-scenario-testcases && git add scripts/generate_excel.py scripts/test_generate_excel.py && git commit -m "feat: v1.22 Excel 美化——行高自适应+斑马纹+优先级着色+统一字体（8 项测试全过）"
```

---

### Task 2: SKILL.md + README.md 同步美化说明

**Files:**
- Modify: `SKILL.md`（第 7 步 Excel 产出说明）
- Modify: `README.md`（§1.1 能力表、§2.1 目录结构注释、更新日志）

- [ ] **Step 1: SKILL.md 第 7 步补充美化说明**

在第 7 步"主 sheet 15 列..."行后追加：
`- 阅读体验：主 sheet 行高按内容自适应（换行完全显示）、斑马纹行带、优先级着色（P0 红/P1 琥珀/P2 灰）、统一字体；覆盖说明分节标题/子表表头统一美化`

- [ ] **Step 2: README §1.1 能力表加行**

在「测试结果跟踪」行后追加：
`| **Excel 美化与阅读体验** | 行高按内容自适应（换行内容完全显示）、数据行斑马纹、优先级着色（P0 红 / P1 琥珀 / P2 灰）、统一微软雅黑字体；覆盖说明分节标题浅蓝、子表蓝底白字表头 |`

- [ ] **Step 3: README §2.1 目录结构注释同步**

把 `主 sheet（15 列 + 样式 + 筛选 + 冻结首行）` 改为 `主 sheet（15 列 + 行高自适应/斑马纹/优先级着色 + 筛选 + 冻结首行）`。

- [ ] **Step 4: 更新日志加 v1.22**

在 v1.21 行之后追加：
`| **v1.22** | Excel 美化与完整显示 | 行高按内容自适应估算（中文 2/ASCII 1 显示宽度模型），换行内容完全显示；主表斑马纹（偶行浅蓝/奇行白）+ 优先级着色（P0 红/P1 琥珀/P2 灰，覆盖场景色）+ 统一微软雅黑字体 + 表头行高；覆盖说明分节标题浅蓝、子表蓝底白字表头、D 列加宽、全表行高自适应；回归测试扩至 8 项 |`

- [ ] **Step 5: Commit**

```bash
cd /d/AgentDev/code-scenario-testcases && git add SKILL.md README.md && git commit -m "docs: v1.22 README 同步 Excel 美化说明 + 更新日志"
```

---

### Task 3: 终验（回归 + 部署一致性）

- [ ] **Step 1: 全量回归**

Run: `cd /d/AgentDev/code-scenario-testcases && python scripts/test_generate_excel.py`
Expected: 8 项 PASS，输出「全部通过」

- [ ] **Step 2: 确认运行时已部署**

Run: `diff <(sed -n '1,999p' /d/AgentDev/code-scenario-testcases/SKILL.md) /c/Users/46018/.claude/skills/code-scenario-testcases/SKILL.md && echo SAME`
Expected: `SAME`；不一致则 `./install.sh` 手动部署

- [ ] **Step 3: 推送**

```bash
cd /d/AgentDev/code-scenario-testcases && git push origin main
```

- [ ] **Step 4: （可选人工）用真实 skill 跑 order-java 抽查**

在 Claude Code 中调用 `code-scenario-testcases -all`（目录 `D:/AgentDev/fixtures/order-java`），人工打开 Excel 确认：内容换行完全显示、斑马纹、优先级着色、覆盖说明分节清晰。此步不阻塞验收。

---

## 自审结论

- **Spec 覆盖**：行高自适应 → Task 1 Step 3（`fit_row_height`）+ Step 4/5；斑马纹 → Task 1 Step 4；优先级着色 → Task 1 Step 4 + 测试 Step 1；覆盖说明美化 → Task 1 Step 5；文档同步 → Task 2；验证标准 → Task 1 Step 6-7 + Task 3。
- **占位符扫描**：无 TBD/TODO；每步含完整可复制代码。
- **类型一致性**：`PRIORITY_COLORS`/`BAND_FILL_*` 在 generate_excel.py（Task 1）与 README 更新日志（Task 2）表述一致；测试断言颜色值与常量一致。
