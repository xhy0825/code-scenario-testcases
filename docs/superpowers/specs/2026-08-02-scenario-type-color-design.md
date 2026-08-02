# code-scenario-testcases 场景类型文字着色 设计文档

> **日期**：2026-08-02
> **版本**：v1.21（本设计对应版本）
> **状态**：待实施

## 背景与目标

v1.19 引入「场景类型」列（正常流程/异常流程）后，用例的成败路径已被显式标注，但打开 Excel 时仍需逐行阅读「场景类型」或「预期结果」才能区分成功与拒绝路径。目标：**让正常流程与异常流程在 Excel 中一眼可辨**——按场景类型给整行文字着色，**正常流程绿色、异常流程红色**，目视即可扫描"成功主路径"与"拒绝/失败分支"，进一步提升自测执行时的可读性。

## 设计方案

### 1. 主 sheet 数据行整行字体着色

- 颜色映射常量 `SCENARIO_TYPE_COLORS`：
  - `正常流程` → `#008000`（绿）
  - `异常流程` → `#FF0000`（红）
- 在 `build_workbook` 写数据行时，按 `case.get("scenario_type")` 取色并应用到**该行全部 15 列**的 `Font(color=...)`
- **未知/空场景类型**（如留空待填）不设色，保持默认黑色——不臆测分类
- 表头、冻结首行、筛选、下拉等既有行为不变

### 2. 测试覆盖

- `scripts/test_generate_excel.py` 新增 `test_scenario_type_row_font_colors`：
  - 正常流程行整行字体为绿色（`008000`）
  - 异常流程行整行字体为红色（`FF0000`）
  - 场景类型列本身颜色一致
- 辅助函数 `_normalize_color` 统一处理 openpyxl 颜色读回格式（`RRGGBB`/`00RRGGBB`/`FFRRGGBB` 取后 6 位），并对 theme/indexed 默认色（访问 `.rgb` 会抛异常）防御返回 None

## 文件改动清单

| 文件 | 改动 |
|---|---|
| `scripts/generate_excel.py` | 新增 `SCENARIO_TYPE_COLORS` 常量；`build_workbook` 数据行循环内按场景类型给整行设置字体色 |
| `scripts/test_generate_excel.py` | 新增 `test_scenario_type_row_font_colors` + `_normalize_color` 辅助函数 |
| `SKILL.md` | 15 列字段表「场景类型」行补充"生成 Excel 时整行文字着色：正常流程绿色、异常流程红色" |
| `README.md` | §1.1 能力表「场景类型标注」行 + §1.2 主 sheet 列说明「场景类型」行补充着色说明；更新日志加 v1.21 |

## 边界与不做

- 不做**整行底色（fill）填充**——用户明确要求"文字颜色"，字体色改动最小、打印/拷贝不污染
- 不改变场景类型判定规则（沿用 v1.19：按路径最终现象分类）
- 不着色「覆盖说明」sheet 的场景类型分布数字（保持中性）
- 不动 `references/*.md`、`evals/`、`enumerate_decision_points.py`

## 验证标准

1. `python scripts/test_generate_excel.py` 全部通过（含新增着色测试）
2. 用含 正常/异常/空 三种场景类型的用例生成 xlsx，抽查：正常行绿、异常行红、空场景类型保持默认黑
3. SKILL.md / README.md 中着色说明与实现一致
4. 现有 4 项回归测试不回归（列数、下拉、覆盖说明分布）

## 变更日志条目（预填）

> **v1.21** | 场景类型文字着色 | 主 sheet 按场景类型给整行字体着色：**正常流程绿色 / 异常流程红色**（`scripts/generate_excel.py` 内置 `SCENARIO_TYPE_COLORS`），未知/空场景类型保持默认黑色，目视即可区分成功/拒绝路径；新增回归测试 `test_scenario_type_row_font_colors`
