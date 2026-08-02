# code-scenario-testcases 自测执行与收尾闭环 实施计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 将 skill 从「验证设计」扩展为「设计 + 执行收尾」完整自测链路，并精简现有结构。

**Architecture:** 三阶段重组——阶段 0 前置界定（新增第 0 步）→ 阶段 1 验证设计（保留第 1~6 步）→ 阶段 2 执行收尾（第 7 步增强 + 新增第 8 步），另附「人工自测 8 列模板」附录。只改 `SKILL.md` 与 `README.md`。

**Tech Stack:** Markdown 文档（无代码改动）。部署：git commit → post-commit → `install.sh` 自动部署到 `~/.claude/skills/code-scenario-testcases`。

## Global Constraints

- 只改 `D:\AgentDev\code-scenario-testcases\SKILL.md` 与 `README.md`；`references/*.md`、`scripts/*.py`、`evals/` 一律不动
- 步骤编号连续可追溯：新增第 0 步、第 8 步，原第 1~7 步编号不变
- 验证方法 / 可自动化程度两处枚举**只保留一处权威定义**（落点在 14 列字段说明表），第 3 步第 8/9 点改为引用
- README §2.5 与 SKILL.md 步骤编号、阶段划分一致
- 所有中文文案沿用现有风格（简体中文、`**粗体**` 强调、行文与现有 SKILL.md 一致）
- 每任务末尾 commit（会触发 post-commit 自动部署，属预期行为；commit 信息格式 `feat:`/`docs:`）

---

### Task 1: SKILL.md 流程总览表 → 三阶段 8 步

**Files:**
- Modify: `D:\AgentDev\code-scenario-testcases\SKILL.md:44-55`（「流程总览」表）

**Interfaces:**
- Produces: 三阶段 + 8 步的流程总览表，作为后续所有步骤编号依据

- [ ] **Step 1: 替换流程总览表**

将 SKILL.md 第 44-55 行的现有 7 行总览表替换为：

```markdown
| 阶段 | 步骤 | 做什么 | 产出 |
|---|---|---|---|
| 阶段0 前置界定 | 0 需求理解与范围界定 | 成功判据三问 + 改动范围/测试边界界定 | 自测范围说明（并入双基线） |
| 阶段1 验证设计 | 1 扫描过滤 + 建立双基线 | 扫描源码；建入口清单（代码基线）+ 业务规则清单（需求基线） | 场景 → 功能骨架 + 双基线清单 |
| 阶段1 验证设计 | 2 识别场景、功能与验证点 | 代码可测点 ∪ 需求验证点 → 每功能验证点清单 | 功能入口清单 + 验证点清单（覆盖分母） |
| 阶段1 验证设计 | 3 设计验证用例 | 逐验证点可执行化（前置+步骤+预期），套 13 类展开模板 | 14 列用例 |
| 阶段1 验证设计 | 4 高级维度补齐 | 数据流/null、时序/幂等、权限矩阵、容量/边界（A-D 默认；MC/DC、链路按需） | 维度用例（标注维度名） |
| 阶段1 验证设计 | 5 覆盖核对 | 判定点分母比对 + 需求验证点 X/Y + 逐功能 100% 承诺 | 覆盖统计 + 缺口 + 追溯矩阵 |
| 阶段1 验证设计 | 6 规则缺口审查 | 对照业务规则基线逐条反向核对（该有却没有） | rule_gaps |
| 阶段2 执行收尾 | 7 生成 Excel 与汇报 | 字段自检门禁 → 生成 xlsx → 汇报（含自测结论） | .xlsx + cases.json |
| 阶段2 执行收尾 | 8 自测执行与收尾（可选） | 冒烟先行 → FAIL 闭环 → 需求偏差确认 → 回补用例 → 结论 | PASS/FAIL 记录 + 自测结论 |
```

- [ ] **Step 2: 验证替换结果**

Run: `grep -n "^| 阶段" D:\AgentDev\code-scenario-testcases\SKILL.md | head -12`
Expected: 9 行，含 `阶段0 前置界定` 一行与 `阶段2 执行收尾` 两行

- [ ] **Step 3: Commit**

```bash
cd /d/AgentDev/code-scenario-testcases && git add SKILL.md && git commit -m "feat: v1.18 流程总览改三阶段 8 步"
```

---

### Task 2: SKILL.md 新增第 0 步（需求理解与范围界定）

**Files:**
- Modify: `D:\AgentDev\code-scenario-testcases\SKILL.md`（在「流程总览」表之后、「### 第 1 步」之前插入）

**Interfaces:**
- Consumes: Task 1 的总览表（第 0 步编号）
- Produces: 第 0 步章节，产出「自测范围说明」供第 1 步双基线引用

- [ ] **Step 1: 插入第 0 步章节**

在 SKILL.md 的 `### 第 1 步 扫描过滤与建立双基线` 之前插入：

```markdown
### 第 0 步 需求理解与范围界定
动笔前先明确"测什么、测到什么程度算过"。回答三个问题（**成功判据三问**）：

1. **给谁用**：本功能的服务角色/入口（哪个页面、哪个接口、谁调用）
2. **输入输出契约**：入参字段、返回结构、状态码（从需求文档/接口契约摘录）
3. **成功判据**：验收标准里可判定的那句话（"订单生成且库存 -1"而非"下单成功"）

再界定本次**测试边界**：
- **改动范围**：本次变更涉及的功能入口（`-diff` 定位受影响函数 + grep 调用方标"间接受影响"）；只测改动波及范围，不无限回溯
- **本功能职责 vs 外部依赖**：外部依赖（DB/缓存/第三方接口/SDK）用 mock 构造，不真测；自测只保证"你的代码在依赖给定返回时行为正确"

产出：**自测范围说明**，并入第 1 步双基线作为"测什么"的起点。
```

- [ ] **Step 2: 验证插入位置**

Run: `grep -n "^### 第 [01] 步" D:\AgentDev\code-scenario-testcases\SKILL.md | head -3`
Expected: 第 0 步紧邻第 1 步之前，第 0 步行号 < 第 1 步行号

- [ ] **Step 3: Commit**

```bash
cd /d/AgentDev/code-scenario-testcases && git add SKILL.md && git commit -m "feat: v1.18 新增第 0 步需求理解与范围界定"
```

---

### Task 3: SKILL.md 第 3 步合并重复枚举（权威定义落到 14 列表）

**Files:**
- Modify: `D:\AgentDev\code-scenario-testcases\SKILL.md:88-104`（14 列字段说明表的两行）+ `SKILL.md:85-86`（第 3 步第 8/9 点）

**Interfaces:**
- Consumes: 现有 14 列表与第 3 步第 8/9 点
- Produces: 唯一权威枚举定义（在 14 列表），第 3 步第 8/9 点引用之

- [ ] **Step 1: 14 列表两行升级为权威定义**

将 14 列字段说明表中这两行替换为：

```markdown
| 验证方法 | **权威定义**：`界面操作`（App 上点选、看弹窗/页面状态）/ `接口响应`（看返回体/状态码）/ `数据状态`（查库/缓存/SharedPreferences）/ `日志`（看 Logcat）/ `需联调`（依赖外部无法本地判定）——执行者如何观察到结果 |
| 可自动化程度 | **权威定义**：`可单测`（纯逻辑分支，JVM+mock 可自动跑）/ `可组件测试`（需 Android/Robolectric 环境）/ `可端到端`（跨界面/跨模块 UI 流程）/ `仅人工+联调`（外部契约/真机环境/需服务端配合）/ `需联调确认`（不确定不臆测） |
```

- [ ] **Step 2: 第 3 步第 8/9 点改为引用**

将第 3 步第 8 点与第 9 点替换为：

```markdown
8. **验证方法标注（怎么验证）**：每条用例标注执行者如何观察到结果——枚举定义见 14 列字段表"验证方法"行（`界面操作 / 接口响应 / 数据状态 / 日志 / 需联调`）。**验证方法与预期结果的可观察对象必须匹配**
9. **可自动化程度标注（能否自动化）**：每条用例标注验证策略分层——枚举定义见 14 列字段表"可自动化程度"行（`可单测 / 可组件测试 / 可端到端 / 仅人工+联调 / 需联调确认`）。不确定时标`需联调确认`，不臆测
```

- [ ] **Step 3: 验证枚举只出现一次完整定义**

Run: `grep -c "JVM+mock 可自动跑" D:\AgentDev\code-scenario-testcases\SKILL.md`
Expected: `1`（完整定义仅在 14 列表出现一次）

- [ ] **Step 4: Commit**

```bash
cd /d/AgentDev/code-scenario-testcases && git add SKILL.md && git commit -m "refactor: v1.18 验证方法/可自动化程度枚举收敛为 14 列表权威定义"
```

---

### Task 4: SKILL.md 第 7 步汇报增强（自测结论 + 需求偏差项）

**Files:**
- Modify: `D:\AgentDev\code-scenario-testcases\SKILL.md`（「汇报：」小节，验证计划列表之后）

**Interfaces:**
- Consumes: 现有第 7 步汇报结构
- Produces: 汇报新增自测结论与需求偏差项，供第 8 步收尾判定引用

- [ ] **Step 1: 汇报追加自测结论与需求偏差项**

在 SKILL.md 第 7 步「汇报：」小节的验证计划 4 项列表之后，追加：

```markdown
- **自测结论**：`全部通过 / 阻塞项已修复 / 存在遗留缺陷 N 项（附清单）`，并明确**是否可提测**
- 风险提示补充**需求偏差项**（代码与需求不一致、待与产品/负责人确认的事项）
```

- [ ] **Step 2: 验证汇报含自测结论**

Run: `grep -n "自测结论" D:\AgentDev\code-scenario-testcases\SKILL.md`
Expected: 出现 ≥1 行（第 7 步汇报；Task 5 完成后还会出现在第 8 步，属预期）

- [ ] **Step 3: Commit**

```bash
cd /d/AgentDev/code-scenario-testcases && git add SKILL.md && git commit -m "feat: v1.18 第 7 步汇报新增自测结论与需求偏差项"
```

---

### Task 5: SKILL.md 新增第 8 步 + 附录模板

**Files:**
- Modify: `D:\AgentDev\code-scenario-testcases\SKILL.md`（在「## 注意」之前插入第 8 步与附录；「## 注意」之后为「## 输出示例与参考模板」）

**Interfaces:**
- Consumes: Task 1 编号、Task 4 自测结论
- Produces: 第 8 步章节（可选执行闭环）+ 附录人工自测 8 列模板（快速场景入口）

- [ ] **Step 1: 插入第 8 步**

在 SKILL.md 的 `## 输出示例与参考模板` 之前插入：

```markdown
### 第 8 步 自测执行与收尾（可选——仅当本次需要执行闭环时启用；纯设计产出场景可跳过）
设计产出交付后，研发按下列顺序执行用例并收尾：

1. **冒烟先行**：先跑 P0 主流程用例，主流程断了先修，不浪费时间测分支
2. **FAIL 闭环**：`用例 FAIL → 定位原因 → 修复代码 → 重跑该条 → 重跑相邻用例`（防修复引入新问题）
3. **需求偏差确认**：代码和需求对不上时**停下确认**（产品/负责人），不擅自按自己理解测
4. **回补用例**：把本次发现的有价值场景回补进用例集，作为下次回归基线
5. **收尾交付**：汇总统计（用例总数 / PASS / FAIL / 覆盖率）+ 遗留问题清单 + **自测结论**（是否可提测）
```

- [ ] **Step 2: 插入附录模板**

在同一位置（第 8 步之后、`## 输出示例与参考模板` 之前）继续插入：

```markdown
## 附录：人工自测 8 列模板（快速场景）

**触发条件**：你只要快速人工自测、不需要 Excel 时，直接输出下列模板；需要完整 Excel 时走第 0~8 步流程。

| 编号 | 用例名称 | 前置条件 | 操作步骤 | 预期结果 | 验证方法 | 优先级 | 结果 |
|---|---|---|---|---|---|---|---|
| TC-ORD-001 | 下单-正常流程 | 登录态有效，库存=5 | 提交订单，数量=1 | HTTP 200，返回订单号，库存变 4 | 接口响应+查库 | P0 | PASS |
| TC-ORD-002 | 下单-库存不足 | 登录态有效，库存=1 | 提交订单，数量=2 | 返回"库存不足"，订单不生成 | 接口响应 | P0 | FAIL |
| ... | | | | | | | |

**字段填写纪律**（与第 3 步一致）：用例名称用"功能名-触发条件"业务语义；前置条件给具体数据、不以"调用/传"开头；操作步骤可复现（入口+参数）；预期结果可判定且单一（精确到文案/状态码）；验证方法为 `界面操作 / 接口响应 / 数据状态 / 日志 / 需联调`。
```

- [ ] **Step 3: 验证第 8 步与附录存在**

Run: `grep -n "^### 第 8 步\|^## 附录" D:\AgentDev\code-scenario-testcases\SKILL.md`
Expected: 两行均命中

- [ ] **Step 4: Commit**

```bash
cd /d/AgentDev/code-scenario-testcases && git add SKILL.md && git commit -m "feat: v1.18 新增第 8 步自测执行与收尾 + 附录人工自测模板"
```

---

### Task 6: README.md 同步（§2.5 改 8 步 + 更新日志 v1.18）

**Files:**
- Modify: `D:\AgentDev\code-scenario-testcases\README.md:253-261`（§2.5 工作流程）+ `README.md:383`（更新日志表格末行）

**Interfaces:**
- Consumes: SKILL.md 的 8 步编号与阶段划分
- Produces: README §2.5 与 SKILL.md 一致；更新日志 v1.18 条目

- [ ] **Step 1: 替换 §2.5 为三阶段 8 步**

将 README.md §2.5 的标题与 1~7 项列表替换为：

```markdown
### 2.5 工作流程（8 步，三阶段，对齐研发自测经验流程）

**阶段 0 · 前置界定**
0. **需求理解与范围界定**：成功判据三问（给谁用 / 输入输出契约 / 成功判据）+ 改动范围与测试边界界定（外部依赖 mock 不真测），产出自测范围说明并入双基线

**阶段 1 · 验证设计**
1. **扫描过滤与建立双基线**：递归收集源码，跳过构建产物/依赖/测试文件；建**入口清单（代码基线）** + **业务规则清单（需求基线，无需求时从业务常识/代码可推断规则建立，标注"待与需求确认"）**
2. **识别场景、功能与验证点**：按 Controller/业务实体/模块聚合场景；每功能建「验证点清单」（代码可测点 ∪ 需求验证点，后者对照第 0/1 步基线）
3. **设计验证用例**：逐验证点可执行化，套 13 类展开模板（复合拆原子/嵌套不笛卡尔积/边界补值），产出 14 列用例 + 字段质量自检
4. **高级维度补齐**：数据流/null 传播、时序/幂等、权限矩阵、容量/边界（A-D 默认）；MC/DC、业务链路（E/F 按需）
5. **覆盖率核对**：判定点机器枚举作分母 + 需求验证点 X/Y + 高级维度检查点 + 逐功能 100% 承诺
6. **规则缺口审查**：对照第 1 步业务规则基线逐条反向核对，无对应分支 → 写入 `coverage.rule_gaps`（疑似缺陷）

**阶段 2 · 执行收尾**
7. **生成 Excel 与汇报**：过发布门禁（字段质量自检）→ `cases.json` → `generate_excel.py` → xlsx（含覆盖说明 sheet）；汇报统计 + 功能级覆盖 + 验证计划 + **自测结论**（是否可提测）
8. **自测执行与收尾（可选）**：冒烟先行（先跑 P0）→ FAIL 闭环（定位 → 修复 → 重跑该条 → 重跑相邻）→ 需求偏差确认 → 回补用例 → 收尾交付（汇总 + 遗留清单 + 自测结论）
```

- [ ] **Step 2: 更新日志追加 v1.18 条目**

在 README.md 更新日志表格末行（v1.17 之后）追加：

```markdown
| **v1.18** | 自测执行与收尾闭环 | 工作流程重构为三阶段 8 步：新增第 0 步"需求理解与范围界定"（成功判据三问 + 改动范围/测试边界）；第 7 步汇报新增自测结论（是否可提测）与需求偏差项；新增第 8 步"自测执行与收尾"（冒烟先行 / FAIL 闭环 / 需求偏差确认 / 回补用例）；新增附录"人工自测 8 列模板"（快速场景免 Excel）；合并验证方法/可自动化程度重复枚举为 14 列表权威定义 |
```

- [ ] **Step 3: 验证 README 与 SKILL 编号一致**

Run: `grep -n "2.5 工作流程" D:\AgentDev\code-scenario-testcases\README.md`
Expected: 含 `（8 步，三阶段` 字样

- [ ] **Step 4: Commit**

```bash
cd /d/AgentDev/code-scenario-testcases && git add README.md && git commit -m "docs: v1.18 README 同步三阶段 8 步流程与更新日志"
```

---

### Task 7: 终验与部署一致性

**Files:**
- Verify: `D:\AgentDev\code-scenario-testcases\SKILL.md`、`README.md`、运行时 `C:\Users\46018\.claude\skills\code-scenario-testcases\SKILL.md`

**Interfaces:**
- Consumes: Task 1-6 全部改动
- Produces: 验收通过 + 运行时目录已同步

- [ ] **Step 1: 验收清单逐项核对**

Run: 依次执行以下 grep，全部通过才继续

```bash
cd /d/AgentDev/code-scenario-testcases
# 1. 8 步编号连续
grep -c "^### 第 [0-8] 步" SKILL.md   # Expected: 9（0~8 共 9 步）
# 2. 三阶段总览存在
grep -c "^| 阶段[0-2]" SKILL.md        # Expected: 9（总览表 9 行）
# 3. 附录模板存在
grep -c "^## 附录" SKILL.md            # Expected: 1
# 4. 自测结论出现在第 7/8 步
grep -c "自测结论" SKILL.md            # Expected: ≥2
# 5. 枚举唯一权威（JVM+mock 完整定义仅 1 处）
grep -c "JVM+mock 可自动跑" SKILL.md   # Expected: 1
# 6. README 同步
grep -c "三阶段 8 步\|8 步，三阶段" README.md  # Expected: ≥1
grep -c "v1.18" README.md              # Expected: ≥1
```

- [ ] **Step 2: 确认运行时已部署（post-commit 已触发）**

Run: `diff <(sed -n '1,999p' /d/AgentDev/code-scenario-testcases/SKILL.md) /c/Users/46018/.claude/skills/code-scenario-testcases/SKILL.md && echo SAME`
Expected: `SAME`（若 commit 均已触发 install.sh，运行时与仓库一致）

- [ ] **Step 3: 若 Step 2 不一致，手动部署**

Run: `cd /d/AgentDev/code-scenario-testcases && ./install.sh`
Expected: 输出 `install: skill 已部署到 /c/Users/46018/.claude/skills/code-scenario-testcases`，重跑 Step 2 确认 SAME

---

## 自审结论

- **Spec 覆盖**：四项优化均有对应任务——需求理解与范围界定（Task 2）、执行与收尾闭环（Task 5）、人工 8 列模板（Task 5）、精简重排（Task 1/3）；第 7 步增强在 Task 4；README 同步在 Task 6；验收在 Task 7。设计文档 §5 验收标准 6 条均有验证步骤。
- **占位符扫描**：无 TBD/TODO/「填细节」类占位；所有插入内容为完整可复制的原文。
- **一致性**：步骤编号 0~8 贯穿总览表、各章标题、README 三处一致；第 3 步引用"14 列字段表"与 Task 3 改动一致；附录触发条件中的"第 0~8 步"与流程一致。
