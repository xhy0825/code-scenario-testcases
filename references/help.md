# code-scenario-testcases · 使用帮助

## 它能做什么

从业务代码（+ 可选需求/契约）自动设计 **业务场景 → 功能 → 验证点 → 用例** 四级验证方案，生成研发自测用例 Excel。每条用例回答三个问题：**验证什么**（验证点）、**怎么验证**（验证方法）、**能否自动化**（可自动化程度），并附覆盖核算（判定点 + 需求验证点）、规则缺口审查、验证计划与风险提示，支持主流语言。

## 指令

| 指令 | 说明 |
|---|---|
| `code-scenario-testcases -all` | 全量扫描指定目录，生成全部自测用例 |
| `code-scenario-testcases -diff` | 只分析**当前未提交**的变更（工作区/暂存区/新文件） |
| `code-scenario-testcases -diff [commitId]` | 分析**该提交之后到当前**的变更（适合"上个版本后我改了什么就测什么"） |
| `code-scenario-testcases -doc` | **仅生成当前项目文档**（6 份版本化 doc：项目背景/需求分析/架构设计/业务流程/功能清单/使用场景），不生成用例 Excel |
| `code-scenario-testcases -doc -diff` | **仅按变更更新文档版本号**：按 `git diff` 判断，有实质变化的主题文档版本 +1、无变化复用当前版本；不生成用例 Excel |
| `code-scenario-testcases -xlsx` | **只生成 Excel，跳过文档**：跳过阶段0A（第 0.5 步），直接扫描生成用例 Excel，不产出 doc 文档 |
| `code-scenario-testcases -help` | 显示本帮助 |

## 用法示例

```
code-scenario-testcases -all
code-scenario-testcases -diff
code-scenario-testcases -diff a1b2c3d
code-scenario-testcases -doc            # 只产出 testcase/<项目名>/doc/ 下 6 份版本化文档
code-scenario-testcases -doc -diff      # 只按变更更新文档版本号（有变化主题 +1、无变化复用）
code-scenario-testcases -xlsx           # 跳过文档，只生成用例 Excel
code-scenario-testcases -all            # 可附加需求/验收标准，触发规则缺口审查
```

**输入方式**：代码目录（推荐）/ 单个文件 / 单个函数 / 直接粘贴代码片段。

## 输出

- 自测用例 Excel：skill 安装目录的 `testcase/<项目名>/<项目名>_<日期>_<版本>.xlsx`
- 前置项目文档（`-doc` / 各模式阶段0A）：`testcase/<项目名>/doc/` 下 6 份版本化 Markdown（项目背景/需求分析/架构设计/业务流程/功能清单/使用场景）
- 「覆盖说明」sheet：覆盖统计（入口/判定点/需求验证点）、可自动化程度分布、验证点→用例追溯矩阵、标注区（缺口 / 验证缺口 / 无法静态验证 / 规则缺口）

## 核心能力

- **验证点设计**：每条用例含验证点（验证什么）、验证方法（怎么验证：界面/接口/数据/日志/联调）、可自动化程度（能否自动化：可单测/组件/端到端/仅人工联调）
- **验证计划**：按可自动化程度分层输出验证策略（冒烟→回归→深度→联调）
- **机器枚举判定点**作覆盖率硬分母——漏测亮红灯，不虚增
- **条件展开**：复合条件拆分、边界值、状态机、调用链、嵌套不笛卡尔积
- **需求驱动**：附带需求/契约时，生成需求验证点覆盖 X/Y，并标出"该有却没有"的规则缺口
- **语言无关**：Java / Go / Python / JS / TS / C# / C++ / Rust

## 小贴士

- 想快速上手：直接说"用 code-scenario-testcases 生成测试用例"，并给一个代码目录
- 想只测改动：改了代码没提交，用 `-diff` 最省
- 完整手册见 skill 内 `README.md`
