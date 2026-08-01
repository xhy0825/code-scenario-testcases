# 语言适配参考

识别不同语言下"功能入口"和"条件分支"的写法差异。**核心建模是语言无关的**——`if/else/switch` 在所有语言里都是"判定点"，差异只在语法书写与入口识别方式。

## 文件扩展名 → 语言

| 语言 | 扩展名 |
|---|---|
| Java | `.java` |
| Kotlin | `.kt`, `.kts` |
| Go | `.go` |
| Python | `.py` |
| JavaScript | `.js`, `.mjs`, `.cjs` |
| TypeScript | `.ts`, `.tsx` |
| C# | `.cs` |
| C / C++ | `.c`, `.h`, `.cpp`, `.hpp`, `.cc` |
| Rust | `.rs` |
| PHP | `.php` |
| Ruby | `.rb` |
| Swift | `.swift` |

## 功能入口识别启发式

"功能" = 可被外部触发的业务入口。各语言常见形态：

### Java / Spring Boot
- 注解方法：`@RequestMapping` / `@GetMapping` / `@PostMapping` / `@PutMapping` / `@DeleteMapping`（Controller 层）
- 每个注解方法 = 一个功能；Controller 类 = 一个业务场景
- 若只有 Service 层：public 业务方法视为功能候选

### Go
- 注册了路由的 handler：`r.GET("/path", handler)`、`router.POST(...)`、`gin.Context` handler
- 导出的 `func Xxx` 但未注册路由的，按业务函数处理

### Python
- 装饰器方法：`@app.route(...)`（Flask）、`@router.get/post(...)`（FastAPI）、Django View 的类方法
- 无装饰器时：类中的业务方法、模块级业务函数

### JS / TS (Node)
- 框架路由：`app.get/post/put/delete`、`router.*`、NestJS 的 `@Controller`/`@Get` 装饰器
- 导出的异步业务函数

### C# / ASP.NET
- `[HttpGet]`/`[HttpPost]`/`[Route]`/`[ApiController]` 特性修饰的 Controller 方法

### 无框架 / 通用
- 导出函数、公开方法、命令行入口（`main`）调用的业务函数
- 以业务语义命名的方法（create/update/delete/query/login 等）

## 条件分支语法对照

| 条件逻辑 | Java/JS/C#/C++ | Go | Python | Rust |
|---|---|---|---|---|
| if/else | `if (a) {...} else {...}` | `if a { ... } else { ... }` | `if a: ... else: ...` | `if a { ... } else { ... }` |
| else-if 链 | `else if` | `else if` | `elif` | `else if` |
| 多路分支 | `switch(x) { case 1: ... default: }` | `switch x { case 1: ... default: }` | `match x: case 1: ...`（无 switch） | `match x { 1 => ..., _ => ... }` |
| 选择表达式 | `x ? a : b` | 无三目（用 if/else） | `a if cond else b` | 无三目（用 match） |
| 提前返回/守卫 | `if (!ok) return error;` | `if err != nil { return err }` | `if not ok: return ...` | `if !ok { return Err }` |

> 说明：
> - Python 没有 `switch`，多路分支用 `if/elif` 链或 `match`；Go 有 `switch`/`select` 但没有三目运算符
> - 语言不同不影响用例生成逻辑——统一按"判定点 → 分支"建模

## 判定点识别（统一）

各语言条件结构形态不同（对照上表），但**用例展开规则统一**——每个判定点 → 分支一条用例。**展开细则（嵌套/复合/边界/状态机/循环/try-catch 等）以 `code-snippet-templates.md` 为准**，本文件只负责"识别判定点的语言写法"。

## 隐含条件识别清单

除了显式 `if/else/switch`，还要识别以下"隐含条件"并生成对应用例（否则是漏测）：

| 隐含条件 | 识别特征 | 示例 |
|---|---|---|
| lambda / stream / 推导式内条件 | Java `filter/map` lambda、Python 列表推导 `if`、Go 无 | `list.stream().filter(x -> x.getPrice() > 0)` |
| 异常路径 | `try/catch`、`except`、Go 多返回值 `if err != nil` | `catch (Exception e) { return error(...) }` |
| 判空 / 可选链 | Java `Objects.isNull`、JS/Kotlin `?.`、`??`、Python `if x is None` | `a?.b?.c ?? "默认"` |
| 类型判断 / 转换 | Java `instanceof`、Python `isinstance`、Go 类型断言 `x.(T)` | `if (obj instanceof User)` |
| switch fall-through / 多 case 合并 | `case 1: case 2:`、Go `case 1, 2:` | `case 1: case 2: doX()` |
| 集合空 / 大小 / 越界 | `.isEmpty()`、`.length() == 0`、`len(...)` | `if (list.isEmpty())` |
| 字符串匹配 / 判空 | `.isEmpty()`、`== ""`、`.matches(...)` | `if (name == null \|\| name.isEmpty())` |

> 复合条件（`&&`/`||`）里的每个原子条件都要能对应到一条用例（见 `code-snippet-templates.md` 场景 6）。枚举脚本会对条件行打 `compound` 标记辅助核对。

## 注释与信息提取

- 读取中文注释、字段名、错误常量可辅助推断业务语义与预期结果文案
- 分支里直接返回/抛出的错误字符串（`throw new BizException("库存不足")`、`return errors.New("...")`、`return Response.error("xxx")`），该字符串即"不通过条件预期展示的结果"，**原样摘录**进预期结果
- 若分支代码调用外部系统（DB/缓存/第三方），预期结果按"调用成功/调用失败"两种情况补充，并在汇报中标注该依赖无法静态验证
