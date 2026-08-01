# 语言适配参考

本文件用于：当被测代码是不同语言时，识别"功能入口"和"条件分支"的写法差异。**核心建模是语言无关的**——`if/else/switch` 在所有语言里都是"判定点"，差异只在语法书写与入口识别方式。

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

## 判定点建模规则（统一）

无论哪种语言，每个判定节点都按同一套规则拆：

- `if (条件)` → 【真分支】+【假分支】
- `else if (条件)` → 各自为独立判定点
- `else` → 前序判定的兜底假分支
- `switch(x)` / `select` → 每个 `case` + `default`
- 三目 / guard clause / 提前 return → 真分支（继续往下走）+ 假分支（返回/报错）
- 循环内 `if` + `break`/`continue` → 记为该循环的一次判定
- `try/catch`（Java/C#/Python/JS）→ catch 分支可视为"异常路径"判定，生成一条用例，预期结果写捕获后的行为（如错误响应、降级文案）

## 注释与信息提取

- 读取中文注释、字段名、错误常量可辅助推断业务语义与预期结果文案
- 分支里直接返回/抛出的错误字符串（`throw new BizException("库存不足")`、`return errors.New("...")`、`return Response.error("xxx")`），该字符串即"不通过条件预期展示的结果"，**原样摘录**进预期结果
- 若分支代码调用外部系统（DB/缓存/第三方），预期结果按"调用成功/调用失败"两种情况补充，并在汇报中标注该依赖无法静态验证
