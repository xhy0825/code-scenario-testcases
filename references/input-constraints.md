# 输入约束分析（参数约束注解 + 等价类）

代码驱动只测"代码里写的分支"，本文件补充**代码隐含的输入约束**：从框架参数注解与类型推断约束，生成等价类用例，独立于分支覆盖。

## 各语言/框架约束注解对照

| 约束 | Java (Bean Validation) | Python (Pydantic) | Go (validator) | JS/TS (class-validator) |
|---|---|---|---|---|
| 非空 | `@NotNull` | `constrain` / 类型必填 | `validate:"required"` | `@IsNotEmpty` |
| 非空串 | `@NotBlank` / `@NotEmpty` | `min_length=1` | `min=1` | `@IsNotEmpty` |
| 最小值 | `@Min(n)` | `ge=n` | `min=n` | `@Min(n)` |
| 最大值 | `@Max(n)` | `le=n` | `max=n` | `@Max(n)` |
| 长度 | `@Size(min,max)` / `@Length` | `min_length`/`max_length` | `len` | `@Length` |
| 格式 | `@Pattern(regexp)` | `pattern=` | `match=` | `@Matches` |
| 邮箱 | `@Email` | `EmailStr` | `email` | `@IsEmail` |
| 正数/负数 | `@Positive` / `@Negative` | `gt=0` / `lt=0` | | `@IsPositive` |

> 枚举脚本会对 `@NotNull/@Min/@Max/...` 等注解行打 `constraint` 标记并计入 `dp.json.total_input_constraints`，辅助核对。

## 等价类生成规则

对每个带约束/带类型的参数，生成：

1. **合法等价类**：满足所有约束的典型值（1 条）
2. **非法等价类**：违反约束的值（每个约束至少 1 条，用例名标注违规约束）
3. **边界值**：约束边界、边界±1（如 `@Min(1)` → 0/1/2）
4. **空/缺省**：null、缺省值、空串/空集合

约束用例标注"输入约束"，优先级：必填/格式校验 P0，范围/长度 P1。

## 示例

```java
public Order create(@NotNull @Min(1) Long userId,
                    @NotBlank String skuId,
                    @Max(10000) BigDecimal amount) {
```

生成约束用例（独立于函数内分支）：

| 用例编号 | 用例名称 | 前置条件 | 操作步骤 | 优先级 | 预期结果 | 备注 |
|---|---|---|---|---|---|---|
| TC-001 | userId非空约束 | 调用接口 | 传 userId=null | P0 | 参数校验失败（框架或代码拦截） | 输入约束：@NotNull |
| TC-002 | userId最小值边界 | 调用接口 | 传 userId=0 | P0 | 参数校验失败 | 输入约束：@Min(1)，边界 0/1 |
| TC-003 | skuId非空串约束 | 调用接口 | 传 skuId="" | P0 | 参数校验失败 | 输入约束：@NotBlank |
| TC-004 | amount最大值边界 | 调用接口 | 传 amount=10001 | P1 | 参数校验失败 | 输入约束：@Max(10000)，边界 10000/10001 |
| TC-005 | 合法输入 | 参数全合法 | 传合法值 | P1 | 进入业务逻辑 | 合法等价类 |

> 说明：约束若由框架层（如 Spring 校验器/Pydantic）拦截，预期结果标"框架校验失败（422/400）"；若代码显式校验，标对应错误文案。
