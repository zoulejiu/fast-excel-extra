# LockHandler 使用指南

## 功能概述

`LockHandler` 是一个 Excel 单元格锁定处理器，支持在导出 Excel 时控制哪些列可以编辑，哪些列被保护。

## 核心特性

✅ **双重配置方式**
- 支持通过 `@ExcelEditable` 注解控制（适合固定规则）
- 支持通过参数动态指定可编辑字段（推荐，更灵活）
- **两种方式可以兼容使用**（AND 逻辑）

✅ **工作表保护**
- 自动启用工作表保护
- 支持设置密码保护
- 允许调整列宽和行高（不影响正常使用）

✅ **表头自动保护**
- 表头行始终被锁定，防止误操作

✅ **灵活的组合策略**（OR 关系 + 参数优先）
- 只使用注解：按注解配置生效
- 只使用参数：参数中的字段可编辑，其他字段按默认规则
- 同时使用：参数中的字段强制可编辑（覆盖注解），参数之外的字段按注解配置

## 使用方式

### 方式一：注解方式（传统方式）

#### 1. 在数据类中使用 `@ExcelEditable` 注解

```kotlin
import cn.idev.excel.annotation.ExcelProperty
import com.excel.annotation.ExcelEditable

data class User(
    @ExcelProperty("用户ID")
    @ExcelEditable(false)  // 不可编辑
    val id: Long,

    @ExcelProperty("姓名")
    val name: String,  // 默认可编辑

    @ExcelProperty("年龄")
    @ExcelEditable(true)  // 明确标记可编辑
    val age: Int
)
```

#### 2. 注册 Handler（不指定 editableFieldNames）

```kotlin
FastExcel.write("output.xlsx")
    .registerWriteHandler(LockHandler(User::class))
    .head(User::class.java)
    .sheet("用户列表")
    .doWrite(dataList)
```

**优点**: 适合固定的业务规则，配置在模型类中
**缺点**: 不够灵活，修改需要改代码和重新编译

---

### 方式二：动态指定（推荐）⭐

#### 1. 数据类无需注解

```kotlin
data class User(
    @ExcelProperty("用户ID")
    val id: Long,

    @ExcelProperty("姓名")
    val name: String,

    @ExcelProperty("年龄")
    val age: Int,

    @ExcelProperty("部门")
    val department: String
)
```

#### 2. 导出时动态指定可编辑字段

```kotlin
// 只允许编辑"姓名"和"年龄"
FastExcel.write("output.xlsx")
    .registerWriteHandler(
        LockHandler(
            User::class,
            editableFieldNames = setOf("name", "age")
        )
    )
    .head(User::class.java)
    .sheet("用户列表")
    .doWrite(dataList)
```

**优点**: 
- ✅ 灵活性高，可以根据业务场景动态调整
- ✅ 同一个数据类可以用于不同的导出场景
- ✅ 无需修改模型类

---

### 方式三：注解与参数兼容使用（灵活组合）⭐⭐

注解和参数可以配合使用，提供更灵活的权限控制。

#### 兼容规则（OR 关系 + 参数优先）

**规则：参数中的字段强制可编辑（覆盖注解），参数之外的字段按注解配置**

| 注解配置 | 参数配置 | 最终结果 | 说明 |
|---------|---------|---------|------|
| `@ExcelEditable(true)` | 在 `editableFieldNames` 中 | ✅ 可编辑 | 参数覆盖（强制可编辑） |
| `@ExcelEditable(true)` | 不在 `editableFieldNames` 中 | ✅ 可编辑 | 按注解配置 |
| `@ExcelEditable(false)` | 在 `editableFieldNames` 中 | ✅ 可编辑 | **参数覆盖注解限制** |
| `@ExcelEditable(false)` | 不在 `editableFieldNames` 中 | ❌ 不可编辑 | 按注解配置 |
| 无注解（默认可编辑） | 在 `editableFieldNames` 中 | ✅ 可编辑 | 参数覆盖（强制可编辑） |
| 无注解（默认可编辑） | 不在 `editableFieldNames` 中 | ✅ 可编辑 | 按注解配置（默认可编辑） |

#### 示例：分权限导出

```kotlin
data class Document(
    @ExcelProperty("文档ID")
    @ExcelEditable(false)  // 注解：ID 永远不可编辑
    val id: Long,

    @ExcelProperty("标题")
    @ExcelEditable(true)  // 注解：标题默认可编辑
    val title: String,

    @ExcelProperty("内容")
    @ExcelEditable(true)  // 注解：内容默认可编辑
    val content: String,

    @ExcelProperty("状态")
    @ExcelEditable(false)  // 注解：状态不可编辑
    val status: String,

    @ExcelProperty("创建时间")
    @ExcelEditable(false)  // 注解：创建时间不可编辑
    val createTime: String
)

// 场景1：普通用户 - 只能编辑标题
// 注解允许：title, content 可编辑；id, status, createTime 不可编辑
// 参数指定：title
// 最终可编辑：title（参数中） + content（不在参数中，按注解可编辑） = title, content
// 但实际想要的是只允许编辑 title，所以需要配合注解或者设计不同的模型
FastExcel.write("document_user.xlsx")
    .registerWriteHandler(
        LockHandler(
            Document::class,
            editableFieldNames = setOf("title")  // 强制 title 可编辑
        )
    )
    .head(Document::class.java)
    .sheet("文档")
    .doWrite(dataList)

// 场景2：编辑 - 可以编辑标题和内容  
// 注解允许：title, content
// 参数指定：title, content
// 最终可编辑：title, content（都在参数中或按注解）
FastExcel.write("document_editor.xlsx")
    .registerWriteHandler(
        LockHandler(
            Document::class,
            editableFieldNames = setOf("title", "content")
        )
    )
    .head(Document::class.java)
    .sheet("文档")
    .doWrite(dataList)

// 场景3：管理员 - 可以编辑状态字段（覆盖注解限制）
// 注解禁止：id, status, createTime
// 参数指定：status（覆盖注解限制）
// 最终可编辑：title, content（按注解） + status（参数覆盖）
FastExcel.write("document_admin.xlsx")
    .registerWriteHandler(
        LockHandler(
            Document::class,
            editableFieldNames = setOf("status")  // 强制 status 可编辑（覆盖注解限制）
        )
    )
    .head(Document::class.java)
    .sheet("文档")
    .doWrite(dataList)
```

#### 优势

✅ **基础安全规则**：通过注解定义默认的可编辑规则  
✅ **灵活覆盖能力**：通过参数可以覆盖注解限制，实现特殊场景需求（如管理员强制编辑某些字段）  
✅ **代码可维护性**：默认规则在模型层，特殊规则在调用层，职责清晰

#### 使用建议

1. **注解用于定义常规规则**：大部分情况下的可编辑性
2. **参数用于特殊场景覆盖**：
   - 管理员需要编辑通常不可编辑的字段（如状态、时间戳）
   - 特定业务流程需要临时开放某些字段的编辑权限

---

## 常见使用场景

### 场景1: 绩效考核分阶段导出

```kotlin
// 草稿阶段：员工只能编辑自评分数和自评说明
FastExcel.write("performance_draft.xlsx")
    .registerWriteHandler(
        LockHandler(
            PerformanceItem::class,
            editableFieldNames = setOf("selfRating", "selfComment")
        )
    )
    .head(PerformanceItem::class.java)
    .sheet("绩效考核")
    .doWrite(dataList)

// 审核阶段：领导只能编辑审核分数和审核说明
FastExcel.write("performance_review.xlsx")
    .registerWriteHandler(
        LockHandler(
            PerformanceItem::class,
            editableFieldNames = setOf("auditRating", "auditComment")
        )
    )
    .head(PerformanceItem::class.java)
    .sheet("绩效考核")
    .doWrite(dataList)

// 终审阶段：部门领导只能编辑最终分数和最终评价
FastExcel.write("performance_final.xlsx")
    .registerWriteHandler(
        LockHandler(
            PerformanceItem::class,
            editableFieldNames = setOf("finalRating", "finalRatingComment"),
            protectPassword = "admin123"  // 设置密码保护
        )
    )
    .head(PerformanceItem::class.java)
    .sheet("绩效考核")
    .doWrite(dataList)
```

### 场景2: 数据导入模板（部分字段可填写）

```kotlin
// 生成导入模板，只允许用户填写"姓名"、"年龄"、"部门"
// "用户ID" 由系统生成，不允许手动填写
FastExcel.write("import_template.xlsx")
    .registerWriteHandler(
        LockHandler(
            User::class,
            editableFieldNames = setOf("name", "age", "department")
        )
    )
    .head(User::class.java)
    .sheet("用户信息导入模板")
    .doWrite(emptyList())  // 空数据，只生成表头
```

### 场景3: 只读导出（完全锁定）

```kotlin
// 导出报表，所有字段都不可编辑
FastExcel.write("readonly_report.xlsx")
    .registerWriteHandler(
        LockHandler(
            Report::class,
            editableFieldNames = emptySet()  // 空集合 = 完全锁定
        )
    )
    .head(Report::class.java)
    .sheet("数据报表")
    .doWrite(dataList)
```

### 场景4: 配合 excludeColumnFieldNames 使用

```kotlin
// 草稿阶段：隐藏审核相关字段，只显示和编辑自评字段
FastExcel.write("draft.xlsx")
    .registerWriteHandler(
        LockHandler(
            PerformanceItem::class,
            editableFieldNames = setOf("selfRating", "selfComment")
        )
    )
    .head(PerformanceItem::class.java)
    .excludeColumnFieldNames(
        listOf("auditRating", "auditComment", "finalRating", "finalRatingComment")
    )
    .sheet("绩效考核")
    .doWrite(dataList)
```

---

## API 参数说明

### LockHandler 构造函数

```kotlin
class LockHandler<T : Any>(
    private val dataClass: KClass<T>,                    // 数据类的 KClass（必填）
    private val editableFieldNames: Set<String>? = null, // 可编辑字段名集合（可选）
    private val enableProtection: Boolean = true,        // 是否启用工作表保护（默认 true）
    private val protectPassword: String? = null          // 工作表保护密码（可选）
)
```

| 参数 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `dataClass` | `KClass<T>` | - | 数据类的 Kotlin Class（必填） |
| `editableFieldNames` | `Set<String>?` | `null` | 可编辑的字段名集合。如果指定，则只有这些字段可编辑，其他字段均被锁定。如果不指定（null），则回退到使用 `@ExcelEditable` 注解 |
| `enableProtection` | `Boolean` | `true` | 是否启用工作表保护。启用后，锁定的单元格将无法编辑 |
| `protectPassword` | `String?` | `null` | 工作表保护密码。如果设置，取消保护时需要输入密码 |

---

## 注意事项

### 1. 字段名匹配

`editableFieldNames` 中指定的是 **字段名（field name）**，而不是 Excel 表头名称。

```kotlin
data class User(
    @ExcelProperty("用户ID")   // 表头是"用户ID"
    val id: Long               // 字段名是 "id"
)

// ❌ 错误：使用表头名称
LockHandler(User::class, editableFieldNames = setOf("用户ID"))

// ✅ 正确：使用字段名
LockHandler(User::class, editableFieldNames = setOf("id"))
```

### 2. 优先级规则

当同时使用注解和 `editableFieldNames` 参数时，**参数优先级更高**。

```kotlin
data class User(
    @ExcelEditable(true)  // 注解标记为可编辑
    val name: String
)

// 参数中未包含 "name"，所以最终 name 字段会被锁定
LockHandler(User::class, editableFieldNames = setOf("age"))
```

### 3. 表头始终被保护

无论如何配置，表头行（第一行）始终被锁定，防止用户误修改列名。

### 4. 列宽和行高调整

即使启用了工作表保护，用户仍然可以调整列宽和行高，这不会影响正常使用。

---

## 测试示例

运行测试文件查看效果：

```bash
# 运行完整测试（包含多种场景）
kotlin src/main/test/com/excel/LockHandlerDynamicTest.kt

# 运行绩效考核场景测试
kotlin src/main/test/com/excel/DynamicHeaderTest.kt
```

生成的文件位置：`D:/test/`

---

## 常见问题

### Q1: 为什么我指定了 editableFieldNames，但所有字段还是可编辑？

**A**: 检查以下几点：
1. 确认 `enableProtection = true`（默认就是 true）
2. 确认字段名拼写正确（使用字段名，不是表头名称）
3. 打开 Excel 后，检查"审阅"选项卡，确认工作表已被保护

### Q2: 如何让所有字段都不可编辑？

**A**: 传入空集合

```kotlin
LockHandler(User::class, editableFieldNames = emptySet())
```

### Q3: 如何让所有字段都可编辑？

**A**: 两种方式：
1. 不传 `editableFieldNames` 参数，且数据类中不使用 `@ExcelEditable(false)` 注解
2. 设置 `enableProtection = false`

### Q4: 注解和参数同时使用时，谁的优先级更高？

**A**: **参数优先级更高**（参数可以覆盖注解）。
- 规则：参数中的字段强制可编辑（覆盖注解限制），参数之外的字段按注解配置
- 这样设计是为了支持特殊场景（如管理员需要临时编辑某些系统字段）

示例：
```kotlin
data class User(
    @ExcelEditable(false)  // 注解禁止
    val id: Long,
    
    @ExcelEditable(true)   // 注解允许
    val name: String,
    
    @ExcelEditable(false)  // 注解禁止
    val createTime: String
)

// 参数中包含 "id"，id 将可编辑（参数覆盖注解）
LockHandler(User::class, editableFieldNames = setOf("id"))
// 结果：
// - id: 可编辑（参数覆盖）
// - name: 可编辑（不在参数中，按注解）
// - createTime: 不可编辑（不在参数中，按注解）
```

### Q5: 密码保护后，如何取消保护？

**A**: 在 Excel 中：
1. 点击"审阅"选项卡
2. 点击"撤销工作表保护"
3. 输入密码（如果设置了密码）

---

## 总结

| 使用场景 | 推荐方式 | 示例 |
|---------|---------|------|
| 固定业务规则 | 注解方式 | `LockHandler(User::class)` |
| 灵活的导出场景 | 动态指定 | `LockHandler(User::class, editableFieldNames = setOf("name"))` |
| 基础安全 + 权限控制 | 注解 + 参数（推荐⭐） | 注解定义永远不可编辑的字段，参数根据角色限制 |
| 分阶段权限控制 | 动态指定 | 不同阶段传不同的 `editableFieldNames` |
| 导入模板 | 动态指定 | 只允许用户填写部分字段 |
| 只读报表 | 动态指定 | `editableFieldNames = emptySet()` |

### 推荐实践

1. **小型项目/简单场景**：使用动态指定方式（方式二），灵活简单

2. **中大型项目/复杂权限**：使用注解 + 参数组合（方式三）⭐
   ```kotlin
   // 模型层：定义基础规则
   data class Order(
       @ExcelEditable(false) val id: Long,        // 系统字段，默认不可编辑
       @ExcelEditable(false) val createTime: String,  // 系统字段，默认不可编辑
       @ExcelEditable(true) val amount: Double,   // 业务字段，默认可编辑
       @ExcelEditable(false) val status: String   // 状态字段，默认不可编辑
   )
   
   // 业务层：根据用户角色动态控制
   when (userRole) {
       "USER" -> LockHandler(Order::class)  // 按注解：amount 可编辑
       "MANAGER" -> LockHandler(Order::class, editableFieldNames = setOf("status"))  // status 可编辑（覆盖注解）
       "ADMIN" -> LockHandler(Order::class, editableFieldNames = setOf("status", "createTime"))  // 管理员可编辑更多字段
   }
   ```

3. **优势**：
   - ✅ 模型层定义常规规则（大部分情况的可编辑性）
   - ✅ 业务层根据角色覆盖特殊需求（如管理员强制编辑某些字段）
   - ✅ 代码清晰，职责分离
   - ✅ 支持灵活的权限控制（参数可以覆盖注解限制）

