# Fast Excel Extra

`fast-excel-extra` 是一个基于 **FastExcel + Apache POI** 的 Kotlin 扩展工具库，提供以下增强功能：

- ✅ **Excel 单元格锁定控制**（🆕 v1.3）
- ✅ **Excel 导出带下拉框的数据**
- ✅ **Excel 列头批注提示**
- ✅ 支持静态和动态下拉数据源
- ✅ 支持自定义列头顺序（基于 `@ExcelProperty.index`）
- ✅ 完美兼容 `excludeColumnFieldNames`（🆕 v1.3）

## Maven 引用

```xml
<dependency>
    <groupId>io.github.zoulejiu</groupId>
    <artifactId>fast-excel-extra</artifactId>
    <version>1.3</version>
</dependency>
```

---

## 核心特性

### 1. 单元格锁定控制 🆕

通过 `LockHandler` 和 `@ExcelEditable` 注解，精确控制 Excel 单元格的可编辑性：

- 🔒 保护关键字段（如 ID、系统字段）不被修改
- ✏️ 允许用户编辑指定字段
- 🔑 支持密码保护
- 📏 允许调整列宽和行高
- 🎯 支持动态权限控制（根据用户角色）
- 🔄 完美兼容 `excludeColumnFieldNames`

### 2. 数据验证（下拉框）

通过 `@ExcelSelect` 注解为字段添加下拉框验证：

- 📋 静态下拉选项（直接在注解中指定）
- 🔄 动态下拉选项（通过参数传入）
- ✅ 自动数据验证

### 3. 列头批注提示

通过 `@ExcelComment` 注解在列头添加批注说明：

- 💡 帮助用户理解字段含义
- 📝 提供填写指引

### 4. 灵活的列控制

- 🎯 支持 `@ExcelProperty.index` 自定义列顺序
- 🚫 完美兼容 `excludeColumnFieldNames` 排除列
- 🔄 自动处理列索引映射

---

## 快速开始

### 依赖配置

```xml
<!-- FastExcel -->
<dependency>
    <groupId>cn.idev.excel</groupId>
    <artifactId>fastexcel</artifactId>
    <version>1.3.0</version>
</dependency>

<!-- Fast Excel Extra -->
<dependency>
    <groupId>io.github.zoulejiu</groupId>
    <artifactId>fast-excel-extra</artifactId>
    <version>1.3</version>
</dependency>
```

---

## 使用示例

### 示例 1：单元格锁定控制 🆕

#### 场景：绩效考核分阶段导出

```kotlin
import com.excel.annotation.ExcelEditable
import com.excel.handler.LockHandler

data class PerformanceItem(
    @ExcelProperty("员工ID")
    @ExcelEditable(false)  // 不可编辑
    val id: Long,

    @ExcelProperty("姓名")
    @ExcelEditable(false)  // 不可编辑
    val name: String,

    @ExcelProperty("自评分数")
    @ExcelEditable(true)  // 可编辑
    val selfRating: Double,

    @ExcelProperty("自评说明")
    @ExcelEditable(true)  // 可编辑
    val selfComment: String,

    @ExcelProperty("评审分数")
    @ExcelEditable(false)  // 默认不可编辑
    val auditRating: Double,

    @ExcelProperty("评审说明")
    @ExcelEditable(false)  // 默认不可编辑
    val auditComment: String
)

// 草稿阶段：员工只能编辑自评字段
FastExcel.write("performance_draft.xlsx")
    .registerWriteHandler(LockHandler(PerformanceItem::class))
    .head(PerformanceItem::class.java)
    .sheet("绩效考核")
    .doWrite(dataList)

// 审核阶段：领导可以编辑审核字段（覆盖注解限制）
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

// 带密码保护
FastExcel.write("performance_protected.xlsx")
    .registerWriteHandler(
        LockHandler(
            PerformanceItem::class,
            editableFieldNames = setOf("selfRating", "selfComment"),
            protectPassword = "123456"
        )
    )
    .head(PerformanceItem::class.java)
    .sheet("绩效考核")
    .doWrite(dataList)
```

#### 配合 excludeColumnFieldNames 使用

```kotlin
// 草稿阶段：隐藏审核字段，只显示自评字段
FastExcel.write("draft.xlsx")
    .registerWriteHandler(
        LockHandler(
            PerformanceItem::class,
            editableFieldNames = setOf("selfRating", "selfComment")
        )
    )
    .head(PerformanceItem::class.java)
    .excludeColumnFieldNames(listOf("auditRating", "auditComment"))  // 排除审核字段
    .sheet("绩效考核")
    .doWrite(dataList)
```

### 示例 2：下拉框和批注

```kotlin
import com.excel.annotation.ExcelSelect
import com.excel.annotation.ExcelComment
import com.excel.handler.UniversalDropdownHandler

data class User(
    @ExcelProperty("序号")
    val id: Int,

    @ExcelProperty("部门")
    @ExcelSelect(key = "dept")  // 动态下拉选项
    @ExcelComment("请从下拉列表中选择部门")
    val dept: String,

    @ExcelProperty("姓名")
    val name: String,

    @ExcelProperty("性别")
    @ExcelSelect(options = ["男", "女"])  // 静态下拉选项
    @ExcelComment("只能选择 男 或 女")
    val gender: String
)

fun main() {
    val users = listOf(
        User(1, "研发部", "张三", "男"),
        User(2, "销售部", "李四", "女")
    )

    // 动态下拉选项（运行时传入）
    val dynamicOptions = mapOf(
        "dept" to arrayOf("研发部", "销售部", "产品部", "人事部")
    )

    FastExcel.write("users.xlsx")
        .registerWriteHandler(UniversalDropdownHandler(User::class, dynamicOptions))
        .head(User::class.java)
        .sheet("用户列表")
        .doWrite(users)
}
```

### 示例 3：组合使用

```kotlin
// 同时使用单元格锁定 + 下拉框 + 批注
data class Product(
    @ExcelProperty("产品ID")
    @ExcelEditable(false)  // ID 不可编辑
    val id: Long,

    @ExcelProperty("产品名称")
    @ExcelEditable(true)  // 名称可编辑
    val name: String,

    @ExcelProperty("分类")
    @ExcelSelect(options = ["电子产品", "家居用品", "食品饮料"])
    @ExcelComment("请选择产品分类")
    @ExcelEditable(true)  // 分类可编辑
    val category: String,

    @ExcelProperty("价格")
    @ExcelComment("请输入正数")
    @ExcelEditable(true)  // 价格可编辑
    val price: Double
)

FastExcel.write("products.xlsx")
    .registerWriteHandler(UniversalDropdownHandler(Product::class))  // 下拉框和批注
    .registerWriteHandler(LockHandler(Product::class))  // 单元格锁定
    .head(Product::class.java)
    .sheet("产品列表")
    .doWrite(productList)
```

---

## API 说明

### LockHandler

用于控制 Excel 单元格的可编辑性。

**构造参数：**

| 参数 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `dataClass` | `KClass<T>` | - | 数据类的 KClass（必填） |
| `editableFieldNames` | `Set<String>?` | `null` | 可编辑的字段名集合（可选）。如果指定，参数中的字段强制可编辑（覆盖注解），其他字段按注解配置 |
| `enableProtection` | `Boolean` | `true` | 是否启用工作表保护 |
| `protectPassword` | `String?` | `null` | 工作表保护密码 |

**注解：`@ExcelEditable`**

| 参数 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `value` | `Boolean` | `false` | `true` 表示可编辑，`false` 表示不可编辑（默认不可编辑） |

**规则：**
- 参数优先级高于注解
- 参数中的字段强制可编辑（覆盖注解限制）
- 参数之外的字段按注解配置
- 表头行始终被锁定

### UniversalDropdownHandler

用于添加下拉框验证和列头批注。

**构造参数：**

| 参数 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `dataClass` | `KClass<T>` | - | 数据类的 KClass（必填） |
| `dynamicOptions` | `Map<String, Array<String>>` | `emptyMap()` | 动态下拉选项（key 对应 `@ExcelSelect.key`） |
| `lastRow` | `Int` | `200` | 下拉框应用到第几行 |

**注解：`@ExcelSelect`**

| 参数 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `options` | `Array<String>` | `[]` | 静态下拉选项（直接指定） |
| `key` | `String` | `""` | 动态下拉选项的 key（从 `dynamicOptions` 获取） |

**注解：`@ExcelComment`**

| 参数 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `value` | `String` | - | 批注内容 |

---

## 高级特性

### 兼容 excludeColumnFieldNames

v1.3 完美支持 `excludeColumnFieldNames`，自动处理列索引映射：

```kotlin
FastExcel.write("output.xlsx")
    .registerWriteHandler(LockHandler(Model::class))
    .registerWriteHandler(UniversalDropdownHandler(Model::class))
    .head(Model::class.java)
    .excludeColumnFieldNames(listOf("id", "createTime"))  // ✅ 自动处理索引
    .sheet()
    .doWrite(dataList)
```

### 动态权限控制

根据用户角色动态控制可编辑字段：

```kotlin
val handler = when (userRole) {
    "USER" -> LockHandler(Order::class)  // 普通用户：按注解
    "MANAGER" -> LockHandler(Order::class, editableFieldNames = setOf("status"))  // 经理：可编辑状态
    "ADMIN" -> LockHandler(Order::class, editableFieldNames = setOf("status", "amount"))  // 管理员：可编辑更多
    else -> LockHandler(Order::class, editableFieldNames = emptySet())  // 其他：完全锁定
}

FastExcel.write("order.xlsx")
    .registerWriteHandler(handler)
    .head(Order::class.java)
    .sheet()
    .doWrite(orderList)
```

---

## 文档

- 📖 [LockHandler 详细使用指南](docs/LockHandler使用指南.md)

---

## 更新日志

### v1.3 (2025-10-27)

- 🆕 新增 `LockHandler` 单元格锁定功能
- 🆕 新增 `@ExcelEditable` 注解控制可编辑性
- ✨ 支持动态指定可编辑字段
- ✨ 支持密码保护
- ✨ 允许调整列宽和行高
- 🐛 修复 `excludeColumnFieldNames` 导致的列索引错位问题
- 🔧 改进 `ExcelFieldUtils.resolveExcelColumns()` 支持排除列

### v1.1

- ✨ 初始版本
- 📋 下拉框验证支持
- 💡 列头批注支持

---

## License

MIT License
