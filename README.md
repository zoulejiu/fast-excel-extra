# Excel Extra

`excel-extra` 是一个基于 **FastExcel + Apache POI** 的 Kotlin 扩展工具，主要用于：

- **Excel 导出带下拉框的数据**
- **Excel 列头批注提示**
- 支持静态和动态下拉数据源
- 支持自定义列头顺序（基于 `@ExcelProperty.index`）
## Maven引用
```xml

<dependency>
    <groupId>io.github.zoulejiu</groupId>
    <artifactId>fast-excel_extra</artifactId>
    <version>1.0</version>
</dependency>
```

---

## 特性

1. **下拉框**：通过 `@ExcelSelect` 注解设置字段下拉，支持静态数组或动态 Map 数据源。
2. **批注提示**：通过 `@ExcelComment` 注解在列头生成批注，方便前端用户填写。
3. **自动列顺序**：默认按字段声明顺序生成列，`@ExcelProperty.index` 可覆盖顺序。
4. **兼容 FastExcel**：直接注册 `UniversalDropdownHandler` 即可导出。

---

## Maven / Gradle 依赖

```xml
<dependency>
    <groupId>cn.idev.excel</groupId>
    <artifactId>fastexcel</artifactId>
    <version>1.3.0</version>
</dependency>
```
### 使用示例
```kotlin
class User {
    @field:ExcelProperty(value = ["序号"], index = 0)
    var id: Int? = null

    @field:ExcelProperty(value = ["名称"], index = 1)
    @field:ExcelSelect(key = "name")
    @field:ExcelComment("请选择部门名称")
    var name: String? = null

    var dept: String? = null

    @field:ExcelProperty(value = ["性别"], index = 5)
    @field:ExcelSelect(options = ["男","女"])
    @field:ExcelComment("只能选择 男 或 女")
    var gender: String? = null
}

fun main() {
    val users = listOf(
        User().apply { id = 1; name = "Alice"; dept = "研发部"; gender = "女" },
        User().apply { id = 2; name = "Bob"; dept = "销售部"; gender = "男" },
        User().apply { id = 3; name = "Charlie"; dept = "产品部"; gender = "女" }
    )

    val dynamicOptions = mapOf(
        "name" to arrayOf("研发部", "销售部", "产品部")
    )

    FastExcel.write("users.xlsx", User::class.java)
        .registerWriteHandler(UniversalDropdownHandler(User::class, dynamicOptions))
        .sheet()
        .doWrite(users)
}

```
