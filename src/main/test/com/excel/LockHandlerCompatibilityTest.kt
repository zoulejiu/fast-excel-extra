package com.excel

import cn.idev.excel.FastExcel
import cn.idev.excel.annotation.ExcelProperty
import com.excel.annotation.ExcelEditable
import com.excel.handler.LockHandler

/**
 * LockHandler 注解与参数兼容性测试
 * 演示注解和参数两种方式如何配合使用
 */
fun main() {
    // 准备测试数据
    val products = listOf(
        Product(1L, "笔记本电脑", 5999.99, 100, "这是一款高性能笔记本"),
        Product(2L, "机械键盘", 599.99, 200, "RGB机械键盘"),
        Product(3L, "显示器", 1999.99, 50, "27寸4K显示器")
    )

    println("=".repeat(80))
    println("LockHandler 注解与参数兼容性测试")
    println("=".repeat(80))

    // ========== 场景1: 只使用注解（不传参数）==========
    println("\n【场景1】只使用注解")
    println("- 注解配置：id(不可编辑), name(可编辑), price(可编辑), stock(不可编辑), description(可编辑)")
    println("- 参数配置：无")
    println("- 预期结果：name、price、description 可编辑；id、stock 不可编辑")
    
    FastExcel.write("D:/test/compatibility_test1_annotation_only.xlsx")
        .registerWriteHandler(LockHandler(Product::class))  // 不传 editableFieldNames
        .head(Product::class.java)
        .sheet("商品列表")
        .doWrite(products)

    // ========== 场景2: 只使用参数（忽略注解）==========
    println("\n【场景2】只使用参数（数据类无注解的情况）")
    println("- 注解配置：无（使用 ProductNoAnnotation 类）")
    println("- 参数配置：只允许 name 和 price")
    println("- 预期结果：name、price 可编辑；其他字段不可编辑")
    
    val productsNoAnn = listOf(
        ProductNoAnnotation(1L, "笔记本电脑", 5999.99, 100, "这是一款高性能笔记本"),
        ProductNoAnnotation(2L, "机械键盘", 599.99, 200, "RGB机械键盘")
    )
    
    FastExcel.write("D:/test/compatibility_test2_parameter_only.xlsx")
        .registerWriteHandler(
            LockHandler(
                ProductNoAnnotation::class,
                editableFieldNames = setOf("name", "price")
            )
        )
        .head(ProductNoAnnotation::class.java)
        .sheet("商品列表")
        .doWrite(productsNoAnn)

    // ========== 场景3: 注解与参数同时使用（参数优先 + OR 关系）==========
    println("\n【场景3】注解与参数同时使用（参数优先 + OR 关系）")
    println("- 注解配置：id(不可编辑), name(可编辑), price(可编辑), stock(不可编辑), description(可编辑)")
    println("- 参数配置：name、price、stock")
    println("- 预期结果：")
    println("  • id: 不在参数中 → 按注解 = 不可编辑 ❌")
    println("  • name: 在参数中 → 强制可编辑 = 可编辑 ✅")
    println("  • price: 在参数中 → 强制可编辑 = 可编辑 ✅")
    println("  • stock: 在参数中 → 强制可编辑（覆盖注解限制）= 可编辑 ✅")
    println("  • description: 不在参数中 → 按注解 = 可编辑 ✅")
    
    FastExcel.write("D:/test/compatibility_test3_annotation_and_parameter.xlsx")
        .registerWriteHandler(
            LockHandler(
                Product::class,
                editableFieldNames = setOf("name", "price", "stock")  // stock 虽然在列表中，但被注解禁止
            )
        )
        .head(Product::class.java)
        .sheet("商品列表")
        .doWrite(products)

    // ========== 场景4: 参数覆盖注解（强制某些字段可编辑）==========
    println("\n【场景4】参数覆盖注解（强制某些字段可编辑）")
    println("- 注解配置：id、stock 不可编辑，其他可编辑")
    println("- 参数配置：只允许 id、stock（覆盖注解）")
    println("- 预期结果：")
    println("  • id: 在参数中 → 强制可编辑（覆盖注解）✅")
    println("  • stock: 在参数中 → 强制可编辑（覆盖注解）✅")
    println("  • name、price、description: 不在参数中 → 按注解 = 可编辑 ✅")
    
    FastExcel.write("D:/test/compatibility_test4_parameter_override.xlsx")
        .registerWriteHandler(
            LockHandler(
                Product::class,
                editableFieldNames = setOf("id", "stock")  // 强制 id 和 stock 可编辑（覆盖注解）
            )
        )
        .head(Product::class.java)
        .sheet("商品列表")
        .doWrite(products)

    // ========== 场景5: 空参数（只按注解配置）==========
    println("\n【场景5】空参数（只按注解配置）")
    println("- 注解配置：id、stock 不可编辑，name、price、description 可编辑")
    println("- 参数配置：空集合 emptySet()")
    println("- 预期结果：")
    println("  • 所有字段都不在参数中 → 按注解配置")
    println("  • id、stock: 不可编辑 ❌")
    println("  • name、price、description: 可编辑 ✅")
    
    FastExcel.write("D:/test/compatibility_test5_empty_parameter.xlsx")
        .registerWriteHandler(
            LockHandler(
                Product::class,
                editableFieldNames = emptySet()  // 空集合
            )
        )
        .head(Product::class.java)
        .sheet("商品列表")
        .doWrite(products)

    println("\n" + "=".repeat(80))
    println("✅ 所有测试文件已生成到 D:/test/ 目录")
    println("=".repeat(80))
    println("\n📋 测试文件列表:")
    println("1️⃣  compatibility_test1_annotation_only.xlsx")
    println("    → 只使用注解：id、stock 不可编辑，其他可编辑")
    println()
    println("2️⃣  compatibility_test2_parameter_only.xlsx")
    println("    → 只使用参数：name、price 可编辑，其他按默认规则")
    println()
    println("3️⃣  compatibility_test3_annotation_and_parameter.xlsx")
    println("    → 注解+参数（参数优先）：name、price、stock 可编辑，description 可编辑，id 不可编辑")
    println()
    println("4️⃣  compatibility_test4_parameter_override.xlsx")
    println("    → 参数覆盖注解：id、stock 强制可编辑（覆盖注解限制），其他按注解")
    println()
    println("5️⃣  compatibility_test5_empty_parameter.xlsx")
    println("    → 空参数：按注解配置，id、stock 不可编辑，其他可编辑")
    println("\n💡 提示：打开 Excel 文件验证可编辑字段是否符合预期")
}

/**
 * 带注解的商品数据类
 */
data class Product(
    @ExcelProperty("商品ID")
    @ExcelEditable(false)  // 不可编辑
    val id: Long,

    @ExcelProperty("商品名称")
    @ExcelEditable(true)  // 可编辑
    val name: String,

    @ExcelProperty("价格")
    @ExcelEditable(true)  // 可编辑
    val price: Double,

    @ExcelProperty("库存")
    @ExcelEditable(false)  // 不可编辑
    val stock: Int,

    @ExcelProperty("描述")
    // 没有注解，默认可编辑
    val description: String
)

/**
 * 无注解的商品数据类（用于测试纯参数模式）
 */
data class ProductNoAnnotation(
    @ExcelProperty("商品ID")
    val id: Long,

    @ExcelProperty("商品名称")
    val name: String,

    @ExcelProperty("价格")
    val price: Double,

    @ExcelProperty("库存")
    val stock: Int,

    @ExcelProperty("描述")
    val description: String
)

