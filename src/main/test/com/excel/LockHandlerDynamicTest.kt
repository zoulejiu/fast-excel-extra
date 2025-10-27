package com.excel

import cn.idev.excel.FastExcel
import cn.idev.excel.annotation.ExcelProperty
import com.excel.handler.LockHandler

/**
 * LockHandler 动态指定可编辑字段测试
 * 演示如何在导出时动态控制哪些列可以编辑
 */
fun main() {
    // 准备测试数据
    val users = listOf(
        User(1L, "张三", 25, "研发部"),
        User(2L, "李四", 30, "市场部"),
        User(3L, "王五", 28, "人事部")
    )

    // ========== 场景1: 只允许编辑"姓名"和"年龄"字段 ==========
    println("生成测试文件1: 只允许编辑姓名和年龄...")
    FastExcel.write("D:/test/lock_test1_editable_name_age.xlsx")
        .registerWriteHandler(
            LockHandler(
                User::class,
                editableFieldNames = setOf("name", "age")  // 只有name和age可编辑
            )
        )
        .head(User::class.java)
        .sheet("用户列表")
        .doWrite(users)

    // ========== 场景2: 只允许编辑"部门"字段，其他字段都锁定 ==========
    println("生成测试文件2: 只允许编辑部门...")
    FastExcel.write("D:/test/lock_test2_editable_department.xlsx")
        .registerWriteHandler(
            LockHandler(
                User::class,
                editableFieldNames = setOf("department")  // 只有department可编辑
            )
        )
        .head(User::class.java)
        .sheet("用户列表")
        .doWrite(users)

    // ========== 场景3: 所有字段都不可编辑（完全锁定） ==========
    println("生成测试文件3: 所有字段都不可编辑...")
    FastExcel.write("D:/test/lock_test3_all_locked.xlsx")
        .registerWriteHandler(
            LockHandler(
                User::class,
                editableFieldNames = emptySet()  // 空集合表示所有字段都锁定
            )
        )
        .head(User::class.java)
        .sheet("用户列表")
        .doWrite(users)

    // ========== 场景4: 带密码保护，只允许编辑"年龄" ==========
    println("生成测试文件4: 带密码保护，只允许编辑年龄...")
    FastExcel.write("D:/test/lock_test4_password_protected.xlsx")
        .registerWriteHandler(
            LockHandler(
                User::class,
                editableFieldNames = setOf("age"),
                protectPassword = "123456"  // 设置密码
            )
        )
        .head(User::class.java)
        .sheet("用户列表")
        .doWrite(users)

    // ========== 场景5: 不指定可编辑字段，回退到注解方式（默认所有字段可编辑） ==========
    println("生成测试文件5: 不指定可编辑字段（所有字段默认可编辑）...")
    FastExcel.write("D:/test/lock_test5_default_all_editable.xlsx")
        .registerWriteHandler(
            LockHandler(User::class)  // 不指定 editableFieldNames
        )
        .head(User::class.java)
        .sheet("用户列表")
        .doWrite(users)

    println("\n✅ 所有测试文件已生成到 D:/test/ 目录")
    println("\n📋 测试文件说明:")
    println("1️⃣  lock_test1_editable_name_age.xlsx - 只能编辑：姓名、年龄")
    println("2️⃣  lock_test2_editable_department.xlsx - 只能编辑：部门")
    println("3️⃣  lock_test3_all_locked.xlsx - 所有字段都不可编辑")
    println("4️⃣  lock_test4_password_protected.xlsx - 只能编辑：年龄（密码：123456）")
    println("5️⃣  lock_test5_default_all_editable.xlsx - 所有字段默认可编辑（除表头）")
    println("\n💡 提示：打开 Excel 文件后，尝试编辑各个单元格，验证保护机制是否生效")
    println("💡 提示：可以调整列宽和行高，不会受到保护限制")
}

/**
 * 测试用户数据类
 */
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

