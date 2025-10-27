package com.excel

import cn.idev.excel.FastExcel
import com.excel.handler.LockHandler
import com.excel.model.PerformanceSettingAnnualAssessmentItem
import java.math.BigDecimal

fun main() {
    // First instance
    val setting1 = PerformanceSettingAnnualAssessmentItem().apply {
        performanceSettingItemId = 1L
        id = 2L
        category = "Sales"
        indicatorType = "Revenue"
        indicatorDetail = "Total sales revenue for the year"
        scoreRule = "Score based on percentage of target achieved"
        difficulty = "Medium"
        cycle = "Annual"
        type = 1
        weight = "30%"
        selfRating = BigDecimal("85.5")
        selfComment = "Exceeded target by 5%"
        auditRating = BigDecimal("87.0")
        auditComment = "Good performance, slightly above expectations"
        finalRating = BigDecimal("86.25")
        finalRatingComment = "Final rating is an average of self and audit ratings"
    }

    // Second instance
    val setting2 = PerformanceSettingAnnualAssessmentItem().apply {
        performanceSettingItemId = 3L
        id = 4L
        category = "Marketing"
        indicatorType = "Campaign Effectiveness"
        indicatorDetail = "Number of successful marketing campaigns"
        scoreRule = "Score based on number of successful campaigns"
        difficulty = "High"
        cycle = "Quarterly"
        type = 2
        weight = "20%"
        selfRating = BigDecimal("78.0")
        selfComment = "Met most of the campaign goals"
        auditRating = BigDecimal("80.0")
        auditComment = "Slightly below expectations but acceptable"
        finalRating = BigDecimal("79.0")
        finalRatingComment = "Final rating is an average of self and audit ratings"
    }

    // Third instance
    val setting3 = PerformanceSettingAnnualAssessmentItem().apply {
        performanceSettingItemId = 5L
        id = 6L
        category = "Customer Service"
        indicatorType = "Customer Satisfaction"
        indicatorDetail = "Customer satisfaction survey results"
        scoreRule = "Score based on customer satisfaction index"
        difficulty = "Low"
        cycle = "Monthly"
        type = 3
        weight = "10%"
        selfRating = BigDecimal("92.0")
        selfComment = "High customer satisfaction levels"
        auditRating = BigDecimal("93.0")
        auditComment = "Excellent performance, above expectations"
        finalRating = BigDecimal("92.5")
        finalRatingComment = "Final rating is an average of self and audit ratings"
    }
    // ========== 方式1: 使用注解方式（需要在 model 类中添加 @ExcelEditable 注解）==========
    // 模拟草稿下导出 - 使用注解控制可编辑字段
//    FastExcel.write("D:/test/dynamicHeader0_annotation.xlsx")
//        .registerWriteHandler(LockHandler(PerformanceSettingAnnualAssessmentItem::class,
//            setOf("selfRating","selfComment")))
//        .head(PerformanceSettingAnnualAssessmentItem::class.java)
//        .excludeColumnFieldNames(listOf("id","auditRating","auditComment","finalRating","finalRatingComment"))
//        .sheet().doWrite(listOf(setting1, setting2, setting3))

    // ========== 方式2: 动态指定可编辑字段（推荐，无需注解）==========
    // 模拟草稿下导出 - 只允许编辑"自评分数"和"自评说明"
    FastExcel.write("D:/test/dynamicHeader0_dynamic.xlsx")
        .registerWriteHandler(LockHandler(
            PerformanceSettingAnnualAssessmentItem::class,
            editableFieldNames = setOf("selfRating", "selfComment")
        ))
        .head(PerformanceSettingAnnualAssessmentItem::class.java)
        .excludeColumnFieldNames(listOf("id","auditRating","auditComment","finalRating","finalRatingComment"))
        .sheet().doWrite(listOf(setting1, setting2, setting3))

    // 模拟直接上级领导导出 - 只允许编辑"评审分数"和"评审说明"
    FastExcel.write("D:/test/dynamicHeader1_dynamic.xlsx")
        .registerWriteHandler(LockHandler(
            PerformanceSettingAnnualAssessmentItem::class,
            editableFieldNames = setOf("auditRating", "auditComment")
        ))
        .head(PerformanceSettingAnnualAssessmentItem::class.java)
        .excludeColumnFieldNames(listOf("finalRating","finalRatingComment"))
        .sheet().doWrite(listOf(setting1, setting2, setting3))

    // 模拟部门领导导出 - 只允许编辑"最终分数"和"最终评价"
    FastExcel.write("D:/test/dynamicHeader2_dynamic.xlsx")
        .registerWriteHandler(LockHandler(
            PerformanceSettingAnnualAssessmentItem::class,
            editableFieldNames = setOf("finalRating", "finalRatingComment"),
            protectPassword = "admin123"  // 设置密码保护
        ))
        .head(PerformanceSettingAnnualAssessmentItem::class.java)
        .excludeColumnFieldNames(listOf())
        .sheet().doWrite(listOf(setting1, setting2, setting3))

    println("✅ Excel 文件已生成到 D:/test/ 目录")
    println("📁 dynamicHeader0_dynamic.xlsx - 草稿模式（可编辑：自评分数、自评说明）")
    println("📁 dynamicHeader1_dynamic.xlsx - 直接上级模式（可编辑：评审分数、评审说明）")
    println("📁 dynamicHeader2_dynamic.xlsx - 部门领导模式（可编辑：最终分数、最终评价，密码：admin123）")

}
