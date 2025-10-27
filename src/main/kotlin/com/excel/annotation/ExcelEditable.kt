package com.excel.annotation
/**
 * 控制 Excel 单元格是否可编辑
 */
@Target(AnnotationTarget.FIELD)
@Retention(AnnotationRetention.RUNTIME)
annotation class ExcelEditable(
    val value: Boolean = false // 默认不可编辑
)
