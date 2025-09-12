package com.excel.annotation

/**
 * 数据验证下拉框
 */
@Target(AnnotationTarget.FIELD)
@Retention(AnnotationRetention.RUNTIME)
annotation class ExcelSelect(
    val options: Array<String> = [], // 静态下拉选项
    val key: String = ""             // 外部下拉选项标识
)
