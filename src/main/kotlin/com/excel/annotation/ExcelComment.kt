package com.excel.annotation

/**
 * 批注注解
 */
@Target(AnnotationTarget.FIELD)
@Retention(AnnotationRetention.RUNTIME)
annotation class ExcelComment(
    val value: String // 批注内容
)
