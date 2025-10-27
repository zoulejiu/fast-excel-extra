package com.excel.util

import cn.idev.excel.annotation.ExcelIgnore
import cn.idev.excel.annotation.ExcelProperty
import kotlin.reflect.KClass

/**
 * Excel 字段工具类
 */
object ExcelFieldUtils {

    /**
     * Excel 列元数据
     */
    data class ExcelColumnMeta(
        val fieldName: String,
        val headerName: String,
        val index: Int
    )

    /**
     * 解析 Excel 列配置（兼容 @ExcelProperty.index 和源码顺序）
     *
     * @param clazz 数据类的 KClass
     * @param excludeColumnFieldNames 要排除的字段名集合（兼容 EasyExcel 的 excludeColumnFieldNames）
     * @return 按列索引排序的列元数据列表
     */
    fun <T : Any> resolveExcelColumns(
        clazz: KClass<T>,
        excludeColumnFieldNames: Collection<String> = emptyList()
    ): List<ExcelColumnMeta> {
        val fields = clazz.java.declaredFields
        var autoIndex = 0

        return fields
            .filter { it.getAnnotation(ExcelIgnore::class.java) == null }  // 排除 @ExcelIgnore
            .filter { it.name !in excludeColumnFieldNames }  // 排除 excludeColumnFieldNames 中的字段
            .map { field ->
                val ann = field.getAnnotation(ExcelProperty::class.java)
                val header = ann?.value?.firstOrNull() ?: field.name
                val annIndex = ann?.index ?: -1

                val finalIndex = if (annIndex != -1) {
                    annIndex
                } else {
                    // 自动分配 index（跳过已有的 index）
                    while (true) {
                        if (fields.none {
                                it.getAnnotation(ExcelProperty::class.java)?.index == autoIndex
                            }) {
                            break
                        }
                        autoIndex++
                    }
                    autoIndex++
                    autoIndex - 1
                }

                ExcelColumnMeta(
                    fieldName = field.name,
                    headerName = header,
                    index = finalIndex
                )
            }
            .sortedBy { it.index }
            .mapIndexed { actualIndex, meta ->
                // 重新计算实际的列索引（考虑排除列后的连续索引）
                meta.copy(index = actualIndex)
            }
    }
}

