package com.excel.handler

import cn.idev.excel.annotation.ExcelIgnore
import cn.idev.excel.annotation.ExcelProperty
import cn.idev.excel.write.handler.CellWriteHandler
import cn.idev.excel.write.handler.WorkbookWriteHandler
import cn.idev.excel.write.handler.context.CellWriteHandlerContext
import cn.idev.excel.write.handler.context.WorkbookWriteHandlerContext
import com.excel.annotation.ExcelEditable
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.xssf.streaming.SXSSFSheet
import org.apache.poi.xssf.usermodel.XSSFSheet
import kotlin.reflect.KClass

/**
 * Excel 单元格锁定处理器
 *
 * 功能说明：
 * 1. 支持两种方式控制列是否可编辑：
 *    a) 通过 @ExcelEditable 注解（基础规则）
 *    b) 通过 editableFieldNames 参数动态指定（覆盖规则）
 * 2. 两种方式可以兼容使用（OR 关系 + 参数优先）：
 *    - 如果只使用注解：按注解配置生效
 *    - 如果只使用参数：参数中的字段可编辑，其他字段按默认规则（可编辑）
 *    - 如果同时使用：参数中的字段强制可编辑（覆盖注解），参数之外的字段按注解配置
 * 3. 在所有数据写入完成后，自动保护工作表
 *
 * 使用示例：
 * ```kotlin
 * data class User(
 *     @ExcelProperty("ID")
 *     val id: Long,
 *
 *     @ExcelProperty("姓名")
 *     val name: String,
 *
 *     @ExcelProperty("年龄")
 *     val age: Int
 * )
 *
 * // 方式1: 使用注解（字段需要标记 @ExcelEditable）
 * EasyExcel.write(file, User::class.java)
 *     .registerWriteHandler(LockHandler(User::class))
 *     .sheet("用户列表")
 *     .doWrite(dataList)
 *
 * // 方式2: 动态指定可编辑字段（推荐）
 * EasyExcel.write(file, User::class.java)
 *     .registerWriteHandler(LockHandler(User::class, editableFieldNames = setOf("name", "age")))
 *     .sheet("用户列表")
 *     .doWrite(dataList)
 *
 * // 方式3: 带密码保护
 * EasyExcel.write(file, User::class.java)
 *     .registerWriteHandler(LockHandler(User::class, editableFieldNames = setOf("name"), protectPassword = "123456"))
 *     .sheet("用户列表")
 *     .doWrite(dataList)
 * ```
 *
 * @param T 数据类型
 * @param dataClass 数据类的 KClass
 * @param editableFieldNames 可编辑的字段名集合（可选）。
 *                           - 如果不指定（null），则完全按照注解配置生效
 *                           - 如果指定，则作为"覆盖规则"：
 *                             • 在参数列表中的字段 → 强制可编辑（覆盖注解的限制）
 *                             • 不在参数列表中的字段 → 按注解配置
 *                           - 优先级：参数 > 注解
 * @param enableProtection 是否启用工作表保护，默认 true
 * @param protectPassword 工作表保护密码（可选）
 */
class LockHandler<T : Any>(
    private val dataClass: KClass<T>,
    private val editableFieldNames: Set<String>? = null,  // 可编辑字段名集合
    private val enableProtection: Boolean = true,  // 默认启用保护
    private val protectPassword: String? = null
) : CellWriteHandler, WorkbookWriteHandler {

    /**
     * 字段名 -> 是否可编辑的映射
     * true: 可编辑（unlocked）
     * false: 不可编辑（locked）
     */
    private val fieldEditableMap = mutableMapOf<String, Boolean>()

    /**
     * 列索引 -> 字段名的映射（运行时动态构建，兼容 excludeColumnFieldNames）
     */
    private val columnIndexToFieldNameMap = mutableMapOf<Int, String>()

    /**
     * 标记是否已初始化
     */
    private var initialized = false

    /**
     * 在单元格创建之前初始化字段配置
     */
    override fun beforeCellCreate(context: CellWriteHandlerContext) {
        if (!initialized) {
            resolveEditableFields(dataClass)
            initialized = true
        }
    }

    /**
     * 在单元格处理完成后记录列映射（运行时构建，兼容 excludeColumnFieldNames）
     */
    override fun afterCellDispose(context: CellWriteHandlerContext) {
        val cell = context.cell ?: return
        val colIndex = cell.columnIndex
        val rowIndex = cell.rowIndex

        // 表头行：记录列索引到字段名的映射
        if (context.head && rowIndex == 0) {
            val headerText = cell.stringCellValue
            if (headerText != null && headerText.isNotBlank()) {
                val fieldName = findFieldNameByHeader(headerText)
                if (fieldName != null) {
                    columnIndexToFieldNameMap[colIndex] = fieldName
                }
            }
        }
    }

    /**
     * 在整个 Workbook 写入完成后保护所有工作表
     */
    override fun afterWorkbookDispose(context: WorkbookWriteHandlerContext) {
        // 通过 writeWorkbookHolder 获取 workbook
        val workbook = context.writeWorkbookHolder.cachedWorkbook

        // 保护所有 sheet
        for (i in 0 until workbook.numberOfSheets) {
            val sheet = workbook.getSheetAt(i)

            // 步骤 1：在保护前，确保所有单元格的 locked 属性正确设置
            val styleCache = mutableMapOf<CellStyle, Pair<CellStyle, CellStyle>>()

            for (rowIdx in 0 until sheet.physicalNumberOfRows) {
                val row = sheet.getRow(rowIdx) ?: continue
                for (colIdx in 0 until row.lastCellNum.toInt()) {
                    val cell = row.getCell(colIdx) ?: continue
                    val fieldName = columnIndexToFieldNameMap[colIdx]

                    val shouldLock = if (rowIdx == 0) {
                        // 表头行始终锁定
                        true
                    } else {
                        // 数据行根据字段配置
                        if (fieldName != null) {
                            val isEditable = fieldEditableMap[fieldName] ?: true
                            !isEditable  // isEditable=true -> locked=false
                        } else {
                            false  // 未知字段默认可编辑
                        }
                    }

                    val originalStyle = cell.cellStyle
                    if (originalStyle != null) {
                        // 从缓存中获取或创建对应的样式
                        val (lockedStyleForThis, unlockedStyleForThis) = styleCache.getOrPut(originalStyle) {
                            // 创建 locked 版本
                            val lockedStyle = workbook.createCellStyle()
                            lockedStyle.cloneStyleFrom(originalStyle)
                            lockedStyle.locked = true

                            // 创建 unlocked 版本
                            val unlockedStyle = workbook.createCellStyle()
                            unlockedStyle.cloneStyleFrom(originalStyle)
                            unlockedStyle.locked = false

                            Pair(lockedStyle, unlockedStyle)
                        }

                        // 应用对应的样式
                        cell.cellStyle = if (shouldLock) lockedStyleForThis else unlockedStyleForThis
                    }
                }
            }

            // 步骤 2：启用工作表保护
            if (enableProtection) {
                // 先启用工作表保护
                sheet.protectSheet(protectPassword ?: "")
                
                // 然后设置允许的操作（关键：必须在 protectSheet 之后）
                when (sheet) {
                    is SXSSFSheet -> {
                        // SXSSFSheet 流式写入工作表
                        sheet.lockFormatColumns(false)  // 允许调整列宽
                        sheet.lockFormatRows(false)     // 允许调整行高
                    }
                    is XSSFSheet -> {
                        // XSSFSheet 标准工作表
                        sheet.lockFormatColumns(false)  // 允许调整列宽
                        sheet.lockFormatRows(false)     // 允许调整行高
                    }
                }
            }
        }
    }

    /**
     * 解析哪些字段是可编辑的
     * 兼容注解和参数两种方式（OR 关系 + 参数优先）：
     * 1. 只使用注解：按注解配置
     * 2. 只使用参数：参数中的字段可编辑，其他字段按默认规则
     * 3. 同时使用：参数中的字段强制可编辑（覆盖注解），参数之外的字段按注解配置
     */
    private fun resolveEditableFields(clazz: KClass<T>) {
        val fields = clazz.java.declaredFields

        for (field in fields) {
            // 忽略被 @ExcelIgnore 标记的字段
            if (field.getAnnotation(ExcelIgnore::class.java) != null) {
                continue
            }

            // 步骤1: 根据注解判断字段是否可编辑
            val editableAnn = field.getAnnotation(ExcelEditable::class.java)
            val editableByAnnotation = editableAnn?.value ?: true  // 默认可编辑

            // 步骤2: 判断最终是否可编辑（参数优先 + OR 关系）
            val isEditable = if (editableFieldNames == null) {
                // 没有指定参数，完全按注解
                editableByAnnotation
            } else {
                // 指定了参数：参数中的字段强制可编辑（覆盖注解），其他字段按注解
                if (field.name in editableFieldNames) {
                    true  // 在参数列表中 → 强制可编辑（覆盖注解）
                } else {
                    editableByAnnotation  // 不在参数列表中 → 按注解配置
                }
            }

            fieldEditableMap[field.name] = isEditable
        }
    }

    /**
     * 根据表头文本查找对应的字段名
     */
    private fun findFieldNameByHeader(headerText: String): String? {
        val fields = dataClass.java.declaredFields

        for (field in fields) {
            if (field.getAnnotation(ExcelIgnore::class.java) != null) {
                continue
            }

            val excelProperty = field.getAnnotation(ExcelProperty::class.java)
            val header = excelProperty?.value?.firstOrNull() ?: field.name

            if (header == headerText) {
                return field.name
            }
        }

        return null
    }
}
