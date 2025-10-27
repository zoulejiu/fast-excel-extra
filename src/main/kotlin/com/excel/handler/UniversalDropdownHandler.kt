package com.excel.handler

import cn.idev.excel.write.handler.SheetWriteHandler
import cn.idev.excel.write.handler.context.SheetWriteHandlerContext
import com.excel.annotation.ExcelComment
import com.excel.annotation.ExcelSelect
import com.excel.util.ExcelFieldUtils
import org.apache.poi.ss.usermodel.DataValidation
import org.apache.poi.ss.usermodel.DataValidationConstraint
import org.apache.poi.ss.util.CellRangeAddressList
import kotlin.reflect.KClass

/**
 * 通用的数据验证支持数据验证和批注
 */
class UniversalDropdownHandler<T : Any>(
    private val dataClass: KClass<T>,
    private val dynamicOptions: Map<String, Array<String>> = emptyMap(),
    private val lastRow: Int = 200
) : SheetWriteHandler {

    override fun afterSheetCreate(context: SheetWriteHandlerContext) {
        val sheet = context.writeSheetHolder.sheet
        val helper = sheet.dataValidationHelper

        // 从 context 中获取排除的列
        val excludedFields = try {
            context.writeSheetHolder?.excludeColumnFieldNames() ?: emptySet()
        } catch (e: Exception) {
            emptySet()
        }

        // 计算字段与列的对应关系（兼容 @ExcelProperty.index、源码顺序和 excludeColumnFieldNames）
        val columnMetas = ExcelFieldUtils.resolveExcelColumns(dataClass, excludedFields)

        for (meta in columnMetas) {
            val field = dataClass.java.getDeclaredField(meta.fieldName)
            field.getAnnotation(ExcelSelect::class.java)?.let { selectAnn ->
                val options = if (selectAnn.options.isNotEmpty()) selectAnn.options
                else dynamicOptions[selectAnn.key] ?: emptyArray()
                if (options.isNotEmpty()) {
                    val range = CellRangeAddressList(1, lastRow, meta.index, meta.index)
                    val constraint: DataValidationConstraint =
                        helper.createExplicitListConstraint(options)
                    val validation: DataValidation = helper.createValidation(constraint, range)
                    validation.showErrorBox = true
                    sheet.addValidationData(validation)

                    println("index=${meta.index}, field=${meta.fieldName}, header=${meta.headerName}, options=${options.contentToString()}")
                }
            }
            field.getAnnotation(ExcelComment::class.java)?.let { commentAnn ->
                val row = sheet.getRow(0) ?: sheet.createRow(0)
                val cell = row.getCell(meta.index) ?: row.createCell(meta.index)
                val drawing = sheet.createDrawingPatriarch()
                val creationHelper = sheet.workbook.creationHelper
                val anchor = creationHelper.createClientAnchor().apply {
                    setCol1(meta.index)
                    setCol2(meta.index + 2)
                    row1 = 0
                    row2 = 3
                }
                val comment = drawing.createCellComment(anchor)
                comment.string = sheet.workbook.creationHelper.createRichTextString(commentAnn.value)
                cell.cellComment = comment

                println("批注 index=${meta.index}, field=${meta.fieldName}, comment=${commentAnn.value}")
            }

        }
    }
}
