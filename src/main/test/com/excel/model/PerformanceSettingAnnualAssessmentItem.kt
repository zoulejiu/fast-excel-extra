package com.excel.model

import cn.idev.excel.annotation.ExcelIgnore
import cn.idev.excel.annotation.ExcelProperty
import com.excel.annotation.ExcelComment
import com.excel.annotation.ExcelEditable
import java.math.BigDecimal

class PerformanceSettingAnnualAssessmentItem {
    @ExcelProperty("明细标识")
    @ExcelEditable
    @ExcelComment("不要修改！！！")
    var performanceSettingItemId: Long? = null
    @ExcelProperty("主键")
    @ExcelEditable
    @ExcelComment("不要修改！！！")
    var id: Long? = null
    @ExcelProperty("指标类型")
    @ExcelEditable
    var category: String? = null
    @ExcelProperty("关键绩效指标")
    @ExcelEditable
    var indicatorType: String? = null
    @ExcelProperty("指标值(完成时限与成果标志)")
    @ExcelEditable
    var indicatorDetail: String? = null
    @ExcelProperty("计分规则")
    @ExcelEditable
    var scoreRule: String? = null
    @ExcelProperty("工作难度系数")
    @ExcelEditable
    var difficulty: String? = null
    @ExcelProperty("分配时点")
    @ExcelEditable
    var cycle: String? = null
    @ExcelIgnore
    var type: Int? = null
    @ExcelProperty("权重")
    @ExcelEditable
    var weight: String? = null
    @ExcelEditable
    @ExcelProperty("自评分数")
    var selfRating: BigDecimal? = null
    @ExcelEditable
    @ExcelProperty("自评描述")
    var selfComment: String? = null
    @ExcelEditable
    @ExcelProperty("审核分数")
    var auditRating: BigDecimal? = null
    @ExcelEditable
    @ExcelProperty("审核描述")
    var auditComment: String? = null

    @ExcelProperty("审定分数")
    @ExcelEditable
    var finalRating: BigDecimal? = null
    @ExcelEditable
    @ExcelProperty("审定描述")
    var finalRatingComment: String? = null
}
