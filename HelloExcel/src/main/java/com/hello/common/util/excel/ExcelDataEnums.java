package com.hello.common.util.excel;

/**
 * 描述：exel格式枚举
 * Created by jay on 2017-9-18.
 */
public enum ExcelDataEnums {
    DEFAULT,//默认
    DOUBLE,//小数 200.00
    YYYYMM,//日期格式 2017-07，将会以字符格式，2017-07形式导出
    CONSTANT_SELECT,//固定值下拉框
    SINGLE_SELECT,//无关联下拉框
    UNION_PARENT_SELECT,//有关联父级下拉框
    UNION_CHILD_SELECT//有关联子级下拉框
}