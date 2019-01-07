package com.adoph.excel.sax;

/**
 * Sax模式解析Excel常量类
 *
 * @author Adoph
 * @version v1.0
 * @date 2018/12/21
 */
class ExcelSaxParseConstant {
    static final String DIMENSION_TAG = "dimension";
    static final String DIMENSION_ATTR_REF = "ref";

    static final String ROW_TAG = "row";

    /**
     * 行位置
     */
    static final String ROW_ATTR_POSITION = "r";

    /**
     * 列位置
     */
    static final String CELL_ATTR_POSITION = "r";

    static final String CELL_TAG = "c";
    static final String CELL_VALUE_TAG = "v";

    /**
     * 单元格数据类型
     */
    static final String CELL_ATTR_DATA_TYPE = "t";

    /**
     * 单元格的样式的索引, 样式记录存储在样式部件中
     */
    static final String CELL_ATTR_TYPE = "s";

    /**
     * 布尔类型
     */
    static final String CELL_ATTR_VALUE_BOOLEAN = "b";

    /**
     * 日期类型
     */
    static final String CELL_ATTR_VALUE_DATE = "1";

    /**
     * 错误格式
     */
    static final String CELL_ATTR_VALUE_ERROR = "e";

    /**
     * 内联元素，不使用共享字符串池的字符索引
     */
    static final String CELL_ATTR_VALUE_INLINE_STR = "inlineStr";

    /**
     * 公式类型
     */
    static final String CELL_ATTR_VALUE_STR = "str";

    /**
     * 共享字符串类型
     */
    static final String CELL_ATTR_VALUE_SHARED = "s";
}
