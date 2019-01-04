package com.adoph.excel.sax;

/**
 * 单元格数据类型
 *
 * @author Tangqiandong
 * @version v1.0
 * @date 2019/1/2
 */
public enum CellDataType {
    /**
     * 共享字符串类型
     */
    SHARED_STR("s"),

    /**
     * 数字
     */
    NUMBER("b"),

    /**
     * 内联元素，允许作为子元素，不存在公式
     */
    INLINE_STR("inlineStr"),

    /**
     * Cell containing a formula string.
     */
    STRING("str"),

    /**
     * Cell containing an error.
     */
    ERROR("e");

    private String index;

    public String getIndex() {
        return index;
    }

    public void setIndex(String index) {
        this.index = index;
    }

    CellDataType(String index) {
        this.index = index;
    }
}