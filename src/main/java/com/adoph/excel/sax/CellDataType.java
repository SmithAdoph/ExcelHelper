package com.adoph.excel.sax;

/**
 * 单元格数据类型:
 * <p>
 * 参考：https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellvalues?view=openxml-2.8.1
 *
 * @author Adoph
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
    NUMBER("n"),

    /**
     * 布尔
     */
    BOOLEAN("b"),

    /**
     * 日期类型：
     * <p>
     * When the item is serialized out as xml, its value is "d".This item is only available in Office2010.
     */
    DATE("d"),

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
    ERROR("e"),

    /**
     * 未设置
     */
    NONE("null");

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

    public static CellDataType getByIndex(String index) {
        for (CellDataType item :
                CellDataType.values()) {
            if (item.getIndex().equals(index)) {
                return item;
            }
        }
        return CellDataType.NONE;
    }
}