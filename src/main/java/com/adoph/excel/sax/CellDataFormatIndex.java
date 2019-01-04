package com.adoph.excel.sax;

/**
 * 单元格格式索引
 *
 * @author Tangqiandong
 * @version v1.0
 * @date 2019/1/4
 */
public enum CellDataFormatIndex {
    /**
     * 常规
     */
    GENERIC(0),

    /**
     * 数值
     */
    NUMERICAL(1),

    /**
     * 货币
     */
    CURRENCY(2),

    /**
     * 会计专用
     */
    ACCOUNTANT_DEDICATED(3),

    /**
     * 日期：年月日
     */
    DATE(4),

    /**
     * 时间：年月日：时分秒
     */
    DATETIME(5),

    /**
     * 百分比：%
     */
    PERCENTAGE(6),

    /**
     * 百分数
     */
    FRACTION(7),

    /**
     * 科学计数
     */
    SCIENTIFIC_NOTATION(8),

    /**
     * 文本
     */
    TEXT(9),

    /**
     * 特殊：比如邮政编码、中文小写数字等
     */
    SPECIAL(10);

    private int index;

    CellDataFormatIndex(int index) {
        this.index = index;
    }

    public int getIndex() {
        return index;
    }

    public void setIndex(int index) {
        this.index = index;
    }

    public static CellDataFormatIndex getByIndex(int index) {
        for (CellDataFormatIndex item : CellDataFormatIndex.values()) {
            if (item.getIndex() == index) {
                return item;
            }
        }
        return GENERIC;
    }
}
