package com.adoph.excel.sax;

import java.util.List;

/**
 * Sax模式行读取监听
 *
 * @author Adoph
 * @version v1.0
 * @date 2019/1/10
 */
public interface ExcelReadListener {

    /**
     * 行读
     *
     * @param currentSheetTotalRow 当前sheet总行数,请注意不是所有sheet的行数合计
     * @param currentRow           当前解析到第几行
     * @param data                 行数据
     */
    public void readRow(int currentSheetTotalRow, int currentRow, List<String> data);

    /**
     * 读取当前sheet结束
     *
     * @param currentSheetTotalRow 当前sheet总行数
     */
    public void readDone(int currentSheetTotalRow);

}
