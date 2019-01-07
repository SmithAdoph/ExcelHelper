package com.adoph.excel.sax;

import java.util.EnumMap;

/**
 * Excel元数据
 *
 * @author Adoph
 * @version v1.0
 * @date 2018/12/29
 */
public class ExcelSheetMetadata extends EnumMap<ExcelMetadata, Integer> {

    public ExcelSheetMetadata(Class<ExcelMetadata> keyType) {
        super(keyType);
        this.put(ExcelMetadata.START_COL, 1);
        this.put(ExcelMetadata.START_ROW, 1);
        this.put(ExcelMetadata.TOTAL_COL, 1);
        this.put(ExcelMetadata.TOTAL_ROW, 1);
    }

    public void startCol(Integer num) {
        this.put(ExcelMetadata.START_COL, num);
    }

    public Integer startCol() {
        return this.get(ExcelMetadata.START_COL);
    }

    public void startRow(Integer num) {
        this.put(ExcelMetadata.START_ROW, num);
    }

    public Integer startRow() {
        return this.get(ExcelMetadata.START_ROW);
    }

    public void totalCol(Integer num) {
        this.put(ExcelMetadata.TOTAL_COL, num);
    }

    public Integer totalCol() {
        return this.get(ExcelMetadata.TOTAL_COL);
    }

    public void totalRow(Integer num) {
        this.put(ExcelMetadata.TOTAL_ROW, num);
    }

    public Integer totalRow() {
        return this.get(ExcelMetadata.TOTAL_ROW);
    }
}
