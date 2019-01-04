package com.adoph.excel.sax;

import org.junit.Test;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * POI Sax模式解析Excel工具类
 *
 * @author Tangqiandong
 * @version v1.0
 * @date 2019/1/2
 */
public class ExcelSaxUtils {

    /**
     * 起始行列、总列数、总行数匹配正则
     */
    private final static String DIMENSION_REF_REG = "([A-Z]+)(\\d+)";

    /**
     * 获取行数
     *
     * @param currentCellIndex 当前单元格下标
     * @return 当前第几行
     */
    public static int getRow(String currentCellIndex) {
        int row = 0;
        if (currentCellIndex != null) {
            String rowStr = currentCellIndex.replaceAll("[A-Z]", "").replaceAll("[a-z]", "");
            row = Integer.parseInt(rowStr) - 1;
        }
        return row;
    }

    /**
     * 获取列数
     *
     * @param currentCellIndex 当前单元格下标
     * @return 当前第几列, 从1开始
     */
    static int getCol(String currentCellIndex) {
        int col = 0;
        if (currentCellIndex != null) {
            char[] currentIndex = currentCellIndex.replaceAll("[0-9]", "").toCharArray();
            for (int i = 0; i < currentIndex.length; i++) {
                col += (currentIndex[i] - '@') * Math.pow(26, (currentIndex.length - i - 1));
            }
        }
        return col;
    }


    /**
     * 获取excel元数据
     *
     * @param dimensionRefVal 对应dimension内容
     * @return ExcelSheetMetadata
     */
    static ExcelSheetMetadata getMetadata(String dimensionRefVal) {
        String[] arr = dimensionRefVal.split(":");
        ExcelSheetMetadata metadata = new ExcelSheetMetadata(ExcelMetadata.class);
        Pattern p = Pattern.compile(DIMENSION_REF_REG);
        Matcher m1 = p.matcher(arr[0]);
        if (m1.find()) {
            metadata.startCol(getCol(m1.group(1)));
            metadata.startRow(Integer.valueOf(m1.group(2)));
        }
        if (arr.length == 2) {
            Matcher m2 = p.matcher(arr[1]);
            if (m2.find()) {
                metadata.totalCol(getCol(m2.group(1)));
                metadata.totalRow(Integer.valueOf(m2.group(2)));
            }
        }
        return metadata;
    }

    @Test
    public void test() {
//        System.out.println(getCol("A3"));
//        System.out.println(getCol("BC"));
        System.out.println(getMetadata("A1:E2"));
    }
}




