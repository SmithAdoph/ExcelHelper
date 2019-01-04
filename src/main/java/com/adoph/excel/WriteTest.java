package com.adoph.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.junit.Test;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/**
 * TODO
 *
 * @author Tangqiandong
 * @version v1.0
 * @date 2018/12/17
 */
public class WriteTest {

    @Test
    public void testWrite() throws IOException {
        OutputStream out = new FileOutputStream("D:\\tmp/test.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook();
//        样式
        XSSFFont font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 12);

//        创建单元格风格对象
        XSSFCellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);//设置水平居中
        style.setFont(font);
        style.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setBorderBottom((short) 2);
        style.setBorderTop((short) 2);
        style.setBorderLeft((short) 2);
        style.setBorderRight((short) 2);
//        sheet
        XSSFSheet sheet = workbook.createSheet("01");
//        rows
        XSSFRow row = sheet.createRow(0);
//        cells
        XSSFCell c1 = row.createCell(0);
        c1.setCellStyle(style);
        c1.setCellValue("工资");
        XSSFCell c2 = row.createCell(1);
        c2.setCellStyle(style);
        c2.setCellValue("工资");
        workbook.write(out);
        out.close();
    }

    @Test
    public void testRead01() throws IOException {
        InputStream is = FileUtil.getResourcesFileInputStream("2007.xlsx");
        long now = System.currentTimeMillis();
//        2007.xlsx生成工作簿
        XSSFWorkbook workbook = new XSSFWorkbook(is);
//        获取总的sheet数量
//        int numberOfSheets = workbook.getNumberOfSheets();
        int numberOfSheets = 1;
//        遍历sheet
        for (int i = 0; i < numberOfSheets; i++) {
//            获取每一个sheet
            XSSFSheet sheet = workbook.getSheetAt(i);
            System.out.println("sheet name = " + sheet.getSheetName());
//            遍历行
            for (int j = 0; j < sheet.getPhysicalNumberOfRows(); j++) {
//                获取行
                XSSFRow row = sheet.getRow(j);
                if (row == null) {
                    continue;
                }
                List<String> rowList = new ArrayList<String>();
//                遍历列
                for (int k = 0; k < row.getPhysicalNumberOfCells(); k++) {
                    XSSFCell cell = row.getCell(k);
                    if (cell != null) {
                        cell.setCellType(1);
                        rowList.add(cell.getStringCellValue());
                    } else {
                        rowList.add(null);
                    }
                }
                System.out.println(j);
                System.out.println(Arrays.toString(rowList.toArray()));
            }
        }
        System.out.println("执行时间" + (System.currentTimeMillis() - now));
    }

}
