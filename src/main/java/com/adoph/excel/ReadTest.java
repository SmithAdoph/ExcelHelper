package com.adoph.excel;

import com.adoph.excel.sax.ExcelReadListener;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.junit.Test;
import org.xml.sax.SAXException;

import javax.xml.parsers.ParserConfigurationException;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

/**
 * 测试
 */
public class ReadTest {

    /**
     * sax读取
     * 支持格式：07版本xlsx
     */
    @Test
    public void readBySax1() throws OpenXML4JException, ParserConfigurationException, SAXException, IOException {
        long start = System.currentTimeMillis();
        InputStream is = FileUtil.getResourcesFileInputStream("测试文件.xlsx");
        List<List<String>> dataList = ExcelHelper.readBySax(is);
        System.out.println("读取总计" + dataList.size() + "条，耗时：" + (System.currentTimeMillis() - start) + "毫秒");
        print(dataList);

//        long start1 = System.currentTimeMillis();
//        InputStream is1 = FileUtil.getResourcesFileInputStream("4月薪酬明细模板2.xlsx");
//        List<List<String>> dataList1 = ExcelHelper.readBySax(is1);
//        System.out.println("读取总计" + dataList1.size() + "条，耗时：" + (System.currentTimeMillis() - start1) + "毫秒");
//        print(dataList1);
    }

    /**
     * sax监听读取
     * 支持格式：07版本xlsx
     */
    @Test
    public void readBySax2() throws OpenXML4JException, ParserConfigurationException, SAXException, IOException {
        final long start = System.currentTimeMillis();
        InputStream is = FileUtil.getResourcesFileInputStream("2007.xlsx");
        ExcelHelper.readBySax(is, new ExcelReadListener() {
            @Override
            public void readRow(int currentSheetTotalRow, int currentRow, List<String> data) {
                System.out.println("第" + currentRow + "行" + data.toString());
            }

            @Override
            public void readDone(int currentSheetTotalRow) {
                System.out.println("读取总计" + currentSheetTotalRow + "条，耗时：" + (System.currentTimeMillis() - start) + "毫秒");
            }
        });

    }

    private static void print(List<List<String>> dataList) {
        for (List<String> row : dataList) {
            System.out.println(row.toString());
        }
    }
}