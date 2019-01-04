package com.adoph.excel.sax;

import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.xml.sax.SAXException;

import javax.xml.parsers.ParserConfigurationException;
import java.io.IOException;
import java.util.List;

/**
 * 测试
 */
public class ReadTest {

    public static void main(String[] args) throws OpenXML4JException, ParserConfigurationException, SAXException, IOException {
        // Excel路径和Excel列数
        long start = System.currentTimeMillis();
        String filePath = "C:\\Users\\Tangqiandong\\Desktop\\测试文件 - 副本.xlsx";
        List<List<String>> dataList = ExcelSaxReader.readExcel(filePath);
        System.out.println("读取总计" + dataList.size() + "条，耗时：" + (System.currentTimeMillis() - start) + "毫秒");
        int count = 0;

        for (List<String> row : dataList) {
            System.out.println(row.toString());
            if ((count++) == 100) {
                return;
            }
        }
    }
}