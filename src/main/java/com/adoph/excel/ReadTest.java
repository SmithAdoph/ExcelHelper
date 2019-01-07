package com.adoph.excel;

import com.adoph.excel.sax.ExcelSaxReader;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.xml.sax.SAXException;

import javax.xml.parsers.ParserConfigurationException;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

/**
 * 测试
 */
public class ReadTest {

    public static void main(String[] args) throws OpenXML4JException, ParserConfigurationException, SAXException, IOException {
        long start = System.currentTimeMillis();
        InputStream is = FileUtil.getResourcesFileInputStream("2007.xlsx");
        List<List<String>> dataList = ExcelSaxReader.readExcel(is);
        System.out.println("读取总计" + dataList.size() + "条，耗时：" + (System.currentTimeMillis() - start) + "毫秒");
        print(dataList);
    }
    private static void print(List<List<String>> dataList) {
        for (List<String> row : dataList) {
            System.out.println(row.toString());
        }
    }
}