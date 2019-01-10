package com.adoph.excel;

import com.adoph.excel.sax.ExcelReadListener;
import com.adoph.excel.sax.ExcelSaxReader;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.xml.sax.SAXException;

import javax.xml.parsers.ParserConfigurationException;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

/**
 * Excel工具
 * 支持方式：读、写
 * 支持版本：03（xls）、07(xlsx)版
 * 特殊：sax模式读取07版大文件，减少内存占用
 *
 * @author Adoph
 * @version v1.0
 * @date 2019/1/10
 */
public class ExcelHelper {

    /**
     * 通过sax解析
     * 支持格式：Excel 2007版(.xlsx)
     * 只读模式
     *
     * @param is 文件流
     * @return 数据集合
     */
    public static List<List<String>> readBySax(InputStream is) throws OpenXML4JException, ParserConfigurationException, SAXException, IOException {
        return ExcelSaxReader.readExcel(is);
    }

    /**
     * 监听模式通过sax解析，行读取
     * 支持格式：Excel 2007版(.xlsx)
     * 只读模式
     *
     * @param is 文件流
     */
    public static void readBySax(InputStream is, ExcelReadListener readListener) throws OpenXML4JException, ParserConfigurationException, SAXException, IOException {
        ExcelSaxReader.readExcel(is, readListener);
    }
}
