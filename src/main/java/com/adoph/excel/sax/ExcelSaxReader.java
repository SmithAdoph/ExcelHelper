package com.adoph.excel.sax;

import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.StylesTable;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * Excel 07版解析工具类
 * Sax模式解析：
 * 1.减少内存占用，解决读xlsx大文件内存溢出的问题
 * 2.特殊单元格数据类型无法处理，比如包含日期、数学公式等单元格
 * <p>
 * 参考文档：https://docs.microsoft.com/en-us/office/open-xml/open-xml-sdk
 *
 * @author Adoph
 * @version v1.0
 * @date 2018/12/29
 */
public class ExcelSaxReader {

    /**
     * 读Excel 07版xlsx格式
     *
     * @param is 文件流
     * @return 数据集合
     */
    public static List<List<String>> readExcel(InputStream is) throws OpenXML4JException, IOException, ParserConfigurationException, SAXException {
        return new ExcelSaxReader().process(is);
    }

    /**
     * 读Excel 07版xlsx格式
     *
     * @param is 文件流
     */
    public static void readExcel(InputStream is, ExcelReadListener readListener) throws OpenXML4JException, IOException, ParserConfigurationException, SAXException {
        new ExcelSaxReader().process(is, readListener);
    }

    /**
     * 读Excel 07版xlsx格式
     *
     * @param path 文件路径
     * @return 数据集合
     */
    public static List<List<String>> readExcel(String path) throws OpenXML4JException, IOException, ParserConfigurationException, SAXException {
        return new ExcelSaxReader().process(path);
    }

    /**
     * 解析excel
     *
     * @return 数据集合
     */
    private List<List<String>> process(InputStream is) throws OpenXML4JException, ParserConfigurationException, SAXException, IOException {
        List<List<String>> data;
        try (OPCPackage pkg = OPCPackage.open(is)) {
            data = processSheets(pkg);
        }
        return data;
    }

    /**
     * 监听模式解析excel
     */
    private void process(InputStream is, ExcelReadListener readListener) throws OpenXML4JException, ParserConfigurationException, SAXException, IOException {
        try (OPCPackage pkg = OPCPackage.open(is)) {
            processSheets(pkg, readListener);
        }
    }

    /**
     * 解析excel
     *
     * @return 数据集合
     */
    private List<List<String>> process(String path) throws OpenXML4JException, ParserConfigurationException, SAXException, IOException {
        List<List<String>> data;
        try (OPCPackage pkg = OPCPackage.open(path, PackageAccess.READ)) {
            data = processSheets(pkg);
        }
        return data;
    }

    /**
     * 解析sheets
     *
     * @param pkg OPCPackage
     * @return 数据集合
     * @see org.apache.poi.openxml4j.opc.OPCPackage
     */
    private List<List<String>> processSheets(OPCPackage pkg) throws IOException, OpenXML4JException, SAXException, ParserConfigurationException {
        ReadOnlySharedStringsTable sst = new ReadOnlySharedStringsTable(pkg);
        XSSFReader xssfReader = new XSSFReader(pkg);
        StylesTable st = xssfReader.getStylesTable();
        List<List<String>> list = new ArrayList<>();
        Iterator<InputStream> sheets = xssfReader.getSheetsData();
        while (sheets.hasNext()) {
            InputStream stream = sheets.next();
            list.addAll(processSheet(sst, st, stream));
            stream.close();
        }
        return list;
    }

    /**
     * 解析sheets
     *
     * @param pkg OPCPackage
     * @see org.apache.poi.openxml4j.opc.OPCPackage
     */
    private void processSheets(OPCPackage pkg, ExcelReadListener readListener) throws IOException, OpenXML4JException, SAXException, ParserConfigurationException {
        ReadOnlySharedStringsTable sst = new ReadOnlySharedStringsTable(pkg);
        XSSFReader xssfReader = new XSSFReader(pkg);
        StylesTable st = xssfReader.getStylesTable();
        Iterator<InputStream> sheets = xssfReader.getSheetsData();
        while (sheets.hasNext()) {
            InputStream stream = sheets.next();
            processSheet(sst, st, stream, readListener);
            stream.close();
        }
    }

    /**
     * 解析sheet
     *
     * @param sst              共享字符串
     * @param sheetInputStream sheet流
     * @return 数据集合
     * @see org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable
     */
    private List<List<String>> processSheet(ReadOnlySharedStringsTable sst, StylesTable st, InputStream sheetInputStream)
            throws IOException, ParserConfigurationException, SAXException {
        InputSource sheetSource = new InputSource(sheetInputStream);
        SAXParserFactory saxFactory = SAXParserFactory.newInstance();
        SAXParser saxParser = saxFactory.newSAXParser();
        XMLReader sheetParser = saxParser.getXMLReader();
        SheetHandler handler = new SheetHandler(sst, st);
        sheetParser.setContentHandler(handler);
        sheetParser.parse(sheetSource);
        return handler.getTable();
    }

    /**
     * 解析sheet
     *
     * @param sst              共享字符串
     * @param sheetInputStream sheet流
     * @see org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable
     */
    private void processSheet(ReadOnlySharedStringsTable sst, StylesTable st, InputStream sheetInputStream, ExcelReadListener readListener)
            throws IOException, ParserConfigurationException, SAXException {
        InputSource sheetSource = new InputSource(sheetInputStream);
        SAXParserFactory saxFactory = SAXParserFactory.newInstance();
        SAXParser saxParser = saxFactory.newSAXParser();
        XMLReader sheetParser = saxParser.getXMLReader();
        SheetHandler handler = new SheetHandler(sst, st, readListener);
        sheetParser.setContentHandler(handler);
        sheetParser.parse(sheetSource);
    }

}