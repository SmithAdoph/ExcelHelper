//package com.adoph.excel;
//
//import org.apache.commons.lang3.StringUtils;
//import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
//import org.apache.poi.openxml4j.opc.OPCPackage;
//import org.apache.poi.openxml4j.opc.PackageAccess;
//import org.apache.poi.ss.usermodel.DateUtil;
//import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
//import org.apache.poi.xssf.eventusermodel.XSSFReader;
//import org.apache.poi.xssf.usermodel.XSSFRichTextString;
//import org.xml.sax.Attributes;
//import org.xml.sax.InputSource;
//import org.xml.sax.SAXException;
//import org.xml.sax.XMLReader;
//import org.xml.sax.helpers.DefaultHandler;
//
//import javax.xml.parsers.ParserConfigurationException;
//import javax.xml.parsers.SAXParser;
//import javax.xml.parsers.SAXParserFactory;
//import java.io.IOException;
//import java.io.InputStream;
//import java.math.BigDecimal;
//import java.text.DecimalFormat;
//import java.text.SimpleDateFormat;
//import java.util.ArrayList;
//import java.util.Iterator;
//import java.util.List;
//
//public class ExcelSaxReader {
//
//    enum CellDataType {
//        SHARED_STR, NUMBER, INLINE_STR, DATE
//    }
//
//    // Excel路径
//    private String path;
//    // 列总数
//    private int columns;
//    //文件流
//    private InputStream fis;
//
//    private ExcelSaxReader(InputStream fis, int columns) {
//        this.fis = fis;
//        this.columns = columns;
//    }
//
//    private ExcelSaxReader(String path, int columns) {
//        this.path = path;
//        this.columns = columns;
//    }
//
//    // Excel处理
//    private List<String[]> processByInputStream() throws IOException, SAXException, OpenXML4JException, ParserConfigurationException {
//        OPCPackage opcPackage = OPCPackage.open(fis);
//        ReadOnlySharedStringsTable stringsTable = new ReadOnlySharedStringsTable(opcPackage);
//        XSSFReader xssfReader = new XSSFReader(opcPackage);
//        List<String[]> list = new ArrayList<>();
//        Iterator<InputStream> sheets = xssfReader.getSheetsData();
//        while (sheets.hasNext()) {
//            InputStream stream = sheets.next();
//            list.addAll(processSheet(stringsTable, stream));
//            stream.close();
//        }
//        opcPackage.close();
//        return list;
//    }
//
//    // Excel处理
//    private List<String[]> process() throws IOException, SAXException, OpenXML4JException, ParserConfigurationException {
//        OPCPackage opcPackage = OPCPackage.open(path, PackageAccess.READ);
//        ReadOnlySharedStringsTable stringsTable = new ReadOnlySharedStringsTable(opcPackage);
//        XSSFReader xssfReader = new XSSFReader(opcPackage);
//        List<String[]> list = new ArrayList<>();
//        Iterator<InputStream> sheets = xssfReader.getSheetsData();
//        while (sheets.hasNext()) {
//            InputStream stream = sheets.next();
//            list.addAll(processSheet(stringsTable, stream));
//            stream.close();
//        }
//        opcPackage.close();
//        return list;
//    }
//
//    // 处理sheet
//    private List<String[]> processSheet(ReadOnlySharedStringsTable strings, InputStream sheetInputStream)
//            throws IOException, ParserConfigurationException, SAXException {
//        InputSource sheetSource = new InputSource(sheetInputStream);
//        SAXParserFactory saxFactory = SAXParserFactory.newInstance();
//        SAXParser saxParser = saxFactory.newSAXParser();
//        XMLReader sheetParser = saxParser.getXMLReader();
//
//        SheetHandler handler = new SheetHandler(strings, columns);
//        sheetParser.setContentHandler(handler);
//        sheetParser.parse(sheetSource);
//        return handler.getRows();
//    }
//
//    // 读取Excel入口方法
//    public static List<String[]> readExcel(InputStream is, int columns) throws OpenXML4JException, IOException, ParserConfigurationException, SAXException {
//        return new ExcelSaxReader(is, columns).processByInputStream();
//    }
//
//    // 读取Excel入口方法
//    public static List<String[]> readExcel(String path, int columns) throws OpenXML4JException, IOException, ParserConfigurationException, SAXException {
//        return new ExcelSaxReader(path, columns).process();
//    }
//
//    private class SheetHandler extends DefaultHandler {
//
//        private ReadOnlySharedStringsTable sharedStringsTable;
//        // 最大列数
//        private final int maxColumnCount;
//        // 是否是Value
//        private boolean isValue;
//        // 下一个数据类型
//        private CellDataType nextDataType;
//        // 当前列
//        private int currentColumn = -1;
//        private StringBuffer value;
//        // Excel行数据
//        private String[] record;
//        // Excel表数据
//        private List<String[]> rows = new ArrayList<>();
//        private DecimalFormat df = new DecimalFormat("0");
//
//        private SheetHandler(ReadOnlySharedStringsTable strings, int maxColumnCount) {
//            this.sharedStringsTable = strings;
//            this.maxColumnCount = maxColumnCount;
//            this.value = new StringBuffer();
//            record = new String[this.maxColumnCount];
//            rows.clear();
//        }
//
//        @Override
//        public void startElement(String uri, String localName, String name, Attributes attributes) {
//            if ("c".equals(name)) {
//                currentColumn = getCurrentColumn(attributes.getValue("r"));
//                nextDataType = CellDataType.NUMBER;
//
//                if ("s".equals(attributes.getValue("t"))) {
//                    nextDataType = CellDataType.SHARED_STR;
//                }
//                if ("inlineStr".equals(attributes.getValue("t"))) {
//                    nextDataType = CellDataType.INLINE_STR;
//                }
//                //日期格式
//                if ("1".equals(attributes.getValue("s"))) {
//                    nextDataType = CellDataType.DATE;
//                }
//            }
//
//            if ("v".equals(name) || "t".equals(name)) {
//                isValue = true;
//                value.setLength(0);
//            }
//        }
//
//        @Override
//        public void endElement(String uri, String localName, String name) {
//            String str = null;
//            if ("v".equals(name) || "t".equals(name)) {
//                if (CellDataType.SHARED_STR.equals(nextDataType)) {
//                    str = new XSSFRichTextString(sharedStringsTable.getEntryAt(Integer.parseInt(value.toString()))).toString();
//                }
//                if (CellDataType.INLINE_STR.equals(nextDataType)) {
//                    str = value.toString();
//                }
//                if (CellDataType.NUMBER.equals(nextDataType)) {
//                    str = df.format(new BigDecimal(value.toString()));
//                }
//                if (CellDataType.DATE.equals(nextDataType)) {
//                    str = new SimpleDateFormat("yyyy-MM-dd").format(DateUtil.getJavaDate(Double.parseDouble(value.toString())));
//                }
//                if (null == str) {
//                    str = "";
//                }
//                record[currentColumn] = str;
//            }
//
//            if ("row".equals(name)) {
//                if (columns > 0) {
//                    if (isNullArray()) {
//                        return;
//                    }
//                    rows.add(record.clone());
//                    for (int i = 0; i < record.length; i++) {
//                        record[i] = null;
//                    }
//                }
//            }
//        }
//
//        @Override
//        public void characters(char[] ch, int start, int length) {
//            if (isValue) {
//                value.append(ch, start, length);
//            }
//        }
//
//        private List<String[]> getRows() {
//            return rows;
//        }
//
//        // 获取当前列数
//        private int getCurrentColumn(String str) {
//            int firstDigit = -1;
//            for (int c = 0; c < str.length(); c++) {
//                if (Character.isDigit(str.charAt(c))) {
//                    firstDigit = c;
//                    break;
//                }
//            }
//            return nameToColumn(str.substring(0, firstDigit));
//        }
//
//        // 通过名字转换为列数（从0开始）
//        private int nameToColumn(String name) {
//
//            int column = -1;
//            for (int i = 0; i < name.length(); i++) {
//                int j = name.charAt(i);
//                column = (column + 1) * 26 + j - 'A';
//            }
//            return column;
//        }
//
//        // 判断是否为空数组
//        private Boolean isNullArray() {
//            int count = 0;
//            if (record.length > 0) {
//                for (String item : record) {
//                    if (StringUtils.isEmpty(item)) {
//                        count++;
//                    }
//                }
//                return count == record.length;
//            }
//            return true;
//        }
//    }
//}