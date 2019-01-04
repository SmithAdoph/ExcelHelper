package com.adoph.excel.sax;

import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.helpers.DefaultHandler;

import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import static com.adoph.excel.sax.CellDataFormatIndex.DATE;
import static com.adoph.excel.sax.ExcelSaxParseConstant.*;
import static com.adoph.excel.sax.CellDataFormatIndex.*;

/**
 * 具体sheet处理器
 *
 * @author Tangqiandong
 * @version v1.0
 * @date 2019/1/2
 */
public class SheetHandler extends DefaultHandler {

    /**
     * @see org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable
     */
    private ReadOnlySharedStringsTable sharedStringsTable;

    /**
     * 是否为值
     */
    private boolean isValue;

    /**
     * 当前单元格内容数据类型
     */
    private CellDataType currentCellDataType;

    /**
     * 当前单元格格式索引值，默认未常规类型
     */
    private CellDataFormatIndex currentCellDataFormatIndex;

    /**
     * 当前第几列
     */
    private int currentColumn = -1;

    /**
     * 当前第几行
     */
    private int currentRow = -1;

    /**
     * 单元格值
     */
    private StringBuilder value;

    /**
     * 行：采用数组，空值占位
     */
    private String[] row;

    /**
     * 总数据
     */
    private List<List<String>> table;

    /**
     * 总列数
     */
    private int totalCol;

    /**
     * 总行数
     */
    private int totalRow;

    /**
     * 数字类型格式化（科学计数法）
     */
    private DecimalFormat df = new DecimalFormat("0");

    SheetHandler(ReadOnlySharedStringsTable strings) {
        this.sharedStringsTable = strings;
        this.value = new StringBuilder(50);
        table = new ArrayList<>();
    }

    @Override
    public void startElement(String uri, String localName, String qName, Attributes attributes) {
        //解析excel元数据，列数、行数等
        if (qName.equals(ExcelSaxParseConstant.DIMENSION_TAG)) {
            ExcelSheetMetadata metadata = ExcelSaxUtils.getMetadata(attributes.getValue(DIMENSION_ATTR_REF));
            this.totalCol = metadata.totalCol();//设置列
            row = new String[this.totalCol];//初始化行
            this.totalRow += metadata.totalRow();//多个sheet需要叠加总行数
            return;
        }

        //当前行数
        if (qName.equals(ROW_TAG)) {
            this.currentRow = Integer.valueOf(attributes.getValue(ROW_ATTR_POSITION));
            return;
        }

        //解析单元格,设置数据类型
        if (qName.equals(CELL_TAG)) {
            //当前列数
            currentColumn = ExcelSaxUtils.getCol(attributes.getValue(CELL_ATTR_POSITION));
            //默认是字符串
//            currentCellDataType = CellDataType.INLINE_STR;
            //数据类型(nullable)
//            String cellDataType = attributes.getValue(CELL_ATTR_DATA_TYPE);
            //样式索引(nullable)
            String cellStyleIndex = attributes.getValue(CELL_ATTR_TYPE);
            this.currentCellDataFormatIndex = cellStyleIndex != null ? CellDataFormatIndex.getByIndex(Integer.valueOf(cellStyleIndex)) : CellDataFormatIndex.GENERIC;
            return;
        }

        //解析cell value
        if (qName.equals(CELL_VALUE_TAG)) {
            isValue = true;
            value.setLength(0);
        } else {
            isValue = false;
        }
    }

    @Override
    public void endElement(String uri, String localName, String qName) {
        String str = null;
        if (qName.equals(CELL_VALUE_TAG)) {
            switch (currentCellDataFormatIndex) {
                case GENERIC:
                    str = new XSSFRichTextString(sharedStringsTable.getEntryAt(Integer.parseInt(value.toString()))).toString();
                    break;
                case NUMERICAL:
                    str = df.format(new BigDecimal(value.toString()));
                    break;
                case CURRENCY:
                    break;
                case ACCOUNTANT_DEDICATED:
                    break;
                case DATE:
                    str = new SimpleDateFormat("yyyy/MM/dd").format(DateUtil.getJavaDate(Double.parseDouble(value.toString())));
                    break;
                case DATETIME:
                    str = new SimpleDateFormat("yyyy/MM/dd hh:mm:ss").format(DateUtil.getJavaDate(Double.parseDouble(value.toString())));
                    break;
                case PERCENTAGE:
                    break;
                case FRACTION:
                    break;
                case SCIENTIFIC_NOTATION:
                    break;
                case TEXT:
                    break;
                case SPECIAL:
                    break;
                default:
                    str = new XSSFRichTextString(sharedStringsTable.getEntryAt(Integer.parseInt(value.toString()))).toString();
            }

            row[currentColumn - 1] = str;
        }

        if (qName.equals(ROW_TAG)) {
            table.add(Arrays.asList(row));
            row = null;
            row = new String[totalCol];
        }
    }

    @Override
    public void characters(char[] ch, int start, int length) {
        if (isValue) {
            value.append(ch, start, length);
        }
    }

    List<List<String>> getTable() {
        return table;
    }
}