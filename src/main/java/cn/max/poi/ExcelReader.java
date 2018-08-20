package cn.max.poi;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFReader.SheetIterator;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Iterator;
import java.util.List;

import static cn.max.poi.value.CellDataType.*;


/**
 * 使用SAX方法解析Excel（只能解析2007以上的版本，即尾缀为.xlsx）
 *
 * @author MaxStar
 * @date 2018/8/3
 */
public class ExcelReader extends DefaultHandler {

    /**
     * 共享字符串表
     */
    private SharedStringsTable sst;

    /**
     * 上一次的内容
     */
    private String lastContents;

    /**
     * 上一个标签是否为C
     */
    private boolean preEleIsC = false;

    /**
     * 当前列数
     */
    private int curCol = 0;

    /**
     * 单元格数据类型，默认为字符串类型
     */
    private String nextDataType;

    /**
     * 格式解析
     */
    private final DataFormatter formatter = new DataFormatter();

    /**
     * 单元格样式，用于格式转换
     */
    private short formatIndex = -1;
    private String formatString = null;

    private StylesTable stylesTable;
    /**
     * 上一个列号
     */
    private int preCellColNum = 0;

    /**
     * 单元格内容
     */
    private List<String> rowValueList = new ArrayList<>();

    /**
     * 存储每一行所有单元格的list
     */
    private List<List<String>> rowList = new ArrayList<>();

    /**
     * 在解析多个sheet时使用，将每个sheet的内容存进List
     */
    private List<List<List<String>>> sheetList = new ArrayList<>();

    public void processOne(InputStream inputStream, int sheetId) throws Exception {
        OPCPackage opcPackage = OPCPackage.open(inputStream);
        processOneSheet(opcPackage, sheetId);
    }


    public void processAll(InputStream inputStream) throws Exception {
        OPCPackage opcPackage = OPCPackage.open(inputStream);
        processAll(opcPackage);
    }


    public void processByName(InputStream inputStream, String sheetName) throws Exception {
        OPCPackage opcPackage = OPCPackage.open(inputStream);
        processBySheetName(opcPackage, sheetName);
    }

    private XMLReader getXMLReader(OPCPackage opcPackage) throws Exception {
        XSSFReader r = new XSSFReader(opcPackage);
        //   this.sharedStringsTable = new ReadOnlySharedStringsTable(opcPackage);
        this.stylesTable = r.getStylesTable();
        SharedStringsTable sst = r.getSharedStringsTable();
        return fetchSheetParser(sst);
    }

    private void parseSheet(XMLReader parser, InputStream sheet) throws IOException, SAXException {
        InputSource sheetSource = new InputSource(sheet);
        parser.parse(sheetSource);
        sheet.close();
    }

    /**
     * 只遍历一个电子表格，其中sheetId为要遍历的sheet索引，从1开始，1-3
     *
     * @throws Exception
     */
    private void processOneSheet(OPCPackage opcPackage, int sheetId) throws Exception {
        XSSFReader r = new XSSFReader(opcPackage);
        //   this.sharedStringsTable = new ReadOnlySharedStringsTable(opcPackage);
        this.stylesTable = r.getStylesTable();
        SharedStringsTable sst = r.getSharedStringsTable();
        XMLReader parser = fetchSheetParser(sst);

        // 根据 rId# 或 rSheet# 查找sheet
        InputStream sheet = r.getSheet("rId" + sheetId);
        InputSource sheetSource = new InputSource(sheet);
        parser.parse(sheetSource);
        sheet.close();
    }

    /**
     * 遍历工作簿中所有的电子表格
     */
    private void processAll(OPCPackage opcPackage) throws Exception {
        XSSFReader r = new XSSFReader(opcPackage);
        this.stylesTable = r.getStylesTable();
        SharedStringsTable sst = r.getSharedStringsTable();
        XMLReader parser = fetchSheetParser(sst);
        SheetIterator sheetiterator = (SheetIterator) r.getSheetsData();
        while (sheetiterator.hasNext()) {
            InputStream sheet = sheetiterator.next();
            System.out.println("当前表格名：" + sheetiterator.getSheetName());
            InputSource sheetSource = new InputSource(sheet);
            parser.parse(sheetSource);
            sheet.close();
            // 添加进sheetList内，并清空rowList
            sheetList.add(rowList);
            rowList = new ArrayList<>();
        }
    }

    /**
     * 遍历指定名称的单元格
     *
     * @param opcPackage
     * @param sheetName
     * @throws Exception
     */
    private void processBySheetName(OPCPackage opcPackage, String sheetName) throws Exception {
        boolean notFindSheet = true;
        XSSFReader r = new XSSFReader(opcPackage);
        this.stylesTable = r.getStylesTable();
        SharedStringsTable sst = r.getSharedStringsTable();
        XMLReader parser = fetchSheetParser(sst);
        SheetIterator sheetiterator = (SheetIterator) r.getSheetsData();
        while (sheetiterator.hasNext()) {
            InputStream sheet = sheetiterator.next();
            if (sheetiterator.getSheetName().equals(sheetName)) {
                InputSource sheetSource = new InputSource(sheet);
                parser.parse(sheetSource);
                sheet.close();
                notFindSheet = false;
                break;
            }
        }

        if (notFindSheet) {
            System.out.println("未找表格:" + sheetName);
        }
    }


    private XMLReader fetchSheetParser(SharedStringsTable sst) throws SAXException {
        XMLReader parser = XMLReaderFactory
                .createXMLReader("org.apache.xerces.parsers.SAXParser");
        this.sst = sst;
        parser.setContentHandler(this);
        return parser;
    }

    /**
     * 读取单元格的格式
     *
     * @param uri
     * @param localName
     * @param name
     * @param attributes
     */
    @Override
    public void startElement(String uri, String localName, String name, Attributes attributes) {
        // c => 单元格
        if (name.equals("c")) {
            // 检测有没有漏行
            String cellCol = attributes.getValue("r").replaceAll("\\d", "").trim();
            int cellColNum = excelColStrToNum(cellCol);
            if (preCellColNum != 0 && cellColNum - preCellColNum > 1) {
                // 计算两个c标签之间的差值
                int diff = cellColNum - preCellColNum;
                for (int i = 0; i < (diff - 1); i++) {
                    rowValueList.add(curCol, null);
                }
                curCol += (diff - 1);
            }
            preCellColNum = cellColNum;

            // 判断上一个标签是否还是c，如果是c则表示漏了一行(使用清除内容会导致没有v标签)
            if (preEleIsC) {
                rowValueList.add(curCol, null);
                curCol++;
                cleanCellFormate();
            }
            preEleIsC = true;
            String cellType = attributes.getValue("t");
            String cellStyleStr = attributes.getValue("s");
            if (cellType != null) {
                // 判断单元格类型
                switch (cellType) {
                    case SSTINDEX:
                        nextDataType = SSTINDEX;
                        break;
                    case "b":
                        nextDataType = BOOL;
                        break;
                    case "inlineStr":
                        nextDataType = INLINESTR;
                        break;
                    case "str":
                        nextDataType = FORMULA;
                        break;
                    default:
                        nextDataType = NEED_FORMAT;
                        break;
                }
            }
            if (nextDataType == null && cellStyleStr != null) {
                int styleIndex = Integer.parseInt(cellStyleStr);
                // 获取单元格样式
                XSSFCellStyle style = stylesTable.getStyleAt(styleIndex);
                formatIndex = style.getDataFormat();
                formatString = style.getDataFormatString();
                if (formatString == null) {
                    formatString = BuiltinFormats.getBuiltinFormat(formatIndex);
                }
                nextDataType = NEED_FORMAT;

                if (formatString == "m/d/yy" || formatString == "yyyy-mm-dd" || formatString.contains("[$-F800]")) {
                    nextDataType = DATE;
                    formatString = "yyyy-MM-dd";
                    //      formatString = "yyyy-MM-dd hh:mm:ss.SSS";
                }

                if (formatString == "h:mm" || formatString.contains("[$-F400]")) {
                    nextDataType = TIME;
                    formatString = "hh:mm:ss.SSS";
                }
            }
        }

        // 置空
        lastContents = "";
    }

    /**
     * 读取单元格的内容
     */
    @Override
    public void endElement(String uri, String localName, String name) {
        // 根据SST的索引值的到单元格的真正要存储的字符串
        // 这时characters()方法可能会被调用多次
        if (nextDataType != null && nextDataType.equals(SSTINDEX)) {
            int idx = Integer.parseInt(lastContents);
            lastContents = new XSSFRichTextString(sst.getEntryAt(idx)).toString().trim();
        }

        // v => 单元格的值， 将单元格内容加入rowlist中
        String value;
        if (name.equals("v")) {
            if (nextDataType == null) {
                value = lastContents;
            } else {
                switch (nextDataType) {
                    case BOOL:
                        value = lastContents.charAt(0) == '0' ? "FALSE" : "TRUE";
                        break;
                    case INLINESTR:
                        value = new XSSFRichTextString(lastContents).toString();
                        break;
                    case FORMULA:
                    case SSTINDEX:
                        value = lastContents;
                        break;
                    case NEED_FORMAT:
                        if (formatString != null) {
                            value = formatter.formatRawCellContents(Double.parseDouble(lastContents), formatIndex, formatString).trim();
                        } else {
                            value = lastContents;
                        }
                        break;
                    case DATE:
                    case TIME:
                        value = formatter.formatRawCellContents(Double.parseDouble(lastContents), formatIndex, formatString);
                        break;
                    default:
                        value = lastContents;
                        break;
                }
            }
            value = value.equals("") ? null : value;
            rowValueList.add(curCol, value);
            preEleIsC = false;
            curCol++;
            cleanCellFormate();
        } else if (name.equals("row")) {
            try {
                if (!rowValueList.isEmpty()) {
                    List<String> tempRowValues = new ArrayList<>(rowValueList);
                    tempRowValues.removeAll(Collections.singleton(null));
                    if (tempRowValues.size() > 0) {
                        rowList.add(rowValueList);
                    }
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
            rowValueList = new ArrayList<>();
            preEleIsC = false;
            preCellColNum = 0;
            curCol = 0;
            cleanCellFormate();
        }
    }

    /**
     * 将列名由英文转换为数字，用于计算跳过了多少空格
     *
     * @param column 列名
     * @return 列数
     */
    private int excelColStrToNum(String column) {
        int result = -1;
        for (int i = 0; i < column.length(); i++) {
            result = (result + 1) * 26 + (column.charAt(i) - 'A');
        }
        return result + 1;
    }

    @Override
    public void characters(char[] ch, int start, int length) {
        //得到单元格内容的值
        lastContents += new String(ch, start, length);
    }

    /**
     * 重置单元格格式
     */
    private void cleanCellFormate() {
        formatIndex = -1;
        formatString = null;
        nextDataType = null;
    }

    /**
     * 获取每一行内容
     *
     * @return 行集合
     */
    public List<List<String>> getRowList() {
        return rowList;
    }

    /**
     * 获取每一个sheet的内容
     *
     * @return 表集合
     */
    public List<List<List<String>>> getSheetList() {
        return sheetList;
    }
}