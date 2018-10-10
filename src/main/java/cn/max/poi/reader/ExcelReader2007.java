package cn.max.poi.reader;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFReader.SheetIterator;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;


/**
 * 使用SAX方法解析Excel（只能解析2007以上的版本，即尾缀为.xlsx）
 *
 * @author MaxStar
 * @date 2018/8/3
 */
public class ExcelReader2007 extends DefaultHandler {

    /**
     * 共享字符串表
     */
    private ReadOnlySharedStringsTable sst;

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
     * 单元格数据类型
     */
    private Integer nextDataType;
    private static final int SSTINDEX = 1;
    private static final int BOOL = 2;
    private static final int INLINESTR = 3;
    private static final int NEED_FORMAT = 4;

    /**
     * 格式解析
     */
    private static final DataFormatter FORMATTER = new DataFormatter();

    /**
     * 单元格样式，用于格式转换
     */
    private short format = -1;
    private String formatString = null;

    private StylesTable stylesTable;
    /**
     * 上一个列号
     */
    private int preCellColNum = 0;

    /**
     * 单元格内容
     */
    private List<String> cellValueList = new ArrayList<>();

    /**
     * 存储每一行所有单元格的list
     */
    private List<List<String>> rowList = new ArrayList<>();

    /**
     * 在解析多个sheet时使用，将每个sheet的内容存进List
     */
    private List<List<List<String>>> sheetList = new ArrayList<>();

    public void process(InputStream inputStream, int sheetId) throws Exception {
        OPCPackage opcPackage = OPCPackage.open(inputStream);
        process(opcPackage, sheetId);
    }

    public void process(InputStream inputStream, String sheetName) throws Exception {
        OPCPackage opcPackage = OPCPackage.open(inputStream);
        process(opcPackage, sheetName);
    }

    public void processAll(InputStream inputStream) throws Exception {
        OPCPackage opcPackage = OPCPackage.open(inputStream);
        processAll(opcPackage);
    }


    /**
     * 只遍历一个电子表格，其中sheetId为要遍历的sheet索引，从1开始，1-3
     *
     * @throws Exception
     */
    private void process(OPCPackage opcPackage, int sheetId) throws Exception {
        XSSFReader r = new XSSFReader(opcPackage);
        this.stylesTable = r.getStylesTable();
        this.sst = new ReadOnlySharedStringsTable(opcPackage);
        XMLReader parser = fetchSheetParser(sst);

        // 根据 rId# 或 rSheet# 查找sheet
        InputStream sheet = r.getSheet("rId" + sheetId);
        InputSource sheetSource = new InputSource(sheet);
        parser.parse(sheetSource);
        sheet.close();
    }


    /**
     * 遍历指定名称的单元格
     *
     * @param opcPackage
     * @param sheetName
     * @throws Exception
     */
    private void process(OPCPackage opcPackage, String sheetName) throws Exception {
        boolean notFindSheet = true;
        XSSFReader r = new XSSFReader(opcPackage);
        this.stylesTable = r.getStylesTable();
        this.sst = new ReadOnlySharedStringsTable(opcPackage);
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

    /**
     * 遍历工作簿中所有的电子表格
     */
    private void processAll(OPCPackage opcPackage) throws Exception {
        XSSFReader r = new XSSFReader(opcPackage);
        this.stylesTable = r.getStylesTable();
        this.sst = new ReadOnlySharedStringsTable(opcPackage);
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

    private XMLReader fetchSheetParser(ReadOnlySharedStringsTable sst) throws SAXException {
        XMLReader parser = XMLReaderFactory.createXMLReader("org.apache.xerces.parsers.SAXParser");
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
                    cellValueList.add(curCol, null);
                }
                curCol += (diff - 1);
            }
            preCellColNum = cellColNum;

            // 判断上一个标签是否还是c，如果是c则表示漏了一行(使用清除内容会导致没有v标签)
            if (preEleIsC) {
                cellValueList.add(curCol, null);
                curCol++;
                cleanCellFormate();
            }
            preEleIsC = true;
            String cellType = attributes.getValue("t");
            String cellStyleStr = attributes.getValue("s");
            if (cellType != null) {
                // 判断单元格类型
                switch (cellType) {
                    case "s":
                        nextDataType = SSTINDEX;
                        break;
                    case "b":
                        nextDataType = BOOL;
                        break;
                    case "inlineStr":
                        nextDataType = INLINESTR;
                        break;
                    default:
                        nextDataType = null;
                        break;
                }
            }
            if (nextDataType == null) {
                int styleIndex;
                try {
                    styleIndex = Integer.parseInt(cellStyleStr);
                } catch (NumberFormatException e) {
                    styleIndex = 0;
                }

                // 获取单元格样式
                XSSFCellStyle style = stylesTable.getStyleAt(styleIndex);
                format = style.getDataFormat();
                formatString = style.getDataFormatString();
                if (formatString == null) {
                    formatString = BuiltinFormats.getBuiltinFormat(format);
                }

                // 判断是否是日期时间
                if (formatString.equals("m/d/yy") || formatString.equals("yyyy-mm-dd") || formatString.contains("[$-F800]")) {
                    formatString = "yyyy-MM-dd";
                    //      formatString = "yyyy-MM-dd hh:mm:ss.SSS";
                } else if (formatString.equals("h:mm") || formatString.contains("[$-F400]")) {
                    formatString = "HH:mm:ss";
                }
                nextDataType = NEED_FORMAT;
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
        // v => 单元格的值， 将单元格内容加入rowlist中
        String value;
        if (name.equals("v")) {
            if (nextDataType == null) {
                value = lastContents;
            } else {
                // 根据SST的索引值的到单元格的真正要存储的字符串
                // 这时characters()方法可能会被调用多次
                if (nextDataType == SSTINDEX) {
                    int idx = Integer.parseInt(lastContents);
                    value = new XSSFRichTextString(sst.getEntryAt(idx)).toString().trim();
                } else if (nextDataType == BOOL) {
                    value = lastContents.charAt(0) == '0' ? "FALSE" : "TRUE";
                } else if (nextDataType == INLINESTR) {
                    value = new XSSFRichTextString(lastContents).toString();
                } else if (nextDataType == NEED_FORMAT && formatString != null) {
                    value = FORMATTER.formatRawCellContents(Double.parseDouble(lastContents), format, formatString).trim();
                } else {
                    value = lastContents;
                }
            }
            value = value.equals("") ? null : value;
            cellValueList.add(curCol, value);
            preEleIsC = false;
            curCol++;
            cleanCellFormate();
        } else if (name.equals("row")) {
            try {
                if (!cellValueList.isEmpty()) {
                    List<String> tempRowValues = new ArrayList<>(cellValueList);
                    tempRowValues.removeAll(Collections.singleton(null));
                    if (tempRowValues.size() > 0) {
                        rowList.add(cellValueList);
                    }
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
            cellValueList = new ArrayList<>();
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
        format = -1;
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