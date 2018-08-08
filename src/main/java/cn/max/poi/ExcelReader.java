package cn.max.poi;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.File;
import java.io.InputStream;
import java.util.*;

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
     * 判断单元格类型是否是字符串
     */
    private boolean nextIsString;

    /**
     * 上一个标签名
     */
    private char preEle = '0';

    /**
     * 当前列数
     */
    private int curCol = 0;

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
    private List<List<List<String>>> SheetList = new ArrayList<>();

    public void processOne(InputStream inputStream, int sheetId) throws Exception {
        OPCPackage opcPackage = OPCPackage.open(inputStream);
        processOneSheet(opcPackage, sheetId);
    }

    public void processOne(String filePath, int sheetId) throws Exception {
        OPCPackage opcPackage = OPCPackage.open(filePath);
        processOneSheet(opcPackage, sheetId);
    }

    public void processOne(File file, int sheetId) throws Exception {
        OPCPackage opcPackage = OPCPackage.open(file);
        processOneSheet(opcPackage, sheetId);
    }

    public void processAll(InputStream inputStream) throws Exception {
        OPCPackage opcPackage = OPCPackage.open(inputStream);
        processAll(opcPackage);
    }

    public void processAll(String filePath) throws Exception {
        OPCPackage opcPackage = OPCPackage.open(filePath);
        processAll(opcPackage);
    }

    public void processAll(File file) throws Exception {
        OPCPackage opcPackage = OPCPackage.open(file);
        processAll(opcPackage);
    }

    /**
     * 只遍历一个电子表格，其中sheetId为要遍历的sheet索引，从1开始，1-3
     *
     * @throws Exception
     */
    private void processOneSheet(OPCPackage opcPackage, int sheetId) throws Exception {
        XSSFReader r = new XSSFReader(opcPackage);
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
        SharedStringsTable sst = r.getSharedStringsTable();
        XMLReader parser = fetchSheetParser(sst);
        Iterator<InputStream> sheets = r.getSheetsData();
        while (sheets.hasNext()) {
            InputStream sheet = sheets.next();
            InputSource sheetSource = new InputSource(sheet);
            parser.parse(sheetSource);
            sheet.close();
            // 添加进SheetList内，并清空rowList
            SheetList.add(rowList);
            rowList = new ArrayList<>();
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
            String cellType = attributes.getValue("t");
            if (cellType != null && cellType.equals("s")) {
                nextIsString = true;
            } else {
                nextIsString = false;
            }

            // 检测有没有漏行
            String cellCol = attributes.getValue("r").replaceAll("\\d", "").trim();
            int cellColNum = excelColStrToNum(cellCol);
            if (preCellColNum != 0 && cellColNum - preCellColNum > 1) {
                // 计算两个c标签之间的差值
                int diff = cellColNum - preCellColNum;

                // 循环赋值null
                for (int i = 0; i < (diff - 1); i++) {
                    rowValueList.add(null);
                }
                curCol += (diff - 1);
            }

            // 将上一列的数值赋值当前列的数值
            preCellColNum = cellColNum;

            // 判断上一个标签是否还是c，如果是c则表示漏了一行(使用清除内容会导致没有v标签)
            if (preEle == 'c') {
                rowValueList.add(null);
                curCol++;
            }
            preEle = 'c';
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
        if (nextIsString) {
            int idx = Integer.parseInt(lastContents);
            lastContents = new XSSFRichTextString(sst.getEntryAt(idx)).toString().trim();
            nextIsString = false;
        }

        // v => 单元格的值， 将单元格内容加入rowlist中
        if (name.equals("v")) {
            String value = lastContents.trim();
            value = value.equals("") ? null : value;
            rowValueList.add(value);

            // 修改当前标签名为v
            preEle = 'v';

            // c标签重复次数重置为0
            curCol++;
        } else if (name.equals("row")) {
            // 行尾
            try {
                if (!rowValueList.isEmpty()) {
                    // 将当前list赋值给一个临时队列，检测是否全为null元素
                    List<String> tempRowValues = new ArrayList<>(rowValueList);
                    tempRowValues.removeAll(Collections.singleton(null));

                    // 判断是否为空行，若部位空则添加进list内
                    if (tempRowValues.size() > 0) {
                        rowList.add(rowValueList);
                    }
                }
            } catch (Exception e) {
                e.printStackTrace();
            }

            // 初始化list，存储下一行的内容
            rowValueList = new ArrayList<>();

            // 修改当前标签名为r
            preEle = 'r';

            // preCellColNum重置为0
            preCellColNum = 0;

            // c标签重复次数重置为0
            curCol = 0;
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
     * 获取每一行内容
     *
     * @return
     */
    public List<List<String>> getRowList() {
        return rowList;
    }

    /**
     * 获取每一个sheet的内容
     *
     * @return
     */
    public List<List<List<String>>> getSheetList() {
        return SheetList;
    }
}