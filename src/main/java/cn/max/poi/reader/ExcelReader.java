package cn.max.poi.reader;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import static org.apache.commons.lang3.StringUtils.isBlank;

/**
 * 使用DOM方法解析Excel（能解析所有版本的Excel）
 *
 * @author MaxStar
 * @date 2018/8/31
 */
public class ExcelReader {

    private static final DataFormatter FORMATTER = new DataFormatter();
    private static FormulaEvaluator evaluator;


    /**
     * 解析单个sheet
     *
     * @param inputStream 输入流
     * @param sheetName   sheet名字（若未填写则默认为Sheet1）
     * @return
     * @throws IOException
     * @throws InvalidFormatException
     */
    public static List<List<String>> process(InputStream inputStream, String sheetName) throws IOException, InvalidFormatException {
        Workbook workbook = createWorkBook(inputStream);
        Sheet sheet = workbook.getSheet(isBlank(sheetName) ? "Sheet1" : sheetName);
        return getRowValueList(sheet);
    }


    /**
     * 解析全部sheet
     *
     * @param inputStream 输入流
     * @return
     * @throws IOException
     * @throws InvalidFormatException
     */
    public static List<List<List<String>>> processAll(InputStream inputStream) throws IOException, InvalidFormatException {
        List<List<List<String>>> sheetList = new ArrayList<>();
        Workbook workbook = createWorkBook(inputStream);
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            sheetList.add(getRowValueList(workbook.getSheetAt(i)));
        }
        return sheetList;
    }

    /**
     * 创建workbook
     *
     * @param inputStream 输入流
     * @return
     * @throws IOException
     * @throws InvalidFormatException
     */
    private static Workbook createWorkBook(InputStream inputStream) throws IOException, InvalidFormatException {
        Workbook workbook = WorkbookFactory.create(inputStream);
        evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        return workbook;
    }

    /**
     * 获取每一行的值
     *
     * @param sheet 单元表
     * @return
     */
    private static List<List<String>> getRowValueList(Sheet sheet) {
        List<List<String>> rowList = new ArrayList<>();
        for (int rowNum = 0; rowNum <= sheet.getLastRowNum(); rowNum++) {
            Row row = sheet.getRow(rowNum);
            List<String> cellValueList = new ArrayList<>();
            for (int j = 0; j < sheet.getRow(rowNum).getLastCellNum(); j++) {
                cellValueList.add(getCellValue(row, j));
            }
            rowList.add(cellValueList);
        }
        return rowList;
    }


    /**
     * 解析所有单元格为String
     * 日期中各数值来源参考： https://poi.apache.org/apidocs/org/apache/poi/ss/usermodel/BuiltinFormats.html
     *
     * @param row    行
     * @param celNum 列
     * @return 单元格的值
     */
    private static String getCellValue(Row row, int celNum) {
        try {
            Cell cell = row.getCell(celNum);
            switch (cell.getCellTypeEnum()) {
                case BLANK:
                    return null;
                case STRING:
                    String value = cell.getStringCellValue().trim();
                    if (isBlank(value)) {
                        value = null;
                    }
                    return value;
                case BOOLEAN:
                    return String.valueOf(cell.getBooleanCellValue());
                case FORMULA:
                    return evaluator.evaluate(cell).formatAsString();
                case _NONE:
                    return null;
                case NUMERIC:
                    short format = cell.getCellStyle().getDataFormat();
                    // 判断是否日期或时间
                    String formatString;
                    if (DateUtil.isCellDateFormatted(cell)) {
                        if (format == (short) 0x12 || format == (short) 0x13 || format == (short) 0x14 || format == (short) 0x15) {
                            formatString = "HH:mm:ss";
                        } else {
                            formatString = "yyyy-MM-dd";
//                            sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                        }
                    } else {
                        formatString = cell.getCellStyle().getDataFormatString();
                        if (formatString == null) {
                            formatString = BuiltinFormats.getBuiltinFormat(format);
                        }
                    }
                    return FORMATTER.formatRawCellContents(cell.getNumericCellValue(), format, formatString);
                default:
                    cell.setCellType(CellType.STRING);
                    return cell.getStringCellValue().trim();
            }
        } catch (NullPointerException e) {
            return null;
        }
    }
}