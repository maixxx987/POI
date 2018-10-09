package cn.max.poi;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.NumberToTextConverter;

import java.io.IOException;
import java.io.InputStream;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;

import static org.apache.commons.lang3.StringUtils.isBlank;

/**
 * usermodel解析法，兼容所有表格
 *
 * @author MaxStar
 * @date 2018/8/31
 */
public class ExcelReaderUsermodel {

    private static FormulaEvaluator evaluator;

    public List<List<String>> process(InputStream inputStream, String sheetName) throws IOException, InvalidFormatException {
        List<List<String>> rowValueList = new ArrayList<>();
        Workbook workbook = WorkbookFactory.create(inputStream);
        Sheet sheet = workbook.getSheet(sheetName == null ? "Sheet1" : sheetName);
        evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        for (int rowNum = 0; rowNum <= sheet.getLastRowNum(); rowNum++) {
            Row row = sheet.getRow(rowNum);
            List<String> cellValueList = new ArrayList<>();
            for (int j = 0; j < sheet.getRow(rowNum).getLastCellNum(); j++) {
                cellValueList.add(getCellValue(row, j));
            }
            rowValueList.add(cellValueList);
        }
        return rowValueList;
    }

    /**
     * 解析所有单元格为String（日期时间等特殊单元格暂时无法解析）
     * 日期中各数值来源参考： https://poi.apache.org/apidocs/org/apache/poi/ss/usermodel/BuiltinFormats.html
     *
     * @param row
     * @param celNum
     * @return
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

                    if (DateUtil.isCellDateFormatted(cell)) {
                        SimpleDateFormat sdf;
                        if (format == (short) 0x12 || format == (short) 0x13 || format == (short) 0x14 || format == (short) 0x15) {
                            sdf = new SimpleDateFormat("HH:mm:ss");
                        } else if (format == (short) 0x16) {
                            sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:sss");
                        } else {
                            sdf = new SimpleDateFormat("yyyy-MM-dd");
                        }
                        return sdf.format(DateUtil.getJavaDate(cell.getNumericCellValue()));
                        // 百分百
                    } else if (format == 9 || format == (short) 0xa) {
                        return NumberFormat.getPercentInstance().format(cell.getNumericCellValue());
                    } else if (format == (short) 0xb || format == (short) 0x30) {
                        return String.format("%E",cell.getNumericCellValue());
                    } else {
                        return String.valueOf(NumberToTextConverter.toText(cell.getNumericCellValue()));
                    }

                default:
                    cell.setCellType(CellType.STRING);
                    return cell.getStringCellValue().trim();
            }
        } catch (
                NullPointerException e) {
            return null;
        }
    }
}
