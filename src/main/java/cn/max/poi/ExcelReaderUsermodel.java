package cn.max.poi;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

/**
 * usermodel解析法，兼容所有表格
 *
 * @author MaxStar
 * @date 2018/8/31
 */
public class ExcelReaderUsermodel {

    public List<List<String>> process(InputStream inputStream, String sheetName) throws IOException, InvalidFormatException {
        List<List<String>> rowValueList = new ArrayList<>();
        Workbook workbook = WorkbookFactory.create(inputStream);
        Sheet sheet = workbook.getSheet(sheetName);

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
     *
     * @param row
     * @param celNum
     * @return
     */
    private static String getCellValue(Row row, int celNum) {
        try {
            Cell cell = row.getCell(celNum, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
            cell.setCellType(CellType.STRING);
            return cell.getStringCellValue().trim();
        } catch (NullPointerException e) {
            return null;
        }
    }
}
