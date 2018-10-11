package cn.max.poi.writer;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

/**
 * @author MaxStar
 * @date 2018/9/6
 */
public class ExcelWriter {

    /**
     * 创建sheet
     *
     * @param sheetTitle  单元格名称
     * @param headers     表头
     * @param rowList     数据
     * @param isExcel2003 是否excel2003
     *                    true:HSSF
     *                    false:XSSF
     */
    public static Workbook createWorkBook(String sheetTitle, String[] headers, List<List<String>> rowList, boolean isExcel2003) {
        Workbook workBook;
        if (isExcel2003) {
            workBook = new HSSFWorkbook();
        } else {
            workBook = new XSSFWorkbook();
        }
        Sheet sheet = workBook.createSheet(sheetTitle);
        createHeader(headers, sheet, setHeaderStyle(workBook));
        createBody(rowList, sheet, setBodyStyle(workBook));
        return workBook;
    }

    /**
     * 创建表头
     *
     * @param headers 表头
     * @param sheet   表
     */
    private static void createHeader(String[] headers, Sheet sheet, CellStyle headerStyle) {
        Row row = sheet.createRow(0);
        for (int i = 0; i < headers.length; i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(headerStyle);
        }
    }

    /**
     * 创建正文单元格
     *
     * @param rowList 数据
     * @param sheet   表
     */
    private static void createBody(List<List<String>> rowList, Sheet sheet, CellStyle bodyStyle) {
        for (int i = 0; i < rowList.size(); i++) {
            Row row = sheet.createRow(i + 1);
            List<String> dataList = rowList.get(i);
            for (int j = 0; j < dataList.size(); j++) {
                Cell cell = row.createCell(j);
                cell.setCellValue(dataList.get(j));
                cell.setCellStyle(bodyStyle);
            }
            sheet.autoSizeColumn(i);
        }
    }

    /**
     * 设置表头格式
     *
     * @param workBook 工作簿
     * @return 表头样式
     */
    private static CellStyle setHeaderStyle(Workbook workBook) {
        CellStyle headerStyle = workBook.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.index);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        Font font = workBook.createFont();
        font.setFontName("微软雅黑");
        font.setFontHeightInPoints((short) 12);
        font.setBold(true);
        headerStyle.setFont(font);
        return headerStyle;
    }

    /**
     * 设置正文单元格格式
     *
     * @param workBook 工作簿
     * @return 单元格样式
     */
    private static CellStyle setBodyStyle(Workbook workBook) {
        CellStyle bodyStyle = workBook.createCellStyle();
        bodyStyle.setAlignment(HorizontalAlignment.CENTER);
        bodyStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        Font font = workBook.createFont();
        font.setFontName("微软雅黑");
        font.setFontHeightInPoints((short) 11);
        bodyStyle.setFont(font);
        return bodyStyle;
    }

    /**
     * 输出工作簿
     *
     * @param path     输出路径
     * @param workbook 工作簿
     * @throws IOException
     */
    public static void write(String path, Workbook workbook) throws IOException {
        OutputStream outputStream = new FileOutputStream(path);
        workbook.write(outputStream);
        outputStream.close();
        workbook.close();
    }
}